"""Bounded Microsoft Graph v1.0 client.

Egress is limited to graph.microsoft.com and, for file bytes only, the SharePoint download host
Graph redirects to. Paths are built by the connector, never taken from model input or handles.
"""

from __future__ import annotations

import re
import uuid
from dataclasses import dataclass
from typing import Any
from urllib.parse import quote, urlsplit

import httpx

from . import __version__
from .auth import TokenSource
from .config import GRAPH_BASE
from .errors import GraphError

DOWNLOAD_HOST_SUFFIXES = (
    ".sharepoint.com",
    ".sharepoint.us",
    ".sharepoint.de",
    ".sharepoint.cn",
    ".sharepoint-mil.us",
)
MAX_RETRY_AFTER_SECONDS = 300
DEFAULT_RETRY_AFTER_SECONDS = 10
_SAFE_PATH_RE = re.compile(r"^/[A-Za-z0-9!_\-.,:/%()=$']*$")

_STATUS_CODES = {
    400: "bad_request",
    401: "unauthorized",
    403: "forbidden",
    404: "not_found",
    409: "conflict",
    410: "not_found",
    413: "file_too_large",
    416: "bad_request",
    423: "locked",
    429: "throttled",
}


def seg(value: str) -> str:
    """Percent-encode one Graph path segment."""
    return quote(value, safe="!,")


@dataclass
class CallInfo:
    request_id: str | None = None
    client_request_id: str | None = None
    status: int | None = None


def _retry_after(resp: httpx.Response) -> int | None:
    raw = resp.headers.get("Retry-After")
    if raw is None:
        return DEFAULT_RETRY_AFTER_SECONDS if resp.status_code in (429, 503) else None
    try:
        return max(1, min(int(float(raw)), MAX_RETRY_AFTER_SECONDS))
    except ValueError:
        return DEFAULT_RETRY_AFTER_SECONDS


def _error(resp: httpx.Response, client_request_id: str) -> GraphError:
    graph_code = None
    request_id = resp.headers.get("request-id")
    try:
        body = resp.json()
        err = body.get("error", {}) if isinstance(body, dict) else {}
        graph_code = err.get("code")
        request_id = request_id or (err.get("innerError") or {}).get("request-id")
    except ValueError:
        pass
    status = resp.status_code
    code = _STATUS_CODES.get(status, "upstream_unavailable" if status >= 500 else "graph_error")
    messages = {
        "unauthorized": "Microsoft Graph rejected the session; sign in again.",
        "forbidden": "You do not have access to this SharePoint resource.",
        "not_found": "The SharePoint resource was not found or is no longer shared with you.",
        "throttled": "Microsoft Graph is throttling requests; retry after the given delay.",
        "upstream_unavailable": "Microsoft Graph is temporarily unavailable.",
    }
    return GraphError(
        code,
        messages.get(code, "Microsoft Graph returned an error."),
        status=status,
        retry_after_seconds=_retry_after(resp) if status in (429, 503) else None,
        request_id=request_id,
        client_request_id=client_request_id,
        details={"graph_code": graph_code} if graph_code else {},
    )


class GraphClient:
    def __init__(
        self,
        tokens: TokenSource,
        *,
        timeout: float = 20.0,
        transport: httpx.BaseTransport | None = None,
        base_url: str = GRAPH_BASE,
    ):
        self._tokens = tokens
        self._base = base_url.rstrip("/")
        self._http = httpx.Client(
            timeout=timeout,
            transport=transport,
            follow_redirects=False,
            headers={"User-Agent": f"xtium-sharepoint-mcp/{__version__}"},
        )
        self.last = CallInfo()

    def close(self) -> None:
        self._http.close()

    def _url(self, path: str) -> str:
        if not _SAFE_PATH_RE.match(path) or ".." in path:
            raise GraphError("invalid_request", "Refusing to call an unsafe Graph path.")
        return self._base + path

    def _send(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        json: Any = None,
        stream: bool = False,
    ) -> httpx.Response:
        url = self._url(path)
        for attempt in (0, 1):
            client_request_id = str(uuid.uuid4())
            token = self._tokens.access_token(force_refresh=attempt == 1)
            request = self._http.build_request(
                method,
                url,
                params=params,
                json=json,
                headers={
                    "Authorization": f"Bearer {token}",
                    "client-request-id": client_request_id,
                    "return-client-request-id": "true",
                    "Accept": "application/json",
                },
            )
            try:
                resp = self._http.send(request, stream=stream)
            except httpx.TimeoutException as exc:
                raise GraphError(
                    "timeout", "Microsoft Graph did not respond in time.",
                    client_request_id=client_request_id,
                ) from exc
            except httpx.HTTPError as exc:
                raise GraphError(
                    "network_error", "Could not reach Microsoft Graph.",
                    client_request_id=client_request_id,
                ) from exc
            self.last = CallInfo(resp.headers.get("request-id"), client_request_id, resp.status_code)
            if resp.status_code == 401 and attempt == 0:
                resp.close()
                continue
            return resp
        return resp

    def _json(self, resp: httpx.Response) -> dict[str, Any]:
        if resp.status_code >= 400:
            resp.read()
            raise _error(resp, self.last.client_request_id or "")
        try:
            data = resp.json()
        except ValueError as exc:
            raise GraphError(
                "invalid_response", "Microsoft Graph returned a non-JSON response.",
                request_id=self.last.request_id, client_request_id=self.last.client_request_id,
            ) from exc
        if not isinstance(data, dict):
            raise GraphError("invalid_response", "Unexpected Graph response shape.")
        return data

    def get(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        return self._json(self._send("GET", path, params=params))

    def post(self, path: str, body: dict[str, Any]) -> dict[str, Any]:
        return self._json(self._send("POST", path, json=body))

    def download(self, path: str, max_bytes: int) -> bytes:
        """Fetch file bytes, following one redirect to a SharePoint host without the Graph token."""
        resp = self._send("GET", path, stream=True)
        try:
            if resp.status_code in (301, 302, 303, 307, 308):
                location = resp.headers.get("Location", "")
                parts = urlsplit(location)
                host = (parts.hostname or "").lower()
                if parts.scheme != "https" or not host.endswith(DOWNLOAD_HOST_SUFFIXES):
                    raise GraphError(
                        "egress_blocked", "Graph redirected the download to an unapproved host.",
                        request_id=self.last.request_id,
                        client_request_id=self.last.client_request_id,
                    )
                resp.close()
                try:
                    resp = self._http.send(
                        self._http.build_request("GET", location), stream=True
                    )
                except httpx.HTTPError as exc:
                    raise GraphError("network_error", "Could not download the file.") from exc
            if resp.status_code >= 400:
                resp.read()
                raise _error(resp, self.last.client_request_id or "")
            declared = resp.headers.get("Content-Length")
            if declared and declared.isdigit() and int(declared) > max_bytes:
                raise GraphError("file_too_large", "The file is larger than the read limit.")
            chunks: list[bytes] = []
            total = 0
            for chunk in resp.iter_bytes():
                total += len(chunk)
                if total > max_bytes:
                    raise GraphError("file_too_large", "The file is larger than the read limit.")
                chunks.append(chunk)
            return b"".join(chunks)
        finally:
            resp.close()
