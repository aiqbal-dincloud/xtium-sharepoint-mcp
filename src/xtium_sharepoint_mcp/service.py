"""SharePoint tool logic. Every call runs as the signed-in user; Graph enforces their ACLs."""

from __future__ import annotations

import re
import time
from collections.abc import Callable
from typing import Any, TypeVar
from urllib.parse import parse_qs, urlsplit

from .audit import AuditLog
from .auth import Identity, TokenSource
from .config import SCOPE_FILES_READ_ALL, SCOPE_SITES_READ_ALL, Settings
from .errors import AuthError, ConnectorError, GraphError
from .extract import extension, extract, is_supported
from .graph import GraphClient, seg
from .handles import HandleCodec

T = TypeVar("T")

STATUS_TOOL = "get_connection_status"
TOOL_SCOPES: dict[str, tuple[frozenset[str], ...]] = {
    "sharepoint_search": (frozenset({SCOPE_FILES_READ_ALL}), frozenset({SCOPE_SITES_READ_ALL})),
    "sharepoint_get_site": (frozenset({SCOPE_SITES_READ_ALL}),),
    "sharepoint_list_folder": (frozenset({SCOPE_FILES_READ_ALL}), frozenset({SCOPE_SITES_READ_ALL})),
    "sharepoint_read_file": (frozenset({SCOPE_FILES_READ_ALL}), frozenset({SCOPE_SITES_READ_ALL})),
}
DATA_TOOLS = tuple(TOOL_SCOPES)

ITEM_SELECT = "id,name,size,file,folder,cTag,lastModifiedDateTime,webUrl,parentReference"
SEARCH_FIELDS = ["id", "name", "size", "webUrl", "lastModifiedDateTime", "parentReference", "file", "folder"]
SHAREPOINT_HOST_RE = re.compile(r"^[a-z0-9\-]{1,63}(-my)?\.sharepoint\.(com|us|de|cn)$")
SITE_PATH_RE = re.compile(r"^/(sites|teams|personal)/[A-Za-z0-9_\-.%]{1,128}(/[A-Za-z0-9_\-.%]{1,128}){0,4}$")
_CONTROL_RE = re.compile(r"[\x00-\x1f\x7f]")


def tools_for(granted: frozenset[str]) -> list[str]:
    return [t for t, options in TOOL_SCOPES.items() if any(o <= granted for o in options)]


class SharePointService:
    def __init__(
        self,
        settings: Settings,
        tokens: TokenSource,
        graph: GraphClient,
        codec_factory: Callable[[Identity], HandleCodec],
        audit: AuditLog,
        rotate_handle_key: Callable[[], None] = lambda: None,
        clock: Callable[[], float] = time.perf_counter,
    ):
        self._settings = settings
        self._limits = settings.limits
        self._tokens = tokens
        self._graph = graph
        self._codec_factory = codec_factory
        self._audit = audit
        self._rotate_handle_key = rotate_handle_key
        self._clock = clock
        self._codec: HandleCodec | None = None
        self._codec_user: str | None = None

    # --- plumbing -------------------------------------------------------------------------

    def _require(self, tool: str) -> tuple[Identity, HandleCodec]:
        identity = self._tokens.identity()
        if identity is None:
            raise AuthError("not_signed_in", "Not signed in. Run sharepoint_sign_in first.")
        if tool not in tools_for(identity.granted_scopes):
            raise AuthError(
                "scope_not_granted",
                "The signed-in session does not have the Graph scope this tool needs.",
                details={"granted_scopes": sorted(identity.granted_scopes)},
            )
        if self._codec is None or self._codec_user != identity.object_id:
            self._codec = self._codec_factory(identity)
            self._codec_user = identity.object_id
        return identity, self._codec

    def reset_handles(self) -> None:
        self._codec = None
        self._codec_user = None

    def signed_in(self, identity: Identity) -> None:
        self.reset_handles()
        self._audit.record(
            event="sign_in",
            outcome="allowed",
            actor_oid=identity.object_id,
            counts={"scopes": len(identity.granted_scopes)},
        )

    def sign_in_failed(self, code: str, client_request_id: str | None = None) -> None:
        self._audit.record(
            event="sign_in",
            outcome="denied",
            error_code=code,
            client_request_id=client_request_id,
        )

    def sign_out(self) -> None:
        """Rotating the key makes every handle issued before sign-out unusable."""
        self.reset_handles()
        self._rotate_handle_key()
        self._audit.record(event="sign_out", outcome="allowed")

    def _run(self, tool: str, fn: Callable[[Identity, HandleCodec, list[str]], T]) -> T:
        started = self._clock()
        resources: list[str] = []
        identity: Identity | None = None
        try:
            identity, codec = self._require(tool)
            result = fn(identity, codec, resources)
        except ConnectorError as exc:
            self._audit.record(
                event=tool,
                outcome="denied" if isinstance(exc, AuthError) else "error",
                actor_oid=identity.object_id if identity else None,
                resource_ids=resources,
                error_code=exc.code,
                duration_ms=int((self._clock() - started) * 1000),
                request_id=exc.request_id or self._graph.last.request_id,
                client_request_id=exc.client_request_id or self._graph.last.client_request_id,
            )
            raise
        counts = result.pop("_audit_counts", None) if isinstance(result, dict) else None
        self._audit.record(
            event=tool,
            outcome="allowed",
            actor_oid=identity.object_id,
            resource_ids=resources,
            duration_ms=int((self._clock() - started) * 1000),
            request_id=self._graph.last.request_id,
            client_request_id=self._graph.last.client_request_id,
            counts=counts,
        )
        return result

    @staticmethod
    def _site_ref(parent: dict[str, Any], codec: HandleCodec) -> str | None:
        site_id = parent.get("siteId")
        if not site_id:
            return None
        try:
            return codec.seal("site", {"s": site_id})
        except ConnectorError:
            return None

    def _item_view(self, item: dict[str, Any], codec: HandleCodec) -> dict[str, Any] | None:
        parent = item.get("parentReference") or {}
        drive_id, item_id = parent.get("driveId"), item.get("id")
        if not drive_id or not item_id:
            return None
        is_folder = "folder" in item
        try:
            handle = codec.seal(
                "folder" if is_folder else "file",
                {"d": drive_id, "i": item_id, **({} if is_folder else {"c": item.get("cTag")})},
            )
        except ConnectorError:
            return None
        view: dict[str, Any] = {
            "handle": handle,
            "kind": "folder" if is_folder else "file",
            "name": item.get("name"),
            "modified": item.get("lastModifiedDateTime"),
            "web_url": item.get("webUrl"),
        }
        if is_folder:
            view["child_count"] = (item.get("folder") or {}).get("childCount")
        else:
            view["size_bytes"] = item.get("size")
            view["mime_type"] = (item.get("file") or {}).get("mimeType")
            view["readable"] = is_supported(item.get("name") or "")
        site = self._site_ref(parent, codec)
        if site:
            view["site"] = site
        return view

    # --- tools ----------------------------------------------------------------------------

    def connection_status(self) -> dict[str, Any]:
        status: dict[str, Any] = {
            "tenant_id": self._settings.tenant_id,
            "client_id": self._settings.client_id,
            "requested_scopes": list(self._settings.scopes),
            "signed_in": False,
            "available_tools": [STATUS_TOOL],
        }
        started = self._clock()
        try:
            identity = self._tokens.identity()
        except AuthError as exc:
            status["auth_error"] = exc.code
            self._audit.record(event=STATUS_TOOL, outcome="denied", error_code=exc.code)
            return status
        if identity is None:
            self._audit.record(event=STATUS_TOOL, outcome="allowed", counts={"signed_in": 0})
            return status
        status.update(
            signed_in=True,
            granted_scopes=sorted(identity.granted_scopes),
            available_tools=[STATUS_TOOL, *tools_for(identity.granted_scopes)],
        )
        try:
            me = self._graph.get("/me", {"$select": "id"})
            status["graph_check"] = "ok" if str(me.get("id", "")).lower() == identity.object_id else "user_mismatch"
        except ConnectorError as exc:
            status["graph_check"] = exc.code
        self._audit.record(
            event=STATUS_TOOL,
            outcome="allowed",
            actor_oid=identity.object_id,
            duration_ms=int((self._clock() - started) * 1000),
            request_id=self._graph.last.request_id,
            client_request_id=self._graph.last.client_request_id,
            counts={"signed_in": 1},
        )
        return status

    def search(
        self,
        query: str,
        site: str | None = None,
        include_list_items: bool = False,
        page_from: int = 0,
        size: int | None = None,
    ) -> dict[str, Any]:
        lim = self._limits

        def run(identity: Identity, codec: HandleCodec, resources: list[str]) -> dict[str, Any]:
            q = _CONTROL_RE.sub(" ", query or "").strip()
            if not q:
                raise ConnectorError("bad_request", "query must not be empty.")
            if len(q) > lim.query_max_chars:
                raise ConnectorError("bad_request", f"query is limited to {lim.query_max_chars} characters.")
            n = lim.search_default_size if size is None else size
            if not 1 <= n <= lim.search_max_size:
                raise ConnectorError("bad_request", f"size must be 1-{lim.search_max_size}.")
            if not 0 <= page_from <= lim.search_max_from:
                raise ConnectorError("bad_request", f"page_from must be 0-{lim.search_max_from}.")

            site_id = None
            if site:
                site_id = codec.open(site, "site")["s"]
                resources.append(site_id)
                web = self._graph.get(f"/sites/{seg(site_id)}", {"$select": "id,webUrl"})
                web_url = str(web.get("webUrl") or "")
                if not web_url.startswith("https://") or '"' in web_url:
                    raise GraphError("invalid_response", "Graph returned an unusable site URL.")
                q = f'({q}) path:"{web_url}"'

            warnings: list[str] = []
            entity_types = ["driveItem"]
            if include_list_items:
                if SCOPE_SITES_READ_ALL in identity.granted_scopes:
                    entity_types.append("listItem")
                else:
                    warnings.append("list_items_need_sites_scope")

            def body(types: list[str]) -> dict[str, Any]:
                return {
                    "requests": [
                        {
                            "entityTypes": types,
                            "query": {"queryString": q},
                            "from": page_from,
                            "size": n,
                            "fields": SEARCH_FIELDS,
                        }
                    ]
                }

            try:
                data = self._graph.post("/search/query", body(entity_types))
            except GraphError as exc:
                if exc.code != "forbidden" or len(entity_types) == 1:
                    raise
                warnings.append("list_items_forbidden_fell_back_to_files")
                data = self._graph.post("/search/query", body(["driveItem"]))

            results: list[dict[str, Any]] = []
            total, more, skipped = 0, False, 0
            for response in data.get("value") or []:
                for container in response.get("hitsContainers") or []:
                    total += int(container.get("total") or 0)
                    more = more or bool(container.get("moreResultsAvailable"))
                    for hit in container.get("hits") or []:
                        resource = hit.get("resource") or {}
                        parent = resource.get("parentReference") or {}
                        if site_id and parent.get("siteId") and parent["siteId"] != site_id:
                            skipped += 1
                            continue
                        if "listItem" in str(resource.get("@odata.type", "")):
                            view = {
                                "kind": "list_item",
                                "name": resource.get("name"),
                                "web_url": resource.get("webUrl"),
                                "modified": resource.get("lastModifiedDateTime"),
                            }
                            site_ref = self._site_ref(parent, codec)
                            if site_ref:
                                view["site"] = site_ref
                        else:
                            view = self._item_view(resource, codec)
                        if view is None:
                            skipped += 1
                            continue
                        view["rank"] = hit.get("rank")
                        results.append(view)
            next_from = page_from + n
            out: dict[str, Any] = {
                "results": results,
                "total_estimate": total,
                "next_page_from": next_from if more and next_from <= lim.search_max_from else None,
                "_audit_counts": {"results": len(results), "skipped": skipped},
            }
            if warnings:
                out["warnings"] = warnings
            return out

        return self._run("sharepoint_search", run)

    def get_site(self, site: str = "root") -> dict[str, Any]:
        def run(identity: Identity, codec: HandleCodec, resources: list[str]) -> dict[str, Any]:
            path = self._site_path(site, codec)
            info = self._graph.get(path, {"$select": "id,displayName,name,webUrl,description"})
            site_id = info.get("id")
            if not site_id:
                raise GraphError("invalid_response", "Graph returned a site without an ID.")
            resources.append(site_id)
            drives = self._graph.get(f"/sites/{seg(site_id)}/drives", {"$select": "id,name,driveType,webUrl"})
            libraries = []
            for drive in drives.get("value") or []:
                if not drive.get("id"):
                    continue
                libraries.append(
                    {
                        "name": drive.get("name"),
                        "drive_type": drive.get("driveType"),
                        "web_url": drive.get("webUrl"),
                        "root_folder": codec.seal("folder", {"d": drive["id"], "i": "root"}),
                    }
                )
            return {
                "site": codec.seal("site", {"s": site_id}),
                "name": info.get("displayName") or info.get("name"),
                "description": info.get("description"),
                "web_url": info.get("webUrl"),
                "libraries": libraries,
                "_audit_counts": {"libraries": len(libraries)},
            }

        return self._run("sharepoint_get_site", run)

    def _site_path(self, site: str, codec: HandleCodec) -> str:
        site = (site or "root").strip()
        if site == "root":
            return "/sites/root"
        if site.startswith("sps_"):
            return f"/sites/{seg(codec.open(site, 'site')['s'])}"
        parts = urlsplit(site if "://" in site else f"https://{site}")
        host = (parts.hostname or "").lower()
        path = parts.path.rstrip("/")
        if parts.scheme != "https" or not SHAREPOINT_HOST_RE.match(host):
            raise ConnectorError("bad_request", "site must be 'root', a site handle or a SharePoint site URL.")
        if not path:
            return f"/sites/{host}"
        if not SITE_PATH_RE.match(path) or ".." in path:
            raise ConnectorError("bad_request", "Only /sites/... or /teams/... site URLs are accepted.")
        return f"/sites/{host}:{path}"

    def list_folder(self, folder: str, cursor: str | None = None) -> dict[str, Any]:
        lim = self._limits

        def run(identity: Identity, codec: HandleCodec, resources: list[str]) -> dict[str, Any]:
            ref = codec.open(folder, "folder")
            drive_id, item_id = ref["d"], ref["i"]
            resources.append(f"{drive_id}/{item_id}")
            base = f"/drives/{seg(drive_id)}/" + (
                "root/children" if item_id == "root" else f"items/{seg(item_id)}/children"
            )
            params: dict[str, Any] = {"$top": lim.folder_page_size, "$select": ITEM_SELECT}
            if cursor:
                state = codec.open(cursor, "cursor")
                if state.get("d") != drive_id or state.get("i") != item_id:
                    raise ConnectorError("bad_request", "The cursor belongs to a different folder.")
                params["$skiptoken"] = state["k"]
            data = self._graph.get(base, params)
            items = [v for v in (self._item_view(i, codec) for i in data.get("value") or []) if v]
            next_cursor = None
            next_link = data.get("@odata.nextLink")
            if next_link:
                token = self._skiptoken(next_link)
                if token:
                    next_cursor = codec.seal("cursor", {"d": drive_id, "i": item_id, "k": token})
            return {
                "items": items,
                "next_cursor": next_cursor,
                "_audit_counts": {"items": len(items)},
            }

        return self._run("sharepoint_list_folder", run)

    @staticmethod
    def _skiptoken(next_link: str) -> str | None:
        parts = urlsplit(next_link)
        if parts.scheme != "https" or (parts.hostname or "").lower() != "graph.microsoft.com":
            return None
        if not (parts.path.startswith("/v1.0/drives/") and parts.path.endswith("/children")):
            return None
        values = parse_qs(parts.query).get("$skiptoken")
        return values[0] if values else None

    def read_file(self, file: str, max_chars: int | None = None, start_page: int = 1) -> dict[str, Any]:
        lim = self._limits

        def run(identity: Identity, codec: HandleCodec, resources: list[str]) -> dict[str, Any]:
            chars = lim.read_default_chars if max_chars is None else max_chars
            if not 1 <= chars <= lim.read_max_chars:
                raise ConnectorError("bad_request", f"max_chars must be 1-{lim.read_max_chars}.")
            ref = codec.open(file, "file")
            drive_id, item_id = ref["d"], ref["i"]
            resources.append(f"{drive_id}/{item_id}")
            item_path = f"/drives/{seg(drive_id)}/items/{seg(item_id)}"
            meta = self._graph.get(item_path, {"$select": ITEM_SELECT})
            name = meta.get("name") or ""
            if "folder" in meta or "file" not in meta:
                raise ConnectorError("not_a_file", "The handle points at a folder, not a file.")
            size = int(meta.get("size") or 0)
            if size > lim.read_max_bytes:
                raise ConnectorError(
                    "file_too_large",
                    f"The file is {size} bytes; the read limit is {lim.read_max_bytes}.",
                )
            if not is_supported(name):
                raise ConnectorError(
                    "unsupported_format",
                    "Only PDF, DOCX and plain-text files can be read in this pilot.",
                    details={"extension": extension(name) or None},
                )
            data = self._graph.download(f"{item_path}/content", lim.read_max_bytes)
            result = extract(name, data, lim, chars, start_page)
            out: dict[str, Any] = {
                "name": name,
                "mime_type": (meta.get("file") or {}).get("mimeType"),
                "size_bytes": size,
                "modified": meta.get("lastModifiedDateTime"),
                "web_url": meta.get("webUrl"),
                "changed_since_handle": bool(ref.get("c")) and ref.get("c") != meta.get("cTag"),
                "format": result.format,
                "content_trust": "untrusted_document_text",
                "text": result.text,
                "truncated": result.truncated,
                "_audit_counts": {"bytes": len(data), "chars": len(result.text)},
            }
            if result.total_pages is not None:
                out["pages"] = {
                    "start": start_page,
                    "read": result.pages_read,
                    "total": result.total_pages,
                    "next_start_page": result.next_page,
                }
            return out

        return self._run("sharepoint_read_file", run)
