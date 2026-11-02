from __future__ import annotations

import httpx
import pytest

from xtium_sharepoint_mcp.errors import GraphError, HandleError
from xtium_sharepoint_mcp.graph import GraphClient
from xtium_sharepoint_mcp.handles import HandleCodec

from .conftest import FIN_SITE, OTHER_TENANT, TENANT, USER_A, USER_B, FakeGraph, FakeTokens

KEY = b"\x07" * 32


class TestHandles:
    def test_round_trip_is_opaque(self) -> None:
        codec = HandleCodec(KEY, TENANT, USER_A)
        handle = codec.seal("file", {"d": "b!FinanceDrive01", "i": "01FINBUDGET", "c": "x"})
        assert handle.startswith("spf_")
        assert "FinanceDrive01" not in handle and "01FINBUDGET" not in handle
        assert codec.open(handle, "file")["i"] == "01FINBUDGET"

    @pytest.mark.parametrize(("tenant", "user"), [(TENANT, USER_B), (OTHER_TENANT, USER_A)])
    def test_bound_to_tenant_and_user(self, tenant: str, user: str) -> None:
        handle = HandleCodec(KEY, TENANT, USER_A).seal("site", {"s": FIN_SITE})
        with pytest.raises(HandleError) as err:
            HandleCodec(KEY, tenant, user).open(handle, "site")
        assert err.value.code == "invalid_handle"

    def test_key_rotation_invalidates(self) -> None:
        handle = HandleCodec(KEY, TENANT, USER_A).seal("site", {"s": FIN_SITE})
        with pytest.raises(HandleError):
            HandleCodec(b"\x08" * 32, TENANT, USER_A).open(handle, "site")

    def test_kind_is_enforced(self) -> None:
        codec = HandleCodec(KEY, TENANT, USER_A)
        folder = codec.seal("folder", {"d": "b!FinanceDrive01", "i": "root"})
        with pytest.raises(HandleError) as err:
            codec.open(folder, "file")
        assert err.value.code == "wrong_handle_kind"
        forged = "spf_" + folder[4:]
        with pytest.raises(HandleError) as err:
            codec.open(forged, "file")
        assert err.value.code == "invalid_handle"

    @pytest.mark.parametrize(
        "value",
        ["https://graph.microsoft.com/v1.0/me", "/drives/x", "b!ok?$expand=x", "../../me", "a" * 300],
    )
    def test_rejects_urls_and_paths_as_ids(self, value: str) -> None:
        with pytest.raises(HandleError):
            HandleCodec(KEY, TENANT, USER_A).seal("file", {"d": value, "i": "01ABC"})

    @pytest.mark.parametrize("raw", ["", "https://contoso.sharepoint.com/x", "spf_!!!", "spf_" + "A" * 3000])
    def test_garbage_is_refused(self, raw: str) -> None:
        with pytest.raises(HandleError):
            HandleCodec(KEY, TENANT, USER_A).open(raw, "file")


def _client(graph: FakeGraph, tokens: FakeTokens | None = None) -> tuple[GraphClient, FakeTokens]:
    tokens = tokens or FakeTokens()
    return GraphClient(tokens, transport=httpx.MockTransport(graph)), tokens


class TestGraphClient:
    def test_sets_correlation_headers_and_captures_request_id(self) -> None:
        graph = FakeGraph()
        graph.on("GET", "/v1.0/me", lambda r: httpx.Response(200, json={"id": "x"}, headers={"request-id": "rid-1"}))
        client, _ = _client(graph)
        client.get("/me")
        sent = graph.requests[0]
        assert sent.headers["Authorization"] == f"Bearer token-{USER_A}-0"
        assert sent.headers["return-client-request-id"] == "true"
        assert client.last.request_id == "rid-1"
        assert client.last.client_request_id == sent.headers["client-request-id"]

    def test_401_refreshes_once(self) -> None:
        graph = FakeGraph()
        graph.on("GET", "/v1.0/me", lambda r: httpx.Response(
            200 if r.headers["Authorization"].endswith("-1") else 401, json={"id": "x"}))
        client, tokens = _client(graph)
        assert client.get("/me") == {"id": "x"}
        assert tokens.refreshes == 1 and len(graph.requests) == 2

    def test_persistent_401_is_typed(self) -> None:
        graph = FakeGraph()
        graph.on("GET", "/v1.0/me", lambda r: httpx.Response(
            401, json={"error": {"code": "InvalidAuthenticationToken"}}))
        client, _ = _client(graph)
        with pytest.raises(GraphError) as err:
            client.get("/me")
        assert err.value.code == "unauthorized"
        assert err.value.details == {"graph_code": "InvalidAuthenticationToken"}
        assert len(graph.requests) == 2

    @pytest.mark.parametrize(
        ("status", "headers", "code", "retry"),
        [
            (403, {}, "forbidden", None),
            (404, {}, "not_found", None),
            (429, {"Retry-After": "7"}, "throttled", 7),
            (429, {}, "throttled", 10),
            (429, {"Retry-After": "86400"}, "throttled", 300),
            (503, {"Retry-After": "2"}, "upstream_unavailable", 2),
            (500, {}, "upstream_unavailable", None),
        ],
    )
    def test_status_mapping(self, status: int, headers: dict, code: str, retry: int | None) -> None:
        graph = FakeGraph()
        graph.on("POST", "/v1.0/search/query", lambda r: httpx.Response(
            status, headers={"request-id": "rid-9", **headers},
            json={"error": {"code": "x", "message": "SECRET name.docx"}}))
        client, _ = _client(graph)
        with pytest.raises(GraphError) as err:
            client.post("/search/query", {})
        assert err.value.code == code
        assert err.value.retry_after_seconds == retry
        assert err.value.request_id == "rid-9"
        assert "SECRET" not in str(err.value)
        assert len(graph.requests) == 1

    @pytest.mark.parametrize("path", ["https://evil.example/x", "/me/../../x", "/me?x=1", "/me#frag"])
    def test_unsafe_paths_refused(self, path: str) -> None:
        client, _ = _client(FakeGraph())
        with pytest.raises(GraphError) as err:
            client.get(path)
        assert err.value.code == "invalid_request"

    def test_download_follows_sharepoint_redirect_without_token(self) -> None:
        graph = FakeGraph()
        graph.on("GET", "/v1.0/drives/b!D/items/01X/content", lambda r: httpx.Response(
            302, headers={"Location": "https://contoso.sharepoint.com/_layouts/15/download.aspx?t=1"}))
        graph.on("GET", "contoso.sharepoint.com/_layouts/15/download.aspx",
                 lambda r: httpx.Response(200, content=b"hello"))
        client, _ = _client(graph)
        assert client.download("/drives/b!D/items/01X/content", 100) == b"hello"
        assert "Authorization" not in graph.requests[1].headers

    def test_download_to_unapproved_host_blocked(self) -> None:
        graph = FakeGraph()
        graph.on("GET", "/v1.0/drives/b!D/items/01X/content", lambda r: httpx.Response(
            302, headers={"Location": "https://contoso.sharepoint.com.evil.example/x"}))
        client, _ = _client(graph)
        with pytest.raises(GraphError) as err:
            client.download("/drives/b!D/items/01X/content", 100)
        assert err.value.code == "egress_blocked"
        assert len(graph.requests) == 1

    @pytest.mark.parametrize("declare", [True, False])
    def test_download_cap(self, declare: bool) -> None:
        graph = FakeGraph()
        body = b"x" * 101
        headers = {} if declare else {"Transfer-Encoding": "chunked"}
        graph.on("GET", "/v1.0/drives/b!D/items/01X/content", lambda r: httpx.Response(
            200, content=body if declare else iter([body[:60], body[60:]]), headers=headers))
        client, _ = _client(graph)
        with pytest.raises(GraphError) as err:
            client.download("/drives/b!D/items/01X/content", 100)
        assert err.value.code == "file_too_large"
