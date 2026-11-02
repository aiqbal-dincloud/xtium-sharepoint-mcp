from __future__ import annotations

import json

import httpx
import pytest

from xtium_sharepoint_mcp.errors import AuthError, ConnectorError, ExtractionError, GraphError, HandleError
from xtium_sharepoint_mcp.service import tools_for

from .conftest import FIN_SITE, USER_A, USER_B, Env, fixture, make_docx, make_pdf

SEARCH = "/v1.0/search/query"
FORBIDDEN_IN_OUTPUT = ["SECRET-EXCERPT", "SECRET-DOWNLOAD", "downloadUrl", "summary", "b!FinanceDrive01",
                       "01FINBUDGET", "token-"]


def _no_leaks(payload: object) -> None:
    text = json.dumps(payload)
    for needle in FORBIDDEN_IN_OUTPUT:
        assert needle not in text, needle


def _file_handle(env: Env, name: str = "notes.txt", item: str = "01NOTESTXT", ctag: str = "c1",
                 user: str = USER_A) -> str:
    return env.codec(user).seal("file", {"d": "b!FinanceDrive01", "i": item, "c": ctag})


def _meta(env: Env, name: str, size: int, ctag: str = "c1", item: str = "01NOTESTXT") -> None:
    env.graph.on("GET", f"/v1.0/drives/b!FinanceDrive01/items/{item}", {
        "id": item, "name": name, "size": size, "cTag": ctag, "file": {"mimeType": "x"},
        "lastModifiedDateTime": "2026-10-02T09:00:00Z",
        "webUrl": f"https://contoso.sharepoint.com/sites/Finance/Shared%20Documents/{name}",
        "parentReference": {"driveId": "b!FinanceDrive01"},
    })


def _content(env: Env, data: bytes, item: str = "01NOTESTXT") -> None:
    env.graph.on("GET", f"/v1.0/drives/b!FinanceDrive01/items/{item}/content",
                 lambda r: httpx.Response(302, headers={"Location": "https://contoso.sharepoint.com/dl"}))
    env.graph.on("GET", "contoso.sharepoint.com/dl", lambda r: httpx.Response(200, content=data))


class TestScopeShaping:
    def test_tools_for_scopes(self) -> None:
        assert tools_for(frozenset({"User.Read"})) == []
        assert tools_for(frozenset({"User.Read", "Files.Read.All"})) == [
            "sharepoint_search", "sharepoint_list_folder", "sharepoint_read_file",
        ]
        assert "sharepoint_get_site" in tools_for(frozenset({"Sites.Read.All"}))

    def test_missing_scope_refuses_and_audits(self, env: Env) -> None:
        env.tokens.scopes = frozenset({"User.Read", "Files.Read.All"})
        with pytest.raises(AuthError) as err:
            env.service.get_site()
        assert err.value.code == "scope_not_granted"
        assert env.graph.requests == []
        assert env.audit_lines()[-1]["outcome"] == "denied"

    def test_not_signed_in(self, env: Env) -> None:
        env.tokens.user = None
        with pytest.raises(AuthError) as err:
            env.service.search("budget")
        assert err.value.code == "not_signed_in"


class TestSearch:
    def test_returns_bounded_metadata_and_handles_only(self, env: Env) -> None:
        env.graph.on("POST", SEARCH, fixture("search_driveitem.json"))
        out = env.service.search("budget")
        body = json.loads(env.graph.requests[0].content)
        assert body["requests"][0]["entityTypes"] == ["driveItem"]
        assert body["requests"][0]["size"] == 10 and body["requests"][0]["from"] == 0
        assert [r["name"] for r in out["results"]] == ["Budget FY27.docx", "HR plan.pdf"]
        assert out["next_page_from"] == 10 and out["total_estimate"] == 3
        first = out["results"][0]
        assert first["handle"].startswith("spf_") and first["site"].startswith("sps_")
        assert first["readable"] is True
        _no_leaks(out)

    def test_handles_open_to_the_right_item(self, env: Env) -> None:
        env.graph.on("POST", SEARCH, fixture("search_driveitem.json"))
        handle = env.service.search("budget")["results"][0]["handle"]
        assert env.codec().open(handle, "file")["i"] == "01FINBUDGET"

    def test_site_scope_adds_path_filter_and_drops_other_sites(self, env: Env) -> None:
        env.graph.on("GET", f"/v1.0/sites/{FIN_SITE}", {"id": FIN_SITE,
                     "webUrl": "https://contoso.sharepoint.com/sites/Finance"})
        env.graph.on("POST", SEARCH, fixture("search_driveitem.json"))
        site = env.codec().seal("site", {"s": FIN_SITE})
        out = env.service.search("budget", site=site)
        query = json.loads(env.graph.requests[1].content)["requests"][0]["query"]["queryString"]
        assert query == '(budget) path:"https://contoso.sharepoint.com/sites/Finance"'
        assert [r["name"] for r in out["results"]] == ["Budget FY27.docx"]

    def test_list_item_403_falls_back_to_files(self, env: Env) -> None:
        calls = []

        def handler(request: httpx.Request) -> httpx.Response:
            types = json.loads(request.content)["requests"][0]["entityTypes"]
            calls.append(types)
            if "listItem" in types:
                return httpx.Response(403, json={"error": {"code": "Forbidden"}})
            return httpx.Response(200, json=fixture("search_driveitem.json"))

        env.graph.on("POST", SEARCH, handler)
        out = env.service.search("budget", include_list_items=True)
        assert calls == [["driveItem", "listItem"], ["driveItem"]]
        assert out["warnings"] == ["list_items_forbidden_fell_back_to_files"]

    def test_throttle_surfaces_retry_after(self, env: Env) -> None:
        env.graph.on("POST", SEARCH, lambda r: httpx.Response(429, headers={"Retry-After": "5"}))
        with pytest.raises(GraphError) as err:
            env.service.search("budget")
        assert err.value.retry_after_seconds == 5
        assert env.audit_lines()[-1]["error"] == "throttled"

    @pytest.mark.parametrize(
        "kwargs",
        [{"query": ""}, {"query": "x" * 257}, {"query": "a", "size": 26}, {"query": "a", "size": 0},
         {"query": "a", "page_from": 201}, {"query": "a", "page_from": -1}],
    )
    def test_bounds(self, env: Env, kwargs: dict) -> None:
        with pytest.raises(ConnectorError) as err:
            env.service.search(**kwargs)
        assert err.value.code == "bad_request"
        assert env.graph.requests == []

    def test_user_b_cannot_use_user_a_site_handle(self, env: Env) -> None:
        site = env.codec(USER_A).seal("site", {"s": FIN_SITE})
        env.tokens.user = USER_B
        with pytest.raises(HandleError):
            env.service.search("budget", site=site)
        assert env.graph.requests == []


class TestSiteAndFolder:
    def test_get_site_by_url(self, env: Env) -> None:
        env.graph.on("GET", "/v1.0/sites/contoso.sharepoint.com:/sites/Finance",
                     {"id": FIN_SITE, "displayName": "Finance", "webUrl": "https://contoso.sharepoint.com/sites/Finance"})
        env.graph.on("GET", f"/v1.0/sites/{FIN_SITE}/drives", {"value": [
            {"id": "b!FinanceDrive01", "name": "Documents", "driveType": "documentLibrary"}]})
        out = env.service.get_site("https://contoso.sharepoint.com/sites/Finance")
        assert out["name"] == "Finance"
        assert out["libraries"][0]["root_folder"].startswith("spd_")
        _no_leaks(out)

    @pytest.mark.parametrize("site", ["https://evil.example/sites/x", "http://contoso.sharepoint.com/sites/x",
                                      "contoso.sharepoint.com/sites/../me", "https://contoso.sharepoint.com/_api/x"])
    def test_get_site_rejects_other_hosts_and_paths(self, env: Env, site: str) -> None:
        with pytest.raises(ConnectorError) as err:
            env.service.get_site(site)
        assert err.value.code == "bad_request"
        assert env.graph.requests == []

    def test_list_folder_pages_with_opaque_cursor(self, env: Env) -> None:
        folder = env.codec().seal("folder", {"d": "b!FinanceDrive01", "i": "root"})
        env.graph.on("GET", "/v1.0/drives/b!FinanceDrive01/root/children", fixture("children_page1.json"))
        out = env.service.list_folder(folder)
        assert [i["kind"] for i in out["items"]] == ["folder", "file"]
        assert out["next_cursor"].startswith("spc_")
        _no_leaks(out)
        env.service.list_folder(folder, cursor=out["next_cursor"])
        assert env.graph.requests[-1].url.params["$skiptoken"] == "UGFnZWQ9VFJVRSZwX0lEPTI"

    def test_cursor_is_bound_to_its_folder(self, env: Env) -> None:
        codec = env.codec()
        cursor = codec.seal("cursor", {"d": "b!FinanceDrive01", "i": "01OTHER", "k": "abc"})
        folder = codec.seal("folder", {"d": "b!FinanceDrive01", "i": "root"})
        with pytest.raises(ConnectorError) as err:
            env.service.list_folder(folder, cursor=cursor)
        assert err.value.code == "bad_request"


class TestReadFile:
    def test_text_file(self, env: Env) -> None:
        _meta(env, "notes.txt", 11)
        _content(env, b"hello world")
        out = env.service.read_file(_file_handle(env))
        assert out["text"] == "hello world" and out["format"] == "text"
        assert out["content_trust"] == "untrusted_document_text"
        assert out["changed_since_handle"] is False
        _no_leaks({k: v for k, v in out.items() if k != "text"})

    def test_ctag_change_is_reported(self, env: Env) -> None:
        _meta(env, "notes.txt", 2, ctag="c2")
        _content(env, b"hi")
        assert env.service.read_file(_file_handle(env, ctag="c1"))["changed_since_handle"] is True

    def test_docx(self, env: Env) -> None:
        data = make_docx("Quarterly plan", "Second line")
        _meta(env, "plan.docx", len(data))
        _content(env, data)
        out = env.service.read_file(_file_handle(env))
        assert out["text"] == "Quarterly plan\nSecond line\nTotal | 42"

    def test_pdf_pages_and_truncation(self, env: Env) -> None:
        data = make_pdf(["Alpha page", "Bravo page", "Charlie page"])
        _meta(env, "doc.pdf", len(data))
        _content(env, data)
        full = env.service.read_file(_file_handle(env))
        assert "Alpha page" in full["text"] and "Charlie page" in full["text"]
        assert full["pages"] == {"start": 1, "read": 3, "total": 3, "next_start_page": None}
        part = env.service.read_file(_file_handle(env), max_chars=25)
        assert part["truncated"] is True and part["pages"]["next_start_page"] == 2
        rest = env.service.read_file(_file_handle(env), start_page=3)
        assert rest["text"].startswith("[page 3]")

    def test_unsupported_format_never_downloads(self, env: Env) -> None:
        _meta(env, "model.xlsx", 10)
        with pytest.raises(ConnectorError) as err:
            env.service.read_file(_file_handle(env))
        assert err.value.code == "unsupported_format"
        assert not any(r.url.path.endswith("/content") for r in env.graph.requests)

    def test_size_cap_checked_before_download(self, env: Env) -> None:
        _meta(env, "big.txt", env.settings.limits.read_max_bytes + 1)
        with pytest.raises(ConnectorError) as err:
            env.service.read_file(_file_handle(env))
        assert err.value.code == "file_too_large"
        assert not any(r.url.path.endswith("/content") for r in env.graph.requests)

    def test_encrypted_office_file(self, env: Env) -> None:
        data = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64
        _meta(env, "locked.docx", len(data))
        _content(env, data)
        with pytest.raises(ExtractionError) as err:
            env.service.read_file(_file_handle(env))
        assert err.value.code == "encrypted_file"

    def test_revoked_access_is_typed_404(self, env: Env) -> None:
        with pytest.raises(GraphError) as err:
            env.service.read_file(_file_handle(env))
        assert err.value.code == "not_found"

    def test_folder_handle_rejected_as_file(self, env: Env) -> None:
        folder = env.codec().seal("folder", {"d": "b!FinanceDrive01", "i": "root"})
        with pytest.raises(HandleError) as err:
            env.service.read_file(folder)
        assert err.value.code == "wrong_handle_kind"


class TestAuditAndSignOut:
    def test_audit_has_hashes_and_correlation_only(self, env: Env) -> None:
        env.graph.on("POST", SEARCH, lambda r: httpx.Response(
            200, json=fixture("search_driveitem.json"), headers={"request-id": "graph-rid-1"}))
        env.service.search("confidential merger")
        line = env.audit_lines()[-1]
        assert line["event"] == "sharepoint_search" and line["outcome"] == "allowed"
        assert line["graph_request_id"] == "graph-rid-1" and line["client_request_id"]
        assert line["counts"] == {"results": 2, "skipped": 1}
        raw = json.dumps(line)
        for needle in ["confidential", "merger", USER_A, "Budget", "token-"]:
            assert needle not in raw
        assert len(line["actor"]) == 24

    def test_sign_out_invalidates_handles(self, env: Env) -> None:
        env.graph.on("POST", SEARCH, fixture("search_driveitem.json"))
        handle = env.service.search("budget")["results"][0]["handle"]
        env.service.sign_out()
        with pytest.raises(HandleError):
            env.service.read_file(handle)

    def test_status_checks_graph_identity(self, env: Env) -> None:
        env.graph.on("GET", "/v1.0/me", {"id": USER_A})
        status = env.service.connection_status()
        assert status["signed_in"] and status["graph_check"] == "ok"
        assert "sharepoint_read_file" in status["available_tools"]
        assert env.graph.requests[0].url.params["$select"] == "id"

    def test_status_signed_out(self, env: Env) -> None:
        env.tokens.user = None
        assert env.service.connection_status()["available_tools"] == ["get_connection_status"]
