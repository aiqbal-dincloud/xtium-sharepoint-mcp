"""WI-29/30: audit retention, Graph network errors, mid-session revoke tool-list sync."""

from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path
from unittest.mock import AsyncMock

import anyio
import httpx
import pytest
from mcp.server.mcpserver.exceptions import ToolError

from xtium_sharepoint_mcp.audit import AuditLog
from xtium_sharepoint_mcp.errors import AuthError, GraphError
from xtium_sharepoint_mcp.graph import GraphClient
from xtium_sharepoint_mcp.server import build_server

from .conftest import USER_A, Env, FakeGraph, FakeTokens


def test_audit_prune_drops_files_past_retention(tmp_path: Path) -> None:
    now = datetime(2026, 10, 29, 12, 0, tzinfo=UTC)
    audit = AuditLog(tmp_path, b"\x01" * 32, "tenant", retention_days=7, clock=lambda: now)
    keep = tmp_path / "audit-2026-10-25.jsonl"
    drop = tmp_path / "audit-2026-10-20.jsonl"
    keep.write_text("{}\n", encoding="utf-8")
    drop.write_text("{}\n", encoding="utf-8")
    audit.prune()
    assert keep.exists() and not drop.exists()


def test_sign_in_failed_records_correlation(env: Env) -> None:
    env.service.sign_in_failed("reauth_required", client_request_id="aad-corr-1")
    lines = list(env.audit_dir.glob("audit-*.jsonl"))
    assert lines
    entry = json.loads(lines[0].read_text(encoding="utf-8").strip().splitlines()[-1])
    assert entry["event"] == "sign_in"
    assert entry["outcome"] == "denied"
    assert entry["error"] == "reauth_required"
    assert entry["client_request_id"] == "aad-corr-1"
    assert "token" not in json.dumps(entry).lower()


def test_connection_status_surfaces_auth_error(env: Env) -> None:
    class BrokenTokens(FakeTokens):
        def identity(self):  # type: ignore[override]
            raise AuthError("reauth_required", "Sign in again.", client_request_id="c1")

    env.tokens = BrokenTokens(user=USER_A)
    env.service._tokens = env.tokens  # noqa: SLF001 — test seam
    status = env.service.connection_status()
    assert status["signed_in"] is False
    assert status["auth_error"] == "reauth_required"


class TestGraphNetworkErrors:
    def test_timeout(self) -> None:
        def boom(_req: httpx.Request) -> httpx.Response:
            raise httpx.ReadTimeout("slow")

        graph = FakeGraph()
        graph.on("GET", "/v1.0/me", boom)
        client = GraphClient(FakeTokens(), transport=httpx.MockTransport(graph))
        with pytest.raises(GraphError) as err:
            client.get("/me")
        assert err.value.code == "timeout"

    def test_network_error(self) -> None:
        def boom(_req: httpx.Request) -> httpx.Response:
            raise httpx.ConnectError("down")

        graph = FakeGraph()
        graph.on("GET", "/v1.0/me", boom)
        client = GraphClient(FakeTokens(), transport=httpx.MockTransport(graph))
        with pytest.raises(GraphError) as err:
            client.get("/me")
        assert err.value.code == "network_error"

    def test_410_is_not_found(self) -> None:
        graph = FakeGraph()
        graph.on(
            "GET",
            "/v1.0/me",
            lambda _r: httpx.Response(410, json={"error": {"code": "gone", "message": "SECRET"}}),
        )
        client = GraphClient(FakeTokens(), transport=httpx.MockTransport(graph))
        with pytest.raises(GraphError) as err:
            client.get("/me")
        assert err.value.code == "not_found"
        assert "SECRET" not in str(err.value)

    def test_401_then_refresh_auth_error(self) -> None:
        class RevokingTokens(FakeTokens):
            def access_token(self, force_refresh: bool = False) -> str:
                if force_refresh:
                    raise AuthError("reauth_required", "Refresh revoked.")
                return super().access_token(force_refresh)

        graph = FakeGraph()
        graph.on(
            "GET",
            "/v1.0/me",
            lambda _r: httpx.Response(401, json={"error": {"code": "InvalidAuthenticationToken"}}),
        )
        client = GraphClient(RevokingTokens(), transport=httpx.MockTransport(graph))
        with pytest.raises(AuthError) as err:
            client.get("/me")
        assert err.value.code == "reauth_required"
        assert len(graph.requests) == 1


class RecordingAuth:
    def __init__(self, env: Env):
        self.env = env

    def identity(self):
        return self.env.tokens.identity()

    def sign_in(self):
        self.env.tokens.user = USER_A
        return self.env.tokens.identity()

    def sign_out(self) -> None:
        self.env.tokens.user = None


def _names(server) -> set[str]:
    return {t.name for t in anyio.run(server.list_tools)}


def test_sign_in_out_sends_tool_list_changed(env: Env, monkeypatch: pytest.MonkeyPatch) -> None:
    env.tokens.user = None
    server, _ = build_server(RecordingAuth(env), env.service)
    notifies: list[str] = []
    monkeypatch.setattr(
        "xtium_sharepoint_mcp.server._notify",
        AsyncMock(side_effect=lambda ctx: notifies.append("changed")),
    )
    result = anyio.run(lambda: server.call_tool("sharepoint_sign_in", {}))
    assert result is not None
    assert "sharepoint_search" in _names(server)
    assert notifies == ["changed"]
    anyio.run(lambda: server.call_tool("sharepoint_sign_out", {}))
    assert "sharepoint_search" not in _names(server)
    assert notifies == ["changed", "changed"]


def test_mid_session_reauth_hides_data_tools(env: Env, monkeypatch: pytest.MonkeyPatch) -> None:
    class RevokingTokens(FakeTokens):
        def access_token(self, force_refresh: bool = False) -> str:
            raise AuthError("reauth_required", "Consent removed.")

    env.tokens = RevokingTokens(user=USER_A)
    env.service._tokens = env.tokens  # noqa: SLF001
    env.client._tokens = env.tokens  # noqa: SLF001
    server, registry = build_server(RecordingAuth(env), env.service)
    assert "sharepoint_search" in registry.active
    notifies: list[str] = []
    monkeypatch.setattr(
        "xtium_sharepoint_mcp.server._notify",
        AsyncMock(side_effect=lambda ctx: notifies.append("changed")),
    )
    with pytest.raises(ToolError) as err:
        anyio.run(lambda: server.call_tool("sharepoint_search", {"query": "x"}))
    message = str(err.value)
    assert json.loads(message[message.index("{") :])["error"] == "reauth_required"
    assert "sharepoint_search" not in registry.active
    assert "sharepoint_search" not in _names(server)
    assert notifies == ["changed"]
