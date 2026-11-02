from __future__ import annotations

import json
from typing import Any

import anyio
import pytest
from mcp.server.mcpserver.exceptions import ToolError

from xtium_sharepoint_mcp.server import build_server

from .conftest import Env, fixture


class FakeAuth:
    def __init__(self, env: Env):
        self.env = env

    def identity(self):
        return self.env.tokens.identity()

    def sign_in(self):
        self.env.tokens.user = "0a0a0a0a-0000-0000-0000-00000000000a"
        return self.env.tokens.identity()

    def sign_out(self) -> None:
        self.env.tokens.user = None


def _names(server) -> set[str]:
    return {t.name for t in anyio.run(server.list_tools)}


def _call(server, name: str, args: dict[str, Any] | None = None):
    return anyio.run(lambda: server.call_tool(name, args or {}))


def test_signed_out_exposes_only_lifecycle_tools(env: Env) -> None:
    env.tokens.user = None
    server, _ = build_server(FakeAuth(env), env.service)
    assert _names(server) == {"get_connection_status", "sharepoint_sign_in", "sharepoint_sign_out"}


def test_files_only_scope_hides_site_tool(env: Env) -> None:
    env.tokens.scopes = frozenset({"User.Read", "Files.Read.All"})
    server, _ = build_server(FakeAuth(env), env.service)
    names = _names(server)
    assert "sharepoint_search" in names and "sharepoint_get_site" not in names


def test_sign_in_and_out_reshape_tools(env: Env) -> None:
    env.tokens.user = None
    auth = FakeAuth(env)
    server, registry = build_server(auth, env.service)
    assert registry.sync(None) is False
    auth.sign_in()
    registry.sync(env.tokens.scopes)
    assert "sharepoint_read_file" in _names(server)
    auth.sign_out()
    registry.sync(None)
    assert "sharepoint_read_file" not in _names(server)


def test_tool_round_trip_and_typed_error(env: Env) -> None:
    env.graph.on("POST", "/v1.0/search/query", fixture("search_driveitem.json"))
    server, _ = build_server(FakeAuth(env), env.service)
    result = _call(server, "sharepoint_search", {"query": "budget"})
    text = json.dumps(result, default=str)
    assert "Budget FY27.docx" in text and "SECRET" not in text

    with pytest.raises(ToolError) as err:
        _call(server, "sharepoint_read_file", {"file": "not-a-handle"})
    message = str(err.value)
    assert json.loads(message[message.index("{"):])["error"] == "invalid_handle"
