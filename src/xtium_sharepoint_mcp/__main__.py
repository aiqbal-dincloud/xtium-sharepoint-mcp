"""``python -m xtium_sharepoint_mcp`` starts the MCP stdio server (MCPB entry)."""

from __future__ import annotations

from .cli import main

raise SystemExit(main(["serve"]))
