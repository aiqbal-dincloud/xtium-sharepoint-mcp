"""Launch the connector over stdio like an MCP host does, list tools, call one, print the result.

Usage: uv run python scripts/smoke_stdio.py [tool_name] [json_args]
Reads XTIUM_SP_TENANT_ID / XTIUM_SP_CLIENT_ID (and optional XTIUM_SP_DATA_DIR) from the environment.
"""

from __future__ import annotations

import json
import os
import sys
import time

import anyio
from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client


async def main() -> None:
    tool = sys.argv[1] if len(sys.argv) > 1 else "get_connection_status"
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    env = {k: v for k, v in os.environ.items() if k.startswith(("XTIUM_SP_", "LOCALAPPDATA", "SYSTEMROOT", "PATH"))}
    params = StdioServerParameters(command=sys.executable, args=["-m", "xtium_sharepoint_mcp.cli", "serve"], env=env)
    async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
        init = await session.initialize()
        print(f"server: {init.server_info.name} {init.server_info.version}")
        tools = await session.list_tools()
        print("tools:", sorted(t.name for t in tools.tools))
        started = time.perf_counter()
        result = await session.call_tool(tool, args)
        elapsed = int((time.perf_counter() - started) * 1000)
        print(f"{tool}: is_error={result.is_error} in {elapsed} ms")
        for block in result.content:
            print(getattr(block, "text", block))


if __name__ == "__main__":
    anyio.run(main)
