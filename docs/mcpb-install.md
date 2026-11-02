# Install as a Claude Desktop extension (MCPB)

One-click local install so users enter **tenant ID** and **client ID** in Claude’s UI instead of editing `claude_desktop_config.json`.

This keeps the AIDATA-3958 model: process on the device, MSAL public-client sign-in, Graph direct. It is **not** a remote Connectors URL.

## Prerequisites

- Claude Desktop (macOS or Windows), signed in to claude.ai
- Organization allows **Extensions** (if Team/Enterprise has disabled them, ask an admin)
- [`uv`](https://docs.astral.sh/uv/) on `PATH` (Claude launches `uv run …` from the bundle)
- Customer Entra **public** app in the target tenant (no secret): redirect `http://localhost`, delegated `User.Read`, `Sites.Read.All`, `Files.Read.All`, plus OIDC scopes

## Build the `.mcpb` (maintainers)

From the repo root (with Node/`mcpb` CLI available):

```powershell
npm install -g @anthropic-ai/mcpb
.\scripts\pack_mcpb.ps1
```

Output: `dist/xtium-sharepoint-0.1.0.mcpb`

Validate:

```powershell
mcpb validate .\manifest.json
mcpb info .\dist\xtium-sharepoint-0.1.0.mcpb
```

## Install (end user)

1. Get the `.mcpb` from your internal share / release (private distribution; directory listing is not used for MCPB).
2. Either:
   - Double-click the file, or
   - Drag it into Claude Desktop, or
   - **Settings → Extensions → Advanced settings → Install Extension…** and choose the file
3. Review permissions. Enter:
   - **Entra tenant ID**
   - **Entra application (client) ID**
4. Finish install. Restart Claude if tools do not appear.
5. In chat: ask Claude to run **`sharepoint_sign_in`**, complete browser SSO for that tenant, then search/list/read.

## After install

| Action | How |
| --- | --- |
| Change tenant/client | Extension settings in Claude (re-enter IDs) or reinstall |
| Sign out | `sharepoint_sign_out` or CLI `uv run xtium-sharepoint-mcp sign-out` |
| Wipe cache | Delete `%LOCALAPPDATA%\xtium-sharepoint-mcp` (Windows) |

## Still supported: developer JSON

For lab debugging without packing:

```json
{
  "mcpServers": {
    "xtium-sharepoint": {
      "command": "C:\\\\path\\\\to\\\\xtium-sharepoint-mcp\\\\.venv\\\\Scripts\\\\xtium-sharepoint-mcp.exe",
      "args": ["serve"],
      "env": {
        "XTIUM_SP_TENANT_ID": "<tenant id>",
        "XTIUM_SP_CLIENT_ID": "<client id>"
      }
    }
  }
}
```

Prefer the MCPB path for demos and pilots after Jordan’s packaging feedback.

## Out of scope here

- Remote **Customize → Connectors → +Add** (Anthropic cloud → URL)
- Agent Manager tool-gateway proxy (later track)
- Marketplace directory listing (would be a plugin wrap later)
