# Install as a Claude Desktop extension (MCPB)

One-click local install so users enter **tenant ID** and **client ID** in Claude’s UI instead of editing `claude_desktop_config.json`.

This keeps the AIDATA-3958 model: process on the device, MSAL public-client sign-in, Graph direct. It is **not** a remote Connectors URL.

## Prerequisites

- Claude Desktop (macOS or Windows), signed in to claude.ai
- Organization allows **Extensions**. If **Install** is greyed out ("Contact your organization owner"), the desktop extension allowlist is on. An owner must add the file under **Organization settings → Connectors → Desktop → Add custom extension**, or turn the allowlist off. Browse Directory will not list it.
- Use **normal Claude Desktop** (Claude.ai sign-in). The Claude 3P / Agent Manager profile is a different app profile and will not show this extension.
- [`uv`](https://docs.astral.sh/uv/) on the user `PATH` (Claude launches `uv run …` from the bundle). Typical Windows location: `%USERPROFILE%\.local\bin`. Fully quit and reopen Claude after changing `PATH`.
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
4. Finish install. The file is unsigned, so Claude warns about that. Restart Claude if tools do not appear.
5. In chat: ask Claude to run **`sharepoint_sign_in`**, complete browser SSO for that tenant, then search/list/read.
6. If `sharepoint_list_folder` and the other data tools are missing afterwards, run `get_connection_status`. When it shows `signed_in: true`, quit Claude fully (tray included) and start a new chat. Names like `sharepoint_folder_search` belong to Claude's Microsoft 365 connector, not this extension.

## Troubleshooting

| Symptom | Cause and fix |
| --- | --- |
| "Unable to connect to extension server" | A cached sign-in makes startup refresh the token before Claude gets its first reply, and Claude times out after 120 seconds. Quit Claude, delete `token_cache.bin` in the data directory, reopen, and sign in again. |
| Claude names the wrong user or shows `@xtium.com` files | That is Claude's own Microsoft 365 connector (your Claude sign-in). This extension returns only Graph data for the account used in `sharepoint_sign_in`, and `get_connection_status` does not return a user name. |
| `ERR_CONNECTION_REFUSED` on a "Sign in with your organization" screen | You are in the Claude 3P profile pointed at a local gateway (`127.0.0.1:8001`). Choose "Or sign in with Claude.ai" for normal Claude. |

## After install

| Action | How |
| --- | --- |
| Change tenant/client | Extension settings in Claude (re-enter IDs) or reinstall |
| Sign out | `sharepoint_sign_out` or CLI `uv run xtium-sharepoint-mcp sign-out` |
| Wipe cache | Delete the data directory: `%LOCALAPPDATA%\xtium-sharepoint-mcp` normally, or `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Local\xtium-sharepoint-mcp` for the Microsoft Store Claude (Windows); `~/.local/share/xtium-sharepoint-mcp` (macOS) |

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
