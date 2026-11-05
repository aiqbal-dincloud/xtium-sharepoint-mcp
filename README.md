# xtium-sharepoint-mcp

A local, read-only SharePoint connector for MCP desktop clients. It runs on the user's device, signs in
as that user with the customer's own Entra public-client app, and calls Microsoft Graph directly. No XTIUM
service sees tokens or file content.

This is the pilot for [AIDATA-3958](https://xtium.atlassian.net/browse/AIDATA-3958). The design is in
[ADR 0001](docs/adr/0001-local-transport-and-oauth.md).

## Tools

| Tool | Needs | What it returns |
| --- | --- | --- |
| `get_connection_status` | nothing | Sign-in state, granted scopes, available tools and a Graph identity check |
| `sharepoint_sign_in` | nothing | Opens the browser for Microsoft sign-in |
| `sharepoint_sign_out` | nothing | Deletes the local session and invalidates all handles |
| `sharepoint_search` | `Files.Read.All` or `Sites.Read.All` | Metadata and opaque handles, with no excerpts or download links |
| `sharepoint_get_site` | `Sites.Read.All` | Site details and document libraries, with root folder handles |
| `sharepoint_list_folder` | `Files.Read.All` or `Sites.Read.All` | Folder contents, with a paging cursor |
| `sharepoint_read_file` | `Files.Read.All` or `Sites.Read.All` | Text from PDF, DOCX or plain-text files, within size caps |

## Setup

### 1. Register the app (tenant admin, once)

Create an Entra app registration in the customer tenant with these settings:

- Supported account types: **this organisational directory only**.
- Platform: **Mobile and desktop applications**, with redirect URI `http://localhost`.
- **Allow public client flows: Yes**. Do not create a secret or certificate.
- API permissions: Microsoft Graph delegated `User.Read`, `Sites.Read.All`, `Files.Read.All`,
  `offline_access`, `openid` and `profile`. Grant admin consent if users cannot consent themselves.

Record the **tenant ID** and **application (client) ID**. Neither is a secret, but keep them out of
tickets anyway; the customer owns them.

For a QA tenant, `scripts/provision_test_app.py --tenant <id> --create --grant-consent` does the
same steps and writes the IDs to `local/test_app.json` (gitignored).

### 2. Install

```powershell
uv sync
```

### 3. Add it to Claude Desktop

**Preferred: Claude Desktop extension (MCPB).** Users enter the tenant and client IDs in Claude’s UI; no JSON editing.

1. Get `dist/xtium-sharepoint-0.1.0.mcpb` from this repo (or build it with `.\scripts\pack_mcpb.ps1`).
2. Install [`uv`](https://docs.astral.sh/uv/) and make sure it is on the user `PATH` (Claude runs `uv run` from the bundle). Then fully quit and reopen Claude.
3. In **normal Claude Desktop** (signed in with Claude.ai, not the Claude 3P / Agent Manager profile), open **Settings → Extensions → Advanced settings → Install Extension…** and pick the `.mcpb`. You can also double-click it or drag it into Claude.
4. Enter the **Entra tenant ID** and **application (client) ID**, then finish.

If **Install** is greyed out, your organization restricts extensions. An owner must add the file under **Organization settings → Connectors → Desktop → Add custom extension**, or turn the extension allowlist off. Searching the Directory will not find it: the plugin is not published, and this is a local MCPB, not a remote connector.

The `.mcpb` is unsigned, so Claude shows an “unsigned” warning. Directory listing (a **plugin** submission, not the remote MCP connector form) is described in [docs/directory-submission.md](docs/directory-submission.md). Full install notes are in [docs/mcpb-install.md](docs/mcpb-install.md).

**Developer fallback** — edit Claude’s `mcpServers` JSON. For the MSIX build the file is `%LOCALAPPDATA%\Claude-Data\claude_pzs8sxrjxfjjc\claude_desktop_config.json` (the `Roaming\Claude` folder is not used). The Claude 3P profile has its own file under `%LOCALAPPDATA%\Claude-3p\`; do not use it for this connector.

```json
{
  "mcpServers": {
    "xtium-sharepoint": {
      "command": "C:\\path\\to\\xtium-sharepoint-mcp\\.venv\\Scripts\\xtium-sharepoint-mcp.exe",
      "args": ["serve"],
      "env": {
        "XTIUM_SP_TENANT_ID": "<tenant id>",
        "XTIUM_SP_CLIENT_ID": "<client id>"
      }
    }
  }
}
```

Claude Code and Cursor take the same `command`, `args` and `env`. Each host is verified separately; see the [test packet](docs/wi-25-test-packet.md).

### 4. Sign in

Ask the assistant to run `sharepoint_sign_in`, or run this first (developer install):

```powershell
.venv\Scripts\xtium-sharepoint-mcp.exe sign-in
```

The search, site, folder and read tools appear only when the session holds the Graph scopes. If you signed in during a chat and the new tools do not show up, quit Claude completely (including the tray icon) and open a new chat.

### Troubleshooting

- **`sharepoint_*` tools missing:** run `get_connection_status` in the same chat. If `signed_in` is true and `available_tools` lists them, start a new chat after a full quit. If Claude only offers `sharepoint_search`, `sharepoint_folder_search` or `read_resource`, those belong to Claude’s own Microsoft 365 connector, not this one.
- **Wrong account or documents:** Claude’s own Microsoft 365 connector uses your Claude sign-in. Only the XTIUM SharePoint (local) tools use the account you signed in with here, and `get_connection_status` does not return your user name.
- **“Unable to connect to extension server”:** if a previous sign-in is cached, startup refreshes the token before answering Claude, and Claude gives up after 120 seconds. Quit Claude, delete `token_cache.bin` in the data directory, reopen, and run `sharepoint_sign_in` again.
- **`uv` not found:** put `%USERPROFILE%\.local\bin` on the user `PATH`, then fully quit and reopen Claude.

## Optional settings

| Variable | Default | Purpose |
| --- | --- | --- |
| `XTIUM_SP_SCOPES` | `User.Read Sites.Read.All Files.Read.All` | Use a smaller set, for example `Files.Read.All` |
| `XTIUM_SP_DATA_DIR` | `%LOCALAPPDATA%\xtium-sharepoint-mcp` | Token cache, keys and audit log. For the Microsoft Store (MSIX) Claude, Windows redirects this to `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Local\xtium-sharepoint-mcp`. On macOS the default is `~/.local/share/xtium-sharepoint-mcp` |
| `XTIUM_SP_AUDIT_RETENTION_DAYS` | `30` | Local audit retention, from 1 to 365 days |

## Support

- **Disconnect**: run `sharepoint_sign_out` or `xtium-sharepoint-mcp sign-out`. To wipe everything,
  delete the data directory (see `XTIUM_SP_DATA_DIR` above; for the Microsoft Store Claude it is under `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Local\`).
- **Revoke a user**: in Entra, revoke the user's sessions or remove the app's consent. The connector asks
  them to sign in again at the next token refresh (within about an hour).
- **Correlate an error**: each error and each audit line carries the Graph `request-id` and a
  `client-request-id`. The audit log is at `%LOCALAPPDATA%\xtium-sharepoint-mcp\audit\audit-YYYY-MM-DD.jsonl`
  and contains hashed identifiers only.

## Privacy Policy

This connector is designed so **XTIUM does not host SharePoint content or Microsoft tokens**.

| Topic | Practice |
| --- | --- |
| **Data collection** | On the user’s device only: MSAL token cache (OS-encrypted), opaque file handles, and a local audit log of hashed identifiers and Graph correlation IDs. No document text is written to disk by the connector. |
| **Usage and storage** | Graph calls run from the user’s machine to Microsoft. Extracted file text is returned once in the MCP response and is not retained by the connector. Default data directory: `%LOCALAPPDATA%\xtium-sharepoint-mcp` (Windows) or the configured `XTIUM_SP_DATA_DIR`. |
| **Third-party sharing** | Microsoft Graph / Entra receive authentication and API traffic as with any Graph client. Anthropic/Claude receives only what the local MCP host sends in the conversation (tool results). XTIUM cloud services are not in the token or download path for this local connector. |
| **Data retention** | Audit lines default to 30 days (`XTIUM_SP_AUDIT_RETENTION_DAYS`). Sign-out deletes the token cache and invalidates handles. Users may delete the data directory to wipe local state. |
| **Contact** | XTIUM SharePoint MCP maintainers — `aaqib.iqbal@xtium.com` (pilot). |

Customer Entra admins remain responsible for app consent, Conditional Access, and user revocation in their tenant.

## Development

```powershell
uv run pytest
uv run ruff check src tests
uv run python scripts/smoke_stdio.py                    # stdio round trip as a host would run it
uv run python scripts/smoke_stdio.py sharepoint_search '{"query":"budget"}'
.\scripts\pack_mcpb.ps1                                 # Claude Desktop / plugin MCPB
```
