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

**Preferred (pilot packaging):** Claude **plugin** + embedded MCPB so users enter tenant/client IDs in Claude’s UI:

1. Build: `.\scripts\pack_mcpb.ps1` → `dist/xtium-sharepoint-0.1.0.mcpb`
2. Install MCPB: **Settings → Extensions → Install Extension…** (or double-click / drag), **or** install the plugin once it is listed / sideloaded
3. Enter **Entra tenant ID** and **application (client) ID**, then finish

Directory listing uses a **plugin** submission (not the remote MCP connector form): [docs/directory-submission.md](docs/directory-submission.md).  
Install details: [docs/mcpb-install.md](docs/mcpb-install.md).


**Developer fallback** — edit Claude’s `mcpServers` JSON (MSIX path is under `%LOCALAPPDATA%\Claude-Data\…\claude_desktop_config.json`):

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

Ask the assistant to run `sharepoint_sign_in`, or run this first:

```powershell
.venv\Scripts\xtium-sharepoint-mcp.exe sign-in
```

If the client does not refresh its tool list after sign-in, restart it.

## Optional settings

| Variable | Default | Purpose |
| --- | --- | --- |
| `XTIUM_SP_SCOPES` | `User.Read Sites.Read.All Files.Read.All` | Use a smaller set, for example `Files.Read.All` |
| `XTIUM_SP_DATA_DIR` | `%LOCALAPPDATA%\xtium-sharepoint-mcp` | Token cache, keys and audit log |
| `XTIUM_SP_AUDIT_RETENTION_DAYS` | `30` | Local audit retention, from 1 to 365 days |

## Support

- **Disconnect**: run `sharepoint_sign_out` or `xtium-sharepoint-mcp sign-out`. To wipe everything,
  delete `%LOCALAPPDATA%\xtium-sharepoint-mcp`.
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
