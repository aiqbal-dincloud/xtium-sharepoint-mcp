# WI-25 test packet: local client, transport and customer public app

Ticket: [AIDATA-3970](https://xtium.atlassian.net/browse/AIDATA-3970). Design: [ADR 0001](adr/0001-local-transport-and-oauth.md).
No secrets, tokens, UPNs or file content belong in this file or in Jira.

## Offline evidence (no tenant needed)

| Check | Result | Where |
| --- | --- | --- |
| Unit and contract tests | 111+ passed (2026-11-02 re-run green) | `uv run pytest` |
| Empty-cache `identity()` skips MSAL / OIDC | Pass | `test_identity_skips_msal_when_cache_empty` |
| Wrong tenant, client, issuer and Graph audience rejected; app-only token refused | Pass | `tests/test_config_auth.py` |
| Cached user cannot change silently; two cached accounts refused | Pass | `tests/test_config_auth.py` |
| Revoked refresh token maps to `reauth_required` | Pass | `tests/test_config_auth.py` |
| Handles bound to tenant, user and kind; forged or URL-shaped IDs refused; sign-out invalidates them | Pass | `tests/test_handles_graph.py`, `tests/test_service.py` |
| Search output has no `summary`, excerpt, download URL, raw ID or token | Pass | `tests/test_service.py` |
| 401 refresh-once; 403, 404, 429 and 5xx typed with request IDs; Retry-After capped; no auto-retry | Pass | `tests/test_handles_graph.py` |
| Download redirect limited to SharePoint hosts, and the token is never sent there | Pass | `tests/test_handles_graph.py` |
| Size and format refused before download; encrypted Office files typed | Pass | `tests/test_service.py` |
| Tool list follows granted scopes and sign-in state | Pass | `tests/test_server.py` |
| Audit lines hold hashed identifiers and correlation IDs only | Pass | `tests/test_service.py` |
| Real stdio round trip (`initialize`, `tools/list`, `get_connection_status`) | Pass, ~581 ms (2026-11-02; signed-out tools only with placeholder tenant/client) | `scripts/smoke_stdio.py` |

## Live evidence (f3zj)

| Check | Result |
| --- | --- |
| f3zj has SharePoint licences | Pass — `DEVELOPERPACK_V2_E5` (25 enabled, 1 consumed) |
| Two licensed users with different site ACLs, plus one item only User A can open | Pass — Aaqib searched and read `wi31-a-only.docx`. Usama’s search returned nothing, and User A’s handle returned `invalid_handle`. |
| Non-production public-client test app created, no secret | Pass — app `f4368549-8052-4a54-86fc-ce8635e0af70`, tenant-wide consent for User.Read / Sites.Read.All / Files.Read.All (+ OIDC). IDs in gitignored `local/test_app.json` |
| Claude Desktop: real browser sign-in and one tool round trip | Partial — signed in on consumer Claude. Search, root site, and Mark8 folder listing returned metadata only. Read of a 0-byte DOCX correctly returned `corrupt_file`. A non-empty file read was skipped (the only search hit is 0 bytes). Audit log for 2026-10-30 passed: hashed actor and resources, Graph request ids, no tokens or file text. Log is under the MSIX local cache, not `%LOCALAPPDATA%\xtium-sharepoint-mcp`. |
| Wrong tenant: configure another tenant ID, sign in, and get `wrong_tenant` | Pass — fake tenant failed closed at startup (`AADSTS90002`). Real tenant restored. |
| Wrong client: configure an unregistered client ID and get `wrong_client` (AADSTS700016) | Pass — Entra `AADSTS700016`, application not found. Real client id restored. |
| Wrong audience: covered offline; live, the connector only ever requests Graph scopes | Offline pass |
| User A vs User B search and read differ by ACL | Pass — see `docs/wi-31-pilot-packet.md` (2026-10-30 Aaqib/Usama on f3zj) |
| Claude Code | Not tested (host MCP config + interactive sign-in still required) |
| Cursor | Not tested (host MCP config + interactive sign-in still required) |
| Private GHE repo | Landed 2026-11-02 — `https://xtium.ghe.com/XTIUM/xtium-sharepoint-mcp` |
