# ADR 0001: Local stdio transport and connector-run delegated OAuth

- Status: Proposed (WI-25, [AIDATA-3970](https://xtium.atlassian.net/browse/AIDATA-3970))
- Epic: [AIDATA-3958](https://xtium.atlassian.net/browse/AIDATA-3958)
- Source: Bundled Inference and SharePoint Implementation Plan v1.1, section 4.4, WI-25 to WI-31
- Decisions covered: DEC-08, DEC-09, DEC-10, DEC-12 (pilot scope only)

## Context

The pilot needs a SharePoint connector that runs on the user's own device, signs in as that user with a
customer-created public-client app, and calls Microsoft Graph directly. There must be no XTIUM URL, no
remote MCP resource URI and no XTIUM-hosted token store.

The MCP authorization specification applies to HTTP transports. For stdio servers it says implementations
should not follow it and should get credentials from the environment instead. So with a local stdio server,
the MCP host does not run OAuth on the connector's behalf. The connector has to sign the user in itself.

## Decision

### Transport and first client (DEC-08)

- The connector is a local **stdio** MCP server (`xtium-sharepoint-mcp serve`), launched by the MCP host.
- The first client to prove is **Claude Desktop**. Claude Code and Cursor use the same stdio config shape,
  but each is recorded separately in the test packet and is not claimed to work until it has been tested.
- The admin enters the tenant ID and client ID manually, as `XTIUM_SP_TENANT_ID` and `XTIUM_SP_CLIENT_ID`
  in the host's server config. Multi-tenant aliases (`common`, `organizations`) are rejected.

### Sign-in flow (DEC-08, DEC-09)

- MSAL Python `PublicClientApplication`, using the authorization code flow with PKCE in the **system
  browser**, with a loopback redirect to `http://localhost` (any port, as Entra allows for public
  clients). The `sharepoint_sign_in` tool or the `sign-in` CLI command starts it.
- The connector validates each token result before using it:
  - The ID token's `tid` and issuer match the configured tenant, and its `aud` matches the configured client.
  - The access token's `aud` is Microsoft Graph, its `tid` is the tenant, and its `appid`/`azp` is the client.
  - The token is delegated (`scp` is present). App-only tokens are refused.
  - The `oid` agrees across both tokens, and the cached user cannot change silently.
- The Windows WAM broker (`msal[broker]`) is the alternative "approved OS broker" path. It is deferred
  until Security asks for it, and it would replace only the cache and sign-in UI, not the validation above.

### Scopes and consent (DEC-09)

- Delegated only: `User.Read`, `Sites.Read.All` and `Files.Read.All`, plus `offline_access`, `openid` and `profile`.
- The tool registry is shaped by the scopes actually granted:

  | Granted | Tools |
  | --- | --- |
  | none (signed out) | `get_connection_status`, `sharepoint_sign_in`, `sharepoint_sign_out` |
  | `Files.Read.All` only | adds search (files only), list folder, read file |
  | `Sites.Read.All` | adds site lookup and list-item search |

- The f3zj test app is a disposable, non-production public client with no secret, and it is named that
  way. Approving it does not approve a customer production app. The app and consent owner for a customer
  pilot is the customer's tenant admin, recorded separately in the WI-31 packet.

### Device cache, logout and revoke (DEC-10)

- The token cache is MSAL's cache in `msal-extensions` encrypted persistence: DPAPI (current user) on
  Windows, Keychain on macOS and libsecret on Linux. It lives in `%LOCALAPPDATA%\xtium-sharepoint-mcp`.
- The cache holds one account per OS profile. If more than one account is found for the tenant, the
  connector refuses (`ambiguous_account`) and does not guess.
- The handle key and audit key are 32 random bytes each, stored with the same encrypted persistence.
- **Logout** (`sharepoint_sign_out` or the `sign-out` CLI command) removes the MSAL accounts, deletes the
  cache file and rotates the handle key, so every handle issued earlier stops working.
- **Revoke** is a tenant-admin action: revoke the user's sessions, remove the consent grant, or disable
  the user. The next refresh then fails with `invalid_grant`, which the connector maps to `reauth_required`.
  An access token that has already been issued can stay valid until it expires (about 60–90 minutes),
  because the pilot does not request CAE `cp1`. This known limitation is tracked to WI-37.
- File content is never persisted. Text is extracted in memory and returned once.

### Graph contract and limits (DEC-12)

- Graph **v1.0** only. Pinned response shapes are in `tests/fixtures`.
- Search uses `POST /search/query` with `driveItem` (and `listItem` only when asked for and `Sites.Read.All`
  is granted).
  - The Graph `summary`, excerpts and `@microsoft.graph.downloadUrl` are dropped.
  - Raw IDs are sealed into AES-GCM handles bound to the tenant, the user and the handle kind.
  - Site scoping adds a KQL `path:` filter and a hard post-filter on `parentReference.siteId`.
  - If a combined `listItem` request gets a 403, the connector retries once with `driveItem` only and
    returns a warning.
- Reads revalidate the item metadata and report a cTag change. Size and format are checked before any
  download. The download follows Graph's single redirect only to a SharePoint host and never sends the
  Graph token there.
- Errors:
  - 401 triggers one forced refresh, then `unauthorized`.
  - 403 maps to `forbidden`, and 404 or 410 to `not_found`.
  - 429 and 503 map to `throttled` or `upstream_unavailable`, with `retry_after_seconds` capped at 300.
    The connector does not retry these automatically, to avoid a retry storm.
  - Every error carries the Graph `request-id` and the connector's `client-request-id`.
- Initial caps:

  | Cap | Value |
  | --- | --- |
  | Query length | 256 characters |
  | Search page size | 10 by default, 25 maximum |
  | Search offset | 200 maximum |
  | Folder page | 50 items |
  | Read size | 10 MB |
  | Read text | 20,000 characters by default, 50,000 maximum |
  | PDF pages per read | 50 |
  | DOCX uncompressed size | 100 MB |
  | HTTP timeout | 20 seconds |

- Readable formats: PDF, DOCX and plain text (`.txt .md .csv .tsv .json .xml .log .yaml .yml .html .htm`).
  Anything else returns `unsupported_format` without a download. Encrypted, IRM-protected or
  password-protected files return `encrypted_file`.
- Multi-geo: delegated search returns results across geos without the `region` parameter, which only
  app-only calls need. f3zj is single-geo, so this is untested (see the gaps below).
- Sensitivity-label gate: not in the pilot. No blocklist is configured, so the connector makes no claim
  to enforce labels (WI-37).

### Egress

The runtime calls only `login.microsoftonline.com` (sign-in), `graph.microsoft.com` and the SharePoint
download host that Graph redirects to. Graph paths are built by the connector and checked against a safe
character set. They are never taken from model input or handles.

## Owners

| Role | Owner |
| --- | --- |
| Runtime and repository (local) | Aaqib Iqbal, proposed; Connector to confirm |
| f3zj test app and consent | Aaqib Iqbal (QA tenant admin for the pilot) |
| Customer pilot app and consent | Customer tenant admin; not yet named, needed for WI-31 |
| Support (device cache removal, reconnect) | Open; needed before WI-31 |
| Security review of cache and egress posture | Open; Security under DEC-10 |

## Consequences and remaining gaps

- The connector owns sign-in and cache security. A host-side OAuth change would not help a stdio server.
- The local loopback redirect means the device needs a browser. A headless or remote host needs a
  separate device-code decision, which this ADR does not make.
- Not yet done:
  - Host-level egress enforcement (firewall or allowlist).
  - CAE and claims challenges.
  - The label gate.
  - Multi-geo testing.
  - Claude Code and Cursor runs.
  - Office formats other than DOCX (XLSX and PPTX).
- The audit log is local JSONL with 30-day retention. Support export goes through the user and no
  XTIUM telemetry exists. The retention and support owner is still open (DEC-10).
