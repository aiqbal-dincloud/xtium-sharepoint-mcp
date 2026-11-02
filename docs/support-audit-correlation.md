# Support: local audit correlation (WI-29)

Pilot support uses files on the **customer device only**. Nothing is sent to XTIUM by default.

## Where events live

- Directory: `{data_dir}/audit/` (default under the connector data dir)
- Files: `audit-YYYY-MM-DD.jsonl`, one JSON object per line
- Retention: `XTIUM_SP_AUDIT_RETENTION_DAYS` (default 30); older files are deleted on startup/prune

## Fields (sanitized)

| Field | Meaning |
|-------|---------|
| `ts` | UTC timestamp |
| `event` | Tool or lifecycle name (`sharepoint_search`, `sign_in`, …) |
| `outcome` | `allowed`, `denied`, or `error` |
| `tenant_id` | Configured Entra tenant (directory id) |
| `actor` | HMAC hash of the Entra `oid` (not UPN) |
| `resources` | HMAC hashes of Graph resource ids used in the call |
| `error` | Stable connector error code |
| `graph_request_id` | Graph `request-id` header when present |
| `client_request_id` | Our/`client-request-id` (or Entra correlation when recorded on sign-in deny) |
| `duration_ms` | Wall time for the tool call |
| `counts` | Small numeric tallies (result sizes), never content |

Ordinary logs must **not** contain access tokens, refresh tokens, UPNs, raw item/drive ids, queries, or file text.

## Correlating to Graph / Entra

1. Reproduce with the affected user on the same device.
2. Open the day’s `audit-*.jsonl` and find the `ts` / `error` / `event`.
3. Give Microsoft support the `graph_request_id` and/or `client_request_id` only — not the full JSONL unless redacted.
4. Entra sign-in failures may only have `client_request_id` (no Graph call).

## Disable / reconnect

1. Call `sharepoint_sign_out` (or delete the MSAL cache under the data dir per the README).
2. Confirm data tools disappear from the host tool list.
3. Call `sharepoint_sign_in` and re-consent if the admin revoked the grant.

Host firewall egress (login.microsoftonline.com / graph.microsoft.com / `*.sharepoint.com`) remains a **device-admin** control; the connector only refuses unsafe Graph paths and non-SharePoint download redirects.
