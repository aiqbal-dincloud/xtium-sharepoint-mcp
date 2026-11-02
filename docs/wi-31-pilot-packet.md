# WI-31 / G6 — two-user standalone MCP pilot packet (draft)

Fill this after f3zj live runs. Do **not** paste tokens, secrets, raw item ids, or document content into Jira.

## Environment

| Item | Value |
|------|-------|
| Tenant | f3zj (`6d21634d-ca87-4a22-99eb-e31240eeea86`) or named licensed tenant |
| Test app | from `local/test_app.json` (not committed) |
| Client | Claude Desktop (first); others separately |
| Connector build | `xtium-sharepoint-mcp` local checkout |
| User A | `aaqib.iqbal@f3zj.onmicrosoft.com` |
| User B | `usama.saeed@f3zj.onmicrosoft.com` |
| A-only fixture | Word file in Alice's OneDrive. Organization link removed. Access list is Aaqib only. Alice remains the file owner and is not User B. |

## Checklist

- [x] Two licensed users confirmed — 2026-10-30 rescan: `DEVELOPERPACK_V2_E5` consumed 4 of 25. Licensed: Aaqib, Alice, Jordan, Usama. Alice is the admin operator, not User B.
- [x] Public client created + tenant admin consent
- [x] Claude Desktop config points at stdio entry + tenant/client
- [x] User A: signed in as Aaqib, search found `wi31-a-only.docx` in 1882 ms, read allowed in 4269 ms (14,449 bytes, 25 extracted characters). No file text stored here.
- [x] User B: Usama signed in at 11:43:01Z. Search for `wi31-a-only` returned 0 results in 1525 ms. Reading User A’s handle returned `invalid_handle` in 1 ms, with no file bytes.
- [ ] Negatives from `docs/wi-30-negative-matrix.md` live column
- [ ] Audit sample redacted and attached/linked (hashed ids only)
- [ ] p95 search &lt; 3 s; read ≤5 MB &lt; 5 s (or Product revises — do not mark pass if missed)
- [ ] Attestation: no content/tokens stored on XTIUM infra
- [ ] Security/connector/pilot owners named for go/no-go

## Results

- 2026-10-30 User A pass. Search 1882 ms (under 3 s). Read of a 14 KB docx 4269 ms (under 5 s). One sample only, not a p95.
- 2026-10-30 User B pass. Search returned 0 results in 1525 ms. Replaying User A’s handle returned `invalid_handle` in 1 ms and did not read the file.
