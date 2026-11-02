# WI-30 negative matrix (local connector)

Offline rows are covered by pytest. Live rows need the f3zj test app and two licensed users (WI-25/31).

| Case | Expected | Offline | Live |
|------|----------|---------|------|
| Wrong tenant id in config | Sign-in / token validation fail (`bad_tenant` / issuer) | Yes | Pass (fail-closed). Fake tenant `11111111-1111-1111-1111-111111111111` made MSAL reject startup (`AADSTS90002` tenant not found). The process exited, so Claude never listed `sharepoint_sign_in`. Real tenant restored; `tools/list` succeeded at 2026-10-30 09:56:46Z. |
| Wrong client id | Sign-in fail | Yes | Pass — fake client `22222222-2222-2222-2222-222222222222` returned `AADSTS700016` (application not found in the f3zj directory) at 2026-10-30 10:26:28Z. Real client id restored. |
| Wrong Graph audience | Token rejected | Yes | Pending |
| Forged / cross-user handle | `invalid_handle` | Yes | Pass — live garbage handle returned `invalid_handle` ("This is not a connector handle.") |
| Logout + cache clear | Data tools hidden; old handles invalid | Yes | Pass — `sharepoint_sign_out` left `signed_in: false` and only `get_connection_status` available |
| Consent removed / refresh revoked | `reauth_required`; tool list drops data tools + `tools/list_changed` | Yes (unit) | Partial — admin sign-out of the signed-in user made a forced refresh fail with `AADSTS50173` (`invalid_grant`, maps to `reauth_required`). Claude searches at 10:58Z–11:07Z still succeeded because Graph accepted the existing access token. The connector refreshes only after a Graph 401. |
| Disabled user | `reauth_required` or Graph unauthorized → reauth path | Partial | Pending |
| 403 inaccessible item | `forbidden` (search may omit; read denies) | Yes | Pending |
| 429 throttle | `throttled` + capped `retry_after_seconds` | Yes | Pending |
| Graph timeout / connect failure | `timeout` / `network_error` | Yes | Partial — one live search returned `upstream_unavailable` (Graph 5xx) at 2026-10-30 10:39:13Z and the connector did not retry that call. The next manual call succeeded. `timeout` / `network_error` not seen live. |
| CA / CAE challenge | Not instant wipe; document window — see ADR → WI-37 | Documented limit | Observed window — refresh token dead by 11:08Z, access token still accepted at 11:07Z (sign-in 10:34Z). Full `reauth_required` in Claude still open until that access token is rejected. |

## Known limitation (CAE)

Conditional Access / continuous access evaluation may leave a short provider-side window after revoke. The connector does not claim instant wipe; measure the real window in the WI-31 packet and point operators at sign-out + cache delete.
