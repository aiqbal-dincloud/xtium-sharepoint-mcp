# Manual verification (plain steps)

Nothing in the connector is hard-coded to Jordan. **User A is whoever signs in** in the browser.  
Jordan only appeared in the readiness scan because he is currently the **only** f3zj account with a SharePoint licence.

To run as **Aaqib**: Aaqib must have an f3zj account **and** a SharePoint seat assigned (24 seats free on `DEVELOPERPACK_V2_E5`). Then sign in as Aaqib at step 3 below.

Do **not** paste tokens, passwords, or file content into chat or Jira.

---

## Part 0 — One-time admin setup (Entra / Microsoft 365 admin)

1. Open the f3zj admin centre (same tenant as app `f4368549-8052-4a54-86fc-ce8635e0af70`).
2. Confirm Aaqib exists as a user (or create `…@f3zj.onmicrosoft.com`).
3. **Licences** → assign a seat that includes SharePoint (same E5 developer pack is fine) to Aaqib.
4. Optional for later ACL tests: create a second licensed user (User B) and a library/file only Aaqib can open.
5. Re-check licences (optional):

```powershell
cd c:\Users\AaqibIqbal\XTAI\xtium-sharepoint-mcp
uv run python scripts/provision_test_app.py --tenant 6d21634d-ca87-4a22-99eb-e31240eeea86
```

You should see Aaqib under `LICENSED_USERS` (and Jordan if still licensed).

App + consent are already done — you do **not** need `--create` again unless the app was deleted.

---

## Part 1 — Use Anthropic Claude (not Gateway), with MCP already configured

You have **two** Claude data profiles on this PC:

| Profile | Folder | Purpose |
| --- | --- | --- |
| **Consumer Claude (use this for SharePoint)** | `%LOCALAPPDATA%\Claude-Data\claude_pzs8sxrjxfjjc\` | Normal Anthropic chat. Account must **not** say “Gateway”. |
| **Claude 3p / Gateway (skip for now)** | `%LOCALAPPDATA%\Claude-3p\` | Local AM pilot (`http://127.0.0.1:8001`). Needs a gateway identity link. |

The SharePoint MCP block is already in **both** configs. For SharePoint-only verification:

1. Fully quit every Claude window (tray → Quit).
2. Open **Claude** from the Start menu (the normal MSIX app).
3. Check the bottom-left account label: it must **not** say `· Gateway`. If it does, you are still on 3p — quit and open the consumer app again.
4. Sign in with your normal Anthropic account if asked (not Entra gateway Setup).
5. Start a **new** chat (Cowork or Chat is fine).
6. Confirm tools like `sharepoint_sign_in` / `get_connection_status` appear.

If tools are missing, open this file and confirm `mcpServers` → `xtium-sharepoint`:

`%LOCALAPPDATA%\Claude-Data\claude_pzs8sxrjxfjjc\claude_desktop_config.json`

Do **not** use **Open Setup** / Gateway credentials for this path. That is only for the AM pilot.

---

## Part 2 — Happy path as Aaqib (copy these prompts)

In Claude Desktop, ask one step at a time. After each step, check what Claude reports.

### Step 1 — Status while signed out

**You say:**  
`Call get_connection_status and show me signed_in and available_tools.`

**Pass if:** `signed_in` is false (or equivalent). Tools are mainly status / sign-in / sign-out. No search/read yet (or they fail with not signed in).

### Step 2 — Sign in as Aaqib

**You say:**  
`Call sharepoint_sign_in.`

- A browser window opens.
- Sign in with **Aaqib’s f3zj account** (not Jordan, not your work XTIUM account unless that *is* the f3zj user).
- Complete consent if prompted.

**Pass if:** Tool returns success / `signed_in: true` and lists granted scopes.  
If the tool list does not refresh, **restart Claude Desktop**, open a new chat, run `get_connection_status` again — you should now see search / site / folder / read tools.

### Step 3 — Search

**You say:**  
`Call sharepoint_search with query "test" (or any word you know exists in SharePoint). Show titles and handles only.`

**Pass if:** You get file/site metadata and handles starting like `spf_…`.  
**Fail if:** You see download URLs, full file text, or raw Graph ids that look like drive paths.

### Step 4 — Site

**You say:**  
`Call sharepoint_get_site with site "root".`

**Pass if:** You get site info and libraries with folder handles.

### Step 5 — List a folder

**You say:**  
`Using a folder handle from the previous step, call sharepoint_list_folder.`

**Pass if:** You see files/folders and optional `next_cursor`.

### Step 6 — Read a small file

**You say:**  
`Call sharepoint_read_file on a small PDF or DOCX or text file handle from the list.`

**Pass if:** You get a bounded text extract.  
**Skip if:** You only have XLSX/PPTX (connector refuses those on purpose).

### Step 7 — Audit (on your PC, not in Claude)

1. Open File Explorer.  
2. Go to: `%LOCALAPPDATA%\xtium-sharepoint-mcp\audit\`  
3. Open today’s `audit-….jsonl` in Notepad.

**Pass if:** Lines have `event`, `outcome`, hashed `actor` / `resources`, maybe `graph_request_id`.  
**Fail if:** You see access tokens, passwords, or full document text.

Tick WI-25 “Claude Desktop round trip” = Pass when steps 1–6 work as Aaqib.

---

## Part 3 — Manual negatives (still as Aaqib / one user)

Do these after the happy path. Restore the real tenant/client IDs when finished.

### N1 — Wrong tenant

1. Edit `claude_desktop_config.json`.  
2. Change `XTIUM_SP_TENANT_ID` to a fake GUID, e.g. `11111111-1111-1111-1111-111111111111`.  
3. Quit + reopen Claude.  
4. Call `sharepoint_sign_in`.

**Pass if:** Sign-in fails (wrong tenant / issuer).  
5. Put the real tenant id back, restart.

### N2 — Wrong client

1. Set `XTIUM_SP_CLIENT_ID` to a random GUID.  
2. Restart Claude, call `sharepoint_sign_in`.

**Pass if:** Fail (app not found / wrong client).  
3. Restore real client id, restart.

### N3 — Bad handle

**You say:**  
`Call sharepoint_read_file with file "not-a-real-handle".`

**Pass if:** Error like `invalid_handle`.

### N4 — Sign out

**You say:**  
`Call sharepoint_sign_out, then get_connection_status.`

**Pass if:** Signed out; search/read gone or denied.  
Then try reading an old `spf_…` handle from earlier — should fail (`invalid_handle`).

### N5 — Revoke (admin)

1. Sign in again as Aaqib.  
2. In Entra: revoke Aaqib’s sessions **or** remove this app’s user consent for Aaqib.  
3. In Claude, call `sharepoint_search` (may need a few tries or wait until token refresh).

**Pass if:** Eventually `reauth_required` (or similar) and data tools disappear after status/retry.  
Sign in again to recover.

Mark live cells in `docs/wi-30-negative-matrix.md` as you go.

---

## Part 4 — Two users (only when User B is licensed)

1. Clear session: `sharepoint_sign_out` or delete `%LOCALAPPDATA%\xtium-sharepoint-mcp`.  
2. Sign in as **Aaqib** → search/read the **A-only** file → success.  
3. Sign out / clear cache.  
4. Sign in as **User B** → same search must **not** show that file; reading A’s handle must fail.  
5. Fill `docs/wi-31-pilot-packet.md`.

Until User B exists, **skip Part 4** — that is the only G6 blocker, not the Aaqib happy path.

---

## What you do *not* need to do manually

These are already covered by `uv run pytest` (offline). You can run that once for confidence:

```powershell
cd c:\Users\AaqibIqbal\XTAI\xtium-sharepoint-mcp
uv run pytest -q
```

You do **not** need to re-test every Graph status code by hand if pytest is green.

---

## Short “did I pass?” checklist

- [ ] Aaqib has SharePoint licence on f3zj  
- [ ] Claude config points at the exe + real tenant/client  
- [ ] Sign-in as **Aaqib** works  
- [ ] Search → site → folder → read works  
- [ ] Audit file has no secrets/content  
- [ ] Wrong tenant / wrong client / bad handle / sign-out behave as above  
- [ ] (Later) User B cannot see A-only content  
