# Submit local SharePoint to Anthropic’s directory

Anthropic’s **MCP connector** form only accepts a **remote `https://` server**.  
Local MCPB listings are deprecated. To list a **local** SharePoint connector in the directory, submit a **plugin** that embeds the `.mcpb`.

Portal: [claude.ai/directory/manage](https://claude.com/directory/manage) → **Submit new** → **Plugin bundle**  
(Not “MCP connector”.)

## What this repo already has

| Piece | Path |
| --- | --- |
| Plugin manifest | `.claude-plugin/plugin.json` → points at `./dist/xtium-sharepoint-0.1.0.mcpb` |
| MCPB (local server + tenant/client `user_config`) | `dist/xtium-sharepoint-0.1.0.mcpb` |
| Privacy policy | `README.md` § Privacy Policy + `privacy_policies` in `manifest.json` |
| License | `LICENSE` |

Rebuild the bundle before submit:

```powershell
.\scripts\pack_mcpb.ps1
```

## Hard requirements from Anthropic

1. **github.com** repository (GHE alone is not enough for the portal).
2. Connected GitHub account with **push** access on that repo.
3. Paid Claude plan / org role that can submit.
4. Repo may stay **private** through validate/submit; it must be **public** before the listing can go **live**.
5. Your Claude **organization** may still block members from installing plugins — that is separate from Anthropic review. Ask an org admin to allow plugin installs for the pilot.

## Recommended flow for XTIUM

1. Mirror (or push) this plugin to a **github.com** repo, e.g. `XTIUM/xtium-sharepoint-mcp` (or an open-source-safe subset if Product requires).
2. Confirm `plugin.json`, `LICENSE`, Privacy Policy, and fresh `.mcpb` are on the default branch.
3. Locally (optional): `claude plugin validate .` from the repo root.
4. Open [claude.ai/directory/manage](https://claude.com/directory/manage) → **Submit new** → **Plugin bundle**.
5. Source:
   - Repository: `owner/repo` on github.com  
   - Plugin path: leave empty if `.claude-plugin/plugin.json` is at repo root  
   - Branch: `main`
6. **Validate** in the portal; fix any findings; validate again.
7. Complete listing / data handling / compliance → **Submit for review**.
8. When checks pass and repo is public → **Publish**.

## What we are not submitting

- A remote SharePoint MCP URL under **MCP connector** (would put Graph behind Anthropic’s cloud unless redesigned).
- Paid Azure hosting on f3zj for the package (use github.com / existing channels for distribution).

## After listing

Users still configure **Entra tenant ID** and **client ID** when the bundled MCPB first starts, then run `sharepoint_sign_in`.
