"""Check a test tenant for SharePoint readiness and create the non-production public-client test app.

Signs in once with device code (Azure CLI public client, Graph delegated) as a tenant admin.
Writes non-secret values to local/test_app.json and local/tenant_readiness.json (gitignored).
The app is a public client: it has no secret or certificate.

Usage: uv run python scripts/provision_test_app.py --tenant <guid> [--create] [--grant-consent]
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import httpx
import msal

AZ_CLI_CLIENT = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
APP_NAME = "AIDATA-3970 SharePoint MCP local test (non-production)"
DELEGATED = {
    "User.Read": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
    "Sites.Read.All": "205e70e5-aba6-4c52-a976-6d2d46c48043",
    "Files.Read.All": "df85f4d6-205c-4ac5-a5ea-6bf408dba283",
    "offline_access": "7427e0e9-2fba-42fe-b0c0-848c9e6a3102",
    "openid": "37f7f235-527c-4136-accd-4a02d197296e",
    "profile": "14dad69e-099b-42c9-810b-d002981feec1",
}
LOCAL = Path(__file__).resolve().parent.parent / "local"


def sign_in(tenant: str) -> str:
    app = msal.PublicClientApplication(AZ_CLI_CLIENT, authority=f"https://login.microsoftonline.com/{tenant}")
    flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])
    if "user_code" not in flow:
        raise SystemExit(f"device flow failed: {flow.get('error')}")
    print(f"DEVICE_CODE {flow['user_code']} at {flow['verification_uri']}", flush=True)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise SystemExit(f"sign-in failed: {result.get('error')} {result.get('error_codes')}")
    claims = result.get("id_token_claims") or {}
    if str(claims.get("tid", "")).lower() != tenant:
        raise SystemExit("signed in to the wrong tenant")
    print("SIGNED_IN", flush=True)
    return result["access_token"]


class Graph:
    def __init__(self, token: str):
        self._http = httpx.Client(base_url=GRAPH, timeout=60, headers={"Authorization": f"Bearer {token}"})

    def call(self, method: str, path: str, **kw: Any) -> dict[str, Any]:
        resp = self._http.request(method, path, **kw)
        if resp.status_code >= 400:
            raise RuntimeError(f"{method} {path} -> {resp.status_code} {resp.text[:300]}")
        return resp.json() if resp.content else {}

    def pages(self, path: str, params: dict[str, Any] | None = None) -> list[dict[str, Any]]:
        out, data = [], self.call("GET", path, params=params)
        while True:
            out += data.get("value", [])
            link = data.get("@odata.nextLink")
            if not link:
                return out
            data = self.call("GET", link.replace(GRAPH, ""))


def readiness(g: Graph) -> dict[str, Any]:
    skus = []
    for sku in g.pages("/subscribedSkus"):
        sp_plans = [p["servicePlanName"] for p in sku.get("servicePlans", [])
                    if p["servicePlanName"].startswith("SHAREPOINT") and p.get("provisioningStatus") == "Success"]
        if sp_plans:
            skus.append({
                "sku": sku["skuPartNumber"],
                "sharepoint_plans": sp_plans,
                "enabled": sku.get("prepaidUnits", {}).get("enabled"),
                "consumed": sku.get("consumedUnits"),
            })
    users = []
    for u in g.pages("/users", {"$select": "id,displayName,userPrincipalName,accountEnabled,assignedPlans",
                                "$top": "999"}):
        has_sp = any(p.get("service") == "SharePoint" and p.get("capabilityStatus") == "Enabled"
                     for p in u.get("assignedPlans", []))
        if has_sp and u.get("accountEnabled"):
            users.append({"id": u["id"], "displayName": u.get("displayName"), "upn": u.get("userPrincipalName")})
    return {"sharepoint_skus": skus, "sharepoint_licensed_users": users}


def ensure_app(g: Graph, tenant: str, grant_consent: bool) -> dict[str, Any]:
    existing = g.call("GET", "/applications", params={"$filter": f"displayName eq '{APP_NAME}'"}).get("value", [])
    if existing:
        app = existing[0]
        print("APP_EXISTS", app["appId"], flush=True)
    else:
        app = g.call("POST", "/applications", json={
            "displayName": APP_NAME,
            "signInAudience": "AzureADMyOrg",
            "isFallbackPublicClient": True,
            "publicClient": {"redirectUris": ["http://localhost"]},
            "notes": "AIDATA-3970 pilot. Non-production. Public client, no secret. Delete after the pilot.",
            "requiredResourceAccess": [{
                "resourceAppId": GRAPH_APP_ID,
                "resourceAccess": [{"id": i, "type": "Scope"} for i in DELEGATED.values()],
            }],
        })
        print("APP_CREATED", app["appId"], flush=True)
    sps = g.call("GET", "/servicePrincipals", params={"$filter": f"appId eq '{app['appId']}'"}).get("value", [])
    sp = sps[0] if sps else g.call("POST", "/servicePrincipals", json={"appId": app["appId"]})

    consent = None
    if grant_consent:
        graph_sp = g.call("GET", "/servicePrincipals", params={"$filter": f"appId eq '{GRAPH_APP_ID}'"})["value"][0]
        scope = " ".join(DELEGATED)
        grants = g.call("GET", "/oauth2PermissionGrants",
                        params={"$filter": f"clientId eq '{sp['id']}' and resourceId eq '{graph_sp['id']}'"})["value"]
        if grants:
            g.call("PATCH", f"/oauth2PermissionGrants/{grants[0]['id']}", json={"scope": scope})
        else:
            g.call("POST", "/oauth2PermissionGrants", json={
                "clientId": sp["id"], "consentType": "AllPrincipals", "resourceId": graph_sp["id"], "scope": scope,
            })
        consent = scope
        print("CONSENT_GRANTED", scope, flush=True)

    return {
        "tenant_id": tenant,
        "client_id": app["appId"],
        "display_name": APP_NAME,
        "redirect_uri": "http://localhost",
        "delegated_scopes": list(DELEGATED),
        "tenant_wide_consent": consent,
        "application_object_id": app["id"],
        "service_principal_id": sp["id"],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--tenant", required=True)
    parser.add_argument("--create", action="store_true")
    parser.add_argument("--grant-consent", action="store_true")
    args = parser.parse_args()
    tenant = args.tenant.lower()

    g = Graph(sign_in(tenant))
    LOCAL.mkdir(exist_ok=True)
    ready = readiness(g)
    (LOCAL / "tenant_readiness.json").write_text(json.dumps(ready, indent=2), encoding="utf-8")
    print("SKUS", json.dumps(ready["sharepoint_skus"]), flush=True)
    print("LICENSED_USERS", len(ready["sharepoint_licensed_users"]), flush=True)
    for u in ready["sharepoint_licensed_users"]:
        print("  USER", u["displayName"], u["upn"], flush=True)

    if args.create:
        values = ensure_app(g, tenant, args.grant_consent)
        (LOCAL / "test_app.json").write_text(json.dumps(values, indent=2), encoding="utf-8")
        print("WROTE", LOCAL / "test_app.json", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
