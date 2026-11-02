"""Delegated sign-in for one local user: MSAL public client, auth code + PKCE, DPAPI-encrypted cache.

The connector never holds a client secret and never sends tokens anywhere except Microsoft Graph.
"""

from __future__ import annotations

import base64
import contextlib
import json
import sys
import threading
from dataclasses import dataclass
from typing import Any, Protocol

from .config import (
    ALLOWED_SCOPES,
    AUTHORITY_HOST,
    GRAPH_APP_ID,
    GRAPH_RESOURCE,
    Settings,
)
from .errors import AuthError

GRAPH_AUDIENCES = frozenset({GRAPH_RESOURCE, f"{GRAPH_RESOURCE}/", GRAPH_APP_ID})

# AADSTS codes that mean the configured tenant or client is wrong, or consent is missing.
_WRONG_TENANT_CODES = {50020, 90072, 90002, 50128, 500011}
_WRONG_CLIENT_CODES = {700016, 7000215, 700054, 50011, 500113}
_CONSENT_CODES = {65001, 65004, 90094, 90095}
_REAUTH_CODES = {50173, 70043, 700082, 50076, 50079, 50078, 50132, 50133, 53003}


@dataclass(frozen=True)
class Identity:
    tenant_id: str
    object_id: str
    granted_scopes: frozenset[str]


class TokenSource(Protocol):
    def identity(self) -> Identity | None: ...

    def access_token(self, force_refresh: bool = False) -> str: ...


def _b64url_json(segment: str) -> dict[str, Any]:
    padded = segment + "=" * (-len(segment) % 4)
    return json.loads(base64.urlsafe_b64decode(padded.encode("ascii")))


def unverified_jwt_claims(token: str) -> dict[str, Any]:
    """Read claims from the Graph access token for defence-in-depth checks only.

    Graph verifies the signature; the connector checks the claims so a misconfigured app,
    tenant or resource fails closed locally instead of reaching Graph.
    """
    parts = token.split(".")
    if len(parts) != 3:
        raise AuthError("token_unverifiable", "The access token is not a JWT; refusing to use it.")
    try:
        return _b64url_json(parts[1])
    except (ValueError, json.JSONDecodeError) as exc:
        raise AuthError("token_unverifiable", "The access token could not be decoded.") from exc


def _error_from_result(result: dict[str, Any]) -> AuthError:
    codes = set(result.get("error_codes") or [])
    error = str(result.get("error") or "auth_failed")
    details = {"aad_error": error, "aad_error_codes": sorted(codes)}
    correlation = result.get("correlation_id")
    if codes & _WRONG_TENANT_CODES:
        code, msg = "wrong_tenant", "The account or app is not in the configured tenant."
    elif codes & _WRONG_CLIENT_CODES or error in {"invalid_client", "unauthorized_client"}:
        code, msg = "wrong_client", "The configured client ID is not a valid public client here."
    elif codes & _CONSENT_CODES or error == "consent_required":
        code, msg = "consent_required", "The tenant has not consented to the requested scopes."
    elif error == "access_denied":
        code, msg = "sign_in_cancelled", "Sign-in was cancelled or denied."
    elif codes & _REAUTH_CODES or error in {"invalid_grant", "interaction_required"}:
        code, msg = "reauth_required", "Sign in again; the session was revoked or expired."
    else:
        code, msg = "auth_failed", "Sign-in failed."
    return AuthError(code, msg, client_request_id=correlation, details=details)


def validate_token_result(result: dict[str, Any], settings: Settings) -> tuple[Identity, str]:
    """Check an MSAL token result against the configured tenant, client, issuer and Graph audience."""
    if not result or "error" in result:
        raise _error_from_result(result or {})
    token = result.get("access_token")
    if not token:
        raise AuthError("auth_failed", "No access token was returned.")

    id_claims = result.get("id_token_claims") or {}
    if id_claims:
        tid = str(id_claims.get("tid", "")).lower()
        if tid != settings.tenant_id:
            raise AuthError("wrong_tenant", "The ID token is from a different tenant.")
        if str(id_claims.get("aud", "")).lower() != settings.client_id:
            raise AuthError("wrong_client", "The ID token was issued to a different client.")
        expected_iss = f"{AUTHORITY_HOST}/{settings.tenant_id}/v2.0"
        if str(id_claims.get("iss", "")).lower() != expected_iss:
            raise AuthError("wrong_issuer", "The ID token issuer does not match the tenant.")

    claims = unverified_jwt_claims(token)
    if str(claims.get("aud", "")).lower() not in GRAPH_AUDIENCES:
        raise AuthError("wrong_audience", "The access token is not for Microsoft Graph.")
    if str(claims.get("tid", "")).lower() != settings.tenant_id:
        raise AuthError("wrong_tenant", "The access token is from a different tenant.")
    client = str(claims.get("appid") or claims.get("azp") or "").lower()
    if client != settings.client_id:
        raise AuthError("wrong_client", "The access token was issued to a different client.")
    scp = str(claims.get("scp") or "")
    if not scp:
        raise AuthError("app_only_token", "Only delegated user tokens are accepted.")
    oid = str(claims.get("oid") or id_claims.get("oid") or "").lower()
    if not oid:
        raise AuthError("auth_failed", "The token has no user object ID.")
    if id_claims and str(id_claims.get("oid", oid)).lower() != oid:
        raise AuthError("user_mismatch", "The ID token and access token name different users.")

    granted = frozenset(s for s in scp.split() if s in ALLOWED_SCOPES)
    return Identity(settings.tenant_id, oid, granted), token


class AuthSession:
    """One signed-in user per OS profile. The cache is encrypted with Windows DPAPI."""

    def __init__(self, settings: Settings, app: Any | None = None):
        self._settings = settings
        self._lock = threading.RLock()
        self._app = app
        self._identity: Identity | None = None

    def _client(self) -> Any:
        if self._app is None:
            import msal
            from msal_extensions import PersistedTokenCache, build_encrypted_persistence

            self._settings.data_dir.mkdir(parents=True, exist_ok=True)
            persistence = build_encrypted_persistence(str(self._settings.cache_path))
            self._app = msal.PublicClientApplication(
                self._settings.client_id,
                authority=self._settings.authority,
                token_cache=PersistedTokenCache(persistence),
            )
        return self._app

    def _tenant_accounts(self) -> list[dict[str, Any]]:
        suffix = f".{self._settings.tenant_id}"
        return [
            a
            for a in self._client().get_accounts()
            if str(a.get("home_account_id", "")).lower().endswith(suffix)
        ]

    def sign_in(self, timeout: int = 180) -> Identity:
        with self._lock:
            app = self._client()
            for account in app.get_accounts():
                app.remove_account(account)
            self._identity = None
            # MSAL may print a fallback URL; stdout is the MCP JSON-RPC channel.
            with contextlib.redirect_stdout(sys.stderr):
                result = app.acquire_token_interactive(
                    scopes=self._settings.graph_scopes,
                    prompt="select_account",
                    timeout=timeout,
                )
            try:
                identity, _ = validate_token_result(result, self._settings)
            except AuthError:
                for account in app.get_accounts():
                    app.remove_account(account)
                raise
            self._identity = identity
            return identity

    def _silent(self, force_refresh: bool) -> tuple[Identity, str]:
        app = self._client()
        accounts = self._tenant_accounts()
        if not accounts:
            raise AuthError("not_signed_in", "Not signed in. Run the sign-in tool or CLI command.")
        if len(accounts) > 1:
            raise AuthError("ambiguous_account", "More than one cached account; sign out first.")
        result = app.acquire_token_silent_with_error(
            self._settings.graph_scopes, account=accounts[0], force_refresh=force_refresh
        )
        if result is None:
            raise AuthError("reauth_required", "Sign in again; no usable cached session.")
        identity, token = validate_token_result(result, self._settings)
        if self._identity and self._identity.object_id != identity.object_id:
            raise AuthError("user_mismatch", "The cached user changed; sign in again.")
        self._identity = identity
        return identity, token

    def identity(self) -> Identity | None:
        with self._lock:
            # Empty or missing cache: do not construct MSAL (it hits OIDC discovery).
            # Lets `serve` and stdio smoke start signed-out without a live tenant.
            cache = self._settings.cache_path
            if (not cache.exists()) or cache.stat().st_size == 0:
                return None
            try:
                return self._silent(force_refresh=False)[0]
            except AuthError as exc:
                if exc.code in {"not_signed_in"}:
                    return None
                raise

    def access_token(self, force_refresh: bool = False) -> str:
        with self._lock:
            return self._silent(force_refresh)[1]

    def sign_out(self) -> None:
        """Remove the local session. Server-side revocation is a tenant admin action (see docs)."""
        with self._lock:
            app = self._client()
            for account in app.get_accounts():
                app.remove_account(account)
            self._identity = None
        with contextlib.suppress(FileNotFoundError):
            self._settings.cache_path.unlink()
