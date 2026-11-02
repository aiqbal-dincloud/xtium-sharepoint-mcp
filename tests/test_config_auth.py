from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from xtium_sharepoint_mcp.auth import AuthSession, validate_token_result
from xtium_sharepoint_mcp.config import load_settings
from xtium_sharepoint_mcp.errors import AuthError, ConfigError

from .conftest import CLIENT, OTHER_TENANT, TENANT, USER_A, USER_B, graph_token, id_claims


def _env(**extra: str) -> dict[str, str]:
    return {"XTIUM_SP_TENANT_ID": TENANT, "XTIUM_SP_CLIENT_ID": CLIENT, **extra}


class TestConfig:
    def test_defaults(self, tmp_path: Path) -> None:
        s = load_settings(_env(XTIUM_SP_DATA_DIR=str(tmp_path)))
        assert s.authority == f"https://login.microsoftonline.com/{TENANT}"
        assert s.graph_scopes == [
            "https://graph.microsoft.com/User.Read",
            "https://graph.microsoft.com/Sites.Read.All",
            "https://graph.microsoft.com/Files.Read.All",
        ]

    @pytest.mark.parametrize("tenant", ["common", "organizations", "contoso.com", ""])
    def test_rejects_non_single_tenant(self, tenant: str) -> None:
        with pytest.raises(ConfigError):
            load_settings(_env(XTIUM_SP_TENANT_ID=tenant))

    def test_rejects_write_or_foreign_scopes(self) -> None:
        with pytest.raises(ConfigError) as err:
            load_settings(_env(XTIUM_SP_SCOPES="Files.ReadWrite.All api://xtium/.default"))
        assert err.value.details["rejected"] == ["Files.ReadWrite.All", "api://xtium/.default"]

    def test_files_only_keeps_user_read(self) -> None:
        assert load_settings(_env(XTIUM_SP_SCOPES="Files.Read.All")).scopes == (
            "User.Read", "Files.Read.All",
        )


def _result(**token_overrides: Any) -> dict[str, Any]:
    id_over = token_overrides.pop("id_claims", {})
    return {"access_token": graph_token(**token_overrides), "id_token_claims": id_claims(**id_over)}


class TestValidateTokenResult:
    def test_accepts_delegated_graph_token(self, settings) -> None:
        identity, token = validate_token_result(_result(), settings)
        assert identity.tenant_id == TENANT and identity.object_id == USER_A
        assert identity.granted_scopes == {"User.Read", "Sites.Read.All", "Files.Read.All"}
        assert token.count(".") == 2

    def test_granted_scopes_ignore_extras(self, settings) -> None:
        identity, _ = validate_token_result(_result(scp="Files.Read.All Mail.Read openid"), settings)
        assert identity.granted_scopes == {"Files.Read.All"}

    @pytest.mark.parametrize(
        ("overrides", "code"),
        [
            ({"tid": OTHER_TENANT}, "wrong_tenant"),
            ({"id_claims": {"tid": OTHER_TENANT}}, "wrong_tenant"),
            ({"aud": "api://xtium-gateway"}, "wrong_audience"),
            ({"aud": "https://graph.windows.net"}, "wrong_audience"),
            ({"appid": "99999999-9999-9999-9999-999999999999"}, "wrong_client"),
            ({"id_claims": {"aud": "99999999-9999-9999-9999-999999999999"}}, "wrong_client"),
            ({"id_claims": {"iss": "https://sts.windows.net/x/"}}, "wrong_issuer"),
            ({"scp": None, "roles": ["Sites.Read.All"]}, "app_only_token"),
            ({"id_claims": {"oid": USER_B}}, "user_mismatch"),
        ],
    )
    def test_negative_claims(self, settings, overrides: dict[str, Any], code: str) -> None:
        with pytest.raises(AuthError) as err:
            validate_token_result(_result(**overrides), settings)
        assert err.value.code == code

    def test_opaque_token_refused(self, settings) -> None:
        with pytest.raises(AuthError) as err:
            validate_token_result({"access_token": "opaque", "id_token_claims": id_claims()}, settings)
        assert err.value.code == "token_unverifiable"

    @pytest.mark.parametrize(
        ("result", "code"),
        [
            ({"error": "invalid_client", "error_codes": [700016]}, "wrong_client"),
            ({"error": "invalid_request", "error_codes": [50020]}, "wrong_tenant"),
            ({"error": "invalid_grant", "error_codes": [70043]}, "reauth_required"),
            ({"error": "invalid_grant", "error_codes": [65001]}, "consent_required"),
            ({"error": "access_denied"}, "sign_in_cancelled"),
        ],
    )
    def test_aad_errors_are_typed(self, settings, result: dict[str, Any], code: str) -> None:
        with pytest.raises(AuthError) as err:
            validate_token_result({**result, "correlation_id": "corr-1"}, settings)
        assert err.value.code == code
        assert err.value.client_request_id == "corr-1"
        assert "access_token" not in str(err.value)


class FakeMsal:
    def __init__(self, interactive: dict[str, Any], accounts: list[dict[str, Any]] | None = None):
        self.interactive = interactive
        self.accounts = accounts or []
        self.silent: dict[str, Any] | None = None
        self.calls: list[str] = []

    def get_accounts(self) -> list[dict[str, Any]]:
        return list(self.accounts)

    def remove_account(self, account: dict[str, Any]) -> None:
        self.accounts.remove(account)

    def acquire_token_interactive(self, **kwargs: Any) -> dict[str, Any]:
        self.calls.append(f"interactive:{kwargs['prompt']}")
        oid = (self.interactive.get("id_token_claims") or {}).get("oid", USER_A)
        self.accounts = [{"home_account_id": f"{oid}.{TENANT}"}]
        return self.interactive

    def acquire_token_silent_with_error(self, scopes, account, force_refresh=False):
        self.calls.append(f"silent:{force_refresh}")
        return self.silent


class TestAuthSession:
    def test_sign_in_replaces_previous_account(self, settings) -> None:
        app = FakeMsal(_result(), accounts=[{"home_account_id": f"{USER_B}.{TENANT}"}])
        identity = AuthSession(settings, app=app).sign_in()
        assert identity.object_id == USER_A
        assert app.accounts == [{"home_account_id": f"{USER_A}.{TENANT}"}]
        assert app.calls == ["interactive:select_account"]

    def test_failed_validation_leaves_no_account(self, settings) -> None:
        app = FakeMsal(_result(aud="api://other"))
        with pytest.raises(AuthError):
            AuthSession(settings, app=app).sign_in()
        assert app.accounts == []

    def test_silent_uses_only_configured_tenant(self, settings) -> None:
        app = FakeMsal({}, accounts=[{"home_account_id": f"{USER_A}.{OTHER_TENANT}"}])
        assert AuthSession(settings, app=app).identity() is None

    def test_two_cached_accounts_is_ambiguous(self, settings) -> None:
        app = FakeMsal({}, accounts=[
            {"home_account_id": f"{USER_A}.{TENANT}"}, {"home_account_id": f"{USER_B}.{TENANT}"},
        ])
        with pytest.raises(AuthError) as err:
            AuthSession(settings, app=app).identity()
        assert err.value.code == "ambiguous_account"

    def test_revoked_refresh_token_requires_reauth(self, settings) -> None:
        app = FakeMsal({}, accounts=[{"home_account_id": f"{USER_A}.{TENANT}"}])
        app.silent = {"error": "invalid_grant", "error_codes": [50173]}
        with pytest.raises(AuthError) as err:
            AuthSession(settings, app=app).access_token(force_refresh=True)
        assert err.value.code == "reauth_required"
        assert app.calls == ["silent:True"]

    def test_user_switch_under_same_session_is_refused(self, settings) -> None:
        app = FakeMsal({}, accounts=[{"home_account_id": f"{USER_A}.{TENANT}"}])
        app.silent = _result()
        session = AuthSession(settings, app=app)
        session.access_token()
        app.silent = _result(oid=USER_B, id_claims={"oid": USER_B})
        with pytest.raises(AuthError) as err:
            session.access_token()
        assert err.value.code == "user_mismatch"

    def test_sign_out_clears_accounts_and_cache_file(self, settings) -> None:
        settings.data_dir.mkdir(parents=True, exist_ok=True)
        settings.cache_path.write_bytes(b"encrypted")
        app = FakeMsal({}, accounts=[{"home_account_id": f"{USER_A}.{TENANT}"}])
        session = AuthSession(settings, app=app)
        session.sign_out()
        assert app.accounts == [] and not settings.cache_path.exists()
        assert session.identity() is None

    def test_identity_skips_msal_when_cache_empty(self, settings) -> None:
        """Signed-out serve/smoke must not hit Entra OIDC discovery."""
        settings.data_dir.mkdir(parents=True, exist_ok=True)

        class Boom:
            def __init__(self, *args: Any, **kwargs: Any) -> None:
                raise AssertionError("MSAL must not be constructed for an empty cache")

        session = AuthSession(settings)
        session._client = Boom  # type: ignore[method-assign]
        assert session.identity() is None
