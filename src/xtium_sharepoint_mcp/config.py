"""Connector settings, read from the environment the MCP host passes to the process."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path

from .errors import ConfigError

GUID_RE = re.compile(r"^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$")
MULTI_TENANT_ALIASES = {"common", "organizations", "consumers"}

GRAPH_RESOURCE = "https://graph.microsoft.com"
GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
GRAPH_BASE = f"{GRAPH_RESOURCE}/v1.0"
AUTHORITY_HOST = "https://login.microsoftonline.com"

SCOPE_USER_READ = "User.Read"
SCOPE_SITES_READ_ALL = "Sites.Read.All"
SCOPE_FILES_READ_ALL = "Files.Read.All"
DEFAULT_SCOPES = (SCOPE_USER_READ, SCOPE_SITES_READ_ALL, SCOPE_FILES_READ_ALL)
ALLOWED_SCOPES = frozenset(DEFAULT_SCOPES)


@dataclass(frozen=True)
class Limits:
    search_default_size: int = 10
    search_max_size: int = 25
    search_max_from: int = 200
    query_max_chars: int = 256
    folder_page_size: int = 50
    read_max_bytes: int = 10 * 1024 * 1024
    read_default_chars: int = 20_000
    read_max_chars: int = 50_000
    pdf_max_pages: int = 50
    http_timeout_seconds: float = 20.0


@dataclass(frozen=True)
class Settings:
    tenant_id: str
    client_id: str
    data_dir: Path
    scopes: tuple[str, ...] = DEFAULT_SCOPES
    audit_retention_days: int = 30
    limits: Limits = field(default_factory=Limits)

    @property
    def authority(self) -> str:
        return f"{AUTHORITY_HOST}/{self.tenant_id}"

    @property
    def graph_scopes(self) -> list[str]:
        return [f"{GRAPH_RESOURCE}/{s}" for s in self.scopes]

    @property
    def cache_path(self) -> Path:
        return self.data_dir / "token_cache.bin"

    @property
    def key_path(self) -> Path:
        return self.data_dir / "handle_key.bin"

    @property
    def audit_dir(self) -> Path:
        return self.data_dir / "audit"


def default_data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(Path.home(), ".local", "share")
    return Path(base) / "xtium-sharepoint-mcp"


def _guid(name: str, value: str | None) -> str:
    if not value:
        raise ConfigError("config_missing", f"{name} is required.")
    value = value.strip()
    if value.lower() in MULTI_TENANT_ALIASES:
        raise ConfigError("config_invalid", f"{name} must be a single tenant ID, not '{value}'.")
    if not GUID_RE.match(value):
        raise ConfigError("config_invalid", f"{name} must be a GUID.")
    return value.lower()


def _scopes(raw: str | None) -> tuple[str, ...]:
    if not raw:
        return DEFAULT_SCOPES
    scopes = tuple(dict.fromkeys(s for s in re.split(r"[\s,]+", raw.strip()) if s))
    unknown = [s for s in scopes if s not in ALLOWED_SCOPES]
    if unknown:
        raise ConfigError(
            "config_invalid",
            "Only delegated Microsoft Graph read scopes are allowed.",
            details={"allowed": sorted(ALLOWED_SCOPES), "rejected": unknown},
        )
    if SCOPE_USER_READ not in scopes:
        scopes = (SCOPE_USER_READ, *scopes)
    return scopes


def load_settings(env: dict[str, str] | None = None) -> Settings:
    env = dict(os.environ if env is None else env)
    data_dir = Path(env["XTIUM_SP_DATA_DIR"]) if env.get("XTIUM_SP_DATA_DIR") else default_data_dir()
    retention = int(env.get("XTIUM_SP_AUDIT_RETENTION_DAYS", "30"))
    return Settings(
        tenant_id=_guid("XTIUM_SP_TENANT_ID", env.get("XTIUM_SP_TENANT_ID")),
        client_id=_guid("XTIUM_SP_CLIENT_ID", env.get("XTIUM_SP_CLIENT_ID")),
        data_dir=data_dir,
        scopes=_scopes(env.get("XTIUM_SP_SCOPES")),
        audit_retention_days=max(1, min(retention, 365)),
    )
