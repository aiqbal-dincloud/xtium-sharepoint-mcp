"""Typed connector errors. Messages are model-visible, so they never carry tokens or content."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any


@dataclass
class ConnectorError(Exception):
    code: str
    message: str
    status: int | None = None
    retry_after_seconds: int | None = None
    request_id: str | None = None
    client_request_id: str | None = None
    details: dict[str, Any] = field(default_factory=dict)

    def __str__(self) -> str:
        return json.dumps(self.to_dict(), sort_keys=True)

    def to_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"error": self.code, "message": self.message}
        if self.status is not None:
            out["status"] = self.status
        if self.retry_after_seconds is not None:
            out["retry_after_seconds"] = self.retry_after_seconds
        if self.request_id:
            out["request_id"] = self.request_id
        if self.client_request_id:
            out["client_request_id"] = self.client_request_id
        if self.details:
            out["details"] = self.details
        return out


class ConfigError(ConnectorError):
    pass


class AuthError(ConnectorError):
    pass


class GraphError(ConnectorError):
    pass


class HandleError(ConnectorError):
    pass


class ExtractionError(ConnectorError):
    pass
