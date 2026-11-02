"""Local JSONL audit log. Identifiers are keyed hashes; no tokens, names, queries or content."""

from __future__ import annotations

import contextlib
import hashlib
import hmac
import json
import threading
from collections.abc import Callable, Iterable
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any


class AuditLog:
    def __init__(
        self,
        directory: Path,
        key: bytes,
        tenant_id: str,
        retention_days: int = 30,
        clock: Callable[[], datetime] = lambda: datetime.now(UTC),
    ):
        self._dir = directory
        self._key = key
        self._tenant = tenant_id
        self._retention = timedelta(days=retention_days)
        self._clock = clock
        self._lock = threading.Lock()
        self._dir.mkdir(parents=True, exist_ok=True)
        self.prune()

    def hash_id(self, value: str | None) -> str | None:
        if not value:
            return None
        return hmac.new(self._key, b"audit|" + value.encode(), hashlib.sha256).hexdigest()[:24]

    def prune(self) -> None:
        cutoff = (self._clock() - self._retention).date()
        for path in self._dir.glob("audit-*.jsonl"):
            with contextlib.suppress(ValueError, OSError):
                if datetime.strptime(path.stem[6:], "%Y-%m-%d").date() < cutoff:
                    path.unlink()

    def record(
        self,
        *,
        event: str,
        outcome: str,
        actor_oid: str | None = None,
        resource_ids: Iterable[str] = (),
        error_code: str | None = None,
        duration_ms: int | None = None,
        request_id: str | None = None,
        client_request_id: str | None = None,
        counts: dict[str, int] | None = None,
    ) -> dict[str, Any]:
        now = self._clock()
        entry: dict[str, Any] = {
            "ts": now.isoformat(timespec="milliseconds"),
            "event": event,
            "outcome": outcome,
            "tenant_id": self._tenant,
            "actor": self.hash_id(actor_oid),
            "resources": [h for h in (self.hash_id(r) for r in resource_ids) if h],
        }
        for key, value in (
            ("error", error_code),
            ("duration_ms", duration_ms),
            ("graph_request_id", request_id),
            ("client_request_id", client_request_id),
            ("counts", counts),
        ):
            if value is not None:
                entry[key] = value
        line = json.dumps(entry, separators=(",", ":"), sort_keys=True)
        with self._lock:
            path = self._dir / f"audit-{now:%Y-%m-%d}.jsonl"
            with path.open("a", encoding="utf-8") as fh:
                fh.write(line + "\n")
        return entry
