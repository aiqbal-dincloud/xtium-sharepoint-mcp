"""Opaque handles: AES-GCM sealed Graph references bound to one tenant, one user and one kind.

The model only ever sees these strings. They cannot be turned into a Graph URL without the local
key, and a handle minted for user A fails to open for user B or after sign-out (key rotation).
"""

from __future__ import annotations

import base64
import contextlib
import json
import os
import re
from pathlib import Path
from typing import Any

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

from .errors import HandleError

KIND_PREFIX = {"file": "spf", "folder": "spd", "site": "sps", "cursor": "spc"}
MAX_HANDLE_CHARS = 2048

DRIVE_ID_RE = re.compile(r"^[A-Za-z0-9!_\-]{1,200}$")
ITEM_ID_RE = re.compile(r"^(root|[A-Za-z0-9!_\-]{1,200})$")
SITE_ID_RE = re.compile(r"^[A-Za-z0-9.\-]{1,253},[0-9a-fA-F\-]{36},[0-9a-fA-F\-]{36}$")
SKIPTOKEN_RE = re.compile(r"^[A-Za-z0-9!_\-=%.~+/]{1,1024}$")

_ID_RULES = {"d": DRIVE_ID_RE, "i": ITEM_ID_RE, "s": SITE_ID_RE, "k": SKIPTOKEN_RE}


def _b64e(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).rstrip(b"=").decode("ascii")


def _b64d(text: str) -> bytes:
    return base64.urlsafe_b64decode(text + "=" * (-len(text) % 4))


def validate_ids(payload: dict[str, Any]) -> dict[str, Any]:
    for key, rule in _ID_RULES.items():
        value = payload.get(key)
        if value is not None and (not isinstance(value, str) or not rule.match(value)):
            raise HandleError("invalid_handle", "The handle contains an invalid Graph reference.")
    return payload


class HandleCodec:
    def __init__(self, key: bytes, tenant_id: str, object_id: str):
        if len(key) != 32:
            raise ValueError("Handle key must be 32 bytes.")
        self._aead = AESGCM(key)
        self._tenant = tenant_id
        self._user = object_id

    def _aad(self, kind: str) -> bytes:
        return f"xsp1|{kind}|{self._tenant}|{self._user}".encode()

    def seal(self, kind: str, payload: dict[str, Any]) -> str:
        validate_ids(payload)
        nonce = os.urandom(12)
        body = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode()
        sealed = self._aead.encrypt(nonce, body, self._aad(kind))
        return f"{KIND_PREFIX[kind]}_{_b64e(nonce + sealed)}"

    def open(self, handle: str, kind: str) -> dict[str, Any]:
        prefix = KIND_PREFIX[kind] + "_"
        if not isinstance(handle, str) or len(handle) > MAX_HANDLE_CHARS:
            raise HandleError("invalid_handle", "The handle is missing or too long.")
        if not handle.startswith(prefix):
            other = next((k for k, p in KIND_PREFIX.items() if handle.startswith(p + "_")), None)
            if other:
                raise HandleError("wrong_handle_kind", f"Expected a {kind} handle, got a {other} one.")
            raise HandleError("invalid_handle", "This is not a connector handle.")
        try:
            raw = _b64d(handle[len(prefix):])
            body = self._aead.decrypt(raw[:12], raw[12:], self._aad(kind))
            payload = json.loads(body)
        except (InvalidTag, ValueError, json.JSONDecodeError) as exc:
            raise HandleError(
                "invalid_handle",
                "The handle is invalid, expired by sign-out, or belongs to another user or tenant.",
            ) from exc
        if not isinstance(payload, dict):
            raise HandleError("invalid_handle", "The handle payload is malformed.")
        return validate_ids(payload)


def load_or_create_key(path: Path) -> bytes:
    """32-byte key stored with DPAPI (Windows) through msal-extensions persistence."""
    from msal_extensions import build_encrypted_persistence
    from msal_extensions.persistence import PersistenceNotFound

    path.parent.mkdir(parents=True, exist_ok=True)
    persistence = build_encrypted_persistence(str(path))
    with contextlib.suppress(PersistenceNotFound):
        existing = persistence.load()
        if existing:
            key = bytes.fromhex(existing)
            if len(key) == 32:
                return key
    key = os.urandom(32)
    persistence.save(key.hex())
    return key


def delete_key(path: Path) -> None:
    with contextlib.suppress(FileNotFoundError):
        path.unlink()
