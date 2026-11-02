"""Entry point: `xtium-sharepoint-mcp [serve|sign-in|sign-out|status]`."""

from __future__ import annotations

import argparse
import json
import logging
import sys

from .audit import AuditLog
from .auth import AuthSession, Identity
from .config import Settings, load_settings
from .errors import ConnectorError
from .graph import GraphClient
from .handles import HandleCodec, delete_key, load_or_create_key
from .service import SharePointService


def build(settings: Settings) -> tuple[AuthSession, SharePointService]:
    auth = AuthSession(settings)
    graph = GraphClient(auth, timeout=settings.limits.http_timeout_seconds)
    audit = AuditLog(
        settings.audit_dir,
        load_or_create_key(settings.data_dir / "audit_key.bin"),
        settings.tenant_id,
        settings.audit_retention_days,
    )

    def codec_for(identity: Identity) -> HandleCodec:
        return HandleCodec(load_or_create_key(settings.key_path), identity.tenant_id, identity.object_id)

    service = SharePointService(
        settings, auth, graph, codec_for, audit,
        rotate_handle_key=lambda: delete_key(settings.key_path),
    )
    return auth, service


def _print(obj: object) -> None:
    print(json.dumps(obj, indent=2, sort_keys=True))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="xtium-sharepoint-mcp")
    parser.add_argument("command", nargs="?", default="serve",
                        choices=["serve", "sign-in", "sign-out", "status"])
    args = parser.parse_args(argv)
    logging.basicConfig(stream=sys.stderr, level=logging.WARNING)

    try:
        settings = load_settings()
        auth, service = build(settings)
        if args.command == "serve":
            from .server import build_server

            server, _ = build_server(auth, service)
            server.run("stdio")
            return 0
        if args.command == "sign-in":
            try:
                identity = auth.sign_in()
            except ConnectorError as exc:
                service.sign_in_failed(exc.code)
                raise
            service.signed_in(identity)
            _print({"signed_in": True, "granted_scopes": sorted(identity.granted_scopes)})
        elif args.command == "sign-out":
            auth.sign_out()
            service.sign_out()
            _print({"signed_in": False})
        else:
            _print(service.connection_status())
        return 0
    except ConnectorError as exc:
        print(str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
