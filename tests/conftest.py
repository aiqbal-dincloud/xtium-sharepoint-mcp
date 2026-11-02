from __future__ import annotations

import base64
import io
import json
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import httpx
import pytest

from xtium_sharepoint_mcp.audit import AuditLog
from xtium_sharepoint_mcp.auth import Identity
from xtium_sharepoint_mcp.config import Settings
from xtium_sharepoint_mcp.graph import GraphClient
from xtium_sharepoint_mcp.handles import HandleCodec
from xtium_sharepoint_mcp.service import SharePointService

FIXTURES = Path(__file__).parent / "fixtures"
TENANT = "6d21634d-ca87-4a22-99eb-e31240eeea86"
OTHER_TENANT = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
CLIENT = "11112222-3333-4444-5555-666677778888"
USER_A = "0a0a0a0a-0000-0000-0000-00000000000a"
USER_B = "0b0b0b0b-0000-0000-0000-00000000000b"
FIN_SITE = "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222"
ALL_SCOPES = frozenset({"User.Read", "Sites.Read.All", "Files.Read.All"})


def fixture(name: str) -> dict[str, Any]:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def b64(obj: dict[str, Any]) -> str:
    return base64.urlsafe_b64encode(json.dumps(obj).encode()).rstrip(b"=").decode()


def jwt(**claims: Any) -> str:
    return f"{b64({'alg': 'RS256', 'typ': 'JWT'})}.{b64(claims)}.c2ln"


def graph_token(**overrides: Any) -> str:
    claims = {
        "aud": "https://graph.microsoft.com",
        "tid": TENANT,
        "appid": CLIENT,
        "oid": USER_A,
        "scp": "Files.Read.All Sites.Read.All User.Read profile openid email",
    }
    claims.update(overrides)
    return jwt(**{k: v for k, v in claims.items() if v is not None})


def id_claims(**overrides: Any) -> dict[str, Any]:
    claims = {
        "aud": CLIENT,
        "tid": TENANT,
        "iss": f"https://login.microsoftonline.com/{TENANT}/v2.0",
        "oid": USER_A,
    }
    claims.update(overrides)
    return claims


@pytest.fixture
def settings(tmp_path: Path) -> Settings:
    return Settings(tenant_id=TENANT, client_id=CLIENT, data_dir=tmp_path)


@dataclass
class FakeTokens:
    user: str | None = USER_A
    scopes: frozenset[str] = ALL_SCOPES
    refreshes: int = 0

    def identity(self) -> Identity | None:
        return Identity(TENANT, self.user, self.scopes) if self.user else None

    def access_token(self, force_refresh: bool = False) -> str:
        if force_refresh:
            self.refreshes += 1
        return f"token-{self.user}-{self.refreshes}"


@dataclass
class FakeGraph:
    """Route table for httpx.MockTransport keyed by (method, path)."""

    routes: dict[tuple[str, str], Callable[[httpx.Request], httpx.Response]] = field(default_factory=dict)
    requests: list[httpx.Request] = field(default_factory=list)

    def on(self, method: str, path: str, handler: Callable[[httpx.Request], httpx.Response] | dict) -> None:
        if isinstance(handler, dict):
            body = handler
            handler = lambda _req, body=body: httpx.Response(200, json=body)  # noqa: E731
        self.routes[(method, path)] = handler

    def __call__(self, request: httpx.Request) -> httpx.Response:
        self.requests.append(request)
        key = (request.method, request.url.host + request.url.path)
        if key not in self.routes:
            key = (request.method, request.url.path)
        if key not in self.routes:
            return httpx.Response(404, json={"error": {"code": "itemNotFound", "message": "nope"}})
        return self.routes[key](request)


@dataclass
class Env:
    settings: Settings
    tokens: FakeTokens
    graph: FakeGraph
    client: GraphClient
    service: SharePointService
    audit_dir: Path
    keys: dict[str, bytes]

    def codec(self, user: str = USER_A, tenant: str = TENANT) -> HandleCodec:
        return HandleCodec(self.keys["handle"], tenant, user)

    def audit_lines(self) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        for path in sorted(self.audit_dir.glob("*.jsonl")):
            out += [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]
        return out


@pytest.fixture
def env(settings: Settings) -> Env:
    tokens = FakeTokens()
    graph = FakeGraph()
    client = GraphClient(tokens, transport=httpx.MockTransport(graph))
    keys = {"handle": b"\x01" * 32}

    def codec_for(identity: Identity) -> HandleCodec:
        return HandleCodec(keys["handle"], identity.tenant_id, identity.object_id)

    def rotate() -> None:
        keys["handle"] = bytes([keys["handle"][0] + 1]) * 32

    audit = AuditLog(settings.audit_dir, b"\x02" * 32, TENANT)
    service = SharePointService(settings, tokens, client, codec_for, audit, rotate_handle_key=rotate)
    return Env(settings, tokens, graph, client, service, settings.audit_dir, keys)


def make_docx(*paragraphs: str) -> bytes:
    import docx

    document = docx.Document()
    for p in paragraphs:
        document.add_paragraph(p)
    table = document.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "Total"
    table.rows[0].cells[1].text = "42"
    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue()


def make_pdf(pages: list[str]) -> bytes:
    """Minimal text PDF with one Helvetica text line per page."""
    objects: list[bytes] = []
    n = len(pages)
    kids = " ".join(f"{3 + i * 2} 0 R" for i in range(n))
    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    objects.append(f"<< /Type /Pages /Kids [{kids}] /Count {n} >>".encode())
    font_id = 3 + n * 2
    for i, text in enumerate(pages):
        content_id = 4 + i * 2
        objects.append(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents {content_id} 0 R "
            f"/Resources << /Font << /F1 {font_id} 0 R >> >> >>".encode()
        )
        stream = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode()
        objects.append(b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream")
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

    out = io.BytesIO()
    out.write(b"%PDF-1.4\n")
    offsets = []
    for idx, body in enumerate(objects, start=1):
        offsets.append(out.tell())
        out.write(f"{idx} 0 obj\n".encode() + body + b"\nendobj\n")
    xref = out.tell()
    out.write(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode())
    for off in offsets:
        out.write(f"{off:010d} 00000 n \n".encode())
    out.write(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode())
    return out.getvalue()
