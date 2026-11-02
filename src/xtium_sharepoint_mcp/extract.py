"""In-memory text extraction for PDF, DOCX and plain-text formats, within byte/page/char caps."""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from pathlib import PurePosixPath

from .config import Limits
from .errors import ExtractionError

TEXT_EXTENSIONS = {
    ".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".log", ".yaml", ".yml", ".html", ".htm",
}
SUPPORTED_EXTENSIONS = TEXT_EXTENSIONS | {".pdf", ".docx"}
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
MAX_DOCX_UNCOMPRESSED = 100 * 1024 * 1024


@dataclass
class Extracted:
    format: str
    text: str
    truncated: bool
    pages_read: int | None = None
    total_pages: int | None = None
    next_page: int | None = None


def extension(name: str) -> str:
    return PurePosixPath(name or "").suffix.lower()


def is_supported(name: str) -> bool:
    return extension(name) in SUPPORTED_EXTENSIONS


def _cap(text: str, max_chars: int) -> tuple[str, bool]:
    return (text, False) if len(text) <= max_chars else (text[:max_chars], True)


def _text(data: bytes, max_chars: int) -> Extracted:
    if data[:8192].count(b"\x00") > 16:
        raise ExtractionError("unsupported_format", "The file looks binary, not text.")
    try:
        decoded = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        decoded = data.decode("cp1252", errors="replace")
    text, truncated = _cap(decoded, max_chars)
    return Extracted("text", text, truncated)


def _pdf(data: bytes, max_chars: int, start_page: int, limits: Limits) -> Extracted:
    from pypdf import PdfReader
    from pypdf.errors import PdfReadError

    try:
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted and not reader.decrypt(""):
            raise ExtractionError("encrypted_file", "The PDF is password-protected.")
        total = len(reader.pages)
    except ExtractionError:
        raise
    except (PdfReadError, ValueError, KeyError, OSError) as exc:
        raise ExtractionError("corrupt_file", "The PDF could not be parsed.") from exc
    if start_page < 1 or (total and start_page > total):
        raise ExtractionError("bad_request", f"start_page must be between 1 and {total}.")

    parts: list[str] = []
    used = 0
    page = start_page
    last = min(total, start_page + limits.pdf_max_pages - 1)
    next_page: int | None = None
    while page <= last:
        try:
            chunk = reader.pages[page - 1].extract_text() or ""
        except Exception as exc:  # noqa: BLE001 - pypdf raises many types on malformed pages
            raise ExtractionError("corrupt_file", f"Page {page} could not be parsed.") from exc
        block = f"[page {page}]\n{chunk.strip()}\n"
        if used + len(block) > max_chars:
            parts.append(block[: max_chars - used])
            next_page = page
            break
        parts.append(block)
        used += len(block)
        page += 1
    else:
        next_page = page if page <= total else None
    pages_read = page - start_page + (1 if next_page == page and page <= last else 0)
    return Extracted(
        "pdf",
        "".join(parts),
        next_page is not None,
        pages_read=min(pages_read, total),
        total_pages=total,
        next_page=next_page,
    )


def _docx(data: bytes, max_chars: int) -> Extracted:
    if data.startswith(OLE_MAGIC):
        raise ExtractionError("encrypted_file", "The document is password-protected or IRM-encrypted.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            if sum(i.file_size for i in zf.infolist()) > MAX_DOCX_UNCOMPRESSED:
                raise ExtractionError("file_too_large", "The document expands beyond the safe limit.")
    except zipfile.BadZipFile as exc:
        raise ExtractionError("corrupt_file", "The document is not a valid DOCX file.") from exc

    import docx

    try:
        document = docx.Document(io.BytesIO(data))
    except Exception as exc:  # noqa: BLE001 - python-docx surfaces lxml/zip errors directly
        raise ExtractionError("corrupt_file", "The DOCX could not be parsed.") from exc
    lines = [p.text for p in document.paragraphs if p.text.strip()]
    for table in document.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                lines.append(" | ".join(cells))
    text, truncated = _cap("\n".join(lines), max_chars)
    return Extracted("docx", text, truncated)


def extract(name: str, data: bytes, limits: Limits, max_chars: int, start_page: int = 1) -> Extracted:
    ext = extension(name)
    if ext == ".pdf":
        return _pdf(data, max_chars, start_page, limits)
    if ext == ".docx":
        return _docx(data, max_chars)
    if ext in TEXT_EXTENSIONS:
        return _text(data, max_chars)
    raise ExtractionError(
        "unsupported_format",
        "Only PDF, DOCX and plain-text files can be read in this pilot.",
        details={"extension": ext or None},
    )
