"""Analyze a tree of .eml files extracted from an Outlook .pst file."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Callable, Iterable, Iterator, Optional

_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"\s+")


def iter_eml_files(root: Path) -> Iterator[Path]:
    """Yield every .eml file under ``root``, sorted for deterministic output."""
    yield from sorted(Path(root).rglob("*.eml"))


def _folder_for(path: Path, root: Path) -> str:
    rel_parent = path.relative_to(root).parent
    return "(root)" if str(rel_parent) == "." else rel_parent.as_posix()


def _html_to_text(html: str) -> str:
    text = _TAG_RE.sub(" ", html)
    return _WS_RE.sub(" ", text).strip()


@dataclass
class Attachment:
    filename: str
    content_type: str
    size_bytes: int


@dataclass
class MessageRecord:
    path: Path
    folder: str
    subject: str = ""
    sender: str = ""
    sender_name: str = ""
    to: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    date: Optional[datetime] = None
    message_id: str = ""
    size_bytes: int = 0
    attachments: list[Attachment] = field(default_factory=list)
    body_text: str = ""
    parse_error: Optional[str] = None

    @property
    def has_attachments(self) -> bool:
        return bool(self.attachments)

    @property
    def sender_domain(self) -> str:
        return self.sender.split("@")[-1].lower() if "@" in self.sender else ""


def parse_eml(path: Path, root: Path, *, extract_body: bool = True) -> MessageRecord:
    """Parse one .eml file.

    ``extract_body`` controls whether the plain-text/HTML body is decoded
    and kept on the record. Folder/sender/date/attachment stats never
    require the body, so callers that only need aggregate counts (e.g. the
    ``analyze`` CLI command) should pass ``extract_body=False`` to avoid
    holding the full text of every message in memory at once -- important
    for multi-gigabyte PST files with hundreds of thousands of messages.
    Attachment *payloads* are never retained on the record either way (only
    filename/type/size), so attachments don't bloat memory here regardless.
    """
    folder = _folder_for(path, root)
    size_bytes = path.stat().st_size
    record = MessageRecord(path=path, folder=folder, size_bytes=size_bytes)

    try:
        with path.open("rb") as fh:
            msg = BytesParser(policy=policy.default).parse(fh)
    except Exception as exc:  # malformed message; keep basic stats, skip details
        record.parse_error = str(exc)
        return record

    record.subject = str(msg.get("Subject", "")).strip()
    record.message_id = str(msg.get("Message-ID", "")).strip()

    from_addrs = getaddresses(msg.get_all("From", []))
    if from_addrs:
        record.sender_name, record.sender = from_addrs[0]

    record.to = [addr for _, addr in getaddresses(msg.get_all("To", [])) if addr]
    record.cc = [addr for _, addr in getaddresses(msg.get_all("Cc", [])) if addr]

    date_hdr = msg.get("Date")
    if date_hdr:
        try:
            record.date = parsedate_to_datetime(date_hdr)
        except (TypeError, ValueError):
            record.date = None

    body_parts: list[str] = []
    for part in msg.walk():
        if part.is_multipart():
            continue
        disposition = part.get_content_disposition()
        filename = part.get_filename()
        if disposition == "attachment" or (filename and disposition != "inline"):
            payload = part.get_payload(decode=True) or b""
            record.attachments.append(
                Attachment(
                    filename=filename or "(unnamed)",
                    content_type=part.get_content_type(),
                    size_bytes=len(payload),
                )
            )
        elif not extract_body:
            continue
        elif part.get_content_type() == "text/plain" and not body_parts:
            try:
                body_parts.append(part.get_content())
            except Exception:
                pass
        elif part.get_content_type() == "text/html" and not body_parts:
            try:
                body_parts.append(_html_to_text(part.get_content()))
            except Exception:
                pass

    record.body_text = "\n".join(body_parts)
    return record


@dataclass
class AnalysisResult:
    root: Path
    total_messages: int = 0
    parse_errors: int = 0
    folder_counts: Counter = field(default_factory=Counter)
    sender_counts: Counter = field(default_factory=Counter)
    recipient_counts: Counter = field(default_factory=Counter)
    domain_counts: Counter = field(default_factory=Counter)
    monthly_counts: Counter = field(default_factory=Counter)
    attachment_extension_counts: Counter = field(default_factory=Counter)
    total_attachments: int = 0
    total_attachment_bytes: int = 0
    total_size_bytes: int = 0
    earliest: Optional[datetime] = None
    latest: Optional[datetime] = None
    messages: list[MessageRecord] = field(default_factory=list)

    def top(self, counter: Counter, n: int = 10):
        return counter.most_common(n)


def analyze(
    root: Path,
    *,
    keep_messages: bool = True,
    extract_body: bool = False,
    on_message: Optional[Callable[[MessageRecord], None]] = None,
) -> AnalysisResult:
    """Walk every .eml under ``root`` once, aggregating stats as it goes.

    For very large mailboxes, pass ``on_message`` (e.g. a CSV/XLSX row
    writer) and ``keep_messages=False`` to stream each record out to disk
    immediately instead of accumulating the full message list in memory.
    """
    root = Path(root)
    result = AnalysisResult(root=root)

    for eml_path in iter_eml_files(root):
        record = parse_eml(eml_path, root, extract_body=extract_body)
        result.total_messages += 1
        result.total_size_bytes += record.size_bytes
        result.folder_counts[record.folder] += 1

        if record.parse_error:
            result.parse_errors += 1
        else:
            if record.sender:
                result.sender_counts[record.sender] += 1
            if record.sender_domain:
                result.domain_counts[record.sender_domain] += 1
            for addr in record.to + record.cc:
                result.recipient_counts[addr] += 1
            if record.date:
                bucket = f"{record.date.year:04d}-{record.date.month:02d}"
                result.monthly_counts[bucket] += 1
                if result.earliest is None or record.date < result.earliest:
                    result.earliest = record.date
                if result.latest is None or record.date > result.latest:
                    result.latest = record.date
            for att in record.attachments:
                result.total_attachments += 1
                result.total_attachment_bytes += att.size_bytes
                ext = Path(att.filename).suffix.lower() or "(no extension)"
                result.attachment_extension_counts[ext] += 1

        if on_message:
            on_message(record)
        if keep_messages:
            result.messages.append(record)

    return result


def _field_value(record: MessageRecord, field_name: str, *, case_sensitive: bool) -> str:
    value = getattr(record, field_name, "")
    if isinstance(value, list):
        value = " ".join(value)
    return value if case_sensitive else value.lower()


def matches_keyword(
    record: MessageRecord,
    keyword: str,
    *,
    case_sensitive: bool = False,
    fields: tuple[str, ...] = ("subject", "body_text", "sender", "to"),
) -> bool:
    if not case_sensitive:
        keyword = keyword.lower()
    return any(keyword in _field_value(record, f, case_sensitive=case_sensitive) for f in fields)


def search_messages(
    messages: Iterable[MessageRecord],
    keyword: str,
    *,
    case_sensitive: bool = False,
    fields: tuple[str, ...] = ("subject", "body_text", "sender", "to"),
) -> list[MessageRecord]:
    return [
        record
        for record in messages
        if matches_keyword(record, keyword, case_sensitive=case_sensitive, fields=fields)
    ]
