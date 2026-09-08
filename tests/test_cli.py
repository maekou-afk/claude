"""CLI smoke tests using --extracted-dir, so no real .pst or readpst binary
is required (extraction itself is exercised only via extractor tests)."""

from __future__ import annotations

import json
from email.message import EmailMessage
from pathlib import Path

from pst_tool.cli import main


def _write_eml(path: Path, subject: str, sender: str, to: str, date: str) -> None:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = to
    msg["Date"] = date
    msg.set_content("body text")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(bytes(msg))


def test_analyze_with_extracted_dir(tmp_path: Path, capsys):
    root = tmp_path / "extracted"
    _write_eml(
        root / "Inbox" / "1.eml",
        "Hello",
        "a@example.com",
        "b@example.com",
        "Mon, 05 Jan 2024 09:00:00 +0000",
    )
    json_out = tmp_path / "out.json"

    main(["analyze", "--extracted-dir", str(root), "--json-out", str(json_out)])

    data = json.loads(json_out.read_text(encoding="utf-8"))
    assert data["total_messages"] == 1
    assert data["folders"]["Inbox"] == 1

    captured = capsys.readouterr()
    assert "JSONレポート" in captured.out


def test_search_with_extracted_dir(tmp_path: Path, capsys):
    root = tmp_path / "extracted"
    _write_eml(
        root / "Inbox" / "1.eml",
        "Invoice attached",
        "a@example.com",
        "b@example.com",
        "Mon, 05 Jan 2024 09:00:00 +0000",
    )

    main(["search", "--extracted-dir", str(root), "invoice"])

    captured = capsys.readouterr()
    assert "1 件のメッセージ" in captured.out
