"""Tests for pst_tool.analyzer / pst_tool.report.

These tests build a synthetic tree of .eml files that mirrors the layout
``readpst -S -r -e`` produces (folder-per-directory, numbered .eml files),
so the analysis logic can be exercised without needing a real .pst file
or the readpst binary.
"""

from __future__ import annotations

from email.message import EmailMessage
from pathlib import Path

import pytest

from pst_tool import analyzer, report


def _write_eml(
    path: Path,
    *,
    subject: str,
    sender: str,
    to: list[str],
    date: str,
    body: str = "Hello",
    attachment: tuple[str, bytes] | None = None,
) -> None:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to)
    msg["Date"] = date
    msg["Message-ID"] = f"<{path.name}@example.com>"
    msg.set_content(body)
    if attachment:
        filename, data = attachment
        msg.add_attachment(data, maintype="application", subtype="octet-stream", filename=filename)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(bytes(msg))


@pytest.fixture()
def mailbox_root(tmp_path: Path) -> Path:
    root = tmp_path / "mailbox"
    _write_eml(
        root / "Inbox" / "1.eml",
        subject="Project kickoff",
        sender="Alice <alice@example.com>",
        to=["bob@example.com"],
        date="Mon, 05 Jan 2024 09:00:00 +0000",
        body="Let's kick off the project on Monday.",
    )
    _write_eml(
        root / "Inbox" / "2.eml",
        subject="Re: Project kickoff",
        sender="Bob <bob@example.com>",
        to=["alice@example.com"],
        date="Mon, 05 Jan 2024 10:30:00 +0000",
        body="Sounds good, see you then.",
        attachment=("agenda.pdf", b"%PDF-1.4 fake pdf content"),
    )
    _write_eml(
        root / "Inbox" / "Archive" / "3.eml",
        subject="Old newsletter",
        sender="news@example.org",
        to=["alice@example.com", "bob@example.com"],
        date="Fri, 15 Mar 2024 08:00:00 +0000",
        body="Monthly newsletter contents.",
        attachment=("report.xlsx", b"fake spreadsheet bytes" * 10),
    )
    _write_eml(
        root / "Sent Items" / "4.eml",
        subject="Budget approval",
        sender="alice@example.com",
        to=["finance@example.com"],
        date="Wed, 20 Mar 2024 12:00:00 +0000",
        body="Please approve the attached budget.",
    )
    return root


def test_iter_eml_files_sorted(mailbox_root: Path):
    files = list(analyzer.iter_eml_files(mailbox_root))
    assert len(files) == 4
    assert all(f.suffix == ".eml" for f in files)


def test_folder_naming(mailbox_root: Path):
    files = {f.name: f for f in analyzer.iter_eml_files(mailbox_root)}
    inbox_msg = analyzer.parse_eml(files["1.eml"], mailbox_root)
    archive_msg = analyzer.parse_eml(files["3.eml"], mailbox_root)
    sent_msg = analyzer.parse_eml(files["4.eml"], mailbox_root)
    assert inbox_msg.folder == "Inbox"
    assert archive_msg.folder == "Inbox/Archive"
    assert sent_msg.folder == "Sent Items"


def test_parse_eml_basic_fields(mailbox_root: Path):
    path = mailbox_root / "Inbox" / "1.eml"
    record = analyzer.parse_eml(path, mailbox_root)
    assert record.subject == "Project kickoff"
    assert record.sender == "alice@example.com"
    assert record.sender_name == "Alice"
    assert record.to == ["bob@example.com"]
    assert record.date is not None
    assert record.date.year == 2024
    assert "kick off" in record.body_text
    assert record.sender_domain == "example.com"
    assert not record.has_attachments


def test_parse_eml_attachment(mailbox_root: Path):
    path = mailbox_root / "Inbox" / "2.eml"
    record = analyzer.parse_eml(path, mailbox_root)
    assert record.has_attachments
    assert len(record.attachments) == 1
    att = record.attachments[0]
    assert att.filename == "agenda.pdf"
    assert att.size_bytes > 0


def test_analyze_aggregates(mailbox_root: Path):
    result = analyzer.analyze(mailbox_root)

    assert result.total_messages == 4
    assert result.parse_errors == 0
    assert result.folder_counts["Inbox"] == 2
    assert result.folder_counts["Inbox/Archive"] == 1
    assert result.folder_counts["Sent Items"] == 1

    assert result.sender_counts["alice@example.com"] == 2
    assert result.sender_counts["bob@example.com"] == 1

    assert result.domain_counts["example.com"] == 3
    assert result.domain_counts["example.org"] == 1

    # bob@example.com and finance@example.com each appear once as a To
    # recipient, alice@example.com appears as To on messages 2 and 3.
    assert result.recipient_counts["alice@example.com"] == 2
    assert result.recipient_counts["finance@example.com"] == 1

    assert result.total_attachments == 2
    assert result.total_attachment_bytes > 0
    assert result.attachment_extension_counts[".pdf"] == 1
    assert result.attachment_extension_counts[".xlsx"] == 1

    assert result.earliest.month == 1
    assert result.latest.month == 3
    assert result.monthly_counts["2024-01"] == 2
    assert result.monthly_counts["2024-03"] == 2


def test_search_messages_keyword(mailbox_root: Path):
    result = analyzer.analyze(mailbox_root)

    hits = analyzer.search_messages(result.messages, "budget")
    assert len(hits) == 1
    assert hits[0].subject == "Budget approval"

    hits_case = analyzer.search_messages(result.messages, "BUDGET", case_sensitive=True)
    assert len(hits_case) == 0

    hits_subject_only = analyzer.search_messages(
        result.messages, "newsletter", fields=("subject",)
    )
    assert len(hits_subject_only) == 1


def test_report_to_dict_and_markdown(mailbox_root: Path):
    result = analyzer.analyze(mailbox_root)
    d = report.to_dict(result, top_n=5)

    assert d["total_messages"] == 4
    assert d["folders"]["Inbox"] == 2
    assert d["attachments"]["total_count"] == 2

    md = report.to_markdown(result, top_n=5)
    assert "総メッセージ数" in md
    assert "Inbox" in md


def test_write_csv(mailbox_root: Path, tmp_path: Path):
    result = analyzer.analyze(mailbox_root)
    out_csv = tmp_path / "messages.csv"
    report.write_csv(result.messages, out_csv)

    content = out_csv.read_text(encoding="utf-8")
    assert "subject" in content.splitlines()[0]
    assert "Project kickoff" in content
    assert content.count("\n") >= 4


def test_analyze_streaming_does_not_retain_messages(mailbox_root: Path):
    """keep_messages=False + on_message must not build up an in-memory list

    -- this is what keeps memory flat when analyzing a multi-gigabyte PST
    with hundreds of thousands of messages."""
    streamed: list[analyzer.MessageRecord] = []
    result = analyzer.analyze(
        mailbox_root, keep_messages=False, on_message=streamed.append
    )

    assert result.messages == []
    assert len(streamed) == 4
    # aggregate stats are still computed correctly without keeping messages
    assert result.total_messages == 4
    assert result.folder_counts["Inbox"] == 2


def test_analyze_extract_body_default_false(mailbox_root: Path):
    result = analyzer.analyze(mailbox_root)
    assert all(m.body_text == "" for m in result.messages)

    result_with_body = analyzer.analyze(mailbox_root, extract_body=True)
    kickoff = next(m for m in result_with_body.messages if m.subject == "Project kickoff")
    assert "kick off" in kickoff.body_text


def test_matches_keyword_streaming_search(mailbox_root: Path):
    matches: list[analyzer.MessageRecord] = []

    def on_message(record: analyzer.MessageRecord) -> None:
        if analyzer.matches_keyword(record, "budget"):
            matches.append(record)

    analyzer.analyze(
        mailbox_root, keep_messages=False, extract_body=True, on_message=on_message
    )
    assert len(matches) == 1
    assert matches[0].subject == "Budget approval"


def test_write_xlsx(mailbox_root: Path, tmp_path: Path):
    pytest.importorskip("openpyxl")
    import openpyxl

    result = analyzer.analyze(mailbox_root)
    out_xlsx = tmp_path / "report.xlsx"
    report.write_xlsx(result, out_xlsx, top_n=5)

    wb = openpyxl.load_workbook(out_xlsx)
    assert "Messages" in wb.sheetnames
    assert "Summary" in wb.sheetnames
    assert "Folders" in wb.sheetnames

    messages_ws = wb["Messages"]
    rows = list(messages_ws.iter_rows(values_only=True))
    assert rows[0][0] == "folder"
    assert len(rows) == 1 + result.total_messages  # header + one row per message


def test_xlsx_report_writer_streaming(mailbox_root: Path, tmp_path: Path):
    pytest.importorskip("openpyxl")
    import openpyxl

    from pst_tool.report import XlsxReportWriter

    out_xlsx = tmp_path / "streamed.xlsx"
    writer = XlsxReportWriter(out_xlsx)

    def on_message(record: analyzer.MessageRecord) -> None:
        writer.write(record)

    result = analyzer.analyze(
        mailbox_root, keep_messages=False, extract_body=False, on_message=on_message
    )
    writer.close(result, top_n=5)

    wb = openpyxl.load_workbook(out_xlsx)
    rows = list(wb["Messages"].iter_rows(values_only=True))
    assert len(rows) == 1 + result.total_messages
