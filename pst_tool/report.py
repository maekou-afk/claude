"""Render AnalysisResult objects as JSON, Markdown, or CSV."""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Iterable

from .analyzer import AnalysisResult, MessageRecord


def _human_size(num_bytes: int) -> str:
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"


def to_dict(result: AnalysisResult, *, top_n: int = 10) -> dict:
    return {
        "root": str(result.root),
        "total_messages": result.total_messages,
        "parse_errors": result.parse_errors,
        "total_size_bytes": result.total_size_bytes,
        "total_size_human": _human_size(result.total_size_bytes),
        "date_range": {
            "earliest": result.earliest.isoformat() if result.earliest else None,
            "latest": result.latest.isoformat() if result.latest else None,
        },
        "folders": dict(result.folder_counts.most_common()),
        "top_senders": result.top(result.sender_counts, top_n),
        "top_recipients": result.top(result.recipient_counts, top_n),
        "top_domains": result.top(result.domain_counts, top_n),
        "messages_per_month": dict(sorted(result.monthly_counts.items())),
        "attachments": {
            "total_count": result.total_attachments,
            "total_bytes": result.total_attachment_bytes,
            "total_human": _human_size(result.total_attachment_bytes),
            "by_extension": dict(result.attachment_extension_counts.most_common()),
        },
    }


def to_json(result: AnalysisResult, *, top_n: int = 10, indent: int = 2) -> str:
    return json.dumps(to_dict(result, top_n=top_n), ensure_ascii=False, indent=indent)


def to_markdown(result: AnalysisResult, *, top_n: int = 10) -> str:
    d = to_dict(result, top_n=top_n)
    lines = [
        f"# PST 解析レポート: `{d['root']}`",
        "",
        f"- 総メッセージ数: **{d['total_messages']}**",
        f"- 解析エラー: {d['parse_errors']}",
        f"- 合計サイズ: {d['total_size_human']}",
        f"- 期間: {d['date_range']['earliest'] or 'N/A'} 〜 {d['date_range']['latest'] or 'N/A'}",
        f"- 添付ファイル数: {d['attachments']['total_count']} ({d['attachments']['total_human']})",
        "",
        "## フォルダ別メッセージ数",
        "",
    ]
    for folder, count in d["folders"].items():
        lines.append(f"- {folder}: {count}")

    lines += ["", f"## 送信者トップ {top_n}", ""]
    for sender, count in d["top_senders"]:
        lines.append(f"- {sender}: {count}")

    lines += ["", f"## 受信者トップ {top_n}", ""]
    for recipient, count in d["top_recipients"]:
        lines.append(f"- {recipient}: {count}")

    lines += ["", f"## ドメイン別送信者トップ {top_n}", ""]
    for domain, count in d["top_domains"]:
        lines.append(f"- {domain}: {count}")

    lines += ["", "## 月別メッセージ数", ""]
    for month, count in d["messages_per_month"].items():
        lines.append(f"- {month}: {count}")

    lines += ["", "## 添付ファイルの拡張子別件数", ""]
    for ext, count in d["attachments"]["by_extension"].items():
        lines.append(f"- {ext}: {count}")

    return "\n".join(lines) + "\n"


CSV_FIELDS = [
    "folder",
    "date",
    "sender",
    "sender_name",
    "to",
    "cc",
    "subject",
    "size_bytes",
    "attachment_count",
    "attachment_names",
    "path",
]


def write_csv(messages: Iterable[MessageRecord], out_path: Path) -> None:
    out_path = Path(out_path)
    with out_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for m in messages:
            writer.writerow(
                {
                    "folder": m.folder,
                    "date": m.date.isoformat() if m.date else "",
                    "sender": m.sender,
                    "sender_name": m.sender_name,
                    "to": "; ".join(m.to),
                    "cc": "; ".join(m.cc),
                    "subject": m.subject,
                    "size_bytes": m.size_bytes,
                    "attachment_count": len(m.attachments),
                    "attachment_names": "; ".join(a.filename for a in m.attachments),
                    "path": str(m.path),
                }
            )
