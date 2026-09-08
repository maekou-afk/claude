"""Command-line interface for the Outlook PST analysis toolkit."""

from __future__ import annotations

import argparse
import re
import sys
import tempfile
from email import policy
from email.parser import BytesParser
from pathlib import Path

from . import analyzer, report
from .extractor import PstExtractionError, ReadPstNotFoundError, extract_pst

_UNSAFE_FILENAME_RE = re.compile(r"[^\w.\-]+")


def _safe_filename(name: str) -> str:
    name = Path(name).name or "attachment"
    return _UNSAFE_FILENAME_RE.sub("_", name)


def _resolve_eml_root(args) -> Path:
    """Return a directory of .eml files, extracting the PST first if needed."""
    if getattr(args, "extracted_dir", None):
        eml_root = Path(args.extracted_dir)
        if not eml_root.is_dir():
            print(f"エラー: --extracted-dir が見つかりません: {eml_root}", file=sys.stderr)
            sys.exit(1)
        return eml_root

    pst_path = Path(args.pst_file)
    if getattr(args, "keep_extracted", None):
        out_dir = Path(args.keep_extracted)
    else:
        out_dir = Path(tempfile.mkdtemp(prefix="pst_tool_"))

    try:
        extract_pst(
            pst_path,
            out_dir,
            include_deleted=args.include_deleted,
            jobs=args.jobs,
        )
    except ReadPstNotFoundError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        sys.exit(1)
    except (FileNotFoundError, PstExtractionError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        sys.exit(1)
    return out_dir


def cmd_extract(args) -> None:
    out_dir = _resolve_eml_root(args)
    print(f"PSTを展開しました: {out_dir}")


def cmd_analyze(args) -> None:
    eml_root = _resolve_eml_root(args)
    result = analyzer.analyze(eml_root)

    if result.total_messages == 0:
        print("警告: メッセージが見つかりませんでした。PSTファイルが空か、展開に失敗している可能性があります。", file=sys.stderr)

    if args.json_out:
        Path(args.json_out).write_text(report.to_json(result, top_n=args.top), encoding="utf-8")
        print(f"JSONレポートを書き出しました: {args.json_out}")
    if args.markdown_out:
        Path(args.markdown_out).write_text(report.to_markdown(result, top_n=args.top), encoding="utf-8")
        print(f"Markdownレポートを書き出しました: {args.markdown_out}")
    if args.csv_out:
        report.write_csv(result.messages, args.csv_out)
        print(f"CSVを書き出しました: {args.csv_out}")

    if not (args.json_out or args.markdown_out or args.csv_out):
        print(report.to_markdown(result, top_n=args.top))


def cmd_search(args) -> None:
    eml_root = _resolve_eml_root(args)
    result = analyzer.analyze(eml_root)
    fields = tuple(args.field) if args.field else ("subject", "body_text", "sender", "to")
    matches = analyzer.search_messages(
        result.messages, args.keyword, case_sensitive=args.case_sensitive, fields=fields
    )

    print(f"{len(matches)} 件のメッセージが '{args.keyword}' に一致しました。\n")
    for m in matches[: args.limit]:
        date_str = m.date.isoformat() if m.date else "不明"
        print(f"[{m.folder}] {date_str} | {m.sender} -> {', '.join(m.to)}")
        print(f"  件名: {m.subject}")
        print(f"  パス: {m.path}")
        print()


def cmd_attachments(args) -> None:
    eml_root = _resolve_eml_root(args)
    result = analyzer.analyze(eml_root)

    def matches_filters(att: analyzer.Attachment) -> bool:
        if args.ext and Path(att.filename).suffix.lower().lstrip(".") not in {
            e.lower().lstrip(".") for e in args.ext
        }:
            return False
        if args.min_size and att.size_bytes < args.min_size:
            return False
        return True

    save_dir = Path(args.save_dir) if args.save_dir else None
    if save_dir:
        save_dir.mkdir(parents=True, exist_ok=True)

    total = 0
    for m in result.messages:
        for att in m.attachments:
            if not matches_filters(att):
                continue
            total += 1
            print(f"[{m.folder}] {m.date} | {att.filename} ({att.size_bytes} bytes) <- {m.path}")
            if save_dir:
                _save_attachment(m.path, att, save_dir, total)

    print(f"\n合計 {total} 件の添付ファイルが条件に一致しました。")


def _save_attachment(eml_path: Path, target: analyzer.Attachment, save_dir: Path, index: int) -> None:
    with eml_path.open("rb") as fh:
        msg = BytesParser(policy=policy.default).parse(fh)
    for part in msg.walk():
        if part.is_multipart():
            continue
        filename = part.get_filename() or "(unnamed)"
        if filename == target.filename and part.get_content_type() == target.content_type:
            payload = part.get_payload(decode=True) or b""
            dest = save_dir / f"{index:05d}_{_safe_filename(filename)}"
            dest.write_bytes(payload)
            return


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pst-tool", description="Outlook PST ファイル解析ツール"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    def add_source_args(p: argparse.ArgumentParser, require_pst: bool = True) -> None:
        if require_pst:
            p.add_argument("pst_file", nargs="?", help="解析対象の .pst ファイル")
        p.add_argument(
            "--extracted-dir",
            help="readpst で既に展開済みの .eml ディレクトリを再利用する（再展開をスキップ）",
        )
        p.add_argument(
            "--keep-extracted",
            help="展開したファイルをこのディレクトリに保存する（未指定なら一時ディレクトリを使用）",
        )
        p.add_argument("--include-deleted", action="store_true", help="削除済みアイテムも含める")
        p.add_argument("--jobs", type=int, default=None, help="readpst の並列ジョブ数")

    p_extract = subparsers.add_parser("extract", help="PSTを.emlツリーに展開する")
    add_source_args(p_extract)
    p_extract.set_defaults(func=cmd_extract)

    p_analyze = subparsers.add_parser("analyze", help="PSTを解析してレポートを出力する")
    add_source_args(p_analyze)
    p_analyze.add_argument("--json-out", help="JSONレポートの出力先パス")
    p_analyze.add_argument("--markdown-out", help="Markdownレポートの出力先パス")
    p_analyze.add_argument("--csv-out", help="メッセージ一覧CSVの出力先パス")
    p_analyze.add_argument("--top", type=int, default=10, help="送信者/受信者ランキングの表示件数")
    p_analyze.set_defaults(func=cmd_analyze)

    p_search = subparsers.add_parser("search", help="件名・本文・差出人・宛先をキーワード検索する")
    add_source_args(p_search)
    p_search.add_argument("keyword", help="検索キーワード")
    p_search.add_argument("--case-sensitive", action="store_true")
    p_search.add_argument(
        "--field",
        action="append",
        choices=["subject", "body_text", "sender", "to"],
        help="検索対象フィールド（複数指定可、未指定なら全フィールド）",
    )
    p_search.add_argument("--limit", type=int, default=50, help="表示する最大件数")
    p_search.set_defaults(func=cmd_search)

    p_att = subparsers.add_parser("attachments", help="添付ファイルの一覧表示・抽出")
    add_source_args(p_att)
    p_att.add_argument("--ext", action="append", help="拡張子でフィルタ（複数指定可、例: --ext pdf --ext xlsx）")
    p_att.add_argument("--min-size", type=int, default=0, help="最小サイズ（バイト）でフィルタ")
    p_att.add_argument("--save-dir", help="条件に一致した添付ファイルを保存するディレクトリ")
    p_att.set_defaults(func=cmd_attachments)

    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.extracted_dir and not getattr(args, "pst_file", None):
        parser.error("pst_file または --extracted-dir のいずれかを指定してください")

    args.func(args)


if __name__ == "__main__":
    main()
