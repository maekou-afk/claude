# pst-tool — Outlook PST 解析ツール

Microsoft Outlook の `.pst` ファイルを解析するコマンドラインツールです。
[libpst](https://www.five-ten-sg.com/libpst/) の `readpst` を使って PST を
`.eml`（RFC822形式）のツリーに展開し、Python 標準の `email` モジュールで
フォルダ構成・送信者/受信者・添付ファイル・日付範囲などを集計します。

## できること

- **フォルダ別のメッセージ数**の集計
- **送信者/受信者ランキング**、**ドメイン別の送信者集計**
- **メッセージの日付範囲**と**月別の推移**
- **添付ファイルの一覧・拡張子別集計**、条件に合う添付ファイルの保存
- **件名・本文・差出人・宛先のキーワード検索**
- 結果を **JSON / Markdown / CSV** で出力

内部処理は「PST展開」と「解析」を分離しているため、一度展開した結果
（`--extracted-dir`）を使い回して何度でも高速に再解析できます。

## インストール

### 1. readpst（libpst）のインストール

このツールは PST の展開に `readpst` コマンドを利用します。

```bash
# Debian / Ubuntu
sudo apt-get install pst-utils

# Fedora / RHEL
sudo dnf install libpst

# macOS (Homebrew)
brew install libpst
```

### 2. 本ツールのインストール

追加の Python 依存パッケージはありません（標準ライブラリのみ）。

```bash
pip install -e .
# もしくはインストールせず python -m pst_tool ... で直接実行可能
```

## 使い方

### 解析レポートを出力する

```bash
pst-tool analyze path/to/mailbox.pst \
  --markdown-out report.md \
  --json-out report.json \
  --csv-out messages.csv \
  --top 20
```

出力先を何も指定しない場合は、Markdown形式のサマリーを標準出力に表示します。

### PSTを.emlツリーに展開するだけ

展開結果を保存しておけば、以後は `--extracted-dir` で再展開なしに
何度でも `analyze` / `search` / `attachments` を実行できます。

```bash
pst-tool extract path/to/mailbox.pst --keep-extracted ./extracted
pst-tool analyze --extracted-dir ./extracted --json-out report.json
```

### キーワード検索

```bash
pst-tool search path/to/mailbox.pst "請求書" --field subject --field body_text
```

### 添付ファイルの一覧・抽出

```bash
# PDFかつ100KB以上の添付ファイルを一覧表示
pst-tool attachments path/to/mailbox.pst --ext pdf --min-size 100000

# 条件に一致した添付ファイルをディスクに保存
pst-tool attachments path/to/mailbox.pst --ext pdf --save-dir ./attachments
```

### 削除済みアイテムを含める / 並列展開

```bash
pst-tool analyze path/to/mailbox.pst --include-deleted --jobs 4
```

## 出力例（Markdown）

```
# PST 解析レポート: `./extracted`

- 総メッセージ数: **4**
- 解析エラー: 0
- 合計サイズ: 12.3 KB
- 期間: 2024-01-05T09:00:00+00:00 〜 2024-03-20T12:00:00+00:00
- 添付ファイル数: 2 (1.5 KB)

## フォルダ別メッセージ数

- Inbox: 2
- Inbox/Archive: 1
- Sent Items: 1
...
```

## 仕組み

1. `readpst -S -r -e -o <出力先> <PSTファイル>` で、PST内部のフォルダ構造を
   ディレクトリツリーとして再現しつつ、各メッセージを添付ファイル込みの
   `.eml` ファイルとして書き出します（ANSI/Unicode 両方のPST形式に対応）。
2. `pst_tool.analyzer` が `.eml` ツリーを走査し、`email` 標準ライブラリで
   各メッセージをパースして集計・検索を行います。
3. `pst_tool.report` が集計結果を JSON / Markdown / CSV に整形します。

展開処理（`extractor.py`）と解析処理（`analyzer.py`）を分離しているため、
テストは実際の `.pst` ファイルなしに、合成した `.eml` ツリーで解析ロジックを
検証しています（`tests/`）。

## テストの実行

```bash
pip install pytest
pytest
```

## 制限事項

- パスワード保護された PST は事前にパスワードを解除する必要があります
  （`readpst` 自体はOutlookのPSTパスワードに対応していません）。
- 破損した PST の一部は `readpst` の警告付きで部分的にしか復元できない
  ことがあります。
- 添付ファイルの保存機能は、同名・同種別の添付が同一メッセージ内に複数ある
  場合、最初に一致したものを保存します。
