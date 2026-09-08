# pst-tool — Outlook PST 解析ツール

Microsoft Outlook の `.pst` ファイルを解析するコマンドラインツールです。
[libpst](https://www.five-ten-sg.com/libpst/) の `readpst` を使って PST を
`.eml`（RFC822形式）のツリーに展開し、Python 標準の `email` モジュールで
フォルダ構成・送信者/受信者・添付ファイル・日付範囲などを集計します。

## ローカル完結・アップロード不要

このツールは **すべてPC上で完結**します。PSTファイルをどこかにアップロード
する必要は一切ありません（インターネット通信自体を行いません）。
`readpst` によるPSTの展開も、その後の集計も、すべてローカルディスク上の
処理です。数GB〜10GB超のPSTファイルでも、お使いのPC内だけで解析できます。

## できること

- **フォルダ別のメッセージ数**の集計
- **送信者/受信者ランキング**、**ドメイン別の送信者集計**
- **メッセージの日付範囲**と**月別の推移**
- **添付ファイルの一覧・拡張子別集計**、条件に合う添付ファイルの保存
- **件名・本文・差出人・宛先のキーワード検索**
- 結果を **Excel(.xlsx) / CSV / JSON / Markdown** で出力
  （.xlsxはシート分けされたレポートとして、そのままExcelで開けます）

## インストール不要な代替ツール（会社PCで管理者権限がない場合）

`readpst` のインストールや `pip install`、スクリプトの実行そのものが
社内ポリシーで禁止されている場合は、代わりに
[`web/mail-ledger/`](web/mail-ledger/) を使ってください。ブラウザで
`mail_ledger.html` を開くだけで動作する、インストール不要のツールです
（対象はPST全体ではなく、Outlookから個別にエクスポートした`.msg`ファイル）。

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

CSV / JSON / Markdown 出力だけなら追加の Python 依存パッケージは不要です
（標準ライブラリのみ）。**Excel(.xlsx) 出力**を使う場合のみ `openpyxl` が
必要です。

```bash
pip install -e .            # 基本機能のみ
pip install -e ".[xlsx]"    # Excel(.xlsx)出力も使う場合
# もしくはインストールせず python -m pst_tool ... で直接実行可能
```

## 使い方

### 解析レポートを出力する

```bash
pst-tool analyze path/to/mailbox.pst \
  --markdown-out report.md \
  --json-out report.json \
  --csv-out messages.csv \
  --xlsx-out report.xlsx \
  --top 20
```

出力先を何も指定しない場合は、Markdown形式のサマリーを標準出力に表示します。

`--xlsx-out` を指定すると、Excelでそのまま開けるレポート（`report.xlsx`）を
生成します。`Messages`（メッセージ一覧）、`Summary`（サマリー）、
`Folders` / `Top Senders` / `Top Recipients` / `Top Domains` / `Monthly` /
`Attachments by Ext` のシートに分かれています。

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

## 大容量PST（数GB〜10GB超）を扱う場合

7GBクラスのPSTでも動きますが、以下を意識すると快適です。

- **ディスク空き容量**: `readpst` はPSTの中身を `.eml` として展開するため、
  一時的に元のPSTと同程度（またはそれ以上）の空き容量が必要です。
  `--keep-extracted ./extracted` で展開先を指定し、SSD上の十分な空き容量
  がある場所を使ってください。
- **展開の並列化**: `--jobs 4` のようにCPUコア数に応じて指定すると、
  `readpst` によるPST展開が速くなります。
- **メモリ効率**: `analyze` コマンドは、メッセージ本文を保持せず集計する
  設計になっており、`--csv-out` / `--xlsx-out` を指定した場合もメッセージ
  一覧を1件ずつファイルへ書き出す（メモリに溜め込まない）ため、メッセージ
  数が数十万件規模でもメモリ使用量は増えにくくなっています。
- **展開結果の使い回し**: 一度 `extract` で展開しておけば、以後は
  `--extracted-dir` で再展開なしに何度でも解析・検索できます。7GBのPSTを
  毎回展開し直すのは時間がかかるため、大容量PSTでは特に有効です。

```bash
# 1. 一度だけ展開（時間がかかる）
pst-tool extract path/to/big-mailbox.pst --keep-extracted ./extracted --jobs 4

# 2. 以降は展開済みディレクトリを使い回して高速に解析・検索
pst-tool analyze --extracted-dir ./extracted --xlsx-out report.xlsx
pst-tool search --extracted-dir ./extracted "キーワード"
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
   各メッセージをパースして集計・検索を行います（大容量PST向けに、
   メッセージを1件ずつ処理してすぐ集計・書き出しに回す設計）。
3. `pst_tool.report` が集計結果を Excel(.xlsx) / CSV / JSON / Markdown に
   整形します。

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
