#!/usr/bin/env bash
# Rebuilds mail_ledger.html from page_top.html + the bundled msgreader
# library + rtf2text.browser.js + app.js.
#
# Usage:
#   cd web/mail-ledger
#   npm install
#   npm run build
set -euo pipefail
cd "$(dirname "$0")"

npx esbuild build.entry.js \
  --bundle --minify --platform=browser --format=iife \
  --define:process.env.NODE_ENV='"production"' \
  --outfile=msgreader.browser.min.js

{
  cat page_top.html
  echo "<script>"
  cat msgreader.browser.min.js
  echo ""
  echo "</script>"
  echo "<script>"
  cat rtf2text.browser.js
  echo ""
  cat app.js
  echo "</script>"
} > mail_ledger.html

rm -f msgreader.browser.min.js
echo "Built mail_ledger.html ($(wc -c < mail_ledger.html) bytes)"
