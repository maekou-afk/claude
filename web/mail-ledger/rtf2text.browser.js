"use strict";

// Naive but reasonably complete RTF -> plain text converter.
// Handles: skippable destination groups (fonttbl/colortbl/htmltag/pict/...),
// \par/\line/\tab, \'hh hex-escaped bytes decoded via the doc's ansicpg
// codepage, \uNNNN unicode escapes (with \ucN fallback-skip), and escaped
// literals \{ \} \\.

const SKIP_DESTINATIONS = new Set([
  "fonttbl", "colortbl", "stylesheet", "generator", "info", "pict",
  "object", "objdata", "htmltag", "mhtmltag", "template", "header",
  "headerl", "headerr", "headerf", "footer", "footerl", "footerr",
  "footerf", "footnote", "listtext", "revtbl", "rsidtbl", "xmlnstbl",
  "themedata", "colorschememapping", "latentstyles", "datastore",
  "companyname", "nonshppict", "bkmkstart", "bkmkend", "fldinst",
  "fldrslt", "shpinst", "shprslt", "sp", "wgrffmtfilter",
]);

const CODEPAGE_TO_ENCODING = {
  932: "shift-jis", 936: "gbk", 949: "euc-kr", 950: "big5",
  1250: "windows-1250", 1251: "windows-1251", 1252: "windows-1252",
  1253: "windows-1253", 1254: "windows-1254", 1255: "windows-1255",
  1256: "windows-1256", 1257: "windows-1257", 1258: "windows-1258",
  874: "windows-874", 65001: "utf-8",
};

function rtfToText(rtfBytes) {
  // rtfBytes: Uint8Array of the raw (decompressed) RTF.
  const bytes = rtfBytes;
  const n = bytes.length;
  let i = 0;
  let out = "";
  let ansicpg = 1252;
  let ucSkip = 1;
  const groupStack = []; // { skip: bool, ucSkip: number }
  let skipDepth = 0; // >0 while inside a skipped destination group

  // Buffer of raw bytes from consecutive \'hh escapes, flushed together so
  // multi-byte codepages (e.g. Shift-JIS) decode correctly.
  let hexBuf = [];
  function flushHexBuf() {
    if (hexBuf.length === 0) return;
    const enc = CODEPAGE_TO_ENCODING[ansicpg] || "windows-1252";
    try {
      out += new TextDecoder(enc).decode(new Uint8Array(hexBuf));
    } catch (e) {
      out += new TextDecoder("windows-1252").decode(new Uint8Array(hexBuf));
    }
    hexBuf = [];
  }

  function emit(str) {
    if (skipDepth > 0) return;
    flushHexBuf();
    out += str;
  }

  while (i < n) {
    const c = bytes[i];

    if (c === 0x7b) { // '{'
      flushHexBuf();
      groupStack.push({ skipDepth, ucSkip });
      i++;
      continue;
    }
    if (c === 0x7d) { // '}'
      flushHexBuf();
      const prev = groupStack.pop();
      if (prev) {
        skipDepth = prev.skipDepth;
        ucSkip = prev.ucSkip;
      }
      i++;
      continue;
    }
    if (c === 0x5c) { // '\'
      i++;
      if (i >= n) break;
      const next = bytes[i];

      if (next === 0x27) { // \'hh
        const hex = String.fromCharCode(bytes[i + 1], bytes[i + 2]);
        const byte = parseInt(hex, 16);
        if (!Number.isNaN(byte) && skipDepth === 0) hexBuf.push(byte);
        i += 3;
        continue;
      }
      if (next === 0x7b || next === 0x7d || next === 0x5c) {
        emit(String.fromCharCode(next));
        i++;
        continue;
      }
      if (next === 0x7e) { emit(" "); i++; continue; } // \~ nbsp
      if (next === 0x2d) { i++; continue; } // \- optional hyphen
      if (next === 0x0a || next === 0x0d) { i++; continue; }

      // control word: letters then optional signed digits, then one
      // optional space delimiter.
      if ((next >= 0x61 && next <= 0x7a) || (next >= 0x41 && next <= 0x5a)) {
        let wordStart = i;
        while (i < n && ((bytes[i] >= 0x61 && bytes[i] <= 0x7a) || (bytes[i] >= 0x41 && bytes[i] <= 0x5a))) i++;
        const word = String.fromCharCode.apply(null, bytes.slice(wordStart, i));
        let numStart = i;
        let neg = false;
        if (i < n && bytes[i] === 0x2d) { neg = true; i++; numStart = i; }
        while (i < n && bytes[i] >= 0x30 && bytes[i] <= 0x39) i++;
        let param = null;
        if (i > numStart) {
          const numStr = String.fromCharCode.apply(null, bytes.slice(numStart, i));
          param = parseInt(numStr, 10) * (neg ? -1 : 1);
        }
        if (i < n && bytes[i] === 0x20) i++; // trailing space delimiter is consumed

        if (word === "ansicpg" && param != null) ansicpg = param;
        else if (word === "uc" && param != null) ucSkip = param;
        else if (word === "u" && param != null) {
          let cp = param;
          if (cp < 0) cp += 65536;
          flushHexBuf();
          if (skipDepth === 0) out += String.fromCharCode(cp);
          // skip the \ucN fallback characters that follow
          let skipped = 0;
          while (skipped < ucSkip && i < n) {
            if (bytes[i] === 0x5c) {
              // a control word/symbol counts as a single fallback char
              i++;
              if (i < n && ((bytes[i] >= 0x61 && bytes[i] <= 0x7a) || (bytes[i] >= 0x41 && bytes[i] <= 0x5a))) {
                while (i < n && ((bytes[i] >= 0x61 && bytes[i] <= 0x7a) || (bytes[i] >= 0x41 && bytes[i] <= 0x5a))) i++;
                while (i < n && bytes[i] >= 0x30 && bytes[i] <= 0x39) i++;
                if (i < n && bytes[i] === 0x20) i++;
              } else if (i < n) {
                i++;
              }
            } else if (bytes[i] === 0x7b || bytes[i] === 0x7d) {
              break;
            } else {
              i++;
            }
            skipped++;
          }
        } else if (word === "par" || word === "line") {
          emit("\n");
        } else if (word === "tab") {
          emit("\t");
        } else if (word === "cell" || word === "row") {
          emit("\t");
        } else if (SKIP_DESTINATIONS.has(word)) {
          skipDepth++;
        }
        continue;
      }

      // control symbol (single non-letter char), e.g. \* \: \_
      if (next === 0x2a) { // \* — marks the *next* destination as ignorable
        // Peek ahead: if followed by \something in the skip set, mark skip.
        // Simple heuristic: just skip this whole group too (safe default
        // for \*\htmltag, \*\generator etc. commonly seen from Outlook).
        skipDepth++;
        i++;
        continue;
      }
      // unknown control symbol: consume one char, ignore
      i++;
      continue;
    }

    // plain literal byte
    if (skipDepth === 0) hexBuf.push(c);
    else flushHexBuf();
    i++;
  }
  flushHexBuf();

  return out
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

