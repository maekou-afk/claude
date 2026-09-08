"use strict";

/* ---------- state ---------- */

const state = {
  records: [],
  loaded: false,
  loading: false,
  processed: 0,
  total: 0,
  errors: 0,
  sortKey: "date",
  sortDir: "desc",
  query: "",
};

const SAMPLE_RECORDS = [
  {
    folder: "受信トレイ/取引先",
    subject: "【サンプル】見積書のご確認をお願いいたします",
    senderName: "山田 花子",
    senderEmail: "hanako.yamada@example.com",
    to: ["経理部 御中 <keiri@example.co.jp>"],
    cc: [],
    date: new Date("2025-11-04T10:12:00+09:00"),
    sizeBytes: 48200,
    attachments: [{ name: "見積書_2025年11月.xlsx", size: 31840 }],
    bodyPreview: "いつもお世話になっております。標記の件、見積書を添付いたします。",
    isSample: true,
  },
  {
    folder: "受信トレイ",
    subject: "【サンプル】週次定例のリマインド",
    senderName: "システム通知",
    senderEmail: "noreply@example.com",
    to: ["営業部一同 <sales-all@example.co.jp>"],
    cc: ["田中 <tanaka@example.co.jp>"],
    date: new Date("2025-11-03T09:00:00+09:00"),
    sizeBytes: 6400,
    attachments: [],
    bodyPreview: "本日15時より週次定例を行います。会議室Aにお集まりください。",
    isSample: true,
  },
  {
    folder: "送信済みアイテム",
    subject: "【サンプル】Re: 契約更新について",
    senderName: "自分",
    senderEmail: "me@example.co.jp",
    to: ["佐藤 次郎 <jiro.sato@example.com>"],
    cc: [],
    date: new Date("2025-10-28T17:45:00+09:00"),
    sizeBytes: 12100,
    attachments: [{ name: "契約書ドラフト_v3.docx", size: 55200 }],
    bodyPreview: "ご連絡ありがとうございます。契約更新の件、以下の通り進めたく存じます。",
    isSample: true,
  },
  {
    folder: "受信トレイ/請求関連",
    subject: "【サンプル】9月分ご請求書送付のご案内",
    senderName: "経理担当",
    senderEmail: "billing@example-partner.com",
    to: ["me@example.co.jp"],
    cc: [],
    date: new Date("2025-10-01T11:20:00+09:00"),
    sizeBytes: 89000,
    attachments: [{ name: "invoice_202509.pdf", size: 74000 }],
    bodyPreview: "9月分のご請求書を送付いたします。ご確認のほどよろしくお願いいたします。",
    isSample: true,
  },
];

/* ---------- dom refs ---------- */

const el = {
  fileInput: document.getElementById("fileInput"),
  folderInput: document.getElementById("folderInput"),
  dropZone: document.getElementById("dropZone"),
  progressWrap: document.getElementById("progressWrap"),
  progressBar: document.getElementById("progressBar"),
  progressLabel: document.getElementById("progressLabel"),
  statsRail: document.getElementById("statsRail"),
  folderList: document.getElementById("folderList"),
  senderList: document.getElementById("senderList"),
  domainList: document.getElementById("domainList"),
  searchInput: document.getElementById("searchInput"),
  ledgerBody: document.getElementById("ledgerBody"),
  ledgerCount: document.getElementById("ledgerCount"),
  exportBtn: document.getElementById("exportBtn"),
  sampleBadge: document.getElementById("sampleBadge"),
  ledgerTable: document.getElementById("ledgerTable"),
  clearBtn: document.getElementById("clearBtn"),
  pickFilesBtn: document.getElementById("pickFilesBtn"),
  pickFolderBtn: document.getElementById("pickFolderBtn"),
};

/* ---------- utils ---------- */

function escapeHtml(str) {
  return String(str == null ? "" : str).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function formatBytes(n) {
  if (!n && n !== 0) return "-";
  const units = ["B", "KB", "MB", "GB"];
  let v = n, i = 0;
  while (v >= 1024 && i < units.length - 1) { v /= 1024; i++; }
  return `${v.toFixed(v >= 10 || i === 0 ? 0 : 1)} ${units[i]}`;
}

function formatDate(d) {
  if (!d || Number.isNaN(d.getTime())) return "(日付不明)";
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

function monthKey(d) {
  if (!d || Number.isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function domainOf(email) {
  if (!email || email.indexOf("@") === -1) return null;
  return email.split("@").pop().toLowerCase();
}

function debounce(fn, ms) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

function yieldToUI() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

/* ---------- .msg parsing ---------- */

function pickDate(fileData) {
  const raw = fileData.clientSubmitTime || fileData.messageDeliveryTime || fileData.creationTime;
  if (!raw) return null;
  const d = new Date(raw);
  return Number.isNaN(d.getTime()) ? null : d;
}

function extractBody(fileData) {
  const plain = (fileData.body || "").trim();
  if (plain) return plain;
  if (fileData.compressedRtf && fileData.compressedRtf.length) {
    try {
      const decompressed = window.decompressRTF(Array.from(fileData.compressedRtf));
      return rtfToText(new Uint8Array(decompressed));
    } catch (e) {
      return "";
    }
  }
  return "";
}

function folderFromPath(file) {
  const rel = file.webkitRelativePath || "";
  if (!rel) return "(フォルダ指定なし)";
  const parts = rel.split("/");
  parts.pop(); // filename
  parts.shift(); // top-level picked folder name itself
  return parts.length ? parts.join("/") : "(ルート)";
}

async function parseMsgFile(file) {
  const record = {
    fileName: file.name,
    folder: folderFromPath(file),
    subject: "", senderName: "", senderEmail: "",
    to: [], cc: [], date: null, sizeBytes: file.size,
    attachments: [], bodyPreview: "", error: null, isSample: false,
  };
  try {
    const buf = await file.arrayBuffer();
    const reader = new window.MsgReader(buf);
    const fileData = reader.getFileData();
    if (fileData.error) throw new Error(fileData.error);

    record.subject = fileData.subject || "(件名なし)";
    record.senderName = fileData.senderName || "";
    record.senderEmail = fileData.senderEmail && fileData.senderEmail.includes("@")
      ? fileData.senderEmail : "";
    const recipients = fileData.recipients || [];
    for (const r of recipients) {
      const label = r.name || r.smtpAddress || r.email || "";
      if (r.recipType === "cc") record.cc.push(label);
      else if (r.recipType !== "bcc") record.to.push(label);
    }
    record.date = pickDate(fileData);
    record.attachments = (fileData.attachments || []).map((a) => ({
      name: a.fileName || a.name || "(無題)",
      size: a.contentLength || 0,
    }));
    record.bodyPreview = extractBody(fileData).slice(0, 2000);
  } catch (e) {
    record.error = e && e.message ? e.message : String(e);
  }
  return record;
}

/* ---------- ingest ---------- */

async function handleFiles(fileList) {
  const files = Array.from(fileList).filter((f) => /\.msg$/i.test(f.name));
  if (files.length === 0) return;

  state.loading = true;
  state.total = files.length;
  state.processed = 0;
  state.errors = 0;
  el.progressWrap.hidden = false;
  el.exportBtn.disabled = true;

  const newRecords = [];
  for (const file of files) {
    const record = await parseMsgFile(file);
    if (record.error) state.errors++;
    newRecords.push(record);
    state.processed++;
    if (state.processed % 15 === 0 || state.processed === state.total) {
      updateProgress();
      await yieldToUI();
    }
  }

  state.records = state.records.filter((r) => !r.isSample).concat(newRecords);
  state.loaded = true;
  state.loading = false;
  el.progressWrap.hidden = true;
  el.exportBtn.disabled = state.records.length === 0;
  el.sampleBadge.hidden = true;
  el.clearBtn.hidden = false;
  render();
}

function updateProgress() {
  const pct = state.total ? Math.round((state.processed / state.total) * 100) : 0;
  el.progressBar.style.width = pct + "%";
  el.progressLabel.textContent = `${state.processed} / ${state.total} 件を解析中…（エラー ${state.errors} 件）`;
}

function clearAll() {
  state.records = [];
  state.loaded = false;
  el.exportBtn.disabled = true;
  el.sampleBadge.hidden = false;
  el.clearBtn.hidden = true;
  el.searchInput.value = "";
  state.query = "";
  render();
}

/* ---------- aggregation ---------- */

function aggregate(records) {
  const real = records.filter((r) => !r.error);
  const folderCounts = new Map();
  const senderCounts = new Map();
  const domainCounts = new Map();
  const monthCounts = new Map();
  let attachTotal = 0, attachBytes = 0;
  let earliest = null, latest = null;

  for (const r of real) {
    folderCounts.set(r.folder, (folderCounts.get(r.folder) || 0) + 1);
    const senderKey = r.senderEmail || r.senderName || "(不明な送信者)";
    senderCounts.set(senderKey, (senderCounts.get(senderKey) || 0) + 1);
    const domain = domainOf(r.senderEmail);
    if (domain) domainCounts.set(domain, (domainCounts.get(domain) || 0) + 1);
    if (r.date) {
      const mk = monthKey(r.date);
      if (mk) monthCounts.set(mk, (monthCounts.get(mk) || 0) + 1);
      if (!earliest || r.date < earliest) earliest = r.date;
      if (!latest || r.date > latest) latest = r.date;
    }
    for (const a of r.attachments) {
      attachTotal++;
      attachBytes += a.size || 0;
    }
  }

  const sortDesc = (map) => [...map.entries()].sort((a, b) => b[1] - a[1]);

  return {
    total: records.length,
    errors: records.length - real.length,
    folderCounts: sortDesc(folderCounts),
    senderCounts: sortDesc(senderCounts),
    domainCounts: sortDesc(domainCounts),
    monthCounts: [...monthCounts.entries()].sort((a, b) => a[0].localeCompare(b[0])),
    attachTotal,
    attachBytes,
    earliest,
    latest,
  };
}

/* ---------- render ---------- */

function currentRecords() {
  return state.records.length ? state.records : SAMPLE_RECORDS;
}

function matchesQuery(r, q) {
  if (!q) return true;
  const hay = [r.subject, r.senderName, r.senderEmail, r.bodyPreview, ...r.to, ...r.cc]
    .join(" ").toLowerCase();
  return hay.includes(q);
}

function sortRecords(records) {
  const dir = state.sortDir === "asc" ? 1 : -1;
  const key = state.sortKey;
  return [...records].sort((a, b) => {
    let av, bv;
    if (key === "date") { av = a.date ? a.date.getTime() : -Infinity; bv = b.date ? b.date.getTime() : -Infinity; }
    else if (key === "size") { av = a.sizeBytes; bv = b.sizeBytes; }
    else if (key === "sender") { av = (a.senderName || a.senderEmail || "").toLowerCase(); bv = (b.senderName || b.senderEmail || "").toLowerCase(); }
    else { av = (a.subject || "").toLowerCase(); bv = (b.subject || "").toLowerCase(); }
    if (av < bv) return -1 * dir;
    if (av > bv) return 1 * dir;
    return 0;
  });
}

const RENDER_CAP = 400;

function render() {
  const records = currentRecords();
  const agg = aggregate(records);
  renderStats(agg, records.length);
  renderLists(agg);
  renderLedger(records);
}

function renderStats(agg, total) {
  const range = agg.earliest && agg.latest
    ? `${formatDate(agg.earliest).slice(0, 10)} 〜 ${formatDate(agg.latest).slice(0, 10)}`
    : "-";
  const tiles = [
    ["総メッセージ数", total.toLocaleString()],
    ["解析エラー", agg.errors.toLocaleString()],
    ["期間", range],
    ["送信者数", agg.senderCounts.length.toLocaleString()],
    ["フォルダ数", agg.folderCounts.length.toLocaleString()],
    ["添付ファイル", `${agg.attachTotal.toLocaleString()}件 / ${formatBytes(agg.attachBytes)}`],
  ];
  el.statsRail.innerHTML = tiles.map(([label, value]) => `
    <div class="tile">
      <div class="tile-label">${escapeHtml(label)}</div>
      <div class="tile-value">${escapeHtml(value)}</div>
    </div>
  `).join("");
}

function renderLists(agg) {
  const renderRows = (entries, cap = 8) => entries.slice(0, cap).map(([name, count]) => `
    <li><span class="li-name">${escapeHtml(name)}</span><span class="li-count">${count}</span></li>
  `).join("") || `<li class="li-empty">データなし</li>`;

  el.folderList.innerHTML = renderRows(agg.folderCounts);
  el.senderList.innerHTML = renderRows(agg.senderCounts);
  el.domainList.innerHTML = renderRows(agg.domainCounts);
}

function renderLedger(records) {
  const q = state.query.trim().toLowerCase();
  const filtered = records.filter((r) => matchesQuery(r, q));
  const sorted = sortRecords(filtered);
  const shown = sorted.slice(0, RENDER_CAP);

  el.ledgerCount.textContent = filtered.length > RENDER_CAP
    ? `${filtered.length.toLocaleString()} 件中 ${RENDER_CAP} 件を表示（検索で絞り込めます）`
    : `${filtered.length.toLocaleString()} 件`;

  if (shown.length === 0) {
    el.ledgerBody.innerHTML = `<tr><td colspan="6" class="empty-row">一致するメッセージがありません</td></tr>`;
    return;
  }

  el.ledgerBody.innerHTML = shown.map((r) => {
    const from = r.senderName || r.senderEmail || "(不明)";
    const to = r.to.join(", ") || "-";
    const att = r.attachments.length
      ? `<span class="att-badge" title="${escapeHtml(r.attachments.map((a) => a.name).join(", "))}">📎 ${r.attachments.length}</span>`
      : "";
    const errClass = r.error ? " row-error" : "";
    const rowTitle = r.error ? ` title="解析エラー: ${escapeHtml(r.error)}"` : "";
    return `
      <tr class="${errClass.trim()}"${rowTitle}>
        <td class="col-date">${formatDate(r.date)}</td>
        <td class="col-folder">${escapeHtml(r.folder)}</td>
        <td class="col-subject">${escapeHtml(r.subject || (r.error ? "(解析失敗)" : "(件名なし)"))} ${att}</td>
        <td class="col-from">${escapeHtml(from)}</td>
        <td class="col-to">${escapeHtml(to)}</td>
        <td class="col-size">${formatBytes(r.sizeBytes)}</td>
      </tr>
    `;
  }).join("");
}

/* ---------- CSV export ---------- */

function toCsvValue(v) {
  const s = String(v == null ? "" : v);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function buildCsv(records) {
  const header = ["folder", "date", "sender_name", "sender_email", "to", "cc", "subject", "size_bytes", "attachment_count", "attachment_names", "file_name", "parse_error"];
  const lines = [header.join(",")];
  for (const r of records) {
    lines.push([
      r.folder, r.date ? r.date.toISOString() : "", r.senderName, r.senderEmail,
      r.to.join("; "), r.cc.join("; "), r.subject, r.sizeBytes, r.attachments.length,
      r.attachments.map((a) => a.name).join("; "), r.fileName || "", r.error || "",
    ].map(toCsvValue).join(","));
  }
  return lines.join("\r\n");
}

async function exportCsv() {
  const records = currentRecords();
  const csv = "﻿" + buildCsv(records);
  const filename = `msg解析レポート_${new Date().toISOString().slice(0, 10)}.csv`;

  const downloads = await claude.use("downloads");
  if (!downloads) {
    el.exportBtn.textContent = "この環境では保存できません";
    setTimeout(() => { el.exportBtn.textContent = "CSVを保存"; }, 2500);
    return;
  }
  try {
    await downloads.save({ filename, data: csv });
  } catch (e) {
    if (e && e.code === "declined") return;
    el.exportBtn.textContent = "保存に失敗しました";
    setTimeout(() => { el.exportBtn.textContent = "CSVを保存"; }, 2500);
  }
}

/* ---------- sort header wiring ---------- */

function wireSortHeaders() {
  el.ledgerTable.querySelectorAll("th[data-sort]").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.getAttribute("data-sort");
      if (state.sortKey === key) {
        state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
      } else {
        state.sortKey = key;
        state.sortDir = key === "date" ? "desc" : "asc";
      }
      el.ledgerTable.querySelectorAll("th[data-sort]").forEach((h) => h.removeAttribute("data-active"));
      th.setAttribute("data-active", state.sortDir);
      render();
    });
  });
}

/* ---------- drag & drop / inputs ---------- */

function wireEvents() {
  el.pickFilesBtn.addEventListener("click", () => el.fileInput.click());
  el.pickFolderBtn.addEventListener("click", () => el.folderInput.click());
  el.fileInput.addEventListener("change", (e) => { handleFiles(e.target.files); e.target.value = ""; });
  el.folderInput.addEventListener("change", (e) => { handleFiles(e.target.files); e.target.value = ""; });
  el.clearBtn.addEventListener("click", clearAll);
  el.exportBtn.addEventListener("click", exportCsv);
  el.searchInput.addEventListener("input", debounce((e) => {
    state.query = e.target.value;
    renderLedger(currentRecords());
  }, 120));

  ["dragenter", "dragover"].forEach((evt) => {
    el.dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      el.dropZone.classList.add("drag-over");
    });
  });
  ["dragleave", "drop"].forEach((evt) => {
    el.dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      el.dropZone.classList.remove("drag-over");
    });
  });
  el.dropZone.addEventListener("drop", (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length) handleFiles(dt.files);
  });
}

/* ---------- init ---------- */

wireSortHeaders();
wireEvents();
render();
