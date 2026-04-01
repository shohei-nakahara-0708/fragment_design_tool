import React, { useEffect, useMemo, useRef, useState } from "react";
import { Link } from "react-router-dom";
import { Document, Page, pdfjs } from "react-pdf";

import "react-pdf/dist/Page/TextLayer.css";
import "react-pdf/dist/Page/AnnotationLayer.css";

pdfjs.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url
).toString();

const API_BASE = import.meta.env.VITE_API_BASE || "";

const IMPORTANT_COLUMN_NAMES = new Set([
  "開催日",
  "時間",
  "講演会名",
  "演題",
  "案内状掲載 医師名",
  "案内状掲載 施設名",
  "案内状掲載 役職",
  "医師名",
  "施設名",
  "役職",
  "title",
  "speaker",
  "facility",
  "role",
]);

const AUXILIARY_COLUMN_NAMES = new Set([
  "案内状掲載 所属科",
  "医師コード",
]);

const DEFAULT_PINNED_HEADERS = ["講演会ID", "演題", "案内状掲載 医師名"];

const ui = {
  page: {
    display: "grid",
    gridTemplateColumns: "minmax(460px, 760px) 1fr",
    gap: 16,
    padding: 16,
    minHeight: "100vh",
    background: "#fafafa",
  },
  leftCol: {
    display: "grid",
    gap: 12,
    alignContent: "start",
    minWidth: 0,
  },
  rightCol: {
    display: "grid",
    gap: 12,
    alignContent: "start",
    position: "sticky",
    top: 16,
    height: "calc(100vh - 32px)",
    minWidth: 0,
  },
  card: {
    background: "#fff",
    border: "1px solid #e7e7e7",
    borderRadius: 14,
    padding: 14,
    boxShadow: "0 1px 2px rgba(0,0,0,0.04)",
  },
  previewOnlyCard: {
    background: "#fff",
    border: "1px solid #e7e7e7",
    borderRadius: 14,
    padding: 10,
    boxShadow: "0 1px 2px rgba(0,0,0,0.04)",
    height: "calc(100vh - 32px)",
    overflow: "auto",
  },
  softBox: {
    border: "1px solid #ececec",
    borderRadius: 12,
    padding: 12,
    background: "#fcfcfc",
  },
  headerRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 12,
    marginBottom: 8,
  },
  sectionTitleRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 8,
    marginBottom: 8,
  },
  h2: {
    margin: "6px 0 0",
    fontSize: 18,
    fontWeight: 800,
    letterSpacing: 0.2,
  },
  h3: {
    margin: 0,
    fontSize: 13,
    fontWeight: 800,
    color: "#222",
  },
  muted: {
    fontSize: 12,
    color: "#666",
  },
  input: {
    border: "1px solid #ddd",
    borderRadius: 10,
    padding: "10px 12px",
    fontSize: 14,
    width: "100%",
    boxSizing: "border-box",
    background: "#fff",
  },
  select: {
    border: "1px solid #ddd",
    borderRadius: 10,
    padding: "8px 10px",
    fontSize: 13,
    background: "#fff",
  },
  textarea: {
    width: "100%",
    minHeight: 120,
    resize: "vertical",
    border: "1px solid #ddd",
    borderRadius: 12,
    padding: 12,
    fontSize: 14,
    lineHeight: 1.6,
    fontFamily: "inherit",
    boxSizing: "border-box",
  },
  badge: (tone = "gray") => {
    const base = {
      display: "inline-flex",
      alignItems: "center",
      gap: 6,
      padding: "5px 10px",
      borderRadius: 999,
      borderWidth: "1px",
      borderStyle: "solid",
      borderColor: "#e3e3e3",
      fontSize: 12,
      lineHeight: 1,
      userSelect: "none",
      whiteSpace: "nowrap",
      background: "#f6f6f6",
      color: "#444",
    };
    if (tone === "green") return { ...base, background: "#ecf8ef", borderColor: "#bfe3c6", color: "#1b6b2f" };
    if (tone === "red") return { ...base, background: "#fff2f2", borderColor: "#f2c2c2", color: "#a00000" };
    if (tone === "yellow") return { ...base, background: "#fff9e8", borderColor: "#f2dda2", color: "#8b6400" };
    if (tone === "blue") return { ...base, background: "#eef5ff", borderColor: "#c7ddff", color: "#1a4fb3" };
    return base;
  },
  btn: (variant = "secondary") => ({
    appearance: "none",
    border: "1px solid " + (variant === "primary" ? "#111" : "#d9d9d9"),
    borderRadius: 12,
    padding: "10px 12px",
    fontWeight: 800,
    cursor: "pointer",
    background: variant === "primary" ? "#111" : "#fff",
    color: variant === "primary" ? "#fff" : "#111",
  }),
  chipBtn: (active = false) => ({
    appearance: "none",
    border: `1px solid ${active ? "#111" : "#d9d9d9"}`,
    borderRadius: 999,
    padding: "6px 10px",
    fontSize: 12,
    fontWeight: 700,
    cursor: "pointer",
    background: active ? "#111" : "#fff",
    color: active ? "#fff" : "#111",
  }),
  toolbarRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: 8,
    alignItems: "center",
  },
  toolbarGrid: {
    display: "grid",
    gap: 10,
  },
  toolbarStats: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
    alignItems: "center",
  },
  columnsBox: {
    display: "grid",
    gap: 8,
    maxHeight: 240,
    overflow: "auto",
    padding: 10,
    border: "1px solid #ececec",
    borderRadius: 10,
    background: "#fff",
  },
  columnsGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
    gap: 8,
  },
  checkLabel: {
    display: "inline-flex",
    gap: 6,
    alignItems: "center",
    fontSize: 13,
    color: "#333",
  },
  sheetWrap: {
    background: "#fff",
    border: "1px solid #d9d9d9",
    borderRadius: 12,
    overflow: "hidden",
    boxShadow: "0 1px 2px rgba(0,0,0,0.04)",
  },
  sheetScroller: {
    overflow: "auto",
    maxHeight: "calc(100vh - 250px)",
    background: "#fff",
  },
  sheetTable: {
    borderCollapse: "separate",
    borderSpacing: 0,
    width: "max-content",
    minWidth: "100%",
    tableLayout: "fixed",
    fontSize: 13,
  },
  sheetCorner: {
    position: "sticky",
    top: 0,
    left: 0,
    zIndex: 6,
    background: "#f1f3f4",
    borderRight: "1px solid #d0d7de",
    borderBottom: "1px solid #d0d7de",
    width: 48,
    minWidth: 48,
    height: 34,
  },
  sheetColHeader: {
    position: "sticky",
    top: 0,
    zIndex: 5,
    background: "#f1f3f4",
    borderBottom: "1px solid #d0d7de",
    borderRight: "1px solid #e5e7eb",
    textAlign: "center",
    fontWeight: 700,
    color: "#444",
    padding: "8px 6px",
    height: 34,
    userSelect: "none",
    whiteSpace: "nowrap",
  },
  sheetRowHeader: {
    position: "sticky",
    left: 0,
    zIndex: 4,
    background: "#f1f3f4",
    borderRight: "1px solid #d0d7de",
    borderBottom: "1px solid #e5e7eb",
    textAlign: "center",
    color: "#555",
    width: 48,
    minWidth: 48,
    fontSize: 12,
    padding: "6px 4px",
    userSelect: "none",
  },
  sheetHeaderCell: {
    background: "#f8fafc",
    borderRight: "1px solid #e5e7eb",
    borderBottom: "1px solid #dfe3e8",
    padding: "8px 10px",
    fontWeight: 700,
    color: "#222",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  sheetCell: {
    position: "relative",
    borderRight: "1px solid #e5e7eb",
    borderBottom: "1px solid #e5e7eb",
    padding: "8px 10px",
    verticalAlign: "top",
    background: "#fff",
    cursor: "pointer",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    lineHeight: 1.45,
    minHeight: 44,
  },
  headerMenu: {
    position: "fixed",
    zIndex: 1000,
    width: 300,
    background: "#fff",
    border: "1px solid #ddd",
    borderRadius: 12,
    boxShadow: "0 8px 24px rgba(0,0,0,0.12)",
    padding: 12,
    display: "grid",
    gap: 10,
  },
  headerMenuValues: {
    maxHeight: 180,
    overflow: "auto",
    border: "1px solid #eee",
    borderRadius: 8,
    padding: 8,
  },
  previewToolbar: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
    alignItems: "center",
    marginBottom: 10,
  },
  previewHint: {
    fontSize: 12,
    color: "#666",
    marginTop: 6,
  },
  miniPill: {
    display: "inline-flex",
    alignItems: "center",
    padding: "2px 6px",
    borderRadius: 999,
    fontSize: 11,
    fontWeight: 700,
    background: "#fff",
    border: "1px solid #ddd",
    color: "#333",
  },
  hitLegendRow: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
    alignItems: "center",
  },
  viewerBox: {
    border: "1px solid #eee",
    borderRadius: 14,
    overflow: "auto",
    maxHeight: "calc(100vh - 130px)",
    background: "#f6f6f6",
    padding: 8,
    minWidth: 0,
  },
  textPane: {
    background: "#fff",
    border: "1px solid #eee",
    borderRadius: 12,
    padding: 12,
    fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
    fontSize: 13,
    lineHeight: 1.6,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  collapseCard: {
    border: "1px solid #ececec",
    borderRadius: 12,
    background: "#fcfcfc",
    overflow: "hidden",
  },
  collapseHeader: {
    width: "100%",
    textAlign: "left",
    background: "#fff",
    border: "none",
    borderBottom: "1px solid #ececec",
    padding: "12px 14px",
    fontWeight: 800,
    cursor: "pointer",
  },
  collapseBody: {
    padding: 12,
  },
  diffGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 12,
  },
  diffPane: {
    background: "#fff",
    border: "1px solid #e5e7eb",
    borderRadius: 12,
    overflow: "hidden",
  },
  diffPaneHeader: {
    padding: "8px 10px",
    borderBottom: "1px solid #e5e7eb",
    background: "#f8fafc",
    fontSize: 12,
    fontWeight: 800,
    color: "#444",
  },
  diffPaneBody: {
    padding: 12,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    lineHeight: 1.7,
    fontSize: 14,
    minHeight: 120,
  },
  diffSame: {},
  diffAdd: {
    background: "#B3DEB6",
    color: "#1b5e20",
    borderRadius: 3,
  },
  diffRemove: {
    background: "#faafba",
    color: "#b71c1c",
    borderRadius: 3,
  },
  auxHeaderCell: {
  background: "#f3f4f6",
  color: "#6b7280",
},
auxCell: {
  background: "#f8f9fb",
  color: "#6b7280",
  },
  sheetFooter: {
    padding: "5px 10px",
  }
};

function normalizeText(s) {
  return String(s || "")
    .replace(/\r/g, "")
    .replace(/\n+/g, "\n")
    .replace(/[ \t　]+/g, " ")
    .trim();
}

function normalizeKey(s) {
  return normalizeText(s)
    .replace(/[\s　]+/g, "")
    .replace(/先生/g, "")
    .replace(/[（(].*?[）)]/g, "")
    .toLowerCase();
}

function normalizeForCompare(s) {
  return String(s || "")
    .replace(/\r/g, "")
    .replace(/[ \t\u3000]+/g, " ")
    .replace(/[：]/g, ":")
    .replace(/[〜～]/g, "～")
    .trim();
}

function statusTone(status) {
  if (status === "match") return "green";
  if (status === "partial") return "yellow";
  if (status === "mismatch") return "red";
  return "gray";
}

function statusText(status) {
  if (status === "match") return "一致";
  if (status === "partial") return "部分一致";
  if (status === "mismatch") return "未検出";
  return "未入力";
}

function flattenBlocks(blocks) {
  const arr = Array.isArray(blocks) ? blocks : [];
  return arr
    .slice()
    .sort((a, b) => {
      const ap = Number(a.page || a.page_number || 1);
      const bp = Number(b.page || b.page_number || 1);
      if (ap !== bp) return ap - bp;
      return (Number(a.top || 0) - Number(b.top || 0)) || (Number(a.left || 0) - Number(b.left || 0));
    })
    .map((b, idx) => ({
      index: idx + 1,
      text: normalizeText(b.text || ""),
      left: Number(b.left || 0),
      top: Number(b.top || 0),
      width: Number(b.width || 0),
      height: Number(b.height || 0),
      max_font_pt: Number(b.max_font_pt || 0),
      page: Number(b.page || b.page_number || 1),
      page_width: Number(b.page_width || b._page_width || 0),
      page_height: Number(b.page_height || b._page_height || 0),
      coord_unit: b.coord_unit || b._coord_unit || "",
    }));
}

function colLetter(n) {
  let s = "";
  let x = n + 1;
  while (x > 0) {
    const mod = (x - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

function getErrorMessageFromResponseData(data, fallback) {
  if (!data) return fallback;
  if (typeof data === "string") return data;
  if (typeof data.detail === "string") return data.detail;
  if (typeof data.message === "string") return data.message;
  return fallback;
}

async function readErrorMessage(res, fallback) {
  try {
    const data = await res.json();
    return getErrorMessageFromResponseData(data, fallback);
  } catch {
    return fallback;
  }
}

function splitJapaneseNameParts(name) {
  const compact = normalizeText(name).replace(/\s+/g, "");
  if (!compact) return [];

  if (compact.length <= 2) return [compact];

  // 雑でも、姓2〜3文字 + 残り の候補を作る
  const parts = [];
  if (compact.length >= 2) parts.push(compact.slice(0, 2));
  if (compact.length >= 3) parts.push(compact.slice(0, 3));
  parts.push(compact.slice(2));
  if (compact.length >= 3) parts.push(compact.slice(3));

  return Array.from(new Set(parts.filter(Boolean)));
}

function scoreBlockMatch(value, blockText, fieldLabel) {
  const raw = normalizeText(value);
  const key = normalizeKey(value);

  const blockRaw = normalizeText(blockText);
  const blockKey = normalizeKey(blockText);

  if (!raw || !key || !blockKey) return 0;

  let score = 0;

  const rawCompact = raw.replace(/\s+/g, "");
  const blockCompact = blockRaw.replace(/\s+/g, "");

  const label = normalizeText(fieldLabel);

  // 完全一致系
  if (blockKey === key) score += 120;
  if (blockRaw === raw) score += 40;
  if (blockCompact === rawCompact) score += 60;

  // 包含系
  if (blockRaw.includes(raw)) score += 60;
  if (raw.includes(blockRaw) && blockRaw.length >= 2) score += 25;

  if (blockKey.includes(key)) score += 55;
  if (key.includes(blockKey) && blockKey.length >= 2) score += 30;

  if (blockCompact.includes(rawCompact)) score += 50;
  if (rawCompact.includes(blockCompact) && blockCompact.length >= 2) score += 25;

  // 人名は姓だけ/名だけでも多少拾う
if (label.includes("医師名") || label === "speaker") {
  const parts = splitJapaneseNameParts(raw);
  for (const p of parts) {
    if (p.length >= 2 && blockCompact.includes(p)) {
      score += 12;
    }
  }
}

  // 施設名は大学・病院などの語があると加点
  if (label.includes("施設名") || label === "facility" || label.includes("所属")) {
    if (/病院|大学|クリニック|センター|科|部|医院|学部/.test(blockRaw)) score += 15;
  }

  // 役職は役職語を含むと加点
  if (label.includes("役職") || label === "role") {
    if (/教授|部長|医長|院長|講師|助教|准教授|センター長|科長/.test(blockRaw)) score += 18;
  }

  // 演題は短すぎる候補を少し減点
  if (label.includes("演題") || label === "title") {
    if (blockRaw.length >= 8) score += 5;
    if (blockRaw.length <= 3) score -= 20;
  }

  return score;
}

function compareValueToBlocks(value, blocks, fieldLabel) {
  const raw = normalizeText(value);
  if (!raw) {
    return { status: "missing", hits: [] };
  }

  const rawKey = normalizeForCompare(raw);

  const hits = blocks
    .map((b) => {
      const score = scoreBlockMatch(value, b.text || "", fieldLabel);

      const blockText = String(b.text || "");
      const blockNorm = normalizeForCompare(blockText);

      let matchedText = blockText;

      let rawMatchStart = -1;
      let rawMatchLength = 0;

      let normMatchStart = -1;
      let normMatchLength = 0;

      // まずは生文字列で探す
      const rawIdx = blockText.indexOf(raw);
      if (rawIdx !== -1) {
        rawMatchStart = Array.from(blockText.slice(0, rawIdx)).length;
        rawMatchLength = Array.from(raw).length;
        matchedText = raw;
      } else {
        // ダメなら正規化後文字列で探す
        const normIdx = blockNorm.indexOf(rawKey);
        if (normIdx !== -1) {
          normMatchStart = normIdx;
          normMatchLength = Array.from(rawKey).length;
          matchedText = raw;
        }
      }

      return {
        ...b,
        score,
        matchType: score >= 100 ? "exact" : score >= 40 ? "partial" : "weak",
        keyword: raw,
        matchedText,
        rawTextLength: Array.from(blockText).length,
        normalizedTextLength: Array.from(blockNorm).length,
        rawMatchStart,
        rawMatchLength,
        normMatchStart,
        normMatchLength,
      };
    })
    .filter((b) => b.score > 0)
    .sort((a, b) => b.score - a.score || a.top - b.top || a.left - b.left);

  if (hits[0]?.score >= 110) {
  return { status: "match", hits };
}
if (hits[0]?.score >= 28) {
  return { status: "partial", hits };
}
return { status: "mismatch", hits };
}

function getVmHeaders(rows) {
  const ordered = [];
  const seen = new Set();

  for (const row of Array.isArray(rows) ? rows : []) {
    for (const key of Object.keys(row || {})) {
      if (!seen.has(key)) {
        seen.add(key);
        ordered.push(key);
      }
    }
  }

  const priority = [
    "講演会ID",
    "システムID",
    "演題",
    "案内状掲載 医師名",
    "案内状掲載 施設名",
    "案内状掲載 役職",
  ];

  const prioritized = [];
  for (const p of priority) {
    if (seen.has(p)) prioritized.push(p);
  }

  const rest = ordered.filter((x) => !priority.includes(x));
  return [...prioritized, ...rest];
}

function shapeVmRows(vmRows, blocks, headersFromApi = []) {
  const rows = Array.isArray(vmRows) ? vmRows : [];
  const flatBlocks = flattenBlocks(blocks);
  const headers =
    Array.isArray(headersFromApi) && headersFromApi.length
      ? headersFromApi
      : getVmHeaders(rows);

  const sections = rows.map((vm, idx) => {
    const fields = headers.map((header) => {
      const value = vm?.[header] ?? "";
      return {
        key: `row-${idx + 1}-${header}`,
        label: header,
        value: value == null ? "" : String(value),
        rowIndex: idx,
        ...compareValueToBlocks(value, flatBlocks, header),
      };
    });

    return {
      id: `vm-row-${idx + 1}`,
      rowNumber: idx + 2,
      label: `VM行 ${idx + 1}`,
      raw: vm,
      fields,
    };
  });

  return {
    headers,
    rows: sections,
  };
}

function getColumnWidth(label) {
  const s = String(label || "");
  if (s === "案内状掲載 所属科") return 120;
  if (s === "医師コード") return 100;
  if (s.includes("演題")) return 320;
  if (s.includes("施設名")) return 260;
  if (s.includes("講演会名")) return 260;
  if (s.includes("講演会ID") || s.includes("システムID")) return 200;
  if (s.includes("日時") || s.includes("開催日") || s.includes("時間")) return 200;
  if (s.includes("医師名")) return 180;
  if (s.includes("役職")) return 180;
  return 180;
}
function useElementWidth() {
  const ref = useRef(null);
  const [width, setWidth] = useState(1);

  useEffect(() => {
    if (!ref.current) return;

    const el = ref.current;
    const update = () => setWidth(el.clientWidth || 1);

    update();

    const ro = new ResizeObserver(update);
    ro.observe(el);

    return () => ro.disconnect();
  }, []);

  return [ref, width];
}

function summarizeRow(row) {
  let match = 0;
  let partial = 0;
  let mismatch = 0;
  let missing = 0;

  for (const f of row.fields || []) {
    if (f.status === "match") match += 1;
    else if (f.status === "partial") partial += 1;
    else if (f.status === "mismatch") mismatch += 1;
    else missing += 1;
  }

  return { match, partial, mismatch, missing };
}

function diffChars(a, b) {
  const left = Array.from(String(a || ""));
  const right = Array.from(String(b || ""));

  const m = left.length;
  const n = right.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (left[i - 1] === right[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  const outLeft = [];
  const outRight = [];
  let i = m;
  let j = n;

  while (i > 0 && j > 0) {
    if (left[i - 1] === right[j - 1]) {
      outLeft.push({ ch: left[i - 1], type: "same" });
      outRight.push({ ch: right[j - 1], type: "same" });
      i--;
      j--;
    } else if (dp[i - 1][j] >= dp[i][j - 1]) {
      outLeft.push({ ch: left[i - 1], type: "remove" });
      i--;
    } else {
      outRight.push({ ch: right[j - 1], type: "add" });
      j--;
    }
  }

  while (i > 0) {
    outLeft.push({ ch: left[i - 1], type: "remove" });
    i--;
  }

  while (j > 0) {
    outRight.push({ ch: right[j - 1], type: "add" });
    j--;
  }

  return {
    left: outLeft.reverse(),
    right: outRight.reverse(),
  };
}

function buildDiffSummary(a, b) {
  const rawA = String(a || "");
  const rawB = String(b || "");
  const normA = normalizeForCompare(rawA);
  const normB = normalizeForCompare(rawB);

  if (!rawA && !rawB) return { tone: "gray", text: "両方空です" };
  if (rawA === rawB) return { tone: "green", text: "原文一致" };
  if (normA === normB) return { tone: "yellow", text: "正規化後一致" };
  return { tone: "red", text: "差分あり" };
}

function CollapseSection({ title, defaultOpen = false, children, right }) {
  const [open, setOpen] = useState(defaultOpen);

  return (
    <div style={ui.collapseCard}>
      <button type="button" style={ui.collapseHeader} onClick={() => setOpen((v) => !v)}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
          <span>{open ? "▼" : "▶"} {title}</span>
          {right ? <span>{right}</span> : null}
        </div>
      </button>
      {open ? <div style={ui.collapseBody}>{children}</div> : null}
    </div>
  );
}

function DiffText({ parts, side }) {
  return (
    <div>
      {parts.map((p, idx) => {
        let style = ui.diffSame;
        if (p.type === "add" && side === "right") style = ui.diffAdd;
        if (p.type === "remove" && side === "left") style = ui.diffRemove;

        if (p.type === "add" && side === "left") return null;
        if (p.type === "remove" && side === "right") return null;

        const ch = p.ch;

        if (ch === "\n") {
          const newlineStyle =
            p.type === "add" && side === "right"
              ? ui.diffAdd
              : p.type === "remove" && side === "left"
              ? ui.diffRemove
              : {};

          return (
            <React.Fragment key={idx}>
              <span
                style={{
                  ...newlineStyle,
                  display: "inline-block",
                  padding: "0 4px",
                  margin: "0 2px",
                  fontSize: 12,
                  lineHeight: 1.4,
                  opacity: 0.9,
                }}
              >
                ↵
              </span>
              <br />
            </React.Fragment>
          );
        }

        return (
          <span key={idx} style={style}>
            {ch}
          </span>
        );
      })}
    </div>
  );
}

function SideBySideComparePanel({ selectedField, manualCompareText, setManualCompareText }) {
  const sheetValue = selectedField?.value || "";
const compareTargetText =
  manualCompareText.trim() || selectedField?.hits?.[0]?.matchedText || selectedField?.hits?.[0]?.text || "";

  const diff = useMemo(
    () => diffChars(sheetValue, compareTargetText),
    [sheetValue, compareTargetText]
  );
  const summary = useMemo(
    () => buildDiffSummary(sheetValue, compareTargetText),
    [sheetValue, compareTargetText]
  );

  if (!selectedField) {
    return <div style={ui.textPane}>セルを選択すると比較を表示します。</div>;
  }

  return (
    <div style={{ display: "grid", gap: 12 }}>
      <div style={ui.sectionTitleRow}>
        <div style={ui.h3}>文字比較</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <span style={ui.badge(summary.tone)}>{summary.text}</span>
          <span style={ui.badge("blue")}>{selectedField.label}</span>
        </div>
      </div>

      <div style={{ display: "grid", gap: 8 }}>
        <div style={ui.muted}>PDFからコピーした文字を貼り付け</div>
        <textarea
          value={manualCompareText}
          onChange={(e) => setManualCompareText(e.target.value)}
          placeholder="右のPDFで選択してコピーした文字をここに貼り付け"
          style={ui.textarea}
        />
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button
            type="button"
            style={ui.btn("secondary")}
            onClick={() => setManualCompareText("")}
          >
            手動入力をクリア
          </button>
        </div>
      </div>

      <div style={ui.diffGrid}>
        <div style={ui.diffPane}>
          <div style={ui.diffPaneHeader}>シート値</div>
          <div style={ui.diffPaneBody}>
            <DiffText parts={diff.left} side="left" />
          </div>
        </div>

        <div style={ui.diffPane}>
          <div style={ui.diffPaneHeader}>
            {manualCompareText.trim() ? "貼り付け文字列" : "PDF候補"}
          </div>
          <div style={ui.diffPaneBody}>
            <DiffText parts={diff.right} side="right" />
          </div>
        </div>
      </div>
    </div>
  );
}

function HeaderFilterMenu({
  headerMenu,
  setHeaderMenu,
  rows,
  columnFilters,
  setColumnFilters,
  columnSort,
  setColumnSort,
}) {
  const menuRef = useRef(null);
  const header = headerMenu?.header || "";
  const filter = columnFilters[header] || {
    query: "",
    statuses: [],
    onlyBlank: false,
    onlyNonBlank: false,
    selectedValues: [],
  };

  const values = useMemo(() => {
    if (!header) return [];
    const uniq = new Set();
    for (const row of rows) {
      const field = row.fields.find((f) => f.label === header);
      uniq.add(String(field?.value || ""));
    }
    return Array.from(uniq).slice(0, 80);
  }, [rows, header]);

  useEffect(() => {
    if (!headerMenu) return;

    function handleMouseDown(e) {
      if (!menuRef.current) return;
      if (!menuRef.current.contains(e.target)) {
        setHeaderMenu(null);
      }
    }

    window.addEventListener("mousedown", handleMouseDown);
    return () => window.removeEventListener("mousedown", handleMouseDown);
  }, [headerMenu, setHeaderMenu]);

  if (!headerMenu) return null;

  return (
    <div
      ref={menuRef}
      style={{
        ...ui.headerMenu,
        left: headerMenu.anchorX,
        top: headerMenu.anchorY,
      }}
    >
      <div style={{ fontWeight: 800, wordBreak: "break-word" }}>{header}</div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <button
          type="button"
          style={ui.btn("secondary")}
          onClick={() => {
            setColumnSort({ header, direction: "asc" });
            setHeaderMenu(null);
          }}
        >
          昇順
        </button>
        <button
          type="button"
          style={ui.btn("secondary")}
          onClick={() => {
            setColumnSort({ header, direction: "desc" });
            setHeaderMenu(null);
          }}
        >
          降順
        </button>
        <button
          type="button"
          style={ui.btn("secondary")}
          onClick={() => {
            setColumnSort({ header: "", direction: "" });
            setHeaderMenu(null);
          }}
        >
          ソート解除
        </button>
      </div>

      <input
        value={filter.query}
        onChange={(e) =>
          setColumnFilters((prev) => ({
            ...prev,
            [header]: { ...filter, query: e.target.value },
          }))
        }
        placeholder="この列で検索"
        style={ui.input}
      />

      <div style={{ display: "grid", gap: 6 }}>
        <label style={ui.checkLabel}>
          <input
            type="checkbox"
            checked={filter.onlyBlank}
            onChange={(e) =>
              setColumnFilters((prev) => ({
                ...prev,
                [header]: {
                  ...filter,
                  onlyBlank: e.target.checked,
                  onlyNonBlank: e.target.checked ? false : filter.onlyNonBlank,
                },
              }))
            }
          />
          空白のみ
        </label>

        <label style={ui.checkLabel}>
          <input
            type="checkbox"
            checked={filter.onlyNonBlank}
            onChange={(e) =>
              setColumnFilters((prev) => ({
                ...prev,
                [header]: {
                  ...filter,
                  onlyNonBlank: e.target.checked,
                  onlyBlank: e.target.checked ? false : filter.onlyBlank,
                },
              }))
            }
          />
          空白以外
        </label>
      </div>

      <div style={ui.headerMenuValues}>
        {values.map((v) => {
          const checked = filter.selectedValues?.includes(v) || false;
          return (
            <label
              key={`${header}-${v}`}
              style={{
                display: "flex",
                gap: 6,
                alignItems: "center",
                fontSize: 12,
                marginBottom: 6,
              }}
            >
              <input
                type="checkbox"
                checked={checked}
                onChange={(e) => {
                  const nextValues = e.target.checked
                    ? [...(filter.selectedValues || []), v]
                    : (filter.selectedValues || []).filter((x) => x !== v);

                  setColumnFilters((prev) => ({
                    ...prev,
                    [header]: { ...filter, selectedValues: nextValues },
                  }));
                }}
              />
              <span style={{ wordBreak: "break-word" }}>{v || "(空白)"}</span>
            </label>
          );
        })}
      </div>

      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <button
          type="button"
          style={ui.btn("secondary")}
          onClick={() => {
            setColumnFilters((prev) => {
              const next = { ...prev };
              delete next[header];
              return next;
            });
            setHeaderMenu(null);
          }}
        >
          フィルタ解除
        </button>
        <button
          type="button"
          style={ui.btn("primary")}
          onClick={() => setHeaderMenu(null)}
        >
          閉じる
        </button>
      </div>
    </div>
  );
}

function SheetToolbar({
  rows,
  headers,
  filteredRowCount,
  sheetQuery,
  setSheetQuery,
  statusFilter,
  setStatusFilter,
  showOnlyRowsWithDiff,
  setShowOnlyRowsWithDiff,
  sortMode,
  setSortMode,
  visibleHeaders,
  setVisibleHeaders,
  pinnedHeaders,
  setPinnedHeaders,
  showImportantOnly,
  setShowImportantOnly,
  onlyHighlightImportant,
  setOnlyHighlightImportant,
}) {
  const [showColumnPicker, setShowColumnPicker] = useState(false);

  const summary = useMemo(() => {
    let match = 0;
    let partial = 0;
    let mismatch = 0;
    let missing = 0;

    for (const row of rows || []) {
      for (const f of row.fields || []) {
        if (f.status === "match") match += 1;
        else if (f.status === "partial") partial += 1;
        else if (f.status === "mismatch") mismatch += 1;
        else missing += 1;
      }
    }

    return { match, partial, mismatch, missing };
  }, [rows]);

  const allVisible = headers.length > 0 && visibleHeaders.length === headers.length;

  function toggleVisibleHeader(header, checked) {
    if (checked) {
      setVisibleHeaders((prev) => (prev.includes(header) ? prev : [...prev, header]));
    } else {
      setVisibleHeaders((prev) => prev.filter((x) => x !== header));
      setPinnedHeaders((prev) => prev.filter((x) => x !== header));
    }
  }

  function togglePinnedHeader(header, checked) {
    if (checked) {
      setPinnedHeaders((prev) => (prev.includes(header) ? prev : [...prev, header]));
    } else {
      setPinnedHeaders((prev) => prev.filter((x) => x !== header));
    }
  }

  return (
    <div style={ui.card}>
      <div style={ui.sectionTitleRow}>
        <div style={ui.h3}>シート操作</div>
        <span style={ui.badge("blue")}>
          {filteredRowCount} / {rows.length} 行
        </span>
      </div>

      <div style={ui.toolbarGrid}>
        <input
          value={sheetQuery}
          onChange={(e) => setSheetQuery(e.target.value)}
          placeholder="列名・セル値で検索"
          style={ui.input}
        />

        <div style={ui.toolbarRow}>
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            style={ui.select}
          >
            <option value="all">全ステータス</option>
            <option value="match">一致あり</option>
            <option value="partial">部分一致あり</option>
            <option value="mismatch">未検出あり</option>
            <option value="missing">未入力あり</option>
          </select>

          <select
            value={sortMode}
            onChange={(e) => setSortMode(e.target.value)}
            style={ui.select}
          >
            <option value="sheet">シート順</option>
            <option value="mismatch_desc">未検出が多い順</option>
            <option value="partial_desc">部分一致が多い順</option>
          </select>

          <button
            type="button"
            style={ui.btn(showOnlyRowsWithDiff ? "primary" : "secondary")}
            onClick={() => setShowOnlyRowsWithDiff((v) => !v)}
          >
            {showOnlyRowsWithDiff ? "差分あり行のみ" : "全行表示"}
          </button>

          <button
            type="button"
            style={ui.btn(showColumnPicker ? "primary" : "secondary")}
            onClick={() => setShowColumnPicker((v) => !v)}
          >
            列設定
          </button>
        </div>

        <div style={ui.toolbarRow}>
          <button
            type="button"
            style={ui.btn(showImportantOnly ? "primary" : "secondary")}
            onClick={() => setShowImportantOnly((v) => !v)}
          >
            {showImportantOnly ? "重要列のみ表示中" : "全列表示中"}
          </button>

          <button
            type="button"
            style={ui.btn(onlyHighlightImportant ? "primary" : "secondary")}
            onClick={() => setOnlyHighlightImportant((v) => !v)}
          >
            {onlyHighlightImportant ? "重要列だけ色付け" : "全列を色付け"}
          </button>
        </div>

        <div style={ui.toolbarStats}>
          <span style={ui.badge("green")}>一致 {summary.match}</span>
          <span style={ui.badge("yellow")}>部分一致 {summary.partial}</span>
          <span style={ui.badge("red")}>未検出 {summary.mismatch}</span>
          <span style={ui.badge()}>未入力 {summary.missing}</span>
          <span style={ui.badge("blue")}>表示列 {visibleHeaders.length}</span>
          <span style={ui.badge("blue")}>固定列 {pinnedHeaders.length}</span>
        </div>

        {showColumnPicker && (
          <div style={ui.softBox}>
            <div style={{ ...ui.sectionTitleRow, marginBottom: 10 }}>
              <div style={ui.h3}>列設定</div>
              <div style={ui.toolbarRow}>
                <button type="button" style={ui.chipBtn(allVisible)} onClick={() => setVisibleHeaders(headers)}>
                  全列表示
                </button>
                <button
                  type="button"
                  style={ui.chipBtn(false)}
                  onClick={() => {
                    setVisibleHeaders([]);
                    setPinnedHeaders([]);
                  }}
                >
                  全列非表示
                </button>
                <button
                  type="button"
                  style={ui.chipBtn(false)}
                  onClick={() => setPinnedHeaders(DEFAULT_PINNED_HEADERS.filter((h) => headers.includes(h)))}
                >
                  固定列リセット
                </button>
              </div>
            </div>

            <div style={ui.columnsBox}>
              <div style={ui.columnsGrid}>
                {headers.map((header) => {
                  const visible = visibleHeaders.includes(header);
                  const pinned = pinnedHeaders.includes(header);

                  return (
                    <div
                      key={header}
                      style={{
                        border: "1px solid #ececec",
                        borderRadius: 10,
                        padding: 10,
                        background: "#fff",
                        display: "grid",
                        gap: 8,
                      }}
                    >
                      <div style={{ fontSize: 13, fontWeight: 700, wordBreak: "break-word" }}>
                        {header}
                      </div>

                      <label style={ui.checkLabel}>
                        <input
                          type="checkbox"
                          checked={visible}
                          onChange={(e) => toggleVisibleHeader(header, e.target.checked)}
                        />
                        表示
                      </label>

                      <label style={ui.checkLabel}>
                        <input
                          type="checkbox"
                          checked={pinned}
                          disabled={!visible}
                          onChange={(e) => togglePinnedHeader(header, e.target.checked)}
                        />
                        左に固定
                      </label>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function SpreadsheetLikeTable({
  rows,
  headers,
  selectedKey,
  onSelect,
  showImportantOnly,
  onlyHighlightImportant,
  pinnedHeaders = [],
  onOpenHeaderMenu,
  columnFilters,
  columnSort,
}) {
  const filteredHeaders = useMemo(() => {
    const base = showImportantOnly
      ? headers.filter((h) => IMPORTANT_COLUMN_NAMES.has(h))
      : headers;

    const pinned = base.filter((h) => pinnedHeaders.includes(h));
    const rest = base.filter((h) => !pinnedHeaders.includes(h));
    return [...pinned, ...rest];
  }, [headers, showImportantOnly, pinnedHeaders]);

  function getField(row, header) {
    return row.fields.find((f) => f.label === header);
  }

  function cellBackground(field, isSelected) {
    if (isSelected) return "#e8f0fe";

    const isImportant = IMPORTANT_COLUMN_NAMES.has(field?.label);
    if (onlyHighlightImportant && !isImportant) return "#fff";

    if (field?.status === "match") return "#f0fbf3";
    if (field?.status === "partial") return "#fff9e8";
    if (field?.status === "mismatch") return "#fff3f3";
    return "#fff";
  }

  return (
    <div style={ui.sheetWrap}>
      <div style={ui.sheetScroller}>
        <table style={ui.sheetTable}>
          <colgroup>
            <col style={{ width: 48 }} />
            {filteredHeaders.map((header) => (
              <col key={header} style={{ width: getColumnWidth(header) }} />
            ))}
          </colgroup>

          <thead>
            <tr>
              <th style={ui.sheetCorner} />
              {filteredHeaders.map((header, idx) => (
                <th key={`col-letter-${header}`} style={ui.sheetColHeader}>
                  {colLetter(idx)}
                </th>
              ))}
            </tr>
            <tr>
              <th style={ui.sheetRowHeader}>1</th>
              {filteredHeaders.map((header) => {
                    const hasFilter = !!columnFilters?.[header];
                    const isSorted = columnSort?.header === header;
                    const isAux = AUXILIARY_COLUMN_NAMES.has(header);

                    return (
                      <th
                        key={`col-header-${header}`}
                        style={{
                          ...ui.sheetHeaderCell,
                          ...(isAux ? ui.auxHeaderCell : null),
                          background: hasFilter
                            ? "#e8f0fe"
                            : isAux
                            ? "#f3f4f6"
                            : IMPORTANT_COLUMN_NAMES.has(header)
                            ? "#eef5ff"
                            : "#f8fafc",
                        }}
                        title={header}
                      >
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        gap: 8,
                      }}
                    >
                      <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>{header}</span>
                      <button
                        type="button"
                        onClick={(e) => {
                          const rect = e.currentTarget.getBoundingClientRect();
                          onOpenHeaderMenu?.({
                            header,
                            anchorX: Math.max(8, rect.left - 220),
                            anchorY: rect.bottom + 4,
                          });
                        }}
                        style={{
                          border: "none",
                          background: "transparent",
                          cursor: "pointer",
                          fontSize: 12,
                          color: isSorted ? "#1a73e8" : "#666",
                          padding: 0,
                        }}
                      >
                        {isSorted
                          ? columnSort.direction === "asc"
                            ? "▲"
                            : "▼"
                          : "▾"}
                      </button>
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>

          <tbody>
            {rows.map((row, rowIdx) => (
              <tr key={row.id}>
                <th style={ui.sheetRowHeader}>{rowIdx + 2}</th>

                {filteredHeaders.map((header, colIdx) => {
                  const field = getField(row, header);
                  const isSelected = selectedKey === field?.key;
                  const cellRef = `${colLetter(colIdx)}${rowIdx + 2}`;
                  const isAux = AUXILIARY_COLUMN_NAMES.has(header);

                  return (
                    <td
                      key={`${row.id}-${header}`}
                      onClick={() => {
                        if (!field) return;

                        if (selectedKey === field.key) {
                          onSelect(null, null); // ← 解除
                        } else {
                          onSelect(field, cellRef); // ← 通常選択
                        }
                      }}
                      style={{
                        ...ui.sheetCell,
                        ...(isAux ? ui.auxCell : null),
                        background: isSelected
                          ? "#e8f0fe"
                          : isAux
                          ? "#f8f9fb"
                          : cellBackground(field, isSelected),
                        boxShadow: isSelected ? "inset 0 0 0 2px #1a73e8" : "none",
                        fontSize: isAux ? 12 : 13,
                      }}
                      title={field?.value || ""}
                    >
                      <div style={{ paddingRight: 48 }}>{field?.value || ""}</div>

                      {(field?.status === "match" ||
                        field?.status === "partial" ||
                        field?.status === "mismatch") && (
                        <div
                          style={{
                            position: "absolute",
                            top: 6,
                            right: 6,
                            width: 10,
                            height: 10,
                            borderRadius: 999,
                            background:
                              field.status === "match"
                                ? "#34a853"
                                : field.status === "partial"
                                ? "#fbbc04"
                                : "#ea4335",
                          }}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={ui.sheetFooter}>
        <span style={ui.muted}>※案内状から情報確認できる列はヘッダーが青色で表示されます。</span>
      </div>
    </div>
  );
}

function PdfPreview({ file, selectedField, activePreviewHitKey }) {
 const keyword = selectedField?.value || "";

const hits = useMemo(
  () => [...(selectedField?.hits || [])].sort((a, b) => (b.score || 0) - (a.score || 0)),
  [selectedField]
);

  const [numPages, setNumPages] = useState(0);
  const [containerRef, viewWidth] = useElementWidth();
  const [pdfError, setPdfError] = useState("");
  const [pageViewports, setPageViewports] = useState({});

  const renderedWidth = Math.max(1, Math.floor(viewWidth));
  const maxScore = Math.max(...hits.map((x) => Number(x.score || 0)), 0);

  function onLoadSuccess(pdf) {
    setNumPages(pdf.numPages || 0);
    setPdfError("");
    setPageViewports({});

    
  }

  function onLoadError(error) {
    console.error("PDF load error:", error);
    setPdfError(String(error?.message || error || "PDFの読み込みに失敗しました。"));
  }

  function handlePageLoadSuccess(page, pageNumber) {
    const viewport = page.getViewport({ scale: 1 });
    setPageViewports((prev) => ({
      ...prev,
      [pageNumber]: {
        width: viewport.width,
        height: viewport.height,
      },
    }));
  }

function getScaledRect(hit, pageNumber) {
  const vp = pageViewports[pageNumber];
  if (!vp?.width) return null;

  const xScale = renderedWidth / vp.width;
  const yScale = xScale;

  const rawLeft = Number(hit.left || 0);
  const rawTop = Number(hit.top || 0);
  const rawWidth = Number(hit.width || 0);
  const rawHeight = Number(hit.height || 0);

  let left = rawLeft * xScale;
  let top = rawTop * yScale;
  let width = Math.max(8, rawWidth * xScale);
  let height = Math.max(8, rawHeight * yScale);

  // 生文字列で位置が取れていればそれを最優先
  if (Number(hit.rawMatchStart) >= 0 && Number(hit.rawMatchLength) > 0 && Number(hit.rawTextLength) > 0) {
    const startRatio = Number(hit.rawMatchStart) / Number(hit.rawTextLength);
    const widthRatio = Number(hit.rawMatchLength) / Number(hit.rawTextLength);

    left += width * startRatio;
    width = Math.max(8, width * widthRatio);
  }
  // 生文字列で無理だった時だけ normalized 比率で寄せる
  else if (
    Number(hit.normMatchStart) >= 0 &&
    Number(hit.normMatchLength) > 0 &&
    Number(hit.normalizedTextLength) > 0
  ) {
    const startRatio = Number(hit.normMatchStart) / Number(hit.normalizedTextLength);
    const widthRatio = Number(hit.normMatchLength) / Number(hit.normalizedTextLength);

    left += width * startRatio;
    width = Math.max(8, width * widthRatio);
  }

  // inset はズレを強めるのでかなり小さくするか、いったん外す
  const inset = 0;
  left += inset;
  top += inset;
  width = Math.max(6, width - inset * 2);
  height = Math.max(6, height - inset * 2);

  return { left, top, width, height };
}
  
  if (!file) {
    return <div style={ui.textPane}>プレビューがありません。</div>;
  }

  return (
    <div ref={containerRef} style={{ width: "100%", maxWidth: 980, margin: "0 auto" }}>
      <div style={ui.previewToolbar}>
        <span style={ui.badge("blue")}>PDFビューア</span>
        <span style={ui.miniPill}>テキスト選択可</span>
        <span style={ui.miniPill}>候補 {hits.length}件</span>
      </div>

      <Document
        file={file}
        onLoadSuccess={onLoadSuccess}
        onLoadError={onLoadError}
        loading={<div style={ui.textPane}>PDFを読み込み中...</div>}
        error={<div style={ui.textPane}>PDFの表示に失敗しました。</div>}
      >
        {pdfError ? (
          <div style={ui.textPane}>
            PDFの表示に失敗しました。
            {"\n"}
            {pdfError}
          </div>
        ) : (
          Array.from({ length: numPages }, (_, index) => {
            const pageNumber = index + 1;
            const pageHits = hits.filter((h) => Number(h.page || h.page_number || 1) === pageNumber);

            return (
              <div
                key={`page-wrap-${pageNumber}`}
                style={{
                  position: "relative",
                  marginBottom: 18,
                  background: "#fff",
                  borderRadius: 12,
                  overflow: "hidden",
                  boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
                }}
              >
                <div
                  style={{
                    position: "sticky",
                    top: 0,
                    zIndex: 2,
                    background: "rgba(255,255,255,0.96)",
                    padding: "6px 10px",
                    borderBottom: "1px solid #eee",
                    fontSize: 12,
                    fontWeight: 700,
                  }}
                >
                  Page {pageNumber}
                </div>

                <div style={{ position: "relative" }}>
                  <Page
                    pageNumber={pageNumber}
                    width={renderedWidth}
                    renderTextLayer
                    renderAnnotationLayer
                    onLoadSuccess={(page) => handlePageLoadSuccess(page, pageNumber)}
                  />

                  <div style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>
                    {pageHits.map((hit) => {
                      const rect = getScaledRect(hit, pageNumber);
                      if (!rect) return null;

                      const isTop = Number(hit.score || 0) === maxScore;
                      const hitKey = `page-${pageNumber}-hit-${hit.index}`;
                      const isActiveJump = activePreviewHitKey === hitKey;

                      return (
                        <div
                          id={hitKey}
                          key={hitKey}
                          style={{
                            position: "absolute",
                            left: rect.left,
                            top: rect.top,
                            width: rect.width,
                            minHeight: rect.height,
                            background: isActiveJump
                              ? "rgba(255,80,80,0.22)"
                              : isTop
                              ? "rgba(255,140,0,0.16)"
                              : "rgba(255,200,0,0.05)",
                            border: isActiveJump
                              ? "3px solid #ff4d4f"
                              : isTop
                              ? "2px solid #ff8c00"
                              : "1px solid rgba(255,180,0,0.55)",
                            borderRadius: 3,
                            boxSizing: "border-box",
                            boxShadow: isActiveJump ? "0 0 0 4px rgba(255,77,79,0.18)" : "none",
                            transition: "all 0.2s ease",
                          }}
                        />
                      );
                    })}
                  </div>
                </div>
              </div>
            );
          })
        )}
      </Document>

      <div style={ui.previewHint}>
        PDF上のテキストは選択可能です。
      </div>
    </div>
  );
}

function MaterialPreview({ uploadFile, selectedField, activePreviewHitKey }) {
  const isPdf =
    !!uploadFile && (uploadFile.name || "").toLowerCase().endsWith(".pdf");

  if (isPdf) {
    return (
      <PdfPreview
        file={uploadFile}
        selectedField={selectedField}
        activePreviewHitKey={activePreviewHitKey}
      />
    );
  }

  return <div style={ui.textPane}>PDFファイルを選択してください。</div>;
}

export default function VmDiffPage() {
  const [blocks, setBlocks] = useState([]);
  const [vmRows, setVmRows] = useState([]);
  const [vmHeaders, setVmHeaders] = useState([]);
  const [selectedFieldKey, setSelectedFieldKey] = useState("");
  const [selectedCellRef, setSelectedCellRef] = useState("");
  const [eventIdInput, setEventIdInput] = useState("");
  const [uploadFile, setUploadFile] = useState(null);
  const [loadingVm, setLoadingVm] = useState(false);
  const [loadingAnalyze, setLoadingAnalyze] = useState(false);
  const [error, setError] = useState("");
  const [manualCompareText, setManualCompareText] = useState("");
  const [activePreviewHitKey, setActivePreviewHitKey] = useState("");

  const [showImportantOnly, setShowImportantOnly] = useState(false);
  const [onlyHighlightImportant, setOnlyHighlightImportant] = useState(true);
  const [showMatchedBlocksOnly, setShowMatchedBlocksOnly] = useState(true);

  const [sheetQuery, setSheetQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [showOnlyRowsWithDiff, setShowOnlyRowsWithDiff] = useState(false);
  const [sortMode, setSortMode] = useState("sheet");
  const [visibleHeaders, setVisibleHeaders] = useState([]);
  const [pinnedHeaders, setPinnedHeaders] = useState(DEFAULT_PINNED_HEADERS);

  const [headerMenu, setHeaderMenu] = useState(null);
  const [columnFilters, setColumnFilters] = useState({});
  const [columnSort, setColumnSort] = useState({ header: "", direction: "" });

  const loading = loadingVm || loadingAnalyze;

  function resetForNewFlyer() {
  setVmRows([]);
  setVmHeaders([]);
  setBlocks([]);

  setSelectedFieldKey("");
  setSelectedCellRef("");
  setActivePreviewHitKey("");

  setManualCompareText("");

  setHeaderMenu(null);
  setColumnFilters({});
  setColumnSort({ header: "", direction: "" });

  setSheetQuery("");
  setStatusFilter("all");
  setShowOnlyRowsWithDiff(false);

  setVisibleHeaders([]);
  setPinnedHeaders(DEFAULT_PINNED_HEADERS);

  setError("");
}

function handleFlyerFileChange(file) {
  resetForNewFlyer();
  setUploadFile(file || null);
}

  function isSupportedFlyerFile(file) {
    if (!file) return false;
    const name = (file.name || "").toLowerCase();
    return name.endsWith(".pdf");
  }

  async function fetchVmRowsOnly() {
    setError("");
    if (!eventIdInput.trim()) {
      setError("講演会IDを入力してください。");
      return;
    }

    setLoadingVm(true);
    try {
      const res = await fetch(`${API_BASE}/vm-diff/by-event-id`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ event_id: eventIdInput.trim() }),
      });

      if (!res.ok) {
        throw new Error(await readErrorMessage(res, `VM取得に失敗しました: ${res.status}`));
      }

      const data = await res.json();
      setVmRows(Array.isArray(data?.vm_rows) ? data.vm_rows : []);
      setVmHeaders(Array.isArray(data?.headers) ? data.headers : []);
    } catch (e) {
      setError(e?.message || "VM取得に失敗しました。");
    } finally {
      setLoadingVm(false);
    }
  }

  async function analyzeUploadedFlyerTextOnly() {
    setError("");

    if (!uploadFile) {
      setError("案内状ファイルをアップロードしてください。");
      return;
    }

    if (!isSupportedFlyerFile(uploadFile)) {
      setError("サポートされていないファイル形式です。PDFファイルをアップロードしてください。");
      return;
    }

    setLoadingAnalyze(true);
    try {
      const formData = new FormData();
      formData.append("file", uploadFile);
      formData.append("eventId", eventIdInput.trim());

      const res = await fetch(`${API_BASE}/vm-diff/extract-text-blocks`, {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        throw new Error(await readErrorMessage(res, `案内状テキスト取得に失敗しました: ${res.status}`));
      }

      const data = await res.json();
      setBlocks(Array.isArray(data?.blocks) ? data.blocks : []);
      if (Array.isArray(data?.vm_rows)) setVmRows(data.vm_rows);
      if (Array.isArray(data?.headers)) setVmHeaders(data.headers);
    } catch (e) {
      setError(e?.message || "案内状テキスト取得に失敗しました。");
    } finally {
      setLoadingAnalyze(false);
    }
  }

  const flatBlocks = useMemo(() => flattenBlocks(blocks), [blocks]);

  const shaped = useMemo(
    () => shapeVmRows(vmRows, blocks, vmHeaders),
    [vmRows, blocks, vmHeaders]
  );

  const headers = shaped.headers;
  const rows = shaped.rows;

  useEffect(() => {
    if (headers.length && visibleHeaders.length === 0) {
      setVisibleHeaders(headers);
    }
  }, [headers, visibleHeaders.length]);

  useEffect(() => {
    setPinnedHeaders((prev) => prev.filter((h) => headers.includes(h)));
  }, [headers]);

  const filteredRows = useMemo(() => {
    let next = [...rows];

    if (sheetQuery.trim()) {
      const q = sheetQuery.trim().toLowerCase();
      next = next.filter((row) =>
        row.fields.some((f) =>
          String(f.label || "").toLowerCase().includes(q) ||
          String(f.value || "").toLowerCase().includes(q)
        )
      );
    }

    if (statusFilter !== "all") {
      next = next.filter((row) =>
        row.fields.some((f) => f.status === statusFilter)
      );
    }

    if (showOnlyRowsWithDiff) {
      next = next.filter((row) =>
        row.fields.some((f) => f.status === "partial" || f.status === "mismatch")
      );
    }

    next = next.filter((row) => {
      return Object.entries(columnFilters).every(([header, filter]) => {
        const field = row.fields.find((f) => f.label === header);
        const value = String(field?.value || "");
        const status = field?.status || "";

        if (filter.query && !value.toLowerCase().includes(filter.query.toLowerCase())) {
          return false;
        }

        if (filter.statuses?.length && !filter.statuses.includes(status)) {
          return false;
        }

        if (filter.onlyBlank && value.trim() !== "") return false;
        if (filter.onlyNonBlank && value.trim() === "") return false;

        if (filter.selectedValues?.length && !filter.selectedValues.includes(value)) {
          return false;
        }

        return true;
      });
    });

    if (columnSort.header && columnSort.direction) {
      next.sort((a, b) => {
        const av = String(a.fields.find((f) => f.label === columnSort.header)?.value || "");
        const bv = String(b.fields.find((f) => f.label === columnSort.header)?.value || "");
        return columnSort.direction === "asc"
          ? av.localeCompare(bv, "ja")
          : bv.localeCompare(av, "ja");
      });
    } else if (sortMode === "mismatch_desc") {
      next.sort((a, b) => summarizeRow(b).mismatch - summarizeRow(a).mismatch);
    } else if (sortMode === "partial_desc") {
      next.sort((a, b) => summarizeRow(b).partial - summarizeRow(a).partial);
    }

    return next;
  }, [
    rows,
    sheetQuery,
    statusFilter,
    showOnlyRowsWithDiff,
    sortMode,
    columnFilters,
    columnSort,
  ]);

  const allFields = useMemo(() => rows.flatMap((x) => x.fields), [rows]);

  const selectedField = useMemo(
    () => allFields.find((x) => x.key === selectedFieldKey) || allFields[0] || null,
    [allFields, selectedFieldKey]
  );

  useEffect(() => {
    if (!allFields.length) {
      setSelectedFieldKey("");
      setSelectedCellRef("");
      return;
    }

    if (!selectedFieldKey || !allFields.some((x) => x.key === selectedFieldKey)) {
      setSelectedFieldKey(allFields[0].key);
      setSelectedCellRef("A2");
    }
  }, [allFields, selectedFieldKey]);

  const blocksForList = useMemo(() => {
    if (!selectedField) return flatBlocks;
    if (!showMatchedBlocksOnly) return flatBlocks;
    if (!Array.isArray(selectedField.hits) || selectedField.hits.length === 0) return flatBlocks;

    const hitIdx = new Set(selectedField.hits.map((x) => x.index));
    return flatBlocks.filter((b) => hitIdx.has(b.index));
  }, [flatBlocks, selectedField, showMatchedBlocksOnly]);

  const summary = useMemo(() => {
    let match = 0;
    let partial = 0;
    let mismatch = 0;
    let missing = 0;

    for (const f of allFields) {
      if (f.status === "match") match += 1;
      else if (f.status === "partial") partial += 1;
      else if (f.status === "mismatch") mismatch += 1;
      else missing += 1;
    }

    return { match, partial, mismatch, missing, total: allFields.length };
  }, [allFields]);

function handleSelectField(field, cellRef) {
  // 別セルへ移動時も、解除時も比較用の手動文字はクリア
  setManualCompareText("");

  if (!field) {
    setSelectedFieldKey("");
    setSelectedCellRef("");
    setActivePreviewHitKey("");
    return;
  }

  // 同じセルを再クリックしたら解除
  if (selectedFieldKey === field.key) {
    setSelectedFieldKey("");
    setSelectedCellRef("");
    setActivePreviewHitKey("");
    return;
  }

  setSelectedFieldKey(field.key);
  setSelectedCellRef(cellRef || "");

  const firstHit = field?.hits?.[0];
  if (firstHit) {
    const page = Number(firstHit.page || firstHit.page_number || 1);
    const key = `page-${page}-hit-${firstHit.index}`;
    setActivePreviewHitKey(key);

    requestAnimationFrame(() => {
      const el = document.getElementById(key);
      if (el) {
        el.scrollIntoView({
          behavior: "smooth",
          block: "center",
          inline: "nearest",
        });
      }
    });
  } else {
    setActivePreviewHitKey("");
  }
}

  return (
    <div style={ui.page}>
      <div style={ui.leftCol}>
        <div style={ui.card}>
          <div style={{ display: "grid", gap: 12 }}>
            <div>
              <Link to="/" style={{ textDecoration: "none" }}>
                ← 一覧へ
              </Link>
            </div>

            <div style={ui.h2}>シート値 × 案内状比較</div>
            <div style={ui.muted}>
              左でシート確認、右でPDFプレビューを表示します。
            </div>

            <div style={{ display: "grid", gap: 8 }}>
              <label style={ui.h3}>講演会ID</label>
              <input
                value={eventIdInput}
                onChange={(e) => setEventIdInput(e.target.value)}
                placeholder="例: EM2603-0381974"
                style={ui.input}
              />
              {/* <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button type="button" style={ui.btn("secondary")} onClick={fetchVmRowsOnly} disabled={loading}>
                  {loadingVm ? "取得中..." : "VM行を取得"}
                </button>
              </div> */}
            </div>

            <div style={{ display: "grid", gap: 8 }}>
              <label style={ui.h3}>案内状ファイル</label>
              <input
                type="file"
                accept=".pdf,application/pdf"
                onChange={(e) => handleFlyerFileChange(e.target.files?.[0] || null)}
              />
              <div style={ui.muted}>
                対応形式はテキスト抽出可能な PDF のみです。スキャンPDF・画像PDFは対応外です。
              </div>
              <button type="button" style={ui.btn("primary")} onClick={analyzeUploadedFlyerTextOnly} disabled={loading}>
                {loadingAnalyze ? "読込中..." : "シートを読み込む"}
              </button>
            </div>

            {error ? <div style={{ ...ui.muted, color: "#b00020" }}>{error}</div> : null}
          </div>
        </div>

        {/* <SheetToolbar
          rows={rows}
          headers={headers}
          filteredRowCount={filteredRows.length}
          sheetQuery={sheetQuery}
          setSheetQuery={setSheetQuery}
          statusFilter={statusFilter}
          setStatusFilter={setStatusFilter}
          showOnlyRowsWithDiff={showOnlyRowsWithDiff}
          setShowOnlyRowsWithDiff={setShowOnlyRowsWithDiff}
          sortMode={sortMode}
          setSortMode={setSortMode}
          visibleHeaders={visibleHeaders}
          setVisibleHeaders={setVisibleHeaders}
          pinnedHeaders={pinnedHeaders}
          setPinnedHeaders={setPinnedHeaders}
          showImportantOnly={showImportantOnly}
          setShowImportantOnly={setShowImportantOnly}
          onlyHighlightImportant={onlyHighlightImportant}
          setOnlyHighlightImportant={setOnlyHighlightImportant}
        /> */}

        <div style={ui.card}>
          <div style={ui.headerRow}>
            <div style={ui.h3}>集計</div>
            <span style={ui.badge("blue")}>全 {summary.total} セル</span>
          </div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <span style={ui.badge("green")}>一致 {summary.match}</span>
            <span style={ui.badge("yellow")}>部分一致 {summary.partial}</span>
            <span style={ui.badge("red")}>未検出 {summary.mismatch}</span>
            <span style={ui.badge()}>未入力 {summary.missing}</span>
          </div>
        </div>

        {rows.length === 0 ? (
          <div style={ui.card}>
            <div style={ui.muted}>
              左で講演会IDを取得し、右で案内状を読み込むと比較表を表示します。
            </div>
          </div>
        ) : (
            
            filteredRows.length === 0 ? (
              <div style={ui.card}>
                <div style={ui.muted}>条件に一致する行がありません。<br />演題演者（VM）シートに該当するデータがあるか確認してください。</div>
              </div>
            ) : ( 
          <SpreadsheetLikeTable
            rows={filteredRows}
            headers={headers.filter((h) => visibleHeaders.includes(h))}
            selectedKey={selectedField?.key || ""}
            onSelect={handleSelectField}
            showImportantOnly={showImportantOnly}
            onlyHighlightImportant={onlyHighlightImportant}
            pinnedHeaders={pinnedHeaders}
            onOpenHeaderMenu={setHeaderMenu}
            columnFilters={columnFilters}
            columnSort={columnSort}
          />
        ))}

        <CollapseSection
          title="文字比較"
          defaultOpen={true}
          right={<span style={ui.badge("blue")}>{selectedField?.label || "未選択"}</span>}
        >
          <SideBySideComparePanel
            selectedField={selectedField}
            manualCompareText={manualCompareText}
            setManualCompareText={setManualCompareText}
          />
        </CollapseSection>

        <CollapseSection
          title="抽出テキスト"
          right={<span style={ui.badge("blue")}>{blocksForList.length ? blocksForList.length : flatBlocks.length}件</span>}
        >
          <div style={{ display: "grid", gap: 8 }}>
            <div style={ui.toolbarRow}>
              <button
                type="button"
                style={ui.btn(showMatchedBlocksOnly ? "primary" : "secondary")}
                onClick={() => setShowMatchedBlocksOnly((v) => !v)}
              >
                {showMatchedBlocksOnly ? "該当候補のみ表示" : "全ブロック表示"}
              </button>
            </div>

            {((blocksForList.length ? blocksForList : flatBlocks) || []).map((b) => {
              const hit = selectedField?.hits?.find((x) => x.index === b.index);
              const isHit = !!hit;

              return (
                <div
                  key={b.index}
                  style={{
                    ...ui.textPane,
                    borderColor: isHit ? "#d9b84a" : "#eee",
                    background: isHit ? "#fffbe8" : "#fff",
                  }}
                >
                  <div style={{ ...ui.muted, marginBottom: 6 }}>
                    [{b.index}]
                    {b.page ? ` page=${b.page}` : ""}
                    {hit ? ` score=${hit.score}` : ""}
                  </div>
                  {b.text || "—"}
                </div>
              );
            })}

            {flatBlocks.length === 0 ? <div style={ui.textPane}>抽出テキストがありません</div> : null}
          </div>
        </CollapseSection>
      </div>

      <div style={ui.rightCol}>
        <div style={ui.previewOnlyCard}>
          <div style={ui.sectionTitleRow}>
            <div style={ui.h3}>プレビュー</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {selectedField ? (
                <>
                  <span style={ui.badge(statusTone(selectedField.status))}>
                    {statusText(selectedField.status)}
                  </span>
                  <span style={ui.badge("blue")}>{selectedField.label}</span>
                </>
              ) : null}
            </div>
          </div>

          <MaterialPreview
            uploadFile={uploadFile}
            selectedField={selectedField}
            activePreviewHitKey={activePreviewHitKey}
          />
        </div>
      </div>

      <HeaderFilterMenu
        headerMenu={headerMenu}
        setHeaderMenu={setHeaderMenu}
        rows={rows}
        columnFilters={columnFilters}
        setColumnFilters={setColumnFilters}
        columnSort={columnSort}
        setColumnSort={setColumnSort}
      />
    </div>
  );
}