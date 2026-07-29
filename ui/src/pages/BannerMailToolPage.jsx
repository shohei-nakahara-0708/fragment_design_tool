import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { pdfjs } from "react-pdf";

pdfjs.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url,
).toString();

const API_BASE = import.meta.env.VITE_API_BASE || "";
const TARGET_WIDTH = 1200;
const EVENT_CODE_RE = /\bE\s*[-‐‑‒–—―ー－]?\s*(\d{3,6})\b/i;
const LIST_CONFIG_HEADERS = ["vaultID", "製品", "カテゴリー", "リリース日", "終了日", "製品2"];
const PRODUCT_TOOL_MAP = {
  TachoSil: "00_メールツール_CAB",
  Beriplast: "00_メールツール_CAB",
  "Beriplast P": "00_メールツール_CAB",
  Berinert: "00_メールツール_HAE",
  "Berinert SC": "00_メールツール_HAE",
  Andembry: "00_メールツール_HAE",
  "Hizentra-SID": "00_メールツール_SID",
  "Privigen-SID": "00_メールツール_SID",
  "Hizentra-CIDP": "00_メールツール_CIDP/PID",
  "Privigen-CIDP": "00_メールツール_CIDP/PID",
  "Hizentra-PID": "00_メールツール_CIDP/PID",
  "Privigen-PID": "00_メールツール_CIDP/PID",
  Idelvion: "00_メールツール_HEM",
  Afstyla: "00_メールツール_HEM",
};
const PRODUCT_TOOL_LINKS = {
  "00_メールツール_CAB": "https://cslbehring-promomats-us-ft.veevavault.com/ui/#doc_info/245871/",
  "00_メールツール_HAE": "https://cslbehring-promomats-us-ft.veevavault.com/ui/#doc_info/255021/",
  "00_メールツール_SID": "https://cslbehring-promomats-us-ft.veevavault.com/ui/#doc_info/245092/",
  "00_メールツール_CIDP/PID": "https://cslbehring-promomats-us-ft.veevavault.com/ui/#doc_info/207722/",
  "00_メールツール_HEM": "https://cslbehring-promomats-us-ft.veevavault.com/ui/#doc_info/242798/",
};
const HEM_PRODUCT_LABELS = {
  Idelvion: "イデルビオン",
  Afstyla: "エイフスチラ",
};

function normalizeProductName(product) {
  const text = String(product || "").trim();
  return text.replace(/^Berinert\s*[-‐‑‒–—―ー－]\s*SC$/i, "Berinert SC");
}

function formatSize(bytes) {
  if (bytes == null) return "";
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  const mb = kb / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MB`;
  return `${(mb / 1024).toFixed(2)} GB`;
}

function makeId() {
  return globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeLookupKey(value) {
  const name = String(value || "").split(/[\\/]/).pop() || "";
  const stem = name.includes(".") ? name.replace(/\.[^.]+$/, "") : name;
  return stem.replace(/\s|\u3000/g, "").toLowerCase();
}

function normalizeEventCode(value) {
  const match = String(value || "").match(EVENT_CODE_RE);
  return match ? `E-${match[1]}` : "";
}

function getEventCodeMatchKeys(...values) {
  return [...new Set(values.map((value) => normalizeLookupKey(value)).filter(Boolean))];
}

function eventCodeKeysMatch(rowKey, matchKeys) {
  return Boolean(rowKey && matchKeys.some((key) => (
    rowKey === key || key.includes(rowKey) || rowKey.includes(key)
  )));
}

function getErrorMessage(payload, fallback) {
  if (!payload) return fallback;
  if (typeof payload === "string") return payload;
  const detail = payload.detail;
  if (typeof detail === "string") return detail;
  if (detail?.message) return detail.message;
  if (payload.message) return payload.message;
  return fallback;
}

function resultDownloadUrl(url) {
  if (!url) return "";
  const separator = url.includes("?") ? "&" : "?";
  return `${API_BASE}${url}${separator}download=1`;
}

function getProductTool(product) {
  const key = normalizeProductName(product);
  return PRODUCT_TOOL_MAP[key] || (key ? `未設定（${key}）` : "未設定");
}

function getProductToolUrl(tool) {
  return PRODUCT_TOOL_LINKS[tool] || "";
}

function getRowKey(row) {
  return String(row?.id || row?.rowNumber || row?.product || row?.name || "row");
}

function getVaultIdKey(packageKey, row) {
  return `${packageKey}::${getRowKey(row)}`;
}

function getVaultIdValue(vaultIds, packageKey, row) {
  const key = getVaultIdKey(packageKey, row);
  if (Object.prototype.hasOwnProperty.call(vaultIds, key)) return vaultIds[key];
  return String(row?.vaultId || "").replace(/[^\d]/g, "");
}

function csvCell(value) {
  const text = String(value ?? "");
  if (!/[",\r\n]/.test(text)) return text;
  return `"${text.replaceAll('"', '""')}"`;
}

function csvLine(headers, values) {
  return headers.map((header) => csvCell(values[header] || "")).join(",");
}

function formatTodayForCsv() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}/${month}/${day}`;
}

function formatListConfigDate(value) {
  const text = String(value ?? "").trim();
  if (!text) return "";

  const compactMatch = text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (compactMatch) {
    const [, year, month, day] = compactMatch;
    return `${year}/${month}/${day}`;
  }

  const usDateMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (usDateMatch) {
    const [, month, day, year] = usDateMatch;
    return `${year}/${month.padStart(2, "0")}/${day.padStart(2, "0")}`;
  }

  const match = text
    .replace(/[年月]/g, "/")
    .replace(/日/g, "")
    .replace(/[.-]/g, "/")
    .match(/(\d{4})\/\s*(\d{1,2})\/\s*(\d{1,2})/);
  if (!match) return text;

  const [, year, month, day] = match;
  return `${year}/${month.padStart(2, "0")}/${day.padStart(2, "0")}`;
}

function makeListConfigCsvValues(row, vaultId) {
  const product = normalizeProductName(row?.product);
  const tool = getProductTool(product);
  const isHem = tool === "00_メールツール_HEM";
  return {
    vaultID: vaultId,
    製品: isHem ? (HEM_PRODUCT_LABELS[product] || product) : product,
    カテゴリー: row?.mailCategory || "",
    リリース日: formatTodayForCsv(),
    終了日: formatListConfigDate(row?.expirationDate || row?.endDate || ""),
    製品2: isHem ? product : "",
  };
}

function makeListConfigCsvRow(row, vaultId) {
  return csvLine(LIST_CONFIG_HEADERS, makeListConfigCsvValues(row, vaultId));
}

function packToRow(pack) {
  return {
    id: `package-${pack?.zipPath || pack?.path || pack?.htmlFilename}`,
    rowNumber: pack?.rowNumber,
    name: pack?.name || pack?.templateName || pack?.htmlFilename || "",
    detailGroup: pack?.detailGroup || "",
    product: normalizeProductName(pack?.product),
    description: pack?.description || "",
    expirationDate: pack?.expirationDate || "",
    endDate: pack?.endDate || "",
    mailCategory: pack?.mailCategory || "",
    requestDate: pack?.requestDate || "",
    responseType: pack?.responseType || "",
    vaultId: pack?.vaultId || "",
  };
}

function getRelatedRowsForPack(pack, rows) {
  const selected = rows.find((row) => String(row.rowNumber) === String(pack?.rowNumber));
  const name = selected?.name || pack?.name || "";
  const requestDate = selected?.requestDate || pack?.requestDate || "";
  const responseType = selected?.responseType || pack?.responseType || "";
  const related = name
    ? rows.filter((row) => (
      String(row.name || "") === String(name || "")
      && String(row.requestDate || "") === String(requestDate || "")
      && String(row.responseType || "") === String(responseType || "")
    ))
    : [];
  const sourceRows = related.length ? related : [selected || packToRow(pack)];
  return [...sourceRows].sort((a, b) => {
    const aSelected = String(a.rowNumber) === String(pack?.rowNumber) ? 0 : 1;
    const bSelected = String(b.rowNumber) === String(pack?.rowNumber) ? 0 : 1;
    return aSelected - bSelected || normalizeProductName(a.product).localeCompare(normalizeProductName(b.product), "ja");
  });
}

function groupRowsByTool(rows) {
  const map = new Map();
  for (const row of rows) {
    const tool = getProductTool(row.product);
    if (!map.has(tool)) map.set(tool, []);
    map.get(tool).push(row);
  }
  return [...map.entries()].map(([tool, toolRows]) => ({ tool, rows: toolRows }));
}

async function writeClipboardText(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }
  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.setAttribute("readonly", "");
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand("copy");
  document.body.removeChild(textarea);
}

function readImageSize(file) {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      resolve({ width: img.naturalWidth, height: img.naturalHeight, previewUrl: url, previewObjectUrl: true });
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve({ width: 0, height: 0, previewUrl: "", previewObjectUrl: false });
    };
    img.src = url;
  });
}

async function readPdfPreview(file) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const loadingTask = pdfjs.getDocument({ data: arrayBuffer });
    const pdf = await loadingTask.promise;
    const page = await pdf.getPage(1);
    const baseViewport = page.getViewport({ scale: 1 });
    const scale = Math.min(2, 420 / baseViewport.width);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    await page.render({ canvasContext: context, viewport }).promise;
    const previewUrl = canvas.toDataURL("image/png");
    const eventCode = await extractEventCodeFromPdfPage(page, baseViewport);
    await pdf.destroy();
    return {
      width: Math.round(baseViewport.width),
      height: Math.round(baseViewport.height),
      previewUrl,
      previewObjectUrl: false,
      eventCode,
    };
  } catch {
    return { width: 0, height: 0, previewUrl: "", previewObjectUrl: false, eventCode: "" };
  }
}

async function extractEventCodeFromPdfPage(page, baseViewport) {
  const textContent = await page.getTextContent();
  const candidates = [];
  for (const [index, item] of (textContent.items || []).entries()) {
    const text = String(item.str || "");
    const eventCode = normalizeEventCode(text);
    if (!eventCode) continue;

    const tx = item.transform || [];
    const rawX = Number(tx[4] || 0);
    const rawY = Number(tx[5] || 0);
    let viewX = rawX;
    let viewY = rawY;
    try {
      [viewX, viewY] = baseViewport.convertToViewportPoint(rawX, rawY);
    } catch {
      // Keep raw coordinates when PDF.js cannot convert them.
    }
    const pageW = Math.max(1, baseViewport.width || 1);
    const pageH = Math.max(1, baseViewport.height || 1);
    const leftScore = 1 - Math.min(1, Math.max(0, viewX / pageW));
    const bottomScore = Math.min(1, Math.max(0, viewY / pageH));
    candidates.push({
      eventCode,
      index,
      score: (leftScore * 2) + (bottomScore * 2),
    });
  }

  candidates.sort((a, b) => b.score - a.score || a.index - b.index);
  if (candidates[0]?.eventCode) return candidates[0].eventCode;
  return normalizeEventCode((textContent.items || []).map((item) => item.str || "").join(" "));
}

function isSupportedFile(file) {
  const name = file?.name || "";
  return file?.type?.startsWith("image/") || file?.type === "application/pdf" || /\.pdf$/i.test(name);
}

function Field({ label, value, link = false }) {
  return (
    <div className="lecture-tool-field">
      <div className="lecture-tool-field__label">{label}</div>
      <div className="lecture-tool-field__value">
        {link && value ? (
          <a href={value} target="_blank" rel="noreferrer">
            {value}
          </a>
        ) : (
          value || "-"
        )}
      </div>
    </div>
  );
}

function CopyValueCell({ label, value, copied, onCopy, changed = false }) {
  return (
    <div className={`stamp-tool-copy-cell${changed ? " stamp-tool-copy-cell--diff" : ""}`}>
      <div className="stamp-tool-copy-cell__label">{label}</div>
      <div className="stamp-tool-copy-cell__value">{value || "-"}</div>
      <button
        className="stamp-tool-copy-cell__button"
        type="button"
        onClick={onCopy}
        disabled={!value}
      >
        {copied ? "コピー済み" : "コピー"}
      </button>
    </div>
  );
}

function BannerCandidateRows({ rows, selectedRowId, onSelect, disabled }) {
  if (!rows.length) {
    return <div className="lecture-tool-empty">Name列に該当する候補がありません。行番号で指定してください。</div>;
  }
  return (
    <div className="lecture-tool-candidates" aria-label="bannerタブ行候補">
      <div className="lecture-tool-candidates__head">
        <span>候補行</span>
        <em>{rows.length} 件</em>
      </div>
      <div className="banner-tool-candidates__list">
        {rows.map((row) => {
          const active = selectedRowId === row.id;
          return (
            <button
              className={`banner-tool-candidate${active ? " banner-tool-candidate--active" : ""}`}
              type="button"
              key={row.id}
              onClick={() => onSelect(row.id)}
              disabled={disabled}
            >
              <span className="banner-tool-candidate__row">
                <small>行</small>
                {row.rowNumber || "-"}
              </span>
              <span>
                <small>Name</small>
                {row.name || "-"}
              </span>
              <span>
                <small>種別</small>
                {row.responseType || "-"}
              </span>
              <span>
                <small>遷移先URL</small>
                {row.destinationUrl || "-"}
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

function SelectedRowInfo({ row }) {
  if (!row) {
    return <div className="lecture-tool-alert">bannerタブの行を選択してください。</div>;
  }
  return (
    <div>
      <div className="lecture-tool-hint mb_10">以下の内容でHTMLを生成します。</div>
      <div className="lecture-tool-rowinfo">
        <Field label="行" value={row.rowNumber} />
        <Field label="Product" value={normalizeProductName(row.product)} />
        <Field label="Name" value={row.name} />
        <Field label="説明(Description)" value={row.description} />
        <Field label="Expiration Date" value={row.expirationDate} />
        <Field label="メールツール カテゴリー" value={row.mailCategory} />
        <Field label="種別" value={row.responseType} />
        <Field label="遷移先URL" value={row.destinationUrl} link />
      </div>
    </div>
  );
}

function BannerRowNumberField({ item, row, onChange, disabled }) {
  const title = item.autoMatchCount > 0
    ? "候補から選ばず、bannerタブの行番号で指定できます。"
    : "Name列の候補が見つからないため、bannerタブの行番号を入力してください。";

  return (
    <div className="lecture-tool-manual">
      <div className="lecture-tool-manual__title">{title}</div>
      <label className="lecture-tool-select-label">
        行番号
        <input
          className="lecture-tool-search"
          inputMode="numeric"
          value={item.rowNumberInput}
          onChange={(event) => onChange(item.id, event.target.value)}
          placeholder="例: 12"
          disabled={disabled}
        />
      </label>
      {item.rowNumberInput && !row ? (
        <div className="lecture-tool-empty">該当行が読み込まれていません。行番号を確認してください。</div>
      ) : null}
    </div>
  );
}

export default function BannerMailToolPage() {
  const fileInputRef = useRef(null);
  const [rows, setRows] = useState([]);
  const [sheetMeta, setSheetMeta] = useState(null);
  const [items, setItems] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const [loadingRows, setLoadingRows] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [currentProgress, setCurrentProgress] = useState("");
  const [previewItem, setPreviewItem] = useState(null);
  const [vaultIds, setVaultIds] = useState({});
  const [copiedKey, setCopiedKey] = useState("");
  const [openResultSections, setOpenResultSections] = useState({});
  const itemsRef = useRef([]);

  const rowsById = useMemo(() => {
    const map = new Map();
    for (const row of rows) map.set(row.id, row);
    return map;
  }, [rows]);

  const loadRows = useCallback(async () => {
    setLoadingRows(true);
    setError("");
    try {
      const response = await fetch(`${API_BASE}/banner-mail-tool/sheet-rows`, { cache: "no-store" });
      const payload = await response.json().catch(() => null);
      if (!response.ok) {
        throw new Error(getErrorMessage(payload, "スプレッドシートの読み込みに失敗しました。"));
      }
      setRows(payload.rows || []);
      setSheetMeta(payload);
    } catch (err) {
      setError(err?.message || "スプレッドシートの読み込みに失敗しました。");
    } finally {
      setLoadingRows(false);
    }
  }, []);

  useEffect(() => {
    loadRows();
  }, [loadRows]);

  useEffect(() => {
    itemsRef.current = items;
  }, [items]);

  useEffect(() => (
    () => {
      for (const item of itemsRef.current) {
        if (item.previewUrl && item.previewObjectUrl) URL.revokeObjectURL(item.previewUrl);
      }
    }
  ), []);

  const buildItem = useCallback(
    async (file) => {
      const isPdf = file.type === "application/pdf" || /\.pdf$/i.test(file.name);
      const size = isPdf ? await readPdfPreview(file) : await readImageSize(file);
      const matchKeys = getEventCodeMatchKeys(size.eventCode, file.name);
      const matchedRows = rows.filter((row) => (
        eventCodeKeysMatch(row.nameKey, matchKeys) || eventCodeKeysMatch(row.eventCodeKey, matchKeys)
      ));
      return {
        id: makeId(),
        file,
        filename: file.name,
        size: file.size,
        isPdf,
        detectedEventCode: size.eventCode || normalizeEventCode(file.name),
        width: size.width,
        height: size.height,
        previewUrl: size.previewUrl,
        previewObjectUrl: size.previewObjectUrl,
        selectedRowId: matchedRows.length === 1 ? matchedRows[0].id : "",
        matchedRowIds: matchedRows.map((row) => row.id),
        autoMatchCount: matchedRows.length,
        rowNumberInput: "",
        rowNumberMode: matchedRows.length === 0,
        mergePdfPages: false,
      };
    },
    [rows],
  );

  const addFiles = useCallback(async (rawFiles) => {
    if (loadingRows || !rows.length || generating) return;
    const files = Array.from(rawFiles || []).filter(isSupportedFile);
    if (!files.length) {
      setError("画像またはPDFをアップロードしてください。");
      return;
    }
    setError("");
    setResult(null);
    setVaultIds({});
    setCopiedKey("");
    setOpenResultSections({});
    setCurrentProgress("");
    const nextItems = await Promise.all(files.map(buildItem));
    setItems((prev) => [...prev, ...nextItems]);
  }, [buildItem, generating, loadingRows, rows.length]);

  const onSelectFiles = async (event) => {
    await addFiles(event.target.files);
    event.target.value = "";
  };

  const onDropFiles = async (event) => {
    event.preventDefault();
    setDragOver(false);
    await addFiles(event.dataTransfer.files);
  };

  const removeItem = (id) => {
    setItems((prev) => {
      const target = prev.find((item) => item.id === id);
      if (target?.previewUrl && target.previewObjectUrl) URL.revokeObjectURL(target.previewUrl);
      if (previewItem?.id === id) setPreviewItem(null);
      return prev.filter((item) => item.id !== id);
    });
  };

  const updateSelectedRow = (id, selectedRowId) => {
    setItems((prev) => prev.map((item) => (
      item.id === id ? { ...item, selectedRowId, rowNumberInput: "", rowNumberMode: false } : item
    )));
  };

  const updateRowNumber = (id, value) => {
    const normalized = value.replace(/[^\d]/g, "");
    setItems((prev) => prev.map((item) => (
      item.id === id ? { ...item, rowNumberInput: normalized, selectedRowId: "", rowNumberMode: true } : item
    )));
  };

  const updateRowMode = (id, rowNumberMode) => {
    setItems((prev) => prev.map((item) => (
      item.id === id ? { ...item, rowNumberMode } : item
    )));
  };

  const updatePdfMergeMode = (id, mergePdfPages) => {
    setItems((prev) => prev.map((item) => (
      item.id === id ? { ...item, mergePdfPages } : item
    )));
  };

  const updateVaultId = (packageKey, row, value) => {
    const normalized = value.replace(/[^\d]/g, "");
    setVaultIds((prev) => ({ ...prev, [getVaultIdKey(packageKey, row)]: normalized }));
  };

  const toggleResultSection = (key) => {
    setOpenResultSections((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const copyText = async (key, value) => {
    if (!value) return;
    try {
      await writeClipboardText(value);
      setCopiedKey(key);
      window.setTimeout(() => {
        setCopiedKey((current) => (current === key ? "" : current));
      }, 1600);
    } catch {
      setError("クリップボードへのコピーに失敗しました。");
    }
  };

  const getEffectiveRowId = (item) => item.rowNumberMode ? item.rowNumberInput : item.selectedRowId;
  const sheetReady = rows.length > 0 && !loadingRows;
  const rowMissing = items.filter((item) => !rowsById.has(getEffectiveRowId(item)));
  const canGenerate = sheetReady && items.length > 0 && rowMissing.length === 0 && !generating;

  const generate = async () => {
    if (!canGenerate) {
      setError("すべてのファイルにbannerタブの行を選択するか、行番号を入力してください。");
      return;
    }

    setGenerating(true);
    setError("");
    setResult(null);
    setVaultIds({});
    setCopiedKey("");
    setOpenResultSections({});
    setCurrentProgress("アップロード準備中...");
    try {
      const formData = new FormData();
      for (const item of items) {
        formData.append("files", item.file, item.filename);
        formData.append("rowIds", getEffectiveRowId(item));
        formData.append("mergePdfPages", item.isPdf && item.mergePdfPages ? "true" : "false");
      }
      setCurrentProgress("HTMLとZIPを生成しています...");
      const response = await fetch(`${API_BASE}/banner-mail-tool/generate`, {
        method: "POST",
        body: formData,
      });
      const payload = await response.json().catch(() => null);
      if (!response.ok) {
        throw new Error(getErrorMessage(payload, "生成に失敗しました。"));
      }
      setResult(payload);
      setCurrentProgress("完了しました。");
    } catch (err) {
      setError(err?.message || "生成に失敗しました。");
      setCurrentProgress("エラーで停止しました。");
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="lecture-tool-page">
      <div className="lecture-tool-page__inner">
        <header className="lecture-tool-header">
          <div>
            <h1 className="lecture-tool-header__title">CSLバナーメール生成</h1>
            <div className="lecture-tool-header__sub">
              bannerタブのName行を選び、アップロード画像またはPDFからメールHTMLとimages.zipを生成します。
            </div>
          </div>
          <div className="lecture-tool-actions">
            <button className="lecture-tool-button" type="button" onClick={loadRows} disabled={loadingRows || generating}>
              {loadingRows ? "読込中" : "シートを再読込"}
            </button>
            <button
              className="lecture-tool-button lecture-tool-button--primary"
              type="button"
              onClick={generate}
              disabled={!canGenerate}
            >
              {generating ? "生成中" : "生成"}
            </button>
          </div>
        </header>

        {error ? <div className="lecture-tool-alert">{error}</div> : null}

        <div className="lecture-tool-grid">
          <div className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">画像 / PDFアップロード</h2>
                  <div className="lecture-tool-panel__sub">
                    画像は幅{TARGET_WIDTH}pxを超える場合のみ縮小します。PDFは1ページ目を幅{TARGET_WIDTH}pxのbanner.jpgに変換します。
                  </div>
                </div>
                <span className="lecture-tool-status lecture-tool-status--ok">
                  {items.length} ファイル
                </span>
              </div>

              <div
                className={`lecture-tool-drop${dragOver ? " lecture-tool-drop--active" : ""}${!sheetReady || generating ? " lecture-tool-drop--disabled" : ""}`}
                onClick={() => {
                  if (sheetReady && !generating) fileInputRef.current?.click();
                }}
                onDragOver={(event) => {
                  event.preventDefault();
                  if (sheetReady && !generating) setDragOver(true);
                }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(event) => {
                  if (!sheetReady || generating) {
                    event.preventDefault();
                    setDragOver(false);
                    return;
                  }
                  onDropFiles(event);
                }}
                role="button"
                tabIndex={0}
              >
                <div className="lecture-tool-drop__title">
                  {sheetReady ? "画像またはPDFを選択 / ドラッグ&ドロップ" : "参照シート読み込み後にアップロードできます"}
                </div>
                <div className="lecture-tool-drop__sub">
                  {loadingRows ? "シートを読み込み中です。" : "JPG / PNG / PDF などを複数アップロードできます。"}
                </div>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept="image/*,application/pdf,.pdf"
                  multiple
                  onChange={onSelectFiles}
                  className="lecture-tool-hidden-input"
                  disabled={!sheetReady || generating}
                />
              </div>
            </div>
          </div>

          <aside className="lecture-tool-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">参照シート</h2>
                </div>
              </div>
              <div className="lecture-tool-meta">
                <div>
                  <strong>spreadsheet</strong>
                  <br />
                  {sheetMeta?.spreadsheetUrl ? (
                    <a href={sheetMeta.spreadsheetUrl} target="_blank" rel="noreferrer">
                      {sheetMeta?.spreadsheetTitle || "スプレッドシートを開く"}
                    </a>
                  ) : (
                    sheetMeta?.spreadsheetTitle || "-"
                  )}
                </div>
                <div>
                  <strong>シート名</strong>
                  <br />
                  {sheetMeta?.sheetTitle || "-"}
                </div>
                <div>
                  <strong>行数</strong>
                  <br />
                  {loadingRows ? "読み込み中" : `${rows.length > 0 ? rows.length + 3 : 0} 件`}
                </div>
                <div>
                  <strong>必要列</strong>
                  <br />
                  Name / 遷移先URL / 種別
                </div>
              </div>
            </div>
          </aside>

          <div className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">行の選択</h2>
                  <div className="lecture-tool-panel__sub">Name列がHTML名になります。候補が複数ある場合は該当行を選択してください。</div>
                </div>
              </div>

              {!items.length ? (
                <div className="lecture-tool-empty">画像またはPDFをアップロードすると、ここに選択欄が表示されます。</div>
              ) : (
                <div className="lecture-tool-image-list">
                  {items.map((item) => {
                    const optionRows = (item.matchedRowIds || [])
                      .map((rowId) => rowsById.get(rowId))
                      .filter(Boolean);
                    const hasCandidateRows = optionRows.length > 0;
                    const showRowSelect = hasCandidateRows && !item.rowNumberMode;
                    const manualRow = item.rowNumberMode ? rowsById.get(item.rowNumberInput) : null;
                    const selectedRow = rowsById.get(getEffectiveRowId(item));
                    const isLargeImage = !item.isPdf && item.width > TARGET_WIDTH;
                    const isSmallImage = !item.isPdf && item.width > 0 && item.width <= TARGET_WIDTH;

                    return (
                      <div className="lecture-tool-image-card stamp-tool-card" key={item.id}>
                        <div className="lecture-tool-image-card__media stamp-tool-card__media">
                          {item.previewUrl ? (
                            <button
                              className="lecture-tool-image-card__preview"
                              type="button"
                              onClick={() => setPreviewItem(item)}
                              aria-label={`${item.filename}を拡大表示`}
                            >
                              <img src={item.previewUrl} alt="" />
                            </button>
                          ) : (
                            <div className="stamp-tool-pdf-preview">
                              <strong>PDF</strong>
                              <span>1ページ目をbanner.jpgに変換</span>
                            </div>
                          )}
                        </div>
                        <div className="lecture-tool-image-card__body">
                          <div className="lecture-tool-image-card__top">
                            <div>
                              <div className="lecture-tool-image-card__title">{item.filename}</div>
                              <div className="lecture-tool-image-card__sub">
                                {formatSize(item.size)}
                                {item.isPdf ? " / PDF" : ` / ${item.width || "-"} x ${item.height || "-"}px`}
                              </div>
                            </div>
                            <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={() => removeItem(item.id)} disabled={generating}>
                              削除
                            </button>
                          </div>

                          <div className="lecture-tool-tag-row">
                            {item.autoMatchCount === 1 ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">Name列で自動一致</span>
                            ) : item.autoMatchCount > 1 ? (
                              <span className="lecture-tool-status lecture-tool-status--warn">同コード候補 {item.autoMatchCount} 件</span>
                            ) : (
                              <span className="lecture-tool-status lecture-tool-status--warn">手動選択</span>
                            )}
                            {isLargeImage ? (
                              <span className="lecture-tool-status lecture-tool-status--warn">幅{TARGET_WIDTH}pxへ縮小</span>
                            ) : null}
                            {isSmallImage ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">原寸JPG</span>
                            ) : null}
                            {item.isPdf && item.mergePdfPages ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">全ページを1枚JPG化</span>
                            ) : item.isPdf ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">1ページ目を幅{TARGET_WIDTH}pxで生成</span>
                            ) : null}
                            <span className="lecture-tool-status">banner.jpg</span>
                          </div>

                          {item.autoMatchCount > 1 && !item.selectedRowId && !item.rowNumberMode ? (
                            <div className="lecture-tool-hint">候補が複数あります。該当するスプレッドシート行を選択してください。</div>
                          ) : null}

                          <div className="lecture-tool-mode" role="group" aria-label="行指定方法">
                            <button
                              className={`lecture-tool-mode__button${!item.rowNumberMode ? " lecture-tool-mode__button--active" : ""}`}
                              type="button"
                              onClick={() => updateRowMode(item.id, false)}
                              disabled={!hasCandidateRows || generating}
                            >
                              候補から選択
                            </button>
                            <button
                              className={`lecture-tool-mode__button${item.rowNumberMode ? " lecture-tool-mode__button--active" : ""}`}
                              type="button"
                              onClick={() => updateRowMode(item.id, true)}
                              disabled={generating}
                            >
                              行番号で指定
                            </button>
                          </div>

                          {item.isPdf ? (
                            <label className="stamp-tool-pdf-mode">
                              <input
                                type="checkbox"
                                checked={item.mergePdfPages}
                                onChange={(event) => updatePdfMergeMode(item.id, event.target.checked)}
                                disabled={generating}
                              />
                              <span>
                                <strong>PDF全ページを1枚画像化</strong>
                                <small>全ページを縦につなげてbanner.jpgを生成します。</small>
                              </span>
                            </label>
                          ) : null}

                          {showRowSelect ? (
                            <BannerCandidateRows
                              rows={optionRows}
                              selectedRowId={item.selectedRowId}
                              onSelect={(rowId) => updateSelectedRow(item.id, rowId)}
                              disabled={loadingRows || generating}
                            />
                          ) : null}

                          {!showRowSelect ? (
                            <BannerRowNumberField item={item} row={manualRow} onChange={updateRowNumber} disabled={generating} />
                          ) : null}

                          <SelectedRowInfo row={selectedRow} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>

          {(generating || currentProgress) ? (
            <div className="lecture-tool-panel lecture-tool-panel--wide">
              <div className="lecture-tool-panel__inner">
                <div className="lecture-tool-progress">
                  <div className="lecture-tool-progress__head">
                    <div>
                      <div className="lecture-tool-progress__title">処理状況</div>
                      <div className="lecture-tool-progress__current">{currentProgress || "待機中"}</div>
                    </div>
                    {generating ? <span className="lecture-tool-progress__badge">実行中</span> : null}
                  </div>
                </div>
              </div>
            </div>
          ) : null}

          <div className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">生成結果</h2>
                </div>
              </div>

              {result?.packages?.length ? (
                <div className="lecture-tool-package-list">
                  {result.packages.map((pack) => {
                    const packageKey = pack.zipPath || pack.path || pack.htmlFilename;
                    const relatedRows = getRelatedRowsForPack(pack, rows);
                    const selectedVaultRow = rows.find((row) => String(row.rowNumber) === String(pack?.rowNumber)) || packToRow(pack);
                    const csvGroups = groupRowsByTool(relatedRows);
                    const vaultSectionKey = `${packageKey}-vault`;
                    const csvSectionKey = `${packageKey}-csv`;
                    const vaultOpen = Boolean(openResultSections[vaultSectionKey]);
                    const csvOpen = Boolean(openResultSections[csvSectionKey]);
                    return (
                      <div className="lecture-tool-package" key={packageKey}>
                        <div className="lecture-tool-package__top">
                          <div>
                            <div className="lecture-tool-package__title">{pack.htmlFilename}</div>
                            <div className="lecture-tool-package__sub">{pack.destinationUrl}</div>
                          </div>
                          <a className="lecture-tool-file__link" href={resultDownloadUrl(pack.zipUrl)}>
                            ZIPダウンロード
                          </a>
                        </div>
                        <div className="lecture-tool-result-statuses">
                          <span className="lecture-tool-status lecture-tool-status--ok">HTML作成済み</span>
                          <span className="lecture-tool-status lecture-tool-status--ok">images.zip作成済み</span>
                          <span className="lecture-tool-status lecture-tool-status--ok">
                            {pack.imageInfo?.width || TARGET_WIDTH}px
                          </span>
                        </div>

                        <div className={`lecture-tool-result-detail stamp-tool-collapsible${vaultOpen ? " stamp-tool-collapsible--open" : ""}`}>
                          <button
                            className="stamp-tool-collapsible__button"
                            type="button"
                            onClick={() => toggleResultSection(vaultSectionKey)}
                            aria-expanded={vaultOpen}
                          >
                            <span>Vault登録用</span>
                            <em>{relatedRows.length} 行 / {vaultOpen ? "閉じる" : "開く"}</em>
                          </button>
                          {vaultOpen ? (
                            <>
                              {relatedRows.length > 1 ? (
                                <div className="lecture-tool-hint">
                                  同じName・依頼日・種別の行が {relatedRows.length} 件あります。Vault登録用に他製品版も表示しています。
                                </div>
                              ) : null}
                              <div className="stamp-tool-vault-list">
                                {relatedRows.map((row) => {
                                  const rowKey = getRowKey(row);
                                  const isSelectedVaultRow = String(row.rowNumber) === String(selectedVaultRow.rowNumber);
                                  const cells = [
                                    ["Detail Group", row.detailGroup, selectedVaultRow.detailGroup],
                                    ["Product", normalizeProductName(row.product), normalizeProductName(selectedVaultRow.product)],
                                    ["Name", row.name, selectedVaultRow.name],
                                    ["説明(Description)", row.description, selectedVaultRow.description],
                                    ["Expiration Date / 失効希望日", row.expirationDate, selectedVaultRow.expirationDate],
                                  ];
                                  return (
                                    <div className="stamp-tool-vault-card" key={rowKey}>
                                      <div className="stamp-tool-vault-card__head">
                                        <strong>{normalizeProductName(row.product) || "Product未設定"}</strong>
                                        <span>行 {row.rowNumber || "-"}</span>
                                        <span>{getProductTool(row.product)}</span>
                                        {row.requestDate ? <span>依頼日 {row.requestDate}</span> : null}
                                        {row.responseType ? <span>種別 {row.responseType}</span> : null}
                                      </div>
                                      <div className="stamp-tool-vault-copy-grid">
                                        {cells.map(([label, value, selectedValue]) => {
                                          const key = `${packageKey}-vault-${rowKey}-${label}`;
                                          const changed = !isSelectedVaultRow && String(value || "") !== String(selectedValue || "");
                                          return (
                                            <CopyValueCell
                                              key={label}
                                              label={label}
                                              value={value}
                                              copied={copiedKey === key}
                                              changed={changed}
                                              onCopy={() => copyText(key, value || "")}
                                            />
                                          );
                                        })}
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            </>
                          ) : null}
                        </div>

                        <div className={`lecture-tool-result-detail stamp-tool-collapsible${csvOpen ? " stamp-tool-collapsible--open" : ""}`}>
                          <button
                            className="stamp-tool-collapsible__button"
                            type="button"
                            onClick={() => toggleResultSection(csvSectionKey)}
                            aria-expanded={csvOpen}
                          >
                            <span>LIST_CONFIG.csv更新用</span>
                            <em>{csvGroups.length} ツール / {csvOpen ? "閉じる" : "開く"}</em>
                          </button>
                          {csvOpen ? (
                            <div className="stamp-tool-csv-groups">
                              {csvGroups.map((group) => {
                                const groupKey = `${packageKey}-csv-${group.tool}`;
                                const toolUrl = getProductToolUrl(group.tool);
                                const getRowVaultId = (row) => getVaultIdValue(vaultIds, packageKey, row);
                                const allVaultIdsReady = group.rows.every((row) => Boolean(getRowVaultId(row)));
                                const csvText = group.rows
                                  .map((row) => makeListConfigCsvRow(row, getRowVaultId(row)))
                                  .join("\n");
                                return (
                                  <div className="stamp-tool-csv-group" key={group.tool}>
                                    <div className="stamp-tool-csv-group__head">
                                      <div>
                                        <strong>{group.tool}</strong>
                                        <span>{group.rows.length} 行 / LIST_CONFIG.csv</span>
                                        {toolUrl ? (
                                          <a className="stamp-tool-csv-tool-link" href={toolUrl} target="_blank" rel="noreferrer">
                                            該当ツールを開く
                                          </a>
                                        ) : null}
                                      </div>
                                      <button
                                        className="lecture-tool-button lecture-tool-button--small"
                                        type="button"
                                        onClick={() => copyText(groupKey, csvText)}
                                        disabled={!allVaultIdsReady}
                                      >
                                        {copiedKey === groupKey ? "コピー済み" : group.rows.length > 1 ? "このツール分をコピー" : "CSV行コピー"}
                                      </button>
                                    </div>
                                    <div className="stamp-tool-csv-table-wrap">
                                      <table className="stamp-tool-csv-table">
                                        <thead>
                                          <tr>
                                            {LIST_CONFIG_HEADERS.map((header) => (
                                              <th key={header}>{header}</th>
                                            ))}
                                          </tr>
                                        </thead>
                                        <tbody>
                                          {group.rows.map((row) => {
                                            const rowVaultId = getRowVaultId(row);
                                            const values = makeListConfigCsvValues(row, rowVaultId);
                                            return (
                                              <tr key={getRowKey(row)}>
                                                {LIST_CONFIG_HEADERS.map((header) => (
                                                  <td key={header} className={values[header] ? "" : "stamp-tool-csv-table__empty"}>
                                                    {header === "vaultID" ? (
                                                      <input
                                                        className="stamp-tool-csv-input"
                                                        inputMode="numeric"
                                                        value={rowVaultId}
                                                        onChange={(event) => updateVaultId(packageKey, row, event.target.value)}
                                                        placeholder="Vault ID"
                                                      />
                                                    ) : (
                                                      values[header] || "-"
                                                    )}
                                                  </td>
                                                ))}
                                              </tr>
                                            );
                                          })}
                                        </tbody>
                                      </table>
                                    </div>
                                    {!allVaultIdsReady ? (
                                      <div className="lecture-tool-empty">Vault IDを入力するとCSVコピーできます。</div>
                                    ) : null}
                                  </div>
                                );
                              })}
                            </div>
                          ) : null}
                        </div>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div className="lecture-tool-empty">まだ生成されていません。</div>
              )}
            </div>
          </div>
        </div>
      </div>

      {previewItem ? (
        <div className="lecture-tool-lightbox" role="dialog" aria-modal="true" aria-label="画像プレビュー" onClick={() => setPreviewItem(null)}>
          <div className="lecture-tool-lightbox__body" onClick={(event) => event.stopPropagation()}>
            <div className="lecture-tool-lightbox__head">
              <div className="lecture-tool-lightbox__title">{previewItem.filename}</div>
              <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={() => setPreviewItem(null)}>
                閉じる
              </button>
            </div>
            <img className="lecture-tool-lightbox__image" src={previewItem.previewUrl} alt="" />
          </div>
        </div>
      ) : null}
    </div>
  );
}
