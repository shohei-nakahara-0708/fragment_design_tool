import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { pdfjs } from "react-pdf";

pdfjs.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url,
).toString();

const API_BASE = import.meta.env.VITE_API_BASE || "";
const TARGET_WIDTH = 2048;
const EVENT_CODE_RE = /\bE\s*[-‐‑‒–—―ー－]?\s*(\d{3,6})\b/i;
const LIST_CONFIG_HEADERS = ["Product", "EventDate", "Name", "FileName", "PresentationId", "ThumnailName", "Pickup", "Date", "EndDate"];
const LECTURE_TOOL_CONFIG_SURVEY_URL = "https://cslbehring-promomats-us-ft.veevavault.com/ui/#doc_info/283601";

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

function getMatchKeys(...values) {
  return [...new Set(values.map((value) => normalizeLookupKey(value)).filter(Boolean))];
}

function keysMatch(rowKey, matchKeys) {
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

function getRowKey(row) {
  return String(row?.id || row?.rowNumber || row?.product || row?.presentationId || row?.mediaFileName || "row");
}

function csvCell(value) {
  const text = String(value ?? "");
  if (!/[",\r\n]/.test(text)) return text;
  return `"${text.replaceAll('"', '""')}"`;
}

function csvLine(headers, values) {
  return headers.map((header) => csvCell(values[header] ?? "")).join(",");
}

function formatTodayForCsv() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function makeListConfigCsvValues(row) {
  return {
    Product: normalizeProductName(row?.product),
    EventDate: row?.eventDate || "",
    Name: row?.listConfigName || row?.presentationName || row?.eventName || "",
    FileName: row?.mediaFileName || "",
    PresentationId: row?.presentationId || "",
    ThumnailName: row?.thumbnailName || "",
    Pickup: row?.pickup || "",
    Date: row?.configDate || formatTodayForCsv(),
    EndDate: row?.endDate || "",
  };
}

function makeListConfigCsvRow(row) {
  return csvLine(LIST_CONFIG_HEADERS, makeListConfigCsvValues(row));
}

function getRelatedRowsForPack(pack, rows) {
  if (Array.isArray(pack?.rows) && pack.rows.length) return pack.rows;
  const lectureId = pack?.lectureId || "";
  const mediaFileName = pack?.mediaFileName || "";
  const related = rows.filter((row) => (
    String(row.lectureId || "") === String(lectureId || "")
    && String(row.mediaFileName || "") === String(mediaFileName || "")
  ));
  return related.length ? related : [pack].filter(Boolean);
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
    const eventCode = await extractEventCodeFromPdfPage(page, baseViewport);
    const previewUrl = canvas.toDataURL("image/png");
    const pageCount = pdf.numPages || 1;
    await pdf.destroy();
    return {
      width: Math.round(baseViewport.width),
      height: Math.round(baseViewport.height),
      previewUrl,
      previewObjectUrl: false,
      eventCode,
      pageCount,
    };
  } catch {
    return { width: 0, height: 0, previewUrl: "", previewObjectUrl: false, eventCode: "", pageCount: 0 };
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
      // PDF.js座標変換に失敗した場合はraw座標で近似します。
    }
    const pageW = Math.max(1, baseViewport.width || 1);
    const pageH = Math.max(1, baseViewport.height || 1);
    const leftScore = 1 - Math.min(1, Math.max(0, viewX / pageW));
    const bottomScore = Math.min(1, Math.max(0, viewY / pageH));
    candidates.push({ eventCode, index, score: (leftScore * 2) + (bottomScore * 2) });
  }

  candidates.sort((a, b) => b.score - a.score || a.index - b.index);
  if (candidates[0]?.eventCode) return candidates[0].eventCode;
  return normalizeEventCode((textContent.items || []).map((item) => item.str || "").join(" "));
}

function isSupportedFile(file) {
  const name = file?.name || "";
  return file?.type?.startsWith("image/") || file?.type === "application/pdf" || /\.pdf$/i.test(name);
}

function Field({ label, value }) {
  return (
    <div className="lecture-tool-field">
      <div className="lecture-tool-field__label">{label}</div>
      <div className="lecture-tool-field__value">{value || "-"}</div>
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

function CandidateRows({ rows, selectedRowId, onSelect, disabled }) {
  if (!rows.length) {
    return <div className="lecture-tool-empty">講演会IDで一致する候補がありません。行番号で指定してください。</div>;
  }
  return (
    <div className="lecture-tool-candidates" aria-label="講演会ツール行候補">
      <div className="lecture-tool-candidates__head">
        <span>候補行</span>
        <em>{rows.length} 件</em>
      </div>
      <div className="lecture-tool-candidates__list">
        {rows.map((row) => {
          const active = selectedRowId === row.id;
          return (
            <button
              className={`lecture-tool-candidate${active ? " lecture-tool-candidate--active" : ""}`}
              type="button"
              key={row.id}
              onClick={() => onSelect(row.id)}
              disabled={disabled}
            >
              <span className="lecture-tool-candidate__cell lecture-tool-candidate__cell--strong">
                <small>行</small>
                {row.rowNumber || "-"}
              </span>
              <span className="lecture-tool-candidate__cell lecture-tool-candidate__cell--strong">
                <small>講演会ID</small>
                {row.lectureId || "-"}
              </span>
              <span className="lecture-tool-candidate__cell">
                <small>Product</small>
                {normalizeProductName(row.product) || "-"}
              </span>
              <span className="lecture-tool-candidate__cell lecture-tool-candidate__cell--wide">
                <small>メディアファイル名</small>
                {row.mediaFileName || "-"}
              </span>
              <span className="lecture-tool-candidate__cell">
                <small>種別</small>
                {row.responseType || "-"}
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
    return <div className="lecture-tool-alert">講演会ツールタブの行を選択してください。</div>;
  }
  return (
    <div>
      <div className="lecture-tool-hint mb_10">以下の行を基準に、同じ講演会IDの他製品版も確認して生成します。</div>
      <div className="lecture-tool-rowinfo">
        <Field label="行" value={row.rowNumber} />
        <Field label="講演会ID" value={row.lectureId} />
        <Field label="Product" value={normalizeProductName(row.product)} />
        <Field label="プレゼンテーション名" value={row.presentationName} />
        <Field label="PresentationId" value={row.presentationId} />
        <Field label="メディアファイル名" value={row.mediaFileName} />
        <Field label="ThumnailName" value={row.thumbnailName} />
        <Field label="種別" value={row.responseType} />
        <Field label="講演会名" value={row.eventName} />
      </div>
    </div>
  );
}

function RowNumberField({ item, row, onChange, disabled }) {
  const title = item.autoMatchCount > 0
    ? "候補から選ばず、講演会ツールタブの行番号で指定できます。"
    : "講演会IDの候補が見つからないため、講演会ツールタブの行番号を入力してください。";

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

export default function CslLectureToolPage() {
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
      const response = await fetch(`${API_BASE}/csl-lecture-tool/sheet-rows`, { cache: "no-store" });
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
      const detectedLectureId = size.eventCode || normalizeEventCode(file.name);
      const matchKeys = getMatchKeys(detectedLectureId, file.name);
      const matchedRows = rows.filter((row) => keysMatch(row.lectureIdKey, matchKeys));
      return {
        id: makeId(),
        file,
        filename: file.name,
        size: file.size,
        isPdf,
        detectedLectureId,
        width: size.width,
        height: size.height,
        pageCount: size.pageCount || 0,
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
      setError("すべてのファイルに講演会ツールタブの行を選択するか、行番号を入力してください。");
      return;
    }

    setGenerating(true);
    setError("");
    setResult(null);
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
      setCurrentProgress("講演会ツールZIPを生成しています...");
      const response = await fetch(`${API_BASE}/csl-lecture-tool/generate`, {
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
            <h1 className="lecture-tool-header__title">CSL講演会ツール生成</h1>
            <div className="lecture-tool-header__sub">
              講演会ツールタブの行を選び、アップロード画像またはPDFからHTML・full/thumb画像・ZIPを生成します。
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
                    1.jpgは幅{TARGET_WIDTH}pxを超える場合のみ縮小します。PDFは1ページ目、または全ページを縦結合して生成できます。
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
                  {loadingRows ? "読み込み中" : `${rows.length} 件`}
                </div>
                <div>
                  <strong>必要列</strong>
                  <br />
                  講演会ID / メディアファイル名
                </div>
              </div>
            </div>
          </aside>

          <div className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">行の選択</h2>
                  <div className="lecture-tool-panel__sub">アップロードファイルまたはPDF内の講演会IDで、講演会ツールタブの候補行を表示します。</div>
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
                    const isSmallImage = !item.isPdf && item.width > 0 && item.width < TARGET_WIDTH;
                    const isExactImage = !item.isPdf && item.width === TARGET_WIDTH;

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
                              <span>1.jpgに変換</span>
                            </div>
                          )}
                        </div>
                        <div className="lecture-tool-image-card__body">
                          <div className="lecture-tool-image-card__top">
                            <div>
                              <div className="lecture-tool-image-card__title">{item.filename}</div>
                              <div className="lecture-tool-image-card__sub">
                                {formatSize(item.size)}
                                {item.isPdf ? ` / PDF${item.pageCount ? ` ${item.pageCount}ページ` : ""}` : ` / ${item.width || "-"} x ${item.height || "-"}px`}
                              </div>
                            </div>
                            <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={() => removeItem(item.id)} disabled={generating}>
                              削除
                            </button>
                          </div>

                          <div className="lecture-tool-tag-row">
                            {item.detectedLectureId ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">講演会ID {item.detectedLectureId}</span>
                            ) : null}
                            {item.autoMatchCount === 1 ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">講演会IDで自動一致</span>
                            ) : item.autoMatchCount > 1 ? (
                              <span className="lecture-tool-status lecture-tool-status--warn">同ID候補 {item.autoMatchCount} 件</span>
                            ) : (
                              <span className="lecture-tool-status lecture-tool-status--warn">手動選択</span>
                            )}
                            {isLargeImage ? (
                              <span className="lecture-tool-status lecture-tool-status--warn">幅{TARGET_WIDTH}pxへ縮小</span>
                            ) : null}
                            {isSmallImage ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">原寸JPG</span>
                            ) : null}
                            {isExactImage ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">幅{TARGET_WIDTH}px</span>
                            ) : null}
                            {item.isPdf && item.mergePdfPages ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">全ページを1枚JPG化</span>
                            ) : item.isPdf ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">1ページ目を幅{TARGET_WIDTH}pxで生成</span>
                            ) : null}
                          </div>

                          {item.autoMatchCount > 1 && !item.selectedRowId && !item.rowNumberMode ? (
                            <div className="lecture-tool-hint">講演会IDの候補が複数あります。該当するスプレッドシート行を選択してください。</div>
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
                                <small>全ページを縦につなげて1.jpgを生成します。</small>
                              </span>
                            </label>
                          ) : null}

                          {showRowSelect ? (
                            <CandidateRows
                              rows={optionRows}
                              selectedRowId={item.selectedRowId}
                              onSelect={(rowId) => updateSelectedRow(item.id, rowId)}
                              disabled={loadingRows || generating}
                            />
                          ) : null}

                          {!showRowSelect ? (
                            <RowNumberField item={item} row={manualRow} onChange={updateRowNumber} disabled={generating} />
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
                    const packageKey = pack.zipPath || pack.path || pack.mediaBase;
                    const relatedRows = getRelatedRowsForPack(pack, rows);
                    const selectedVaultRow = relatedRows.find((row) => String(row.rowNumber) === String(pack?.selectedRowNumber || pack?.rowNumber))
                      || relatedRows[0]
                      || pack;
                    const vaultSectionKey = `${packageKey}-vault`;
                    const csvSectionKey = `${packageKey}-csv`;
                    const vaultOpen = Boolean(openResultSections[vaultSectionKey]);
                    const csvOpen = Boolean(openResultSections[csvSectionKey]);
                    const csvText = relatedRows.map((row) => makeListConfigCsvRow(row)).join("\n");
                    return (
                      <div className="lecture-tool-package" key={packageKey}>
                        <div className="lecture-tool-package__top">
                          <div>
                            <div className="lecture-tool-package__title">{pack.mediaFileName || pack.mediaBase}</div>
                            <div className="lecture-tool-package__sub">
                              {pack.lectureId || "-"} / {normalizeProductName(pack.product) || "-"} / 行 {pack.rowNumber || "-"}
                            </div>
                          </div>
                          <a className="lecture-tool-file__link" href={resultDownloadUrl(pack.zipUrl)}>
                            ZIPダウンロード
                          </a>
                        </div>
                        <div className="lecture-tool-result-statuses">
                          <span className="lecture-tool-status lecture-tool-status--ok">ZIP作成済み</span>
                          <span className="lecture-tool-status lecture-tool-status--ok">{pack.mediaBase}.html</span>
                          <span className="lecture-tool-status lecture-tool-status--ok">1.jpg {pack.imageInfo?.width || TARGET_WIDTH}px</span>
                          {pack.isRelatedProduct ? (
                            <span className="lecture-tool-status lecture-tool-status--warn">他製品版</span>
                          ) : null}
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
                                  同じ講演会ID・FileNameの行が {relatedRows.length} 件あります。Vault登録用に他製品版も表示しています。
                                </div>
                              ) : null}
                              <div className="stamp-tool-vault-list">
                                {relatedRows.map((row) => {
                                  const rowKey = getRowKey(row);
                                  const isSelectedVaultRow = String(row.rowNumber) === String(selectedVaultRow.rowNumber);
                                  const cells = [
                                    ["Detail Group", row.detailGroup, selectedVaultRow.detailGroup],
                                    ["製品", normalizeProductName(row.product), normalizeProductName(selectedVaultRow.product)],
                                    ["プレゼンテーション名", row.presentationName, selectedVaultRow.presentationName],
                                    ["PresentationId", row.presentationId, selectedVaultRow.presentationId],
                                  ];
                                  return (
                                    <div className="stamp-tool-vault-card" key={rowKey}>
                                      <div className="stamp-tool-vault-card__head">
                                        <strong>{normalizeProductName(row.product) || "製品未設定"}</strong>
                                        <span>行 {row.rowNumber || "-"}</span>
                                        {row.mediaFileName ? <span>{row.mediaFileName}</span> : null}
                                        {row.presentationId ? <span>{row.presentationId}</span> : null}
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
                            <em>{relatedRows.length} 行 / {csvOpen ? "閉じる" : "開く"}</em>
                          </button>
                          {csvOpen ? (
                            <div className="stamp-tool-csv-groups">
                              <div className="stamp-tool-csv-group">
                                <div className="stamp-tool-csv-group__head">
                                  <div>
                                    <strong>LIST_CONFIG.csv</strong>
                                    <span>{relatedRows.length} 行 / Date空欄は{formatTodayForCsv()}</span>
                                    <a className="stamp-tool-csv-tool-link" href={LECTURE_TOOL_CONFIG_SURVEY_URL} target="_blank" rel="noreferrer">
                                      LECTURE_TOOL_CONFIG_SURVEYを開く
                                    </a>
                                  </div>
                                  <button
                                    className="lecture-tool-button lecture-tool-button--small"
                                    type="button"
                                    onClick={() => copyText(`${packageKey}-csv`, csvText)}
                                  >
                                    {copiedKey === `${packageKey}-csv` ? "コピー済み" : relatedRows.length > 1 ? "このZIP分をコピー" : "CSV行コピー"}
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
                                      {relatedRows.map((row) => {
                                        const values = makeListConfigCsvValues(row);
                                        return (
                                          <tr key={getRowKey(row)}>
                                            {LIST_CONFIG_HEADERS.map((header) => (
                                              <td key={header} className={values[header] ? "" : "stamp-tool-csv-table__empty"}>
                                                {values[header] || "-"}
                                              </td>
                                            ))}
                                          </tr>
                                        );
                                      })}
                                    </tbody>
                                  </table>
                                </div>
                                {pack.thumbnailImages?.length ? (
                                  <>
                                    <div className="stamp-tool-result-links">
                                      {pack.thumbnailImages.map((image) => (
                                        <a key={image.path} href={resultDownloadUrl(image.url)} download={image.name}>
                                          {image.name}ダウンロード
                                        </a>
                                      ))}
                                    </div>
                                    <div className="stamp-tool-thumbnail-preview-list">
                                      {pack.thumbnailImages.map((image) => (
                                        <a
                                          className="stamp-tool-stampimg-preview"
                                          href={`${API_BASE}${image.url}`}
                                          target="_blank"
                                          rel="noreferrer"
                                          key={`${image.path}-preview`}
                                          aria-label={`${image.name}を開く`}
                                        >
                                          <img src={`${API_BASE}${image.url}`} alt={`${image.name}プレビュー`} />
                                        </a>
                                      ))}
                                    </div>
                                  </>
                                ) : (
                                  <div className="lecture-tool-empty">ThumnailNameが空のため、追加サムネイル画像はありません。</div>
                                )}
                              </div>
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
