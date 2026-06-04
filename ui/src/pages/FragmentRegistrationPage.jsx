import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

const API_BASE = import.meta.env.VITE_API_BASE || "";
const TARGET_WIDTH = 2048;
const FALLBACK_VAULT_ACCOUNTS = [
  "mika.hirawatari@msd.com",
  "maika.mori@msd.com",
  "yura.fukuhara@msd.com",
  "hidenori.sonohata@msd.com",
  "Hayato.Seto@vv-agency.com",
];

function formatSize(bytes) {
  if (bytes == null) return "";
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  const mb = kb / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MB`;
  return `${(mb / 1024).toFixed(2)} GB`;
}

function formatDate(value) {
  if (!value) return "";
  return new Date(value).toLocaleString("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function normalizeImageKey(value) {
  const name = String(value || "").split(/[\\/]/).pop() || "";
  const stem = name.includes(".") ? name.replace(/\.[^.]+$/, "") : name;
  return stem.replace(/\s|\u3000/g, "").toLowerCase();
}

function normalizeLookupKey(value) {
  return String(value || "").replace(/\s|\u3000/g, "").toLowerCase();
}

function makeId() {
  return globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
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

function resultFileUrl(sessionId, path) {
  const encodedPath = path.split("/").map(encodeURIComponent).join("/");
  return `${API_BASE}/lecture-tool/results/${sessionId}/${encodedPath}`;
}

function resultDownloadUrl(url) {
  if (!url) return "";
  const separator = url.includes("?") ? "&" : "?";
  return `${API_BASE}${url}${separator}download=1`;
}

function driveImageContentUrl(fileId) {
  return `${API_BASE}/fragment-tool/drive-images/${encodeURIComponent(fileId)}/content`;
}

function isProgressError(item) {
  const text = `${item?.step || ""} ${item?.message || ""}`;
  return /(error|fatal|cancelled|失敗|エラー|TimeoutError|Error:|停止しました|中断しました)/i.test(text);
}

function parseSseBlock(block) {
  const lines = block.split(/\r?\n/);
  let event = "message";
  const dataLines = [];
  for (const line of lines) {
    if (line.startsWith("event:")) event = line.slice(6).trim();
    if (line.startsWith("data:")) dataLines.push(line.slice(5).trimStart());
  }
  if (!dataLines.length) return null;
  return { event, data: JSON.parse(dataLines.join("\n")) };
}

function readImageSize(file) {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      resolve({ width: img.naturalWidth, height: img.naturalHeight, previewUrl: url });
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve({ width: 0, height: 0, previewUrl: "" });
    };
    img.src = url;
  });
}

function Field({ label, value }) {
  return (
    <div className="lecture-tool-field">
      <div className="lecture-tool-field__label">{label}</div>
      <div className="lecture-tool-field__value">{value || "-"}</div>
    </div>
  );
}

function getRowProcessMode(row) {
  const category = String(row?.category || "").trim();
  const compactCategory = category.replace(/\s|\u3000/g, "");
  const isRevision = compactCategory.includes("修正");
  return {
    isRevision,
    label: isRevision ? "修正" : "新規",
    category: category || "-",
    title: isRevision ? "修正として処理されます" : "新規として登録されます",
    note: isRevision
      ? "プレゼンテーションIDで既存バインダーを検索し、対象が1件だけ見つかった場合のみ下書き作成とZIP更新を行います。見つからない場合や複数件ある場合は停止します。"
      : "新規登録として処理します。プレゼンテーションIDの既存登録が見つかった場合は、重複防止のため処理を停止します。",
  };
}

function RowModeBadge({ row }) {
  const mode = getRowProcessMode(row);
  return (
    <span
      className={`lecture-tool-mode-badge${mode.isRevision ? " lecture-tool-mode-badge--revision" : " lecture-tool-mode-badge--new"}`}
      title={mode.category}
    >
      {mode.label}
    </span>
  );
}

function RowProcessNotice({ row }) {
  const mode = getRowProcessMode(row);
  return (
    <div className={`lecture-tool-process-note${mode.isRevision ? " lecture-tool-process-note--revision" : " lecture-tool-process-note--new"}`}>
      <div className="lecture-tool-process-note__title">
        <RowModeBadge row={row} />
        <span>{mode.title}</span>
      </div>
      <div className="lecture-tool-process-note__body">{mode.note}</div>
    </div>
  );
}

function CandidateRows({ rows, selectedRowId, onSelect, disabled }) {
  return (
    <div className="lecture-tool-candidates" aria-label="スプレッドシート行候補">
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
                <small>受付日</small>
                {row.receptionDate || "-"}
              </span>
              <span className="lecture-tool-candidate__cell">
                <small>区分</small>
                <RowModeBadge row={row} />
              </span>
              <span className="lecture-tool-candidate__cell lecture-tool-candidate__cell--wide">
                <small>開催日</small>
                {row.eventDate || "-"}
              </span>
              <span className="lecture-tool-candidate__cell">
                <small>時間</small>
                {row.eventTime || "-"}
              </span>
              <span className="lecture-tool-candidate__cell">
                <small>講演会名</small>
                {row.eventName || "-"}
              </span>


            </button>
          );
        })}
      </div>
    </div>
  );
}


function RowCellPreview({ row }) {
  return (
    <div className="lecture-tool-selected-cells-wrap">
      <div className="lecture-tool-selected-cells__head">セル確認用</div>
      <div className="lecture-tool-candidate lecture-tool-selected-cells">
        <span className="lecture-tool-candidate__cell lecture-tool-candidate__cell--strong">
          <small>行</small>
          <span className="lecture-tool-candidate__value">{row.rowNumber || "-"}</span>
        </span>
        <span className="lecture-tool-candidate__cell lecture-tool-candidate__cell--strong">
          <small>講演会ID</small>
          <span className="lecture-tool-candidate__value">{row.lectureId || "-"}</span>
        </span>
        <span className="lecture-tool-candidate__cell">
          <small>受付日</small>
          <span className="lecture-tool-candidate__value">{row.receptionDate || "-"}</span>
        </span>
        <span className="lecture-tool-candidate__cell">
          <small>区分</small>
          <RowModeBadge row={row} />
        </span>
        <span className="lecture-tool-candidate__cell lecture-tool-candidate__cell--wide">
          <small>開催日</small>
          <span className="lecture-tool-candidate__value">{row.eventDate || "-"}</span>
        </span>
        <span className="lecture-tool-candidate__cell">
          <small>時間</small>
          <span className="lecture-tool-candidate__value">{row.eventTime || "-"}</span>
        </span>
        <span className="lecture-tool-candidate__cell lecture-tool-candidate__cell--wide">
          <small>講演会名</small>
          <span className="lecture-tool-candidate__value">{row.eventName || "-"}</span>
        </span>
      </div>
    </div>
  );
}

function SelectedRowInfo({ row, showCellPreview = false }) {
  if (!row) {
    return <div className="lecture-tool-alert">スプレッドシートの行を指定してください。</div>;
  }



  return (
    <div>
      {showCellPreview ? <RowCellPreview row={row} /> : null}
      <RowProcessNotice row={row} />
      <div className="lecture-tool-hint mb_10">以下の内容で処理が実行されます。</div>
      <div className="lecture-tool-rowinfo">
        {/* <Field label="講演会ID" value={row.lectureId} /> */}
        <Field label="Product" value={row.product} />
        {/* <Field label="受付日" value={row.receptionDate} /> */}
        <Field label="区分" value={row.category} />
        <Field label="プレゼンテーションID" value={row.presentationId} />
        <Field label="メディアファイル名" value={row.mediaFileName} />
        <Field label="プレゼンテーション/キーメッセージ名" value={row.presentationName} />
        <Field label="画像名" value={row.imageName} />
      </div>
    </div>

  );
}

function RowNumberField({ item, row, onChange, disabled }) {
  const title = item.autoMatchCount === 1
    ? "画像名または講演会IDで一致した行番号を自動入力しています。"
    : item.autoMatchCount > 1
      ? "候補が複数あるため、スプレッドシート行番号を入力してください。"
      : "候補が見つからないため、スプレッドシート行番号を入力してください。";

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
          placeholder="例: 1234"
          disabled={disabled}
        />
      </label>
      {item.rowNumberInput && !row ? (
        <div className="lecture-tool-empty">該当行が読み込まれていません。行番号を確認してください。</div>
      ) : null}
    </div>
  );
}

function FileResultList({ sessionId, files }) {
  if (!files?.length) {
    return <div className="lecture-tool-empty">生成物はまだありません。</div>;
  }

  return (
    <div className="lecture-tool-list">
      {files.map((file) => (
        <div className="lecture-tool-file" key={file.path}>
          <div>
            <div className="lecture-tool-file__name" title={file.path}>
              {file.path}
            </div>
            <div className="lecture-tool-file__meta">
              {formatSize(file.size)} / {formatDate(file.modified)}
            </div>
          </div>
          <a className="lecture-tool-file__link" href={resultFileUrl(sessionId, file.path)} target="_blank" rel="noreferrer">
            開く
          </a>
        </div>
      ))}
    </div>
  );
}

export default function FragmentRegistrationPage() {
  const [rows, setRows] = useState([]);
  const [sheetMeta, setSheetMeta] = useState(null);
  const [images, setImages] = useState([]);
  const [driveFiles, setDriveFiles] = useState([]);
  const [driveMeta, setDriveMeta] = useState(null);
  const [driveSearchInput, setDriveSearchInput] = useState("");
  const [driveSearch, setDriveSearch] = useState("");
  const [driveLoading, setDriveLoading] = useState(false);
  const [driveError, setDriveError] = useState("");
  const [addingDriveIds, setAddingDriveIds] = useState(() => new Set());
  const [loadingRows, setLoadingRows] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [vaultAccounts, setVaultAccounts] = useState(FALLBACK_VAULT_ACCOUNTS);
  const [vaultAccount, setVaultAccount] = useState(FALLBACK_VAULT_ACCOUNTS[0]);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [progressItems, setProgressItems] = useState([]);
  const [currentProgress, setCurrentProgress] = useState("");
  const [previewImage, setPreviewImage] = useState(null);
  const [activeSessionId, setActiveSessionId] = useState("");
  const [cancelling, setCancelling] = useState(false);
  const imagesRef = useRef([]);
  const activeSessionIdRef = useRef("");
  const generatingRef = useRef(false);
  const cancelSentRef = useRef("");
  const streamAbortRef = useRef(null);
  const streamAbortReasonRef = useRef("");
  const mountedRef = useRef(true);

  const rowsById = useMemo(() => {
    const map = new Map();
    for (const row of rows) map.set(row.id, row);
    return map;
  }, [rows]);

  const selectedDriveIds = useMemo(() => (
    new Set(images.map((item) => item.driveFileId).filter(Boolean))
  ), [images]);

  const loadRows = useCallback(async () => {
    setLoadingRows(true);
    setError("");
    try {
      const response = await fetch(`${API_BASE}/lecture-tool/sheet-rows`, { cache: "no-store" });
      const payload = await response.json().catch(() => null);
      if (!response.ok) {
        throw new Error(getErrorMessage(payload, "スプレッドシートの読み込みに失敗しました。"));
      }
      setRows(payload.rows || []);
      setSheetMeta(payload);
      const nextVaultAccounts = payload.vaultAccounts?.length ? payload.vaultAccounts : FALLBACK_VAULT_ACCOUNTS;
      setVaultAccounts(nextVaultAccounts);
      setVaultAccount((current) => current && nextVaultAccounts.includes(current) ? current : nextVaultAccounts[0] || "");
    } catch (err) {
      setError(err?.message || "スプレッドシートの読み込みに失敗しました。");
    } finally {
      setLoadingRows(false);
    }
  }, []);

  useEffect(() => {
    loadRows();
  }, [loadRows]);

  const loadDriveFiles = useCallback(async ({ append = false, pageToken = "" } = {}) => {
    setDriveLoading(true);
    setDriveError("");
    try {
      const params = new URLSearchParams();
      params.set("pageSize", "80");
      if (driveSearch.trim()) params.set("search", driveSearch.trim());
      if (pageToken) params.set("pageToken", pageToken);
      const response = await fetch(`${API_BASE}/fragment-tool/drive-images?${params.toString()}`, { cache: "no-store" });
      const payload = await response.json().catch(() => null);
      if (!response.ok) {
        throw new Error(getErrorMessage(payload, "Drive画像の読み込みに失敗しました。"));
      }
      const files = payload?.files || [];
      setDriveFiles((prev) => {
        if (!append) return files;
        const seen = new Set(prev.map((file) => file.id));
        return [...prev, ...files.filter((file) => !seen.has(file.id))];
      });
      setDriveMeta(payload);
    } catch (err) {
      setDriveError(err?.message || "Drive画像の読み込みに失敗しました。");
    } finally {
      setDriveLoading(false);
    }
  }, [driveSearch]);

  useEffect(() => {
    loadDriveFiles();
  }, [loadDriveFiles]);

  useEffect(() => {
    imagesRef.current = images;
  }, [images]);

  useEffect(() => {
    generatingRef.current = generating;
  }, [generating]);

  useEffect(() => {
    return () => {
      for (const item of imagesRef.current) {
        if (item.previewUrl) URL.revokeObjectURL(item.previewUrl);
      }
    };
  }, []);

  const rememberSessionId = useCallback((sessionId) => {
    if (!sessionId) return;
    activeSessionIdRef.current = sessionId;
    setActiveSessionId(sessionId);
  }, []);

  const sendCancelRequest = useCallback((sessionId, { beacon = false } = {}) => {
    if (!sessionId || cancelSentRef.current === sessionId) return false;
    cancelSentRef.current = sessionId;
    const url = `${API_BASE}/lecture-tool/cancel/${sessionId}`;
    if (beacon && navigator.sendBeacon) {
      navigator.sendBeacon(url, new Blob([], { type: "text/plain" }));
      return true;
    }
    fetch(url, { method: "POST", keepalive: beacon }).catch(() => { });
    return true;
  }, []);

  const cancelOnLeave = useCallback(() => {
    if (!generatingRef.current) return;
    streamAbortReasonRef.current = "leave";
    sendCancelRequest(activeSessionIdRef.current, { beacon: true });
    streamAbortRef.current?.abort();
  }, [sendCancelRequest]);

  useEffect(() => {
    const handleBeforeUnload = (event) => {
      if (!generatingRef.current) return;
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    window.addEventListener("pagehide", cancelOnLeave);
    return () => {
      mountedRef.current = false;
      window.removeEventListener("beforeunload", handleBeforeUnload);
      window.removeEventListener("pagehide", cancelOnLeave);
      cancelOnLeave();
    };
  }, [cancelOnLeave]);

  const buildImageItem = useCallback(
    async (file, source = {}) => {
      const filename = source.filename || file.name;
      const key = normalizeImageKey(filename);
      const matchedRows = rows.filter((row) => (
        (row.imageKey && row.imageKey === key) ||
        (row.lectureId && normalizeLookupKey(row.lectureId) === key)
      ));
      const size = await readImageSize(file);
      return {
        id: makeId(),
        file,
        filename,
        size: file.size,
        width: size.width,
        height: size.height,
        previewUrl: size.previewUrl,
        sourceType: source.sourceType || "drive",
        driveFileId: source.driveFileId || "",
        driveWebViewLink: source.webViewLink || "",
        driveModifiedTime: source.modifiedTime || "",
        selectedRowId: matchedRows.length === 1 ? matchedRows[0].id : "",
        autoMatchCount: matchedRows.length,
        matchedRowIds: matchedRows.map((row) => row.id),
        rowNumberInput: "",
        rowNumberMode: matchedRows.length === 0,
        driveEnabled: true,
        vaultEnabled: true,
      };
    },
    [rows],
  );

  const addDriveImage = useCallback(async (driveFile) => {
    if (loadingRows || !rows.length || generating || !driveFile?.id || selectedDriveIds.has(driveFile.id)) return;
    setError("");
    setResult(null);
    setAddingDriveIds((prev) => new Set(prev).add(driveFile.id));
    try {
      const response = await fetch(driveImageContentUrl(driveFile.id), { cache: "no-store" });
      const contentType = response.headers.get("content-type") || driveFile.mimeType || "image/jpeg";
      if (!response.ok) {
        const payload = contentType.includes("json") ? await response.json().catch(() => null) : null;
        throw new Error(getErrorMessage(payload, `${driveFile.name || "Drive画像"} の取得に失敗しました。`));
      }
      const blob = await response.blob();
      const file = new File(
        [blob],
        driveFile.name || `drive-${driveFile.id}.jpg`,
        {
          type: blob.type || driveFile.mimeType || "image/jpeg",
          lastModified: Date.parse(driveFile.modifiedTime || "") || Date.now(),
        },
      );
      const item = await buildImageItem(file, {
        sourceType: "drive",
        driveFileId: driveFile.id,
        filename: driveFile.name,
        webViewLink: driveFile.webViewLink || "",
        modifiedTime: driveFile.modifiedTime || "",
      });
      setImages((prev) => (
        prev.some((existing) => existing.driveFileId === driveFile.id) ? prev : [...prev, item]
      ));
    } catch (err) {
      setError(err?.message || "Drive画像の追加に失敗しました。");
    } finally {
      setAddingDriveIds((prev) => {
        const next = new Set(prev);
        next.delete(driveFile.id);
        return next;
      });
    }
  }, [buildImageItem, generating, loadingRows, rows.length, selectedDriveIds]);

  const updateSelectedRow = (id, selectedRowId) => {
    setImages((prev) => prev.map((item) => (
      item.id === id ? { ...item, selectedRowId, rowNumberMode: false } : item
    )));
  };

  const updateRowNumber = (id, value) => {
    const normalized = value.replace(/[^\d]/g, "");
    setImages((prev) => prev.map((item) => (
      item.id === id ? { ...item, rowNumberInput: normalized, selectedRowId: "", rowNumberMode: true } : item
    )));
  };

  const updateRowMode = (id, rowNumberMode) => {
    setImages((prev) => prev.map((item) => (
      item.id === id ? { ...item, rowNumberMode } : item
    )));
  };

  const updateOperation = (id, field, checked) => {
    setImages((prev) => prev.map((item) => (
      item.id === id ? { ...item, [field]: checked } : item
    )));
  };

  const removeImage = (id) => {
    setImages((prev) => {
      const target = prev.find((item) => item.id === id);
      if (target?.previewUrl) URL.revokeObjectURL(target.previewUrl);
      if (previewImage?.id === id) setPreviewImage(null);
      return prev.filter((item) => item.id !== id);
    });
  };

  const getEffectiveRowId = (item) => item.rowNumberMode ? item.rowNumberInput : item.selectedRowId;
  const sheetReady = rows.length > 0 && !loadingRows;
  const rowMissing = images.filter((item) => !rowsById.has(getEffectiveRowId(item)));
  const hasVaultEnabled = images.some((item) => item.vaultEnabled);
  const canGenerate = sheetReady && images.length > 0 && rowMissing.length === 0 && (!hasVaultEnabled || vaultAccount) && !generating;

  const generate = async () => {
    if (!canGenerate) {
      if (hasVaultEnabled && !vaultAccount) {
        setError("Vault登録対象がある場合はVaultアカウントを選択してください。");
        return;
      }
      setError("すべての画像にスプレッドシート行を選択するか、行番号を入力してください。");
      return;
    }

    generatingRef.current = true;
    cancelSentRef.current = "";
    activeSessionIdRef.current = "";
    streamAbortRef.current?.abort();
    streamAbortReasonRef.current = "";
    const abortController = new AbortController();
    streamAbortRef.current = abortController;
    setGenerating(true);
    setError("");
    setResult(null);
    setProgressItems([]);
    setCurrentProgress("生成準備中...");
    setActiveSessionId("");
    setCancelling(false);

    try {
      const formData = new FormData();
      formData.append("vaultAccount", vaultAccount);
      for (const item of images) {
        formData.append("files", item.file, item.filename);
        formData.append("rowIds", getEffectiveRowId(item));
        formData.append("driveEnabled", item.driveEnabled ? "true" : "false");
        formData.append("vaultEnabled", item.vaultEnabled ? "true" : "false");
      }

      const response = await fetch(`${API_BASE}/lecture-tool/generate-stream`, {
        method: "POST",
        body: formData,
        signal: abortController.signal,
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => null);
        throw new Error(getErrorMessage(payload, "生成に失敗しました。"));
      }
      if (!response.body) {
        throw new Error("進捗ストリームを読み込めませんでした。");
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = "";
      let finalPayload = null;

      while (true) {
        const { value, done } = await reader.read();
        buffer += decoder.decode(value || new Uint8Array(), { stream: !done });
        const blocks = buffer.split(/\r?\n\r?\n/);
        buffer = blocks.pop() || "";

        for (const block of blocks) {
          const parsed = parseSseBlock(block);
          if (!parsed) continue;
          if (parsed.data?.sessionId) {
            rememberSessionId(parsed.data.sessionId);
          }
          if (parsed.event === "progress") {
            const item = {
              id: makeId(),
              at: new Date().toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit", second: "2-digit" }),
              step: parsed.data.step || "progress",
              message: parsed.data.message || "",
            };
            setCurrentProgress(item.message);
            setProgressItems((prev) => [item, ...prev].slice(0, 80));
          } else if (parsed.event === "done") {
            finalPayload = parsed.data;
            setResult(parsed.data);
            setCurrentProgress("完了しました。");
            generatingRef.current = false;
            setGenerating(false);
            setCancelling(false);
            setActiveSessionId("");
            activeSessionIdRef.current = "";
            setProgressItems((prev) => [
              {
                id: makeId(),
                at: new Date().toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit", second: "2-digit" }),
                step: "done",
                message: "すべての処理が完了しました。",
              },
              ...prev,
            ].slice(0, 80));
          } else if (parsed.event === "partial") {
            setResult(parsed.data.result);
            if (parsed.data.result?.sessionId) {
              rememberSessionId(parsed.data.result.sessionId);
            }
            const partialPackages = parsed.data.result?.packages || [];
            const latestPackage = partialPackages[partialPackages.length - 1];
            if (latestPackage) {
              setCurrentProgress(`ZIPを表示しました: ${latestPackage.mediaFileName}`);
            }
          } else if (parsed.event === "vault-result") {
            setResult((prev) => {
              if (!prev) return prev;
              const item = parsed.data.result || {};
              const nextRegistrations = item.error
                ? prev.vaultRegistrations || []
                : [...(prev.vaultRegistrations || []).filter((hit) => hit.mediaFileName !== item.mediaFileName), item];
              const nextErrors = item.error
                ? [...(prev.vaultErrors || []).filter((hit) => hit.mediaFileName !== item.mediaFileName), item]
                : (prev.vaultErrors || []).filter((hit) => hit.mediaFileName !== item.mediaFileName);
              return { ...prev, vaultRegistrations: nextRegistrations, vaultErrors: nextErrors };
            });
            const vaultResult = parsed.data.result || {};
            if (vaultResult?.error) {
              const message = `Vault登録エラー（スキップ）: ${vaultResult.mediaFileName || ""} ${vaultResult.error || ""}`.trim();
              setCurrentProgress(message);
              setProgressItems((prev) => [
                {
                  id: makeId(),
                  at: new Date().toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit", second: "2-digit" }),
                  step: "vault",
                  message,
                },
                ...prev,
              ].slice(0, 80));
            } else if (vaultResult?.mediaFileName) {
              setCurrentProgress(`Vault登録結果を表示しました: ${vaultResult.mediaFileName}`);
            }
          } else if (parsed.event === "cancelled") {
            setCurrentProgress(parsed.data.message || "処理を中断しました。");
            generatingRef.current = false;
            setGenerating(false);
            setCancelling(false);
            setProgressItems((prev) => [
              {
                id: makeId(),
                at: new Date().toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit", second: "2-digit" }),
                step: "cancelled",
                message: parsed.data.message || "処理を中断しました。",
              },
              ...prev,
            ].slice(0, 80));
          } else if (parsed.event === "error") {
            throw new Error(parsed.data.message || "生成に失敗しました。");
          }
        }

        if (done) break;
      }

      if (finalPayload) setResult(finalPayload);
    } catch (err) {
      if (err?.name === "AbortError") {
        if (mountedRef.current && streamAbortReasonRef.current !== "vault-error") {
          setCurrentProgress("処理を中断しました。");
        }
      } else if (mountedRef.current) {
        setError(err?.message || "生成に失敗しました。");
        setCurrentProgress("エラーで停止しました。");
      }
    } finally {
      generatingRef.current = false;
      if (streamAbortRef.current === abortController) {
        streamAbortRef.current = null;
      }
      if (mountedRef.current) {
        setGenerating(false);
        setCancelling(false);
      }
    }
  };

  const cancelGenerate = async () => {
    const sessionId = activeSessionIdRef.current || activeSessionId || result?.sessionId;
    if (!sessionId || cancelling) {
      streamAbortRef.current?.abort();
      generatingRef.current = false;
      setCurrentProgress("処理を中断しました。");
      setGenerating(false);
      setCancelling(false);
      return;
    }
    if (cancelSentRef.current === sessionId) {
      streamAbortRef.current?.abort();
      generatingRef.current = false;
      setCurrentProgress("処理を中断しました。");
      setGenerating(false);
      setCancelling(false);
      return;
    }
    cancelSentRef.current = sessionId;
    setCancelling(true);
    setCurrentProgress("中断リクエストを送信しています...");
    try {
      const response = await fetch(`${API_BASE}/lecture-tool/cancel/${sessionId}`, { method: "POST" });
      const payload = await response.json().catch(() => null);
      if (!response.ok) {
        throw new Error(getErrorMessage(payload, "中断リクエストに失敗しました。"));
      }
      setProgressItems((prev) => [
        {
          id: makeId(),
          at: new Date().toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit", second: "2-digit" }),
          step: "cancel",
          message: payload?.message || "中断リクエストを送信しました。",
        },
        ...prev,
      ].slice(0, 80));
      streamAbortReasonRef.current = "user-cancel";
      streamAbortRef.current?.abort();
      generatingRef.current = false;
      setGenerating(false);
      setCancelling(false);
    } catch (err) {
      const message = err?.message || "中断リクエストに失敗しました。";
      setError(message);
      if (/見つからない|すでに完了|404/.test(message)) {
        streamAbortRef.current?.abort();
        generatingRef.current = false;
        setGenerating(false);
      }
      setCancelling(false);
    }
  };

  return (
    <div className="lecture-tool-page">
      <div className="lecture-tool-page__inner">
        <header className="lecture-tool-header">
          <div>
            <h1 className="lecture-tool-header__title">Fragment登録ツール</h1>
            <div className="lecture-tool-header__sub">
              Drive画像とスプレッドシート行を紐づけ、HTML・サムネイル・ZIPを生成します。
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
            {generating ? (
              <button
                className="lecture-tool-button lecture-tool-button--danger"
                type="button"
                onClick={cancelGenerate}
                disabled={cancelling}
              >
                {cancelling ? "中断中" : "中断"}
              </button>
            ) : null}
          </div>
        </header>

        {error ? <div className="lecture-tool-alert">{error}</div> : null}
        <div className="lecture-tool-grid">
          <div className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">Drive画像選択</h2>
                  <div className="lecture-tool-panel__sub">
                    画像名が一致する行は自動で選択されます。ZIP内画像は幅{TARGET_WIDTH}pxで作成し、Drive格納用画像は別途リサイズで100KB前後に調整します。
                  </div>
                </div>
                <span className="lecture-tool-status lecture-tool-status--ok">
                  {images.length} 画像
                </span>
              </div>

              <div className="lecture-tool-drive-picker">
                <div className="lecture-tool-drive-picker__toolbar">
                  <label className="lecture-tool-select-label">
                    Drive内検索
                    <input
                      className="lecture-tool-search"
                      value={driveSearchInput}
                      onChange={(event) => setDriveSearchInput(event.target.value)}
                      onKeyDown={(event) => {
                        if (event.key === "Enter") {
                          event.preventDefault();
                          setDriveSearch(driveSearchInput.trim());
                        }
                      }}
                      placeholder="画像名で検索"
                      disabled={driveLoading || generating}
                    />
                  </label>
                  <div className="lecture-tool-drive-picker__actions">
                    <button
                      className="lecture-tool-button"
                      type="button"
                      onClick={() => setDriveSearch(driveSearchInput.trim())}
                      disabled={driveLoading || generating}
                    >
                      検索
                    </button>
                    <button
                      className="lecture-tool-button"
                      type="button"
                      onClick={() => loadDriveFiles()}
                      disabled={driveLoading || generating}
                    >
                      {driveLoading ? "読込中" : "再読込"}
                    </button>
                    {driveMeta?.folderUrl ? (
                      <a className="lecture-tool-file__link" href={driveMeta.folderUrl} target="_blank" rel="noreferrer">
                        Driveを開く
                      </a>
                    ) : null}
                  </div>
                </div>

                {driveError ? <div className="lecture-tool-alert">{driveError}</div> : null}

                <div className="lecture-tool-drive-picker__meta">
                  {driveLoading ? "Drive画像を読み込み中です。" : `${driveFiles.length} 件表示`}
                  {driveMeta?.recursive ? ` / サブフォルダ含む ${driveMeta.folderCount || 0} フォルダ` : ""}
                  {driveMeta?.foldersTruncated ? " / フォルダ数上限に達しました" : ""}
                  {!sheetReady ? " / 参照シート読み込み後に追加できます。" : ""}
                </div>

                {!driveLoading && !driveFiles.length ? (
                  <div className="lecture-tool-empty">Driveフォルダ内に選択できる画像がありません。</div>
                ) : (
                  <div className="lecture-tool-drive-grid">
                    {driveFiles.map((driveFile) => {
                      const selected = selectedDriveIds.has(driveFile.id);
                      const adding = addingDriveIds.has(driveFile.id);
                      const dimensions = driveFile.imageMediaMetadata?.width && driveFile.imageMediaMetadata?.height
                        ? `${driveFile.imageMediaMetadata.width} x ${driveFile.imageMediaMetadata.height}px`
                        : "";
                      return (
                        <div className={`lecture-tool-drive-card${selected ? " lecture-tool-drive-card--selected" : ""}`} key={driveFile.id}>
                          <div className="lecture-tool-drive-card__thumb">
                            {driveFile.thumbnailLink ? <img src={driveFile.thumbnailLink} alt="" loading="lazy" /> : <span>IMAGE</span>}
                          </div>
                          <div className="lecture-tool-drive-card__body">
                            <div className="lecture-tool-drive-card__name" title={driveFile.name}>{driveFile.name}</div>
                            <div className="lecture-tool-drive-card__meta">
                              {[driveFile.size ? formatSize(Number(driveFile.size)) : "", dimensions, formatDate(driveFile.modifiedTime)].filter(Boolean).join(" / ")}
                            </div>
                            <div className="lecture-tool-drive-card__actions">
                              <button
                                className="lecture-tool-button lecture-tool-button--small"
                                type="button"
                                onClick={() => addDriveImage(driveFile)}
                                disabled={!sheetReady || generating || selected || adding}
                              >
                                {selected ? "追加済み" : adding ? "追加中" : "追加"}
                              </button>
                              {driveFile.webViewLink ? (
                                <a className="lecture-tool-file__link" href={driveFile.webViewLink} target="_blank" rel="noreferrer">
                                  表示
                                </a>
                              ) : null}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}

                {driveMeta?.nextPageToken ? (
                  <button
                    className="lecture-tool-button"
                    type="button"
                    onClick={() => loadDriveFiles({ append: true, pageToken: driveMeta.nextPageToken })}
                    disabled={driveLoading || generating}
                  >
                    さらに読み込む
                  </button>
                ) : null}
              </div>
            </div>
          </div>

          <aside className="lecture-tool-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">参照シート</h2>
                  {/* <div className="lecture-tool-panel__sub">gid: {sheetMeta?.sheetGid || "-"}</div> */}
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
                  {loadingRows ? "読み込み中" : `${rows.length > 0 ? rows.length + 4 : 0} 件`}
                </div>
                {/* <div>
                  <strong>columns</strong>
                  <br />
                  講演会ID / Product / 画像名 / メディアファイル名 / プレゼンテーション
                </div> */}
              </div>
            </div>
          </aside>

          <aside className="lecture-tool-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">Vaultアカウント</h2>
                  <div className="lecture-tool-panel__sub">Vault登録に使用するアカウントを選択してください。</div>
                </div>
              </div>
              <div className="lecture-tool-settings">
                <label className="lecture-tool-select-label">

                  <select
                    className="lecture-tool-select"
                    value={vaultAccount}
                    onChange={(event) => setVaultAccount(event.target.value)}
                    disabled={generating}
                  >
                    {vaultAccounts.map((account) => (
                      <option value={account} key={account}>
                        {account}
                      </option>
                    ))}
                  </select>
                </label>
              </div>

              {vaultAccount === "Hayato.Seto@vv-agency.com" ? (
                <div className="lecture-tool-hint mt_10">嵐丸環境でのテスト実行になります。一部の処理は省略されます。</div>
              ) : null}


            </div>
          </aside>

          <div className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">行の選択</h2>
                  <div className="lecture-tool-panel__sub">誤選択防止のため、選択中の行情報を画像ごとに表示します。</div>
                </div>
              </div>

              {!images.length ? (
                <div className="lecture-tool-empty">Drive画像を追加すると、ここに選択欄が表示されます。</div>
              ) : (
                <div className="lecture-tool-image-list">
                  {images.map((item) => {
                    const manualRow = item.rowNumberMode ? rowsById.get(item.rowNumberInput) : null;
                    const selectedRow = item.rowNumberMode ? manualRow : rowsById.get(item.selectedRowId);
                    const isSmall = item.width > 0 && item.width < TARGET_WIDTH;
                    const optionRows = (item.matchedRowIds || [])
                      .map((rowId) => rowsById.get(rowId))
                      .filter(Boolean);
                    const hasCandidateRows = optionRows.length > 0;
                    const showRowSelect = hasCandidateRows && !item.rowNumberMode;
                    return (
                      <div className="lecture-tool-image-card" key={item.id}>
                        <div className="lecture-tool-image-card__media">
                          {item.previewUrl ? (
                            <button
                              className="lecture-tool-image-card__preview"
                              type="button"
                              onClick={() => setPreviewImage(item)}
                              aria-label={`${item.filename}を拡大表示`}
                            >
                              <img src={item.previewUrl} alt="" />
                            </button>
                          ) : null}
                        </div>
                        <div className="lecture-tool-image-card__body">
                          <div className="lecture-tool-image-card__top">
                            <div>
                              <div className="lecture-tool-image-card__title">{item.filename}</div>
                              <div className="lecture-tool-image-card__sub">
                                {formatSize(item.size)} / {item.width || "-"} x {item.height || "-"}px
                                {item.sourceType === "drive" ? " / Drive選択" : ""}
                              </div>
                            </div>
                            <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={() => removeImage(item.id)} disabled={generating}>
                              削除
                            </button>
                          </div>

                          <div className="lecture-tool-tag-row">
                            {item.autoMatchCount === 1 ? (
                              <span className="lecture-tool-status lecture-tool-status--ok">画像名で自動一致</span>
                            ) : item.autoMatchCount > 1 ? (
                              <span className="lecture-tool-status lecture-tool-status--warn">同名候補 {item.autoMatchCount} 件</span>
                            ) : (
                              <span className="lecture-tool-status lecture-tool-status--warn">手動選択</span>
                            )}
                            {isSmall ? (
                              <span className="lecture-tool-status lecture-tool-status--warn">幅{TARGET_WIDTH}pxへリサイズ</span>
                            ) : null}
                            {item.driveWebViewLink ? (
                              <a className="lecture-tool-file__link" href={item.driveWebViewLink} target="_blank" rel="noreferrer">
                                Drive画像
                              </a>
                            ) : null}
                          </div>
                          {item.autoMatchCount > 1 && !item.selectedRowId && !item.rowNumberMode ? (
                            <div className="lecture-tool-hint">候補が複数あります。該当するスプレッドシート行を選択してください。</div>
                          ) : null}

                          <div className="lecture-tool-operation-row">
                            <label className="lecture-tool-check">
                              <input
                                type="checkbox"
                                checked={item.driveEnabled}
                                onChange={(event) => updateOperation(item.id, "driveEnabled", event.target.checked)}
                                disabled={generating}
                              />
                              <span>Drive格納</span>
                            </label>
                            <label className="lecture-tool-check">
                              <input
                                type="checkbox"
                                checked={item.vaultEnabled}
                                onChange={(event) => updateOperation(item.id, "vaultEnabled", event.target.checked)}
                                disabled={generating}
                              />
                              <span>Vault登録</span>
                            </label>
                          </div>

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

                          {showRowSelect ? (
                            <CandidateRows
                              rows={optionRows}
                              selectedRowId={item.selectedRowId}
                              onSelect={(rowId) => updateSelectedRow(item.id, rowId)}
                              disabled={loadingRows || generating}
                            />
                          ) : null}





                          {!showRowSelect ? (
                            <RowNumberField item={item} row={manualRow} onChange={updateRowNumber} hasCandidateRows={hasCandidateRows} disabled={generating} />
                          ) : null}






                          <SelectedRowInfo row={selectedRow} showCellPreview={item.rowNumberMode} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>

          {(generating || progressItems.length > 0) ? (
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
                  {progressItems.length ? (
                    <div className="lecture-tool-progress__list">
                      {progressItems.map((progressItem) => (
                        <div
                          className={`lecture-tool-progress__item${isProgressError(progressItem) ? " lecture-tool-progress__item--error" : ""}`}
                          key={progressItem.id}
                        >
                          <span>{progressItem.at}</span>
                          <strong>{progressItem.step}</strong>
                          <em>{progressItem.message}</em>
                        </div>
                      ))}
                    </div>
                  ) : null}
                </div>
              </div>
            </div>
          ) : null}

          <div className="lecture-tool-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">生成結果</h2>
                  {/* <div className="lecture-tool-panel__sub">
                    {result?.sessionId ? `session: ${result.sessionId}` : "ZIP とHTMLを確認できます。"}
                  </div> */}
                </div>
              </div>

              {result?.packages?.length ? (
                <>
                  <div className="lecture-tool-package-list">
                    {result.packages.map((pack) => {
                      const driveItems = (result.driveUploads || []).filter((item) => item.mediaFileName === pack.mediaFileName);
                      const driveErrors = (result.driveErrors || []).filter((item) => item.mediaFileName === pack.mediaFileName);
                      const driveUploadedCount = driveItems.filter((item) => item.uploaded || item.webViewLink).length;
                      const vaultHit = (result.vaultRegistrations || []).find((item) => item.zipPath === pack.zipPath || item.mediaFileName === pack.mediaFileName);
                      const vaultError = (result.vaultErrors || []).find((item) => item.zipPath === pack.zipPath || item.mediaFileName === pack.mediaFileName);
                      return (
                        <div className="lecture-tool-package" key={pack.zipPath}>
                          <div className="lecture-tool-package__top">
                            <div>
                              <div className="lecture-tool-package__title">{pack.presentationName}</div>

                              <div className="lecture-tool-package__sub">{pack.presentationId}</div>
                            </div>
                            <a className="lecture-tool-file__link" href={`${API_BASE}${pack.zipUrl}`} target="_blank" rel="noreferrer">
                              ZIPダウンロード
                            </a>
                          </div>
                          <div className="lecture-tool-result-statuses">
                            <span className="lecture-tool-status lecture-tool-status--ok">ZIP作成済み</span>
                            {pack.driveEnabled ? (
                              <span className={`lecture-tool-status ${driveUploadedCount ? "lecture-tool-status--ok" : "lecture-tool-status--warn"}`}>
                                Drive {driveUploadedCount ? `${driveUploadedCount} 件完了` : "待機中"}
                              </span>
                            ) : null}
                            {pack.vaultEnabled ? (
                              <span className={`lecture-tool-status ${vaultHit ? "lecture-tool-status--ok" : vaultError ? "lecture-tool-status--warn" : "lecture-tool-status--warn"}`}>
                                Vault {vaultHit ? "登録済み" : vaultError ? "エラー" : "待機中"}
                              </span>
                            ) : null}
                          </div>
                          {vaultHit ? (
                            <div className="lecture-tool-result-detail">
                              <div className="lecture-tool-result-detail__label">Vault登録</div>
                              {vaultHit.url ? (
                                <div>Binder URL: <a href={vaultHit.url} target="_blank" rel="noreferrer">{vaultHit.url}</a></div>
                              ) : null}
                              {vaultHit.slideUrl ? (
                                <div>Slide URL: <a href={vaultHit.slideUrl} target="_blank" rel="noreferrer">{vaultHit.slideUrl}</a></div>
                              ) : null}
                            </div>
                          ) : null}
                          {vaultError ? (
                            <div className="lecture-tool-result-detail lecture-tool-result-detail--error">
                              <div className="lecture-tool-result-detail__label">Vault登録エラー</div>
                              <div>{vaultError.error || "詳細不明のエラーです。"}</div>
                            </div>
                          ) : null}
                          {(driveItems.length || driveErrors.length || result.driveFolderUrl) ? (
                            <div className="lecture-tool-result-detail">
                              <div className="lecture-tool-result-detail__label">Drive用画像 / Google Drive</div>
                              {result.driveFolderUrl ? (
                                <div>フォルダ: <a href={result.driveFolderUrl} target="_blank" rel="noreferrer">{result.driveFolderUrl}</a></div>
                              ) : null}
                              {driveItems.map((driveItem) => (
                                <div key={`${pack.zipPath}-${driveItem.filename}`}>
                                  {driveItem.filename}（{driveItem.uploaded || driveItem.webViewLink ? "アップロード済み" : "画像生成済み"}）
                                  {driveItem.downloadUrl ? (
                                    <> / <a href={resultDownloadUrl(driveItem.downloadUrl)} download={driveItem.filename || true}>Drive用画像をダウンロード</a></>
                                  ) : null}
                                  {driveItem.webViewLink ? (
                                    <> / <a href={driveItem.webViewLink} target="_blank" rel="noreferrer">Drive上の画像</a></>
                                  ) : null}
                                  {driveItem.uploadSkipReason ? (
                                    <> / {driveItem.uploadSkipReason}</>
                                  ) : null}
                                </div>
                              ))}
                              {driveErrors.map((driveError, index) => (
                                <div className="lecture-tool-result-detail__error" key={`${pack.zipPath}-drive-error-${index}`}>
                                  {driveError.filename || "Drive"}: {driveError.error || "詳細不明のエラーです。"}
                                  {driveError.downloadUrl ? (
                                    <> / <a href={resultDownloadUrl(driveError.downloadUrl)} download={driveError.filename || true}>Drive用画像をダウンロード</a></>
                                  ) : null}
                                </div>
                              ))}
                            </div>
                          ) : null}
                        </div>
                      );
                    })}
                  </div>
                </>
              ) : (
                <div className="lecture-tool-empty">まだ生成されていません。</div>
              )}
            </div>
          </div>

          {/* <div className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">生成ファイル一覧</h2>
                  <div className="lecture-tool-panel__sub">HTML、リサイズ画像、ZIPを個別に開けます。</div>
                </div>
              </div>
              <FileResultList sessionId={result?.sessionId} files={result?.resultFiles || []} />
            </div>
          </div> */}
        </div>
      </div>
      {previewImage ? (
        <div className="lecture-tool-lightbox" role="dialog" aria-modal="true" aria-label="画像プレビュー" onClick={() => setPreviewImage(null)}>
          <div className="lecture-tool-lightbox__body" onClick={(event) => event.stopPropagation()}>
            <div className="lecture-tool-lightbox__head">
              <div className="lecture-tool-lightbox__title">{previewImage.filename}</div>
              <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={() => setPreviewImage(null)}>
                閉じる
              </button>
            </div>
            <img className="lecture-tool-lightbox__image" src={previewImage.previewUrl} alt="" />
          </div>
        </div>
      ) : null}
    </div>
  );
}
