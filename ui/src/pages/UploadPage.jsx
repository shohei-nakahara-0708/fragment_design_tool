// UploadPage.jsx
import React, { useEffect, useMemo, useRef, useState, useCallback } from "react";
import { Link, useNavigate } from "react-router-dom";

const ERROR_MESSAGES = {
  event_id_not_found: "システムIDがスプレッドシートに存在しません",
  event_id_required: "event_id が未入力です",
  not_supported_file: "サポートされていないファイル形式です（pptx または pdf を指定してください）",
  batch_failed: "1件以上のファイルでエラーが発生しました",
  unknown_error: "不明なエラーが発生しました",
};

function humanizeError(err) {
  if (!err) return "";
  if (typeof err === "string") return ERROR_MESSAGES[err] || err;
  if (typeof err === "object") {
    const code = err.code || err.error;
    return ERROR_MESSAGES[code] || err.message || JSON.stringify(err);
  }
  return String(err);
}

function parseSseChunk(buffer) {
  // buffer: string
  // returns { events: [{event, data}], rest }
  const events = [];
  const parts = buffer.split("\n\n");
  const rest = parts.pop() ?? "";

  for (const block of parts) {
    let event = "message";
    let dataStr = "";
    for (const line of block.split("\n")) {
      if (line.startsWith("event:")) event = line.slice(6).trim();
      if (line.startsWith("data:")) dataStr += line.slice(5).trim();
    }
    if (dataStr) {
      try {
        events.push({ event, data: JSON.parse(dataStr) });
      } catch {
        events.push({ event, data: dataStr });
      }
    }
  }
  return { events, rest };
}

function fmtSize(bytes) {
  if (bytes == null) return "";
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  const mb = kb / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MB`;
  const gb = mb / 1024;
  return `${gb.toFixed(2)} GB`;
}

function makeSessionId() {
  return crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}
function normalizeEventId(s) {
  return (s ?? "").trim();
}
function toIsoStart(dateStr) {
  if (!dateStr) return "";
  return new Date(`${dateStr}T00:00:00`).toISOString();
}
function toIsoEnd(dateStr) {
  if (!dateStr) return "";
  const d = new Date(`${dateStr}T00:00:00`);
  d.setHours(23, 59, 59, 999);
  return d.toISOString();
}

function StatusPill({ status }) {
  const v =
    status === "ok"
      ? { bg: "rgba(34,197,94,0.12)", bd: "rgba(34,197,94,0.25)", fg: "#166534", label: "OK" }
      : status === "error"
      ? { bg: "rgba(239,68,68,0.12)", bd: "rgba(239,68,68,0.25)", fg: "#991b1b", label: "ERROR" }
      : { bg: "rgba(148,163,184,0.18)", bd: "rgba(148,163,184,0.28)", fg: "#334155", label: "PENDING" };

  return (
    <span
      style={{
        fontSize: 11,
        fontWeight: 900,
        padding: "5px 10px",
        borderRadius: 999,
        border: `1px solid ${v.bd}`,
        background: v.bg,
        color: v.fg,
        whiteSpace: "nowrap",
      }}
    >
      {v.label}
    </span>
  );
}

const ui = {
  page: {
    minHeight: "100vh",
    background: "#f6f7f9",
    color: "#0f172a",
    fontFamily:
      '"Invention JP", ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial',
  },
  container: { maxWidth: 1100, margin: "0 auto", padding: "22px 18px 26px" },

  header: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, marginBottom: 12 },
  h1: { margin: 0, fontSize: 22, fontWeight: 900, letterSpacing: "-0.02em" },
  sub: { marginTop: 6, fontSize: 13, color: "#475569", lineHeight: 1.4 },
  headerRight: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },

  btnPrimary: {
    background: "#2563eb",
    color: "#fff",
    padding: "10px 14px",
    borderRadius: 12,
    border: "1px solid rgba(37,99,235,0.35)",
    fontWeight: 900,
    fontSize: 13,
    cursor: "pointer",
    boxShadow: "0 10px 22px rgba(37,99,235,0.18)",
  },
  btnGhost: {
    background: "rgba(255,255,255,0.85)",
    color: "#0f172a",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid rgba(226,232,240,0.95)",
    fontWeight: 900,
    fontSize: 13,
    cursor: "pointer",
  },
  btnDanger: {
    background: "rgba(239,68,68,0.10)",
    color: "#991b1b",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid rgba(239,68,68,0.25)",
    fontWeight: 950,
    fontSize: 13,
    cursor: "pointer",
  },

  card: {
    background: "rgba(255,255,255,0.9)",
    border: "1px solid rgba(226,232,240,0.95)",
    borderRadius: 16,
    boxShadow: "0 12px 30px rgba(2,6,23,0.06)",
    overflow: "hidden",
  },
  cardInner: { padding: 14 },

  drop: (dragOver) => ({
    border: "2px dashed " + (dragOver ? "rgba(37,99,235,0.65)" : "rgba(148,163,184,0.55)"),
    borderRadius: 16,
    padding: 16,
    background: dragOver ? "rgba(37,99,235,0.05)" : "rgba(248,250,252,0.75)",
    transition: "all .12s ease",
  }),
  row: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" },
  hint: { marginTop: 10, fontSize: 12, color: "#64748b", lineHeight: 1.5 },

  alertErr: {
    marginTop: 12,
    padding: "10px 12px",
    borderRadius: 14,
    background: "#fff4f4",
    border: "1px solid #ffd0d0",
    color: "#a40000",
    whiteSpace: "pre-wrap",
  },
  alertInfo: {
    marginTop: 10,
    padding: "10px 12px",
    borderRadius: 14,
    background: "rgba(37,99,235,0.06)",
    border: "1px solid rgba(37,99,235,0.18)",
    color: "#1e3a8a",
    whiteSpace: "pre-wrap",
    fontSize: 13,
  },

  listHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, marginTop: 14 },
  listMeta: { fontSize: 13, color: "#64748b" },

  item: (status) => ({
    display: "grid",
    gridTemplateColumns: "44px 1fr 240px 120px 120px 86px",
    gap: 10,
    alignItems: "center",
    padding: "12px 12px",
    borderTop: "1px solid rgba(226,232,240,0.75)",
    background:
      status === "error" ? "rgba(239,68,68,0.04)" : status === "ok" ? "rgba(34,197,94,0.04)" : "#fff",
  }),
  input: (invalid, disabled) => ({
    width: "100%",
    padding: "10px 10px",
    borderRadius: 12,
    border: "1px solid " + (invalid ? "rgba(239,68,68,0.35)" : "rgba(226,232,240,0.95)"),
    background: disabled ? "#f1f5f9" : "#fff",
    outline: "none",
    fontSize: 13,
    fontWeight: 700,
    maxWidth: 210,
  }),
  fileTitle: { fontWeight: 900, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  fileSub: { marginTop: 4, fontSize: 12, color: "#64748b", lineHeight: 1.4 },
  msgErr: { marginTop: 6, fontSize: 12, color: "#991b1b", whiteSpace: "pre-wrap" },
   tabs: {
    display: "inline-flex",
    gap: 6,
    padding: 6,
    borderRadius: 14,
    border: "1px solid rgba(226,232,240,0.95)",
    background: "rgba(255,255,255,0.85)",
  },
  tabBtn: (active) => ({
    padding: "9px 12px",
    borderRadius: 12,
    border: "1px solid " + (active ? "rgba(37,99,235,0.35)" : "transparent"),
    background: active ? "rgba(37,99,235,0.10)" : "transparent",
    color: active ? "#1d4ed8" : "#0f172a",
    fontWeight: 950,
    fontSize: 13,
    cursor: "pointer",
    whiteSpace: "nowrap",
  }),
  headerLeft: { display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" },
  headerTitleBlock: { minWidth: 220 },

  divider: { height: 1, background: "rgba(226,232,240,0.85)", margin: "12px 0" },

  btnSmall: {
    background: "rgba(255,255,255,0.85)",
    color: "#0f172a",
    padding: "8px 10px",
    borderRadius: 12,
    border: "1px solid rgba(226,232,240,0.95)",
    fontWeight: 900,
    fontSize: 12,
    cursor: "pointer",
  },

  progressCard: {
    marginTop: 10,
    background: "#fff",
    border: "1px solid rgba(226,232,240,0.95)",
    borderRadius: 12,
    padding: 10,
  },

  tableCard: {
    background: "rgba(255,255,255,0.9)",
    border: "1px solid rgba(226,232,240,0.95)",
    borderRadius: 14,
    boxShadow: "0 12px 30px rgba(2,6,23,0.06)",
    overflow: "hidden",
  },

  restoreHead: {
  display: "grid",
  gridTemplateColumns: "44px 1fr 140px 120px 220px",
  gap: 10,
  background: "#f8fafc",
  borderBottom: "1px solid rgba(226,232,240,0.85)",
  fontSize: 12,
  color: "#64748b",
  padding: "10px 12px",
  fontWeight: 900,
},

restoreRow: (tone) => ({
  display: "grid",
  gridTemplateColumns: "44px 1fr 140px 120px 220px",
  gap: 10,
  alignItems: "center",
  padding: "12px 12px",
  borderTop: "1px solid rgba(226,232,240,0.75)",
  background:
    tone === "error"
      ? "rgba(239,68,68,0.04)"
      : tone === "ok"
      ? "rgba(34,197,94,0.04)"
      : "#fff",
  transition: "background .15s ease",
  }),

  codeChip: {
  display: "inline-block",
  padding: "3px 8px",
  borderRadius: 999,
  border: "1px solid rgba(226,232,240,0.95)",
  background: "rgba(248,250,252,0.9)",
  fontSize: 12,
  fontWeight: 900,
  color: "#0f172a",
  },
  
  dropRestore: (dragOver) => ({
  border: "2px dashed " + (dragOver ? "rgba(37,99,235,0.7)" : "rgba(148,163,184,0.55)"),
  borderRadius: 16,
  padding: 18,
  background: dragOver ? "rgba(37,99,235,0.06)" : "rgba(248,250,252,0.75)",
  transition: "all .15s ease",
  }),
  
  progressCard: {
  marginTop: 10,
  background: "#fff",
  border: "1px solid rgba(226,232,240,0.95)",
  borderRadius: 12,
  padding: 12,
  boxShadow: "0 6px 16px rgba(2,6,23,0.04)",
},
};

export default function UploadPage() {
  const nav = useNavigate();
  const [sessionId] = useState(() => makeSessionId());

  const [rows, setRows] = useState([]); // { id, file, filename, size, eventId, status, message, jobId, previewUrl }
  const [dragOver, setDragOver] = useState(false);

  const [uploading, setUploading] = useState(false);
  const [progressText, setProgressText] = useState("");
  const [globalError, setGlobalError] = useState("");

  const inputRef = useRef(null);

const [tab, setTab] = useState("upload"); // "upload" | "restore"
const [restoring, setRestoring] = useState(false);
const [restoreError, setRestoreError] = useState("");
const restoreInputRef = useRef(null);

const [restoreFiles, setRestoreFiles] = useState([]); // File[]
  const [restoreRows, setRestoreRows] = useState([]);   // 結果の配列
  const [restoreDragOver, setRestoreDragOver] = useState(false);

  useEffect(() => {
    const handler = (e) => {
      if (!uploading) return;
      e.preventDefault();
      e.returnValue = "";
      return "";
    };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [uploading]);

  const addFiles = (fileList) => {
    const files = Array.from(fileList || []);
    if (!files.length) return;

    const next = files
      .filter((f) => /\.pptx$/i.test(f.name) || /\.pdf$/i.test(f.name))
      .map((file) => ({
        id: crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        file,
        filename: file.name,
        size: file.size,
        eventId: "",
        status: "pending",
        message: "",
      }));

    if (next.length !== files.length) setGlobalError("※ .pptx / .pdf 以外は無視しました");
    else setGlobalError("");

    setRows((prev) => [...prev, ...next]);
    if (inputRef.current) inputRef.current.value = "";
  };

  const addRestoreFiles = (fileList) => {
  const files = Array.from(fileList || []);
  if (!files.length) return;

  const jsons = files.filter((f) => /\.json$/i.test(f.name));
  if (jsons.length !== files.length) {
    setRestoreError("※ .json 以外は無視しました");
  } else {
    setRestoreError("");
  }

  setRestoreFiles((prev) => {
    const mp = new Map(prev.map((f) => [f.name + ":" + f.size, f]));
    for (const f of jsons) mp.set(f.name + ":" + f.size, f); // 同名同サイズは重複排除
    return Array.from(mp.values());
  });

  if (restoreInputRef.current) restoreInputRef.current.value = "";
};

  const removeRow = (id) => {
    if (uploading) return;
    setRows((prev) => prev.filter((r) => r.id !== id));
  };

  const clearAll = () => {
    if (uploading) return;
    setRows([]);
    setGlobalError("");
  };

  const setEventId = (id, v) => {
    setRows((prev) =>
      prev.map((r) => (r.id === id ? { ...r, eventId: v, status: r.status === "error" ? "pending" : r.status } : r))
    );
  };

  const validate = useMemo(() => {
    if (!rows.length) return { ok: false, msg: "ファイルを追加してください" };
    const empty = rows.filter((r) => !normalizeEventId(r.eventId));
    if (empty.length) return { ok: false, msg: `event_id 未入力が ${empty.length} 件あります` };
    return { ok: true, msg: "" };
  }, [rows]);

  const onDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    addFiles(e.dataTransfer.files);
  };

 const uploadBatch = async () => {
  setGlobalError("");
  if (!validate.ok) {
    setGlobalError(validate.msg);
    return;
  }
   if (uploading) return;
   



  setUploading(true);
  setProgressText("準備中…");

  try {
    setRows((prev) => prev.map((r) => ({ ...r, status: "pending", message: "" })));

    const fd = new FormData();
    // fd.append("sessionId", sessionId); // サーバで受けるなら追加してOK
    rows.forEach((r) => fd.append("files", r.file, r.filename));
    rows.forEach((r) => fd.append("eventIds", normalizeEventId(r.eventId)));

    const API_BASE = import.meta.env.VITE_API_BASE || "";



    const r = await fetch(`${API_BASE}/upload/batch/stream`, { method: "POST", body: fd });
    if (!r.ok || !r.body) throw new Error(await r.text());

    const reader = r.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buf = "";

    let total = rows.length;
    let doneCount = 0;
    let returnedSessionId = null;

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;

      buf += decoder.decode(value, { stream: true });
      const parsed = parseSseChunk(buf);
      buf = parsed.rest;

      for (const ev of parsed.events) {
        if (ev.event === "start") {
          total = ev.data.total ?? total;
          returnedSessionId = ev.data.sessionId ?? returnedSessionId;
          setProgressText(`開始（${total}件）`);
        }

        if (ev.event === "phase") {
          setProgressText(ev.data.message || ev.data.phase || "処理中…");
        }

        if (ev.event === "item_start") {
          const i = ev.data.index;
          // 行の表示を「処理中」に寄せたいなら message に入れる
          setRows((prev) => {
            const out = [...prev];
            if (out[i]) out[i] = { ...out[i], status: "pending", message: "" };
            return out;
          });
        }

        if (ev.event === "item_done") {
          const i = ev.data.index;
          doneCount += 1;

          setRows((prev) => {
            const out = [...prev];
            if (!out[i]) return out;
            out[i] = {
              ...out[i],
              status: ev.data.ok ? "ok" : "error",
              jobId: ev.data.jobId || out[i].jobId || "",
              message: ev.data.ok ? "" : humanizeError(ev.data.error || ev.data.message),
            };
            return out;
          });

          setProgressText(`処理中… ${doneCount}/${total}`);
        }

        if (ev.event === "fatal") {
          throw new Error(ev.data?.message || "server fatal");
        }

        if (ev.event === "done") {
          returnedSessionId = ev.data.sessionId ?? returnedSessionId;
          setProgressText("完了");
          // 少し待って一覧へ
          setTimeout(() => {
            nav(`/jobs?session=${encodeURIComponent(returnedSessionId || "")}`);
          }, 250);
        }
      }
    }
  } catch (e) {
    console.error(e);
    setGlobalError(String(e?.message || e));
    setProgressText("");
  } finally {
    setUploading(false);
  }
 };
  
  
const restoreBatchFromJson = async () => {
  setRestoreError("");
  setRestoreRows([]);
  if (!restoreFiles.length) return;
  if (uploading || restoring) return;

  let returnedSessionId = null;

  setRestoring(true);
  try {
    const API_BASE = import.meta.env.VITE_API_BASE || "";
    const fd = new FormData();
    restoreFiles.forEach((f) => fd.append("files", f, f.name));

    const r = await fetch(`${API_BASE}/jobs/restore/batch`, { method: "POST", body: fd });
    if (!r.ok) {
      const t = await r.text().catch(() => "");
      throw new Error(`restore failed: ${r.status}\n${t}`);
    }
    const data = await r.json();
    setRestoreRows(data.results || []);

    returnedSessionId = data.sessionId ?? returnedSessionId;
    setTimeout(() => {
            nav(`/jobs?session=${encodeURIComponent(returnedSessionId || "")}`);
    }, 250);
    
  } catch (e) {
    setRestoreError(String(e?.message || e));
  } finally {
    setRestoring(false);
  }
};
  
  const restoreItems = useMemo(() => {
  // filename -> result
  const mp = new Map((restoreRows || []).map((r) => [String(r.filename || ""), r]));
  return (restoreFiles || []).map((f) => {
    const r = mp.get(String(f.name)) || null;
    const status = r ? (r.ok ? "ok" : "error") : (restoring ? "pending" : "pending");
    return {
      key: f.name + ":" + f.size,
      file: f,
      filename: f.name,
      size: f.size,
      status,
      result: r, // {ok, jobId, eventId, previewUrl, error, ...} or null
    };
  });
}, [restoreFiles, restoreRows, restoring]);


const progress = useMemo(() => {
  if (!rows.length) return 0;
  const done = rows.filter((r) => r.status === "ok" || r.status === "error").length;
  return Math.round((done / rows.length) * 100);
}, [rows]);

  const hasFiles = rows.length > 0;

  return (
    <div style={ui.page}>
      <div style={ui.container}>
        {/* Header */}
    <div style={ui.header}>
  {/* Left */}
  <div style={ui.headerLeft}>
            <div style={ui.headerTitleBlock}>
              <div style={{marginBottom:5}}>
                              <Link to="/" style={{ textDecoration: "none" }}>
                                ← 一覧へ
                              </Link>
                            </div>
      <h1 style={ui.h1}>Upload</h1>
      <div style={ui.sub}>
        {tab === "upload"
          ? "複数PPTX / PDFをまとめてアップロード → event_id を付与 → 生成"
          : "backup用json（複数可）をアップロード → 復元して再生成"}
      </div>
    </div>

  
  </div>

  {/* Right */}
          <div style={ui.headerRight}>
            
            

            
              <div style={ui.tabs}>
      <button
        style={ui.tabBtn(tab === "upload")}
        onClick={() => setTab("upload")}
        disabled={uploading || restoring}
      >
        アップロード
      </button>
      <button
        style={ui.tabBtn(tab === "restore")}
        onClick={() => setTab("restore")}
        disabled={uploading || restoring}
      >
        JSONから復元
      </button>
    </div>

    {tab === "upload" ? (
      <>
        <button style={ui.btnDanger} onClick={clearAll} disabled={!hasFiles || uploading}>
          全削除
        </button>
        <button style={ui.btnPrimary} onClick={uploadBatch} disabled={!hasFiles || uploading || !validate.ok}>
          {uploading ? "アップロード中…" : "アップロードして生成"}
        </button>
      </>
    ) : (
      <>
        <button
          style={ui.btnDanger}
          onClick={() => {
            if (restoring || uploading) return;
            setRestoreFiles([]);
            setRestoreRows([]);
            setRestoreError("");
            if (restoreInputRef.current) restoreInputRef.current.value = "";
          }}
          disabled={restoring || uploading}
        >
          全削除
        </button>
        <button
          style={ui.btnPrimary}
          onClick={restoreBatchFromJson}
          disabled={!restoreFiles.length || restoring || uploading}
        >
          {restoring ? "復元中…" : "復元して再生成"}
        </button>
      </>
    )}
  </div>
</div>

         {tab === "upload" && (
          <>
       

        {/* Dropzone card */}
        <div style={ui.card}>
          <div style={ui.cardInner}>
            <div
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              style={ui.drop(dragOver)}
            >
              <div style={ui.row}>
                <div style={{ fontWeight: 950 }}>
                  ここに <span style={{ color: "#2563eb" }}>.pptx / .pdf</span> をドロップ
                </div>
                <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                  <input
                    ref={inputRef}
                    type="file"
                    multiple
                    accept=".pptx,.pdf"
                    disabled={uploading}
                    style={{display: "none"}}
                    onChange={(e) => addFiles(e.target.files)}
                  />
                  <button style={ui.btnGhost} onClick={() => inputRef.current?.click?.()} disabled={uploading}>
                    ファイル選択
                  </button>
                </div>
              </div>

              <div style={ui.hint}>
                セッションID: <code>{sessionId.slice(0, 8)}…</code>（同時アップロードをまとめて表示する用）
              </div>
            </div>

{uploading && (
  <div style={ui.progressCard}>
    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#475569", fontWeight: 900 }}>
      <span>⏳ {progressText || "処理中…"}</span>
      <span>{progress}%</span>
    </div>
    <div style={{ marginTop: 8, height: 10, background: "rgba(226,232,240,0.9)", borderRadius: 999, overflow: "hidden" }}>
      <div style={{ width: `${progress}%`, height: "100%", background: "#2563eb", transition: "width 200ms ease" }} />
    </div>
  </div>
)}

            {!!globalError && <div style={ui.alertErr}>{globalError}</div>}
          </div>
        </div>

        {/* List header */}
        <div style={ui.listHead}>
          <div style={ui.listMeta}>
            {rows.length} 件（event_id は必須）
            {!validate.ok && rows.length > 0 ? (
              <span style={{ marginLeft: 10, color: "#991b1b", fontWeight: 900 }}>※ {validate.msg}</span>
            ) : null}
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            <button style={ui.btnGhost} onClick={clearAll} disabled={!hasFiles || uploading}>
              クリア
            </button>
          </div>
        </div>

        {/* List card */}
        <div style={{ ...ui.card, marginTop: 10 }}>
          {/* header row */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "44px 1fr 240px 120px 120px 86px",
              gap: 10,
              background: "#f8fafc",
              borderBottom: "1px solid rgba(226,232,240,0.85)",
              fontSize: 12,
              color: "#64748b",
              padding: "10px 12px",
              fontWeight: 900,
            }}
          >
            <div>#</div>
            <div>ファイル</div>
            <div>講演会ID</div>
            <div>サイズ</div>
            <div>状態</div>
            <div></div>
          </div>

          {rows.map((r, idx) => {
            const invalid = !normalizeEventId(r.eventId);
            return (
              <div key={r.id} style={ui.item(r.status)}>
                <div style={{ color: "#64748b", fontWeight: 900 }}>{idx + 1}</div>

                <div style={{ minWidth: 0 }}>
                  <div style={ui.fileTitle}>{r.filename}</div>
                  {(r.jobId || r.previewUrl) && (
                    <div style={ui.fileSub}>
                      {r.jobId ? (
                        <>
                          jobId: <code>{String(r.jobId).slice(0, 8)}…</code>
                        </>
                      ) : null}
                      {r.previewUrl ? (
                        <>
                          {" "}
                          ・{" "}
                          <a href={r.previewUrl} target="_blank" rel="noreferrer">
                            preview
                          </a>
                        </>
                      ) : null}
                    </div>
                  )}
                  {r.message ? <div style={ui.msgErr}>{r.message}</div> : null}
                </div>

                <div>
                  <input
                    value={r.eventId}
                    disabled={uploading}
                    onChange={(e) => setEventId(r.id, e.target.value)}
                    placeholder="例: EM0000-00000000"
                    style={ui.input(invalid, uploading)}
                  />
                </div>

                <div style={{ color: "#64748b", fontSize: 12, fontWeight: 900 }}>{fmtSize(r.size)}</div>

                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <StatusPill status={r.status} />
                </div>

                <div style={{ display: "flex", justifyContent: "flex-end" }}>
                  <button style={ui.btnGhost} onClick={() => removeRow(r.id)} disabled={uploading}>
                    削除
                  </button>
                </div>
              </div>
            );
          })}

          {!rows.length && (
            <div style={{ padding: 16, color: "#64748b", fontSize: 13 }}>
              まだファイルがありません。上の領域に .pptx / .pdf をドロップするか、ファイル選択してください。
            </div>
          )}
        </div>

        {/* footer tips */}
        <div style={{ marginTop: 10, fontSize: 12, color: "#64748b", lineHeight: 1.6 }}>
          ・event_id は必須（未入力があると生成できません）<br />
          ・同名ファイルが複数あってもOK（ここでは重複排除していません）
        </div>
        </>
        )}
        
        {tab === "restore" && (
  <div style={{ marginTop: 12 }}>
   <div style={ui.card}>
  <div style={ui.cardInner}>
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setRestoreDragOver(true);
      }}
      onDragLeave={() => setRestoreDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setRestoreDragOver(false);
        addRestoreFiles(e.dataTransfer.files);
      }}
      style={ui.drop(restoreDragOver)}
    >
      <div style={ui.row}>
        <div style={{ fontWeight: 950 }}>
          ここに <span style={{ color: "#2563eb" }}>.json</span> をドロップ
        </div>

        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <input
            ref={restoreInputRef}
            type="file"
            multiple
            accept="application/json,.json"
            disabled={restoring || uploading}
            style={{ display: "none" }}
            onChange={(e) => addRestoreFiles(e.target.files)}
          />
          <button
            style={ui.btnGhost}
            onClick={() => restoreInputRef.current?.click?.()}
            disabled={restoring || uploading}
          >
            JSON選択
          </button>

          


        </div>
      </div>

      <div style={ui.hint}>
        選択中: <b>{restoreFiles.length}</b> 件
        {restoreFiles.length ? (
          <>
            {" "}
            ・例: <code>{restoreFiles[0].name}</code>
          </>
        ) : null}
      </div>
    </div>

                {!!restoreError && <div style={ui.alertErr}>{restoreError}</div>}
                
  {restoreFiles.length > 0 && (
  <div style={{ ...ui.card, marginTop: 12 }}>
    <div style={ui.restoreHead}>
      <div>#</div>
      <div>JSON</div>
      <div>eventId</div>
      <div>状態</div>
      <div style={{ textAlign: "right" }}>操作</div>
    </div>

    {restoreItems.map((it, idx) => {
      const r = it.result;
      return (
        <div key={it.key} style={ui.restoreRow(it.status === "pending" ? "pending" : it.status)}>
          <div style={{ color: "#64748b", fontWeight: 900 }}>{idx + 1}</div>

          <div style={{ minWidth: 0 }}>
            <div style={{ fontWeight: 900, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {it.filename}
            </div>

            <div style={ui.mini}>
              {fmtSize(it.size)}
              {r?.jobId ? (
                <>
                  {" "}
                  ・ jobId: <span style={ui.codeChip}>{String(r.jobId).slice(0, 8)}…</span>
                </>
              ) : null}
            </div>

            {r?.error ? (
              <div style={{ marginTop: 6, fontSize: 12, color: "#991b1b", whiteSpace: "pre-wrap" }}>
                {String(r.error)}
              </div>
            ) : null}
          </div>

          <div style={ui.mini}>
            {r?.eventId ? <span style={ui.codeChip}>{String(r.eventId)}</span> : "-"}
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <StatusPill status={it.status} />
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, flexWrap: "wrap" }}>
            {/* pendingでも消せる */}
            <button
              style={ui.btnSmall}
              disabled={restoring || uploading}
              onClick={() => {
                if (restoring || uploading) return;
                setRestoreFiles((prev) => prev.filter((x) => !(x.name === it.file.name && x.size === it.file.size)));
                // 行を消したら結果も掃除（任意）
                setRestoreRows((prev) => (prev || []).filter((x) => String(x.filename || "") !== it.file.name));
              }}
            >
              削除
            </button>

            {r?.ok && r?.previewUrl ? (
              <a
                href={`${(import.meta.env.VITE_API_BASE || "")}${r.previewUrl}?v=${Date.now()}`}
                target="_blank"
                rel="noreferrer"
                style={{ textDecoration: "none" }}
              >
                <button style={ui.btnSmall}>preview</button>
              </a>
            ) : null}

            {r?.ok && r?.jobId ? (
              <button style={ui.btnSmall} onClick={() => nav(`/edit/${encodeURIComponent(r.jobId)}`)}>
                編集
              </button>
            ) : null}

            {r?.ok && r?.jobId ? (
              <a
                href={`${(import.meta.env.VITE_API_BASE || "")}/debug/${encodeURIComponent(r.jobId)}/latest.json`}
                target="_blank"
                rel="noreferrer"
                style={{ textDecoration: "none" }}
              >
                <button style={ui.btnSmall}>json</button>
              </a>
            ) : null}
          </div>
        </div>
      );
    })}
  </div>
)}

      
      </div>
    </div>
  </div>
)}
        
      </div>

    </div>


    
  );
}
