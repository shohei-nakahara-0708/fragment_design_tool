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

    const r = await fetch("/upload/batch/stream", { method: "POST", body: fd });
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
          <div>
            <h1 style={ui.h1}>Upload</h1>
            <div style={ui.sub}>
              複数PPTX / PDFをまとめてアップロード → event_id を付与 → 生成
            </div>
          </div>
          <div style={ui.headerRight}>
            <Link to="/jobs" style={{ textDecoration: "none" }}>
              <button style={ui.btnGhost}>一覧へ</button>
            </Link>
            <button style={ui.btnDanger} onClick={clearAll} disabled={!hasFiles || uploading}>
              全削除
            </button>
            <button style={ui.btnPrimary} onClick={uploadBatch} disabled={!hasFiles || uploading || !validate.ok}>
              生成する
            </button>
          </div>
        </div>

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
                    style={{opacity:0}}
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
  <div style={{ marginTop: 10, background:"#fff", border:"1px solid #eee", borderRadius:12, padding:10 }}>
    <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, color:"#555" }}>
      <span>⏳ {progressText || "処理中…"}</span>
      <span>{progress}%</span>
    </div>
    <div style={{ marginTop:8, height:10, background:"#eee", borderRadius:999, overflow:"hidden" }}>
      <div style={{ width:`${progress}%`, height:"100%", background:"#2563eb", transition:"width 200ms ease" }} />
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
      </div>
    </div>
  );
}
