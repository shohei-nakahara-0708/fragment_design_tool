import React, { useEffect, useMemo, useState, useCallback } from "react";
import { Link, useNavigate } from "react-router-dom";
const API_BASE = import.meta.env.VITE_API_BASE || "";
/** ---------- helpers ---------- */
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
function dateformated(s) {
  if (!s) return "";
  const d = new Date(s);
  return d.toLocaleString("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}
function clamp(v, min, max) {
  return Math.max(min, Math.min(max, v));
}

/** Blob download with custom filename */
async function downloadWithFilename(url, filename) {
  
  const r = await fetch(`${API_BASE}${url}`, { cache: "no-store" });
  if (!r.ok) throw new Error(`download failed: ${r.status} ${await r.text()}`);
  const blob = await r.blob();
  const a = document.createElement("a");
  const objectUrl = URL.createObjectURL(blob);
  a.href = objectUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(objectUrl);
}

/** ---------- UI tokens ---------- */
const ui = {
  page: {
    minHeight: "100vh",
    background: "#f6f7f9",
    color: "#0f172a",
    fontFamily:
      '"Invention JP", ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial',
  },
  container: {
    maxWidth: 1320,
    margin: "0 auto",
    padding: "22px 18px 26px",
  },
  top: {
    position: "sticky",
    top: 0,
    zIndex: 20,
    background: "rgba(246,247,249,0.78)",
    backdropFilter: "blur(10px)",
    borderBottom: "1px solid rgba(226,232,240,0.8)",
    padding: "14px 0 12px",
  },
  headerRow: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: 12,
  },
  title: { margin: 0, fontSize: 22, fontWeight: 900, letterSpacing: "-0.02em" },
  subtitle: { marginTop: 6, fontSize: 13, color: "#475569", lineHeight: 1.4 },
  headerRight: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },

  buttonPrimary: {
    background: "#2563eb",
    color: "#fff",
    padding: "10px 14px",
    borderRadius: 12,
    border: "1px solid rgba(37,99,235,0.35)",
    fontWeight: 800,
    fontSize: 13,
    cursor: "pointer",
    boxShadow: "0 10px 22px rgba(37,99,235,0.18)",
    transition: "transform .12s ease, background .12s ease, box-shadow .12s ease",
    userSelect: "none",
  },
  buttonGhost: {
    background: "rgba(255,255,255,0.85)",
    color: "#0f172a",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid rgba(226,232,240,0.95)",
    fontWeight: 800,
    fontSize: 13,
    cursor: "pointer",
    transition: "transform .12s ease, box-shadow .12s ease",
    userSelect: "none",
  },
   buttonGhost2: {
    background: "rgba(255,255,255,0.85)",
    color: "#0f172a",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid rgba(226,232,240,0.95)",
    fontWeight: 800,
    fontSize: 13,
    cursor: "pointer",
    transition: "transform .12s ease, box-shadow .12s ease",
    userSelect: "none",
  },
  disabled: {opacity: 0.5, cursor: "not-allowed", pointerEvents: "none" },
  buttonDanger: {
    background: "rgba(239,68,68,0.10)",
    color: "#991b1b",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid rgba(239,68,68,0.25)",
    fontWeight: 900,
    fontSize: 13,
    cursor: "pointer",
    transition: "transform .12s ease",
    userSelect: "none",
  },

  controls: { display: "grid", gap: 10, marginTop: 12 },
  controlRow: { display: "flex", gap: 10, alignItems: "center" },
  input: {
    width: "100%",
    padding: "11px 12px",
    borderRadius: 12,
    border: "1px solid rgba(226,232,240,0.95)",
    background: "#fff",
    outline: "none",
    fontSize: 13,
  },
  select: {
    width: "100%",
    padding: "11px 12px",
    borderRadius: 12,
    border: "1px solid rgba(226,232,240,0.95)",
    background: "#fff",
    outline: "none",
    fontSize: 13,
  },
  miniRow: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },
  metaText: { fontSize: 13, color: "#64748b" },

  layout: {
    display: "grid",
    gridTemplateColumns: "1fr",
    gap: 14,
    alignItems: "start",
    marginTop: 14,
  },
  panel: {
    background: "rgba(255,255,255,0.9)",
    border: "1px solid rgba(226,232,240,0.95)",
    borderRadius: 16,
    boxShadow: "0 12px 30px rgba(2,6,23,0.06)",
    overflow: "hidden",
  },
  panelInner: { padding: 14 },

  // list
  listGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))",
    gap: 12,
  },
  sessionHeader: {
    gridColumn: "1 / -1",
    marginTop: 6,
    padding: "10px 12px",
    border: "1px solid rgba(226,232,240,0.95)",
    borderRadius: 14,
    background: "#f8fafc",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 12,
  },
  sessionTitle: { fontWeight: 900, color: "#0f172a" },

  card: {
    border: "1px solid rgba(226,232,240,0.95)",
    borderRadius: 16,
    overflow: "hidden",
    background: "#fff",
    boxShadow: "0 1px 0 rgba(0,0,0,0.03)",
    transition: "transform .12s ease, box-shadow .12s ease, border-color .12s ease",
  },
  cardActive: {
    borderColor: "rgba(37,99,235,0.40)",
    boxShadow: "0 16px 30px rgba(37,99,235,0.12)",
  },
  thumbWrap: { position: "relative", background: "#0b1220" },
  thumb: { width: "100%", display: "block", cursor: "pointer" },
  selectBadge: {
    position: "absolute",
    top: 10,
    left: 10,
    background: "rgba(255,255,255,0.92)",
    border: "1px solid rgba(226,232,240,0.95)",
    borderRadius: 12,
    padding: "6px 8px",
    display: "flex",
    gap: 8,
    alignItems: "center",
    cursor: "pointer",
  },
  cardBody: { padding: 12, display: "grid", gap: 8 },
  cardTitle: { fontWeight: 900, lineHeight: 1.25, fontSize: 13, color: "#0f172a" },
  cardSub: { fontSize: 13, color: "#475569" },
  chips: { display: "flex", gap: 6, flexWrap: "wrap" },
  chip: {
    display: "inline-flex",
    alignItems: "center",
    fontSize: 11,
    fontWeight: 900,
    padding: "4px 8px",
    borderRadius: 999,
    border: "1px solid rgba(226,232,240,0.95)",
    background: "#f8fafc",
    color: "#0f172a",
  },
  cardActions: { display: "flex", gap: 8, alignItems: "center", marginTop: 2 },
  smallBtn: {
    padding: "8px 10px",
    borderRadius: 12,
    border: "1px solid rgba(226,232,240,0.95)",
    background: "#fff",
    fontSize: 12,
    fontWeight: 900,
    cursor: "pointer",
    position: "absolute",
    top: 10,
    right: 10,
  },

  // right preview
 sticky: {
  position: "sticky",
  top: 88,
  height: "calc(100vh - 104px)",   // ← ここが重要（top分引く）
  overflow: "hidden",              // 外枠は隠す
  alignSelf: "start",
},
  previewTop: {
    padding: "12px 12px",
    borderBottom: "1px solid rgba(226,232,240,0.95)",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
    background: "linear-gradient(180deg, rgba(248,250,252,0.8), rgba(255,255,255,0.85))",
  },
  previewTitle: { fontWeight: 950, fontSize: 12, letterSpacing: "0.08em", color: "#334155" },
  previewMainTitle: { fontWeight: 900, fontSize: 14, color: "#0f172a" },
  previewMeta: { marginTop: 4, fontSize: 12, color: "#64748b", lineHeight: 1.4 },
  previewBtns: { display: "flex", gap: 8, flexWrap: "wrap" },
  previewBody: { background: "#0b1220",overflow: "auto",    },
  previewImgWrap: { padding: 12, display: "grid", placeItems: "center" },
  previewImg: {
    maxWidth: 600,
    width: "100%",
    height: "auto",
    borderRadius: 12,
    border: "1px solid rgba(255,255,255,0.10)",
    boxShadow: "0 18px 50px rgba(0,0,0,0.35)",
    background: "#0b1220",
  },
  empty: { padding: 14, color: "#64748b", fontSize: 13, lineHeight: 1.5 },

  pager: { marginTop: 14, display: "flex", gap: 10, alignItems: "center" },

  
};

export default function JobListPage() {
  const nav = useNavigate();

  const [items, setItems] = useState([]);
  const [selected, setSelected] = useState(new Set());
  const [q, setQ] = useState("");
  const [sort, setSort] = useState("created_desc"); // created_desc | filename_asc
  const [fromDate, setFromDate] = useState(""); // "YYYY-MM-DD"
  const [toDate, setToDate] = useState(""); // "YYYY-MM-DD"
  const [page, setPage] = useState(1);
  const [pageSize] = useState(30);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(false);

  // right preview (sticky)
  const [preview, setPreview] = useState(null);

  // jpg download cache buster (manual)
  const [previewBuster, setPreviewBuster] = useState(Date.now());

  useEffect(() => {
    const run = async () => {
      setLoading(true);
      try {
        const params = new URLSearchParams();
        if (q.trim()) params.set("q", q.trim());
        params.set("order", "created_desc");
        params.set("page", String(page));
        params.set("page_size", String(pageSize));

        const cf = toIsoStart(fromDate);
        const ct = toIsoEnd(toDate);
        if (cf) params.set("created_from", cf);
        if (ct) params.set("created_to", ct);

        const API_BASE = import.meta.env.VITE_API_BASE || "";
        const r = await fetch(`${API_BASE}/jobs?${params.toString()}`);
        if (!r.ok) throw new Error(await r.text());
        const d = await r.json();

        setItems(d.items || []);
        setTotal(d.total || 0);

        // previewが消えてたら先頭を選ぶ（任意）
        // setPreview((p) => p || ((d.items || [])[0] ? buildPreview((d.items || [])[0]) : null));
      } catch (e) {
        console.error(e);
        alert(`list failed: ${String(e)}`);
      } finally {
        setLoading(false);
      }
    };

    run();
  }, [q, fromDate, toDate, page, pageSize]);

  const visibleItems = useMemo(() => {
    let arr = items;

    if (sort === "filename_asc") {
      arr = [...arr].sort((a, b) => (a.filename || "").localeCompare(b.filename || "", "ja"));
    } else {
      arr = [...arr].sort((a, b) => String(b.createdAt || b.created_at || "").localeCompare(String(a.createdAt || a.created_at || "")));
    }
    return arr;
  }, [items, sort]);

  const groups = useMemo(() => {
    const mp = new Map();
    for (const it of visibleItems) {
      const key = it.session_id || "no_session";
      if (!mp.has(key)) mp.set(key, []);
      mp.get(key).push(it);
    }
    const keys = Array.from(mp.keys()).sort((a, b) => {
      const maxA = Math.max(...mp.get(a).map((x) => +new Date(x.createdAt || x.created_at || 0)));
      const maxB = Math.max(...mp.get(b).map((x) => +new Date(x.createdAt || x.created_at || 0)));
      return maxB - maxA;
    });
    return keys.map((k) => ({
      sessionId: k,
      items: [...mp.get(k)].sort((a, b) =>
        String(b.createdAt || b.created_at || "").localeCompare(String(a.createdAt || a.created_at || ""))
      ),
    }));
  }, [visibleItems]);

  const toggle = (jobId) => {
    setSelected((prev) => {
      const n = new Set(prev);
      if (n.has(jobId)) n.delete(jobId);
      else n.add(jobId);
      return n;
    });
  };
  const clearSelected = () => setSelected(new Set());

  const exportZip = async () => {
    if (selected.size === 0) return;
    const jobIds = Array.from(selected);

    const API_BASE = import.meta.env.VITE_API_BASE || "";

    const r = await fetch(`${API_BASE}/jobs/export.zip`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ jobIds, nameMode: "filename", includeJson: false }),
    });

    if (!r.ok) {
      const t = await r.text();
      alert(`export failed: ${r.status}\n${t}`);
      return;
    }

    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "export.zip";
    a.click();
    URL.revokeObjectURL(url);
  };

  const deleteSelected = async () => {
    if (selected.size === 0) return;
    if (!confirm(`選択した ${selected.size} 件を削除しますか？`)) return;

    const jobIds = Array.from(selected);
    const API_BASE = import.meta.env.VITE_API_BASE || "";
    for (const id of jobIds) {
      const r = await fetch(`${API_BASE}/job/${id}`, { method: "DELETE" });
      if (!r.ok) {
        const t = await r.text();
        alert(`delete failed: ${id}\n${t}`);
        return;
      }
    }

    setItems((prev) => prev.filter((it) => !selected.has(it.job_id || it.jobId)));
    
    clearSelected();
    setPreview(null);
  };

  const selectAllOnPage = () => {
    setSelected((prev) => {
      const n = new Set(prev);
      visibleItems.forEach((it) => n.add(it.job_id || it.jobId));
      return n;
    });
  };

  const buildPreview = useCallback((it) => {
    const id = it.job_id || it.jobId;
    const updatedKey =
      it.updated_at || it.updatedAt || it.preview_updated_at || it.previewUpdatedAt || it.created_at || it.createdAt || "";
    const previewUrl = it.previewUrl || `/preview/${id}.jpg`;
    const src = `${previewUrl}?v=${encodeURIComponent(updatedKey || Date.now())}`;

    return {
      id,
      src,
      title: it.event_title || it.filename || id,
      filename: it.filename || "",
      createdAt: it.createdAt || it.created_at || "",
      eventId: it.event_id || "",
      warnings: it.warnings || [],
    };
  }, []);

  const openPreview = (it) => setPreview(buildPreview(it));

  // ESCでクリア（右ペインなので「閉じる」＝未選択にする）
  useEffect(() => {
    const onKey = (e) => {
      if (e.key === "Escape") setPreview(null);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  const onDownloadJpg = useCallback(async () => {
    if (!preview?.id) return;
    const eventIdLike = preview.eventId || preview.id || "event";
    const filename = `${eventIdLike}_招聘.jpg`;

    // 既存の /download/:id.jpg を利用（キャッシュ対策に t を付ける）
    const url = `/download/${preview.id}.jpg?t=${encodeURIComponent(previewBuster)}`;
    await downloadWithFilename(url, filename);
  }, [preview, previewBuster]);

  return (
    <div style={ui.page}>
      <div style={ui.container}>
        {/* Top */}
        <div style={ui.top}>
          <div style={ui.headerRow}>
            <div>
              <h1 style={ui.title}>イベント一覧</h1>
              <div style={ui.subtitle}>
                {total} 件 / 選択 {selected.size} 件{loading ? "（読み込み中…）" : ""}
              </div>
            </div>

            <div style={ui.headerRight}>
              <Link to="/upload" style={{ textDecoration: "none" }}>
                <button
                  style={ui.buttonPrimary}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.transform = "translateY(-1px)";
                    e.currentTarget.style.background = "#1d4ed8";
                    e.currentTarget.style.boxShadow = "0 14px 26px rgba(37,99,235,0.22)";
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.transform = "none";
                    e.currentTarget.style.background = "#2563eb";
                    e.currentTarget.style.boxShadow = "0 10px 22px rgba(37,99,235,0.18)";
                  }}
                >
                  ＋ Upload
                </button>
              </Link>

              <button style={{ ...ui.buttonGhost2, ...(selected.size === 0 ? ui.disabled : {}) }} onClick={exportZip} disabled={selected.size === 0}>
                ZIP Export
              </button>

              <button style={{ ...ui.buttonDanger, ...(selected.size === 0 ? ui.disabled : {}) }} onClick={deleteSelected} disabled={selected.size === 0}>
                削除
              </button>

              
            </div>
          </div>

          <div style={ui.controls}>
            <div style={ui.controlRow}>
              <input
                placeholder="検索（ファイル名/タイトル）"
                value={q}
                onChange={(e) => {
                  setQ(e.target.value);
                  setPage(1);
                }}
                style={ui.input}
              />

              <input
                type="date"
                value={fromDate}
                onChange={(e) => {
                  setFromDate(e.target.value);
                  setPage(1);
                }}
                style={{ ...ui.input, width: 160 }}
              />

              <input
                type="date"
                value={toDate}
                onChange={(e) => {
                  setToDate(e.target.value);
                  setPage(1);
                }}
                 style={{ ...ui.input, width: 160 }}
              />

              {/* <select value={sort} onChange={(e) => setSort(e.target.value)} style={ui.select}>
                <option value="created_desc">新しい順</option>
                <option value="filename_asc">ファイル名順</option>
              </select> */}
            </div>

            <div style={ui.miniRow}>
              <button style={ui.buttonGhost} onClick={() => { setFromDate(""); setToDate(""); setPage(1); }} disabled={!fromDate && !toDate}>
                日付クリア
              </button>
              <button style={ui.buttonGhost} onClick={selectAllOnPage}>
                表示分を全選択
              </button>
              <button style={ui.buttonGhost} onClick={clearSelected} disabled={selected.size === 0}>
                選択解除
              </button>

              
            </div>
          </div>
        </div>

        {/* Main layout */}
        <div style={ui.layout}>
          {/* Left panel */}
          <div style={ui.panel}>
            <div style={ui.panelInner}>
              {groups.length === 0 && <div style={ui.empty}>該当なし</div>}

              <div style={ui.listGrid}>
                {groups.map((group) => (
                  <React.Fragment key={group.sessionId}>
                    <div style={ui.sessionHeader}>
                      <div style={ui.sessionTitle}>
                        {dateformated(group.items?.[0]?.createdAt || group.items?.[0]?.created_at)}（{group.items.length}件）
                      </div>
                      <button
                        style={ui.buttonGhost}
                        onClick={() => {
                          setSelected((prev) => {
                            const n = new Set(prev);
                            group.items.forEach((it) => n.add(it.job_id || it.jobId));
                            return n;
                          });
                        }}
                      >
                        このセッションを全選択
                      </button>
                    </div>

                    {group.items.map((it) => {
                      const id = it.job_id || it.jobId;
                      const checked = selected.has(id);

                      const API_BASE = import.meta.env.VITE_API_BASE || "";

                      const previewUrl =
                      it.previewUrl
                        ? `${API_BASE}${it.previewUrl}`
                        : `${API_BASE}/preview/${id}.jpg`;
                      const updatedKey =
                        it.updated_at || it.updatedAt || it.preview_updated_at || it.previewUpdatedAt || it.created_at || it.createdAt || "";
                      const previewSrc = `${previewUrl}?v=${encodeURIComponent(updatedKey || Date.now())}`;

                      const active = preview?.id === id;

                      return (
                        <div key={id} style={{ ...ui.card, ...(active ? ui.cardActive : null) }}>
                          <div style={ui.thumbWrap}>
                            <img src={previewSrc} loading="lazy" alt="" style={ui.thumb} onClick={() => openPreview(it)} />
                            <label style={ui.selectBadge}>
                              <input type="checkbox" checked={checked} onChange={() => toggle(id)} />
                              <span style={{ fontSize: 12, fontWeight: 900 }}>選択</span>
                            </label>
                            <button style={ui.smallBtn} onClick={() => nav(`/job/${id}`)}>
                                編集
                              </button>
                          </div>

                          <div style={ui.cardBody}>
                            <div style={ui.cardTitle}>{it.event_title || it.filename || "（タイトルなし）"}</div>

                            {it.event_id && <div style={ui.cardSub}>講演会ID：{it.event_id}</div>}

                            <div style={ui.chips}>
                              {(it.warnings || []).slice(0, 4).map((w) => (
                                <span key={w} style={ui.chip}>
                                  {w}
                                </span>
                              ))}
                            </div>

                            <div style={ui.cardActions}>
                              
                              <button
                                style={{
                                  padding: "10px 14px",
                                  borderRadius: 12,
                                  border: "none",
                                  background: "#2563eb",
                                  color: "#fff",
                                  fontWeight: 800,
                                  cursor: "pointer",
                                  boxShadow: "0 10px 24px rgba(37,99,235,0.25)",
                                }}
                                onClick={async () => {
                                  const eventIdLike = it.event_id || id || "event";
                                  const filename = `${eventIdLike}_招聘.jpg`;
                                  const url = `/download/${id}.jpg?t=${encodeURIComponent(previewBuster)}`;
                                  await downloadWithFilename(url, filename);
                                }}
                              >
                                ⬇ Download
                              </button>
          
                              {/* <button style={ui.smallBtn} onClick={() => setPreview(buildPreview(it))}>
                                Preview
                              </button> */}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </React.Fragment>
                ))}
              </div>

              {/* pager */}
              <div style={ui.pager}>
                <button style={ui.buttonGhost} onClick={() => setPage((p) => Math.max(1, p - 1))} disabled={page <= 1 || loading}>
                  前へ
                </button>
                <span style={{ color: "#64748b", fontSize: 13 }}>
                  {total}件 / {page}ページ目
                </span>
                <button style={ui.buttonGhost} onClick={() => setPage((p) => p + 1)} disabled={loading || page * pageSize >= total}>
                  次へ
                </button>
                <button style={ui.buttonGhost} onClick={() => setPreviewBuster(Date.now())} title="download/preview のキャッシュを更新">
                  ↻ キャッシュ更新
                </button>
              </div>
            </div>
          </div>

         {preview && (
  <div
    onMouseDown={(e) => {
      if (e.target === e.currentTarget) setPreview(null);
    }}
    style={{
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.65)",
      backdropFilter: "blur(6px)",
      zIndex: 1000,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: 20,
    }}
  >
    <div
      style={{
        width: "min(1400px, 96vw)",
        height: "min(94vh, 1000px)",
        background: "#fff",
        borderRadius: 18,
        overflow: "hidden",
        boxShadow: "0 30px 80px rgba(0,0,0,0.35)",
        display: "grid",
        gridTemplateRows: "auto 1fr",
      }}
    >
      {/* Header */}
      <div
        style={{
          padding: "14px 18px",
          borderBottom: "1px solid #eee",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          gap: 16,
        }}
      >
        <div style={{ minWidth: 0 }}>
          <div
            style={{
              fontWeight: 900,
              fontSize: 16,
              whiteSpace: "nowrap",
              overflow: "hidden",
              textOverflow: "ellipsis",
            }}
          >
            {preview.title}
          </div>

          <div style={{ fontSize: 13, color: "#666", marginTop: 4 }}>
            {preview.eventId && <>講演会ID：{preview.eventId}　</>}
            {preview.filename && <>ファイル：{preview.filename}　</>}
            {preview.createdAt && <>作成：{dateformated(preview.createdAt)}</>}
          </div>
        </div>

        <div style={{ display: "flex", gap: 10 }}>
          <button
            style={{
              padding: "10px 14px",
              borderRadius: 12,
              border: "1px solid #ddd",
              background: "#fff",
              fontWeight: 700,
              cursor: "pointer",
            }}
            onClick={() => nav(`/job/${preview.id}`)}
          >
            編集
          </button>

          <button
            style={{
              padding: "10px 14px",
              borderRadius: 12,
              border: "none",
              background: "#2563eb",
              color: "#fff",
              fontWeight: 800,
              cursor: "pointer",
              boxShadow: "0 10px 24px rgba(37,99,235,0.25)",
            }}
            onClick={onDownloadJpg}
          >
            ⬇ Download
          </button>

          <button
            style={{
              padding: "10px 14px",
              borderRadius: 12,
              border: "1px solid #ddd",
              background: "#fff",
              fontWeight: 700,
              cursor: "pointer",
            }}
            onClick={() => setPreview(null)}
          >
            ✕
          </button>
        </div>
      </div>

      {/* Body */}
      <div
        style={{
          overflow: "scroll",
          background: "#111",
          display: "flex",
          justifyContent: "center",
          alignItems: "flex-start",
          padding: 24,
        }}
      >
        <img
          src={`${API_BASE}${preview.src}`}
          alt=""
          style={{
            maxWidth: "100%",
            background: "#fff",
          }}
        />
      </div>
    </div>
  </div>
)}
        </div>
      </div>
    </div>
  );
}
