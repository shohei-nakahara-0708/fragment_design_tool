import React, { useEffect, useMemo, useRef, useState, useCallback } from "react";
import { Link, useParams } from "react-router-dom";

/**
 * ✅ この編集画面は以下を満たします
 * - 右のPreviewは sticky + スクロール枠
 * - 0.5秒デバウンスでリアルタイムレンダ（autoRender ON）
 * - /render へ { jobId, design } をPOST（※サーバ側 RenderReq に合わせて）
 * - hero の applyOverridesToLines 用に title_overrides を編集できる（配列追加/削除）
 * - talks も同じ仕組みで talk.title_overrides を編集できる
 * - datetime_parts と datetime の同期、timeはtextareaで改行OK
 */

function tryParseJson(text) {
  try {
    return { ok: true, value: JSON.parse(text) };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/** ---------- UI primitives (no deps) ---------- **/
const ui = {
  page: {
    display: "grid",
    gridTemplateColumns: "minmax(420px, 560px) 1fr",
    gap: 16,
    padding: 16,
    minHeight: "100vh",
    background: "#fafafa",
  },
  leftCol: { display: "grid", gap: 12, alignContent: "start" },
  rightCol: {
    display: "grid",
    gap: 12,
    alignContent: "start",
    position: "sticky",
    top: 16,
    height: "calc(100vh - 32px)",
  },
  card: {
    background: "#fff",
    border: "1px solid #e7e7e7",
    borderRadius: 14,
    padding: 14,
    boxShadow: "0 1px 2px rgba(0,0,0,0.04)",
  },
  headerRow: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12,marginBottom: 8 },
  h2: { margin: "6px 0 0", fontSize: 18, fontWeight: 800, letterSpacing: 0.2 },
  h3: { margin: 0, fontSize: 13, fontWeight: 800, color: "#222" },
  muted: { fontSize: 12, color: "#666" },
  badge: (tone = "gray") => {
    const base = {
      display: "inline-flex",
      alignItems: "center",
      gap: 6,
      padding: "5px 10px",
      borderRadius: 999,
      border: "1px solid #e3e3e3",
      fontSize: 12,
      lineHeight: 1,
      userSelect: "none",
      whiteSpace: "nowrap",
    };
    if (tone === "green") return { ...base, background: "#ecf8ef", borderColor: "#bfe3c6", color: "#1b6b2f" };
    if (tone === "red") return { ...base, background: "#fff2f2", borderColor: "#f2c2c2", color: "#a00000" };
    if (tone === "blue") return { ...base, background: "#eef5ff", borderColor: "#c7ddff", color: "#1a4fb3" };
    return { ...base, background: "#f6f6f6", color: "#444" };
  },

  grid2: { display: "grid", gridTemplateColumns: "1fr 160px", gap: 12, alignItems: "start" },
  field: { display: "grid", gap: 6, marginTop: 10 },
  label: { fontSize: 12, fontWeight: 700, color: "#333" },
  help: { fontSize: 11, color: "#777" },
  row: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },

  controlBase: {
    width: "100%",
    padding: "10px 10px",
    borderRadius: 12,
    border: "1px solid #ddd",
    outline: "none",
    background: "#fff",
    boxSizing: "border-box",
  },
  textarea: { resize: "vertical" },

  btn: (variant = "primary", disabled = false) => {
    const base = {
      appearance: "none",
      border: "1px solid transparent",
      borderRadius: 12,
      padding: "10px 12px",
      fontWeight: 800,
      letterSpacing: 0.2,
      cursor: disabled ? "not-allowed" : "pointer",
      userSelect: "none",
      display: "inline-flex",
      alignItems: "center",
      justifyContent: "center",
      gap: 8,
      lineHeight: 1,
      transition: "transform 0.05s ease, box-shadow 0.12s ease, background 0.12s ease, border-color 0.12s ease",
      boxShadow: disabled ? "none" : "0 8px 18px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.06)",
      transform: "translateY(0)",
      textDecoration: "none",
      opacity: disabled ? 0.55 : 1,
      whiteSpace: "nowrap",
    };

    if (variant === "primary") {
      return {
        ...base,
        color: "#fff",
        background: "linear-gradient(180deg, #222 0%, #0f0f0f 100%)",
        borderColor: "#0f0f0f",
      };
    }
    if (variant === "secondary") {
      return {
        ...base,
        color: "#111",
        background: "linear-gradient(180deg, #ffffff 0%, #f3f3f3 100%)",
        borderColor: "#d9d9d9",
      };
    }
    if (variant === "danger") {
      return {
        ...base,
        color: "#a00000",
        background: "linear-gradient(180deg, #ffffff 0%, #fff4f4 100%)",
        borderColor: "#f0caca",
      };
    }
    return {
      ...base,
      color: "#111",
      background: "linear-gradient(180deg, #ffffff 0%, #f6f6f6 100%)",
      borderColor: "#ddd",
    };
  },

  divider: { height: 1, background: "#eee", margin: "12px 0" },
  softBox: { border: "1px dashed #ddd", borderRadius: 12, padding: 12, background: "#fcfcfc" },
};

const styles = {
  previewFrame: {
    border: "1px solid #eee",
    borderRadius: 14,
    overflow: "auto",
    maxHeight: "calc(100vh - 140px)",
    background: "#F2F2F2",
    padding: 8,
  },
};

function Control({ as = "input", style, ...props }) {
  const Tag = as;
  const merged = {
    ...ui.controlBase,
    ...(as === "textarea" ? ui.textarea : {}),
    ...style,
  };
  return <Tag style={merged} {...props} />;
}

function Field({ label, help, children }) {
  return (
    <div style={ui.field}>
      <div style={ui.label}>{label}</div>
      {children}
      {help ? <div style={ui.help}>{help}</div> : null}
    </div>
  );
}

function Card({ title, right, children }) {
  return (
    <div style={ui.card}>
      {title ? (
        <div style={ui.headerRow}>
          <div style={ui.h3}>{title}</div>
          {right ? <div>{right}</div> : <div />}
        </div>
      ) : null}
      {children}
    </div>
  );
}

/** ---------- helpers ---------- **/
function ensureBaseDefaults(j) {
  const next = { ...(j || {}) };

  if (!next.event_title_lines) next.event_title_lines = [];
  if (!next.title_font_size) next.title_font_size = 30;

  if (!next.datetime_parts) {
    next.datetime_parts = { year: "", month: "", day: "", dow: "", time: "" };
  } else {
    next.datetime_parts = {
      year: next.datetime_parts.year ?? "",
      month: next.datetime_parts.month ?? "",
      day: next.datetime_parts.day ?? "",
      dow: next.datetime_parts.dow ?? "",
      time: next.datetime_parts.time ?? "",
    };
  }

  if (!next.chair) next.chair = { name_display: "", affiliation: "" };
  if (!Array.isArray(next.talks)) next.talks = [];

  // hero override (applyOverridesToLines の入力)
  if (!Array.isArray(next.title_overrides)) next.title_overrides = [];

  // talks override
  next.talks = next.talks.map((t) => ({
    time: t?.time ?? "",
    title: t?.title ?? "",
    title_lines: Array.isArray(t?.title_lines) ? t.title_lines : [],
    speaker: t?.speaker ?? "",
    speaker_display: t?.speaker_display ?? "",
    affiliation: t?.affiliation ?? "",
    // ここが追加ポイント
    title_overrides: Array.isArray(t?.title_overrides) ? t.title_overrides : [],
  }));

  return next;
}

/** ---------- Editors ---------- **/
const ChairEditor = React.memo(function ChairEditor({ chair, updateAtPath }) {
  const c = chair || {};
  return (
    <Card title="座長">
      <Field label="名前">
        <Control
          value={c.name_display || ""}
          onChange={(e) => updateAtPath(["chair", "name_display"], e.target.value)}
        />
      </Field>

      <Field label="所属">
        <Control
          as="textarea"
          rows={3}
          value={c.affiliation || ""}
          onChange={(e) => updateAtPath(["chair", "affiliation"], e.target.value)}
        />
      </Field>
    </Card>
  );
});

function OverrideRow({ o, onChange, onDelete, labelPrefix = "" }) {
  const set = (k, v) => onChange({ ...(o || {}), [k]: v });

  return (
    <div style={{ ...ui.softBox, padding: 10 }}>
      <div style={ui.headerRow}>
        <div style={{ fontWeight: 800, fontSize: 12 }}>{labelPrefix} 装飾指定</div>
        <button type="button" onClick={onDelete} style={ui.btn("danger")}>
          削除
        </button>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "90px 1fr", gap: 10, marginTop: 8, alignItems: "center" }}>
        <div style={ui.muted}>対象行数</div>
        <Control
          type="number"
          value={o?.index ?? ""}
          placeholder="例: 2"
          onChange={(e) => set("index", e.target.value === "" ? null : Number(e.target.value))}
        />

        {/* <div style={ui.muted}>target</div>
        <Control
          value={o?.target ?? ""}
          placeholder="部分一致（例: ～肺炎球菌ワクチン）"
          onChange={(e) => set("target", e.target.value)}
        /> */}

        <div style={ui.muted}>フォントサイズ</div>
        <Control
          type="number"
          value={o?.font_size ?? ""}
          placeholder="例: 28"
          onChange={(e) => set("font_size", e.target.value === "" ? null : Number(e.target.value))}
        />

        {/* <div style={ui.muted}>font_weight</div>
        <Control
          type="number"
          value={o?.font_weight ?? ""}
          placeholder="例: 700"
          onChange={(e) => set("font_weight", e.target.value === "" ? null : Number(e.target.value))}
        /> */}

        {/* <div style={ui.muted}>color</div>
        <Control
          value={o?.color ?? ""}
          placeholder="#4b8d41"
          onChange={(e) => set("color", e.target.value)}
        /> */}

        <div style={ui.muted}>letter_spacing</div>
        <Control
          type="number"
          value={o?.letter_spacing ?? ""}
          placeholder="px"
          onChange={(e) => set("letter_spacing", e.target.value === "" ? null : Number(e.target.value))}
        />

        <div style={ui.muted}>line_height</div>
        <Control
          type="number"
          value={o?.line_height ?? ""}
          placeholder="例: 35"
          onChange={(e) => set("line_height", e.target.value === "" ? null : Number(e.target.value))}
        />
{/* 
        <div style={ui.muted}>font_family</div>
        <Control
          value={o?.font_family ?? ""}
          placeholder='例: "Invention JP"'
          onChange={(e) => set("font_family", e.target.value)}
        /> */}
      </div>

      {/* <div style={{ marginTop: 8, fontSize: 11, color: "#777" }}>
        index があれば index 優先。無ければ target（部分一致）で適用。
      </div> */}
    </div>
  );
}

const HeroOverridesEditor = React.memo(function HeroOverridesEditor({ json, updateAtPath }) {
  const arr = Array.isArray(json.title_overrides) ? json.title_overrides : [];

  const add = () => {
    const next = [...arr, { index: null, target: "", font_size: 22, font_weight: 700, color: "" }];
    updateAtPath(["title_overrides"], next);
  };

  const updateOne = (idx, obj) => {
    const next = arr.slice();
    next[idx] = obj;
    updateAtPath(["title_overrides"], next);
  };

  const remove = (idx) => {
    const next = arr.filter((_, i) => i !== idx);
    updateAtPath(["title_overrides"], next);
  };

  return (
    <Card
      title="タイトル行ごとの装飾"
      right={
        <button type="button" onClick={add} style={ui.btn("secondary")}>
          + 装飾追加
        </button>
      }
    >
     

      {arr.length === 0 ? <div style={{ ...ui.muted, marginTop: 8 }}>まだ指定された装飾はありません。</div> : null}

      <div style={{ display: "grid", gap: 10, marginTop: 10 }}>
        {arr.map((o, i) => (
          <OverrideRow
            key={i}
            o={o}
            onChange={(v) => updateOne(i, v)}
            onDelete={() => remove(i)}
            labelPrefix="タイトル "
          />
        ))}
      </div>
    </Card>
  );
});

const TalksEditor = React.memo(function TalksEditor({ talks, updateAtPath }) {
  const arr = Array.isArray(talks) ? talks : [];
  const setTalkField = (idx, key, value) => updateAtPath(["talks", idx, key], value);

  const addTalk = () => {
    const next = [
      ...arr,
      {
        time: "",
        title: "",
        title_lines: [],
        speaker: "",
        speaker_display: "",
        affiliation: "",
        title_overrides: [],
      },
    ];
    updateAtPath(["talks"], next);
  };

  const removeTalk = (idx) => {
    const next = arr.filter((_, i) => i !== idx);
    updateAtPath(["talks"], next);
  };

  const addTalkOverride = (idx) => {
    const cur = arr[idx] || {};
    const curOv = Array.isArray(cur.title_overrides) ? cur.title_overrides : [];
    const nextOv = [...curOv, { index: null, target: "", font_size: "", font_weight: "", color: "" }];
    setTalkField(idx, "title_overrides", nextOv);
  };

  const updateTalkOverride = (talkIdx, ovIdx, obj) => {
    const cur = arr[talkIdx] || {};
    const curOv = Array.isArray(cur.title_overrides) ? cur.title_overrides : [];
    const nextOv = curOv.slice();
    nextOv[ovIdx] = obj;
    setTalkField(talkIdx, "title_overrides", nextOv);
  };

  const removeTalkOverride = (talkIdx, ovIdx) => {
    const cur = arr[talkIdx] || {};
    const curOv = Array.isArray(cur.title_overrides) ? cur.title_overrides : [];
    const nextOv = curOv.filter((_, i) => i !== ovIdx);
    setTalkField(talkIdx, "title_overrides", nextOv);
  };

  return (
    <Card
      title="講演"
      right={
        <button type="button" onClick={addTalk} style={ui.btn("secondary")}>
          + 講演追加
        </button>
      }
    >
      {arr.length === 0 ? <div style={ui.muted}>講演がありません。右上から追加できます。</div> : null}

      {arr.map((t, idx) => (
        <div key={idx} style={{ marginTop: 12 }}>
          <div style={ui.softBox}>
            <div style={ui.headerRow}>
              <div style={{ fontWeight: 800 }}>講演 #{idx + 1}</div>
              <button type="button" onClick={() => removeTalk(idx)} style={ui.btn("danger")}>
                削除
              </button>
            </div>

            <Field label="時間" help="例: 19:00〜19:20">
              <Control value={t.time || ""} onChange={(e) => setTalkField(idx, "time", e.target.value)} />
            </Field>

            <Field label="タイトル" help="改行で行分割">
              <Control
                as="textarea"
                rows={3}
                value={(t.title_lines || []).join("\n")}
                onChange={(e) => setTalkField(idx, "title_lines", e.target.value.split("\n"))}
              />
            </Field>

            <Field label="演者">
              <Control
                value={t.speaker_display || ""}
                onChange={(e) => setTalkField(idx, "speaker_display", e.target.value)}
              />
            </Field>

            <Field label="所属">
              <Control
                as="textarea"
                rows={2}
                value={t.affiliation || ""}
                onChange={(e) => setTalkField(idx, "affiliation", e.target.value)}
              />
            </Field>

            <div style={ui.divider} />

            <div style={ui.headerRow}>
              <div style={{ fontWeight: 800, fontSize: 12 }}>タイトル 行ごとの装飾</div>
              <button type="button" onClick={() => addTalkOverride(idx)} style={ui.btn("secondary")}>
                + 装飾追加
              </button>
            </div>

            {(Array.isArray(t.title_overrides) ? t.title_overrides : []).length === 0 ? (
              <div style={{ ...ui.muted, marginTop: 8 }}>まだ指定された装飾はありません。</div>
            ) : null}

            <div style={{ display: "grid", gap: 10, marginTop: 10 }}>
              {(Array.isArray(t.title_overrides) ? t.title_overrides : []).map((o, ovIdx) => (
                <OverrideRow
                  key={ovIdx}
                  o={o}
                  onChange={(v) => updateTalkOverride(idx, ovIdx, v)}
                  onDelete={() => removeTalkOverride(idx, ovIdx)}
                  labelPrefix={`講演#${idx + 1} `}
                />
              ))}
            </div>
          </div>
        </div>
      ))}
    </Card>
  );
});

/** ---------- Page ---------- **/
export default function JobEditorPage() {
  const { jobId } = useParams();
  const [json, setJson] = useState(null);
  const [busy, setBusy] = useState(false);
  const [previewBuster, setPreviewBuster] = useState(Date.now());
  const [errors, setErrors] = useState({}); // reserved (JSON textarea使うなら)

  // --- realtime ---
  const [autoRender, setAutoRender] = useState(true);
  const debounceRef = useRef(null);
  const lastSentRef = useRef("");
  const inFlightRef = useRef(false);
  const pendingRef = useRef(false);

  useEffect(() => {
    fetch(`/job/${jobId}`)
      .then((r) => r.json())
      .then((d) => {
        const j = ensureBaseDefaults(d.json || {});
        setJson(j);
      });
  }, [jobId]);

  const saveRender = async (payload) => {
    if (!payload) return;
    if (Object.keys(errors).length > 0) return;

    if (inFlightRef.current) {
      pendingRef.current = true;
      return;
    }

    const s = JSON.stringify(payload);
    if (s === lastSentRef.current) return;

    inFlightRef.current = true;
    setBusy(true);
    try {
      const r = await fetch("/render", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ jobId, design: payload }), // ← サーバの RenderReq に合わせて
      });
      if (!r.ok) throw new Error("render failed");
      lastSentRef.current = s;
      setPreviewBuster(Date.now());
    } finally {
      setBusy(false);
      inFlightRef.current = false;

      if (pendingRef.current) {
        pendingRef.current = false;
        Promise.resolve().then(() => {
          setJson((cur) => {
            if (cur) saveRender(cur);
            return cur;
          });
        });
      }
    }
  };

  useEffect(() => {
    if (!json || !autoRender) return;
    if (Object.keys(errors).length > 0) return;

    if (debounceRef.current) clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => saveRender(json), 500);
    return () => debounceRef.current && clearTimeout(debounceRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [json, autoRender, errors]);

  const updateAtPath = useCallback((path, value) => {
    setJson((prev) => {
      if (!prev) return prev;

      const next = Array.isArray(prev) ? [...prev] : { ...prev };
      let curPrev = prev;
      let curNext = next;

      for (let i = 0; i < path.length - 1; i++) {
        const k = path[i];
        const prevChild = curPrev?.[k];

        let nextChild;
        if (Array.isArray(prevChild)) nextChild = [...prevChild];
        else if (prevChild && typeof prevChild === "object") nextChild = { ...prevChild };
        else nextChild = typeof path[i + 1] === "number" ? [] : {};

        curNext[k] = nextChild;
        curPrev = prevChild;
        curNext = nextChild;
      }

      curNext[path[path.length - 1]] = value;
      return next;
    });
  }, []);

  if (!json) return <div style={{ padding: 16 }}>loading...</div>;

  const dt = json.datetime_parts || { year: "", month: "", day: "", dow: "", time: "" };
  const rebuildDatetime = (parts) => {
    if (!parts.year || !parts.month || !parts.day) return "";
    // time は textarea なので改行もあり得る → そのまま連結
    return `${parts.year}年${parts.month}月${parts.day}日（${parts.dow || ""}）${parts.time || ""}`;
  };

  const hasJsonErrors = Object.keys(errors).length > 0;
  const statusTone = hasJsonErrors ? "red" : busy ? "blue" : autoRender ? "green" : "gray";
  const statusText = hasJsonErrors ? "JSON error" : busy ? "Rendering..." : autoRender ? "Auto" : "Manual";

  return (
    <div style={ui.page}>
      {/* Left */}
      <div style={ui.leftCol}>
        <Card
          title={
            <div style={{ display: "grid", gap: 6 }}>
              <div>
                <Link to="/" style={{ textDecoration: "none" }}>
                  ← 一覧へ
                </Link>
              </div>
              <div style={ui.h2}>編集: {(json.event_title_lines || []).join("")}_{json.event_id || ""}</div>
            </div>
          }
          right={<span style={ui.badge(statusTone)}>{statusText}</span>}
        >
          <div style={{ ...ui.row, marginTop: 10 }}>
            <label style={ui.badge(autoRender ? "green" : "gray")}>
              <input type="checkbox" checked={autoRender} onChange={(e) => setAutoRender(e.target.checked)} />
              リアルタイム反映（0.5s）
            </label>

            <div style={ui.muted}>
              {hasJsonErrors ? "エラーを修正すると自動反映が再開します。" : "変更を検知してプレビュー更新します。"}
            </div>
          </div>
        </Card>

        <Card title="基本">
          <Field label="VP/PH/ONC">
            <Control
              as="select"
              value={json.region || ""}
              onChange={(e) => updateAtPath(["region"], e.target.value)}
            >
              <option value="">-- 選択してください --</option>
              <option value="VP">VP</option>
              <option value="PH">PH</option>
              <option value="ONC">ONC</option>
            </Control>
          </Field>

          <div style={ui.grid2}>
            <div>
              <Field label="イベントタイトル" help="改行で行分割（event_title_lines）">
                <Control
                  as="textarea"
                  rows={4}
                  value={(json.event_title_lines || []).join("\n")}
                  onChange={(e) => updateAtPath(["event_title_lines"], e.target.value.split("\n"))}
                />
              </Field>
            </div>

            <div>
              <Field label="ベース文字サイズ" help="heroTitle の基本サイズ（例: 30）">
                <Control
                  type="number"
                  value={json.title_font_size || 30}
                  onChange={(e) => updateAtPath(["title_font_size"], Number(e.target.value))}
                />
              </Field>
            </div>
          </div>
        </Card>

        <HeroOverridesEditor json={json} updateAtPath={updateAtPath} />

        <Card title="日時">
          <div style={ui.row}>
            <Control
              style={{ width: 100 }}
              placeholder="2026"
              value={dt.year}
              onChange={(e) => {
                updateAtPath(["datetime_parts", "year"], e.target.value);
                updateAtPath(["datetime"], rebuildDatetime({ ...dt, year: e.target.value }));
              }}
            />
            <div style={ui.muted}>年</div>

            <Control
              style={{ width: 80 }}
              placeholder="3"
              value={dt.month}
              onChange={(e) => {
                updateAtPath(["datetime_parts", "month"], e.target.value);
                updateAtPath(["datetime"], rebuildDatetime({ ...dt, month: e.target.value }));
              }}
            />
            <div style={ui.muted}>月</div>

            <Control
              style={{ width: 80 }}
              placeholder="6"
              value={dt.day}
              onChange={(e) => {
                updateAtPath(["datetime_parts", "day"], e.target.value);
                updateAtPath(["datetime"], rebuildDatetime({ ...dt, day: e.target.value }));
              }}
            />
            <div style={ui.muted}>日</div>

            <div style={ui.muted}>(</div>

            <Control
              style={{ width: 70 }}
              placeholder="水"
              value={dt.dow}
              onChange={(e) => {
                updateAtPath(["datetime_parts", "dow"], e.target.value);
                updateAtPath(["datetime"], rebuildDatetime({ ...dt, dow: e.target.value }));
              }}
            />

            <div style={ui.muted}>)</div>

            <Control
              as="textarea"
              rows={2}
              style={{ width: 220 }}
              placeholder="19:00~20:20（改行もOK）"
              value={dt.time}
              onChange={(e) => {
                updateAtPath(["datetime_parts", "time"], e.target.value);
                updateAtPath(["datetime"], rebuildDatetime({ ...dt, time: e.target.value }));
              }}
            />

            <label style={ui.badge(!!json.datetime_time_newline ? "green" : "gray")}>
              <input
                type="checkbox"
                checked={!!json.datetime_time_newline}
                onChange={(e) => updateAtPath(["datetime_time_newline"], e.target.checked)}
              />
              時間を改行表示
            </label>
          </div>

          {/* <div style={{ marginTop: 10 }}>
            <Field label="時間の表示テキスト（上書き）" help="template側で datetime_time_text を使う想定。改行はそのまま反映。">
              <Control
                as="textarea"
                rows={2}
                value={json.datetime_time_text || ""}
                placeholder="例: 19:00~20:20\n（入室 18:50〜）"
                onChange={(e) => updateAtPath(["datetime_time_text"], e.target.value)}
              />
            </Field>
          </div> */}
        </Card>

        <ChairEditor chair={json.chair} updateAtPath={updateAtPath} />
        <TalksEditor talks={json.talks} updateAtPath={updateAtPath} />

        <Card title="その他">
          <Field label="取得単位">
            <Control as="textarea" rows={2} value={json.unit || ""} onChange={(e) => updateAtPath(["unit"], e.target.value)} />
          </Field>

          <Field label="主催/共催" help='例: 主催：MSD株式会社'>
            <Control
              as="textarea"
              value={json.organizer || ""}
              placeholder="主催：MSD株式会社"
              onChange={(e) => updateAtPath(["organizer"], e.target.value)}
            />
          </Field>

          <div style={ui.divider} />

          <button
            disabled={busy || hasJsonErrors}
            onClick={() => saveRender(json)}
            style={{
              ...ui.btn("primary"),
              width: "100%",
              opacity: busy || hasJsonErrors ? 0.6 : 1,
              cursor: busy || hasJsonErrors ? "not-allowed" : "pointer",
            }}
          >
            {hasJsonErrors ? "Fix JSON errors to Save" : busy ? "Rendering..." : "Save & Render"}
          </button>
        </Card>
      </div>

      {/* Right */}
      <div style={ui.rightCol}>
        <Card
          title="Preview"
          right={
            <div style={{ display: "flex", gap: 8 }}>
              <a
                href={`/download/${jobId}.jpg?t=${previewBuster}`}
               style={{
                                  padding: "10px 14px",
                                  borderRadius: 12,
                                  border: "none",
                                  background: "#2563eb",
                 color: "#fff",
                                  fontSize: 12,
                                  fontWeight: 800,
                                  cursor: "pointer",
                 boxShadow: "0 10px 24px rgba(37,99,235,0.25)",
                  textDecoration: "none",
                                }}
              >
                ⬇ Download
              </a>
              <span style={ui.badge(busy ? "blue" : "green")}>{busy ? "Rendering…" : "Ready"}</span>
            </div>
          }
        >
          <div style={styles.previewFrame}>
            <img
              src={`/preview/${jobId}.jpg?t=${previewBuster}`}
              style={{ width: "100%", maxWidth: 600, display: "block", margin: "0 auto" }}
              alt=""
            />
          </div>
        </Card>

        {/* <Card title="メモ" right={<span style={ui.badge("gray")}>tips</span>}>
          <div style={{ fontSize: 12, color: "#666", lineHeight: 1.6 }}>
            <div>• hero/talks の部分装飾は override でやる（index or target）</div>
            <div>• template側は applyOverridesToLines(...) で HTML を作って innerHTML に入れる</div>
            <div>• autoRender は 0.5秒デバウンス</div>
          </div>
        </Card> */}
      </div>
    </div>
  );
}