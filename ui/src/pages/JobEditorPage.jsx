import React, { useEffect, useRef, useState, useCallback } from "react";
import { Link, useParams } from "react-router-dom";

/**
 * ✅ この編集画面は以下を満たします
 * - 右のPreviewは sticky + スクロール枠
 * - 0.5秒デバウンスでリアルタイムレンダ（autoRender ON）
 * - /render へ { jobId, design } をPOST
 * - title_overrides / talk.title_overrides を編集可能
 * - datetime_parts と datetime の同期
 * - 必須項目未入力時に赤枠表示
 * - Save & Render時にバリデーション
 */

const API_BASE = import.meta.env.VITE_API_BASE || "";

function tryParseJson(text) {
  try {
    return { ok: true, value: JSON.parse(text) };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/** ---------- UI ---------- **/
const ui = {
  page: {
    display: "grid",
    gridTemplateColumns: "minmax(420px, 560px) 1fr",
    gap: 16,
    padding: 16,
    minHeight: "100vh",
    background: "#fafafa",
  },

  smallBtn2: {
    padding: "8px 10px",
    borderRadius: 12,
    border: "1px solid rgba(226,232,240,0.95)",
    background: "#fff",
    fontSize: 12,
    fontWeight: 900,
    cursor: "pointer",
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
  headerRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 12,
    marginBottom: 8,
  },
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
      borderWidth: "1px",
      borderStyle: "solid",
      borderColor: "#e3e3e3",
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
  controlError: {
    border: "1px solid #ef4444",
    background: "#fff7f7",
    boxShadow: "0 0 0 3px rgba(239,68,68,0.08)",
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
  errorText: {
    fontSize: 12,
    color: "#dc2626",
    fontWeight: 700,
  },
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

function Control({ as = "input", invalid = false, style, ...props }) {
  const Tag = as;
  const merged = {
    ...ui.controlBase,
    ...(as === "textarea" ? ui.textarea : {}),
    ...(invalid ? ui.controlError : null),
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

  if (!next.chair) next.chair = { role: "", name_display: "", affiliation: "", honorific_title: "先生" };
  if (!Array.isArray(next.talks)) next.talks = [];
  if (!Array.isArray(next.title_overrides)) next.title_overrides = [];

  next.talks = next.talks.map((t) => ({
    time: t?.time ?? "",
    title: t?.title ?? "",
    title_lines: Array.isArray(t?.title_lines) ? t.title_lines : [],
    speaker: t?.speaker ?? "",
    speaker_display: t?.speaker_display ?? "",
    affiliation: t?.affiliation ?? "",
    honorific_title: t?.honorific_title ?? "先生",
    title_overrides: Array.isArray(t?.title_overrides) ? t.title_overrides : [],
  }));

  return next;
}

function isBlank(v) {
  return v == null || String(v).trim() === "";
}

function hasAnyLine(lines) {
  return Array.isArray(lines) && lines.some((x) => !isBlank(x));
}

function validateJob(json) {
  const e = {};

  if (!json) return e;

  if (!hasAnyLine(json.event_title_lines)) {
    e.event_title_lines = "イベントタイトルは必須です";
  }

  if (isBlank(json.region)) {
    e.region = "VP/PH/ONC は必須です";
  }

  if (isBlank(json.datetime_parts?.year)) e.datetime_year = "年は必須です";
  if (isBlank(json.datetime_parts?.month)) e.datetime_month = "月は必須です";
  if (isBlank(json.datetime_parts?.day)) e.datetime_day = "日は必須です";
  if (isBlank(json.datetime_parts?.dow)) e.datetime_dow = "曜日は必須です";
  if (isBlank(json.datetime_parts?.time)) e.datetime_time = "時間は必須です";

  // if (isBlank(json.chair?.role)) e.chair_role = "役職は必須です";
  // if (isBlank(json.chair?.name_display)) e.chair_name_display = "名前は必須です";
  // if (isBlank(json.chair?.affiliation)) e.chair_affiliation = "所属は必須です";

  if (!Array.isArray(json.talks) || json.talks.length === 0) {
    e.talks = "講演を1件以上入力してください";
  } else {
    json.talks.forEach((t, i) => {
      // if (isBlank(t?.time)) e[`talk_${i}_time`] = "時間は必須です";
      if (!hasAnyLine(t?.title_lines)) e[`talk_${i}_title_lines`] = "タイトルは必須です";
      if (isBlank(t?.speaker_display)) e[`talk_${i}_speaker_display`] = "演者は必須です";
      if (isBlank(t?.affiliation)) e[`talk_${i}_affiliation`] = "所属は必須です";
    });
  }

  return e;
}

/** ---------- Editors ---------- **/
const ChairEditor = React.memo(function ChairEditor({ chair, updateAtPath, errors }) {
  const c = chair || {};
  return (
    <Card title={c.role || "座長 / 総合司会"}>
      <Field label="役職">
        <Control
          as="select"

          value={c.role || ""}
          onChange={(e) => updateAtPath(["chair", "role"], e.target.value)}
        >
          <option value="">-- 選択してください --</option>
          <option value="座長">座長</option>
          <option value="総合司会">総合司会</option>
        </Control>

      </Field>

      <Field label="名前">
        <Control
          value={c.name_display || ""}
          onChange={(e) => updateAtPath(["chair", "name_display"], e.target.value)}
        />
      </Field>

      <Field label="敬称">
        <Control
          as="select"
          value={c.honorific_title || ""}
          onChange={(e) => updateAtPath(["chair", "honorific_title"], e.target.value)}
        >
          <option value="">-- 選択してください --</option>
          <option value="先生">先生</option>
          <option value="様">様</option>
          <option value="さん">さん</option>
          <option value=""></option>
        </Control>
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
        <div style={{ fontWeight: 800, fontSize: 12 }}>{labelPrefix}装飾指定</div>
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

        <div style={ui.muted}>フォントサイズ</div>
        <Control
          type="number"
          value={o?.font_size ?? ""}
          placeholder="例: 28"
          onChange={(e) => set("font_size", e.target.value === "" ? null : Number(e.target.value))}
        />

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
      </div>
    </div>
  );
}

const HeroOverridesEditor = React.memo(function HeroOverridesEditor({ json, updateAtPath }) {
  const arr = Array.isArray(json.title_overrides) ? json.title_overrides : [];

  const add = () => {
    const next = [...arr, { index: null, target: "", font_size: 30, color: "" }];
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

const TalksEditor = React.memo(function TalksEditor({ talks, updateAtPath, errors }) {
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
        honorific_title: "先生",
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
    const nextOv = [...curOv, { index: null, target: "", font_size: 25, color: "" }];
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
      {errors.talks ? <div style={{ ...ui.errorText, marginTop: 8 }}>{errors.talks}</div> : null}

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
              <Control

                value={t.time || ""}
                onChange={(e) => setTalkField(idx, "time", e.target.value)}
              />

            </Field>

            <Field label="タイトル" help="改行で行分割">
              <Control
                as="textarea"
                rows={3}
                invalid={!!errors[`talk_${idx}_title_lines`]}
                value={(t.title_lines || []).join("\n")}
                onChange={(e) => setTalkField(idx, "title_lines", e.target.value.split("\n"))}
              />
              {errors[`talk_${idx}_title_lines`] ? <div style={ui.errorText}>{errors[`talk_${idx}_title_lines`]}</div> : null}
            </Field>

            <Field label="演者">
              <Control
                invalid={!!errors[`talk_${idx}_speaker_display`]}
                value={t.speaker_display || ""}
                onChange={(e) => setTalkField(idx, "speaker_display", e.target.value)}
              />
              {errors[`talk_${idx}_speaker_display`] ? (
                <div style={ui.errorText}>{errors[`talk_${idx}_speaker_display`]}</div>
              ) : null}
            </Field>

            <Field label="敬称">
              <Control
                as="select"
                value={t.honorific_title || ""}
                onChange={(e) => setTalkField(idx, "honorific_title", e.target.value)}
              >
                <option value="">-- 選択してください --</option>
                <option value="先生">先生</option>
                <option value="様">様</option>
                <option value="さん">さん</option>
                <option value=""></option>
              </Control>
            </Field>

            <Field label="所属">
              <Control
                as="textarea"
                rows={2}
                invalid={!!errors[`talk_${idx}_affiliation`]}
                value={t.affiliation || ""}
                onChange={(e) => setTalkField(idx, "affiliation", e.target.value)}
              />
              {errors[`talk_${idx}_affiliation`] ? (
                <div style={ui.errorText}>{errors[`talk_${idx}_affiliation`]}</div>
              ) : null}
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
  const [previewSrc, setPreviewSrc] = useState("");
  const [errors, setErrors] = useState({});
  const [submitTried, setSubmitTried] = useState(false);

  // realtime
  const [autoRender, setAutoRender] = useState(true);
  const debounceRef = useRef(null);
  const lastSentRef = useRef("");
  const inFlightRef = useRef(false);
  const pendingRef = useRef(false);

  useEffect(() => {
    fetch(`${API_BASE}/job/${jobId}`)
      .then((r) => r.json())
      .then((d) => {
        const j = ensureBaseDefaults(d.json || {});
        setJson(j);
      });
  }, [jobId]);

  const saveRender = async (payload, { validate = true } = {}) => {
    if (!payload) return;

    const nextErrors = validateJob(payload);
    setErrors(nextErrors);

    if (validate && Object.keys(nextErrors).length > 0) {
      setSubmitTried(true);
      return;
    }

    if (inFlightRef.current) {
      pendingRef.current = true;
      return;
    }

    const s = JSON.stringify(payload);
    if (s === lastSentRef.current) return;

    inFlightRef.current = true;
    setBusy(true);
    try {
      const r = await fetch(`${API_BASE}/render`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ jobId, design: payload }),
      });
      if (!r.ok) throw new Error("render failed");
      const d = await r.json();
      lastSentRef.current = s;
      if (d.previewDataUrl) {
        setPreviewSrc(d.previewDataUrl);
      } else {
        setPreviewSrc(`${API_BASE}/preview/${jobId}.jpg?t=${Date.now()}`);
      }
      setPreviewBuster(Date.now());
    } finally {
      setBusy(false);
      inFlightRef.current = false;

      if (pendingRef.current) {
        pendingRef.current = false;
        Promise.resolve().then(() => {
          setJson((cur) => {
            if (cur) saveRender(cur, { validate: false });
            return cur;
          });
        });
      }
    }
  };

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

  const downloadBlob = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const getFilenameFromDisposition = (disposition) => {
    if (!disposition) return null;
    const m = /filename\*=UTF-8''([^;]+)|filename="([^"]+)"/i.exec(disposition);
    const raw = m && (m[1] || m[2]) ? m[1] || m[2] : null;
    if (!raw) return null;
    try {
      return decodeURIComponent(raw);
    } catch {
      return raw;
    }
  };

  const exportOneZip = async (jobId, filename) => {
    const r = await fetch(`${API_BASE}/export/${encodeURIComponent(jobId)}.zip`);
    if (!r.ok) {
      const t = await r.text().catch(() => "");
      alert(`export failed: ${r.status}\n${t}`);
      return;
    }

    const blob = await r.blob();
    const cd = r.headers.get("Content-Disposition");
    const suggested = getFilenameFromDisposition(cd);
    downloadBlob(blob, suggested || filename);
  };

  useEffect(() => {
    if (!submitTried || !json) return;
    setErrors(validateJob(json));
  }, [json, submitTried]);

  useEffect(() => {
    if (!json || !autoRender) return;

    if (debounceRef.current) clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => {
      const s = JSON.stringify(json);
      if (s !== lastSentRef.current) {
        saveRender(json, { validate: false });
      }
    }, 500);

    return () => debounceRef.current && clearTimeout(debounceRef.current);
  }, [json, autoRender]);

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

  if (!json) return (
    <div style={ui.page}>
      <div style={ui.leftCol}>
        {/* Header skeleton */}
        <div style={ui.card}>
          <div style={{ display: "grid", gap: 10 }}>
            <div style={{ height: 14, borderRadius: 6, background: "#eee", width: 80 }} />
            <div style={{ height: 20, borderRadius: 6, background: "#e5e5e5", width: "60%" }} />
          </div>
        </div>
        {/* Basic info skeleton */}
        <div style={ui.card}>
          <div style={{ display: "grid", gap: 14 }}>
            <div style={{ height: 13, borderRadius: 6, background: "#eee", width: 100 }} />
            <div style={{ height: 40, borderRadius: 12, background: "#f3f3f3" }} />
            <div style={{ display: "grid", gridTemplateColumns: "1fr 160px", gap: 12 }}>
              <div style={{ height: 80, borderRadius: 12, background: "#f3f3f3" }} />
              <div style={{ height: 40, borderRadius: 12, background: "#f3f3f3" }} />
            </div>
          </div>
        </div>
        {/* Date skeleton */}
        <div style={ui.card}>
          <div style={{ display: "grid", gap: 14 }}>
            <div style={{ height: 13, borderRadius: 6, background: "#eee", width: 50 }} />
            <div style={{ display: "flex", gap: 10 }}>
              {[100, 80, 80, 70].map((w, i) => (
                <div key={i} style={{ height: 40, borderRadius: 12, background: "#f3f3f3", width: w }} />
              ))}
            </div>
          </div>
        </div>
        {/* Chair skeleton */}
        <div style={ui.card}>
          <div style={{ display: "grid", gap: 14 }}>
            <div style={{ height: 13, borderRadius: 6, background: "#eee", width: 120 }} />
            {[1, 2, 3].map((i) => (
              <div key={i} style={{ height: 40, borderRadius: 12, background: "#f3f3f3" }} />
            ))}
          </div>
        </div>
        {/* Talk skeleton */}
        <div style={ui.card}>
          <div style={{ display: "grid", gap: 14 }}>
            <div style={{ height: 13, borderRadius: 6, background: "#eee", width: 60 }} />
            <div style={{ border: "1px dashed #ddd", borderRadius: 12, padding: 12, display: "grid", gap: 12 }}>
              {[1, 2, 3, 4].map((i) => (
                <div key={i} style={{ height: 40, borderRadius: 12, background: "#f3f3f3" }} />
              ))}
            </div>
          </div>
        </div>
        <style>{`@keyframes shimmer { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }`}</style>
      </div>
      {/* Right preview skeleton */}
      <div style={ui.rightCol}>
        <div style={ui.card}>
          <div style={{ display: "grid", gap: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ height: 13, borderRadius: 6, background: "#eee", width: 70 }} />
            </div>
            <div style={{
              height: 400,
              borderRadius: 14,
              background: "linear-gradient(90deg, #eee 25%, #f5f5f5 50%, #eee 75%)",
              backgroundSize: "400% 100%",
              animation: "shimmer 1.2s ease-in-out infinite",
            }} />
          </div>
        </div>
      </div>
    </div>
  );

  const dt = json.datetime_parts || { year: "", month: "", day: "", dow: "", time: "" };

  const rebuildDatetime = (parts) => {
    if (!parts.year || !parts.month || !parts.day) return "";
    return `${parts.year}年${parts.month}月${parts.day}日（${parts.dow || ""}）${parts.time || ""}`;
  };

  const hasJsonErrors = Object.keys(errors).length > 0;
  const statusTone = hasJsonErrors ? "red" : busy ? "blue" : autoRender ? "green" : "gray";
  const statusText = hasJsonErrors ? "必須項目未入力" : busy ? "Rendering..." : autoRender ? "Auto" : "Manual";

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
              <div style={ui.h2}>
                編集: {json.event_id || ""}_{(json.event_title_lines || []).join("")}
              </div>
            </div>
          }
          right={<span style={ui.badge(statusTone)}>{statusText}</span>}
        >
          <div style={{ ...ui.row, marginTop: 10 }}>
            <label style={ui.badge(autoRender ? "green" : "gray")}>
              <input type="checkbox" checked={autoRender} onChange={(e) => setAutoRender(e.target.checked)} />
              リアルタイム反映（0.5s）
            </label>

          </div>
        </Card>

        <Card title="基本">
          <Field label="VP/PH/ONC">
            <Control
              as="select"
              invalid={!!errors.region}
              value={json.region || ""}
              onChange={(e) => updateAtPath(["region"], e.target.value)}
            >
              <option value="">-- 選択してください --</option>
              <option value="VP">VP</option>
              <option value="PH">PH</option>
              <option value="ONC">ONC</option>
            </Control>
            {errors.region ? <div style={ui.errorText}>{errors.region}</div> : null}
          </Field>

          <div style={ui.grid2}>
            <div>
              <Field label="イベントタイトル" help="改行で行分割">
                <Control
                  as="textarea"
                  rows={4}
                  invalid={!!errors.event_title_lines}
                  value={(json.event_title_lines || []).join("\n")}
                  onChange={(e) => updateAtPath(["event_title_lines"], e.target.value.split("\n"))}
                />
                {errors.event_title_lines ? <div style={ui.errorText}>{errors.event_title_lines}</div> : null}
              </Field>
            </div>

            <div>
              <Field label="ベース文字サイズ" help="イベントタイトル の基本サイズ（例: 30）">
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
              invalid={!!errors.datetime_year}
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
              invalid={!!errors.datetime_month}
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
              invalid={!!errors.datetime_day}
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
              invalid={!!errors.datetime_dow}
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
              invalid={!!errors.datetime_time}
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

            <Field label="注釈" help="日時の下に小さめの文字で表示される行。改行も可能。例: 各講演35分（Q&A含む）など ※VM(本社)の時のみ自動処理が実行されます。">
              <Control
                as="textarea"
                rows={2}
                style={{ width: "100%" }}
                placeholder=" 例:※各講演35分 (Q&A含む)"
                value={json.datetime_note || ""}
                onChange={(e) => updateAtPath(["datetime_note"], e.target.value)}
              />
            </Field>

            {json.datetime_note ? (
              <div style={ui.grid2}>
                <div>
                  <Field label="注釈の文字サイズ" help="注釈の基本サイズ（例: 14）">
                    <Control
                      type="number"
                      value={json.datetime_note_font_size || 14}
                      onChange={(e) => updateAtPath(["datetime_note_font_size"], Number(e.target.value))}
                    />
                  </Field>
                </div>

                {/* <div>
                  <Field label="注釈の左の余白" help="注釈の左位置の余白（例: 3）">
                    <Control
                      type="number"
                      value={json.datetime_note_left || 3}
                      onChange={(e) => updateAtPath(["datetime_note_left"], Number(e.target.value))}
                    />
                  </Field>
                </div> */}
              </div>
            ) : null}
          </div>

          {errors.datetime_year ||
            errors.datetime_month ||
            errors.datetime_day ||
            errors.datetime_dow ||
            errors.datetime_time ? (
            <div style={{ ...ui.errorText, marginTop: 10 }}>
              {[errors.datetime_year, errors.datetime_month, errors.datetime_day, errors.datetime_dow, errors.datetime_time]
                .filter(Boolean)
                .join(" / ")}
            </div>
          ) : null}
        </Card>

        <ChairEditor chair={json.chair} updateAtPath={updateAtPath} errors={errors} />
        <TalksEditor talks={json.talks} updateAtPath={updateAtPath} errors={errors} />

        <Card title="その他">
          <Field label="取得単位">
            <Control as="textarea" rows={2} value={json.unit || ""} onChange={(e) => updateAtPath(["unit"], e.target.value)} />
          </Field>

          <Field label="主催/共催" help="例: 主催：MSD株式会社">
            <Control
              as="textarea"
              value={json.organizer || ""}
              placeholder="主催：MSD株式会社"
              onChange={(e) => updateAtPath(["organizer"], e.target.value)}
            />
          </Field>

          <div style={ui.divider} />

          <button
            disabled={busy}
            onClick={() => saveRender(json)}
            style={{
              ...ui.btn("primary"),
              width: "100%",
              opacity: busy ? 0.6 : 1,
              cursor: busy ? "not-allowed" : "pointer",
            }}
          >
            {busy ? "Rendering..." : "Save & Render"}
          </button>

          {hasJsonErrors ? (
            <div style={{ ...ui.errorText, marginTop: 10 }}>
              必須項目が未入力です。赤枠の項目を入力してください。
            </div>
          ) : null}
        </Card>
      </div>

      {/* Right */}
      <div style={ui.rightCol}>
        <Card
          title="Preview"
          right={
            <div style={{ marginTop: 6, display: "flex", gap: 8, flexWrap: "wrap" }}>
              <div style={{ fontSize: 14, color: "#64748b", width: "100%", paddingTop: 8 }}>Download</div>

              <button
                style={ui.smallBtn2}
                onClick={async () => {
                  const eventIdLike = json.event_id || jobId || "event";
                  const filename = `${eventIdLike}_招聘.jpg`;
                  const url = `/download/${jobId}.jpg?t=${encodeURIComponent(previewBuster)}`;
                  await downloadWithFilename(url, filename);
                }}
              >
                JPG
              </button>

              <button
                style={ui.smallBtn2}
                onClick={async () => {
                  const eventIdLike = json.event_id || jobId || "event";
                  const filename = `${eventIdLike}_backup.json`;
                  const url = `/debug/${jobId}/latest.json?t=${encodeURIComponent(previewBuster)}`;
                  await downloadWithFilename(url, filename);
                }}
              >
                JSON
              </button>

              <button
                style={ui.smallBtn2}
                onClick={async () => {
                  const eventIdLike = json.event_id || jobId || "event";
                  const filename = `${eventIdLike}_export.zip`;
                  await exportOneZip(jobId, filename);
                }}
              >
                まとめてダウンロード
              </button>
            </div>
          }
        >
          <div style={styles.previewFrame}>
            <img
              src={previewSrc || `${API_BASE}/preview/${jobId}.jpg?t=${previewBuster}`}
              style={{ width: "100%", maxWidth: 600, display: "block", margin: "0 auto" }}
              alt=""
            />
          </div>
        </Card>
      </div>
    </div>
  );
}