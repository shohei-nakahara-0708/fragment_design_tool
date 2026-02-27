# app.py
from __future__ import annotations

import base64
import json
import os
import re
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Dict, Any, Literal, Tuple

import httpx
from dotenv import load_dotenv
from fastapi import FastAPI, UploadFile, File, HTTPException, Body,Form
from fastapi.responses import FileResponse, JSONResponse,StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pydantic import BaseModel, Field

import time
import sqlite3
import zipfile
import tempfile
from datetime import datetime, timezone, date
import traceback

import gspread
from google.oauth2.service_account import Credentials
from collections import Counter

from difflib import SequenceMatcher

import random
from gspread.exceptions import APIError
import psycopg
from psycopg.rows import dict_row

import fitz  # pymupdf

import shutil

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL", "")
API_BASE_URL = os.getenv("API_BASE_URL", "")

APP_DIR = Path(__file__).resolve().parent

def resolve_data_dir() -> Path:
    v = (os.getenv("DATA_DIR") or "").strip()
    if v:
        p = Path(v)
        try:
            p.mkdir(parents=True, exist_ok=True)
            return p
        except PermissionError:
            pass

    p = APP_DIR / "_data"
    p.mkdir(parents=True, exist_ok=True)
    return p

DATA_DIR = resolve_data_dir()

TEMPLATE_PATH = APP_DIR / "template.html"
_cached_template: Optional[str] = None

# Postgres移行したなら不要。残すならローカル専用に。
DB_PATH = DATA_DIR / "index.sqlite"

EXPORT_DIR = DATA_DIR / "_exports"
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

MAX_HEIGHT = 2000
BASE_VIEWPORT = {"width": 600, "height": 800}

AI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
AI_TIMEOUT = 30

# EMU -> pt (pptx uses EMU units for font size)
EMU_PER_PT = 12700

TIME_PAT = re.compile(r"(\d{1,2}:\d{2}\s*[～〜\-ー~]\s*\d{1,2}:\d{2})")

ORG_CANON = {
    "MSD": "MSD株式会社",
    "MSD KK": "MSD株式会社",
    "MSD K.K.": "MSD株式会社",
    "Merck": "MSD株式会社",
}

SESSION_TIME_RE = re.compile(r"[①②③④⑤⑥⑦⑧⑨⑩]?\s*(\d{1,2}[:：]\d{2}\s*[～〜\-ー~]\s*\d{1,2}[:：]\d{2})")
TYPESET_JS = r"""
({ data }) => {
  // ---- helpers: normalize ----
  const norm = (s) => String(s ?? "")
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const unifyTilde = (s) => String(s ?? "")
    .replace(/～/g, "〜")
    .trim();

  // ---- measurer element (CSS font/letter-spacing reflected) ----
  const getMeasurer = () => {
    let el = document.getElementById("__measurer__");
    if (!el) {
      el = document.createElement("div");
      el.id = "__measurer__";
      el.style.position = "fixed";
      el.style.left = "-10000px";
      el.style.top = "-10000px";
      el.style.whiteSpace = "pre";
      el.style.padding = "0";
      el.style.margin = "0";
      el.style.border = "0";
      document.body.appendChild(el);
    }
    return el;
  };

  const measure = (text, style) => {
    const el = getMeasurer();
    el.style.fontFamily = style.fontFamily;
    el.style.fontWeight = String(style.fontWeight);
    el.style.fontSize = style.fontSize;
    el.style.letterSpacing = style.letterSpacing ?? "normal";
    el.textContent = text;
    return el.scrollWidth; // px
  };

  // ---- break candidates ----
  const breakPositions = (s) => {
    const out = [];
    const breakers = new Set([" ", "、", "。",  ",", "，", ":", "：", "-", "－", "—", "–", "−"]);
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (breakers.has(ch)) out.push(i + 1);
    }
    // 助詞（ざっくり）
    const re = /(を|の|に|と|へ|や|で)/g;
    let m;
    while ((m = re.exec(s)) !== null) out.push(m.index + m[0].length);
    return Array.from(new Set(out)).sort((a, b) => a - b);
  };

  // ---- force subtitle separators to 2nd-line head ----
  // 〜 / ~ / dash（- – — − －）を「必ず次行」にする
  const splitBySubtitle = (s) => {
    s = unifyTilde(norm(s));
    if (!s) return null;

    // 1) Japanese tilde
    let idx = s.indexOf("〜");
    if (idx > 0) {
      const a = s.slice(0, idx).trimEnd();
      const b = s.slice(idx).trimStart(); // 2行目は必ず「〜」から
      if (a && b) return [a, b];
    }

    // 2) ASCII tilde
    idx = s.indexOf("~");
    if (idx > 0) {
      const a = s.slice(0, idx).trimEnd();
      const b = s.slice(idx).trimStart(); // 2行目は必ず「~」から
      if (a && b) return [a, b];
    }

    // 3) Dash separator: prefer " space-dash-space "
    const dashRe = /([\-–—−－])/;  // まずダッシュを見つける
const m = dashRe.exec(s);
if (m && m.index > 0) {
  const dashPos = m.index;
  const a = s.slice(0, dashPos).trimEnd();
  const dashChar = s[dashPos];
  const rest = s.slice(dashPos + 1).trim();
  const b = (dashChar + rest).trimStart(); // 2行目は dash から（"-当院に…"）
  if (a && b) return [a, b];
}

    return null;
  };

  // ---- wrap into <=maxLines with px constraint ----
  const wrapPx = (s, maxPx, style, maxLines, forceSubtitle2ndHead) => {
    s = unifyTilde(norm(s));
    if (!s) return [];

    // 1) subtitle rule (always newline)
    if (forceSubtitle2ndHead) {
      const sp = splitBySubtitle(s);
      if (sp) {
        const a = sp[0], b = sp[1];
        if (measure(a, style) <= maxPx && measure(b, style) <= maxPx) {
          return [a, b];
        }
        // 収まらなくても「区切りは次行」を守る：a/bをそれぞれwrapして連結
        const aLines = wrapPx(a, maxPx, style, maxLines, false);
        const bLines = wrapPx(b, maxPx, style, maxLines, false);
        return [...aLines, ...bLines].filter(Boolean).slice(0, maxLines);
      }
    }

    // 2) one line if fits
    if (measure(s, style) <= maxPx) return [s];

    // 3) try 2 lines by candidates: choose minimal slack^2
    const cand = breakPositions(s);
    let best = null;
    let bestScore = null;

    for (const p of cand) {
      const a = s.slice(0, p).trim();
      const b = s.slice(p).trim();
      if (!a || !b) continue;

      const wa = measure(a, style);
      const wb = measure(b, style);

      if (wa <= maxPx && wb <= maxPx) {
        const score = (maxPx - wa) ** 2 + (maxPx - wb) ** 2;
        if (bestScore === null || score < bestScore) {
          bestScore = score;
          best = [a, b];
        }
      }
    }
    if (best) return best;

    // 4) try 3 lines
    if (maxLines >= 3) {
      for (const p of cand) {
        const a = s.slice(0, p).trim();
        const rest = s.slice(p).trim();
        if (!a || !rest) continue;
        if (measure(a, style) > maxPx) continue;

        const cand2 = breakPositions(rest);
        let best2 = null;
        let best2Score = null;

        for (const q of cand2) {
          const b = rest.slice(0, q).trim();
          const c = rest.slice(q).trim();
          if (!b || !c) continue;

          const wb = measure(b, style);
          const wc = measure(c, style);

          if (wb <= maxPx && wc <= maxPx) {
            const score = (maxPx - wb) ** 2 + (maxPx - wc) ** 2;
            if (best2Score === null || score < best2Score) {
              best2Score = score;
              best2 = [a, b, c];
            }
          }
        }
        if (best2) return best2;
      }
    }

    // 5) last resort: force break by char
    const out = [];
    let cur = "";
    for (const ch of s) {
      const nxt = cur + ch;
      if (!cur || measure(nxt, style) <= maxPx) {
        cur = nxt;
      } else {
        out.push(cur.trim());
        cur = ch;
        if (out.length >= maxLines - 1) break;
      }
    }
    if (cur.trim()) out.push(cur.trim());
    return out.slice(0, maxLines);
  };

  // ---- compute widths from DOM ----
  const wrapEl = document.querySelector(".wrap");
  const wrapW = wrapEl ? wrapEl.clientWidth : 600;

  // タイトル領域（左padding+右padding=約28, time(74)+gap(28)=102）
  const talkMax = (wrapW - 28 - 102); // ≒470

  // talk-info領域（role-pill(41)+gap(14)+左右padding(28)）を引く
  const rolePillW = 41;
  const infoGap = 14;
  const affMax = (wrapW - 28 - rolePillW - infoGap);

  // ---- styles (match CSS) ----
  const heroStyle = { fontFamily: "Invention JP", fontWeight: 700, fontSize: "30px", letterSpacing: "normal" };
  const talkStyle = { fontFamily: "Invention JP", fontWeight: 700, fontSize: "25px", letterSpacing: "normal" };
  const affStyle  = { fontFamily: "Invention JP", fontWeight: 400, fontSize: "14px", letterSpacing: "normal" };

  // ---- event title ----
  const evBase = norm((data.event_title_lines?.length ? data.event_title_lines.join(" ") : data.event_title) ?? "");
  const evLines = wrapPx(evBase, wrapW, heroStyle, 5, false); // ★subtitle強制ON
  if (evLines.length) {
    data.event_title_lines = evLines;
    data.event_title = evLines.join("\n");
  }

  // ---- talks title lines + affiliation wrap ----
  if (Array.isArray(data.talks)) {
    data.talks = data.talks.map((t) => {
      const title = t?.title ?? "";
      const title_lines = wrapPx(title, talkMax, talkStyle, 5, true); // ★subtitle強制ON

      const affRaw = String(t?.affiliation ?? "");
      const paras = affRaw.split("\n").map(x => norm(x)).filter(Boolean);
      const affLines = [];
      const baseParas = paras.length ? paras : [norm(affRaw)];
      for (const para of baseParas) {
        if (!para) continue;
        // affiliationは最大2行くらいにしたいなら maxLines=2 に変えてOK
        const lines = wrapPx(para, affMax, affStyle, 5, true);
        for (const ln of lines) affLines.push(ln);
      }
      const affiliation = affLines.join("\n");

      return { ...t, title_lines, affiliation };
    });
  }

  return data;
}
"""

# ---------------- Models ----------------

class DatetimeParts(BaseModel):
    year: str = ""
    month: str = ""
    day: str = ""
    dow: str = ""      # "月" "火" ...
    time: str = ""     # "19:00~20:20"

class TextOverride(BaseModel):
    # どれで対象を特定するか（どれか1つ使えればOK）
    target: Optional[str] = ""    # 対象テキスト（完全一致/部分一致に使う）
    index: Optional[int] = None   # title_lines の行番号

    # 変更したい見た目
    font_size: Optional[int] = None
    font_weight: Optional[int] = None
    color: Optional[str] = None


    
class Talk(BaseModel):
    time: str = ""
    title: str = "" 
    title_lines: List[str] = Field(default_factory=list)  # 改行保持 + ~...~ は別行
    speaker: str = ""
    speaker_display: str = ""
    affiliation: str = ""
    title_overrides: List[TextOverride] = Field(default_factory=list)


class Chair(BaseModel):
    name: str = ""
    name_display: str = ""
    affiliation: str = ""


class DesignJSON(BaseModel):
    event_title_lines: List[str] = Field(default_factory=list)  # 改行保持 + ~...~ は別行
    event_title: str = ""  # 互換/テンプレ移行用（event_title_linesを \n で結合）
    title_overrides: List[TextOverride] = Field(default_factory=list)
    datetime: str = ""
    datetime_parts: Optional[DatetimeParts] = None
    datetime_time_newline: bool = False  # datetime_parts.time を改行するか
    organizer: str = ""
    chair: Chair = Chair()
    talks: List[Talk] = Field(default_factory=list)
    warnings: List[str] = Field(default_factory=list)
    confidence: float = 0.0
    manual_override: bool = False
    note: str = ""
    locked: bool = False
    title_font_size: int = 30
    region: str = ""  # 追加: 地域
    unit: str = ""    # 追加: 取得単位
    event_id: str = "" # 追加: イベントID


class RenderReq(BaseModel):
    jobId: str
    design: DesignJSON

@dataclass
class TimeCand:
    text: str
    top: int
    left: int


# ---------------- Utilities ----------------
BREAK_CHARS = ["／", "/", "・", " ", "　", "～", "~", "-", "－", "—", "–", "（", "(", "）", ")", "、", "。", ":", "："]
DROP_AT_BREAK = set(["-", "－", "—", "–"])

EMU_PER_PT = 12700

TIME_RE = re.compile(r"(\d{1,2}\s*:\s*\d{2})\s*[~〜～\-–—]\s*(\d{1,2}\s*:\s*\d{2})")

def _norm_time(s: str) -> str:
    s = str(s or "").replace("\u3000", " ")
    s = re.sub(r"\s+", "", s)
    s = s.translate(str.maketrans("０１２３４５６７８９：", "0123456789:"))
    s = s.replace("～", "~").replace("〜", "~")
    s = re.sub(r"[–—−－\-]", "-", s)
    return s

def extract_time_cands_with_pos(blocks: list[TextBlock]) -> list[TimeCand]:
    out = []
    for b in blocks:
        t = _norm_time(b.text)
        m = TIME_RE.search(t.replace("-", "~"))  # ダッシュも許容してまとめる
        if not m:
            continue
        a, c = m.group(1), m.group(2)
        out.append(TimeCand(text=f"{a}~{c}", top=b.top, left=b.left))
    out.sort(key=lambda x: (x.top, x.left))
    return out

def parse_event_time_range(dt_text: str) -> tuple[str, str] | None:
    # "2026年4月18日（土）14:00～16:00" みたいなのから取る
    t = _norm_time(dt_text)
    m = TIME_RE.search(t.replace("-", "~"))
    if not m:
        return None
    return (m.group(1), m.group(2))

def time_to_minutes(hhmm: str) -> int | None:
    m = re.match(r"^(\d{1,2}):(\d{2})$", hhmm)
    if not m:
        return None
    h = int(m.group(1)); mi = int(m.group(2))
    return h * 60 + mi

def is_within_event(tc: TimeCand, ev: tuple[str,str] | None) -> bool:
    if not ev:
        return True
    s, e = ev
    s0 = time_to_minutes(s); e0 = time_to_minutes(e)
    m = TIME_RE.search(tc.text)
    if not m or s0 is None or e0 is None:
        return True
    a = time_to_minutes(m.group(1)); b = time_to_minutes(m.group(2))
    if a is None or b is None:
        return True
    # 多少の誤差許容
    return (s0 - 10) <= a and b <= (e0 + 10)

def assign_talk_times_by_proximity(blocks: list[TextBlock], payload: DesignJSON) -> DesignJSON:
    # 1) イベント全体の時間枠（VM由来のdatetimeでも、blocks由来でもOK）
    ev = parse_event_time_range(getattr(payload, "datetime", "") or "")

    # 2) time候補を抽出＆イベント範囲でフィルタ
    cands = [c for c in extract_time_cands_with_pos(blocks) if is_within_event(c, ev)]
    if not cands or not getattr(payload, "talks", None):
        return payload

    # 3) talkアンカー(top)を作る：ここは既存の「talkを当てる」ロジックがあるならそれを使ってOK
    #    最低限、title_lines/title/speakerが含まれるブロックのtopを探す（雑でも効く）
    def find_anchor_top(talk) -> int:
        keys = []
        if getattr(talk, "title", ""): keys += [talk.title]
        if getattr(talk, "speaker_display", ""): keys += [talk.speaker_display]
        if getattr(talk, "speaker", ""): keys += [talk.speaker]
        keys = [k for k in keys if k]
        if not keys:
            return 10**18
        best_top = 10**18
        best_score = 0
        for b in blocks:
            bt = b.text or ""
            score = 0
            for k in keys:
                kk = k.replace(" ", "")
                if kk and kk in bt.replace(" ", ""):
                    score += 2
            if score > best_score:
                best_score = score
                best_top = b.top
        return best_top

    talks = list(payload.talks or [])
    talk_infos = []
    for idx, t in enumerate(talks):
        talk_infos.append((idx, find_anchor_top(t)))
    # 上→下に並べる
    talk_infos.sort(key=lambda x: x[1])

    used = set()

    for idx, anchor_top in talk_infos:
        t = talks[idx]
        if _norm_time(getattr(t, "time", "")):
            continue

        # ルール：アンカーより上側の候補のうち、最も近いものを優先（未使用）
        best = None
        best_dist = None

        for ci, c in enumerate(cands):
            if ci in used:
                continue
            # timeはアンカーより少し上にあることが多いので、上側は距離を軽く優遇
            dist = abs(anchor_top - c.top)
            if c.top <= anchor_top:
                dist *= 0.7
            # leftも少しだけ効かせたいならここで加点（time列は左寄り）
            if best_dist is None or dist < best_dist:
                best_dist = dist
                best = (ci, c)

        if best:
            ci, c = best
            used.add(ci)
            t.time = c.text

    payload.talks = talks
    return payload

def extract_blocks_from_pdf(pdf_path: Path, first_page_only: bool = True) -> List[TextBlock]:
    doc = fitz.open(str(pdf_path))
    blocks: List[TextBlock] = []

    pages = [doc[0]] if (first_page_only and doc.page_count > 0) else [doc[i] for i in range(doc.page_count)]

    for page in pages:
        d = page.get_text("dict")

        for b in d.get("blocks", []):
            if b.get("type") != 0:  # 0=text
                continue

            x0, y0, x1, y1 = b.get("bbox", (0, 0, 0, 0))

            # 段落/行を組み立て（PDFのline/spansを尊重）
            lines = []
            max_font_pt = 0.0

            for line in b.get("lines", []):
                spans = line.get("spans", [])
                # spanをそのまま連結（余計な空白は後でnormalize）
                t = "".join(s.get("text", "") for s in spans)
                if t and t.strip():
                    lines.append(t.strip())

                for s in spans:
                    try:
                        max_font_pt = max(max_font_pt, float(s.get("size") or 0.0))
                    except Exception:
                        pass

            text = normalize_keep_newlines("\n".join(lines))
            if not text:
                continue

            # PDF(pt) -> EMU に変換
            left_emu   = int(round(x0 * EMU_PER_PT))
            top_emu    = int(round(y0 * EMU_PER_PT))
            width_emu  = int(round((x1 - x0) * EMU_PER_PT))
            height_emu = int(round((y1 - y0) * EMU_PER_PT))

            blocks.append(
                TextBlock(
                    text=text,
                    left=left_emu,
                    top=top_emu,
                    width=width_emu,
                    height=height_emu,
                    max_font_pt=float(max_font_pt or 0.0),
                )
            )

    doc.close()
    blocks.sort(key=lambda b: (b.top, b.left))
    return blocks

def merge_event_title_blocks_strict(blocks: list[TextBlock]) -> list[TextBlock]:
    # 上部の大フォントだけ抽出
    candidates = [b for b in blocks if b.max_font_pt >= 22]
    if not candidates:
        return blocks

    candidates.sort(key=lambda b: b.top)

    # 上から連続しているものだけ取る
    merged_group = [candidates[0]]

    for b in candidates[1:]:
        # 前のブロックと縦距離が近ければ同グループ
        if abs(b.top - merged_group[-1].top) < 500000:
            merged_group.append(b)
        else:
            break  # 離れたら終了

    merged_text = "\n".join(b.text for b in merged_group)

    merged_block = TextBlock(
        text=merged_text,
        left=min(b.left for b in merged_group),
        top=min(b.top for b in merged_group),
        width=max(b.width for b in merged_group),
        height=sum(b.height for b in merged_group),
        max_font_pt=max(b.max_font_pt for b in merged_group),
    )

    # 元ブロックから除去
    new_blocks = [b for b in blocks if b not in merged_group]
    new_blocks.append(merged_block)
    new_blocks.sort(key=lambda b: (b.top, b.left))

    return new_blocks

def merge_pdfish_blocks(blocks):
    blocks = sorted(blocks, key=lambda b: (b["top"], b["left"]))
    merged = []

    for b in blocks:
        if not merged:
            merged.append(b)
            continue

        prev = merged[-1]

        # 縦距離が近くて左が近いなら同じ行扱い
        if abs(b["top"] - prev["top"]) < 6 and abs(b["left"] - prev["left"]) < 10:
            prev["text"] += b["text"]
            prev["width"] = max(prev["width"], b["width"])
        else:
            merged.append(b)

    return merged

def extract_blocks_any(path: Path, first_only: bool = True) -> List[TextBlock]:
    suf = path.suffix.lower()
    if suf == ".pdf":
        blocks = extract_blocks_from_pdf(path, first_page_only=first_only)
        # PDFは分断が多いので前処理（後で強化できる）
        # blocks = merge_pdfish_blocks(blocks)
        return blocks
    # .ppt は python-pptx では基本ダメなのでここでは弾くか、事前変換前提
    return extract_blocks_from_pptx(path, first_slide_only=first_only)

def pget(obj, key, default=None):
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)

def pset(obj, key, value):
    if isinstance(obj, dict):
        obj[key] = value
        return
    setattr(obj, key, value)

def extract_session_times(s: str) -> list[str]:
    s0 = normalize_datetime_text(s)  # ここで全角コロン等も正規化される :contentReference[oaicite:8]{index=8}
    out = []
    for m in SESSION_TIME_RE.finditer(s0):
        t = normalize_datetime_text(m.group(1)).replace(" ", "")
        if t not in out:
            out.append(t)
    return out

def split_tilde_subtitle(s: str) -> list[str]:
    """
    末尾の ～xxx～ / ~xxx~ を別行に分離（あれば）
    """
    s = normalize_space(s)
    if not s:
        return []
    m = re.search(r"(.*?)(\s*[~～].+[~～]\s*)$", s)
    if m:
        a = normalize_space(m.group(1))
        b = normalize_space(m.group(2))
        return [x for x in [a, b] if x]
    return [s]

def wrap_by_chars(s: str, max_len: int, *, back: int = 8) -> list[str]:
    """
    文字数近似で自然改行（直近の区切り文字を優先）
    """
    s = normalize_space(s)
    if not s:
        return []
    if len(s) <= max_len:
        return [s]

    out = []
    rest = s
    while rest and len(rest) > max_len:
        cut = -1
        start = max(0, max_len - back)
        for i in range(min(max_len, len(rest) - 1), start - 1, -1):
            if rest[i] in BREAK_CHARS:
                cut = i + 1
                break
        if cut == -1:
            cut = max_len

        head = rest[:cut].strip()
        tail = rest[cut:].strip()

        # ハイフンなど “落としたい区切り” を行末/行頭に残さない
        if head and head[-1] in DROP_AT_BREAK:
            head = head[:-1].rstrip()
        if tail and tail[0] in DROP_AT_BREAK:
            tail = tail[1:].lstrip()

        if head:
            out.append(head)
        rest = tail

    if rest:
        out.append(rest)

    return [x for x in out if x]

def join_short_suffix(lines: list[str]) -> list[str]:
    """
    “るために” みたいな短い尻尾行ができたら、前行に戻して繋げる（軽い補正）
    """
    if not lines:
        return lines
    out = [lines[0]]
    for l in lines[1:]:
        if len(l) <= 4 and out:
            out[-1] = (out[-1] + l).strip()
        else:
            out.append(l)
    return out


MEASURE_JS = """
({ text, font }) => {
  const c = document.createElement('canvas');
  const ctx = c.getContext('2d');
  ctx.font = font;
  return ctx.measureText(text).width;
}
"""

async def measure_px(page, text: str, font_css: str) -> float:
    return float(await page.evaluate(MEASURE_JS, {"text": text, "font": font_css}))

def split_tilde_head_2nd(s: str):
    s = (s or "").replace("～", "〜").strip()
    i = s.find("〜")
    if i <= 0:
        return None
    a = s[:i].rstrip()
    b = s[i:].lstrip()  # 2行目は必ず「〜」から
    if not a or not b:
        return None
    return a, b

BREAK_CHARS_PX = [" ", "　", "、", "。", "・", "／", "/", ":", "：", "）", ")", "】", "]"]

def candidate_breaks(s: str) -> list[int]:
    pos = []
    for i, ch in enumerate(s):
        if ch in BREAK_CHARS_PX:
            pos.append(i + 1)  # その直後で折る
    return sorted(set(pos))

async def wrap_px(page, text: str, max_px: int, font_css: str, max_lines: int = 3, force_tilde: bool = False) -> list[str]:
    s = (text or "").replace("\n", " ").strip()
    if not s:
        return []

    # 〜強制（talk用）
    if force_tilde:
        sp = split_tilde_head_2nd(s)
        if sp:
            a, b = sp
            if await measure_px(page, a, font_css) <= max_px and await measure_px(page, b, font_css) <= max_px:
                return [a, b]

    # 1行で入るなら1行
    if await measure_px(page, s, font_css) <= max_px:
        return [s]

    # 2行以上：候補位置で分割して、収まりつつ “余白が少ない” ものを選ぶ
    breaks = candidate_breaks(s)
    best = None
    best_score = None

    # まずは2行を狙う（ダメなら後段で3行）
    for p in breaks:
        a = s[:p].strip()
        b = s[p:].strip()
        if not a or not b:
            continue
        wa = await measure_px(page, a, font_css)
        wb = await measure_px(page, b, font_css)
        if wa <= max_px and wb <= max_px:
            score = (max_px - wa) ** 2 + (max_px - wb) ** 2
            if best_score is None or score < best_score:
                best_score = score
                best = [a, b]
    if best:
        return best

    # 3行まで許す：2回折る（粗いが強い）
    if max_lines >= 3:
        for p in breaks:
            a = s[:p].strip()
            rest = s[p:].strip()
            if not a or not rest:
                continue
            if await measure_px(page, a, font_css) > max_px:
                continue
            # rest を2行にする
            breaks2 = candidate_breaks(rest)
            for q in breaks2:
                b = rest[:q].strip()
                c = rest[q:].strip()
                if not b or not c:
                    continue
                wb = await measure_px(page, b, font_css)
                wc = await measure_px(page, c, font_css)
                if wb <= max_px and wc <= max_px:
                    return [a, b, c]

    # 最後の保険：強制分割（絶対はみ出さない）
    out = []
    cur = ""
    for ch in s:
        nxt = cur + ch
        if not cur or await measure_px(page, nxt, font_css) <= max_px:
            cur = nxt
        else:
            out.append(cur.strip())
            cur = ch
            if len(out) >= max_lines - 1:
                break
    if cur.strip():
        out.append(cur.strip())
    return out[:max_lines]

async def apply_precise_typeset_initial(payload: DesignJSON, page=None) -> DesignJSON:
    # ---- 初期値のみ：編集・ロック・上書き指定がある場合は何もしない ----
    if getattr(payload, "manual_override", False):
        return payload
    if getattr(payload, "locked", False):
        return payload
    if (getattr(payload, "title_overrides", None) or []):
        return payload
    for t in (getattr(payload, "talks", None) or []):
        ov = (t.get("title_overrides") if isinstance(t, dict) else getattr(t, "title_overrides", None)) or []
        if ov:
            return payload

    # ---- payload -> dict ----
    data_json = payload.model_dump_json() if hasattr(payload, "model_dump_json") else payload.json(ensure_ascii=False)
    data_obj = json.loads(data_json)

    async def _run(pg):
        global _cached_template
        if _cached_template is None:
            _cached_template = TEMPLATE_PATH.read_text(encoding="utf-8")

        await pg.set_content(_cached_template, wait_until="domcontentloaded")
        await pg.evaluate("() => document.fonts && document.fonts.ready")

        # TYPESET_JS: data_obj を px実測で event_title_lines/title_lines/affiliation に整形して返す
        return await pg.evaluate(TYPESET_JS, {"data": data_obj})

    # ---- page が無ければここで起動 ----
    if page is None:
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            pg = await browser.new_page(viewport=BASE_VIEWPORT)
            try:
                new_obj = await _run(pg)
            finally:
                await browser.close()
    else:
        new_obj = await _run(page)

    # ---- dict -> payload に戻す ----
    if hasattr(payload.__class__, "model_validate"):
        payload = payload.__class__.model_validate(new_obj)
    else:
        payload = payload.__class__.parse_obj(new_obj)

    # ---- template の参照差でズレないように同期（重要）----
    if getattr(payload, "event_title_lines", None):
        payload.event_title = "\n".join(payload.event_title_lines)

    for t in (getattr(payload, "talks", None) or []):
        if getattr(t, "title_lines", None):
            t.title = "\n".join(t.title_lines)

    # payload.typeset_done = True
    return payload



def format_title_initial(
    raw: str,
    *,
    max_len: int,
    max_lines: int = 3,
    force_tilde_second_line: bool = False,
) -> list[str]:
    """
    talk.title / event_title の初期整形
    """
    raw = normalize_space(raw)
    if not raw:
        return []

    # 表記ゆれ統一
    raw = raw.replace("～", "〜")

    # ★talk用：最初の「〜」を必ず2行目先頭へ
    # 例) 「…治療〜GLP-1…〜」→
    #   1行目「…治療」
    #   2行目「〜GLP-1…〜」
    if force_tilde_second_line and "〜" in raw:
        idx = raw.find("〜")
        if idx > 0:
            first = raw[:idx].rstrip()
            second = raw[idx:].lstrip()  # 2行目は必ず「〜」から
            lines = []
            lines.extend(wrap_by_chars(first, max_len=max_len))
            lines.extend(wrap_by_chars(second, max_len=max_len))
            lines = [x for x in lines if x]
            lines = join_short_suffix(lines)
            if max_lines and len(lines) > max_lines:
                lines = lines[: max_lines - 1] + [" ".join(lines[max_lines - 1 :]).strip()]
            return lines

    # 既存ロジック：末尾の ~xxx~ / ～xxx～ は別行に分離
    lines: list[str] = []
    for part in split_tilde_subtitle(raw):
        lines.extend(wrap_by_chars(part, max_len=max_len))

    lines = [x for x in lines if x]
    lines = join_short_suffix(lines)

    if max_lines and len(lines) > max_lines:
        lines = lines[: max_lines - 1] + [" ".join(lines[max_lines - 1 :]).strip()]
    return lines


def format_affiliation_initial(raw: str, *, max_len: int, max_lines: int = 2) -> str:
    """
    affiliation 初期整形（最大2行）
    """
    raw = normalize_space(raw)
    if not raw:
        return ""
    lines = wrap_by_chars(raw, max_len=max_len)
    return "\n".join(lines[:max_lines]).strip()

def post_format_design_initial(payload):
    """
    manual_override=False の “初期生成” だけ適用する想定
    """
    # event_title（600px ≒ 30px太字 → だいたい 20〜22文字/行）
    if getattr(payload, "event_title_lines", None):
        base = " ".join(payload.event_title_lines).strip()  # ← 改行をスペースで結合
    else:
        base = getattr(payload, "event_title", "") or ""

    # ★ まず1行に戻せるか試す
    one_line = normalize_space(base)

    if len(one_line) <= 24:   # 600px相当（少し余裕を持たせる）
        payload.event_title_lines = [one_line]
    else:
        payload.event_title_lines = format_title_initial(
            one_line,
            max_len=22,
            max_lines=3,
            force_tilde_second_line=False
        )

    payload.event_title = "\n".join(payload.event_title_lines).strip()

    # talks（470px）
    for t in getattr(payload, "talks", []) or []:
        # title: 25px太字 → だいたい 17〜19文字/行
        raw_title = (t.title or "")
        t.title_lines = format_title_initial(
    raw_title,
    max_len=18,
    max_lines=3,
    force_tilde_second_line=True   # ←ここ重要
)

        # affiliation: 16px太字 → だいたい 26〜30文字/行（最大2行）
        t.affiliation = format_affiliation_initial(t.affiliation or "", max_len=28, max_lines=2)

    return payload

def _get_field(payload, key: str, default=None):
    # dict
    if isinstance(payload, dict):
        return payload.get(key, default)
    # pydantic model
    return getattr(payload, key, default)

def _set_field(payload, key: str, value):
    if isinstance(payload, dict):
        payload[key] = value
    else:
        setattr(payload, key, value)

CIRCLED_NUM = "①②③④⑤⑥⑦⑧⑨⑩"

def normalize_time_range(s: str) -> str:
    s = str(s or "")
    # 全角コロン→半角
    s = s.replace("：", ":")
    # 波線統一
    s = re.sub(r"[～〜\-ー]", "~", s)
    # ★空白は全部消さない。~の前後だけ詰める
    s = re.sub(r"\s*~\s*", "~", s)
    s = s.strip()
    return s

def extract_session_times_from_blocks(blocks) -> list[str]:
    if not blocks:
        return []
    # ① 12：20～12：50 / ② 13：00～13：30 のようなブロックを拾う
    buf = []
    for b in blocks:
        t = (b.get("text") if isinstance(b, dict) else getattr(b, "text", "")) or ""
        if any(c in t for c in CIRCLED_NUM) and ("～" in t or "〜" in t or "~" in t or "：" in t or ":" in t):
            buf.append(str(t))
    # 既存の SESSION_TIME_RE を流用（dtに対して使ってたのと同じ）
    joined = "\n".join(buf)
    out = []
    for m in SESSION_TIME_RE.finditer(joined):
        tt = normalize_time_range(m.group(1))
        if tt and tt not in out:
            out.append(tt)
    return out

def extract_session_times_from_datetime(dt: str) -> list[str]:
    dt = str(dt or "")
    # ★①②が無いならセッション抽出しない（ここが重要）
    if not any(c in dt for c in CIRCLED_NUM):
        return []

    out = []
    for m in SESSION_TIME_RE.finditer(dt):
        t = normalize_time_range(m.group(1))
        if t and t not in out:
            out.append(t)
    return out

def should_hide_talk_times(payload) -> bool:
    dt = str(_get_field(payload, "datetime", "") or "")
    parts = _get_field(payload, "datetime_parts", None)
    time_str = ""
    if parts:
        time_str = parts.get("time", "") if isinstance(parts, dict) else getattr(parts, "time", "")

    if any(c in dt for c in CIRCLED_NUM):
        return True
    if "1回目" in time_str or "2回目" in time_str:
        return True
    return False

def clear_talk_times(payload):
    talks = _get_field(payload, "talks", []) or []
    for t in talks:
        if isinstance(t, dict):
            t["time"] = ""
        else:
            setattr(t, "time", "")
    _set_field(payload, "talks", talks)
    return payload

def _ensure_datetime_parts(parts):
    if parts is None:
        return DatetimeParts(year="", month="", day="", dow="", time="")
    if isinstance(parts, DatetimeParts):
        return parts
    if isinstance(parts, dict):
        return DatetimeParts(**parts)
    return DatetimeParts(
        year=getattr(parts, "year", "") or "",
        month=getattr(parts, "month", "") or "",
        day=getattr(parts, "day", "") or "",
        dow=getattr(parts, "dow", "") or "",
        time=getattr(parts, "time", "") or "",
    )

def fill_datetime_parts(payload, blocks=None):
    def pget(obj, key, default=None):
        if isinstance(obj, dict):
            return obj.get(key, default)
        return getattr(obj, key, default)

    def pset(obj, key, value):
        if isinstance(obj, dict):
            obj[key] = value
        else:
            setattr(obj, key, value)

    dt = str(pget(payload, "datetime", "") or "")

    session_times = extract_session_times_from_datetime(dt)
    if not session_times and blocks:
        session_times = extract_session_times_from_blocks(blocks)

    talks = pget(payload, "talks", []) or []
    if session_times and len(session_times) == 1 and talks:
        first = talks[0] or {}
        t2_raw = first.get("time") if isinstance(first, dict) else getattr(first, "time", "")
        t2 = normalize_time_range(t2_raw or "")
        if t2 and t2 not in session_times:
            session_times.append(t2)

    m = re.search(
        r"(?P<y>\d{4})\s*年\s*(?P<mo>\d{1,2})\s*月\s*(?P<d>\d{1,2})\s*日"
        r"(?:\s*[（(]\s*(?P<dow>[^）)\s]+)\s*[）)])?",
        normalize_datetime_text(dt)
    )

    # ★必ず DatetimeParts に統一
    parts = _ensure_datetime_parts(pget(payload, "datetime_parts", None))

    if m:
        y, mo, d, dow = (m.group("y") or "", m.group("mo") or "", m.group("d") or "", m.group("dow") or "")
        parts.year, parts.month, parts.day, parts.dow = y, mo, d, dow

    if session_times:
        if len(session_times) == 1:
            time_joined = session_times[0]
        else:
            time_joined = ", ".join([f"{i+1}回目{t}" for i, t in enumerate(session_times)])
        newline = (len(session_times) >= 2)
    else:
        mt = TIME_RANGE_RE.search(normalize_datetime_text(dt))
        time_joined = normalize_time_range(mt.group(0)) if mt else ""
        newline = False

    parts.time = time_joined

    pset(payload, "datetime_parts", parts)
    pset(payload, "datetime_time_newline", newline)

    if should_hide_talk_times(payload):
        clear_talk_times(payload)

    return payload

def _norm1(s: str) -> str:
    return normalize_space(s).replace("\n", " ").strip()

_TIME_RE = re.compile(
    r"""^\s*
    (?:\d{1,2}\s*[:：]\s*\d{2})      # 19:05
    \s*(?:[〜～~\-－–—]\s*)          # ～ / 〜 / ~ / - 系
    (?:\d{1,2}\s*[:：]\s*\d{2})      # 19:35
    \s*$""",
    re.VERBOSE
)

def looks_like_affil_line(s: str) -> bool:
    s = _norm1(s)
    if not s:
        return False

    # ラベル/案内っぽいのは除外
    if any(k in s for k in ["演題", "演者", "座長", "日時", "会場", "共催", "主催", "提供", "企画", "運営", "詳細は"]):
        return False
    
    if _TIME_RE.match(s):
        return False

    # 敬称入りは名前行の可能性
    if "先生" in s:
        return False

    # ★演題っぽい記号がある行は基本除外（所属に入ることは稀）
    #   - ダッシュ/波線/カギカッコ があると演題率が高い
    if any(p in s for p in ["「", "」", "～", "—", "－", "–"]):
        return False
    # ハイフンは施設名にも混ざり得るので「両側に空白がない長文」の時だけ除外
    if "-" in s and len(s) >= 18:
        return False

    # ★所属として強い接頭辞（法人格など）
    entity_prefix = ("医療法人", "一般財団法人", "一般社団法人", "公益財団法人", "公益社団法人",
                     "独立行政法人", "国立病院機構", "学校法人")
    if s.startswith(entity_prefix):
        return True

    # ★施設キーワード（これが無いなら所属扱いしない）
    facility_kw = ["病院", "クリニック", "医院", "診療所", "大学", "機構", "センター", "医師会", "総合病院"]
    has_facility = any(k in s for k in facility_kw)

    # 役職・部署っぽい語（施設名と一緒に出ると所属確度UP）
    role_kw = ["内科", "外科", "科", "部", "室", "教授", "准教授", "講師", "主任", "部長", "院長", "理事長"]
    has_role = any(k in s for k in role_kw)

    # ★施設名があるなら所属。施設名がなく役職だけはNG。
    if has_facility:
        return True

    # 施設名が無い場合でも「県立中央」など施設っぽい固有パターンを少し救う（任意）
    if ("県立" in s or "市立" in s) and has_role:
        return True

    return False



def wrap_vm_rows_for_rank(vm_rows: list[dict]) -> list[dict]:
    out = []
    for r in (vm_rows or []):
        # すでに {"data": ...} 形式ならそのまま
        if isinstance(r, dict) and "data" in r and isinstance(r["data"], dict):
            out.append(r)
            continue

        # rowdict -> {"data": rowdict} に包む
        if isinstance(r, dict):
            out.append({"data": r})
            continue

    return out

def _key(v: Any) -> str:
    return str(v or "").strip()

def _col_to_a1(col_idx_1based: int) -> str:
    s = ""
    n = col_idx_1based
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _find_col(headers: List[str], name: str) -> int:
    for i, h in enumerate(headers):
        if h == name:
            return i + 1
    raise KeyError(f"column not found: {name}")

def _rowdict(headers: List[str], row: List[Any], *, sheet: str, rownum: int) -> Dict[str, Any]:
    if len(row) < len(headers):
        row = list(row) + [""] * (len(headers) - len(row))
    row = row[:len(headers)]
    d = {headers[i]: row[i] for i in range(len(headers))}
    d["_sheet"] = sheet
    d["_row"] = rownum
    return d

def _extract_1col(vr: dict) -> List[Any]:
    """
    values_batch_get の valueRanges から「1列データ」を取り出す（空行/欠損に強い）
    返り値は [v0, v1, ...]
    """
    vals = vr.get("values") or []
    if not vals:
        return []

    # ケース1: [["a","b","c"]] (1行に横並びで入ってくる)
    if len(vals) == 1 and isinstance(vals[0], list):
        return vals[0]

    # ケース2: [["a"], ["b"], [], ["c"]] (縦で返る＋空行が混ざる)
    if isinstance(vals[0], list):
        out: List[Any] = []
        for r in vals:
            if not r:          # [] をスキップ（または "" を入れるのでもOK）
                out.append("") # ← 行番号ズレ防止のため空を入れるのがおすすめ
                continue
            out.append(r[0])
        return out

    # ケース3: まれに ["a","b"] のような形（そのまま）
    return vals

def _pad_list(xs: List[Any], need_len: int) -> List[Any]:
    if len(xs) < need_len:
        xs = list(xs) + [""] * (need_len - len(xs))
    return xs

def pick_last_rownum(rownums: List[int]) -> int:
    return max(rownums) if rownums else 0

def _build_id_index_from_column(values: List[Any], *, start_row: int) -> Dict[str, List[int]]:
    idx: Dict[str, List[int]] = {}
    for offset, v in enumerate(values):
        k = str(v or "").strip()
        if not k:
            continue
        rownum = start_row + offset
        idx.setdefault(k, []).append(rownum)
    return idx


def _retry_gspread(fn, *, tries=6, base=0.5, jitter=0.2):
    last = None
    for i in range(tries):
        try:
            return fn()
        except APIError as e:
            last = e
            # INTERNAL(500) とかのときだけリトライ
            msg = str(getattr(e, "response", "")) + " " + str(e)
            if "INTERNAL" not in msg and "'code': 500" not in msg:
                raise
            sleep = base * (2 ** i) + random.uniform(0, jitter)
            time.sleep(sleep)
    raise last

def batch_fetch_system_and_vm_rows(
    workbook,
    *,
    ws_map,  # 
    event_ids: List[str],
    presence_sheets: List[str],
    presence_header_row: int,
    presence_id_col: str,
    vm_sheet: str,
    vm_header_row: int,
    vm_id_col_candidates: List[str],
    col_end: str = "N",
    chunk_size: int = 200,
) -> Tuple[Dict[str, List[Dict]], Dict[str, List[Dict]], str]:

    # ---- headers: 1回ずつ ----
    presence_headers_by_sheet: Dict[str, List[str]] = {}
    presence_id_col_letter_by_sheet: Dict[str, str] = {}

    presence_col_end_by_sheet: Dict[str, str] = {}

    for s in presence_sheets:
        ws = ws_map[s]
        headers = make_unique(ws.row_values(presence_header_row))
        presence_headers_by_sheet[s] = headers

        id_col_idx = _find_col(headers, presence_id_col)
        presence_id_col_letter_by_sheet[s] = _col_to_a1(id_col_idx)

        # ★行取得で「後ろ列が空」にならないよう、ヘッダの最終列を col_end にする
        presence_col_end_by_sheet[s] = _col_to_a1(len(headers))

    ws_vm = ws_map[vm_sheet]
    vm_headers = make_unique(ws_vm.row_values(vm_header_row))

    vm_id_col_used = ""
    for c in vm_id_col_candidates:
        if c in vm_headers:
            vm_id_col_used = c
            break
    if not vm_id_col_used:
        raise KeyError(f"VM id column not found. candidates={vm_id_col_candidates}")

    vm_id_col_letter = _col_to_a1(_find_col(vm_headers, vm_id_col_used))

    # ---- batchGet: ID列だけ（まとめて）----
    ranges = []
    # presence
    for s in presence_sheets:
        col_letter = presence_id_col_letter_by_sheet[s]
        start = presence_header_row + 1
        # 「列の下全部」(末尾空はAPIが返さないことがある)
        ranges.append(f"{s}!{col_letter}{start}:{col_letter}")
    # vm
    start_vm = vm_header_row + 1
    ranges.append(f"{vm_sheet}!{vm_id_col_letter}{start_vm}:{vm_id_col_letter}")

    print(f"[INFO] batch fetching ID columns: {ranges}")
    resp = _retry_gspread(lambda: workbook.values_batch_get(ranges))

    # ---- index化（event_id -> rownums）----
    presence_index_by_sheet: Dict[str, Dict[str, List[int]]] = {}
    for i, s in enumerate(presence_sheets):
        ws = ws_map[s]
        start = presence_header_row + 1
        col_values = _extract_1col(resp["valueRanges"][i])

        # ★row_countで膨らませない。返ってきた範囲内だけで十分
        # ただし「途中に空行がある」ケースは _extract_1col が "" を入れてくれるのでズレにくい
        presence_index_by_sheet[s] = _build_id_index_from_column(col_values, start_row=start)

    # VM index
    start = vm_header_row + 1
    end_row = ws_vm.row_count

    vm_col_values = _extract_1col(resp["valueRanges"][-1])
    need_len = end_row - start + 1
    vm_col_values = _pad_list(vm_col_values, need_len)

    vm_index = _build_id_index_from_column(vm_col_values, start_row=start)

    uniq_event_ids = [str(e or "").strip() for e in event_ids if str(e or "").strip()]
    presence_rows_by_event: Dict[str, List[Dict]] = {eid: [] for eid in uniq_event_ids}
    vm_rows_by_event: Dict[str, List[Dict]] = {eid: [] for eid in uniq_event_ids}

    # ---- 必要行rangeを組み立てて batchGet（行だけ）----
    presence_row_ranges = []
    presence_row_meta = []  # (eid, sheet, rownum)
    vm_row_ranges = []
    vm_row_meta = []        # (eid, rownum)

    for eid in uniq_event_ids:
        for s in presence_sheets:
            rownums = presence_index_by_sheet.get(s, {}).get(eid, [])
            last_row = pick_last_rownum(rownums)
            if last_row:
                end_letter = presence_col_end_by_sheet[s]  # ★ここ
                presence_row_ranges.append(f"{s}!A{last_row}:{end_letter}{last_row}")
                presence_row_meta.append((eid, s, last_row))
    
    for eid in uniq_event_ids:
        # ... presence ...
        for rownum in vm_index.get(eid, []):
            vm_row_ranges.append(f"{vm_sheet}!A{rownum}:{col_end}{rownum}")
            vm_row_meta.append((eid, rownum))

    def _chunks(xs, n):
        for i in range(0, len(xs), n):
            yield xs[i:i+n]

    # presence 行取得
    for rchunk, mchunk in zip(_chunks(presence_row_ranges, chunk_size), _chunks(presence_row_meta, chunk_size)):
        if not rchunk:
            continue
        rresp = _retry_gspread(lambda: workbook.values_batch_get(rchunk))
        for vr, meta in zip(rresp["valueRanges"], mchunk):
            rowvals = vr.get("values") or [[]]
            row = rowvals[0] if rowvals else []
            eid, sheet, rownum = meta
            headers = presence_headers_by_sheet[sheet]
            presence_rows_by_event[eid].append(_rowdict(headers, row, sheet=sheet, rownum=rownum))

    # VM 行取得
    for rchunk, mchunk in zip(_chunks(vm_row_ranges, chunk_size), _chunks(vm_row_meta, chunk_size)):
        if not rchunk:
            continue
        rresp = _retry_gspread(lambda: workbook.values_batch_get(rchunk))
        for vr, meta in zip(rresp["valueRanges"], mchunk):
            rowvals = vr.get("values") or [[]]
            row = rowvals[0] if rowvals else []
            eid, rownum = meta
            vm_rows_by_event[eid].append(_rowdict(vm_headers, row, sheet=vm_sheet, rownum=rownum))
    print(vm_rows_by_event)

    return presence_rows_by_event, vm_rows_by_event, vm_id_col_used

def make_unique(headers):
    counter = Counter()
    out = []
    for h in headers:
        key = (h or "").strip()
        if key == "":
            key = "EMPTY"
        counter[key] += 1
        out.append(key if counter[key] == 1 else f"{key}_{counter[key]}_{counter[key]}")
    return out

def pick_last_presence_hit(hits: list[dict]) -> dict | None:
    if not hits:
        return None
    return max(hits, key=lambda h: int(h.get("row") or 0))

def normalize_day(v) -> str:
    """
    'YYYY-MM-DD' に正規化
    """
    if v is None:
        return ""

    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()

    s = str(v).strip()
    if not s:
        return ""

    m = re.search(r"(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})", s)
    if not m:
        return ""

    try:
        return date(int(m[1]), int(m[2]), int(m[3])).isoformat()
    except Exception:
        return ""
    
def pick_best_presence_row(rows: list[dict]) -> dict | None:
    if not rows: return None
    def filled_count(d): return sum(1 for k,v in d.items() if not k.startswith("_") and str(v).strip())
    return max(rows, key=filled_count)

def fetch_row_dict(workbook, sheet_name: str, row_num: int, *, header_row: int = 2, col_end: str = "Z") -> dict:
    """
    指定シートの指定行(row_num)を、ヘッダ(header_row)の列名でdict化して返す。
    """
    ws = workbook.worksheet(sheet_name)

    # ヘッダ
    headers = make_unique(ws.row_values(header_row))
    # 1行分（A{row}:Z{row}）
    row_values = ws.get(f"A{row_num}:{col_end}{row_num}")
    row = (row_values[0] if row_values else [])

    # 列数合わせ
    if len(row) < len(headers):
        row = row + [""] * (len(headers) - len(row))
    elif len(row) > len(headers):
        row = row[:len(headers)]

    data = {headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))}

    # 便利情報として行番号も入れる（不要なら消してOK）
    data["_sheet"] = sheet_name
    data["_row"] = row_num
    return data

def build_system_id_to_rows(workbook, sheet_name: str, header_row: int = 2, id_col_name: str = "講演会ID"):
    ws = workbook.worksheet(sheet_name)

    headers = make_unique(ws.row_values(header_row))
    if id_col_name not in headers:
        raise RuntimeError(f"'{id_col_name}' not found in '{sheet_name}' header_row={header_row}")

    id_col = headers.index(id_col_name) + 1
    id_values = ws.col_values(id_col)  # 列だけ（軽め）

    index: dict[str, list[int]] = {}
    for row_num, v in enumerate(id_values, start=1):
        if row_num <= header_row:
            continue
        key = str(v).strip()
        if not key:
            continue
        index.setdefault(key, []).append(row_num)  # 上から順に溜まる

    return index

def fetch_rows_for_system_id_fast(
    workbook,
    sheet_name: str,
    index: dict[str, list[int]],
    system_id: str,
    header_row: int = 2,
    col_end: str = "Z",
):
    ws = workbook.worksheet(sheet_name)
    headers = make_unique(ws.row_values(header_row))

    rows = index.get(str(system_id).strip(), [])
    if not rows:
        return []

    # 該当範囲を一括取得（例: A10:Z80）
    min_row, max_row = min(rows), max(rows)
    values = ws.get(f"A{min_row}:{col_end}{max_row}")  # まとめて取る

    out = []
    for row_num in rows:
        rel = row_num - min_row  # values内のindex
        row_values = values[rel] if 0 <= rel < len(values) else []

        if len(row_values) < len(headers):
            row_values += [""] * (len(headers) - len(row_values))

        out.append({"sheet": sheet_name, "row": row_num, "data": dict(zip(headers, row_values))})

    return out


def preload_system_id_index(workbook, sheetname_list, *, header_row: int = 2, id_col_name: str = "システムID"):
    """
    3シート側は「存在チェック」用途。
    system_id -> [{sheet,row}, ...]
    """
    index: dict[str, list[dict]] = {}

    for sheet_name in sheetname_list:
        ws = workbook.worksheet(sheet_name)

        headers = make_unique(ws.row_values(header_row))
        if id_col_name not in headers:
            # シート構成が変わっても落とさずSKIP
            print(f"[SKIP] {sheet_name}: '{id_col_name}' not found (header_row={header_row})")
            continue

        id_col = headers.index(id_col_name) + 1
        ids = ws.col_values(id_col)  # 列だけ（軽い）

        for row_num, v in enumerate(ids, start=1):
            if row_num <= header_row:
                continue
            key = str(v).strip()
            if not key:
                continue
            index.setdefault(key, []).append({"sheet": sheet_name, "row": row_num})

    return index


def find_event_rows_by_system_id(workbook, event_id, sheetname_list, header_row=1, id_col_name="システムID"):
    results = []

    for sheet_name in sheetname_list:
        ws = workbook.worksheet(sheet_name)

        # ヘッダー取得（重複対策でユニーク化）
        headers_raw = ws.row_values(header_row)
        headers = make_unique(headers_raw)

        if id_col_name not in headers:
            print(f"[SKIP] {sheet_name}: '{id_col_name}' column not found")
            continue

        id_col_idx = headers.index(id_col_name) + 1  # gspreadは1始まり

        # 「システムID」列だけ取得して検索
        id_values = ws.col_values(id_col_idx)

        for row_num, v in enumerate(id_values, start=1):
            if row_num <= header_row:
                continue

            if str(v).strip() == str(event_id).strip():
                row_values = ws.row_values(row_num)
                # 列数ズレ対策
                if len(row_values) < len(headers):
                    row_values += [""] * (len(headers) - len(row_values))

                results.append({
                    "sheet": sheet_name,
                    "row": row_num,
                    "data": dict(zip(headers, row_values))
                })

    return results

def _norm(s: str) -> str:
    return " ".join((s or "").replace("　"," ").split()).strip().lower()

def sim(a: str, b: str) -> float:
    a = _norm(a); b = _norm(b)
    if not a or not b: return 0.0
    return SequenceMatcher(None, a, b).ratio()

def rank_vm_candidates(pptx_title: str, pptx_speaker: str, vm_rows: list[dict], k: int = 5):
    scored = []
    for r in vm_rows:
        d = r["data"]
        s_title = d.get("演題","")
        s_name  = d.get("案内状掲載 医師名","")
        # タイトル重視＋名前少し
        sc = 0.75*sim(pptx_title, s_title) + 0.25*sim(pptx_speaker, s_name)
        scored.append((sc, r))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[:k]


def apply_vm_correction_no_ai(payload, vm_rows: list[dict], *, hi=0.90, gap=0.06, k=5):
    """
    AIなし：スコアが高い時だけ talk.title を補正。
    演者はシート抜けもあるので基本 keep（触らない）。
    """
    if not vm_rows or not getattr(payload, "talks", None):
        return payload

    warnings = getattr(payload, "warnings", None) or []
    payload.warnings = warnings

    for talk in payload.talks:
        pptx_title = (getattr(talk, "title", "") or "").strip()
        pptx_speaker = (getattr(talk, "speaker", "") or "").strip()

        if not pptx_title and getattr(talk, "title_lines", None):
            pptx_title = "\n".join(talk.title_lines).strip()

        top = rank_vm_candidates(pptx_title, pptx_speaker, vm_rows, k=k)
        if not top:
            continue

        best_score = top[0][0]
        second_score = top[1][0] if len(top) > 1 else 0.0

        if best_score >= hi and (best_score - second_score) >= gap:
            chosen = top[0][1]["data"]
            # 演題のみ補正（必要ならここで項目追加）
            if chosen.get("演題"):
                talk.title = chosen["演題"]
        else:
            payload.warnings.append("vm_match_not_confident")

    return payload

def get_gsa_credentials(scopes):
    gsa_json = (os.getenv("GSA_JSON") or "").strip()
    if gsa_json:
        return Credentials.from_service_account_info(
            json.loads(gsa_json),
            scopes=scopes,
        )

    # ローカル用 fallback
    sa_path = (os.getenv("GOOGLE_SA_PATH") or str(APP_DIR / "_master" / "client_secret.json")).strip()
    if Path(sa_path).exists():
        return Credentials.from_service_account_file(
            sa_path,
            scopes=scopes,
        )

    raise RuntimeError("GSA_JSON (or GOOGLE_SA_PATH) is not set")

def read_spreadsheet(event_id) -> str:
    print("Reading spreadsheet...")
    print("Event ID:", event_id)
    scope = ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive']
    
    credentials = get_gsa_credentials(scope)
    gc = gspread.authorize(credentials)
    SPREADSHEET_KEY = '1hiV0Ve2cnYyrPkBuZcZIcLWeAnJ-ucNiB0P4owZpXug'
    workbook = gc.open_by_key(SPREADSHEET_KEY)
    sheetname_list = ["VM(GWET)","VM(例外)","VM(本社)"]
    
    hits = find_event_rows_by_system_id(
    workbook=workbook,
    event_id=event_id,
    sheetname_list=sheetname_list,
    header_row=2,      # ヘッダーが1行目なら1
    id_col_name="システムID"
)

    return hits

def dump_json(obj) -> str:
    if hasattr(obj, "model_dump_json"):
        return obj.model_dump_json(indent=2)
    return obj.json(ensure_ascii=False, indent=2)


def new_session_id() -> str:
    return uuid.uuid4().hex


def normalize_space(s: str) -> str:
    s = (s or "").replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_keep_newlines(s: str) -> str:
    """
    改行は保持したまま、各行の空白だけ正規化する
    """
    s = (s or "").replace("\u3000", " ")
    lines = []
    for line in s.splitlines():
        line = re.sub(r"[ \t]+", " ", line).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


def normalize_key(s: str) -> str:
    s = (s or "").replace("\u3000", " ")
    s = re.sub(r"\s+", "", s)
    return s


def norm_name(s: str) -> str:
    s = normalize_space(s).replace("先生", "").strip()
    s = s.replace("　", " ")
    s = s.replace(" ", "")
    return s


def split_tilde_subtitle_lines(line: str) -> List[str]:
    """
    1行内の ~...~ / ～...～ を「別行扱い」にする
    """
    s = normalize_space(line)
    if not s:
        return []

    if re.fullmatch(r"[~～].+[~～]", s):
        return [s]

    m = re.search(r"([~～].+[~～])", s)
    if not m:
        return [s]

    before = normalize_space(s[: m.start()])
    sub = normalize_space(m.group(1))
    after = normalize_space(s[m.end() :])

    out: List[str] = []
    if before:
        out.append(before)
    if sub:
        out.append(sub)
    if after:
        out.append(after)
    return out


def normalize_lines_keep_order(lines: List[str]) -> List[str]:
    out: List[str] = []
    seen = set()
    for l in lines:
        for x in split_tilde_subtitle_lines(l):
            x = normalize_space(x)
            if not x:
                continue
            if x not in seen:
                out.append(x)
                seen.add(x)
    return out


def job_paths(job_id: str):
    d = DATA_DIR / job_id
    d.mkdir(parents=True, exist_ok=True)
    return {
        "dir": d,
        "input": d / "input.bin",
        "pptx": d / "input.pptx",
        "json": d / "latest.json",
        "jpg": d / "preview.jpg",
        "debug_html": d / "debug.html",
        "debug_blocks": d / "blocks.json",
    }


def fix_warnings(payload: DesignJSON) -> None:
    w = set(payload.warnings or [])
    if payload.organizer:
        w.discard("missing_organizer")

    # ★ほぼ埋まっていて confidence が高いなら ai_refined は外す（運用用）
    core_ok = bool(payload.event_title) and bool(payload.datetime) and bool(payload.organizer)
    if core_ok and (payload.confidence or 0) >= 0.98:
        w.discard("ai_refined")

    payload.warnings = sorted(w)


ORG_LABEL_PAT = re.compile(r"^(主催|共催|提供|企画|運営)\s*[:：]\s*(.+)$")

ORG_LABEL_PAT = re.compile(r"^(主催|共催|提供|企画|運営)\s*[:：]\s*(.+)$")
ORG_BRACKET_PAT = re.compile(r"^[【\[]\s*(主催|共催|提供|企画|運営)\s*[】\]]\s*(.+)$")

def _organizer_seps_to_space(s: str) -> str:
    # ／,、・ などは全部スペースに寄せる
    s = s.replace("／", " ").replace("/", " ")
    s = s.replace(",", " ").replace("，", " ")
    s = s.replace("、", " ").replace("・", " ")
    s = normalize_space(s)
    return s

def normalize_organizer(org: str) -> str:
    s = normalize_space(org).replace("（", "(").replace("）", ")")

    m = ORG_LABEL_PAT.match(s)
    if not m:
        m = ORG_BRACKET_PAT.match(s)

    if m:
        label = m.group(1)
        body = m.group(2)
        body = ORG_CANON.get(body, body)
        body = _organizer_seps_to_space(body)
        return f"{label}: {body}"   # ★半角コロン + 半角スペース

    # ラベル無し
    s2 = ORG_CANON.get(s, s)
    return _organizer_seps_to_space(s2)

KANJI_NAME_PAT = re.compile(r"^[\u4E00-\u9FFF]{2,6}$")

def add_space_to_jp_name(name: str) -> str:
    s = str(name or "").replace(" ", "").replace("\u3000", "").strip()
    if not s:
        return ""

    # すでにスペース入りならそのまま
    if " " in str(name) or "\u3000" in str(name):
        return normalize_space(name)

    # 日本語っぽい（ざっくりCJK）だけを対象
    # ※ ここはあなたの既存判定があるならそれを使ってOK
    if not all("\u3040" <= ch <= "\u30ff" or "\u4e00" <= ch <= "\u9fff" for ch in s):
        return s

    n = len(s)
    if n == 2:
        # 2文字は分けない（姓1名1を決め打ちすると事故るので）
        return s
    if n == 3:
        return s[:2] + " " + s[2:]     # 前田 潤
    if n == 4:
        return s[:2] + " " + s[2:]     # 山田 太郎
    if n == 5:
        return s[:3] + " " + s[3:]     # 佐々木 健太 など寄せ
    if n == 6:
        return s[:3] + " " + s[3:]
    return s


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()

def safe_title_for_list(payload: DesignJSON) -> str:
    # 一覧用（短く）
    if payload.event_title_lines:
        return payload.event_title_lines[0]
    return (payload.event_title or "").splitlines()[0] if payload.event_title else ""


def join_lines(lines: list[str]) -> str:
    lines = [str(l).rstrip() for l in (lines or [])]
    lines = [l for l in lines if l != ""]
    return "\n".join(lines)

def pull_affil_out_of_title_lines(talk) -> None:
    if getattr(talk, "affiliation", ""):
        return

    lines = list(getattr(talk, "title_lines", []) or [])
    if not lines:
        return

    affils = [ln for ln in lines if looks_like_affil_line(ln)]
    if not affils:
        return

    new_lines = [ln for ln in lines if ln not in affils]

    # まず affiliation は確定させる（全部所属でもOK）
    talk.affiliation = " / ".join(_norm1(a) for a in affils)

    # ★全部所属だった場合：title_lines は空にするが、title は保持する（または空なら所属を残す）
    if not new_lines:
        talk.title_lines = []
        if hasattr(talk, "title"):
            # title が空なら affiliation を見える場所に残す（運用しやすい）
            if not (talk.title or "").strip():
                talk.title = join_lines(lines)  # 元の表示を維持
        return

    # 通常ケース
    talk.title_lines = new_lines
    if hasattr(talk, "title"):
        talk.title = join_lines(new_lines)


def normalize_for_render(payload: DesignJSON) -> DesignJSON:
    # event title 合成
    if hasattr(payload, "event_title_lines") and payload.event_title_lines:
        payload.event_title = join_lines(payload.event_title_lines)

    # ★座長ラベルの掃除（「座長：」が name に残るのを防ぐ）
    if getattr(payload, "chair", None):
        if getattr(payload.chair, "name", ""):
            payload.chair.name = normalize_space(payload.chair.name).replace("座長：", "").replace("座長:", "").strip()
        if getattr(payload.chair, "name_display", ""):
            payload.chair.name_display = normalize_space(payload.chair.name_display).replace("座長：", "").replace("座長:", "").strip()

    # talk title 合成（titleフィールドが存在する時だけ代入）
    for t in (payload.talks or []):
        # ★ここで「所属が title_lines に混ざったやつ」を剥がす
        pull_affil_out_of_title_lines(t)

        if hasattr(t, "title_lines") and t.title_lines:
            if hasattr(t, "title"):
                t.title = join_lines(t.title_lines)

    return payload


DT_RE = re.compile(
    r"(20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日.*?\d{1,2}\s*[:：]\s*\d{2}\s*[～〜\-ー~]\s*\d{1,2}\s*[:：]\s*\d{2})"
)

DATE_ONLY_RE = re.compile(r"(20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)")

TIME_RANGE_RE = re.compile(r"(\d{1,2}\s*[:：]\s*\d{2}\s*[～〜\-ー~]\s*\d{1,2}\s*[:：]\s*\d{2})")

def normalize_time_colon(s: str) -> str:
    # "19：00" -> "19:00"
    return (s or "").replace("：", ":")

_ZEN2HAN = str.maketrans("０１２３４５６７８９：", "0123456789:")

def normalize_datetime_text(s: str) -> str:
    s = (s or "").translate(_ZEN2HAN)
    s = normalize_space(s)
    s = s.replace("：", ":")
    # 年月日まわりの空白を削除
    s = re.sub(r"\s*年\s*", "年", s)
    s = re.sub(r"\s*月\s*", "月", s)
    s = re.sub(r"\s*日\s*", "日", s)
    # コロン前後
    s = re.sub(r"\s*:\s*", ":", s)
    # 20 :00 → 20:00
    s = re.sub(r"(\d)\s+(\d{2})", r"\1\2", s)
    return s.strip()

def looks_like_datetime_text(s: str) -> bool:
    s0 = normalize_datetime_text(s)
    if DT_RE.search(s0):
        return True
    # 「年/月/日 + ～」があれば日時濃厚
    if ("年" in s0 and "月" in s0 and "日" in s0 and ("～" in s0 or "〜" in s0 or "-" in s0)):
        return True
    return False

def looks_like_label(s: str) -> bool:
    k = normalize_key(s)
    # 「日 時」みたいな分割にも強い
    return any(x in k for x in ["日時", "日時", "日", "時", "座長", "演者", "主催", "共催", "提供", "企画", "運営", "会場", "形式", "登録", "視聴"])

def looks_like_talk_anchor(s: str) -> bool:
    k = normalize_key(s)
    # 講演１ / 講演1 / 演題1 等
    return bool(re.search(r"(講演|演題)([0-9]|[１-９])", k))



# ---------------- Jobs list / filter ----------------

def parse_warnings(warnings_json: str) -> List[str]:
    try:
        return json.loads(warnings_json or "[]")
    except Exception:
        return []

# def row_to_job_item(r: sqlite3.Row) -> Dict[str, Any]:
#     job_id = r["job_id"]
#     session_id = r["session_id"]
#     event_id = r["event_id"]
#     return {
#         "jobId": job_id,
#         "filename": r["filename"],
#         "session_id": session_id, 
#         "event_id": event_id, 
#         "status": r["status"],
#         "createdAt": r["created_at"],
#         "updatedAt": r["updated_at"],
#         "title": r["title"] or "",
#         "organizer": r["organizer"] or "",
#         "datetime": r["datetime"] or "",
#         "confidence": float(r["confidence"] or 0.0),
#         "warnings": parse_warnings(r["warnings_json"]),
#         "manualOverride": bool(r["manual_override"]),
#         "note": r["note"] or "",
#         "locked": bool(r["locked"]),
#         "errorMessage": r["error_message"],
#         "previewUrl": f"/preview/{job_id}.jpg",
#         "downloadUrl": f"/download/{job_id}.jpg",
#     }

def row_to_job_item(r) -> Dict[str, Any]:
    job_id = r["job_id"]
    return {
        "jobId": job_id,
        "filename": r.get("filename") or "",
        "session_id": r.get("session_id") or "",
        "event_id": r.get("event_id") or "",
        "status": r["status"],
        "createdAt": r["created_at"].isoformat(),
        "updatedAt": r["updated_at"].isoformat(),
        "title": r.get("title") or "",
        "organizer": r.get("organizer") or "",
        "datetime": r.get("datetime") or "",
        "confidence": float(r.get("confidence") or 0.0),
        "warnings": r.get("warnings_json") or [],
        "manualOverride": bool(r.get("manual_override")),
        "note": r.get("note") or "",
        "locked": bool(r.get("locked")),
        "errorMessage": r.get("error_message"),
        "previewUrl": f"/preview/{job_id}.jpg",
        "downloadUrl": f"/download/{job_id}.jpg",
    }




# ---------------- SQLite index ----------------

# def db_connect() -> sqlite3.Connection:
#     con = sqlite3.connect(DB_PATH)
#     con.row_factory = sqlite3.Row
#     return con

def db_connect():
    # row_factory = dict_row で r["job_id"] 形式を維持
    return psycopg.connect(DATABASE_URL, row_factory=dict_row)


def init_db():
    con = db_connect()
    try:
        con.execute("""
        CREATE TABLE IF NOT EXISTS jobs (
            job_id TEXT PRIMARY KEY,
            filename TEXT,
            session_id TEXT,
            event_id TEXT NOT NULL,     
            status TEXT NOT NULL,                 -- ok / error
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,

            title TEXT,
            organizer TEXT,
            datetime TEXT,
            
            region TEXT,
            unit TEXT,

            confidence REAL,
            warnings_json TEXT,

            manual_override INTEGER NOT NULL DEFAULT 0,
            note TEXT NOT NULL DEFAULT '',
            locked INTEGER NOT NULL DEFAULT 0,

            error_message TEXT
            
        );
        """)
        con.execute("CREATE INDEX IF NOT EXISTS idx_jobs_updated_at ON jobs(updated_at);")
        con.execute("CREATE INDEX IF NOT EXISTS idx_jobs_status ON jobs(status);")
        con.commit()
    finally:
        con.close()

# def upsert_job_ok(
#     job_id: str,
#     filename: str,
#     payload: DesignJSON,
#     session_id: Optional[str] = None,
#     event_id: Optional[str] = None,
# ):
#     con = db_connect()
#     try:
#         created_at = now_iso()
#         updated_at = created_at

#         warnings_json = json.dumps(payload.warnings or [], ensure_ascii=False)

#         # 既存があれば created_at 維持 + session_id/event_id を未指定なら引き継ぐ
#         row = con.execute(
#             "SELECT created_at, session_id, event_id FROM jobs WHERE job_id=?",
#             (job_id,),
#         ).fetchone()
#         if row:
#             created_at = row["created_at"]
#             updated_at = now_iso()
#             if session_id is None:
#                 session_id = row["session_id"]
#             if event_id is None:
#                 event_id = row["event_id"]

#         # 新規insertで None が残ると DB制約で死ぬなら空文字に落とす（安全策）
#         if session_id is None:
#             session_id = ""
#         if event_id is None:
#             event_id = ""

#         con.execute(
#             """
#             INSERT INTO jobs (
#                 job_id, filename, session_id, event_id, status, created_at, updated_at,
#                 title, organizer, datetime,
#                 region, unit,
#                 confidence, warnings_json,
#                 manual_override, note, locked,
#                 error_message
#             ) VALUES (?, ?, ?, ?, 'ok', ?, ?,
#                       ?, ?, ?,
#                       ?, ?,
#                       ?, ?,
#                       ?, ?, ?,
#                       NULL)
#             ON CONFLICT(job_id) DO UPDATE SET
#                 filename=excluded.filename,
#                 session_id=excluded.session_id,
#                 event_id=excluded.event_id,
#                 status='ok',
#                 updated_at=excluded.updated_at,
#                 title=excluded.title,
#                 organizer=excluded.organizer,
#                 datetime=excluded.datetime,
#                 region=excluded.region,
#                 unit=excluded.unit,
#                 confidence=excluded.confidence,
#                 warnings_json=excluded.warnings_json,
#                 manual_override=excluded.manual_override,
#                 note=excluded.note,
#                 locked=excluded.locked,
#                 error_message=NULL
#             """,
#             (
#                 job_id,
#                 filename,
#                 session_id,
#                 event_id,  # ★ここはそのまま正しい並び
#                 created_at,
#                 updated_at,
#                 safe_title_for_list(payload),
#                 payload.organizer,
#                 payload.datetime,
#                 payload.region,
#                 payload.unit,
#                 float(payload.confidence or 0.0),
#                 warnings_json,
#                 1 if getattr(payload, "manual_override", False) else 0,
#                 getattr(payload, "note", "") or "",
#                 1 if getattr(payload, "locked", False) else 0,
#             ),
#         )
#         con.commit()
#     finally:
#         con.close()

def now_ts():
    return datetime.now(timezone.utc)

def upsert_job_ok(job_id: str, filename: str, payload, session_id: str = "", event_id: str = ""):
    created_at = now_ts()
    updated_at = created_at
    warnings = payload.warnings or []

    # jsonb array として入れる（["missing_organizer", ...]）
    warnings_jsonb = json.dumps(warnings, ensure_ascii=False)

    with db_connect() as con:
        row = con.execute(
            "SELECT created_at, session_id, event_id FROM jobs WHERE job_id=%s",
            (job_id,),
        ).fetchone()

        if row:
            created_at = row["created_at"]
            updated_at = now_ts()
            if not session_id:
                session_id = row.get("session_id") or ""
            if not event_id:
                event_id = row.get("event_id") or ""

        con.execute(
            """
            INSERT INTO jobs (
              job_id, filename, session_id, event_id, status, created_at, updated_at,
              title, organizer, datetime, region, unit, confidence, warnings_json,
              manual_override, note, locked, error_message
            )
            VALUES (%s,%s,%s,%s,'ok',%s,%s,
                    %s,%s,%s,%s,%s,%s,%s::jsonb,
                    %s,%s,%s,NULL)
            ON CONFLICT (job_id) DO UPDATE SET
              filename=excluded.filename,
              session_id=excluded.session_id,
              event_id=excluded.event_id,
              status='ok',
              updated_at=excluded.updated_at,
              title=excluded.title,
              organizer=excluded.organizer,
              datetime=excluded.datetime,
              region=excluded.region,
              unit=excluded.unit,
              confidence=excluded.confidence,
              warnings_json=excluded.warnings_json,
              manual_override=excluded.manual_override,
              note=excluded.note,
              locked=excluded.locked,
              error_message=NULL
            """,
            (
                job_id, filename, session_id or "", event_id or "",
                created_at, updated_at,
                safe_title_for_list(payload),
                payload.organizer or "",
                payload.datetime or "",
                payload.region or "",
                payload.unit or "",
                float(payload.confidence or 0.0),
                warnings_jsonb,
                bool(getattr(payload, "manual_override", False)),
                getattr(payload, "note", "") or "",
                bool(getattr(payload, "locked", False)),
            ),
        )
        con.commit()  # ★必須


def upsert_job_error(job_id: str, filename: str, error_message: str, event_id: str = ""):
    created_at = now_ts()
    updated_at = created_at

    with db_connect() as con:
        row = con.execute(
            "SELECT created_at, event_id FROM jobs WHERE job_id=%s",
            (job_id,),
        ).fetchone()

        if row:
            created_at = row["created_at"]
            updated_at = now_ts()
            if not event_id:
                event_id = (row.get("event_id") or "")

        con.execute(
            """
            INSERT INTO jobs (
              job_id, filename, session_id, event_id, status, created_at, updated_at,
              title, organizer, datetime, region, unit, confidence, warnings_json,
              manual_override, note, locked, error_message
            )
            VALUES (%s,%s,%s,%s,'error',%s,%s,
                    %s,%s,%s,%s,%s,%s,%s::jsonb,
                    %s,%s,%s,%s)
            ON CONFLICT (job_id) DO UPDATE SET
              filename=excluded.filename,
              status='error',
              event_id=excluded.event_id,
              updated_at=excluded.updated_at,
              error_message=excluded.error_message
            """,
            (
                job_id,
                filename,
                "",                 # session_id は error の時は空でOK（必要なら保持）
                event_id or "",
                created_at,
                updated_at,
                "", "", "",         # title/organizer/datetime
                "", "",             # region/unit
                0.0,
                "[]",               # warnings_json
                False,
                "",
                False,
                error_message or "",
            ),
        )
        con.commit()

# ---------------- PPTX (Blocks) ----------------
@dataclass
class TextBlock:
    text: str
    left: int
    top: int
    width: int
    height: int
    max_font_pt: float


def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        # GROUP = 6
        if getattr(sh, "shape_type", None) == 6:
            for sub in iter_shapes(sh.shapes):
                yield sub


def extract_blocks_from_pptx(pptx_path: Path, first_slide_only: bool = True) -> List[TextBlock]:
    prs = Presentation(str(pptx_path))
    blocks: List[TextBlock] = []

    slides = [prs.slides[0]] if (first_slide_only and len(prs.slides) > 0) else prs.slides
    for slide in slides:
        for sh in iter_shapes(slide.shapes):
            if not getattr(sh, "has_text_frame", False):
                continue
            tf = sh.text_frame
            if not tf:
                continue

            paras = []
            max_font = 0.0
            for p in tf.paragraphs:
                t = (p.text or "").strip()
                if t:
                    paras.append(t)
                for run in p.runs:
                    if run.font and run.font.size:
                        max_font = max(max_font, run.font.size / EMU_PER_PT)

            # ★改行保持する
            text = normalize_keep_newlines("\n".join(paras))
            if not text:
                continue

            blocks.append(
                TextBlock(
                    text=text,
                    left=int(sh.left),
                    top=int(sh.top),
                    width=int(sh.width),
                    height=int(sh.height),
                    max_font_pt=float(max_font or 0.0),
                )
            )

    blocks.sort(key=lambda b: (b.top, b.left))
    return blocks


def blocks_to_lines(blocks: List[TextBlock]) -> List[str]:
    # 改行は潰して良い用途（datetime/organizer検出など）向け
    out: List[str] = []
    seen = set()
    for b in blocks:
        s = normalize_space(b.text.replace("\n", " "))
        if not s:
            continue
        if s not in seen:
            out.append(s)
            seen.add(s)
    return out


def in_region(b: TextBlock, x0: float, y0: float, x1: float, y1: float) -> bool:
    cx = b.left + b.width / 2.0
    cy = b.top + b.height / 2.0
    return (x0 <= cx <= x1) and (y0 <= cy <= y1)

def looks_like_body_text_for_title(s: str) -> bool:
    s2 = normalize_space(s)
    if not s2:
        return False
    # long polite-body sentences
    if len(s2) >= 30:
        if any(k in s2 for k in ["謹啓", "謹白", "時下", "平素", "ご高配", "ご案内", "ご多用", "厚く御礼", "お慶び"]):
            return True
    # explicitly exclude these keywords even if short
    if any(k in s2 for k in ["謹啓", "謹白"]):
        return True
    
    body_kw = [
        "本会は", "事前参加登録", "参加をご希望", "担当者へご連絡",
        "医療従事者", "医療系資格", "学生", "受付", "医療事務",
        "ご参加はご遠慮", "お願い申し上げます", "ご了承ください",
        "ご視聴には", "事前参加予約", "芳名録", "個人情報",
    ]
    if any(k in s2 for k in body_kw):
        return True
    
    return False


# def looks_like_body_text_for_title(s: str) -> bool:
#     s2 = normalize_space(s)
#     if not s2:
#         return False
        

#     # 既存: 挨拶文
#     if any(k in s2 for k in ["謹啓", "謹白", "時下", "平素", "ご高配", "厚く御礼"]):
#         return True

#     # ★追加: 案内・注意文（今回の混入パターン）
#     body_kw = [
#         "本会は", "事前参加登録", "参加をご希望", "担当者へご連絡",
#         "医療従事者", "医療系資格", "学生", "受付", "医療事務",
#         "ご参加はご遠慮", "お願い申し上げます", "ご了承ください",
#         "ご視聴には", "事前参加予約", "芳名録", "個人情報",
#     ]
#     if any(k in s2 for k in body_kw):
#         return True

#     # 長文は本文率高い（ただしタイトル長めもあるので閾値は控えめに）
#     if len(s2) >= 40:
#         return True

#     return False

def looks_like_format_value(s: str) -> bool:
    s2 = normalize_space(s)
    return (
        "Live配信" in s2
        or "Web（" in s2
        or s2.endswith("による開催")
        or s2.endswith("による配信")
        or s2.startswith("Web")
    )

def extract_event_title_lines_from_blocks(blocks: List[TextBlock]) -> List[str]:
    """
    タイトルは「上部」「大きめフォント」「日時っぽくない」を優先。
    複数行（サブタイトル含む）は上から連続ブロックとして拾う。
    """
    if not blocks:
        return []

    # 上半分を優先（テンプレ安定）
    tops = [b.top for b in blocks]
    min_top, max_top = min(tops), max(tops)
    mid_top = min_top + (max_top - min_top) * 0.55

    cand = []
    for b in blocks:
        s = normalize_space(b.text)
        if not s:
            continue
        if looks_like_body_text_for_title(s):
            continue
        if looks_like_datetime_text(s):
            continue
        if looks_like_format_value(s):
            continue
        # フッター/ラベル除外
        if any(x in s for x in ["主催", "共催", "座長", "演者", "会場", "形式"]):
            continue
        # 上の方 & それなりに大きい
        if b.top <= mid_top and (b.max_font_pt or 0) >= 18:
            cand.append(b)

    if not cand:
        # fallback：日時除外だけして最大フォント
        cand2 = [b for b in blocks if b.text and not looks_like_datetime_text(b.text)]
        if not cand2:
            return []
        b0 = max(cand2, key=lambda x: x.max_font_pt or 0)
        return [normalize_space(b0.text)]

    # まず “タイトル本体” を決める（大きさ優先、同点なら上）
    head = sorted(cand, key=lambda b: ((b.max_font_pt or 0), -b.top), reverse=True)[0]

    # タイトルは head と “近い縦位置の連続” を拾う（サブタイトルを残す）
    # 同じカラム（左寄り）にあるものを拾う
    lines = []
    # head の周辺範囲
    x0 = head.left - 500000
    x1 = head.left + head.width + 500000
    y0 = head.top - 200000
    y1 = head.top + 2500000  # 下に2〜3行分

    near = [b for b in blocks if in_region(b, x0, y0, x1, y1)]
    near = sorted(near, key=lambda b: (b.top, b.left))

    for b in near:
        s = normalize_space(b.text)
        if not s:
            continue
        if looks_like_body_text_for_title(s):
            continue
        if looks_like_datetime_text(s):
            continue
  
        if looks_like_format_value(s):
            continue
        if any(x in s for x in ["主催", "共催", "座長", "演者", "会場", "形式"]):
            continue
        # ここは subtitle "~...~" を消さない（要望対応）
        lines.append(s)

    # 重複除去（順序維持）
    out = []
    seen = set()
    for s in lines:
        if s not in seen:
            out.append(s)
            seen.add(s)

    # それでも空なら head のみ
    return out or [normalize_space(head.text)]





def extract_datetime_from_blocks(blocks: List[TextBlock]) -> str:
    lines = blocks_to_lines(blocks)
    dt_pat = re.compile(
        r"(20\d{2}年\s*\d{1,2}月\s*\d{1,2}日.*?\d{1,2}:\d{2}\s*[～〜\-ー~]\s*\d{1,2}:\d{2})"
    )
    for l in lines:
        m = dt_pat.search(l)
        if m:
            return normalize_space(m.group(1))

    for l in lines:
        if "日時" in l:
            s = normalize_space(l)
            s = re.sub(r"^.*日時\s*[:：]?\s*", "", s)
            return s
    return ""


def extract_organizer_from_blocks(blocks: List[TextBlock]) -> str:
    lines = blocks_to_lines(blocks)

    # ラベル行をそのまま返す（主催/共催/提供/企画/運営）
    pat = re.compile(r"^(主催|共催|提供|企画|運営)\s*[:：]\s*(.+)$")
    for l in lines:
        s = normalize_space(l)
        m = pat.match(s)
        if m:
            # 例: "共催：刈谷内科医会 MSD 株式会社"
            return s

    # fallback（会社名っぽい行）
    corp_pat = re.compile(r"(株式会社|有限会社|合同会社|Inc\.|LLC|Ltd\.|Co\.,?\s*Ltd\.|GmbH)")
    for l in reversed(lines):
        s = normalize_space(l)
        if corp_pat.search(s):
            return s
    return ""



def extract_datetime_from_blocks(blocks: List[TextBlock]) -> str:
    if not blocks:
        return ""

    # 1) 「日時/日 時」ラベルを探す（分割にも強い）
    label_blocks = []
    for b in blocks:
        k = normalize_key(b.text)
        if "日時" in k or (("日" in k) and ("時" in k) and len(k) <= 6):
            label_blocks.append(b)

    # ラベルが見つかったら、その右・下の近傍を最優先で探す
    if label_blocks:
        anchor = sorted(label_blocks, key=lambda x: x.top, reverse=False)[0]
        x0 = anchor.left - 200000
        x1 = anchor.left + 6500000
        y0 = anchor.top - 250000
        y1 = anchor.top + 2000000

        near = [b for b in blocks if in_region(b, x0, y0, x1, y1)]
        near = sorted(near, key=lambda b: (b.top, b.left))

        for b in near:
            s = normalize_datetime_text(b.text)
            if not s:
                continue
            m = DT_RE.search(s)
            if m:
                return normalize_datetime_text(m.group(1))
            # “時間帯だけ”の行があるなら拾う（後で日付と組む余地）
            if DATE_ONLY_RE.search(s) and TIME_RANGE_RE.search(s):
                return normalize_datetime_text(s)

    # 2) 全ブロックから正規表現で拾う（fallback）
    for b in sorted(blocks, key=lambda x: (x.top, x.left)):
        s = normalize_datetime_text(b.text)
        if not s:
            continue
        m = DT_RE.search(s)
        if m:
            return normalize_datetime_text(m.group(1))

    # 3) 「日時：」形式 fallback
    for b in sorted(blocks, key=lambda x: (x.top, x.left)):
        s = normalize_datetime_text(b.text)
        if "日時" in s:
            s2 = re.sub(r"^.*日時\s*[:：]?\s*", "", s).strip()
            if s2:
                return s2

    return ""

def extract_time_candidates_from_blocks(blocks: List[TextBlock]) -> List[str]:
    out: List[str] = []
    seen = set()
    for b in blocks:
        m = TIME_PAT.search(b.text.replace("\n", " "))
        if m:
            t = normalize_space(m.group(1))
            if t not in seen:
                out.append(t)
                seen.add(t)
    return out

def split_name_affil_inline(text: str) -> tuple[str, str]:
    """
    例:
    '永井　英明　先生　（ Web 講演） 独立行政法人...感染症センター長'
    -> ('永井英明', '独立行政法人...感染症センター長')
    """
    s = normalize_space(text)
    if "先生" not in s:
        return "", ""
    # 先生より前を名前、後ろを所属（括弧の補足は捨てる）
    m = re.search(r"(.+?)\s*先生\s*(?:（.*?）\s*)?(.*)$", s)
    if not m:
        return "", ""
    raw_name = normalize_space(m.group(1))
    raw_aff = normalize_space(m.group(2))
    key = norm_name(raw_name)
    return key, raw_aff

def is_overall_datetime(tm: str, overall: str) -> bool:
    if not tm or not overall:
        return False
    return normalize_space(tm) in normalize_space(overall)


def _ns(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").replace("\u3000", " ")).strip()

def _is_greeting(s: str) -> bool:
    s = _ns(s)
    return any(k in s for k in ["謹啓","謹白","平素は","厚く御礼","さてこの度","ご清祥","お慶び","ご多用"])

def extract_chair_from_blocks(blocks, speaker_map):
    ordered = sorted(blocks, key=lambda b: (b.top, b.left))
    filtered = [b for b in ordered if not _is_greeting(b.text)]

    def norm_key_for_map(name_disp: str) -> str:
        return norm_name(_ns(name_disp))

    def looks_like_affil_line(s: str) -> bool:
        s = _ns(s).replace("\n", " ")
        if not s:
            return False
        if "先生" in s:
            return False
        if any(w in s for w in ["座長", "演者", "講演", "日時", "会場", "主催", "共催", "提供", "視聴", "登録", "お願い", "ご注意"]):
            return False
        kw = ["大学", "病院", "クリニック", "センター", "科", "部", "教授", "講師", "医師", "部長", "院長", "医療"]
        return any(w in s for w in kw) or len(s) >= 10

    def find_affil_right_same_row(target_block) -> str:
        # ★同じ高さ帯の右側ブロックを拾う（今回の blocks にドンピシャ）
        cand = []
        for b in filtered:
            if b is target_block:
                continue
            if b.left <= target_block.left:
                continue
            # 高さ帯が近い
            if abs(b.top - target_block.top) > 450000:
                continue
            s = _ns(b.text)
            if not looks_like_affil_line(s):
                continue
            dx = b.left - target_block.left
            dy = abs(b.top - target_block.top)
            score = dx + dy * 0.3
            cand.append((score, s))
        cand.sort(key=lambda x: x[0])
        return cand[0][1] if cand else ""

    def pick_affil_near_lines(lines, i, fallback_key):
        for j in range(i + 1, min(i + 5, len(lines))):
            if "【講演" in lines[j] or "講演" == lines[j].replace(" ", ""):
                break
            if looks_like_affil_line(lines[j]):
                return _ns(lines[j])
        return (speaker_map.get(fallback_key) or "").strip()

    # (1) 最優先：座長：◯◯先生 が同一ブロックにある
    for b in filtered:
        t = _ns(b.text)
        if "座長" in t and "先生" in t:
            m = re.search(r"座長[：:\s]*([^\n]+?)\s*先生", t)
            if m:
                name_disp = _ns(m.group(1))
                key = norm_key_for_map(name_disp)
                # ★まず横（右側）→ ダメなら speaker_map
                aff = find_affil_right_same_row(b) or (speaker_map.get(key) or "").strip()
                return {"name": key, "name_display": name_disp, "affiliation": aff}

    # (2) 次：座長 ラベル単独
    chair_labels = []
    for b in filtered:
        t = _ns(b.text).replace("：", "").replace(":", "").replace(" ", "")
        if t == "座長":
            chair_labels.append(b)

    if chair_labels:
        lbl = sorted(chair_labels, key=lambda b: (b.top, b.left))[0]
        x0 = lbl.left - 200000
        x1 = lbl.left + 6500000
        y0 = lbl.top - 200000
        y1 = lbl.top + 1600000

        cands = []
        for b in filtered:
            if b is lbl:
                continue
            if not in_region(b, x0, y0, x1, y1):
                continue
            if "先生" not in (b.text or ""):
                continue
            if "【講演" in (b.text or ""):
                continue
            cands.append(b)

        cands.sort(key=lambda b: (abs(b.top - lbl.top) + abs(b.left - lbl.left), b.top, b.left))

        for b in cands[:5]:
            lines = [_ns(x) for x in str(b.text).split("\n") if _ns(x)]
            for i, line in enumerate(lines):
                if "先生" in line:
                    name_disp = _ns(line.replace("先生", ""))
                    key = norm_key_for_map(name_disp)
                    if not key:
                        continue
                    aff = pick_affil_near_lines(lines, i, key)
                    return {"name": key, "name_display": name_disp, "affiliation": aff}

    # (3) fallback：巨大ブロックの「講演より前の最初の先生」
    has_chair_label = any(_ns(b.text).replace("：", "").replace(":", "").replace(" ", "") == "座長" for b in filtered)
    if has_chair_label:
        bigs = [b for b in filtered if "先生" in (b.text or "")]
        bigs.sort(key=lambda b: (-(b.width * b.height), b.top, b.left))
        for b in bigs[:3]:
            lines = [_ns(x) for x in str(b.text).split("\n") if _ns(x)]
            for i, line in enumerate(lines):
                if "【講演" in line:
                    break
                if "先生" in line:
                    name_disp = _ns(line.replace("先生", ""))
                    key = norm_key_for_map(name_disp)
                    if not key:
                        continue
                    aff = pick_affil_near_lines(lines, i, key)
                    return {"name": key, "name_display": name_disp, "affiliation": aff}

    # (4) 最後：上から最初の先生
    for b in filtered:
        if "先生" not in (b.text or ""):
            continue
        lines = [_ns(x) for x in str(b.text).split("\n") if _ns(x)]
        for i, line in enumerate(lines):
            if "先生" in line and "【講演" not in line:
                name_disp = _ns(line.replace("先生", ""))
                key = norm_key_for_map(name_disp)
                if key:
                    aff = pick_affil_near_lines(lines, i, key)
                    return {"name": key, "name_display": name_disp, "affiliation": aff}

    return None

def ensure_display_fields(payload: DesignJSON) -> DesignJSON:
    # chair
    if getattr(payload, "chair", None):
        c = payload.chair
        if (getattr(c, "name", "") or "").strip() and not (getattr(c, "name_display", "") or "").strip():
            c.name_display = add_space_to_jp_name(c.name) or c.name

    # talks
    for t in (payload.talks or []):
        # speaker_display を必ず作る（speaker優先）
        sp = (getattr(t, "speaker", "") or "").strip()
        if sp and not (getattr(t, "speaker_display", "") or "").strip():
            t.speaker_display = add_space_to_jp_name(sp) or sp

        # speaker が空で display だけある場合は speaker を作る（逆補完）
        disp = (getattr(t, "speaker_display", "") or "").strip()
        if (not sp) and disp:
            t.speaker = norm_name(disp) or disp.replace(" ", "").replace("\u3000", "")

    return payload

def extract_speaker_affil_map_by_blocks(blocks: List[TextBlock]) -> Dict[str, str]:
    mp: Dict[str, str] = {}

    # ① 同一ブロック内（先生 + 所属）優先（既存のまま）
    for b in blocks:
        key, aff = split_name_affil_inline(b.text)
        if key:
            mp.setdefault(key, "")
        if key and aff:
            mp[key] = aff

    def looks_like_affil(s: str) -> bool:
        s = normalize_space(s)
        if not s:
            return False
        if s.startswith("※"):
            return False
        if "先生" in s:
            return False
        k = normalize_key(s)
        if any(x in k for x in ["座長","演者","講演","日時","会場","主催","共催","提供","企画","運営","登録","視聴"]):
            return False
        kw = ["大学", "病院", "クリニック", "センター", "内科", "外科", "部", "科",
              "講師", "教授", "医師", "部長", "院長", "研究", "機構"]
        return any(w in s for w in kw) or len(s) >= 10

    name_blocks = [b for b in blocks if ("先生" in b.text and "座長" not in b.text and "演者" not in b.text)]
    ordered = sorted(blocks, key=lambda b: (b.top, b.left))

    for nb in name_blocks:
        raw = normalize_space(nb.text.replace("先生", ""))
        key = norm_name(raw)
        if not key:
            continue
        mp.setdefault(key, "")  # ★所属が見つからなくてもキーだけ作る（後段の照合が安定）

        # ★まず「直下」を最優先（このテンプレで一番多い）
        below = []
        x0 = nb.left - 400000
        x1 = nb.left + nb.width + 400000
        y0 = nb.top + nb.height - 100000
        y1 = nb.top + nb.height + 1200000
        for b in ordered:
            if b is nb:
                continue
            if not in_region(b, x0, y0, x1, y1):
                continue
            s = normalize_space(b.text.replace("\n", " "))
            if looks_like_affil(s):
                below.append((abs(b.top - nb.top), s))
        below.sort(key=lambda x: x[0])
        if below:
            mp[key] = below[0][1]
            continue

        # ★次に「右+下」広め（2カラム/右寄せ対策）
        cand = []
        x0 = nb.left - 200000
        x1 = nb.left + 6500000
        y0 = nb.top - 200000
        y1 = nb.top + 1800000

        for b in ordered:
            if b is nb:
                continue
            if not in_region(b, x0, y0, x1, y1):
                continue
            s = normalize_space(b.text.replace("\n", " "))
            if not looks_like_affil(s):
                continue

            cx = b.left + b.width / 2.0
            cy = b.top + b.height / 2.0
            nx = nb.left + nb.width / 2.0
            ny = nb.top + nb.height / 2.0

            # ★「下方向」を強く優遇（所属は下に来ることが多い）
            dy = max(0, cy - ny)
            dx = abs(cx - nx)
            dist = dx + dy * 0.6  # 下を優遇

            cand.append((dist, s))

        cand.sort(key=lambda x: x[0])
        if cand:
            mp[key] = cand[0][1]

    return mp

def extract_chair_by_blocks(blocks: List[TextBlock], speaker_map: Dict[str, str]) -> Chair:
    """
    「座長」アンカー近傍から
    - raw（表示用：スペース保持）
    - key（照合用：スペース除去）
    を分離して取得する
    """
    chair_anchor = None
    for b in blocks:
        if "座長" in b.text:
            chair_anchor = b
            break
    if not chair_anchor:
        return Chair()

    # 座長ラベル近傍
    x0 = chair_anchor.left - 200000
    x1 = chair_anchor.left + 5000000
    y0 = chair_anchor.top - 200000
    y1 = chair_anchor.top + 1200000

    near = [b for b in blocks if in_region(b, x0, y0, x1, y1)]

    # ① 「◯◯ 先生」を最優先
    for b in near:
        if "先生" in b.text:
            raw = normalize_space(b.text.replace("先生", ""))
            key = norm_name(raw)
            aff = speaker_map.get(key, "")
            return Chair(
                name=key,
                name_display=raw,     # ★ スペース保持
                affiliation=normalize_space(aff),
            )

    # ② フォールバック：speaker_map のキーが含まれるか
    joined = normalize_key("\n".join(b.text for b in near))
    for key, aff in speaker_map.items():
        if key and key in joined:
            return Chair(
                name=key,
                name_display=add_space_to_jp_name(key),
                affiliation=normalize_space(aff),
            )

    return Chair()



def pick_time(texts: List[str], time_candidates: List[str]) -> str:
    for t in texts:
        m = TIME_PAT.search((t or "").replace("\n", " "))
        if m:
            tm = normalize_space(m.group(1))
            if tm in time_candidates or not time_candidates:
                return tm
    return ""

def pick_time_from_near_texts(texts: List[str]) -> Optional[str]:
    """
    講演近傍に明示された時間のみ拾う。
    全体時間は絶対に拾わない。
    """
    for t in texts:
        m = TIME_PAT.search(t)
        if m:
            return normalize_space(m.group(1))
    return ""


def pick_speaker(texts: List[str], speaker_map: Dict[str, str]) -> tuple[str, str]:
    """
    return (speaker_key, speaker_display)
    """
    for t in texts:
        if "先生" in t:
            raw = normalize_space(t.replace("先生", ""))
            key = norm_name(raw)
            if key in speaker_map:
                return key, raw  # ★ raw を保持
    return "", ""


def pick_title_lines(texts: List[str]) -> List[str]:
    """
    近傍テキスト群から「演題行」を改行保持で抽出し、
    - 講演1 / 講 演１ などは完全除外
    - ~...~ は必ず別行扱い
    """
    skip_keywords = [
        "講演", "演者", "座長", "日時", "会場", "開催",
        "主催", "共催", "提供", "企画", "運営",
        "登録", "視聴", "Web", "Live"
    ]

    lines: List[str] = []

    for t in texts:
        for raw in (t or "").split("\n"):
            s = normalize_space(raw)
            if not s:
                continue

            sk = normalize_key(s)

            # ★★★ 講演1 / 講 演１ / 講演２… を完全除外 ★★★
            if re.fullmatch(r"講演[0-9１-９]+", sk):
                continue

            # ラベル系除外（正規化キーで判定）
            if any(k in sk for k in skip_keywords):
                continue

            # 時間除外
            if TIME_PAT.search(s):
                continue

            # 人名除外
            if s.endswith("先生") and len(s) <= 14:
                continue

            # 短すぎるノイズ除外
            if len(s) < 4:
                continue

            # ~...~ は別行扱い
            lines.extend(split_tilde_subtitle_lines(s))

    # 重複排除（順序維持）
    uniq, seen = [], set()
    for l in lines:
        if l not in seen:
            uniq.append(l)
            seen.add(l)

    return uniq

def extract_talks_by_blocks(blocks: List[TextBlock], speaker_map: Dict[str, str]) -> List[Talk]:
    """
    まず「講演1/演題1」等のアンカーを優先。
    ただし、時間帯（HH:MM～HH:MM）が複数あるテンプレでは
    “時間行を起点” にセグメント分割して talk を構築する（EM2512対策）。
    """
    talks: List[Talk] = []
    if not blocks:
        return talks

    ordered = sorted(blocks, key=lambda b: (b.top, b.left))
    lines = [normalize_space(b.text) for b in ordered if normalize_space(b.text)]

    def is_time_line(s: str) -> str:
        s2 = normalize_time_colon(normalize_space(s))
        m = TIME_RANGE_RE.search(s2)
        return normalize_space(m.group(1)) if m else ""

    def is_aff_line(s: str) -> bool:
        if not s:
            return False
        if is_time_line(s):
            return False
        return any(k in s for k in s for s in [])  # dummy to keep mypy calm (ignored)

    # ↑ 上のダミーは不要なら削除してOK。ここから本物:
    def is_aff_line(s: str) -> bool:
        if not s:
            return False
        if is_time_line(s):
            return False
        return any(k in s for k in [
            "病院", "クリニック", "医院", "診療所", "大学", "センター", "機構", "総合病院",
            "内科", "外科", "部", "科",
            "教授", "准教授", "講師", "医長", "部長", "院長", "主任", "理事長"
        ])

    def strip_label(prefixes, s: str) -> str:
        s2 = normalize_space(s)
        for p in prefixes:
            if s2.startswith(p):
                s2 = re.sub(rf"^{re.escape(p)}\s*[:：]?\s*", "", s2).strip()
        return s2
    
    def _key_variants(name: str) -> List[str]:
        n = normalize_space(name or "")
        n2 = n.replace("\u3000", " ").replace(" ", "")
        return [n, n2, normalize_key(n), normalize_key(n2)]

    speaker_map_norm: Dict[str, str] = {}
    for k, v in (speaker_map or {}).items():
        for kk in _key_variants(k):
            if kk and v and kk not in speaker_map_norm:
                speaker_map_norm[kk] = v

    def aff_from_speaker_map(name: str) -> str:
        for kk in _key_variants(name):
            if kk in speaker_map_norm:
                return speaker_map_norm[kk]
        return ""

    # ------------------------------------------------------------------
    # ★ 先に time_idxs を計算（anchors を使うか判定するため）
    #    日時行（2026/03/06 ... 19:00～20:20）を time として数えない
    # ------------------------------------------------------------------
    time_idxs: List[int] = []
    for i, s in enumerate(lines):
        if looks_like_datetime_text(s):
            continue
        if is_time_line(s):
            time_idxs.append(i)

    # ------------------------------------------------------------------
    # 1) 講演アンカー（ただし time が複数あるテンプレでは使わない）
    # ------------------------------------------------------------------
    anchors = [b for b in ordered if looks_like_talk_anchor(b.text)]
    anchors = sorted(anchors, key=lambda b: (b.top, b.left))

    def near_texts(a: TextBlock):
        x0 = a.left - 200000
        x1 = a.left + 6500000
        y0 = a.top - 200000
        y1 = a.top + 2000000
        near = [b for b in ordered if in_region(b, x0, y0, x1, y1)]
        near = sorted(near, key=lambda b: (b.top, b.left))
        return near, [normalize_space(b.text) for b in near if normalize_space(b.text)]

    def pick_local_time(near_texts_list: List[str]) -> str:
        for t in near_texts_list:
            if looks_like_datetime_text(t):
                continue
            tm = is_time_line(t)
            if tm:
                return tm
        return ""

    def pick_speaker_from_texts(texts: List[str]) -> str:
        # “演者：” 優先
        for t in texts:
            if "演者" in normalize_key(t):
                cand = strip_label(["演者", "演者:", "演者："], t)
                cand = norm_name(cand)
                if cand:
                    return cand
        # speaker_map の key を含むかでも見る
        joined = normalize_key("\n".join(texts))
        for k in speaker_map.keys():
            if k and k in joined:
                return k
        return ""

    def pick_title_lines(texts: List[str]) -> List[str]:
        out: List[str] = []
        for t in texts:
            s = normalize_space(t)
            if not s:
                continue
            if looks_like_talk_anchor(s):
                continue
            if looks_like_datetime_text(s):
                continue
            if any(x in s for x in ["座長", "演者", "主催", "共催", "会場", "形式", "登録", "視聴"]):
                continue
            if is_time_line(s):
                continue
            # 「演題：」のときはラベル除去
            if "演題" in normalize_key(s):
                s = strip_label(["演題", "演題:", "演題："], s)
            # 人名単体は除外
            if s.endswith("先生") and len(s) <= 16:
                continue
            if len(s) >= 6:
                out.append(s)

        res: List[str] = []
        seen = set()
        for s in out:
            if s not in seen:
                res.append(s)
                seen.add(s)
        return res[:4]

    # ★ time が複数あるなら anchors を使わず 2) に任せる（EM2512対策）
    if anchors and len(time_idxs) < 2:
        for a in anchors[:4]:
            _, texts = near_texts(a)
            time = pick_local_time(texts)
            speaker = pick_speaker_from_texts(texts)
            aff = speaker_map.get(speaker, "") if speaker else ""
            title_lines = pick_title_lines(texts)
            if title_lines or speaker or time:
                talks.append(Talk(time=time, title_lines=title_lines, speaker=speaker, affiliation=aff))
        return talks[:4]

    # ------------------------------------------------------------------
    # 2) 時間行が複数ある場合：時間起点で分割
    # ------------------------------------------------------------------
    if len(time_idxs) >= 2:
        # 時間ブロックを収集（日時行は除外）
        time_blocks: List[tuple[TextBlock, str]] = []
        for b in ordered:
            if looks_like_datetime_text(b.text):
                continue
            tm = is_time_line(b.text)
            if tm:
                time_blocks.append((b, tm))

        time_blocks.sort(key=lambda x: (x[0].left, x[0].top))

        def _content_left_for_time(tb: TextBlock) -> int:
            """時間ブロック(tb)に紐づく“本文側”の left を推定する。"""
            y0 = tb.top - 250000
            y1 = tb.top + 900000

            cand: List[TextBlock] = []
            # まず「演題」ラベルを探す
            for b in ordered:
                if b.top < y0 or b.top > y1:
                    continue
                s = normalize_space(b.text)
                if not s:
                    continue
                if looks_like_datetime_text(s):
                    continue
                if is_time_line(s):
                    continue
                if "演題" in normalize_key(s):
                    cand.append(b)

            # 無ければ“それっぽい本文”を探す（長めで、所属/ラベル/名前ではない）
            if not cand:
                for b in ordered:
                    if b.top < y0 or b.top > y1:
                        continue
                    s = normalize_space(b.text)
                    if not s:
                        continue
                    if looks_like_datetime_text(s):
                        continue
                    if is_time_line(s) or is_aff_line(s):
                        continue
                    if any(x in normalize_key(s) for x in ["演者", "座長", "主催", "共催", "会場", "形式", "登録", "視聴"]):
                        continue
                    if len(s) >= 10:
                        cand.append(b)

            if not cand:
                return tb.left

            cand.sort(key=lambda b: (abs(b.left - tb.left), b.left))
            return cand[0].left

        # time_blocks それぞれに対して“本文側left”を計算
        time_blocks2: List[tuple[TextBlock, str, int]] = []
        for tb, tm in time_blocks:
            time_blocks2.append((tb, tm, _content_left_for_time(tb)))

        # カラム境界推定
        col_lefts: List[int] = []
        for _, _, left in time_blocks2:
            if not col_lefts:
                col_lefts.append(left)
                continue
            if min(abs(left - x) for x in col_lefts) > 900000:
                col_lefts.append(left)
        col_lefts = sorted(col_lefts)

        def col_right_bound(left: int) -> int:
            for x in col_lefts:
                if x > left + 900000:
                    return x - 300000
            return left + 6500000

        def looks_like_name_line(s: str) -> bool:
            s2 = normalize_space(s).replace("先生", "").strip()
            if TIME_RANGE_RE.search(normalize_time_colon(s2)):
                return False
            if any(k in s2 for k in ["演題", "演者", "座長", "病院", "クリニック", "大学", "内科", "外科", "教授", "講師", "部長", "院長", "理事長"]):
                return False
            return bool(re.fullmatch(r"[一-龥々]{2,6}\s*[一-龥々]{1,6}", s2))

        # 本文側leftで安定ソート
        time_blocks2.sort(key=lambda x: (x[2], x[0].top))

        used = set()

        def looks_like_affiliation(s: str) -> bool:
            s = normalize_space(s or "")
            if not s:
                return False
            # 施設・所属・役職っぽい語が入ってたら「タイトル継続」ではない
            keywords = [
                "大学", "病院", "センター", "研究科", "学部", "診療科", "内科", "外科",
                "教授", "准教授", "講師", "部長", "科長", "主任", "医長",
                "先生", "MD", "PhD"
            ]
            return any(k in s for k in keywords)

        for idx_tb, (tb, tm, base_left) in enumerate(time_blocks2):
            if id(tb) in used:
                continue

            # 次の同カラム時間を探して下限にする
            next_top = None
            for j in range(idx_tb + 1, len(time_blocks2)):
                tb2, _, base_left2 = time_blocks2[j]
                if tb2.top <= tb.top:
                    continue
                if abs(base_left - base_left2) <= 900000:
                    next_top = tb2.top
                    break

            x0 = base_left - 300000
            x1 = col_right_bound(base_left)
            y0 = tb.top - 200000
            y1 = (next_top - 200000) if next_top is not None else (tb.top + 3500000)

            near = [b for b in ordered if in_region(b, x0, y0, x1, y1)]
            near = sorted(near, key=lambda b: (b.top, b.left))
            seg_lines = [normalize_space(b.text) for b in near if normalize_space(b.text)]

            # ★重要：同一セグメントに「次の時間行」が混ざったらそこで打ち切る
            seg2: List[str] = []
            started = False
            for s in seg_lines:
                if looks_like_datetime_text(s):
                    continue
                tm2 = is_time_line(s)
                if tm2:
                    if not started:
                        started = True
                        seg2.append(s)
                        continue
                    if tm2 != tm:
                        break
                if started:
                    seg2.append(s)
            if seg2:
                seg_lines = seg2

            title_lines: List[str] = []
            speaker = ""
            affiliation = ""
            aff_candidates: List[str] = []

            for j, s in enumerate(seg_lines):
                k = normalize_key(s)

                if not title_lines and "演題" in k:
                    t = strip_label(["演題", "演題:", "演題："], s)
                    if t:
                        for ln in (t or "").split("\n"):
                            ln = normalize_space(ln)
                            if not ln:
                                continue
                            if is_aff_line(ln):
                                continue
                            title_lines.append(ln)

                    # 次行がサブタイトルっぽければ追加（～で始まる等）
                    if j + 1 < len(seg_lines):
                        nxt = normalize_space(seg_lines[j + 1])
                        # raw_nxt = str(seg_lines[j + 1] or "")
                        if nxt and (nxt.startswith(("～", "~"))) and (not looks_like_affiliation(nxt)):
                            title_lines.append(nxt)

                if not speaker and "演者" in k:
                    sp = strip_label(["演者", "演者:", "演者："], s)
                    sp = norm_name(sp)
                    if sp:
                        speaker = sp

                # 「演者」ラベルが無いテンプレ用：名前っぽい行
                if not speaker and looks_like_name_line(s):
                    speaker = norm_name(s)

                if is_aff_line(s):
                    aff_candidates.append(s)

            # affiliation確定
            if not affiliation and aff_candidates:
                def aff_score(a: str) -> int:
                    score = 0
                    if any(k in a for k in ["病院", "クリニック", "大学", "センター", "総合病院", "機構"]):
                        score += 2
                    if any(k in a for k in ["内科", "外科", "科", "部"]):
                        score += 2
                    if any(k in a for k in ["教授", "准教授", "講師", "部長", "院長", "理事長", "主任", "医長"]):
                        score += 2
                    return score
                aff_candidates.sort(key=lambda a: (-aff_score(a), len(a)))
                affiliation = aff_candidates[0]

            # speaker_map から補完（欠損のみ）
            if speaker and not affiliation:
                affiliation = aff_from_speaker_map(speaker)

            # タイトル fallback: 「演題」ラベルが無いテンプレ用
            if not title_lines:
                for s in seg_lines:
                    if looks_like_datetime_text(s):
                        continue
                    if is_time_line(s):
                        continue
                    if is_aff_line(s):
                        continue
                    if "演者" in normalize_key(s) or "座長" in normalize_key(s):
                        continue
                    if looks_like_name_line(s):
                        continue
                    if len(s) >= 10:
                        title_lines.append(s)
                        break

            talks.append(Talk(time=tm, title_lines=title_lines[:4], speaker=speaker, affiliation=affiliation))

            used.add(id(tb))
            if len(talks) >= 4:
                break

        talks = [t for t in talks if t.time or t.title_lines or t.speaker or t.affiliation]

        def _is_notice_lines(tl: List[str]) -> bool:
            if not tl:
                return False
            tl2 = [normalize_space(x) for x in tl if normalize_space(x)]
            if not tl2:
                return False
            joined = normalize_key("\n".join(tl2))
            # 箇条書き + 注意語
            if all(x.startswith("・") for x in tl2[:min(3, len(tl2))]):
                if any(k in joined for k in ["事前", "参加", "登録", "ご遠慮", "医療従事者", "資格", "担当者へご連絡"]):
                    return True
            return False

        def _is_empty_chair_only(t: Talk) -> bool:
            # タイトルなし/演者なし で affiliation だけ（=座長/施設だけ拾ったゴミ）を落とす
            if (t.speaker or "").strip():
                return False
            if (t.title_lines or []) and any(normalize_space(x) for x in t.title_lines):
                return False
            # affiliation だけがあるケース
            if (t.affiliation or "").strip():
                return True
            return False

        def _should_drop(t: Talk) -> bool:
            # 注意書きtalk
            if _is_notice_lines(t.title_lines or []):
                return True
            # speaker/title無しの “座長/施設だけ”
            if _is_empty_chair_only(t):
                return True
            return False

        talks = [t for t in talks if (t.time or t.title_lines or t.speaker or t.affiliation)]
        talks = [t for t in talks if not _should_drop(t)]
  

        return talks[:4]

    # ------------------------------------------------------------------
    # 3) 最後のfallback：演題ラベル周辺を複数件拾う（2) に入らない時の保険）
    # ------------------------------------------------------------------
    label_idxs: List[int] = []
    for i, s in enumerate(lines):
        if "演題" in normalize_key(s):
            label_idxs.append(i)

    seen = set()

    for label_idx in label_idxs[:6]:
        seg = lines[max(0, label_idx - 3): min(len(lines), label_idx + 25)]

        tm = ""
        for s in seg:
            if looks_like_datetime_text(s):
                continue
            tm2 = is_time_line(s)
            if tm2:
                tm = tm2
                break

        title_lines: List[str] = []
        speaker = ""
        affiliation = ""

        for j, s in enumerate(seg):
            k = normalize_key(s)

            if not title_lines and "演題" in k:
                t = strip_label(["演題", "演題:", "演題："], s).strip()
                if t:
                    for ln in t.split("\n"):
                        ln = normalize_space(ln)
                        if not ln or is_aff_line(ln):
                            continue
                        title_lines.append(ln)

                if j + 1 < len(seg):
                    nxt = normalize_space(seg[j + 1])
                    if nxt and (nxt.startswith("～") or nxt.startswith("~")):
                        title_lines.append(nxt)

            if not speaker and "演者" in k:
                speaker = norm_name(strip_label(["演者", "演者:", "演者："], s))

        # affiliation は time 行の次の所属っぽい行を優先
        for j, s in enumerate(seg):
            if looks_like_datetime_text(s):
                continue
            if is_time_line(s):
                for kk in range(j + 1, min(len(seg), j + 6)):
                    ss = normalize_space(seg[kk])
                    if not ss:
                        continue
                    if "演題" in normalize_key(ss) or "演者" in normalize_key(ss) or "座長" in normalize_key(ss):
                        continue
                    if is_aff_line(ss):
                        affiliation = ss
                        break
                break

        if speaker and not affiliation:
            affiliation = speaker_map.get(speaker, "") or ""

        if title_lines or speaker or tm or affiliation:
            key = (normalize_space(tm), normalize_space(speaker), join_lines(title_lines))
            if key not in seen:
                talks.append(Talk(time=tm, title_lines=title_lines[:4], speaker=speaker, affiliation=affiliation))
                seen.add(key)

        if len(talks) >= 4:
            break

    

        

    return talks[:4]



def find_sponsor_logo_blobs(pptx_path: Path) -> list[bytes]:
    prs = Presentation(str(pptx_path))
    if len(prs.slides) == 0:
        return []

    slide = prs.slides[0]

    sponsor_text_shapes = []
    for sh in iter_shapes(slide.shapes):
        if not getattr(sh, "has_text_frame", False):
            continue
        txt = normalize_space(getattr(sh.text_frame, "text", "") or "")
        if not txt:
            continue
        if "主催" in normalize_key(txt):
            sponsor_text_shapes.append(sh)

    if not sponsor_text_shapes:
        return []

    anchor = sorted(sponsor_text_shapes, key=lambda s: int(getattr(s, "top", 0)), reverse=True)[0]

    x0 = int(anchor.left + anchor.width) - 200000
    x1 = int(anchor.left + anchor.width) + 5000000
    y0 = int(anchor.top) - 400000
    y1 = int(anchor.top + anchor.height) + 900000

    blobs: list[bytes] = []
    for sh in iter_shapes(slide.shapes):
        if getattr(sh, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
            cx = int(sh.left + sh.width / 2)
            cy = int(sh.top + sh.height / 2)
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                try:
                    blobs.append(sh.image.blob)
                except Exception:
                    pass

    return blobs


async def ocr_company_name_with_openai(image_bytes: bytes) -> str:
    if not OPENAI_API_KEY:
        return ""

    b64 = base64.b64encode(image_bytes).decode("utf-8")
    data_url = f"data:image/jpg;base64,{b64}"

    prompt = """この画像はセミナー案内の「主催」ロゴです。
ロゴから読み取れる会社名/団体名のみを日本語で1行で返してください。
不明なら空文字を返してください。余計な説明は禁止。"""

    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }

    body = {
        "model": AI_MODEL,
        "messages": [
            {"role": "system", "content": "Return only plain text. No extra words."},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            },
        ],
        "temperature": 0.0,
    }

    async with httpx.AsyncClient(timeout=AI_TIMEOUT) as client:
        r = await client.post(f"{OPENAI_BASE_URL}/chat/completions", headers=headers, json=body)
        r.raise_for_status()
        data = r.json()

    text = (data["choices"][0]["message"]["content"] or "").strip()
    text = text.splitlines()[0].strip()
    return text


async def try_fill_organizer_from_logo(pptx_path: Path) -> str:
    blobs = find_sponsor_logo_blobs(pptx_path)
    if not blobs:
        return ""

    for blob in blobs[:2]:
        name = await ocr_company_name_with_openai(blob)
        if name:
            return name
    return ""


# ---------------- AI ----------------
def build_ai_prompt(
    blocks: List[TextBlock],
    draft: DesignJSON,
    speaker_map: Dict[str, str],
    time_candidates: List[str],
) -> str:
    blocks_json: List[Dict[str, Any]] = [
        {
            "text": b.text,  # 改行含む
            "left": b.left,
            "top": b.top,
            "width": b.width,
            "height": b.height,
            "max_font_pt": round(b.max_font_pt, 2),
        }
        for b in blocks
    ]

    return f"""あなたは日本語の医療セミナー案内スライド(PPTX)から情報を抽出してJSONに整形するアシスタントです。

# 重要ルール（厳守）
- 出力は JSONのみ（前後に文章を付けない）
- 誤推測禁止：抽出データに根拠がない値は空文字
- datetime は日時のみ（形式/配信方法は入れない）
- organizer は主催の会社名（抽出データにある場合のみ）
- chair は座長
- talks は1〜4件、順序はスライドの登場順（top/left順）
- 空のtalkは禁止（title_lines/speaker/timeが全て空の要素を作らない）
- event_title_lines / talk.title_lines は改行を保持して配列で返す（統合しない）
- 1行内の ~...~ / ～...～ は必ず別行（別要素）扱いにする
- talk.speaker は speaker_map のキーから選ぶ。不明なら空
- talk.affiliation は speaker_map[talk.speaker] をそのまま使用（推測禁止）
- talk.time は「その講演ブロック近傍に明示された時間」のみ使用する
- 全体の開催時間(datetime)を talk.time に使ってはいけない
- 「先生」はspeakerから除去する（例：河 良崇 先生 → 河良崇）
- talksは「講演1/2/3...」アンカー付近（座標的に近いブロック）から構成すること

# speaker_map（所属の根拠。ここに無い所属は書かない）
{json.dumps(speaker_map, ensure_ascii=False, indent=2)}

# time_candidates（talk.timeはここから選ぶ）
{json.dumps(time_candidates, ensure_ascii=False, indent=2)}

# 抽出ブロック（座標付き）
{json.dumps(blocks_json, ensure_ascii=False, indent=2)}

# 下書きJSON（ルールベース）
{dump_json(draft)}

# 出力JSONスキーマ（厳守）
{{
  "event_title_lines": ["string"],
  "event_title": "string",
  "datetime": "string",
  "organizer": "string",
  "chair": {{"name":"string","affiliation":"string"}},
  "talks": [
    {{"time":"string","title_lines":["string"],"speaker":"string","affiliation":"string"}}
  ],
  "warnings": ["string"],
  "confidence": 0.0
}}
"""


async def ai_refine_json(
    blocks: List[TextBlock],
    draft: DesignJSON,
    speaker_map: Dict[str, str],
    time_candidates: List[str],
) -> DesignJSON:
    if not OPENAI_API_KEY:
        return draft

    prompt = build_ai_prompt(blocks, draft, speaker_map, time_candidates)

    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }

    body = {
        "model": AI_MODEL,
        "messages": [
            {"role": "system", "content": "Return ONLY valid JSON. No extra text."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.0,
    }

    async with httpx.AsyncClient(timeout=AI_TIMEOUT) as client:
        r = await client.post(f"{OPENAI_BASE_URL}/chat/completions", headers=headers, json=body)
        r.raise_for_status()
        data = r.json()

    content = data["choices"][0]["message"]["content"]

    try:
        parsed = json.loads(content)
    except Exception:
        content2 = re.sub(r"^```json\s*|\s*```$", "", content.strip())
        parsed = json.loads(content2)

    refined = DesignJSON(**parsed)
    refined.talks = refined.talks[:4]

    # 正規化
    refined.event_title_lines = normalize_lines_keep_order(refined.event_title_lines or [])
    refined.event_title = "\n".join(refined.event_title_lines).strip() if refined.event_title_lines else normalize_space(refined.event_title)

    refined.datetime = normalize_space(refined.datetime)
    refined.organizer = normalize_space(refined.organizer)
    refined.chair.name = norm_name(refined.chair.name)
    refined.chair.affiliation = normalize_space(refined.chair.affiliation)

    for t in refined.talks:
        t.time = normalize_space(t.time)
        t.title_lines = normalize_lines_keep_order(t.title_lines or [])
        t.speaker = norm_name(t.speaker)
        t.affiliation = normalize_space(t.affiliation)
        

    refined = postprocess_refined(refined, speaker_map, time_candidates)

    if not refined.organizer:
        refined.organizer = extract_organizer_from_blocks(blocks)

    w = set(refined.warnings or [])
    w.add("ai_refined")
    if not refined.datetime:
        w.add("missing_datetime")
    if not refined.organizer:
        w.add("missing_organizer")
    if not refined.chair.name:
        w.add("missing_chair")
    if len(refined.talks) == 0:
        w.add("no_talks")
    if not refined.event_title_lines and not refined.event_title:
        w.add("missing_event_title")
    refined.warnings = sorted(w)

    rule_org = extract_organizer_from_blocks(blocks)
    if rule_org:
        refined.organizer = normalize_organizer(rule_org)

    return refined


# ---------------- Postprocess ----------------
def postprocess_refined(refined: DesignJSON, speaker_map: Dict[str, str], time_candidates: List[str]) -> DesignJSON:
    # event_title_lines を優先
    refined.event_title_lines = normalize_lines_keep_order(refined.event_title_lines or [])
    if refined.event_title_lines:
        refined.event_title = "\n".join(refined.event_title_lines).strip()
    else:
        # event_title しか来ない場合の救済
        if refined.event_title:
            # 1行内の ~...~ を別行化したい
            lines = []
            for raw in refined.event_title.split("\n"):
                lines.extend(split_tilde_subtitle_lines(raw))
            refined.event_title_lines = normalize_lines_keep_order(lines)
            refined.event_title = "\n".join(refined.event_title_lines).strip()

    cleaned: List[Talk] = []
    for t in refined.talks:
        if not (t.title_lines or t.speaker or t.affiliation or t.time):
            continue

        # title_lines 正規化（~...~ 別行化 + 重複排除）
        t.title_lines = normalize_lines_keep_order(t.title_lines or [])

        sp = norm_name(t.speaker)
        t.speaker = sp
        t.speaker_display = add_space_to_jp_name(t.speaker)

        # if sp in speaker_map:
        #     t.affiliation = speaker_map[sp] or ""
        # else:
        #     t.affiliation = ""

        if not (t.affiliation or "").strip():
            cand = (speaker_map.get(sp) or "").strip()
            # 長すぎる所属（PDFの注意文混入）を弾く
            if cand and len(cand) <= 80 and ("ご視聴" not in cand) and ("お願い" not in cand):
                t.affiliation = cand

        tm = normalize_space(t.time)
        if time_candidates and tm and tm not in time_candidates:
            tm = ""
        t.time = tm

        if t.title_lines or t.speaker or t.time:
            cleaned.append(t)
        
        if t.time and refined.datetime and normalize_space(t.time) in normalize_space(refined.datetime):
            t.time = ""
        
        

    refined.talks = cleaned[:4]

    if refined.chair.name and not refined.chair.name_display:
        refined.chair.name_display = add_space_to_jp_name(refined.chair.name)

    
    return refined


# ---------------- Parse (Rule + AI) ----------------
def parse_blocks_to_design_json(blocks: List[TextBlock]) -> DesignJSON:
    warnings: List[str] = []
    confidence = 0.78

    event_title_lines = extract_event_title_lines_from_blocks(blocks)
    event_title = "\n".join(event_title_lines).strip()

    print("event_title_lines:", event_title_lines)

    dt = extract_datetime_from_blocks(blocks)

    org = extract_organizer_from_blocks(blocks)  # ←主催: を含めたいなら別途調整（必要なら次で直す）

    speaker_map = extract_speaker_affil_map_by_blocks(blocks)
    chair = extract_chair_by_blocks(blocks, speaker_map)

    talks = extract_talks_by_blocks(blocks, speaker_map)

    if not event_title:
        warnings.append("missing_event_title"); confidence -= 0.2
    if not dt:
        warnings.append("missing_datetime"); confidence -= 0.15
    if not org:
        warnings.append("missing_organizer"); confidence -= 0.1
    if not chair.name:
        warnings.append("missing_chair"); confidence -= 0.1
    if len(talks) == 0:
        warnings.append("no_talks"); confidence -= 0.35

    confidence = float(min(max(confidence, 0.0), 1.0))

    return DesignJSON(
        event_title_lines=event_title_lines,
        event_title=event_title,
        datetime=normalize_space(dt),
        organizer=normalize_space(org),
        chair=Chair(name=chair.name, name_display=chair.name_display, affiliation=chair.affiliation),
        talks=talks[:4],
        warnings=sorted(set(warnings)),
        confidence=confidence,
    )


async def pptx_to_json(pptx_path: Path, debug_blocks_path: Optional[Path] = None) -> DesignJSON:
    blocks = extract_blocks_from_pptx(pptx_path, first_slide_only=True)

    if debug_blocks_path:
        dbg = [
            {
                "text": b.text,
                "left": b.left,
                "top": b.top,
                "width": b.width,
                "height": b.height,
                "max_font_pt": round(b.max_font_pt, 2),
            }
            for b in blocks
        ]
        debug_blocks_path.write_text(json.dumps(dbg, ensure_ascii=False, indent=2), encoding="utf-8")

    speaker_map = extract_speaker_affil_map_by_blocks(blocks)
    time_candidates = extract_time_candidates_from_blocks(blocks)

    draft = parse_blocks_to_design_json(blocks)
    refined = await ai_refine_json(blocks, draft, speaker_map, time_candidates)
    refined = fill_chair_affiliation_from_blocks(refined, blocks)

    if refined.confidence < draft.confidence:
        refined.warnings = sorted(set(refined.warnings + ["ai_lower_confidence"]))

    # organizer OCR（空のときだけ）
    if not refined.organizer:
        org2 = await try_fill_organizer_from_logo(pptx_path)
        if org2:
            refined.organizer = f"主催：{normalize_organizer(org2)}"
            refined.warnings = sorted(set(refined.warnings + ["organizer_from_logo_ocr"]))
        else:
            refined.warnings = sorted(set(refined.warnings + ["organizer_logo_ocr_failed"]))

    fix_warnings(refined)

    # 互換: event_title は常に event_title_lines から作る
    refined.event_title_lines = normalize_lines_keep_order(refined.event_title_lines or [])
    refined.event_title = "\n".join(refined.event_title_lines).strip() if refined.event_title_lines else refined.event_title

    return refined


# ---------------- Render (HTML→PNG) ----------------

# ---------------- VMヒント（演題演者）: PPTX探索精度UP + 欠損のみ補完 ----------------

def _norm_person_name(v: str) -> str:
    s = str(v or "").replace("\u3000", " ")
    s = " ".join(s.split())
    return s.replace(" ", "")

def _vm_aff_str(vm: dict) -> str:
    facility = (vm.get("案内状掲載 施設名") or "").strip()
    dept = (vm.get("案内状掲載 所属科") or "").strip()
    role = (vm.get("案内状掲載 役職") or "").strip()
    parts = [p for p in [facility, dept, role] if p]
    return " ".join(parts).strip()

def _norm_title_key(s: str) -> str:
    s = normalize_space(str(s or ""))
    # 記号ゆらぎを減らす
    s = s.replace("～", "〜").replace("−", "-").replace("—", "-").replace("–", "-").replace("－", "-")
    # かっこ/引用符などを除去（マッチ安定）
    for ch in ['"', "“", "”", "「", "」", "’", "‘", "（", "）", "(", ")", "【", "】", "[", "]"]:
        s = s.replace(ch, "")
    # スペース除去
    return s.replace(" ", "").replace("\u3000", "")

_UNWANTED_TITLE_WORDS = [
    "開会", "閉会", "開会の辞", "閉会の辞",
    "挨拶", "ご挨拶", "総合司会", "司会",
    "休憩", "休", "intermission",
    "動画上映", "ビデオ", "Video", "上映",
    "事務連絡", "諸連絡", "注意事項", "ご案内",
    "総合討論", "討論", "質疑", "Q&A",
    "オープニング", "エンディング",
]

def _is_unwanted_talk(title: str) -> bool:
    t = normalize_space(title)
    if not t:
        return True
    k = _norm_title_key(t)
    return any(_norm_title_key(w) in k for w in _UNWANTED_TITLE_WORDS)

def looks_like_real_talk(t: Talk) -> bool:
    title = normalize_space("\n".join(t.title_lines or []).strip() or (t.title or ""))
    if not title:
        return False
    if "演題" in title:
        return True
    if len(_norm_title_key(title)) >= 12:
        return True
    if (t.time or "").strip() and (t.speaker or t.speaker_display or "").strip():
        return True
    return False

def _vm_speaker_titles(vm_rows: list[dict]) -> list[str]:
    titles = []
    for r in (vm_rows or []):
        if (r.get("役職") or "") != "演者":
            continue
        v = (r.get("演題") or "").strip()
        if v:
            titles.append(v)
    return titles

def _time_start_minutes(t: str) -> int:
    m = re.search(r"(\d{1,2}):(\d{2})", str(t or ""))
    if not m:
        return 10**9
    hh, mm = map(int, m.groups())
    return hh * 60 + mm

def _strip_outer_quotes(s: str) -> str:
    s2 = str(s or "").strip()
    if s2.startswith("「"):
        s2 = s2[1:]
    if s2.endswith("」"):
        s2 = s2[:-1]
    return s2.strip()

def _clean_title_lines(t):
    if getattr(t, "title_lines", None):
        t.title_lines = [_strip_outer_quotes(x) for x in t.title_lines]
    if getattr(t, "title", None):
        t.title = _strip_outer_quotes(t.title)
    return t

TIME_RANGE_PAT = re.compile(
    r"\d{1,2}[:：]\d{2}\s*[~\-–—−－〜～]\s*\d{1,2}[:：]\d{2}"
)

ROLE_PAT = re.compile(r"(演者|座長)")

def clean_speaker_text(s: str) -> str:
    s = str(s or "")
    s = TIME_RANGE_PAT.sub("", s)      # 時間帯を除去
    s = ROLE_PAT.sub("", s)            # 演者/座長を除去
    s = re.sub(r"\s+", " ", s).strip()

    # 漢字間の変なスペースは潰す（前 田潤 → 前田潤）
    s = re.sub(r"(?<=[一-龥])\s+(?=[一-龥])", "", s)

    # 最後に姓名スペースを付け直す（あなたの add_space_to_jp_name を使う）
    s = add_space_to_jp_name(s)
    return s


def normalize_talk_speakers(payload: DesignJSON) -> DesignJSON:
    for t in (payload.talks or []):
        base = (t.speaker_display or "").strip() or (t.speaker or "").strip()
        cleaned = clean_speaker_text(base)

        # display は姓名スペースあり
        t.speaker_display = cleaned
        # speaker はスペースなしに統一（検索/キー用）
        t.speaker = cleaned.replace(" ", "")
    return payload


def _clean_speaker_text(s: str) -> str:
    s = str(s or "")
    s = TIME_PAT.sub("", s)
    s = re.sub(r"(演者|座長)", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _clean_speaker_obj(t):
    base = (getattr(t, "speaker_display", "") or "").strip() \
           or (getattr(t, "speaker", "") or "").strip()

    cleaned = _clean_speaker_text(base)
    cleaned = add_space_to_jp_name(cleaned)

    t.speaker_display = cleaned
    t.speaker = cleaned.replace(" ", "")
    return t

def prune_talks_using_vm_titles(payload: DesignJSON, vm_rows: list[dict]) -> DesignJSON:
    talks = list(payload.talks or [])
    if not talks:
        return payload

    before_keys = [
        _norm_title_key("\n".join(t.title_lines or []).strip() or (t.title or ""))
        for t in talks
    ]

    # 1) まず不要ワードは常に落とす（VMなくても効く）
    filtered = []
    for t in talks:
        tt = normalize_space("\n".join(t.title_lines or []).strip() or (t.title or ""))
        if _is_unwanted_talk(tt):
            continue
        filtered.append(t)

    # 2) 重複除去
    seen = set()
    dedup = []
    for t in filtered:
        key = _norm_title_key("\n".join(t.title_lines or []).strip() or (t.title or ""))
        if not key or key in seen:
            continue
        seen.add(key)
        dedup.append(t)

    # ---- ここから先は VMがあるときだけ ----
    vm_titles = _vm_speaker_titles(vm_rows)
    if not vm_titles:
        dedup = [t for t in dedup if looks_like_real_talk(t)]
        payload.talks = dedup
        payload.warnings = sorted(set((payload.warnings or []) + ["talks_pruned_heuristic_only"]))
        for t in payload.talks:
            _clean_title_lines(t)
        payload = normalize_talk_speakers(payload)
        return payload

    # 3) VM演題にマッチするtalkだけ残す
    vm_keys = [_norm_title_key(x) for x in vm_titles]
    kept = []
    for t in dedup:
        k = _norm_title_key("\n".join(t.title_lines or []).strip() or (t.title or ""))
        if any(k == vk or (vk and vk in k) or (k and k in vk) for vk in vm_keys):
            kept.append(t)

    if kept:
        dedup = kept

    # 4) VM演題数へ寄せる（最大4）
    payload.talks = dedup[: min(len(vm_titles), 4)]
    for t in payload.talks:
            _clean_title_lines(t)
    payload = normalize_talk_speakers(payload)

    after_keys = [
        _norm_title_key("\n".join(t.title_lines or []).strip() or (t.title or ""))
        for t in (payload.talks or [])
    ]

    # ★「VM処理で結果が変わった」時だけ warning
    if after_keys != before_keys:
        payload.warnings = sorted(set((payload.warnings or []) + ["talks_pruned_by_vm_hint"]))

    return payload





def fill_chair_affiliation_from_vm_hint(payload: DesignJSON, blocks: list[TextBlock], vm_rows: list[dict]) -> DesignJSON:
    if not vm_rows or not payload.chair or (payload.chair.affiliation or "").strip() == "":
        pass
    else:
        return payload  # 既に入ってる

    # name -> vm（役職で絞らない。座長行が来る前提）
    vm_by_name = {}
    for vm in vm_rows:
        n = _norm_person_name(vm.get("案内状掲載 医師名", ""))
        if n:
            vm_by_name[n] = vm

    chair_key = _norm_person_name(payload.chair.name or payload.chair.name_display or "")
    vm = vm_by_name.get(chair_key)
    if not vm:
        return payload

    # 1) 施設名ヒントで blocks から拾い直す（talkと同じ戦略）
    facility = (vm.get("案内状掲載 施設名") or "").strip()
    if facility:
        pptx_aff = _join_affiliation_near_facility(blocks, facility)
        if pptx_aff:
            payload.chair.affiliation = pptx_aff
            return payload

    # 2) それでもダメなら VM 文字列を入れる（最後の保険）
    aff = _vm_aff_str(vm)
    if aff:
        payload.chair.affiliation = aff

    return payload

def fill_chair_affiliation_from_blocks(payload: DesignJSON, blocks: list[TextBlock]) -> DesignJSON:
    if not getattr(payload, "chair", None):
        return payload
    if (payload.chair.affiliation or "").strip():
        return payload

    def _nospace(s: str) -> str:
        return normalize_space(s).replace(" ", "").replace("\u3000", "")

    key_ns = _nospace(payload.chair.name)

    ordered = sorted(blocks, key=lambda b: (b.top, b.left))

    target = None
    for b in ordered:
        t = normalize_space(b.text)
        t_ns = _nospace(t)
        if "座長" in t and key_ns and key_ns in t_ns:
            target = b
            break

    if not target:
        return payload

    def looks_like_affil(s: str) -> bool:
        s = normalize_space(s).replace("\n", " ")
        if not s:
            return False
        if "先生" in s or "座長" in s:
            return False
        if any(w in s for w in ["日時", "会場", "共催", "主催", "提供", "視聴", "登録"]):
            return False
        kw = ["大学", "病院", "クリニック", "センター", "科", "部", "教授", "講師", "医師", "部長", "院長"]
        return any(w in s for w in kw)

    # ★ 横並び優先（高さ帯一致）
    cand = []
    for b in ordered:
        if b is target:
            continue

        # 右側
        if b.left <= target.left:
            continue

        # 高さ帯が重なる（重要）
        if not (abs(b.top - target.top) < 400000):
            continue

        s = normalize_space(b.text.replace("\n", " "))
        if not looks_like_affil(s):
            continue

        dx = b.left - target.left
        dy = abs(b.top - target.top)
        score = dx + dy * 0.3
        cand.append((score, s))

    if cand:
        cand.sort(key=lambda x: x[0])
        payload.chair.affiliation = cand[0][1]
        return payload

    # fallback: 下方向
    for b in ordered:
        if b.top <= target.top:
            continue
        s = normalize_space(b.text.replace("\n", " "))
        if looks_like_affil(s):
            payload.chair.affiliation = s
            break

    return payload

def _find_best_block_idx_by_hint(blocks: List[TextBlock], hint: str, *, min_sim: float = 0.60) -> int:
    if not hint:
        return -1
    hint = str(hint).strip()
    if not hint:
        return -1

    # 部分一致優先
    for i, b in enumerate(blocks):
        t = (b.text or "").strip()
        if t and hint in t:
            return i

    # 類似一致
    best_i = -1
    best_sc = 0.0
    for i, b in enumerate(blocks):
        t = (b.text or "").strip()
        if not t:
            continue
        sc = sim(hint, t)
        if sc > best_sc:
            best_sc = sc
            best_i = i
    return best_i if best_sc >= min_sim else -1

def _join_affiliation_near_facility(blocks: List[TextBlock], facility_hint: str, *, max_follow: int = 4, y_limit: int = 260) -> str:
    """PPTX上の所属表記を優先して作る: 施設名ブロック + 近傍（科/役職）を結合"""
    if not facility_hint:
        return ""
    idx = _find_best_block_idx_by_hint(blocks, facility_hint, min_sim=0.60)
    if idx < 0:
        return ""

    b0 = blocks[idx]
    base_top = b0.top
    base_left = b0.left

    parts = [(b0.text or "").strip()]
    taken = 0

    # 施設名の直後に並ぶ「科/部/役職」っぽい行を拾う
    for j in range(idx + 1, min(len(blocks), idx + 1 + 30)):
        bj = blocks[j]
        tj = (bj.text or "").strip()
        if not tj:
            continue

        # 近傍制約（縦位置/横位置）
        if abs(bj.top - base_top) > y_limit:
            continue
        if abs(bj.left - base_left) > 220:
            # 横が大きくズレるものは別カラムの可能性
            continue

        if not any(k in tj for k in ["科", "部", "センター", "内科", "外科", "教授", "准教授", "講師", "医長", "部長"]):
            continue

        parts.append(tj)
        taken += 1
        if taken >= max_follow:
            break

    return "\n".join([p for p in parts if p]).strip()

def _pick_vm_row_by_talk(vm_by_name: dict, t) -> dict | None:
    # speaker を優先してVMを引く（displayが壊れてても耐える）
    cand = []
    sp1 = _norm_person_name(getattr(t, "speaker", "") or "")
    sp2 = _norm_person_name(getattr(t, "speaker_display", "") or "")
    if sp1: cand.append(sp1)
    if sp2 and sp2 != sp1: cand.append(sp2)

    for key in cand:
        vm = vm_by_name.get(key)
        if vm:
            return vm
    return None

def apply_vm_hints_from_blocks(blocks: List[TextBlock], payload: DesignJSON, vm_rows: List[dict]) -> DesignJSON:
    """VMはヒントとしてのみ使用。最終値はPPTX(blocks)から取得して埋める（上書きは欠損/矛盾時のみ）。"""
    if not vm_rows or not getattr(payload, "talks", None):
        return payload

    # name -> vm
    vm_by_name: Dict[str, dict] = {}
    for vm in vm_rows:
        n = _norm_person_name(vm.get("案内状掲載 医師名", ""))
        if n:
            vm_by_name[n] = vm

    # facilities list (to detect mismatch)
    facilities = [ (vm.get("案内状掲載 施設名") or "").strip() for vm in vm_rows if (vm.get("案内状掲載 施設名") or "").strip() ]

    for t in payload.talks:
        # sp = _norm_person_name(getattr(t, "speaker_display", "") or getattr(t, "speaker", ""))
        # vm = vm_by_name.get(sp)
        vm = _pick_vm_row_by_talk(vm_by_name, t)
        if not vm:
            continue

        facility = (vm.get("案内状掲載 施設名") or "").strip()
        if not facility:
            continue

        pptx_aff = _join_affiliation_near_facility(blocks, facility)

        if not pptx_aff:
            continue

        cur = (t.affiliation or "").strip()

        # 欠損なら入れる。入ってるが別施設っぽければPPTX値で修正（PPTX由来なのでOK）
        if not cur:
            t.affiliation = pptx_aff
        else:
            if facility not in cur:
                # もしcurが他の施設名を含んでいたら矛盾とみなす
                if facility not in cur and any(f and f in cur for f in facilities):
                    t.affiliation = pptx_aff

    return payload

def fill_missing_from_vm(payload: DesignJSON, vm_rows: List[dict]) -> DesignJSON:
    if not vm_rows or not payload.talks:
        return payload

    vm_by_name = {
        _norm_person_name(r.get("案内状掲載 医師名", "")): r
        for r in vm_rows
        if _norm_person_name(r.get("案内状掲載 医師名", ""))
    }

    for t in payload.talks:
        sp = _norm_person_name(t.speaker or t.speaker_display or "")
        vm = vm_by_name.get(sp)
        if not vm:
            continue

        # タイトル補完（完全空のときのみ）
        if not (t.title or "").strip():
            v = (vm.get("演題") or "").strip()
            if v:
                t.title = v
                t.title_lines = [ln for ln in v.split("\n") if ln.strip()]

        # 所属補完（完全空のときのみ）
        if not (t.affiliation or "").strip():
            aff = _vm_aff_str(vm)
            if aff:
                t.affiliation = aff

    return payload

def build_vm_title_map(vm_rows: list[dict]) -> dict[str, dict]:
    """
    return: { normalized_title: {"speaker": "...", "affiliation": "...", "title": "..."} }
    """
    def norm_title(s: str) -> str:
        s = normalize_space(s or "")
        # 記号ゆれ吸収（必要なら増やす）
        s = s.replace("〜", "～")
        s = re.sub(r"[‐-–—−]", "-", s)
        s = s.replace(" ", "").replace("\u3000", "")
        return s

    title_map: dict[str, dict] = {}
    for r in (vm_rows or []):
        if (r.get("役職") or "") != "演者":
            continue
        title = (r.get("演題") or "").strip()
        sp = norm_name(r.get("案内状掲載 医師名") or "")
        fac = normalize_space(r.get("案内状掲載 施設名") or "")
        dept = normalize_space(r.get("案内状掲載 所属科") or "")
        pos = normalize_space(r.get("案内状掲載 役職") or "")

        aff = " ".join([x for x in [fac, dept, pos] if x]).strip()
        if not title or not sp:
            continue

        key = norm_title(title)
        title_map[key] = {"speaker": sp, "affiliation": aff, "title": title}

    return title_map



def normalize_speaker_display(payload: DesignJSON) -> DesignJSON:
    if not getattr(payload, "talks", None):
        return payload

    for t in payload.talks:
        if (t.speaker or "").strip():
            # speaker_display が空 or 不正なら再生成
            if not (t.speaker_display or "").strip():
                t.speaker_display = add_space_to_jp_name(t.speaker)

    # chair も同様
    if getattr(payload, "chair", None):
        ch = payload.chair
        if (ch.name or "").strip():
            if not (ch.name_display or "").strip():
                ch.name_display = add_space_to_jp_name(ch.name)

    return payload


def prune_talks_heuristic_only(payload: DesignJSON) -> DesignJSON:
    talks = list(payload.talks or [])
    if not talks:
        return payload

    # 不要語
    filtered = []
    for t in talks:
        tt = normalize_space("\n".join(t.title_lines or []).strip() or (t.title or ""))
        if _is_unwanted_talk(tt):
            continue
        filtered.append(t)

    # 重複
    seen = set()
    dedup = []
    for t in filtered:
        key = _norm_title_key(normalize_space("\n".join(t.title_lines or []).strip() or (t.title or "")))
        if not key or key in seen:
            continue
        seen.add(key)
        dedup.append(t)

    payload.talks = dedup
    payload.warnings = sorted(set((payload.warnings or []) + ["talks_pruned_heuristic_only"]))
    return payload

async def pptx_to_json_vm_hint(pptx_path: Path, vm_rows: List[dict], debug_blocks_path: Optional[Path] = None) -> DesignJSON:
    """PPTX優先。VMは精度を上げるヒントとして blocks からの拾い直しにのみ使用し、欠損時のみVMで補完する。"""
    blocks = extract_blocks_any(pptx_path, first_only=True)
    blocks = merge_event_title_blocks_strict(blocks)

    if debug_blocks_path:
        dbg = [
            {
                "text": b.text,
                "left": b.left,
                "top": b.top,
                "width": b.width,
                "height": b.height,
                "max_font_pt": round(b.max_font_pt, 2),
            }
            for b in blocks
        ]
        debug_blocks_path.write_text(json.dumps(dbg, ensure_ascii=False, indent=2), encoding="utf-8")

    speaker_map = extract_speaker_affil_map_by_blocks(blocks)
    time_candidates = extract_time_candidates_from_blocks(blocks)

    draft = parse_blocks_to_design_json(blocks)
    refined = await ai_refine_json(blocks, draft, speaker_map, time_candidates)
    refined = assign_talk_times_by_proximity(blocks, refined)

    refined = apply_vm_hints_from_blocks(blocks, refined, vm_rows)
    refined = fill_missing_from_vm(refined, vm_rows)

    # ★VM演題がある時だけVM prune
    vm_titles = _vm_speaker_titles(vm_rows)
    if vm_titles:
        refined = prune_talks_using_vm_titles(refined, vm_rows)
    else:
        refined = prune_talks_heuristic_only(refined)
        
    print("after VM hint application:")
    print(refined)


    # # chair
    # chair = extract_chair_from_blocks(blocks, speaker_map)
    # if chair and (chair.get("name") or "").strip():
    #     refined.chair.name = chair["name"]
    #     refined.chair.name_display = chair.get("name_display") or refined.chair.name_display
    #     if not (refined.chair.affiliation or "").strip():
    #         refined.chair.affiliation = chair.get("affiliation","").strip()

    refined = fill_chair_affiliation_from_vm_hint(refined, blocks, vm_rows)
   
    # ★それでも座長所属が空なら blocks から拾う（欠損時のみ）
    refined = fill_chair_affiliation_from_blocks(refined, blocks)

   
    
    # ★speaker_display を必ず作る（VMなしでも）
    refined = normalize_speaker_display(refined)

    refined = fill_datetime_parts(refined, blocks)

    
    # if refined.confidence < draft.confidence:
    #     refined.warnings = list(set(refined.warnings + ["ai_lower_confidence"]))
    #     return draft
    
    return refined






async def render_png(payload: DesignJSON, out_path: Path, debug_html_path: Path):
    global _cached_template
    if _cached_template is None:
        _cached_template = TEMPLATE_PATH.read_text(encoding="utf-8")

    async with async_playwright() as p:
        browser = await p.chromium.launch(args=["--no-sandbox", "--disable-dev-shm-usage"],)
        page = await browser.new_page(viewport=BASE_VIEWPORT)

        page.on("pageerror", lambda e: print("[pageerror]", e))
        page.on("console", lambda m: print("[console]", m.type, m.text))

        # await page.set_content(_cached_template, wait_until="domcontentloaded")
        await page.goto(TEMPLATE_PATH.resolve().as_uri(), wait_until="domcontentloaded")
        await page.evaluate("() => document.fonts && document.fonts.ready")

        data_json = payload.model_dump_json() if hasattr(payload, "model_dump_json") else payload.json(ensure_ascii=False)

        data_obj = json.loads(data_json)

        # ★ここで初期値のみ“精密組版”して data を書き換える
        # data_obj = await page.evaluate(TYPESET_JS, {"data": data_obj})

        # ★ここが重要：DATA注入 → render呼び出し
        await page.evaluate(
            """(data) => {
                window.__DATA__ = data;
                if (typeof window.__render === "function") window.__render();
            }""",
            data_obj  # ← evaluateに渡すのは object が安全（文字列直埋めより事故らない）
        )

        # render完了フラグを待つ（render内で data-ready=1 が立つ）
        await page.wait_for_selector('html[data-ready="1"]', timeout=30000)
        await page.wait_for_selector(".wrap", timeout=30000)

        wrap = page.locator(".wrap")

        for _ in range(60):
            box = await wrap.bounding_box()
            if box and box["height"] and box["height"] > 10:
                break
            await page.wait_for_timeout(100)
        else:
            html = await page.content()
            debug_html_path.write_text(html, encoding="utf-8")
            raise RuntimeError(f"wrap bounding box not ready; wrote {debug_html_path}")

        box = await wrap.bounding_box()
        h = min(int(box["height"]), MAX_HEIGHT)

        await page.set_viewport_size({"width": 600, "height": max(h, 1)})
        await wrap.screenshot(
    path=str(out_path),
    type="jpeg",
    quality=100  # 0〜100（PNGには無い）
)


        await browser.close()




# ---------------- App ----------------
app = FastAPI(title="PPTX → JSON → HTML → jpg (Keep Newlines + Split ~...~)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://<your-vercel-app>.vercel.app",
        "http://localhost:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/fonts", StaticFiles(directory=str(APP_DIR / "fonts")), name="fonts")

@app.on_event("startup")
async def startup():
    global _cached_template

    init_db()

    if not TEMPLATE_PATH.exists():
        raise RuntimeError(f"template.html not found: {TEMPLATE_PATH}")
    _cached_template = TEMPLATE_PATH.read_text(encoding="utf-8")

    # Playwright browser path (Render disk)
    browsers_path = os.getenv("PLAYWRIGHT_BROWSERS_PATH")
    if browsers_path:
        Path(browsers_path).mkdir(parents=True, exist_ok=True)

    # Ensure Chromium exists
    try:
        subprocess.check_call(["python", "-m", "playwright", "install", "chromium"])
    except Exception as e:
        raise RuntimeError(f"Playwright install failed: {e}")


@app.post("/upload")
async def upload(pptx: UploadFile = File(...)):
    if not pptx.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Only .pptx is supported")
    session_id = new_session_id()
    job_id = uuid.uuid4().hex
    p = job_paths(job_id)
    filename = pptx.filename
    
    p["pptx"].write_bytes(await pptx.read())

    try:
        payload = await pptx_to_json(p["pptx"], debug_blocks_path=p["debug_blocks"])
        p["json"].write_text(dump_json(payload), encoding="utf-8")
        await render_png(payload, p["jpg"], p["debug_html"])

        upsert_job_ok(job_id, filename, payload, session_id)

        return JSONResponse({
            "sessionId": session_id,
            "jobId": job_id,
            "json": payload.model_dump(exclude_none=True) if hasattr(payload, "model_dump") else json.loads(payload.json(ensure_ascii=False)),
            "previewUrl": f"/preview/{job_id}.jpg",
            "downloadUrl": f"/download/{job_id}.jpg",
            "debugBlocksUrl": f"/debug/{job_id}/blocks.json",
        })
    except Exception as e:
        upsert_job_error(job_id, filename, str(e))
        raise

def _sse(event: str, data: dict) -> str:
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"

@app.post("/upload/batch/stream")
async def upload_batch_stream(
    files: List[UploadFile] = File(...),
    eventIds: List[str] = Form(...),
):
    session_id = new_session_id()

    if not files:
        raise HTTPException(400, "files is empty")
    if len(eventIds) != len(files):
        raise HTTPException(400, f"eventIds length mismatch: {len(eventIds)} != {len(files)}")

    total = len(files)

    # ✅ ここで UploadFile を全部 “生きてるうちに” 退避する
    buffered: List[Dict[str, Any]] = []
    # sessionごとの temp dir（好きな場所でOK）
    session_dir = Path("jobs") / f"session_{session_id}"
    session_dir.mkdir(parents=True, exist_ok=True)

    for i, f in enumerate(files):
        filename = f.filename or f"file_{i}"
        suffix = Path(filename).suffix.lower()

        item = {"index": i, "filename": filename, "suffix": suffix, "eventId": (eventIds[i] or "").strip()}

        if suffix not in [".pptx", ".pdf"]:
            item["precheck"] = {"ok": False, "error": "not_supported_file"}
            buffered.append(item)
            continue
        if not item["eventId"]:
            item["precheck"] = {"ok": False, "error": "event_id_required"}
            buffered.append(item)
            continue

        try:
            data = await f.read()              # ✅ return前に読む
            in_path = session_dir / f"{i}_{uuid.uuid4().hex}{suffix}"
            in_path.write_bytes(data)
            item["precheck"] = {"ok": True}
            item["in_path"] = str(in_path)
        except Exception as e:
            item["precheck"] = {"ok": False, "error": f"upload_read_failed: {e}"}
        finally:
            # 任意：明示的に閉じておく（なくてもOK）
            try:
                await f.close()
            except Exception:
                pass

        buffered.append(item)

    async def gen():
        yield _sse("start", {"sessionId": session_id, "total": total})

        try:
            yield _sse("phase", {"phase": "sheet_open", "message": "スプレッドシート接続中…"})
            # ---- Spreadsheet open (once) ----
            scope = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ]
            credentials = get_gsa_credentials(scope)
            gc = gspread.authorize(credentials)

            SPREADSHEET_KEY = "1hiV0Ve2cnYyrPkBuZcZIcLWeAnJ-ucNiB0P4owZpXug"
            workbook = gc.open_by_key(SPREADSHEET_KEY)

            PRESENCE_SHEETS = ["VM(GWET)", "VM(例外)", "VM(本社)"]
            VM_SHEET = "演題演者（VM）"
            PRESENCE_HEADER_ROW = 2
            VM_HEADER_ROW = 1
            PRESENCE_ID_COL = "システムID"

            yield _sse("phase", {"phase": "precheck", "message": "事前チェック中…"})

            valid_event_ids = []
            for it in buffered:
                if it["precheck"]["ok"]:
                    valid_event_ids.append(it["eventId"])

            ws_map = _retry_gspread(lambda: {ws.title: ws for ws in workbook.worksheets()})

            yield _sse("phase", {"phase": "batch_fetch", "message": "VM/Presence一括取得中…"})
            presence_rows_by_event, vm_rows_by_event, _ = batch_fetch_system_and_vm_rows(
                workbook,
                ws_map=ws_map,
                event_ids=valid_event_ids,
                presence_sheets=PRESENCE_SHEETS,
                presence_header_row=PRESENCE_HEADER_ROW,
                presence_id_col=PRESENCE_ID_COL,
                vm_sheet=VM_SHEET,
                vm_header_row=VM_HEADER_ROW,
                vm_id_col_candidates=["講演会ID"],
                col_end="N",
            )

            # （あなたの _parse_ymd / presence_rows_by_file_index / vm_rows_by_file_index のロジックは
            #    buffered を元に組み直すのが一番安全。ここでは “生成部分” の直しだけ見せます）

            yield _sse("phase", {"phase": "processing", "message": "生成を開始します…"})
            out: List[Dict[str, Any]] = []

            for it in buffered:
                i = it["index"]
                filename = it["filename"]
                event_id = it["eventId"]

                yield _sse("item_start", {"index": i, "filename": filename, "eventId": event_id})

                if not it["precheck"]["ok"]:
                    err = it["precheck"]["error"]
                    out.append({"filename": filename, "ok": False, "error": err})
                    yield _sse("item_done", {"index": i, "filename": filename, "ok": False, "error": err})
                    continue

                # ✅ もう UploadFile は触らない。退避したパスだけ使う
                in_path = Path(it["in_path"])

                # presence/vm は event_id から引く（ここはあなたの既存ロジックに合わせてOK）
                presence_rows = presence_rows_by_event.get(event_id, []) or []
                if not presence_rows:
                    out.append({"filename": filename, "ok": False, "error": "event_id_not_found"})
                    yield _sse("item_done", {"index": i, "filename": filename, "ok": False, "error": "event_id_not_found"})
                    continue
                vm_rows = vm_rows_by_event.get(event_id, []) or []

                job_id = uuid.uuid4().hex
                p = job_paths(job_id)

                try:
                    payload = await pptx_to_json_vm_hint(in_path, vm_rows, debug_blocks_path=p.get("debug_blocks"))
                    payload = normalize_for_render(payload)
                    payload = post_format_design_initial(payload)
                    payload = await apply_precise_typeset_initial(payload)
                    payload = ensure_display_fields(payload)

                    payload.region = presence_rows[0].get("VP/PH/ONC", "")
                    payload.unit = presence_rows[0].get("取得単位：フラグメントデザインへの内容記載", "")
                    payload.event_id = event_id

                    payload.talks = sorted(payload.talks or [], key=lambda x: _time_start_minutes(getattr(x, "time", "")))
                    p["json"].write_text(dump_json(payload), encoding="utf-8")
                    await render_png(payload, p["jpg"], p["debug_html"])

                    upsert_job_ok(job_id, filename, payload, session_id, event_id)

                    out.append({"filename": filename, "jobId": job_id, "ok": True})
                    yield _sse("item_done", {"index": i, "filename": filename, "ok": True, "jobId": job_id})

                except Exception as e:
                    tb = traceback.format_exc()
                    print("[upload/batch error]", filename, job_id)
                    print(tb)
                    out.append({"filename": filename, "jobId": job_id, "ok": False, "error": str(e)})
                    yield _sse("item_done", {"index": i, "filename": filename, "ok": False, "jobId": job_id, "error": str(e)})

            ok_count = sum(1 for r in out if r.get("ok"))
            yield _sse("done", {"sessionId": session_id, "count": ok_count, "results": out})

        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            yield _sse("fatal", {"message": str(e)})

    return StreamingResponse(
        gen(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )

@app.post("/upload/batch")
async def upload_batch(
    files: List[UploadFile] = File(...),
    eventIds: List[str] = Form(...),  # filesと同じ順
):
    session_id = new_session_id()

    if not files:
        raise HTTPException(400, "files is empty")
    if len(eventIds) != len(files):
        raise HTTPException(400, f"eventIds length mismatch: {len(eventIds)} != {len(files)}")

    # ---- Spreadsheet open (once) ----
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = get_gsa_credentials(scope)
    gc = gspread.authorize(credentials)

    SPREADSHEET_KEY = "1hiV0Ve2cnYyrPkBuZcZIcLWeAnJ-ucNiB0P4owZpXug"
    workbook = gc.open_by_key(SPREADSHEET_KEY)

    # ---- sheets ----
    PRESENCE_SHEETS = ["VM(GWET)", "VM(例外)", "VM(本社)"]
    VM_SHEET = "演題演者（VM）"

    PRESENCE_HEADER_ROW = 2
    VM_HEADER_ROW = 1  # ←VMが1行ヘッダなら1。あなたの現状は2が多そうなので2推奨

    PRESENCE_ID_COL = "システムID"
    VM_PRESENCE_ID_COL = "講演会ID"  # フォールバック

    # # ---- build presence index (once) ----
    # presence_index = preload_system_id_index(
    #     workbook,
    #     PRESENCE_SHEETS,
    #     header_row=PRESENCE_HEADER_ROW,
    #     id_col_name=PRESENCE_ID_COL,
    # )

    # # ---- build VM index (once) ----
    # vm_index = None
    # vm_index_err = None
    # try:
    #     vm_index = build_system_id_to_rows(
    #             workbook,
    #             sheet_name=VM_SHEET,
    #             header_row=VM_HEADER_ROW,
    #             id_col_name=VM_PRESENCE_ID_COL,
    #         )
    #     vm_index_err = None
    
    # except Exception as e:
    #     vm_index_err = e
        

    # if vm_index is None:
    #     raise HTTPException(
    #         status_code=500,
    #         detail={"code": "vm_sheet_index_failed", "message": str(vm_index_err)},
    #     )

    # ---- まず全件チェック＆データ取得（1件でもNGなら422）----
    # 0) precheck（pptx / event_id）
    valid_event_ids = []
    precheck = [None] * len(files)
    for i, f in enumerate(files):
        filename = f.filename or f"file_{i}"
        if not filename.lower().endswith(".pptx") and not filename.lower().endswith(".pdf"):
            precheck[i] = {"ok": False, "error": "not_supported_file"}
            continue
        eid = (eventIds[i] or "").strip()
        if not eid:
            precheck[i] = {"ok": False, "error": "event_id_required"}
            continue
        precheck[i] = {"ok": True}
        valid_event_ids.append(eid)

    ws_map = _retry_gspread(lambda: {ws.title: ws for ws in workbook.worksheets()})

    # 1) presence+VM を一括取得（ここが速度の肝）
    presence_rows_by_event, vm_rows_by_event, _ = batch_fetch_system_and_vm_rows(
        workbook,
        ws_map=ws_map,
        event_ids=valid_event_ids,
        presence_sheets=PRESENCE_SHEETS,
        presence_header_row=PRESENCE_HEADER_ROW,
        presence_id_col=PRESENCE_ID_COL,                 # "システムID"
        vm_sheet=VM_SHEET,
        vm_header_row=VM_HEADER_ROW,
        vm_id_col_candidates=["講演会ID"],  # あなたのVMは講演会IDっぽいので先に
        col_end="N",  # 申請日(A)〜演題(N) までなら N でOK
    )

    print("================= batch fetch done =================")
    print(presence_rows_by_event)
    print('---------------------------------------------------')

    results = [None] * len(files)
    presence_rows_by_file_index = [None] * len(files)
    vm_rows_by_file_index = [None] * len(files)

    # 2) files順に合わせて、申請日でVMを絞る
    
    def _parse_ymd(s: str):
        m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", str(s or ""))
        if not m:
            return None
        y, mo, d = map(int, m.groups())
        return (y, mo, d)
    # 2) files順に合わせて…
    for i, f in enumerate(files):
        filename = f.filename or f"file_{i}"
        if not precheck[i]["ok"]:
            results[i] = {"filename": filename, "ok": False, "error": precheck[i]["error"]}
            continue
        eid = (eventIds[i] or "").strip()
        presence_rows = presence_rows_by_event.get(eid, [])
        if not presence_rows:
            results[i] = {"filename": filename, "ok": False, "error": "event_id_not_found"}
            continue
        presence_last = max(presence_rows, key=lambda r: int(r.get("_row") or 0))
        presence_apply_day = _parse_ymd((presence_last.get("申請日") or "").strip())
        # VMは任意
        vm_rows_all = vm_rows_by_event.get(eid, [])  # ← 全量
        vm_rows = []
        if vm_rows_all:
            vm_by_day: dict[tuple[int,int,int] | None, list[dict]] = {}
            for r in vm_rows_all:
                k = _parse_ymd((r.get("申請日") or "").strip())
                vm_by_day.setdefault(k, []).append(r)
            # 1) presenceと同日
            vm_rows = vm_by_day.get(presence_apply_day, [])
            # 2) フォールバック
            if not vm_rows:
                days = sorted([d for d in vm_by_day.keys() if d is not None])
                if days:
                    if presence_apply_day is not None:
                        le = [d for d in days if d <= presence_apply_day]
                        pick = le[-1] if le else days[-1]
                    else:
                        pick = days[-1]
                    vm_rows = vm_by_day[pick]
        presence_rows_by_file_index[i] = presence_rows
        vm_rows_by_file_index[i] = vm_rows
        results[i] = {"filename": filename, "ok": True}

    failed = [r for r in results if r and not r.get("ok")]
    if failed:
        raise HTTPException(
            status_code=422,
            detail={
                "code": "batch_failed",
                "message": "1件以上のファイルでエラーが発生しました",
                "results": results,
            },
        )

    # ---- ここから生成（全件OKのときだけ）----
    out: List[Dict[str, Any]] = []

    for i, f in enumerate(files):
        filename = f.filename or f"file_{i}"
        event_id = (eventIds[i] or "").strip()

        vm_rows = vm_rows_by_file_index[i] or []
        presence_rows = presence_rows_by_file_index[i] or []
        # presence_best = pick_best_presence_row(presence_rows)  # 1つに絞るなら
        print('--------------------------------')
        print(f'batch] processing{vm_rows}')
        print(f"[batch] processing file {i}: {presence_rows}, event_id={event_id}")
        print('--------------------------------')
        job_id = uuid.uuid4().hex
        p = job_paths(job_id)

        try:
            # p["pptx"].write_bytes(await f.read())


            suffix = Path(filename).suffix.lower()  # ".pptx" or ".pdf"
            in_path = p["dir"] / f"input{suffix}"
            in_path.write_bytes(await f.read())
            

            payload = await pptx_to_json_vm_hint(in_path, vm_rows, debug_blocks_path=p.get("debug_blocks"))
            print(f"[batch] parsed json: {payload}")
            payload = normalize_for_render(payload)
            payload = post_format_design_initial(payload)
            payload = await apply_precise_typeset_initial(payload)
            print(f"[batch] normalized json: {payload}")
            payload = fill_datetime_parts(payload)
            payload = ensure_display_fields(payload)
            
            print(f"[batch] vm corrected json: {payload}")

            # ★3シート側の値も payload / DB に入れる（どっちでも）
            # 例: payloadに埋め込む
            # try:
            #     payload.meta = getattr(payload, "meta", {}) or {}
            #     payload.meta["presence_rows"] = presence_rows
            #     # payload.meta["presence_best"] = presence_best
            # except Exception:
            #     # payloadがdictならこちら
            #     if isinstance(payload, dict):
            #         payload.setdefault("meta", {})
            #         payload["meta"]["presence_rows"] = presence_rows
            #         # payload["meta"]["presence_best"] = presence_best

            payload.region = presence_rows[0]["VP/PH/ONC"] if presence_rows else ""
            payload.unit = presence_rows[0]["取得単位：フラグメントデザインへの内容記載"] if presence_rows else ""
            payload.event_id = event_id
            payload.talks = sorted(payload.talks or [], key=lambda x: _time_start_minutes(getattr(x, "time", "")))

            p["json"].write_text(dump_json(payload), encoding="utf-8")
            await render_png(payload, p["jpg"], p["debug_html"])
            
            print(f"[batch] payload{payload}") 

            # DBにも保存したいなら upsert_job_ok の引数拡張 or payload内metaで保存
            upsert_job_ok(job_id, filename, payload, session_id, event_id)

            out.append({"filename": filename, "jobId": job_id, "ok": True})

        except Exception as e:
            tb = traceback.format_exc()
            print("[upload/batch error]", filename, job_id)
            print(tb)
            out.append({"filename": filename, "jobId": job_id, "ok": False, "error": str(e)})

    ok_count = sum(1 for r in out if r.get("ok"))
    return {"sessionId": session_id, "count": ok_count, "results": out}


@app.post("/render")
async def render(req: RenderReq):
    p = job_paths(req.jobId)

    # --- DB: lock/メタ取得 ---
    with db_connect() as con:
        row = con.execute(
            "SELECT locked, filename, session_id, event_id FROM jobs WHERE job_id=%s",
            (req.jobId,),
        ).fetchone()
    if not row:
        raise HTTPException(404, "job not found")

    if bool(row.get("locked")):
        raise HTTPException(400, "This job is locked.")

    filename = row.get("filename") or ""
    session_id = row.get("session_id") or ""
    event_id = row.get("event_id") or ""

    payload = req.design  # DesignJSON (Pydantic)

    # --- ファイル保存（従来通り）---
    p["json"].write_text(dump_json(payload), encoding="utf-8")
    await render_png(payload, p["jpg"], p["debug_html"])

    # --- DB upsert（Supabase版 upsert_job_ok を呼ぶ想定）---
    upsert_job_ok(req.jobId, filename, payload, session_id, event_id)

    # --- response ---
    data = payload.model_dump(exclude_none=True) if hasattr(payload, "model_dump") else json.loads(payload.json(ensure_ascii=False))
    return JSONResponse({
        "jobId": req.jobId,
        "json": data,
        "warnings": getattr(payload, "warnings", None),
        "previewUrl": f"/preview/{req.jobId}.jpg",
        "downloadUrl": f"/download/{req.jobId}.jpg",
    })

def _parse_date_start(s: str) -> Optional[datetime]:
    s = (s or "").strip()
    if not s:
        return None
    # "YYYY-MM-DD" をUTC 00:00:00として扱う（必要ならJSTに変更）
    return datetime.strptime(s, "%Y-%m-%d").replace(tzinfo=timezone.utc)

def _parse_date_end(s: str) -> Optional[datetime]:
    s = (s or "").strip()
    if not s:
        return None
    # inclusive end にしたいなら 23:59:59.999999
    dt = datetime.strptime(s, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    return dt.replace(hour=23, minute=59, second=59, microsecond=999999)

@app.get("/jobs")
async def list_jobs(
    q: str = "",
    status: Optional[Literal["ok", "error"]] = None,
    warning: str = "",
    manual: Optional[bool] = None,
    locked: Optional[bool] = None,
    min_conf: Optional[float] = None,
    max_conf: Optional[float] = None,
    created_from: str = "",
    created_to: str = "",
    page: int = 1,
    page_size: int = 30,
    order: Literal["updated_desc", "created_desc"] = "updated_desc",
):
    page = max(page, 1)
    page_size = min(max(page_size, 1), 200)
    offset = (page - 1) * page_size

    where: List[str] = []
    params: List[Any] = []

    cf = _parse_date_start(created_from)
    ct = _parse_date_end(created_to)
    if cf:
        where.append("created_at >= %s")
        params.append(cf)
    if ct:
        where.append("created_at <= %s")
        params.append(ct)

    if status:
        where.append("status = %s")
        params.append(status)

    if q and q.strip():
        where.append("(filename ILIKE %s OR title ILIKE %s OR organizer ILIKE %s OR event_id ILIKE %s)")
        like = f"%{q.strip()}%"
        params.extend([like, like, like, like])

    if manual is not None:
        where.append("manual_override = %s")
        params.append(bool(manual))

    if locked is not None:
        where.append("locked = %s")
        params.append(bool(locked))

    if min_conf is not None:
        where.append("confidence >= %s")
        params.append(float(min_conf))

    if max_conf is not None:
        where.append("confidence <= %s")
        params.append(float(max_conf))

    # jsonb array contains: warnings_json @> '["missing_x"]'
    if warning and warning.strip():
        where.append("warnings_json @> %s::jsonb")
        params.append(json.dumps([warning.strip()], ensure_ascii=False))

    where_sql = ("WHERE " + " AND ".join(where)) if where else ""

    order_sql = "ORDER BY updated_at DESC" if order == "updated_desc" else "ORDER BY created_at DESC"

    with db_connect() as con:
        total = con.execute(
            f"SELECT COUNT(*) AS c FROM jobs {where_sql}",
            params,
        ).fetchone()["c"]

        rows = con.execute(
            f"""
            SELECT * FROM jobs
            {where_sql}
            {order_sql}
            LIMIT %s OFFSET %s
            """,
            params + [page_size, offset],
        ).fetchall()

    items = [row_to_job_item(r) for r in rows]
    return {
        "page": page,
        "pageSize": page_size,
        "total": total,
        "items": items,
    }




# ------------------------------------------------------------
# ジョブメタ更新（manual_override / note / locked）
# ------------------------------------------------------------

class JobPatch(BaseModel):
    manual_override: Optional[bool] = None
    note: Optional[str] = None
    locked: Optional[bool] = None

@app.get("/job/{job_id}")
async def get_job(job_id: str):
    p = job_paths(job_id)

    # 1) まずファイル
    if p["json"].exists():
        data = json.loads(p["json"].read_text(encoding="utf-8"))
        return JSONResponse({"jobId": job_id, "json": data})

    # 2) ファイルが無い場合：DBに存在するか確認（存在しないなら404）
    with db_connect() as con:
        row = con.execute("SELECT job_id FROM jobs WHERE job_id=%s", (job_id,)).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="job not found")

    # # 3) DBにはあるがファイルが無い → 409とかで「再生成して」系にするのが親切
    # raise HTTPException(status_code=409, detail="job exists but json file is missing")



@app.get("/preview/{job_id}.jpg")
async def preview(job_id: str):
    p = job_paths(job_id)
    if not p["jpg"].exists():
        raise HTTPException(status_code=404, detail="preview not found")
    return FileResponse(p["jpg"], media_type="image/jpg")


@app.get("/download/{job_id}.jpg")
async def download(job_id: str):
    p = job_paths(job_id)
    file_path = p["jpg"]
    if not file_path.exists():
        raise HTTPException(404, "preview not found")

    event_id = job_id

    # 1) jsonから
    if p["json"].exists():
        try:
            data = json.loads(p["json"].read_text(encoding="utf-8"))
            event_id = (data.get("event_id") or data.get("eventId") or job_id).strip()
        except Exception:
            pass
    else:
        # 2) DBから（フォールバック）
        with db_connect() as con:
            row = con.execute("SELECT event_id FROM jobs WHERE job_id=%s", (job_id,)).fetchone()
        if row and (row.get("event_id") or "").strip():
            event_id = row["event_id"].strip()

    filename = f"{event_id}_招聘.jpg"
    return FileResponse(file_path, media_type="image/jpeg", filename=filename)





@app.get("/debug/{job_id}/blocks.json")
async def debug_blocks(job_id: str):
    p = job_paths(job_id)
    if not p["debug_blocks"].exists():
        raise HTTPException(status_code=404, detail="debug blocks not found")
    return FileResponse(p["debug_blocks"], media_type="application/json")


# ------------------------------------------------------------
# 選択ジョブのPNGをまとめてZIP（納品用）
# ------------------------------------------------------------


def sanitize_basename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[\\/]", "_", s)
    s = re.sub(r"\.pptx$", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    return s or "file"

def unique_name(base: str, used: dict[str, int]) -> str:
    # baseは拡張子なし
    n = used.get(base, 0) + 1
    used[base] = n
    return base if n == 1 else f"{base} ({n})"


class ExportReq(BaseModel):
    jobIds: List[str] = Field(default_factory=list)
    nameMode: Literal["jobId", "filename"] = "filename"  # zip内ファイル名
    includeJson: bool = False

@app.post("/jobs/export.zip")
async def export_zip(req: ExportReq):
    if not req.jobIds:
        raise HTTPException(400, "jobIds is empty")

    # job_id -> filename（Postgres版）
    with db_connect() as con:
        rows = con.execute(
            "SELECT job_id, filename FROM jobs WHERE job_id = ANY(%s)",
            (req.jobIds,),  # ←タプルで包むのが重要
        ).fetchall()

    mp = {r["job_id"]: (r.get("filename") or "") for r in rows}

    export_id = f"export_{int(time.time())}_{uuid.uuid4().hex}"
    zip_path = EXPORT_DIR / f"{export_id}.zip"

    used = {}
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for job_id in req.jobIds:
            p = job_paths(job_id)
            if not p["jpg"].exists():
                continue

            base0 = job_id if req.nameMode == "jobId" else (mp.get(job_id) or job_id)
            base0 = sanitize_basename(base0)

            used[base0] = used.get(base0, 0) + 1
            base = base0 if used[base0] == 1 else f"{base0} ({used[base0]})"

            z.write(p["jpg"], arcname=f"{base}_招聘.jpg")
            if req.includeJson and p["json"].exists():
                z.write(p["json"], arcname=f"{base}.json")

    return FileResponse(str(zip_path), media_type="application/zip", filename="export.zip")



class JobDeleteReq(BaseModel):
    delete_files: bool = True
    force: bool = False  # lockedでも消したい場合だけTrue

@app.delete("/job/{job_id}")
async def delete_job(job_id: str, req: JobDeleteReq = JobDeleteReq()):
    # まずDBを確認
    with db_connect() as con:
        row = con.execute(
            "SELECT locked FROM jobs WHERE job_id=%s",
            (job_id,),
        ).fetchone()
        if not row:
            raise HTTPException(404, "job not found")

        if bool(row.get("locked")) and not req.force:
            raise HTTPException(400, "This job is locked.")

        # DB削除
        con.execute("DELETE FROM jobs WHERE job_id=%s", (job_id,))

    # ファイル削除（任意）
    deleted_files = False
    if req.delete_files:
        try:
            p = job_paths(job_id)
            # job_paths(job_id)["dir"] がある前提（無ければ _data/job_id に合わせてください）
            shutil.rmtree(p["dir"], ignore_errors=True)
            deleted_files = True
        except Exception:
            # ファイル削除失敗でもDBは消えてるので、ここは 200 で返す方が運用楽
            deleted_files = False

    return {"ok": True, "jobId": job_id, "deletedFiles": deleted_files}

