"""
Regeneron 申請書 Excel のパースサービス

申請書フォーマット (講演会フラグメント作成申請書_v3.0):
  シート名: 講演会フラグメント申請
  行1-2: タイトル/サブヘッダー (スキップ)
  行3: カラムヘッダー
  行4: 記入例 (スキップ)
  行5+: データ行

カラム (1-indexed):
  1. # (スキップ)
  2. 区分① (MKT/Sales)
  3. 区分② (講演会/学会共催セミナー)
  4. 講演会ID (Veeva Event ID)
  5. 開催日
  6. 時間
  7. 講演会名
  8. 領域
  9. 製品
  10. 視聴予約ページのURL (ハイパーリンク対応)
  11. 会議ID
  12. パスコード
  13. 都道府県
  14. 他エリアへの案内可否
  15. 使用期限
  16. 元の案内状の文書番号
  17. 元の案内状の承認日
  18. ドキュメント名
"""

from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, time
from io import BytesIO
from typing import Any

from openpyxl import load_workbook


_FRAGMENT_SHEET_NAME = "講演会フラグメント申請"
_DATA_START_ROW = 5   # 1-indexed。行4 は記入例なのでスキップ
_WEEKDAY_INDEX = {"月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6}
_ZEN_TO_HALF = str.maketrans("０１２３４５６７８９：", "0123456789:")


def _is_url(text: str) -> bool:
    """単純な URL 判定"""
    return text.startswith("http://") or text.startswith("https://")


def _clean_text(value: Any) -> str:
    return str(value or "").translate(_ZEN_TO_HALF).strip()


def _infer_year(month: int, day: int, weekday: str = "") -> int:
    today = datetime.now().date()
    candidates: list[tuple[int, int]] = []
    # Events are within about a year; a mistyped weekday must not push the date years ahead.
    for year in range(today.year - 1, today.year + 2):
        try:
            d = date(year, month, day)
        except ValueError:
            continue
        if weekday and d.weekday() != _WEEKDAY_INDEX.get(weekday):
            continue
        candidates.append((abs((d - today).days), year))
    if candidates:
        return min(candidates)[1]
    return _infer_year(month, day) if weekday else today.year


def _format_date(value: Any) -> str:
    """
    openpyxl が返す値 (datetime / str / None) を "YYYY-MM-DD" 文字列に変換する。
    Excel の日付シリアル値は openpyxl が自動で datetime に変換するため int は来ない想定。
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    text = _clean_text(value)
    if not text:
        return ""
    # "2026年 1 2月 10日" has stray spaces inside the month/day digits.
    text = re.sub(r"(\d)\s+(?=\d+\s*[年月日])", r"\1", text)

    m = re.search(r"(20\d{2})\s*[年/.-]\s*(\d{1,2})\s*[月/.-]\s*(\d{1,2})", text)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except ValueError:
            pass

    # US-style text such as "12/15/2026" (month/day/year).
    m = re.search(r"(?<!\d)(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(20\d{2})(?!\d)", text)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(1)), int(m.group(2))).strftime("%Y-%m-%d")
        except ValueError:
            pass

    m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日(?:\s*[(（]([月火水木金土日])[)）])?", text)
    if m:
        month = int(m.group(1))
        day = int(m.group(2))
        year = _infer_year(month, day, m.group(3) or "")
        try:
            return date(year, month, day).strftime("%Y-%m-%d")
        except ValueError:
            pass

    m = re.search(r"(?<!\d)(\d{1,2})\s*[/.]\s*(\d{1,2})(?!\d)", text)
    if m:
        month = int(m.group(1))
        day = int(m.group(2))
        year = _infer_year(month, day)
        try:
            return date(year, month, day).strftime("%Y-%m-%d")
        except ValueError:
            pass

    return text.split()[0].strip()


def _format_time(value: Any) -> str:
    """開催時間から日付部分を取り除き、HH:MM または HH:MM-HH:MM に整える。複数枠は改行でつなぐ。"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return ""
        return value.strftime("%H:%M")
    if isinstance(value, time):
        return value.strftime("%H:%M")

    text = _clean_text(value)
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)

    time_part = r"(\d{1,2})\s*(?::|時)\s*(\d{2})(?:\s*分)?"
    range_re = re.compile(time_part + r"\s*[-~～〜－–—]\s*" + time_part)
    ranges = range_re.findall(text)
    if ranges:
        return "\n".join(f"{int(h1):02d}:{m1}-{int(h2):02d}:{m2}" for h1, m1, h2, m2 in ranges)

    m = re.search(time_part, text)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"

    without_date = re.sub(r"(?:20\d{2}\s*[年/.-]\s*)?\d{1,2}\s*月\s*\d{1,2}\s*日(?:\s*[(（][月火水木金土日][)）])?", "", text)
    without_date = re.sub(r"20\d{2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{1,2}", "", without_date)
    return without_date.strip()


DATE_IN_NAME = (r'(?:20\d{2}\s*[年/.\-]\s*\d{1,2}\s*[月/.\-]\s*\d{1,2}|\d{1,2}\s*[月/]\s*\d{1,2})\s*日?(?:\s*[(（][月火水木金土日][)）])?'
                r'|(?:20)?\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])')
NAME_SEPARATOR = r'[\s_\-－:：]*'


def one_line(text):
    return re.sub(r'[ \t　]*[\r\n]+[ \t　]*', ' ', text).strip()


def strip_event_date(name):
    # The naming rule adds the date itself, so a leading/trailing date in 講演会名 is removed.
    stripped = re.sub(rf'^(?:{DATE_IN_NAME})(?!\d){NAME_SEPARATOR}', '', name)
    stripped = re.sub(rf'(?<!\d){NAME_SEPARATOR}(?:{DATE_IN_NAME})$', '', stripped).strip()
    return stripped or name


def _calc_fragment_name(kubun1: str, kubun2: str, event_date: str,
                        event_name: str, prefecture: str, area_guide: str) -> str:
    """
    フラグメント名命名規則:
      MKT + 講演会         → MKT_〇月〇日_講演会名_他エリア案内可/不可
      MKT + 学会共催セミナー → MKT_〇月〇日_講演会名
      Sales                → Sales_〇月〇日_都道府県_講演会名_他エリア案内可/不可
    """
    date_str = ""
    if event_date:
        try:
            d = datetime.strptime(event_date[:10], "%Y-%m-%d")
            date_str = f"{d.year}年{d.month}月{d.day}日"
        except ValueError:
            date_str = event_date

    if kubun1 == "MKT" and kubun2 == "学会共催セミナー":
        return f"MKT_{date_str}_{event_name}"
    if kubun1 == "MKT":
        return f"MKT_{date_str}_{event_name}_他エリア案内{area_guide}"
    # Sales
    return f"Sales_{date_str}_{prefecture}_{event_name}_他エリア案内{area_guide}"


def application_fragment_name(fields: dict[str, Any]) -> str:
    return one_line(_calc_fragment_name(
        fields.get("kubun1", ""), fields.get("kubun2", ""), fields.get("eventDate", ""),
        strip_event_date(one_line(fields.get("eventName", ""))), fields.get("prefecture", ""), fields.get("areaGuide", ""),
    ))


# (field, default column for the v3.0+ layout, header keyword)
_COLUMNS = (
    # v4.0 splits the MKT target ("NSCLC/CSCC/Cervical") from 講演会/学会共催セミナー, so match on the choices, not "区分②".
    ("kubun1", 2, "区分①"), ("kubun2", 3, "講演会/学会共催"), ("indication", None, "nsclc/cscc"), ("event_id", 4, "講演会id"), ("event_date", 5, "開催日"),
    ("event_time", 6, "時間"), ("event_name", 7, "講演会名"), ("area", 8, "領域"), ("product", 9, "製品"),
    ("url", 10, "視聴予約"), ("meeting_id", 11, "会議id"), ("passcode", 12, "パスコード"), ("prefecture", 13, "都道府県"),
    ("area_guide", 14, "他エリア"), ("expiration", 15, "使用期限"), ("linked_doc", 16, "文書番号"),
    ("approval_date", 17, "承認日"), ("doc_name", 18, "ドキュメント名"),
)


def _find_header(ws: Any) -> tuple[dict[str, int | None], int] | None:
    """Locate the header row and columns by name, so old and new form layouts (e.g. without 承認日) both work."""
    normalize = lambda value: re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value or ""))).lower()
    for header_row in range(1, min(ws.max_row, 10) + 1):
        headers = {col: normalize(ws.cell(row=header_row, column=col).value) for col in range(1, min(ws.max_column, 40) + 1)}
        found = {key: next((col for col, text in headers.items() if normalize(keyword) in text), None) for key, _, keyword in _COLUMNS}
        if sum(col is not None for col in found.values()) >= 10:
            return found, header_row + 1
    return None


def _header_columns(ws: Any) -> tuple[dict[str, int | None], int]:
    return _find_header(ws) or ({key: default for key, default, _ in _COLUMNS}, _DATA_START_ROW)


_DATE_PREFIX = re.compile(
    r"^\s*(?:(?:20\d{2}\s*年\s*)?[\d\s]+月\s*[\d\s]+日|20\d{2}\s*[/.\-]\s*\d{1,2}\s*[/.\-]\s*\d{1,2})"
    r"\s*(?:[(（]\s*[月火水木金土日]\s*[)）])?\s*")


def _sheet_time(value: Any) -> str:
    """Time cell as staff transcribe it: leading date removed, "~" kept, slots on separate lines."""
    if not isinstance(value, str):
        return _format_time(value)
    text = _DATE_PREFIX.sub("", value.strip())
    return re.sub(r"(?<=\d)\s*[/／]\s*(?=\d)", "\n", text).strip()


def is_excel(filename: str) -> bool:
    return filename.lower().endswith((".xlsx", ".xlsm", ".xls"))


def is_regeneron_申請書(filename: str) -> bool:
    """ファイル名から Regeneron の申請書 Excel かどうか判定する"""
    return ("申請書" in filename or "フラグメント" in filename) and is_excel(filename)


def is_application_workbook(data: bytes, filename: str) -> bool:
    """申請書かどうかをファイル名か中身（専用シート名・見出し）で判定する"""
    # A named but broken/legacy .xls still counts, so the parser can report it instead of silently skipping.
    if is_regeneron_申請書(filename):
        return True
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        return False
    try:
        wb = load_workbook(BytesIO(data), data_only=True)
    except Exception:
        return False
    try:
        return _FRAGMENT_SHEET_NAME in wb.sheetnames or _find_header(wb.active) is not None
    finally:
        wb.close()


def parse_regeneron_申請書(data: bytes, filename: str) -> list[dict[str, Any]]:
    """
    Regeneron 申請書 Excel バイナリを解析してフラグメント出力リストを返す。

    Returns:
        list of output dicts compatible with the outputs table:
          { type, name, status, fields }
    """
    try:
        wb = load_workbook(BytesIO(data), data_only=True)
    except Exception as e:
        raise ValueError("申請書を読み込めません。.xlsx 形式で保存し直してください。") from e

    # シート選択: 専用シートを優先、なければ最初のシート
    if _FRAGMENT_SHEET_NAME in wb.sheetnames:
        ws = wb[_FRAGMENT_SHEET_NAME]
    else:
        ws = wb.active

    outputs: list[dict[str, Any]] = []
    fragment_count = 0

    columns, data_start = _header_columns(ws)
    for row_idx in range(data_start, ws.max_row + 1):
        def cell(key: str) -> Any:
            return ws.cell(row=row_idx, column=columns[key]).value if columns[key] else None

        # The sample row under the headers is marked "例" in the # column.
        if str(ws.cell(row=row_idx, column=1).value or "").strip() == "例":
            continue

        # v4.1 forms add a note such as "MKT（CSCC）"; the sheet and naming rule need the bare MKT/Sales.
        kubun1_original = str(cell("kubun1") or "").strip()
        kubun1_match = re.match(r"(MKT|Sales)(?![A-Za-z])", kubun1_original, re.I)
        if not kubun1_match:
            continue
        kubun1 = "MKT" if kubun1_match.group(1).upper() == "MKT" else "Sales"
        indication = str(cell("indication") or "").strip()
        if kubun1 == "MKT" and indication and kubun1_original == kubun1:
            kubun1_original = f"MKT（{indication}）"

        kubun2    = str(cell("kubun2") or "").strip()
        event_id  = str(cell("event_id") or "").strip()
        raw_event_date = cell("event_date")
        raw_event_time = cell("event_time")
        event_date = _format_date(raw_event_date)
        event_time = _format_time(raw_event_time) or _format_time(raw_event_date)
        # The sheet keeps the time cell as written ("18:30~19:50", "第1部：…"); only an empty cell falls back.
        event_time_raw = _sheet_time(raw_event_time) or event_time
        event_name = str(cell("event_name") or "").strip()
        area       = str(cell("area") or "").strip()
        product    = str(cell("product") or "").strip()

        # 視聴予約ページのURL: セル値が URL でなければハイパーリンクを試みる
        url_cell = ws.cell(row=row_idx, column=columns["url"]) if columns["url"] else None
        url_raw = str(url_cell.value or "").strip() if url_cell else ""
        url_value = url_raw
        if not _is_url(url_value) and url_cell is not None and url_cell.hyperlink:
            url_value = url_cell.hyperlink.target or url_value
        # URL として解釈できない（ハイパーリンクも無い）場合は空にし、フラグを立てる
        viewing_url_missing = bool(url_raw) and not _is_url(url_value)
        if viewing_url_missing:
            url_value = ""

        meeting_id  = str(cell("meeting_id") or "").strip()
        passcode    = str(cell("passcode") or "").strip()
        prefecture  = str(cell("prefecture") or "").strip()
        area_guide  = str(cell("area_guide") or "").strip()
        expiration  = _format_date(cell("expiration"))
        linked_doc  = str(cell("linked_doc") or "").strip()
        approval_date = _format_date(cell("approval_date"))
        doc_name    = str(cell("doc_name") or "").strip()

        # 空行スキップ (講演会名と event_id の両方が空)
        if not event_name and not event_id:
            continue

        # eventName keeps the raw value so the sheet transcript can highlight what was cleaned.
        fragment_name = application_fragment_name({
            "kubun1": kubun1, "kubun2": kubun2, "eventDate": event_date,
            "eventName": event_name, "prefecture": prefecture, "areaGuide": area_guide,
        })

        fragment_count += 1
        outputs.append({
            "type": "fragment",
            "name": f"フラグメント {fragment_count}",
            "status": "not_registered",
            "row_number": row_idx,
            "fields": {
                "fragmentName":   fragment_name,
                "kubun1":         kubun1,
                "kubun1Original": kubun1_original if kubun1_original != kubun1 else "",
                "kubun2":         kubun2,
                "eventId":        event_id,
                "eventDate":      event_date,
                "eventTime":      event_time,
                "eventTimeRaw":   event_time_raw,
                "eventName":      event_name,
                "area":           area,
                "product":        product,
                "viewingUrl":     url_value,
                "viewingUrlMissing": viewing_url_missing,
                "viewingUrlOriginal": url_raw if viewing_url_missing else "",
                "meetingId":      meeting_id,
                "passcode":       passcode,
                "prefecture":     prefecture,
                "areaGuide":      area_guide,
                "expiration":     expiration,
                "linkedDocument": linked_doc,
                "approvalDate":   approval_date,
                "documentName":   doc_name,
            },
        })

    wb.close()
    return outputs
