"""Small, independent Regeneron inbox. No event-manager runtime dependency."""
from __future__ import annotations

import base64
import hashlib
import html
import io
import ipaddress
import json
import logging
import os
import re
import threading
import uuid
import zipfile
from datetime import datetime, timedelta, timezone
from email.utils import parseaddr
from pathlib import Path
from urllib.parse import quote, urlsplit

import fitz
import requests
from PIL import Image, ImageOps
from fastapi import APIRouter, Depends, File, HTTPException, Request, UploadFile
from fastapi.responses import Response
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.credentials import Credentials
from pydantic import BaseModel, Field
from typing import Literal, Optional

from regeneron_application import application_fragment_name, is_application_workbook, is_excel, is_regeneron_申請書, parse_regeneron_申請書

log = logging.getLogger(__name__)
MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_PAGES = 30
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


JST = timezone(timedelta(hours=9))


def now():
    return datetime.now(timezone.utc).isoformat()


def banner_name(index):
    return "banner.png" if index == 0 else f"banner{index + 1}.png"


def linked_document(value):
    match = re.search(r"[A-Za-z]{2,}(?:[.-][A-Za-z0-9]+){2,}", value or "")
    document = match.group() if match else (value or "").strip()
    # Requesters sometimes drop the country code ("P.LIB.26.09.0006" / "LIB.26.09.0006"); numbers are JP.<product>.<yy>.<mm>.<seq>.
    return "JP." + document if re.fullmatch(r"[A-Za-z]{3}\.\d{2}\.\d{2}\.\d+", document) else document


def valid_viewing_url(value):
    try:
        parsed = urlsplit(value)
        return parsed.scheme in ("http", "https") and bool(parsed.netloc)
    except ValueError:
        return False


def message_text(part):
    if part.get("mimeType") == "text/plain" and part.get("body", {}).get("data"):
        encoded = part["body"]["data"]
        return base64.urlsafe_b64decode(encoded + "=" * (-len(encoded) % 4)).decode("utf-8", "ignore")
    return "\n".join(message_text(child) for child in part.get("parts", []))


def requester_name(headers, message):
    # Request mail has an "氏名" field; otherwise the signature shows "Kenta Asada　浅田　健太" or "相澤　信一郎　Shinichiro Aizawa".
    text = message_text(message.get("payload", {}))
    kanji = r"([一-鿿々]+)"
    match = re.search(r"氏名[ \t　]*(?:\r?\n[ \t　]*)+" + kanji, text)
    raw = parseaddr(headers.get("reply-to") or headers.get("from", ""))[0]
    name = re.sub(r"^'(.*)' via .*$", r"\1", raw).strip()
    if not match and name:
        match = (re.search(re.escape(name) + r"[ \t　]+" + kanji, text)
                 or re.search(kanji + r"[ \t　]+[一-鿿々]+[ \t　]+" + re.escape(name), text))
    if not match and not name:
        return ""
    return (match.group(1) if match else name) + "様"


def transcript_message(entry):
    sheet, rows = entry["sheet"], entry["rows"]
    first, last = sheet["row"], sheet["row"] + len(rows) - 1
    multiple = len(rows) > 1
    row_label = lambda index: f"{first + index}行目 " if multiple else ""
    corrections = [f"{first + index}行目" for index, row in enumerate(rows) if row.get("category") == "修正"]
    all_corrections = len(corrections) == len(rows)
    lines = ["【リジェネロン】フラグメントの修正申請が届きました" if all_corrections else "【リジェネロン】フラグメント申請が届きました",
             f'件名: {entry["subject"].strip()}']
    if entry.get("requester"):
        lines.append(f'申請者: {entry["requester"]}')
    if sheet.get("release_date"):
        lines.append(f'本番化予定日: {sheet["release_date"]}')
    lines.append(f'管理表: {first}行目{f"〜{last}行目" if multiple else ""}に転記済み')
    if corrections and not all_corrections:
        lines.append(f'区分: {"、".join(corrections)}は修正')
    changes = [f'{row_label(index)}{change["label"]}: 「{change["before"].replace(chr(13), "").replace(chr(10), "⏎")}」→「{change["after"]}」'
               for index, row in enumerate(rows) for change in row.get("changes", [])]
    if changes:
        lines += ["", "【申請書から整えて転記した箇所】", *changes]
    warnings = [f'{row_label(index)}{warning["label"]}: {warning["message"]}' for index, row in enumerate(rows) for warning in row.get("warnings", [])]
    if warnings:
        lines += ["", "【シートで確認・入力が必要な箇所】", *warnings]
    lines += ["", "内容を確認の上、作業者に依頼してください", "", f'<{sheet["url"]}|スプレッドシートを開く>']
    if entry.get("rfc822_id"):
        lines.append(mail_link(entry["rfc822_id"]))
    return "\n".join(lines)


REQUEST_SUBJECT = re.compile(r"新規依頼|修正依頼|フラグメント申請")
# Replies, forwards and Outlook recall notices ("取り消し: …") are not new requests.
REPLY_SUBJECT = re.compile(r"^\s*(?:re|fw|fwd|recall|取り消し|返信|転送)\s*[:：]", re.I)


def missing_application_message(entry):
    lines = ["【リジェネロン】申請書が添付されていない申請メールがあります（管理表には転記していません）", f'件名: {entry["subject"].strip()}']
    if entry.get("requester"):
        lines.append(f'申請者: {entry["requester"]}')
    lines += ["", "申請者に申請書の添付を依頼してください"]
    if entry.get("rfc822_id"):
        lines += ["", mail_link(entry["rfc822_id"])]
    return "\n".join(lines)


def mail_link(rfc822_id):
    # Search by the Message-ID header: unlike Gmail thread ids it is the same in every list member's mailbox.
    return f'<https://mail.google.com/mail/u/0/#search/{quote("rfc822msgid:" + rfc822_id.strip(), safe="")}|該当メールを開く>'


def import_error_message(message, error):
    headers = {h["name"].lower(): h["value"] for h in message.get("payload", {}).get("headers", [])}
    lines = ["【リジェネロン】申請メールを取り込めませんでした（管理表には転記していません）", f'件名: {headers.get("subject", "（件名なし）").strip()}',
             f"理由: {error}", "", "申請書の内容を確認し、必要に応じて申請者に再送を依頼してください"]
    if headers.get("message-id"):
        lines += ["", mail_link(headers["message-id"])]
    return "\n".join(lines)


def notify_chat(text):
    url = os.getenv("REGENERON_CHAT_WEBHOOK_URL", "")
    if url:
        requests.post(url, json={"text": text}, timeout=20).raise_for_status()


MAP_WORDS = re.compile(r"地図|案内図|アクセスマップ|交通アクセス|会場マップ|\bmap\b", re.I)


def map_pages(data, filename):
    """Pages whose text names a map; a map drawn only as an image without any label is not detectable."""
    if not filename.lower().endswith(".pdf"):
        return set()
    with fitz.open(stream=data, filetype="pdf") as pdf:
        # Headings are often letter-spaced ("地 図"), so test both the raw and the space-free text.
        return {index for index, page in enumerate(pdf) if MAP_WORDS.search(page.get_text()) or MAP_WORDS.search(re.sub(r"\s+", "", page.get_text()))}


def selected_banners(card):
    banners = card.get("banners", [])
    if "selected_pages" in card:
        pages = sorted({index for index in card["selected_pages"] if 0 <= index < len(banners)})
        return [banners[index] for index in pages] or banners[:1]
    # Cards saved before page selection used the all_pages switch.
    return banners if card.get("all_pages", False) else banners[:1]


def render_images(data, filename):
    if not data or len(data) > MAX_FILE_BYTES:
        raise ValueError("ファイルは25MB以下で指定してください。")
    images = []
    if filename.lower().endswith(".pdf"):
        with fitz.open(stream=data, filetype="pdf") as pdf:
            if pdf.needs_pass or not 1 <= len(pdf) <= MAX_PAGES:
                raise ValueError("パスワードなし・30ページ以下のPDFを指定してください。")
            for page in pdf:
                if page.rect.width <= 0 or page.rect.height / page.rect.width > 10:
                    raise ValueError("PDFのページサイズを確認してください。")
                scale = 1200 / page.rect.width
                images.append(page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False).tobytes("png"))
    else:
        with Image.open(io.BytesIO(data)) as original:
            image = ImageOps.exif_transpose(original)
            if image.height / image.width > 10:
                raise ValueError("画像の縦横比を確認してください。")
            image = image.convert("RGBA")
            image = image.resize((1200, max(1, round(image.height * 1200 / image.width))), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            image.save(buf, format="PNG")
            images.append(buf.getvalue())
    return images


def fragment_html(card, image_count, image_source=None, button_source="images/btn.png"):
    url = card.get("viewing_url", "").strip()
    if url and not valid_viewing_url(url):
        raise ValueError("参加登録URLは http / https のURLを指定してください。")
    safe_url = html.escape(url, quote=True)
    rows = []
    has_button = bool(url and card.get("include_button", True))
    if has_button:
        rows.extend([
            "                <tr>",
            '                  <td class="cell" width="550" align="center"',
            '                    style="padding:0 5px 30px; border-collapse:collapse; text-align:center; font-size:0;"><a',
            f'                      href="{safe_url}"><img src="{html.escape(button_source, quote=True)}" alt="参加登録はこちら" width="100%" border="0"',
            '                        style="margin:0; padding:0; border:none; outline:none; line-height:100%; text-decoration:none; display:block;" /></a>',
            "                  </td>",
            "                </tr>",
        ])
    for index in range(image_count):
        source = image_source(index) if image_source else f"images/{banner_name(index)}"
        padding = "0" if index == 0 else "10px 0 0"
        escaped_source = html.escape(source, quote=True)
        if url:
            image_start = f'                    style="padding:{padding}; border-collapse:collapse; text-align:center; font-size:0;"><a'
            image_link = f'                      href="{safe_url}"><img'
        else:
            image_start = f'                    style="padding:{padding}; border-collapse:collapse; text-align:center; font-size:0;"><img'
            image_link = None
        rows.extend([
            "                <tr>",
            '                  <td class="cell" width="560" align="center"',
            image_start,
        ])
        if image_link:
            rows.append(image_link + (f' src="{escaped_source}" alt="バナー" width="100%" border="0"' if has_button else ""))
        if not has_button:
            rows.append(f'                        src="{escaped_source}" alt="バナー" width="100%" border="0"')
        rows.extend([
            '                        style="margin:0; padding:0; border:none; outline:none; line-height:100%; text-decoration:none; display:block;" />' + ("</a>" if url else ""),
            "                  </td>",
            "                </tr>",
        ])
    return "\n".join([
        "<tr>",
        '  <td style="word-break: break-word;border-collapse: collapse !important;vertical-align: top;padding-bottom: 10px;">',
        '    <table cellpadding="0" cellspacing="0" align="center" width="100%" border="0"',
        '      style="border-spacing: 0;border-collapse: collapse;vertical-align: top;max-width: 560px;">',
        "      <tbody>",
        "        <tr>",
        '          <td style="word-break: break-word;border-collapse: collapse !important;vertical-align: top;padding: 0;">',
        '            <table cellpadding="0" cellspacing="0" align="center" width="100%" border="0"',
        '              style="border-spacing: 0;border-collapse: collapse;vertical-align: top;">',
        "              <tbody>",
        *rows,
        "              </tbody>",
        "            </table>",
        "          </td>",
        "        </tr>",
        "      </tbody>",
        "    </table>",
        "  </td>",
        "</tr>",
    ])


def build_zip(card, images):
    name = card["fragment_name"].strip()
    if not name or not images:
        raise ValueError("フラグメント名とバナー画像を指定してください。")
    if card.get("viewing_url_missing"):
        raise ValueError("申請書の参加登録URLを読み取れません。URLを修正してください。")
    markup = fragment_html(card, len(images))
    inner = io.BytesIO()
    with zipfile.ZipFile(inner, "w", zipfile.ZIP_DEFLATED) as archive:
        if card.get("viewing_url", "").strip() and card.get("include_button", True):
            archive.writestr("images/btn.png", (Path(__file__).parent / "regeneron_assets/btn.png").read_bytes())
        for index, data in enumerate(images):
            archive.writestr(f"images/{banner_name(index)}", data)
    safe_name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name).strip(". ")[:160] or "fragment"
    outer = io.BytesIO()
    with zipfile.ZipFile(outer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"{safe_name}.html", markup)
        archive.writestr("images.zip", inner.getvalue())
    return safe_name + ".zip", outer.getvalue()


class Store:
    def __init__(self, connect):
        self.connect = connect

    def init(self):
        with self.connect() as con:
            con.execute("CREATE TABLE IF NOT EXISTS regeneron_cards (id TEXT PRIMARY KEY, thread_id TEXT NOT NULL, payload JSONB NOT NULL, received_at TEXT NOT NULL)")
            con.execute("CREATE INDEX IF NOT EXISTS regeneron_cards_thread ON regeneron_cards(thread_id)")
            con.execute("CREATE INDEX IF NOT EXISTS regeneron_cards_received ON regeneron_cards(received_at DESC)")
            con.execute("CREATE TABLE IF NOT EXISTS regeneron_messages (id TEXT PRIMARY KEY, imported_at TEXT NOT NULL)")
            con.execute("CREATE TABLE IF NOT EXISTS regeneron_settings (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
            con.execute("CREATE TABLE IF NOT EXISTS regeneron_exports (id TEXT PRIMARY KEY, payload JSONB NOT NULL, created_at TEXT NOT NULL, done_at TEXT, error TEXT)")
            con.execute("INSERT INTO regeneron_settings(key,value) VALUES ('export_lock','1') ON CONFLICT DO NOTHING")
            claimed = con.execute("INSERT INTO regeneron_settings(key,value) VALUES ('single_page_default_v2','1') ON CONFLICT DO NOTHING RETURNING key").fetchone()
            if claimed:
                con.execute("UPDATE regeneron_cards SET payload=jsonb_set(payload, '{all_pages}', 'false'::jsonb) WHERE NOT (COALESCE(payload->'edited_fields', '[]'::jsonb) ? 'all_pages')")
            claimed = con.execute("INSERT INTO regeneron_settings(key,value) VALUES ('clean_fragment_name_v1','1') ON CONFLICT DO NOTHING RETURNING key").fetchone()
            if claimed:
                # Re-derive names imported before newline/date cleanup; manual edits stay as they are.
                for row in con.execute("SELECT id,payload FROM regeneron_cards").fetchall():
                    card, fields = row["payload"], row["payload"].get("application") or {}
                    if not fields.get("kubun1") or "fragment_name" in card.get("edited_fields", []):
                        continue
                    name = application_fragment_name(fields)
                    if name != card.get("fragment_name"):
                        card["fragment_name"] = fields["fragmentName"] = name
                        con.execute("UPDATE regeneron_cards SET payload=%s::jsonb WHERE id=%s", (json.dumps(card), row["id"]))
            con.commit()

    def start_at(self):
        with self.connect() as con:
            configured = os.getenv("REGENERON_GMAIL_START_AT") or now()
            value = datetime.fromisoformat(configured.replace("Z", "+00:00"))
            if value.tzinfo is None:
                raise ValueError("REGENERON_GMAIL_START_AT requires a timezone")
            con.execute("INSERT INTO regeneron_settings (key,value) VALUES ('start_at',%s) ON CONFLICT DO NOTHING", (value.isoformat(),))
            row = con.execute("SELECT value FROM regeneron_settings WHERE key='start_at'").fetchone()
            con.commit()
            return row["value"]

    def cards(self, thread_id=None):
        with self.connect() as con:
            if thread_id:
                rows = con.execute("SELECT payload FROM regeneron_cards WHERE thread_id=%s ORDER BY id", (thread_id,)).fetchall()
            else:
                rows = con.execute("SELECT payload FROM regeneron_cards ORDER BY received_at DESC,id").fetchall()
            return [row["payload"] for row in rows]

    def card(self, card_id):
        with self.connect() as con:
            row = con.execute("SELECT payload FROM regeneron_cards WHERE id=%s", (card_id,)).fetchone()
        if not row:
            raise HTTPException(404, "カードが見つかりません。")
        return row["payload"]

    def known(self, ids):
        with self.connect() as con:
            rows = con.execute("SELECT id FROM regeneron_messages WHERE id=ANY(%s)", (ids,)).fetchall()
            return {row["id"] for row in rows}

    def save_message(self, message_id, cards, export=None):
        with self.connect() as con:
            # Atomic claim also prevents duplicate imports with multiple API workers.
            claimed = con.execute("INSERT INTO regeneron_messages VALUES (%s,%s) ON CONFLICT DO NOTHING RETURNING id", (message_id, now())).fetchone()
            if not claimed:
                return 0
            if export:
                con.execute("INSERT INTO regeneron_exports (id,payload,created_at) VALUES (%s,%s::jsonb,%s) ON CONFLICT DO NOTHING", (message_id, json.dumps(export), now()))
            for card in cards:
                old = con.execute("SELECT payload FROM regeneron_cards WHERE id=%s FOR UPDATE", (card["id"],)).fetchone()
                if old:
                    previous = old["payload"]
                    if previous["received_at"] > card["received_at"]:
                        continue
                    for field in previous.get("edited_fields", []):
                        card[field] = previous[field]
                    card["edited_fields"] = previous.get("edited_fields", [])
                    card["sources"] = list({s["key"]: s for s in previous.get("sources", []) + card.get("sources", [])}.values())
                con.execute("INSERT INTO regeneron_cards VALUES (%s,%s,%s::jsonb,%s) ON CONFLICT(id) DO UPDATE SET payload=EXCLUDED.payload, received_at=EXCLUDED.received_at", (card["id"], card["thread_id"], json.dumps(card), card["received_at"]))
            con.commit()
        return len(cards)

    def queue_notice(self, notice_id, text):
        with self.connect() as con:
            # The id makes each alert one-shot (per mail, or per day for outages) across retries and API workers.
            con.execute("INSERT INTO regeneron_exports (id,payload,created_at) VALUES (%s,%s::jsonb,%s) ON CONFLICT DO NOTHING",
                        (notice_id, json.dumps({"kind": "notice", "message_id": notice_id, "text": text}), now()))
            con.commit()

    def dismiss(self, message_id):
        with self.connect() as con:
            # Claiming the mail stops import retries; finishing its export stops transcript/Chat retries.
            con.execute("INSERT INTO regeneron_messages VALUES (%s,%s) ON CONFLICT DO NOTHING", (message_id, now()))
            con.execute("UPDATE regeneron_exports SET done_at=%s, error=%s WHERE id=%s AND done_at IS NULL", (now(), "dismissed", message_id))
            con.commit()

    def process_exports(self, handle):
        with self.connect() as con:
            # Serialize sheet appends across API workers so two batches never pick the same empty row.
            con.execute("SELECT key FROM regeneron_settings WHERE key='export_lock' FOR UPDATE")
            rows = con.execute("SELECT id,payload FROM regeneron_exports WHERE done_at IS NULL ORDER BY created_at,id").fetchall()
            for row in rows:
                payload, error = handle(row["payload"])
                con.execute("UPDATE regeneron_exports SET payload=%s::jsonb, done_at=%s, error=%s WHERE id=%s", (json.dumps(payload), None if error else now(), error, row["id"]))
            con.commit()

    def update(self, card_id, changes):
        with self.connect() as con:
            row = con.execute("SELECT payload FROM regeneron_cards WHERE id=%s FOR UPDATE", (card_id,)).fetchone()
            if not row:
                raise HTTPException(404, "カードが見つかりません。")
            card = row["payload"]
            card.update(changes)
            card["edited_fields"] = sorted((set(card.get("edited_fields", [])) | set(changes)) - {"sources"})
            card["updated_at"] = now()
            con.execute("UPDATE regeneron_cards SET payload=%s::jsonb WHERE id=%s", (json.dumps(card), card_id))
            con.commit()
        return card


class Gmail:
    def __init__(self):
        raw = os.getenv("GMAIL_TOKEN_JSON", "")
        if raw:
            creds = Credentials.from_authorized_user_info(json.loads(raw), SCOPES)
        else:
            creds = Credentials.from_authorized_user_file(os.environ["GMAIL_TOKEN_FILE"], SCOPES)
        self.session = AuthorizedSession(creds)

    def get(self, path, **params):
        res = self.session.get("https://gmail.googleapis.com/gmail/v1/users/me/" + path, params=params, timeout=40)
        res.raise_for_status()
        return res.json()

    def messages(self, start):
        # Epoch seconds avoid Gmail's PST interpretation of date-only queries.
        epoch = int(datetime.fromisoformat(start).timestamp()) - 1
        query = f'({os.getenv("REGENERON_GMAIL_QUERY", "from:regeneron@ar40.jp")}) after:{epoch}'
        token = None
        result = []
        while True:
            params = {"q": query, "maxResults": 500}
            if token:
                params["pageToken"] = token
            page = self.get("messages", **params)
            result.extend(page.get("messages", []))
            token = page.get("nextPageToken")
            if not token:
                return result

    def part_data(self, message_id, part):
        body = part.get("body", {})
        if body.get("size", 0) > MAX_FILE_BYTES:
            raise ValueError("添付ファイルは25MB以下で指定してください。")
        if body.get("attachmentId"):
            body = self.get(f'messages/{quote(message_id, safe="")}/attachments/{quote(body["attachmentId"], safe="")}')
        encoded = body.get("data", "")
        data = base64.urlsafe_b64decode(encoded + "=" * (-len(encoded) % 4))
        if len(data) > MAX_FILE_BYTES:
            raise ValueError("添付ファイルは25MB以下で指定してください。")
        return data

    def thread_attachments(self, thread_id):
        thread = self.get(f'threads/{quote(thread_id, safe="")}', format="full")
        result = []
        for message in thread.get("messages", []):
            def visit(part, path="0"):
                if part.get("filename"):
                    result.append({"id": hashlib.sha256(f'{message["id"]}:{path}'.encode()).hexdigest()[:32],
                        "message_id": message["id"], "filename": part["filename"],
                        "size": part.get("body", {}).get("size", 0), "mime_type": part.get("mimeType", ""),
                        "received_at": datetime.fromtimestamp(int(message["internalDate"]) / 1000, timezone.utc).isoformat(), "part": part})
                for index, child in enumerate(part.get("parts", [])):
                    visit(child, f"{path}.{index}")
            visit(message.get("payload", {}))
        return sorted(result, key=lambda item: item["received_at"], reverse=True)

    def attachments(self, message):
        result = []
        def visit(part):
            filename = part.get("filename", "")
            if filename and (is_excel(filename) or filename.lower().endswith(".pdf")):
                body = part.get("body", {})
                # Unnamed Excel files are only candidates (checked by content later); a huge one must not block the mail.
                if body.get("size", 0) > MAX_FILE_BYTES and is_excel(filename) and not is_regeneron_申請書(filename):
                    for child in part.get("parts", []):
                        visit(child)
                    return
                if body.get("size", 0) > MAX_FILE_BYTES:
                    raise ValueError("メール添付は25MB以下で指定してください。")
                if body.get("attachmentId"):
                    body = self.get(f'messages/{message["id"]}/attachments/{quote(body["attachmentId"], safe="")}')
                encoded = body.get("data", "")
                data = base64.urlsafe_b64decode(encoded + "=" * (-len(encoded) % 4))
                if not data or len(data) > MAX_FILE_BYTES:
                    raise ValueError("添付ファイルが空か、25MBを超えています。")
                result.append((filename, data))
            for child in part.get("parts", []):
                visit(child)
        visit(message.get("payload", {}))
        return result


class FragmentService:
    def __init__(self, store, upload, download, sheet_reader=None, sheet_writer=None, notify=notify_chat):
        self.sheet_reader, self.sheet_writer, self.notify = sheet_reader, sheet_writer, notify
        self.store, self.upload, self.download = store, upload, download
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        self.thread = None
        self.status_lock = threading.Lock()
        self.last = {"last_sync_at": None, "error": None, "errors": [], "imported": 0}
        self.sync_failures = 0

    def configured(self):
        return bool(os.getenv("GMAIL_TOKEN_JSON") or os.getenv("GMAIL_TOKEN_FILE"))

    def enabled(self):
        return os.getenv("REGENERON_GMAIL_ENABLED", "1").lower() not in ("0", "false", "no")

    def status(self):
        with self.status_lock:
            return {**self.last, "configured": self.configured(), "enabled": self.enabled(), "running": self.lock.locked(), "interval_seconds": max(30, int(os.getenv("REGENERON_GMAIL_INTERVAL_SECONDS", "60")))}

    def asset(self, data, filename, content_type):
        key = f"regeneron/{uuid.uuid4().hex}/{filename}"
        self.upload(data, key, content_type)
        return key

    def banners(self, data, filename):
        maps = map_pages(data, filename)
        return [{"key": self.asset(image, banner_name(i), "image/png"), "filename": banner_name(i), "has_map": i in maps}
                for i, image in enumerate(render_images(data, filename))]

    def import_message(self, message, attachments):
        headers = {h["name"].lower(): h["value"] for h in message.get("payload", {}).get("headers", [])}
        subject = headers.get("subject", "（件名なし）")
        received = datetime.fromtimestamp(int(message["internalDate"]) / 1000, timezone.utc).isoformat()
        existing = {c["id"]: c for c in self.store.cards(message["threadId"])}
        outputs, sources = [], []
        for filename, data in attachments:
            if is_application_workbook(data, filename):
                parsed = parse_regeneron_申請書(data, filename)
                if not parsed:
                    raise ValueError("申請書に有効なフラグメント行がありません。")
                outputs.extend(o["fields"] for o in parsed)
            elif filename.lower().endswith(".pdf"):
                sources.append({"key": self.asset(data, "source.pdf", "application/pdf"), "filename": filename, "data": data})
        missing = None
        # A new request (not a reply/forward) whose thread has no application yet: the form was probably forgotten.
        if (not outputs and self.sheet_writer and REQUEST_SUBJECT.search(subject) and not REPLY_SUBJECT.match(subject)
                and not any(card.get("application") for card in existing.values())):
            missing = {"kind": "missing_application", "message_id": message["id"], "thread_id": message["threadId"], "rfc822_id": headers.get("message-id", ""),
                       "subject": subject, "received_at": received, "requester": requester_name(headers, message)}
        if not outputs and not sources:
            return self.store.save_message(message["id"], [], missing)
        has_application = bool(outputs)
        if not outputs:
            outputs = [c.get("application", {}) for c in existing.values()] or [{}]
        cards, transcript, occurrences = [], [], {}
        for fields in outputs:
            identity = (fields.get("eventId", "") + "|" + fields.get("documentName", "")) if fields.get("eventId") else fields.get("fragmentName", "mail")
            occurrences[identity] = occurrences.get(identity, 0) + 1
            import_key = f'{identity}|{occurrences[identity]}'
            card_id = hashlib.sha256(f'{message["threadId"]}|{import_key}'.encode()).hexdigest()[:32]
            matched = next((c for c in existing.values() if c.get("import_key") == import_key), None)
            # If the PDF arrived first, enrich its placeholder instead of creating a duplicate.
            if not matched and fields and not cards:
                placeholders = [c for c in existing.values() if not c.get("application")]
                matched = placeholders[0] if len(placeholders) == 1 else None
            if matched:
                card_id = matched["id"]
            old = existing.get(card_id, {})
            if old and old["received_at"] > received:
                continue
            candidates = sources or old.get("sources", [])
            selected = candidates[0] if len(candidates) == 1 else None
            if len(candidates) > 1:
                hints = [fields.get("documentName"), linked_document(fields.get("linkedDocument")), fields.get("eventId")]
                matches = [s for s in candidates if any(h and h in s["filename"] for h in hints)]
                selected = matches[0] if len(matches) == 1 else None
            banners = old.get("banners", [])
            source_key = old.get("source_key", "")
            selected_pages = old.get("selected_pages", [0])
            if sources:
                banners = self.banners(selected["data"], selected["filename"]) if selected else []
                source_key = selected["key"] if selected else ""
                selected_pages = [0]
            card = {**old, "id": card_id, "thread_id": message["threadId"], "message_id": message["id"], "subject": subject,
                    "received_at": received, "updated_at": now(), "application": fields, "import_key": import_key,
                    "fragment_name": fields.get("fragmentName") or old.get("fragment_name") or subject,
                    "linked_document": linked_document(fields.get("linkedDocument")) or old.get("linked_document", ""),
                    "viewing_url": fields.get("viewingUrl", old.get("viewing_url", "")),
                    "viewing_url_missing": fields.get("viewingUrlMissing", old.get("viewing_url_missing", False)),
                    "registration_status": old.get("registration_status", "pending"),
                    "all_pages": old.get("all_pages", False), "include_button": old.get("include_button", True), "banners": banners, "source_key": source_key, "selected_pages": selected_pages,
                    "sources": [{k: v for k, v in source.items() if k != "data"} for source in candidates]}
            cards.append(card)
            transcript.append({**fields, "card_id": card_id, "fragment_name": card["fragment_name"], "linkedDocument": card["linked_document"],
                               "linkedDocumentRaw": fields.get("linkedDocument", ""),
                               "category": "修正" if old.get("application") or "修正" in subject else "新規"})
        export = missing
        if has_application and transcript and self.sheet_writer:
            export = {"message_id": message["id"], "thread_id": message["threadId"], "rfc822_id": headers.get("message-id", ""), "subject": subject, "received_at": received,
                      "requester": requester_name(headers, message), "rows": transcript}
        return self.store.save_message(message["id"], cards, export)

    def export_entry(self, entry, summary):
        if entry.get("kind") in ("missing_application", "notice"):
            try:
                self.notify(entry["text"] if entry["kind"] == "notice" else missing_application_message(entry))
                return entry, None
            except Exception as exc:
                log.warning("Regeneron missing-application notice failed: %s (%s)", entry["message_id"], type(exc).__name__)
                error = "Google Chatに通知できませんでした。次回再試行します。"
                summary["errors"].append({"message_id": entry["message_id"], "message": error})
                return entry, error
        try:
            # The sheet result is persisted before notifying, so a Chat failure never appends the rows twice.
            if not entry.get("sheet"):
                entry["sheet"] = self.sheet_writer(entry)
            self.notify(transcript_message(entry))
            return entry, None
        except Exception as exc:
            log.warning("Regeneron transcript failed: %s (%s)", entry["message_id"], type(exc).__name__)
            error = str(exc) if isinstance(exc, ValueError) else ("Google Chatへの通知" if entry.get("sheet") else "スプレッドシートへの転記") + "に失敗しました。次回再試行します。"
            summary["errors"].append({"message_id": entry["message_id"], "message": error})
            if not entry.get("sheet"):
                entry["sheet_failures"] = entry.get("sheet_failures", 0) + 1
                # Alert once a failure persists (~3 syncs); a single network blip retries silently.
                if entry["sheet_failures"] == 3:
                    try:
                        self.store.queue_notice(f'transcript-error:{entry["message_id"]}', "\n".join([
                            "【リジェネロン】管理表に転記できていない申請があります（自動で再試行中）",
                            f'件名: {entry.get("subject", "")}', f"理由: {error}（{type(exc).__name__}）",
                            "シートの共有設定・関数入りの空き行を確認してください。"]))
                    except Exception:
                        log.warning("Regeneron transcript alert could not be queued: %s", entry["message_id"])
            return entry, error

    def sync(self):
        if not self.enabled() or not self.configured():
            return self.status()
        if not self.lock.acquire(blocking=False):
            return self.status()
        summary = {"last_sync_at": now(), "error": None, "errors": [], "imported": 0}
        gmail = None
        try:
            start = self.store.start_at()
            gmail = Gmail()
            messages = gmail.messages(start)
            known = self.store.known([m["id"] for m in messages]) if messages else set()
            pending = []
            for item in messages:
                if self.stop_event.is_set():
                    break
                if item["id"] in known:
                    continue
                try:
                    message = gmail.get(f'messages/{item["id"]}', format="full")
                    if int(message["internalDate"]) / 1000 >= datetime.fromisoformat(start).timestamp():
                        pending.append(message)
                except Exception:
                    summary["errors"].append({"message_id": item["id"], "message": "メールを取得できませんでした。次回再試行します。"})
            # Import in received order so a correction never gets replaced by an older request.
            for message in sorted(pending, key=lambda m: int(m["internalDate"])):
                if self.stop_event.is_set():
                    break
                try:
                    summary["imported"] += self.import_message(message, gmail.attachments(message))
                except Exception as exc:
                    log.warning("Regeneron import failed: %s (%s)", message["id"], type(exc).__name__)
                    summary["errors"].append({"message_id": message["id"], "message": str(exc) if isinstance(exc, ValueError) else "添付ファイルの取込に失敗しました。次回再試行します。"})
                    # ValueError means the mail itself is unusable (form layout, .xls, broken file); it will not fix itself.
                    if isinstance(exc, ValueError) and self.sheet_writer:
                        self.store.queue_notice(f'import-error:{message["id"]}', import_error_message(message, str(exc)))
            if self.sheet_writer and not self.stop_event.is_set():
                self.store.process_exports(lambda entry: self.export_entry(entry, summary))
            self.sync_failures = 0
        except Exception as exc:
            log.warning("Regeneron sync failed (%s)", type(exc).__name__)
            summary["error"] = "メール連携に失敗しました。Gmail認証・接続設定を確認してください。"
            self.sync_failures += 1
            if self.sync_failures >= 3 and self.sheet_writer:
                self.alert_outage(exc, summary)
        finally:
            if gmail:
                gmail.session.close()
            with self.status_lock:
                self.last = summary
            self.lock.release()
        return self.status()

    def alert_outage(self, exc, summary):
        # One alert per JST day while the inbox sync keeps failing (e.g. an expired Gmail token).
        day = datetime.now(timezone.utc).astimezone(JST).strftime("%Y-%m-%d")
        text = "\n".join(["【リジェネロン】申請メールの取り込みが止まっています", f"理由: {type(exc).__name__}",
                          "新しい申請はシートへの転記・通知がされません。Gmail認証（GMAIL_TOKEN）・DB・接続設定を確認してください。"])
        try:
            self.store.queue_notice(f"sync-error:{day}", text)
            self.store.process_exports(lambda entry: self.export_entry(entry, summary))
        except Exception:
            # The store itself may be what is down; post directly, only on the third consecutive failure of this process.
            try:
                if self.sync_failures == 3:
                    self.notify(text)
            except Exception:
                log.warning("Regeneron outage alert failed")

    def dismiss(self, message_id):
        self.store.dismiss(message_id)
        with self.status_lock:
            self.last = {**self.last, "errors": [e for e in self.last["errors"] if e["message_id"] != message_id]}
        return self.status()

    def start(self):
        if not self.enabled() or not self.configured() or self.thread:
            return
        self.store.start_at()
        self.stop_event.clear()
        def run():
            while not self.stop_event.is_set():
                self.sync()
                self.stop_event.wait(self.status()["interval_seconds"])
        self.thread = threading.Thread(target=run, name="regeneron-inbox", daemon=True)
        self.thread.start()

    def stop(self):
        self.stop_event.set()


class CardUpdate(BaseModel):
    fragment_name: Optional[str] = Field(default=None, min_length=1, max_length=250)
    linked_document: Optional[str] = Field(default=None, max_length=500)
    viewing_url: Optional[str] = Field(default=None, max_length=4000)
    all_pages: Optional[bool] = None
    selected_pages: Optional[list[int]] = Field(default=None, min_length=1, max_length=MAX_PAGES)
    include_button: Optional[bool] = None
    registration_status: Optional[Literal["pending", "registered", "hidden"]] = None


def validated_changes(payload):
    values = payload.model_dump(exclude_unset=True)
    if any(value is None for value in values.values()):
        raise HTTPException(422, "空の値には空文字を指定してください。")
    values = {key: value.strip() if isinstance(value, str) else value for key, value in values.items()}
    if "selected_pages" in values and any(index < 0 for index in values["selected_pages"]):
        raise HTTPException(422, "ページの指定が正しくありません。")
    if "fragment_name" in values and not values["fragment_name"]:
        raise HTTPException(422, "フラグメント名を入力してください。")
    if "viewing_url" in values:
        if values["viewing_url"] and not valid_viewing_url(values["viewing_url"]):
            raise HTTPException(422, "参加登録URLは http / https のURLを指定してください。")
        values["viewing_url_missing"] = False
    return values


class SourceSelection(BaseModel):
    source_key: str


class BatchDownload(BaseModel):
    card_ids: list[str] = Field(min_length=1, max_length=100)


class BatchStatusUpdate(BatchDownload):
    registration_status: Literal["pending", "registered", "hidden"]


def client_ip(request):
    """Client IP as configured by REGENERON_CLIENT_IP_HEADER ("header" or "header:index", e.g. "x-forwarded-for:-1").

    Without it the socket peer is used, which behind a proxy is the proxy itself, so an allowlist then rejects everyone.
    """
    setting = os.getenv("REGENERON_CLIENT_IP_HEADER", "").strip()
    if not setting:
        return request.client.host if request.client else ""
    name, _, index = setting.partition(":")
    values = [value.strip() for value in request.headers.get(name.strip(), "").split(",") if value.strip()]
    try:
        return values[int(index or 0)]
    except (IndexError, ValueError):
        return ""


def require_allowed_ip(request: Request):
    allowed = [item.strip() for item in os.getenv("REGENERON_ALLOWED_IPS", "").split(",") if item.strip()]
    if not allowed or request.url.path.endswith("/regeneron/client-ip"):
        return
    try:
        address = ipaddress.ip_address(client_ip(request))
        if any(address in ipaddress.ip_network(item, strict=False) for item in allowed):
            return
    except ValueError:
        pass
    raise HTTPException(403, "このネットワークからはアクセスできません。")


def create_router(service):
    # Only HTTP access is restricted; the background inbox sync, transcripts and Chat alerts run inside the API.
    router = APIRouter(prefix="/regeneron", tags=["Regeneron"], dependencies=[Depends(require_allowed_ip)])

    @router.get("/client-ip")
    def get_client_ip(request: Request):
        # Lets you see which header carries your real IP before turning the allowlist on.
        return {"detected": client_ip(request), "socket_peer": request.client.host if request.client else "",
                "headers": {name: request.headers.get(name, "") for name in ("x-forwarded-for", "x-real-ip", "true-client-ip", "cf-connecting-ip")}}

    @router.get("/cards")
    def list_cards():
        return {"cards": service.store.cards(), "sync": service.status()}

    @router.post("/sync")
    def sync():
        return service.sync()

    @router.post("/sync/errors/{message_id}/dismiss")
    def dismiss_error(message_id: str):
        return service.dismiss(message_id)

    @router.patch("/cards/{card_id}")
    def update(card_id: str, payload: CardUpdate):
        return service.store.update(card_id, validated_changes(payload))

    @router.get("/cards/{card_id}")
    def get_card(card_id: str):
        return service.store.card(card_id)

    @router.post("/cards/download")
    def download_cards(payload: BatchDownload):
        result = io.BytesIO()
        used_names = set()
        try:
            with zipfile.ZipFile(result, "w", zipfile.ZIP_DEFLATED) as archive:
                for card_id in dict.fromkeys(payload.card_ids):
                    card = service.store.card(card_id)
                    filename, data = build_zip(card, [service.download(banner["key"]) for banner in selected_banners(card)])
                    stem, suffix = filename.rsplit(".", 1)
                    unique_name = filename
                    counter = 2
                    while unique_name in used_names:
                        unique_name = f"{stem}_{counter}.{suffix}"
                        counter += 1
                    used_names.add(unique_name)
                    archive.writestr(unique_name, data)
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from exc
        return Response(result.getvalue(), media_type="application/zip", headers={"Content-Disposition": "attachment; filename*=UTF-8''regeneron_fragments.zip"})

    @router.post("/cards/status")
    def update_card_statuses(payload: BatchStatusUpdate):
        return {"cards": [service.store.update(card_id, {"registration_status": payload.registration_status}) for card_id in dict.fromkeys(payload.card_ids)]}

    @router.get("/sheet/rows")
    def sheet_rows():
        try:
            if not service.sheet_reader:
                raise ValueError("スプレッドシートの接続が未設定です。")
            return service.sheet_reader()
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from exc
        except Exception as exc:
            raise HTTPException(502, "スプレッドシートを読み込めません。共有権限と接続を確認してください。") from exc

    @router.get("/button-image")
    def button_image():
        return Response((Path(__file__).parent / "regeneron_assets/btn.png").read_bytes(), media_type="image/png")

    @router.post("/cards/{card_id}/preview")
    def preview(card_id: str, payload: CardUpdate):
        card = {**service.store.card(card_id), **validated_changes(payload)}
        banners = selected_banners(card)
        pages = [next(index for index, item in enumerate(card["banners"]) if item["key"] == banner["key"]) for banner in banners]
        markup = fragment_html(card, len(banners),
            image_source=lambda i: f'/regeneron/cards/{quote(card_id, safe="")}/banners/{pages[i]}?v={quote(banners[i]["key"], safe="")}',
            button_source="/regeneron/button-image") if banners else ""
        return {"html": markup, "image_count": len(banners)}

    @router.get("/cards/{card_id}/attachments")
    def list_attachments(card_id: str):
        card = service.store.card(card_id)
        gmail = Gmail()
        try:
            return {"attachments": [{k: v for k, v in item.items() if k != "part"} for item in gmail.thread_attachments(card["thread_id"])]}
        except Exception as exc:
            raise HTTPException(502, "メール添付を取得できません。時間をおいて再試行してください。") from exc
        finally:
            gmail.session.close()

    @router.get("/cards/{card_id}/attachments/{attachment_id}")
    def download_attachment(card_id: str, attachment_id: str):
        card = service.store.card(card_id)
        gmail = Gmail()
        try:
            attachment = next((item for item in gmail.thread_attachments(card["thread_id"]) if item["id"] == attachment_id), None)
            if not attachment:
                raise HTTPException(404, "添付ファイルが見つかりません。")
            data = gmail.part_data(attachment["message_id"], attachment["part"])
            filename = re.sub(r'[\\/\r\n\x00]', '_', attachment["filename"])
            return Response(data, media_type="application/octet-stream", headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"})
        except HTTPException:
            raise
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from exc
        except Exception as exc:
            raise HTTPException(502, "添付ファイルをダウンロードできませんでした。") from exc
        finally:
            gmail.session.close()

    @router.post("/cards/{card_id}/attachments/{attachment_id}/application")
    def read_application_attachment(card_id: str, attachment_id: str):
        card = service.store.card(card_id)
        gmail = Gmail()
        try:
            attachment = next((item for item in gmail.thread_attachments(card["thread_id"]) if item["id"] == attachment_id), None)
            if not attachment:
                raise HTTPException(404, "添付ファイルが見つかりません。")
            data = gmail.part_data(attachment["message_id"], attachment["part"]) if is_excel(attachment["filename"]) else b""
            if not is_application_workbook(data, attachment["filename"]):
                raise HTTPException(422, "講演会フラグメント申請書を選択してください。")
            parsed = parse_regeneron_申請書(data, attachment["filename"])
            if not parsed:
                raise HTTPException(422, "申請書に有効なフラグメント行がありません。")
            return {"filename": attachment["filename"], "rows": [{
                "row_number": item.get("row_number", index + 5),
                "fragment_name": item["fields"].get("fragmentName", ""),
                "linked_document": linked_document(item["fields"].get("linkedDocument", "")),
                "viewing_url": item["fields"].get("viewingUrl", ""),
            } for index, item in enumerate(parsed)]}
        except HTTPException:
            raise
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from exc
        except Exception as exc:
            raise HTTPException(502, "申請書を読み込めませんでした。") from exc
        finally:
            gmail.session.close()

    @router.post("/cards/{card_id}/source")
    def select_source(card_id: str, payload: SourceSelection):
        card = service.store.card(card_id)
        source = next((s for s in card.get("sources", []) if s["key"] == payload.source_key), None)
        if not source:
            raise HTTPException(404, "添付PDFが見つかりません。")
        try:
            banners = service.banners(service.download(source["key"]), source["filename"])
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from exc
        return service.store.update(card_id, {"banners": banners, "source_key": source["key"], "selected_pages": [0]})

    @router.post("/cards/{card_id}/banner")
    def upload_banner(card_id: str, file: UploadFile = File(...)):
        card = service.store.card(card_id)
        data = file.file.read(MAX_FILE_BYTES + 1)
        filename = file.filename or "banner.png"
        try:
            banners = service.banners(data, filename)
        except Exception as exc:
            raise HTTPException(422, str(exc) if isinstance(exc, ValueError) else "PDF・PNG・JPEG画像を指定してください。") from exc
        changes = {"banners": banners, "source_key": "", "selected_pages": [0]}
        if filename.lower().endswith(".pdf"):
            key = service.asset(data, "source.pdf", "application/pdf")
            changes.update(source_key=key, sources=card.get("sources", []) + [{"key": key, "filename": filename}])
        return service.store.update(card_id, changes)

    @router.get("/cards/{card_id}/banners/{index}")
    def image(card_id: str, index: int):
        card = service.store.card(card_id)
        if index < 0 or index >= len(card.get("banners", [])):
            raise HTTPException(404, "画像が見つかりません。")
        return Response(service.download(card["banners"][index]["key"]), media_type="image/png", headers={"Cache-Control": "private, no-cache"})

    @router.get("/cards/{card_id}/download")
    def download(card_id: str):
        card = service.store.card(card_id)
        try:
            filename, data = build_zip(card, [service.download(b["key"]) for b in selected_banners(card)])
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from exc
        return Response(data, media_type="application/zip", headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"})

    return router
