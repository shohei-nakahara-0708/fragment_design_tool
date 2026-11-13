import base64
import copy
import io
import json
import os
import sqlite3
import sys
import unittest
import zipfile
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fitz
from fastapi import FastAPI
from fastapi.testclient import TestClient
from openpyxl import Workbook
from PIL import Image

from regeneron_application import parse_regeneron_申請書
from regeneron_fragments import FragmentService, Gmail, Store, build_zip, create_router, fragment_html, map_pages, render_images, selected_banners, transcript_message
from regeneron_sheet import parse_sheet_rows, strip_document_version, strip_event_date, transcript_rows


def pdf_bytes(pages=2):
    with fitz.open() as pdf:
        for i in range(pages):
            pdf.new_page(width=600, height=800).insert_text((30, 30), f'Page {i + 1}')
        return pdf.tobytes()


def application(rows=1, url='https://example.com/register?a=1&b=2'):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '講演会フラグメント申請'
    for index in range(rows):
        for column, value in {2: 'MKT', 3: '講演会', 4: f'EV{index}', 5: '2026-10-01', 7: f'講演会{index}', 10: '参加登録', 14: '可', 16: '文書番号\nJP.LIB.26.04.0072', 18: f'invitation{index}'}.items():
            sheet.cell(index + 5, column, value)
        if url:
            sheet.cell(index + 5, 10).hyperlink = url
    buf = io.BytesIO()
    workbook.save(buf)
    return buf.getvalue()


class MemoryStore:
    def __init__(self):
        self.data = {}
        self.messages = set()
        self.exports = {}

    def start_at(self):
        return '2026-09-24T00:00:00+00:00'

    def known(self, ids):
        return self.messages & set(ids)

    def cards(self, thread_id=None):
        return [copy.deepcopy(c) for c in self.data.values() if not thread_id or c['thread_id'] == thread_id]

    def card(self, card_id):
        return copy.deepcopy(self.data[card_id])

    def save_message(self, message_id, cards, export=None):
        if message_id in self.messages:
            return 0
        self.messages.add(message_id)
        if export:
            self.exports[message_id] = {'payload': copy.deepcopy(export), 'done': False}
        for card in cards:
            self.data[card['id']] = copy.deepcopy(card)
        return len(cards)

    def update(self, card_id, changes):
        self.data[card_id].update(changes)
        return self.card(card_id)

    def queue_notice(self, notice_id, text):
        self.exports.setdefault(notice_id, {'payload': {'kind': 'notice', 'message_id': notice_id, 'text': text}, 'done': False})

    def dismiss(self, message_id):
        self.messages.add(message_id)
        if message_id in self.exports:
            self.exports[message_id]['done'] = True

    def process_exports(self, handle):
        for item in self.exports.values():
            if not item['done']:
                item['payload'], error = handle(item['payload'])
                item['done'] = not error


class RegeneronTests(unittest.TestCase):
    def setUp(self):
        self.store = MemoryStore()
        self.assets = {}
        def upload(data, key, content_type):
            self.assets[key] = data
        self.service = FragmentService(self.store, upload, self.assets.__getitem__)
        self.message = {'id': 'mail1', 'threadId': 'thread1', 'internalDate': '1790258400000', 'payload': {'headers': [{'name': 'Subject', 'value': '制作依頼'}]}}
        self.pdf = pdf_bytes()

    def import_card(self):
        self.service.import_message(self.message, [('フラグメント申請書.xlsx', application()), ('invitation0.pdf', self.pdf)])
        return self.store.cards()[0]

    def test_application_hyperlink_and_document(self):
        card = self.import_card()
        self.assertEqual(card['linked_document'], 'JP.LIB.26.04.0072')
        self.assertEqual(card['viewing_url'], 'https://example.com/register?a=1&b=2')
        self.assertEqual(card['subject'], '制作依頼')
        self.assertTrue(card['fragment_name'].startswith('MKT_2026年10月1日'))

    def test_multiple_rows_make_cards_and_match_pdfs(self):
        count = self.service.import_message(self.message, [('申請書.xlsx', application(2)), ('invitation0.pdf', self.pdf), ('invitation1.pdf', self.pdf)])
        self.assertEqual(count, 2)
        for card in self.store.cards():
            selected = next(s for s in card['sources'] if s['key'] == card['source_key'])
            self.assertEqual(selected['filename'], card['application']['documentName'] + '.pdf')

    def test_ambiguous_pdfs_require_selection(self):
        self.service.import_message(self.message, [('申請書.xlsx', application()), ('a.pdf', self.pdf), ('b.pdf', self.pdf)])
        card = self.store.cards()[0]
        self.assertEqual(len(card['sources']), 2)
        self.assertEqual(card['banners'], [])

    def test_pdf_pages_and_zip_structure(self):
        card = self.import_card()
        images = [self.assets[b['key']] for b in card['banners']]
        self.assertEqual(len(images), 2)
        for data in images:
            with Image.open(io.BytesIO(data)) as image:
                self.assertEqual(image.size, (1200, 1600))
        filename, raw = build_zip(card, images)
        self.assertTrue(filename.endswith('.zip'))
        with zipfile.ZipFile(io.BytesIO(raw)) as outer:
            self.assertEqual(len(outer.namelist()), 2)
            markup = outer.read(next(n for n in outer.namelist() if n.endswith('.html'))).decode()
            self.assertIn('a=1&amp;b=2', markup)
            self.assertIn('images/banner2.png', markup)
            with zipfile.ZipFile(io.BytesIO(outer.read('images.zip'))) as inner:
                self.assertEqual(inner.namelist(), ['images/btn.png', 'images/banner.png', 'images/banner2.png'])

    def test_fragment_html_matches_supplied_layout_patterns(self):
        url = 'https://example.com/register'
        no_button = fragment_html({'viewing_url': url, 'include_button': False}, 1)
        one_banner = fragment_html({'viewing_url': url, 'include_button': True}, 1)
        two_banners = fragment_html({'viewing_url': url, 'include_button': True}, 2)

        self.assertEqual(no_button.splitlines()[1], '  <td style="word-break: break-word;border-collapse: collapse !important;vertical-align: top;padding-bottom: 10px;">')
        self.assertIn('      style="border-spacing: 0;border-collapse: collapse;vertical-align: top;max-width: 560px;">', no_button.splitlines())
        self.assertNotIn('images/btn.png', no_button)
        self.assertEqual(no_button.count('<tr>'), 3)
        self.assertIn('                      href="https://example.com/register"><img', no_button.splitlines())
        self.assertIn('                        src="images/banner.png" alt="バナー" width="100%" border="0"', no_button.splitlines())
        self.assertIn('                    style="padding:0 5px 30px; border-collapse:collapse; text-align:center; font-size:0;"><a', one_banner.splitlines())
        self.assertIn('                      href="https://example.com/register"><img src="images/btn.png" alt="参加登録はこちら" width="100%" border="0"', one_banner.splitlines())
        self.assertIn('                      href="https://example.com/register"><img src="images/banner.png" alt="バナー" width="100%" border="0"', one_banner.splitlines())
        self.assertEqual(one_banner.count('<tr>'), 4)
        self.assertIn('                    style="padding:10px 0 0; border-collapse:collapse; text-align:center; font-size:0;"><a', two_banners.splitlines())
        self.assertIn('                      href="https://example.com/register"><img src="images/banner2.png" alt="バナー" width="100%" border="0"', two_banners.splitlines())
        self.assertEqual(two_banners.count('<tr>'), 5)

    def test_no_url_omits_button_and_filename_is_safe(self):
        name, raw = build_zip({'fragment_name': '../name', 'viewing_url': ''}, render_images(self.pdf, 'a.pdf'))
        self.assertNotIn('/', name)
        with zipfile.ZipFile(io.BytesIO(raw)) as outer:
            with zipfile.ZipFile(io.BytesIO(outer.read('images.zip'))) as inner:
                self.assertNotIn('images/btn.png', inner.namelist())

    def test_missing_url_blocks_export(self):
        self.service.import_message(self.message, [('申請書.xlsx', application(url=None)), ('a.pdf', self.pdf)])
        card = self.store.cards()[0]
        self.assertTrue(card['viewing_url_missing'])
        with self.assertRaises(ValueError):
            build_zip(card, [b'png'])

    def test_duplicate_mail_and_followup_do_not_duplicate_cards(self):
        self.import_card()
        self.import_card()
        reply = {**self.message, 'id': 'mail2', 'internalDate': '1790258460000'}
        self.service.import_message(reply, [('invitation0.pdf', self.pdf)])
        self.assertEqual(len(self.store.cards()), 1)
        self.assertEqual(self.store.cards()[0]['message_id'], 'mail2')

    def test_bad_application_is_retryable(self):
        with self.assertRaises(ValueError):
            self.service.import_message(self.message, [('申請書.xls', b'not an xlsx')])
        self.assertNotIn('mail1', self.store.messages)

    def test_application_after_pdf_enriches_existing_card(self):
        self.service.import_message(self.message, [('a.pdf', self.pdf)])
        original_id = self.store.cards()[0]['id']
        reply = {**self.message, 'id': 'mail2', 'internalDate': '1790258460000'}
        self.service.import_message(reply, [('申請書.xlsx', application())])
        self.service.import_message({**reply, 'id': 'mail3'}, [('申請書.xlsx', application())])
        self.assertEqual(len(self.store.cards()), 1)
        card = self.store.cards()[0]
        self.assertEqual(card['id'], original_id)
        self.assertTrue(card['fragment_name'].startswith('MKT_'))
        self.assertEqual(len(card['banners']), 2)

    def test_unrelated_mail_creates_no_card(self):
        self.service.import_message(self.message, [])
        self.assertEqual(self.store.cards(), [])
        self.assertIn('mail1', self.store.messages)

    def test_api_edit_source_download_and_validation(self):
        card = self.import_card()
        app = FastAPI()
        app.include_router(create_router(self.service))
        with TestClient(app) as client:
            path = '/regeneron/cards/' + card['id']
            self.assertEqual(client.get('/regeneron/cards').status_code, 200)
            self.assertEqual(client.patch(path, json={'fragment_name': '  '}).status_code, 422)
            self.assertEqual(client.patch(path, json={'fragment_name': 'Test', 'viewing_url': 'javascript:alert(1)'}).status_code, 422)
            self.assertEqual(client.patch(path, json={'fragment_name': 'Test', 'viewing_url': 'https://[invalid'}).status_code, 422)
            self.assertEqual(client.post(path + '/source', json={'source_key': 'unknown'}).status_code, 404)
            self.assertEqual(client.patch(path, json={'fragment_name': 'Test', 'all_pages': False}).status_code, 200)
            response = client.get(path + '/download')
            self.assertEqual(response.status_code, 200)
            with zipfile.ZipFile(io.BytesIO(response.content)) as outer:
                with zipfile.ZipFile(io.BytesIO(outer.read('images.zip'))) as inner:
                    self.assertEqual(inner.namelist(), ['images/btn.png', 'images/banner.png'])
            upload = client.post(path + '/banner', files={'file': ('three.pdf', pdf_bytes(3), 'application/pdf')})
            self.assertEqual(upload.status_code, 200)
            self.assertEqual(len(upload.json()['banners']), 3)
            self.assertEqual(client.get(path + '/banners/2').status_code, 200)
            self.assertEqual(client.get(path + '/banners/3').status_code, 404)

    def test_registration_status_and_batch_download(self):
        card = self.import_card()
        self.assertEqual(card['registration_status'], 'pending')
        app = FastAPI()
        app.include_router(create_router(self.service))
        with TestClient(app) as client:
            path = '/regeneron/cards/' + card['id']
            saved = client.patch(path, json={'registration_status': 'registered'}).json()
            self.assertEqual(saved['registration_status'], 'registered')
            self.assertEqual(client.patch(path, json={'registration_status': 'unknown'}).status_code, 422)
            batch_status = client.post('/regeneron/cards/status', json={'card_ids': [card['id']], 'registration_status': 'hidden'})
            self.assertEqual(batch_status.status_code, 200)
            self.assertEqual(batch_status.json()['cards'][0]['registration_status'], 'hidden')
            self.assertEqual(client.post('/regeneron/cards/status', json={'card_ids': [card['id']], 'registration_status': 'unknown'}).status_code, 422)
            response = client.post('/regeneron/cards/download', json={'card_ids': [card['id']]})
            self.assertEqual(response.status_code, 200)
            with zipfile.ZipFile(io.BytesIO(response.content)) as archive:
                self.assertEqual(len(archive.namelist()), 1)
                self.assertTrue(archive.namelist()[0].endswith('.zip'))
                with zipfile.ZipFile(io.BytesIO(archive.read(archive.namelist()[0]))) as card_archive:
                    self.assertIn('images.zip', card_archive.namelist())

    def test_sync_filters_old_mail_retries_failure_and_skips_known(self):
        newer = {**self.message, 'id': 'newer', 'internalDate': '1790262000000'}
        old = {**self.message, 'id': 'old', 'internalDate': '1790100000000'}
        service = self.service
        class FakeGmail:
            session = type('Session', (), {'close': lambda self: None})()
            def messages(self, start):
                return [{'id': m['id']} for m in [newer, old, service_message]]
            def get(self, path, **params):
                return next(m for m in [newer, old, service_message] if path.endswith(m['id']))
            def attachments(self, message):
                if message['id'] == 'newer':
                    raise ValueError('broken file')
                return [('申請書.xlsx', application())]
        service_message = self.message
        with patch.dict(os.environ, {'GMAIL_TOKEN_JSON': '{}', 'REGENERON_GMAIL_ENABLED': '1'}), patch('regeneron_fragments.Gmail', FakeGmail):
            result = service.sync()
            self.assertEqual(result['imported'], 1)
            self.assertEqual(len(result['errors']), 1)
            self.assertEqual(service.store.messages, {'mail1'})
            self.assertEqual(service.sync()['imported'], 0)

    def test_default_single_page_and_optional_button_match_preview_and_zip(self):
        card = self.import_card()
        self.assertEqual(card['selected_pages'], [0])
        app = FastAPI()
        app.include_router(create_router(self.service))
        with TestClient(app) as client:
            path = '/regeneron/cards/' + card['id']
            saved = client.patch(path, json={'include_button': False}).json()
            self.assertEqual(saved['fragment_name'], card['fragment_name'])
            self.assertEqual(saved['viewing_url'], card['viewing_url'])
            preview = client.post(path + '/preview', json={}).json()
            self.assertEqual(preview['image_count'], 1)
            self.assertNotIn('button-image', preview['html'])
            self.assertIn('href="https://example.com/register?a=1&amp;b=2"', preview['html'])
            with zipfile.ZipFile(io.BytesIO(client.get(path + '/download').content)) as outer:
                with zipfile.ZipFile(io.BytesIO(outer.read('images.zip'))) as inner:
                    self.assertEqual(inner.namelist(), ['images/banner.png'])
            preview = client.post(path + '/preview', json={'selected_pages': [1, 0], 'include_button': True}).json()
            self.assertEqual(preview['image_count'], 2)
            self.assertIn('/regeneron/button-image', preview['html'])
            self.assertEqual(self.store.card(card['id'])['selected_pages'], [0], 'preview must not persist a draft')
            preview = client.post(path + '/preview', json={'selected_pages': [1]}).json()
            self.assertIn(f'/banners/1?v=', preview['html'])
            self.assertNotIn(f'/banners/0?v=', preview['html'])
            client.patch(path, json={'selected_pages': [1]})
            with zipfile.ZipFile(io.BytesIO(client.get(path + '/download').content)) as outer:
                with zipfile.ZipFile(io.BytesIO(outer.read('images.zip'))) as inner:
                    self.assertEqual(inner.read('images/banner.png'), self.assets[card['banners'][1]['key']])
            self.assertEqual(client.patch(path, json={'selected_pages': []}).status_code, 422)
            self.assertEqual(client.patch(path, json={'selected_pages': [-1]}).status_code, 422)
            self.assertEqual(client.patch(path, json={'viewing_url': None}).status_code, 422)

    def test_map_pages_are_flagged_and_legacy_cards_keep_all_pages_choice(self):
        with fitz.open() as pdf:
            for text in ('Lung Cancer Forum', 'Access', '地 図'):
                pdf.new_page(width=600, height=800).insert_text((30, 30), text, fontname='japan')
            data = pdf.tobytes()
        self.assertEqual(map_pages(data, 'invitation.pdf'), {2})
        self.assertEqual([b['has_map'] for b in self.service.banners(data, 'invitation.pdf')], [False, False, True])
        banners = [{'key': 'a'}, {'key': 'b'}, {'key': 'c'}]
        self.assertEqual(selected_banners({'banners': banners, 'all_pages': True}), banners)
        self.assertEqual(selected_banners({'banners': banners}), banners[:1])
        self.assertEqual(selected_banners({'banners': banners, 'selected_pages': [2, 0, 9]}), [banners[0], banners[2]])
        self.assertEqual(selected_banners({'banners': banners, 'selected_pages': [9]}), banners[:1])

    def test_sheet_mapping_normalizes_multiline_headers_and_hyperlinks(self):
        rows = parse_sheet_rows([['title'], ['フラグメント名', 'Linked Document\n（文書番号）', '視聴予約ページのURL'], ['名前', 'JP.LIB.26.01.0001', '視聴予約'], ['空欄を反映', '', '']], {(2, 2): 'https://example.com'})
        self.assertEqual(rows, [{'row_number': 3, 'fragment_name': '名前', 'linked_document': 'JP.LIB.26.01.0001', 'viewing_url': 'https://example.com'}, {'row_number': 4, 'fragment_name': '空欄を反映', 'linked_document': '', 'viewing_url': ''}])
        with self.assertRaises(ValueError):
            parse_sheet_rows([['フラグメント名', '別の列']])

    def test_attachment_download_is_scoped_to_card_thread_and_keeps_original_file(self):
        card = self.import_card()
        original = b'original spreadsheet bytes'
        part = {'filename': '申請書.xlsx', 'body': {'data': base64.urlsafe_b64encode(original).decode(), 'size': len(original)}}
        gmail = Gmail.__new__(Gmail)
        gmail.session = type('Session', (), {'close': lambda self: None})()
        thread = {'messages': [{**self.message, 'payload': {'parts': [part]}}]}
        calls = []
        def get(path, **params):
            calls.append(path)
            return thread
        gmail.get = get
        app = FastAPI()
        app.include_router(create_router(self.service))
        with patch('regeneron_fragments.Gmail', return_value=gmail), TestClient(app) as client:
            path = '/regeneron/cards/' + card['id'] + '/attachments'
            items = client.get(path).json()['attachments']
            self.assertEqual(items[0]['filename'], '申請書.xlsx')
            self.assertNotIn('part', items[0])
            self.assertEqual(client.get(path + '/' + items[0]['id']).content, original)
            self.assertEqual(client.get(path + '/not-in-this-thread').status_code, 404)
            self.assertTrue(all(call == 'threads/' + card['thread_id'] for call in calls))

    def test_application_attachment_can_be_parsed_without_updating_card(self):
        card = self.import_card()
        original_name = card['fragment_name']
        data = application()
        part = {'filename': '講演会フラグメント申請書.xlsx', 'body': {'data': base64.urlsafe_b64encode(data).decode(), 'size': len(data)}}
        gmail = Gmail.__new__(Gmail)
        gmail.session = type('Session', (), {'close': lambda self: None})()
        gmail.get = lambda path, **params: {'messages': [{**self.message, 'payload': {'parts': [part]}}]}
        app = FastAPI()
        app.include_router(create_router(self.service))
        with patch('regeneron_fragments.Gmail', return_value=gmail), TestClient(app) as client:
            attachment = client.get(f'/regeneron/cards/{card["id"]}/attachments').json()['attachments'][0]
            response = client.post(f'/regeneron/cards/{card["id"]}/attachments/{attachment["id"]}/application')
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json()['rows'][0]['row_number'], 5)
            self.assertEqual(response.json()['rows'][0]['linked_document'], 'JP.LIB.26.04.0072')
            self.assertEqual(self.store.card(card['id'])['fragment_name'], original_name)

    def test_application_mail_is_transcribed_once_and_notified(self):
        body = base64.urlsafe_b64encode('浅田様\nKenta Asada\u3000浅田\u3000健太\n'.encode()).decode()
        message = {**self.message, 'payload': {'headers': [
            {'name': 'Subject', 'value': '【新規依頼】講演会フラグメント申請'},
            {'name': 'From', 'value': '"\'Kenta Asada\' via REGENERON" <regeneron@ar40.jp>'},
            {'name': 'Reply-To', 'value': 'Kenta Asada <kenta.asada@regeneron.com>'}],
            'parts': [{'mimeType': 'text/plain', 'body': {'data': body}}]}}
        written, notices = [], []
        def writer(entry):
            written.append(copy.deepcopy(entry))
            return {'row': 393, 'url': 'https://sheet.example/#range=A393'}
        def notify(text):
            notices.append(text)
            if len(notices) == 1:
                raise RuntimeError('chat down')
        service = FragmentService(self.store, self.service.upload, self.service.download, sheet_writer=writer, notify=notify)
        service.import_message(message, [('フラグメント申請書.xlsx', application(rows=2))])
        service.import_message({**self.message, 'id': 'pdf-only'}, [('invitation0.pdf', self.pdf)])
        summary = {'errors': []}
        self.store.process_exports(lambda entry: service.export_entry(entry, summary))
        self.store.process_exports(lambda entry: service.export_entry(entry, summary))
        self.assertEqual(len(written), 1)
        self.assertEqual(list(self.store.exports), ['mail1'])
        self.assertIn('Google Chat', summary['errors'][0]['message'])
        self.assertEqual(len(notices), 2)
        self.assertIn('申請者: 浅田様', notices[1])
        self.assertIn('管理表: 393行目〜394行目に転記済み', notices[1])
        received, base, mail = transcript_rows(written[0])[0]
        self.assertEqual(received, ['2026/9/24', '新規', '申請内容確認中'])
        self.assertEqual(len(base), 17)
        self.assertEqual(base[:4], ['MKT', '講演会', 'EV0', '2026/10/1'])
        self.assertEqual(base[8], 'https://example.com/register?a=1&b=2')
        self.assertEqual(base[14], 'JP.LIB.26.04.0072')
        self.assertEqual(mail, ['浅田様', '【新規依頼】講演会フラグメント申請'])

    def test_transcript_quotes_values_sheets_would_reinterpret(self):
        entry = {'received_at': '2026-10-26T16:00:00+00:00', 'subject': '=IMPORTXML()', 'requester': '',
                 'rows': [{'category': '修正', 'eventTime': '18:30-19:40', 'meetingId': '224 100 418 288', 'eventDate': '12月10日'}]}
        received, base, mail = transcript_rows(entry)[0]
        self.assertEqual(received[0], '2026/10/27')
        self.assertEqual(base[4], "'18:30-19:40")
        self.assertEqual(base[9], "'224 100 418 288")
        self.assertEqual(base[3], '12月10日')
        self.assertEqual(mail[1], "'=IMPORTXML()")

    def test_event_name_newlines_and_dates_are_cleaned_and_reported(self):
        entry = {'received_at': '2026-10-06T00:29:27+00:00', 'subject': '件名', 'requester': '',
                 'rows': [{'category': '新規', 'eventName': '20260929_REGENERON Lung Cancer\nWebinar', 'documentName': 'A \nB',
                           'fragment_name': 'MKT_2026年9月29日_REGENERON Lung Cancer Webinar_他エリア案内可'}]}
        base = transcript_rows(entry)[0][1]
        self.assertEqual(base[5], 'REGENERON Lung Cancer Webinar')
        self.assertEqual(base[16], 'A B')
        row = entry['rows'][0]
        self.assertEqual([c['column'] for c in row['changes']], ['K', 'V'])
        message = transcript_message({**entry, 'sheet': {'row': 400, 'url': 'https://sheet'}})
        self.assertIn('【申請書から整えて転記した箇所】\n講演会名: 「20260929_REGENERON Lung Cancer⏎Webinar」→「REGENERON Lung Cancer Webinar」', message)
        self.assertNotRegex(message, '[📝🔴⚠️🚨]')
        for name in ('第65回日本肺癌学会 LS1', 'Session 1-2', 'Seminar 2026'):
            self.assertEqual(strip_event_date(name), name)

    def test_dates_and_multiple_time_slots_are_normalized(self):
        from regeneron_application import _format_date, _format_time
        self.assertEqual(_format_date('2026年 12月 16日(水) 19:00-19:45'), '2026-12-16')
        self.assertEqual(_format_time('2026年 12月 16日(水) 19:00-19:45'), '19:00-19:45')
        self.assertEqual(_format_time('12:30-13:00\n13:15-13:45'), '12:30-13:00\n13:15-13:45')
        self.assertEqual(_format_time('2026年12月16日(水) 12:30～13:00 / 13:15～13:45'), '12:30-13:00\n13:15-13:45')
        self.assertEqual(_format_time('12月16日（水）'), '')
        self.assertEqual(_format_date('2026年 1 2月 10日(木) 12:00-13:00'), '2026-12-10')
        self.assertEqual(_format_time('2026年 1 2月 10日(木) 12:00-13:00'), '12:00-13:00')
        self.assertEqual(_format_date('2026/12/1 12:00'), '2026-12-01')
        entry = {'received_at': '2026-10-16T01:00:00+00:00', 'subject': 's', 'rows': [{'category': '新規', 'eventTime': '12:30-13:00\n13:15-13:45', 'eventTimeRaw': '第1部：12:30-13:00\n第2部：13:15-13:45'}]}
        self.assertEqual(transcript_rows(entry)[0][1][4], '第1部：12:30-13:00\n第2部：13:15-13:45')
        entry['rows'][0]['eventTimeRaw'] = '18:30~19:50'
        self.assertEqual(transcript_rows(entry)[0][1][4], "'18:30~19:50")
        self.assertEqual(entry['rows'][0]['changes'], [])

    def test_document_name_drops_parenthesized_version_only(self):
        self.assertEqual(strip_document_version('20261125_ONC中日本営業所_Regeneron肺癌セミナー in 新潟_案内状 (v0.1)'),
                         '20261125_ONC中日本営業所_Regeneron肺癌セミナー in 新潟_案内状')
        self.assertEqual(strip_document_version('案内状（Ver.1.0）'), '案内状')
        self.assertEqual(strip_document_version('名前 Fix 案内状_20261223'), 'Fix 案内状_20261223')
        for name in ('案内状Ver.2', '案内状_v2', 'LBT13-M25002-2.0_動画 (2.0)'):
            self.assertEqual(strip_document_version(name), name)

    def test_card_fragment_name_uses_cleaned_event_name(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = '講演会フラグメント申請'
        for column, value in {2: 'Sales', 3: '講演会', 4: 'EV1', 5: '2026-11-18', 7: '20261118_REGENERON Lung Cancer \nWebinar', 13: '福岡県', 14: '可'}.items():
            sheet.cell(5, column, value)
        buf = io.BytesIO()
        workbook.save(buf)
        fields = parse_regeneron_申請書(buf.getvalue(), '申請書.xlsx')[0]['fields']
        self.assertEqual(fields['fragmentName'], 'Sales_2026年11月18日_福岡県_REGENERON Lung Cancer Webinar_他エリア案内可')
        self.assertEqual(fields['eventName'], '20261118_REGENERON Lung Cancer \nWebinar')

    def test_v41_form_with_annotated_kubun_and_us_dates(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = '講演会フラグメント申請'
        for column, value in {2: 'MKT（CSCC）', 3: '講演会', 4: 'E-0000034098', 5: '12/15/2026', 6: '12:00-13:00', 7: 'REGENERON Skin Cancer Webinar', 14: '可', 15: '12/15/2026', 17: '10/18/2026'}.items():
            sheet.cell(5, column, value)
        buf = io.BytesIO()
        workbook.save(buf)
        fields = parse_regeneron_申請書(buf.getvalue(), '申請書.xlsx')[0]['fields']
        self.assertEqual((fields['kubun1'], fields['eventDate'], fields['expiration'], fields['approvalDate']), ('MKT', '2026-12-15', '2026-12-15', '2026-10-18'))
        self.assertEqual(fields['fragmentName'], 'MKT_2026年12月15日_REGENERON Skin Cancer Webinar_他エリア案内可')
        entry = {'received_at': '2026-10-28T01:00:00+00:00', 'subject': 's', 'rows': [{**fields, 'category': '新規'}]}
        options = ['Sales', 'MKT（NSCLC）', 'MKT（CSCC）', 'MKT（Cervical）']
        self.assertEqual(transcript_rows(entry, kubun1_options=options)[0][1][0], 'MKT（CSCC）')
        self.assertEqual((entry['rows'][0]['changes'], entry['rows'][0]['warnings']), ([], []))
        entry['rows'][0]['kubun1Original'] = 'mkt (cscc)'
        self.assertEqual(transcript_rows(entry, kubun1_options=options)[0][1][0], 'MKT（CSCC）')
        self.assertEqual(entry['rows'][0]['changes'][0]['after'], 'MKT（CSCC）')
        legacy = {**entry, 'rows': [{'category': '新規', 'kubun1': 'MKT'}]}
        self.assertEqual(transcript_rows(legacy, kubun1_options=options)[0][1][0], 'MKT')
        self.assertEqual(legacy['rows'][0]['warnings'][0]['column'], 'F')
        self.assertIn('区分①: 「MKT」はシートの選択肢にありません', transcript_message({**legacy, 'subject': 's', 'sheet': {'row': 5, 'url': 'u'}, 'rows': [{**legacy['rows'][0], 'fragment_name': 'f'}]}))

    def test_failed_mail_can_be_dismissed(self):
        app = FastAPI()
        app.include_router(create_router(self.service))
        self.service.last = {**self.service.last, 'errors': [{'message_id': 'bad', 'message': '申請書に有効なフラグメント行がありません。'}]}
        with TestClient(app) as client:
            status = client.post('/regeneron/sync/errors/bad/dismiss').json()
        self.assertEqual(status['errors'], [])
        self.assertIn('bad', self.store.known(['bad']))

    def test_older_form_without_approval_column_is_read_by_headers(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = '講演会フラグメント申請'
        headers = ['＃', '区分①\nMKT/Sales', '区分②', '講演会ID', '開催日', '時間', '講演会名', '領域', '製品', '視聴予約ページのURL',
                   '会議ID', 'パスコード', '都道府県', '他エリアへの\n案内可否', '使用期限', '文書番号', 'ドキュメント名']
        for column, header in enumerate(headers, start=1):
            sheet.cell(3, column, header)
        for column, value in {2: 'MKT', 3: '講演会', 5: '2027-01-19', 7: 'Seminar', 14: '可', 16: 'JP.LIB.26.11.0013', 17: '20270119_MKT_案内状'}.items():
            sheet.cell(5, column, value)
        buf = io.BytesIO()
        workbook.save(buf)
        fields = parse_regeneron_申請書(buf.getvalue(), '申請書.xlsx')[0]['fields']
        self.assertEqual((fields['linkedDocument'], fields['approvalDate'], fields['documentName']), ('JP.LIB.26.11.0013', '', '20270119_MKT_案内状'))
        sheet.insert_rows(1)
        sheet.cell(5, 1, '例')
        sheet.cell(5, 2, 'Sales')
        buf = io.BytesIO()
        workbook.save(buf)
        parsed = parse_regeneron_申請書(buf.getvalue(), '申請書.xlsx')
        self.assertEqual([item['fields']['documentName'] for item in parsed], ['20270119_MKT_案内状'])

    def test_time_cell_keeps_staff_notation_without_date(self):
        from regeneron_application import _sheet_time
        for raw, expected in [('2026年 11月25日(水) 19:00-20:15', '19:00-20:15'), ('2026年12月9日(水) 18:30～19:50', '18:30～19:50'),
                              ('2026年 1 2月 10日(木) 12:00-13:00', '12:00-13:00'), ('2026 年 11 月 30 日 (月) 19:00-20:20', '19:00-20:20'),
                              ('2026年 11月 25日(水) \n12:15-12:45 / 13:00-13:30', '12:15-12:45\n13:00-13:30'), ('2026年10月30日(金)18:30~19:30', '18:30~19:30'),
                              ('18:30~19:50', '18:30~19:50'), ('第1部：12:30-13:00\n第2部：13:15-13:45', '第1部：12:30-13:00\n第2部：13:15-13:45'), ('18.30-19.30', '18.30-19.30')]:
            self.assertEqual(_sheet_time(raw), expected, raw)

    def test_v40_form_combines_mkt_with_indication(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = '講演会フラグメント申請'
        headers = ['＃', '区分①\nMKT/Sales', '区分②\nNSCLC/CSCC/Cervical', '区分③\n講演会/学会共催セミナー', '講演会ID', '開催日', '時間', '講演会名', '領域', '製品',
                   '視聴予約ページのURL', '会議ID', 'パスコード', '都道府県', '他エリアへの\n案内可否', '使用期限', '文書番号', 'ドキュメント名']
        for column, header in enumerate(headers, start=1):
            sheet.cell(3, column, header)
        for column, value in {1: '例', 2: 'Sales', 3: 'NSCLC', 4: '講演会', 8: '例の講演会'}.items():
            sheet.cell(4, column, value)
        for column, value in {2: 'MKT', 3: 'CSCC', 4: '学会共催セミナー', 6: '2026-12-19', 8: 'セミナー', 15: '可'}.items():
            sheet.cell(5, column, value)
        buf = io.BytesIO()
        workbook.save(buf)
        parsed = parse_regeneron_申請書(buf.getvalue(), '申請書.xlsx')
        self.assertEqual(len(parsed), 1)
        fields = parsed[0]['fields']
        self.assertEqual((fields['kubun1'], fields['kubun1Original'], fields['kubun2']), ('MKT', 'MKT（CSCC）', '学会共催セミナー'))
        self.assertEqual(fields['fragmentName'], 'MKT_2026年12月19日_セミナー')

    def test_linked_document_restores_country_code_and_missing_approval_is_flagged(self):
        from regeneron_fragments import linked_document
        self.assertEqual(linked_document('P.LIB.26.09.0006'), 'JP.LIB.26.09.0006')
        self.assertEqual(linked_document('文書番号\nJP.LIB.26.11.0015'), 'JP.LIB.26.11.0015')
        entry = {'received_at': '2026-11-01T00:00:00+00:00', 'subject': 's', 'rows': [
            {'category': '新規', 'linkedDocument': 'JP.LIB.26.09.0006', 'linkedDocumentRaw': 'P.LIB.26.09.0006', 'approvalDate': ''},
            {'category': '新規', 'linkedDocument': 'JP.LIB.26.11.0015', 'linkedDocumentRaw': '文書番号\nJP.LIB.26.11.0015', 'approvalDate': '2026-11-06'}]}
        transcript_rows(entry)
        first, second = entry['rows']
        self.assertEqual(first['changes'], [{'column': 'T', 'label': 'Linked Document', 'before': 'P.LIB.26.09.0006', 'after': 'JP.LIB.26.09.0006'}])
        self.assertEqual([w['column'] for w in first['warnings']], ['U'])
        self.assertEqual((second['changes'], second['warnings']), ([], []))

    def test_application_is_detected_by_content_not_filename(self):
        from regeneron_application import is_application_workbook
        data = application()
        self.assertTrue(is_application_workbook(data, '20261212_v3.0.xlsx'))
        self.assertTrue(is_application_workbook(data, '20261212_v3.0.xlsm'))
        other = Workbook()
        other.active['A1'] = '見積'
        buf = io.BytesIO()
        other.save(buf)
        self.assertFalse(is_application_workbook(buf.getvalue(), '見積.xlsx'))
        self.assertFalse(is_application_workbook(b'broken', 'data.xlsx'))
        self.assertTrue(is_application_workbook(b'broken', '申請書.xls'), 'named files still surface a parse error')
        self.service.import_message(self.message, [('20261212_v3.0.xlsx', data), ('見積.xlsx', buf.getvalue())])
        self.assertEqual(len(self.store.cards()), 1)

    def test_request_mail_without_application_notifies_once(self):
        notices = []
        service = FragmentService(self.store, self.service.upload, self.service.download, sheet_writer=lambda entry: self.fail('no sheet write'), notify=notices.append)
        def mail(message_id, subject, thread='t-missing'):
            return {'id': message_id, 'threadId': thread, 'internalDate': '1790258400000', 'payload': {'headers': [{'name': 'Subject', 'value': subject}]}}
        service.import_message(mail('m1', '【新規依頼】講演会フラグメント申請（2026年12月1日 Seminar）'), [('invitation.pdf', self.pdf)])
        service.import_message(mail('m2', 'RE: 【新規依頼】講演会フラグメント申請'), [])
        service.import_message(mail('m3', '【新規依頼】講演会フラグメント申請（添付なし）', 't-other'), [])
        service.import_message(mail('m4', 'ご連絡', 't-third'), [])
        service.import_message(mail('m4b', '取り消し: 【新規依頼】講演会フラグメント申請', 't-recall'), [])
        summary = {'errors': []}
        self.store.process_exports(lambda entry: service.export_entry(entry, summary))
        self.store.process_exports(lambda entry: service.export_entry(entry, summary))
        self.assertEqual(len(notices), 2)
        self.assertTrue(all(n.startswith('【リジェネロン】申請書が添付されていない') for n in notices))
        service.import_message(mail('m5', '【修正依頼】講演会フラグメント申請', 't-done'), [('申請書.xlsx', application())])
        service.import_message(mail('m6', '【修正依頼】講演会フラグメント申請 再送', 't-done'), [])
        self.assertNotIn('m6', self.store.exports)

    def test_import_errors_and_outages_and_stuck_transcripts_alert_chat_once(self):
        notices = []
        failing_writer = {'fail': True}
        def writer(entry):
            if failing_writer['fail']:
                raise RuntimeError('sheet down')
            return {'row': 10, 'url': 'u'}
        service = FragmentService(self.store, self.service.upload, self.service.download, sheet_writer=writer, notify=notices.append)
        bad = {**self.message, 'id': 'bad', 'threadId': 'tb', 'payload': {'headers': [{'name': 'Subject', 'value': '【新規依頼】申請'}]}}
        good = {**self.message, 'id': 'good', 'threadId': 'tg'}
        class FakeGmail:
            session = type('Session', (), {'close': lambda self: None})()
            def messages(self, start): return [{'id': 'bad'}, {'id': 'good'}]
            def get(self, path, **params): return bad if path.endswith('bad') else good
            def attachments(self, message):
                if message['id'] == 'bad':
                    raise ValueError('申請書に有効なフラグメント行がありません。')
                return [('申請書.xlsx', application())]
        with patch.dict(os.environ, {'GMAIL_TOKEN_JSON': '{}'}), patch('regeneron_fragments.Gmail', FakeGmail):
            for _ in range(4):
                service.sync()
        # A: the broken mail is reported once even though it is retried every sync.
        self.assertEqual(sum('取り込めませんでした' in n and '有効なフラグメント行がありません' in n for n in notices), 1)
        # C: the transcript failed 4 times; one alert after the third failure.
        self.assertEqual(sum('転記できていない' in n for n in notices), 1)
        failing_writer['fail'] = False
        with patch.dict(os.environ, {'GMAIL_TOKEN_JSON': '{}'}), patch('regeneron_fragments.Gmail', FakeGmail):
            service.sync()
        self.assertEqual(sum('申請が届きました' in n for n in notices), 1)
        # B: Gmail keeps failing; alert on the third consecutive failure, once per day.
        def broken():
            raise RuntimeError('token expired')
        with patch.dict(os.environ, {'GMAIL_TOKEN_JSON': '{}'}), patch('regeneron_fragments.Gmail', broken):
            for _ in range(5):
                service.sync()
        self.assertEqual(sum('止まっています' in n for n in notices), 1)
        self.assertEqual(service.sync_failures, 5)

    def test_disabled_instance_neither_imports_nor_transcribes_nor_notifies(self):
        notices = []
        service = FragmentService(self.store, self.service.upload, self.service.download, sheet_writer=lambda entry: self.fail('sheet'), notify=notices.append)
        self.store.queue_notice('pending', 'queued before the switch')
        def gmail():
            raise AssertionError('Gmail must not be called')
        with patch.dict(os.environ, {'GMAIL_TOKEN_JSON': '{}', 'REGENERON_GMAIL_ENABLED': '0'}), patch('regeneron_fragments.Gmail', gmail):
            service.sync()
            service.start()
        self.assertIsNone(service.thread)
        self.assertEqual(notices, [])
        self.assertFalse(self.store.exports['pending']['done'])

    def test_ip_allowlist_protects_api_but_not_the_ip_probe(self):
        app = FastAPI()
        app.include_router(create_router(self.service))
        with TestClient(app) as client:
            self.assertEqual(client.get('/regeneron/cards').status_code, 200, 'no allowlist means no restriction')
            with patch.dict(os.environ, {'REGENERON_ALLOWED_IPS': '203.0.113.5, 198.51.100.0/24', 'REGENERON_CLIENT_IP_HEADER': 'x-forwarded-for:-1'}):
                self.assertEqual(client.get('/regeneron/cards', headers={'X-Forwarded-For': '203.0.113.5'}).status_code, 200)
                self.assertEqual(client.get('/regeneron/cards', headers={'X-Forwarded-For': '198.51.100.77'}).status_code, 200)
                # A client-supplied first entry cannot spoof past the proxy-appended last entry.
                self.assertEqual(client.get('/regeneron/cards', headers={'X-Forwarded-For': '203.0.113.5, 192.0.2.9'}).status_code, 403)
                self.assertEqual(client.post('/regeneron/sync').status_code, 403)
                probe = client.get('/regeneron/client-ip', headers={'X-Forwarded-For': '192.0.2.9'}).json()
                self.assertEqual(probe['detected'], '192.0.2.9')
            with patch.dict(os.environ, {'REGENERON_ALLOWED_IPS': '203.0.113.5'}):
                self.assertEqual(client.get('/regeneron/cards', headers={'X-Forwarded-For': '203.0.113.5'}).status_code, 403, 'headers are ignored unless configured')

    def test_transcript_notice_layout(self):
        from datetime import date
        from regeneron_sheet import add_workdays
        self.assertEqual(add_workdays(date(2026, 11, 12), 2, set()), date(2026, 11, 16))
        self.assertEqual(add_workdays(date(2026, 11, 20), 2, {date(2026, 11, 23)}), date(2026, 11, 25))
        entry = {'subject': ' 【新規依頼】申請 ', 'requester': '坂口様', 'thread_id': 'th1', 'rfc822_id': '<ABC@mail.example>',
                 'sheet': {'row': 401, 'url': 'https://sheet', 'release_date': '2026/11/11'},
                 'rows': [{'category': '修正', 'changes': [{'label': '講演会名', 'before': 'A\nB', 'after': 'A B'}], 'warnings': []}]}
        self.assertEqual(transcript_message(entry), '\n'.join([
            '【リジェネロン】フラグメントの修正申請が届きました', '件名: 【新規依頼】申請', '申請者: 坂口様',
            '本番化予定日: 2026/11/11', '管理表: 401行目に転記済み', '',
            '【申請書から整えて転記した箇所】', '講演会名: 「A⏎B」→「A B」', '',
            '内容を確認の上、作業者に依頼してください', '',
            '<https://sheet|スプレッドシートを開く>', '<https://mail.google.com/mail/u/0/#search/rfc822msgid%3A%3CABC%40mail.example%3E|該当メールを開く>']))
        entry['rows'].append({'category': '新規', 'changes': [], 'warnings': [{'label': '承認日', 'message': '入力してください'}]})
        message = transcript_message(entry)
        self.assertTrue(message.startswith('【リジェネロン】フラグメント申請が届きました'))
        self.assertIn('管理表: 401行目〜402行目に転記済み', message)
        self.assertIn('区分: 401行目は修正', message)
        self.assertIn('401行目 講演会名: 「A⏎B」→「A B」', message)
        self.assertIn('【シートで確認・入力が必要な箇所】\n402行目 承認日: 入力してください', message)

    def test_known_event_id_in_sheet_marks_correction(self):
        entry = {'received_at': '2026-10-06T00:29:27+00:00', 'subject': 'RE: 【新規依頼】講演会フラグメント申請',
                 'rows': [{'category': '新規', 'eventId': 'E-0000033351'}, {'category': '新規', 'eventId': ''}]}
        rows = transcript_rows(entry, {'E-0000033351', 'n/a'})
        self.assertEqual([row[0][1] for row in rows], ['修正', '新規'])
        placeholder = {**entry, 'rows': [{'category': '新規', 'eventId': 'n/a'}]}
        self.assertEqual(transcript_rows(placeholder, {'n/a'})[0][0][1], '新規')
        self.assertEqual(entry['rows'][0]['category'], '修正')

    def test_gmail_pagination_and_nested_inline_attachments(self):
        gmail = Gmail.__new__(Gmail)
        calls = []
        def get(path, **params):
            calls.append(params)
            if not params.get('pageToken'):
                return {'messages': [{'id': 'a'}], 'nextPageToken': 'next'}
            return {'messages': [{'id': 'b'}]}
        gmail.get = get
        with patch.dict(os.environ, {'REGENERON_GMAIL_QUERY': 'from:regeneron@ar40.jp'}):
            self.assertEqual(len(gmail.messages('2026-09-24T00:00:00+00:00')), 2)
        self.assertEqual(calls[1]['pageToken'], 'next')
        raw = base64.urlsafe_b64encode(self.pdf).decode().rstrip('=')
        message = {**self.message, 'payload': {'parts': [{'parts': [{'filename': 'a.pdf', 'body': {'data': raw}}]}]}}
        self.assertEqual(gmail.attachments(message), [('a.pdf', self.pdf)])


# Exercise repository conflict/update behaviour with a small SQLite SQL adapter.
# Production remains PostgreSQL; no live application database is touched by tests.
class StoreTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(':memory:')
        self.db.row_factory = sqlite3.Row
        class Cursor:
            def __init__(self, cursor): self.cursor = cursor
            def fetchone(self):
                row = self.cursor.fetchone()
                if row is None: return None
                row = dict(row)
                if 'payload' in row: row['payload'] = json.loads(row['payload'])
                return row
            def fetchall(self):
                result = []
                while True:
                    row = self.fetchone()
                    if row is None: return result
                    result.append(row)
        class Connection:
            def execute(_, sql, args=()):
                if sql.startswith("UPDATE regeneron_cards SET payload=jsonb_set"):
                    for row in self.db.execute('SELECT id,payload FROM regeneron_cards').fetchall():
                        data = json.loads(row['payload'])
                        if 'all_pages' not in data.get('edited_fields', []):
                            data['all_pages'] = False
                            self.db.execute('UPDATE regeneron_cards SET payload=? WHERE id=?', (json.dumps(data), row['id']))
                    return Cursor(self.db.execute('SELECT 1'))
                sql = sql.replace('::jsonb', '').replace(' FOR UPDATE', '').replace('%s', '?')
                return Cursor(self.db.execute(sql, args))
            def commit(_): self.db.commit()
        @contextmanager
        def connect():
            try:
                yield Connection()
                self.db.commit()
            except Exception:
                self.db.rollback()
                raise
        self.store = Store(connect)
        self.store.init()

    def tearDown(self): self.db.close()

    def test_default_migration_preserves_explicit_multi_page_choice(self):
        base = {'thread_id': 't', 'received_at': '2026-09-24', 'all_pages': True}
        self.store.save_message('a', [{**base, 'id': 'implicit'}])
        self.store.save_message('b', [{**base, 'id': 'explicit', 'edited_fields': ['all_pages']}])
        self.db.execute("DELETE FROM regeneron_settings WHERE key='single_page_default_v2'")
        self.db.commit()
        self.store.init()
        self.assertFalse(self.store.card('implicit')['all_pages'])
        self.assertTrue(self.store.card('explicit')['all_pages'])

    def test_existing_fragment_names_are_cleaned_once_unless_edited(self):
        application = {'kubun1': 'MKT', 'kubun2': '講演会', 'eventDate': '2026-09-29', 'eventName': '20260929_LC\nWebinar', 'areaGuide': '可'}
        base = {'thread_id': 't', 'received_at': '2026-09-24', 'fragment_name': 'MKT_2026年9月29日_20260929_LC\nWebinar_他エリア案内可'}
        self.store.save_message('a', [{**base, 'id': 'auto', 'application': application}])
        self.store.save_message('b', [{**base, 'id': 'manual', 'application': application, 'fragment_name': 'manual', 'edited_fields': ['fragment_name']}])
        self.db.execute("DELETE FROM regeneron_settings WHERE key='clean_fragment_name_v1'")
        self.db.commit()
        self.store.init()
        self.assertEqual(self.store.card('auto')['fragment_name'], 'MKT_2026年9月29日_LC Webinar_他エリア案内可')
        self.assertEqual(self.store.card('manual')['fragment_name'], 'manual')

    def test_dismiss_claims_mail_and_finishes_export(self):
        self.store.save_message('m1', [], {'message_id': 'm1', 'rows': []})
        self.store.dismiss('m1')
        self.store.dismiss('m2')
        self.assertEqual({row['id'] for row in self.db.execute('SELECT id FROM regeneron_messages')}, {'m1', 'm2'})
        self.store.process_exports(lambda entry: self.fail('dismissed export must not run'))

    def test_notices_are_queued_once_per_id(self):
        self.store.queue_notice('sync-error:2026-11-12', 'first')
        self.store.queue_notice('sync-error:2026-11-12', 'second')
        sent = []
        self.store.process_exports(lambda entry: (sent.append(entry['text']), (entry, None))[1])
        self.store.process_exports(lambda entry: (sent.append(entry['text']), (entry, None))[1])
        self.assertEqual(sent, ['first'])

    def test_start_boundary_persists_across_restarts(self):
        with patch.dict(os.environ, {'REGENERON_GMAIL_START_AT': '2026-09-24T10:00:00+09:00'}):
            start = self.store.start_at()
        with patch.dict(os.environ, {'REGENERON_GMAIL_START_AT': '2026-10-01T10:00:00+09:00'}):
            self.assertEqual(self.store.start_at(), start)

    def test_mail_claim_and_manual_changes_survive_followup(self):
        card = {'id': 'a', 'thread_id': 't', 'received_at': '2026-09-24T01:00:00+00:00', 'fragment_name': 'original', 'sources': []}
        self.assertEqual(self.store.save_message('m1', [card]), 1)
        self.store.update('a', {'fragment_name': 'edited', 'banners': [{'key': 'manual'}]})
        updated = {**card, 'received_at': '2026-09-24T02:00:00+00:00', 'fragment_name': 'incoming', 'banners': []}
        self.assertEqual(self.store.save_message('m2', [updated]), 1)
        saved = self.store.card('a')
        self.assertEqual(saved['fragment_name'], 'edited')
        self.assertEqual(saved['banners'][0]['key'], 'manual')
        self.assertEqual(self.store.save_message('m2', [updated]), 0)
        self.store.save_message('late-old-mail', [card])
        self.assertEqual(self.store.card('a')['received_at'], updated['received_at'])

    def test_export_queue_is_claimed_with_mail_and_retried_until_done(self):
        card = {'id': 'a', 'thread_id': 't', 'received_at': '2026-09-24T01:00:00+00:00', 'sources': []}
        self.store.save_message('m1', [card], {'message_id': 'm1', 'rows': []})
        self.store.save_message('m1', [card], {'message_id': 'duplicate', 'rows': []})
        seen = []
        def handle(entry):
            seen.append(entry['message_id'])
            return {**entry, 'sheet': {'row': 5}}, 'retry' if len(seen) == 1 else None
        self.store.process_exports(handle)
        self.store.process_exports(handle)
        self.store.process_exports(handle)
        self.assertEqual(seen, ['m1', 'm1'])
        row = self.db.execute('SELECT payload FROM regeneron_exports').fetchone()
        self.assertEqual(json.loads(row['payload'])['sheet']['row'], 5)


if __name__ == '__main__':
    unittest.main()
