"""Read fragment fields from the shared Google Sheet and transcribe new applications to the tracking sheet."""
import os
import re
import unicodedata
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from google.auth.transport.requests import AuthorizedSession

from regeneron_application import one_line, strip_event_date

SHEET_ID = '1dcFZLY6aD0i9jJS5NDjZM3twDoj4MYXyehxrhXsGnMI'
SHEET_URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?gid=0'
TRANSCRIPT_SHEET_ID = '1M1j_cJehL8Po7VHWr2gjXWen5AAUj6-Rh9aebYpzqtk'
TRANSCRIPT_FIRST_ROW = 4
# Formula columns (E, W:Z) and manual columns (A, AA) are left untouched.
TRANSCRIPT_SEGMENTS = (('B', 'D'), ('F', 'V'), ('AB', 'AC'))
EVENT_ID_COLUMN = 'H'
KUBUN1_COLUMN = 'F'
JST = timezone(timedelta(hours=9))


def normalized(value):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value or ''))).lower()


def parse_sheet_rows(values, hyperlinks=None):
    hyperlinks = hyperlinks or {}
    headers = None
    for index, row in enumerate(values[:20]):
        names = [normalized(cell) for cell in row]
        required = {'fragment_name': 'フラグメント名', 'linked_document': 'linkeddocument(文書番号)', 'viewing_url': '視聴予約ページのurl'}
        if all(label in names for label in required.values()):
            headers = {key: names.index(label) for key, label in required.items()}
            start = index + 1
            break
    if headers is None:
        raise ValueError('フラグメント名・Linked Document（文書番号）・視聴予約ページのURLの列が見つかりません。')
    rows = []
    for index, row in enumerate(values[start:], start=start):
        fields = {key: str(row[col] if col < len(row) else '').strip() for key, col in headers.items()}
        if not fields['fragment_name']:
            continue
        fields['viewing_url'] = hyperlinks.get((index, headers['viewing_url'])) or fields['viewing_url']
        rows.append({'row_number': index + 1, **fields})
    return rows


def first_sheet_properties(session, base):
    response = session.get(base, params={'fields': 'sheets(properties)'}, timeout=30)
    response.raise_for_status()
    properties = next((s['properties'] for s in response.json().get('sheets', []) if s['properties']['sheetId'] == 0), None)
    if not properties:
        raise ValueError('指定されたシート（gid=0）が見つかりません。')
    return properties


def read_sheet(credentials_factory):
    credentials = credentials_factory(['https://www.googleapis.com/auth/spreadsheets.readonly'])
    with AuthorizedSession(credentials) as session:
        base = f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}'
        title = first_sheet_properties(session, base)['title']
        escaped = title.replace("'", "''")
        response = session.get(base, params={
            'ranges': f"'{escaped}'",
            'includeGridData': 'true',
            'fields': 'sheets(data(startRow,startColumn,rowData(values(formattedValue,hyperlink))))',
        }, timeout=40)
        response.raise_for_status()
        values, hyperlinks = [], {}
        for sheet in response.json().get('sheets', []):
            for block in sheet.get('data', []):
                start_row, start_col = block.get('startRow', 0), block.get('startColumn', 0)
                for offset, row in enumerate(block.get('rowData', [])):
                    index = start_row + offset
                    while len(values) <= index:
                        values.append([])
                    cells = row.get('values', [])
                    values[index] = [''] * start_col + [c.get('formattedValue', '') for c in cells]
                    for col, cell in enumerate(cells, start=start_col):
                        if cell.get('hyperlink'):
                            hyperlinks[(index, col)] = cell['hyperlink']
        return {'title': title, 'url': SHEET_URL, 'rows': parse_sheet_rows(values, hyperlinks)}


def sheet_text(value):
    text = str(value or '').strip()
    # Values are USER_ENTERED so dates stay dates; quote text Sheets would turn into a formula or number.
    if re.match(r"[=+\-@']", text) or re.fullmatch(r'[\d\s.,:/~～\-%()]+', text):
        return "'" + text
    return text


def sheet_date(value):
    text = str(value or '').strip()
    match = re.fullmatch(r'(\d{4})-(\d{2})-(\d{2})', text)
    # Same display as the hand-entered rows: 2026/11/5, not 2026/11/05.
    return f'{match[1]}/{int(match[2])}/{int(match[3])}' if match else sheet_text(text)


# F:W in sheet order: (column, application field, header label, kind)
TRANSCRIPT_COLUMNS = (
    ('F', 'kubun1', '区分①', 'kubun1'), ('G', 'kubun2', '区分②', 'text'),
    ('H', 'eventId', '講演会ID', 'text'), ('I', 'eventDate', '開催日', 'date'), ('J', 'eventTime', '時間', 'time'),
    ('K', 'eventName', '講演会名', 'event_name'), ('L', 'area', '領域', 'text'), ('M', 'product', '製品', 'text'),
    ('N', 'viewingUrl', '視聴予約ページのURL', 'url'), ('O', 'meetingId', '会議ID', 'text'), ('P', 'passcode', 'パスコード', 'text'),
    ('Q', 'prefecture', '都道府県', 'text'), ('R', 'areaGuide', '他エリアへの案内可否', 'text'), ('S', 'expiration', '使用期限', 'date'),
    ('T', 'linkedDocument', 'Linked Document', 'text'), ('U', 'approvalDate', '承認日', 'date'), ('V', 'documentName', 'ドキュメント名', 'document_name'),
)
def strip_document_version(name):
    # Some forms prefill the cell with a "名前" label line, which staff remove when transcribing.
    name = re.sub(r'^名前\s+', '', name)
    # Drop the file version suffix such as "(v0.1)"; unparenthesized "Ver.2" / "_v2" name a revised document and stay.
    return re.sub(r'\s*[(（]\s*v(?:er\.?)?\s*\d+(?:\.\d+)*\s*[)）]$', '', name, flags=re.I).strip() or name


def sheet_kubun1(value, options):
    # The dropdown uses full-width brackets ("MKT（CSCC）") and the approver lookup needs an exact option.
    normalized = re.sub(r'\s+', '', value).replace('(', '（').replace(')', '）')
    return next((option for option in options if option.lower() == normalized.lower()), normalized)


def column_index(letter):
    index = 0
    for char in letter:
        index = index * 26 + ord(char) - 64
    return index - 1


def transcript_rows(entry, existing_event_ids=(), kubun1_options=()):
    received_at = datetime.fromisoformat(entry['received_at']).astimezone(JST)
    received = f'{received_at.year}/{received_at.month}/{received_at.day}'
    rows = []
    for row in entry['rows']:
        # Corrections often keep the 新規依頼 subject or start a new thread; an already listed 講演会ID is the reliable signal.
        # Placeholders such as "n/a" or "なし" are shared by unrelated rows, so only real IDs (with digits) count.
        event_id = str(row.get('eventId') or '').strip()
        if re.search(r'\d', event_id) and event_id in existing_event_ids:
            row['category'] = '修正'
        values, changes, warnings = [], [], []
        for column, key, label, kind in TRANSCRIPT_COLUMNS:
            original = str(row.get(key) or '').strip() if key else ''
            if key == 'kubun1':
                original = str(row.get('kubun1Original') or original).strip()
            if kind == 'url':
                original = original or str(row.get('viewingUrlOriginal') or '').strip()
            if kind == 'time':
                original = str(row.get('eventTimeRaw') or original).strip()
            if kind == 'date':
                values.append(sheet_date(original))
                if key == 'approvalDate' and not original:
                    warnings.append({'column': column, 'label': label, 'message': '申請書に承認日がありません。Promomatsで確認して入力してください'})
                continue
            # Multiple time slots are joined with newlines on purpose.
            value = original if kind in ('url', 'time') else one_line(original)
            if kind == 'kubun1':
                value = sheet_kubun1(value, kubun1_options)
                if kubun1_options and value not in kubun1_options:
                    warnings.append({'column': column, 'label': label, 'message': f'「{value}」はシートの選択肢にありません（{" / ".join(kubun1_options)}）。選び直してください'})
            if kind == 'event_name':
                value = strip_event_date(value)
            if kind == 'document_name':
                value = strip_document_version(value)
            raw_document = one_line(str(row.get('linkedDocumentRaw') or ''))
            if key == 'linkedDocument' and value.startswith('JP.') and raw_document and 'JP.' not in raw_document.upper():
                original = raw_document
            if value != original:
                changes.append({'column': column, 'label': label, 'before': original, 'after': value})
            values.append(sheet_text(value))
        row['changes'], row['warnings'] = changes, warnings
        rows.append((
            [received, row['category'], '申請内容確認中'],
            values,
            [sheet_text(entry.get('requester')), sheet_text(entry.get('subject'))],
        ))
    return rows


def add_workdays(day, count, holidays):
    while count:
        day += timedelta(days=1)
        if day.weekday() < 5 and day not in holidays:
            count -= 1
    return day


def release_date(session, base, title, row, received_at):
    """本番化予定日: the sheet's E column (WORKDAY(受領日, 2, 祝日)); computed the same way if it shows an error."""
    try:
        response = session.get(f"{base}/values:batchGet", params={'ranges': [f'{title}!E{row}', "'祝日'!A:A"]}, timeout=30)
        response.raise_for_status()
        cell, holiday_column = [block.get('values', []) for block in response.json()['valueRanges']]
        value = str(cell[0][0]).strip() if cell and cell[0] else ''
        if re.fullmatch(r'\d{4}/\d{1,2}/\d{1,2}', value):
            return value
        holidays = set()
        for item in holiday_column:
            match = re.fullmatch(r'(\d{4})/(\d{1,2})/(\d{1,2})', str(item[0]).strip()) if item else None
            if match:
                holidays.add(datetime(int(match[1]), int(match[2]), int(match[3])).date())
        day = add_workdays(datetime.fromisoformat(received_at).astimezone(JST).date(), 2, holidays)
        return f'{day.year}/{day.month}/{day.day}'
    except Exception:
        # The notice is still useful without this line; never fail a finished transcript over it.
        return ''


def append_transcript(credentials_factory, entry):
    sheet_id = os.getenv('REGENERON_TRANSCRIPT_SHEET_ID', TRANSCRIPT_SHEET_ID)
    credentials = credentials_factory(['https://www.googleapis.com/auth/spreadsheets'])
    with AuthorizedSession(credentials) as session:
        base = f'https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}'
        properties = first_sheet_properties(session, base)
        title = "'" + properties['title'].replace("'", "''") + "'"
        response = session.get(f"{base}/values:batchGet", params={'ranges': [f'{title}!B:B', f'{title}!{EVENT_ID_COLUMN}:{EVENT_ID_COLUMN}'], 'majorDimension': 'COLUMNS'}, timeout=30)
        response.raise_for_status()
        received_column, event_column = [((block.get('values') or [[]])[0])[TRANSCRIPT_FIRST_ROW - 1:] for block in response.json()['valueRanges']]
        filled = [index for index, value in enumerate(received_column, start=TRANSCRIPT_FIRST_ROW) if str(value).strip()]
        start = filled[-1] + 1 if filled else TRANSCRIPT_FIRST_ROW
        response = session.get(base, params={'ranges': f'{title}!{KUBUN1_COLUMN}{start}', 'includeGridData': 'true',
                                             'fields': 'sheets(data(rowData(values(dataValidation))))'}, timeout=30)
        response.raise_for_status()
        cells = [value for sheet in response.json().get('sheets', []) for block in sheet.get('data', []) for item in block.get('rowData', []) for value in item.get('values', [])]
        condition = (cells[0].get('dataValidation') or {}).get('condition', {}) if cells else {}
        kubun1_options = [item.get('userEnteredValue', '') for item in condition.get('values', [])] if condition.get('type') == 'ONE_OF_LIST' else []
        rows = transcript_rows(entry, {str(value).strip() for value in event_column if str(value).strip()}, kubun1_options)
        end = start + len(rows) - 1
        if end > properties['gridProperties']['rowCount']:
            raise ValueError('転記先シートに空き行がありません。関数入りの行を追加してください。')
        # Format before values: if the value write fails, the retry targets the same empty rows again.
        red = [{'repeatCell': {
            'range': {'sheetId': properties['sheetId'], 'startRowIndex': start + offset - 1, 'endRowIndex': start + offset,
                      'startColumnIndex': column_index(change['column']), 'endColumnIndex': column_index(change['column']) + 1},
            'cell': {'userEnteredFormat': {'textFormat': {'foregroundColorStyle': {'rgbColor': {'red': 1}}}}},
            'fields': 'userEnteredFormat.textFormat.foregroundColorStyle',
        }} for offset, row in enumerate(entry['rows']) for change in row['changes'] + row['warnings']]
        if red:
            session.post(f'{base}:batchUpdate', json={'requests': red}, timeout=40).raise_for_status()
        data = [{'range': f'{title}!{first}{start}:{last}{end}', 'values': [row[index] for row in rows]}
                for index, (first, last) in enumerate(TRANSCRIPT_SEGMENTS)]
        response = session.post(f'{base}/values:batchUpdate', json={'valueInputOption': 'USER_ENTERED', 'data': data}, timeout=40)
        response.raise_for_status()
        return {'row': start, 'url': f'https://docs.google.com/spreadsheets/d/{sheet_id}/edit#gid=0&range=A{start}',
                'release_date': release_date(session, base, title, start, entry['received_at'])}
