"""Общий импорт решений UI/CLI; формат первых колонок адаптирован из PR #4 Ивана."""

from copy import deepcopy
from collections import defaultdict
from datetime import datetime, timezone
import hashlib
import io
import json
from pathlib import Path
import zipfile

from openpyxl import Workbook, load_workbook

from .export import add_table
from .letters import valid_inn
from .normalization import normalize


DEFAULT_STORE = Path(__file__).resolve().parents[2] / 'manual_decisions.json'
MAX_DECISION_BYTES = 6 * 1024 * 1024
COLUMNS = [
    ('group_id', 'Группа'), ('source_name', 'Исходное название'), ('region', 'Регион'),
    ('city', 'Город'), ('inn', 'ИНН (выбор оператора)'), ('comment', 'Комментарий'),
    ('candidates', 'Кандидаты (ИНН | название | адрес | балл | источник)'),
    ('reason', 'Причина проверки'), ('result_id', 'Идентификатор результата'),
]


class ManualReviewError(ValueError):
    pass


def _hash(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def manual_key(record):
    # Год, вид образования и риски не позволяют переносить решение на другой контекст.
    return _hash([normalize(record.get(key, '')) for key in
                  ('name', 'region', 'city', 'education_kind', 'year', 'inn_raw',
                   'school_address_raw', 'postal_code_raw')] + [sorted(record.get('flags', []))])


def result_id(result):
    return _hash([result.get('created_at_utc'),
                  [(r['record_id'], r.get('name'), r.get('region'), r.get('city')) for r in result['records']],
                  [(s.get('sha256'), s.get('selected_sheet')) for s in result.get('sources', [])]])


def load_result(path):
    try:
        result = json.loads(Path(path).read_text(encoding='utf-8'))
        if (not isinstance(result, dict) or not isinstance(result.get('records'), list)
                or not isinstance(result.get('summary'), dict)
                or not isinstance(result.get('resolution'), dict)
                or not isinstance(result['resolution'].get('groups'), list)
                or not isinstance(result['resolution'].get('candidates'), list)):
            raise ValueError()
        ids = {r['record_id'] for r in result['records']}
        if any(not isinstance(c, dict) for c in result['resolution']['candidates']):
            raise ValueError()
        if any(not group['record_ids'] or any(rid not in ids for rid in group['record_ids'])
               for group in result['resolution']['groups']):
            raise ValueError()
        return result
    except (OSError, ValueError, KeyError, TypeError):
        raise ManualReviewError('Не удалось прочитать result.json указанного запуска. Выберите исходный файл результата, созданный программой.') from None


def _valid_candidate(candidate, inn):
    return (isinstance(candidate, dict) and isinstance(inn, str) and valid_inn(inn)
            and set(inn) != {'0'} and candidate.get('candidate_inn') == inn
            and bool(candidate.get('candidate_name'))
            and candidate.get('legal_status') in {'ACTIVE', 'Действующее'})


def load_store(path):
    path = Path(path)
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
        entries = data['decisions']
        if data['version'] != 1 or not isinstance(entries, dict):
            raise ValueError()
        for key, item in entries.items():
            if (not isinstance(key, str) or not isinstance(item, dict)
                    or not _valid_candidate(item.get('candidate'), item.get('inn'))
                    or not isinstance(item.get('comment'), str) or not isinstance(item.get('reviewed_at'), str)):
                raise ValueError()
        return entries
    except (OSError, ValueError, KeyError, TypeError):
        raise ManualReviewError('Не удалось прочитать справочник ручных решений. Проверьте файл --decisions-store; он не перезаписан.') from None


def save_store(path, entries):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + '.tmp')
    try:
        temporary.write_text(json.dumps({'version': 1, 'decisions': entries}, ensure_ascii=False, indent=2), encoding='utf-8')
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def confirmed_for(records, entries):
    if not entries:
        return None
    if any(not r.get('region') or not r.get('city') or any(
            'КОНФЛИКТ' in flag or flag in {'ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ', 'СТАРЫЙ_ВЫПУСК'}
            for flag in r.get('flags', [])) for r in records):
        # В изолированных строках одинаковый текст не доказывает тождество учреждений.
        return None
    decisions = [entries.get(manual_key(record)) for record in records]
    if not decisions or not all(decisions):
        return None
    first = decisions[0]
    if not _valid_candidate(first.get('candidate'), first.get('inn')):
        return None
    if any(d['inn'] != first['inn'] or d['candidate'] != first['candidate'] for d in decisions):
        return None
    return deepcopy(first)


def review_rows(result):
    records = {r['record_id']: r for r in result['records']}
    candidates = defaultdict(list)
    for candidate in result['resolution'].get('candidates', []) + result.get('dadata_candidates', []):
        candidates[candidate['group_id']].append(candidate)
    identity = result_id(result)
    rows = []
    for group in result['resolution']['groups']:
        if group.get('selected'):
            continue
        member = records[group['record_ids'][0]]
        options = sorted(candidates[group['group_id']],
                         key=lambda c: (bool(c.get('eligible')), c.get('score', 0)), reverse=True)
        distinct = {}
        for option in options:
            if option.get('candidate_inn'):
                distinct.setdefault(option['candidate_inn'], option)
        top = list(distinct.values())[:3]
        reasons = sorted({reason for c in top for reason in c.get('reasons', [])})
        rows.append({'group_id': group['group_id'], 'source_name': member.get('name', ''),
                     'region': member.get('region', ''), 'city': member.get('city', ''),
                     'status': group['status'], 'options': top,
                     'reason': '\n'.join(reasons) or ('Нужен выбор оператора' if top else 'Кандидаты не найдены'),
                     'candidates': '\n'.join(f"{c['candidate_inn']} | {c['candidate_name']} | {c.get('address', '')} | {c.get('score', '')} | {c.get('source', '')}" for c in top),
                     'result_id': identity})
    return rows


def template_bytes(result):
    book = Workbook()
    book.remove(book.active)
    add_table(book, 'Решения', review_rows(result), COLUMNS)
    stream = io.BytesIO()
    book.save(stream)
    return stream.getvalue()


def _read_rows(content):
    if not content or len(content) > MAX_DECISION_BYTES:
        raise ManualReviewError('Файл решений должен быть .xlsx размером до 6 МБ.')
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            if sum(i.file_size for i in archive.infolist()) > 32 * 1024 * 1024:
                raise ManualReviewError('Файл решений слишком большой после распаковки.')
        book = load_workbook(io.BytesIO(content), read_only=True, data_only=False)
        try:
            sheet = book['Решения']
            if sheet.max_row > 5001 or sheet.max_column != len(COLUMNS):
                raise ManualReviewError('Нужен неизменённый шаблон решений, не более 5 000 строк.')
            rows = list(sheet.iter_rows(max_col=len(COLUMNS)))
            if not rows or [c.value for c in rows[0]] != [label for _, label in COLUMNS]:
                raise ManualReviewError('Заголовки файла решений изменены. Скачайте новый шаблон.')
            return [[(c.value, c.data_type) for c in row] for row in rows[1:]]
        finally:
            book.close()
    except ManualReviewError:
        raise
    except Exception:
        raise ManualReviewError('Не удалось прочитать .xlsx с листом «Решения». Скачайте шаблон из результата обработки.') from None


def apply_import(result, content, entries):
    """Сначала проверить весь файл, затем вернуть новые данные без изменения аргументов."""
    available = {row['group_id']: row for row in review_rows(result)}
    staged, seen = [], set()
    for number, cells in enumerate(_read_rows(content), 2):
        row = [cell[0] for cell in cells]
        if row[4] is None or not str(row[4]).strip():
            continue
        if any(kind == 'f' for _, kind in cells):
            raise ManualReviewError(f'Строка {number}: формулы в файле решений не допускаются.')
        group_id = str(row[0] or '')
        current = available.get(group_id)
        if group_id in seen or current is None:
            raise ManualReviewError(f'Строка {number}: неизвестная, повторная или уже подтверждённая группа.')
        seen.add(group_id)
        if (row[8] != current['result_id'] or any((row[i] or '') != current[key]
                for i, key in ((1, 'source_name'), (2, 'region'), (3, 'city'), (6, 'candidates'), (7, 'reason')))):
            raise ManualReviewError(f'Строка {number}: файл относится к другому запуску или изменены исходные данные.')
        inn = str(row[4]).strip()
        options = [c for c in current['options'] if c.get('candidate_inn') == inn]
        if not options or not _valid_candidate(options[0], inn):
            raise ManualReviewError(f'Строка {number}: нужен корректный ИНН действующего кандидата из списка этой группы.')
        comment = str(row[5] or '').strip()
        if len(comment) > 1000:
            raise ManualReviewError(f'Строка {number}: комментарий длиннее 1 000 символов.')
        staged.append((group_id, options[0], comment))
    if not staged:
        raise ManualReviewError('В файле нет заполненных решений.')
    updated, saved = deepcopy(result), deepcopy(entries)
    groups = {g['group_id']: g for g in updated['resolution']['groups']}
    records = {r['record_id']: r for r in updated['records']}
    now = datetime.now(timezone.utc).isoformat()
    for group_id, candidate, comment in staged:
        group = groups[group_id]
        decision = {'inn': candidate['candidate_inn'], 'candidate': deepcopy(candidate),
                    'comment': comment, 'reviewed_at': now}
        # Группа нового запуска может иметь другой ID: храним только данные организации.
        decision['candidate'].pop('group_id', None)
        for rid in group['record_ids']:
            saved[manual_key(records[rid])] = deepcopy(decision)
        group.update(selected=deepcopy(candidate), status='ПРОВЕРЕНО', decision_reason='ПРОВЕРЕНО',
                     review_candidate=None, manual_decision=True,
                     manual_comment=comment, manual_reviewed_at=now)
        if candidate not in updated['resolution']['candidates']:
            updated['resolution']['candidates'].append(deepcopy(candidate))
    updated['manual_import'] = {'imported': len(staged), 'reviewed_at': now}
    return updated, saved
