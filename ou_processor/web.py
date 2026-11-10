"""Локальный интерфейс: python -m ou_processor.web. Сервер слушает только этот компьютер."""
import argparse
import base64
import binascii
from datetime import datetime, timezone
import getpass
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
import os
from pathlib import Path
import re
import secrets
import shutil
import sqlite3
import threading
import webbrowser
import zipfile

from .dadata import DaDataClient, candidate_row
from openpyxl import Workbook, load_workbook
from .key_storage import TokenStore
from .full_run import aggregate, save
from .input import read_records
from .pipeline import process
from .registry import current_snapshot
from .resolution import RegistrySearch, SearchCache, resolve
from .registry_worker import RegistryWorker

ROOT = Path(__file__).resolve().parents[1]
MAX_BYTES = 100 * 1024 * 1024
MAX_BODY = 140 * 1024 * 1024
MAX_UNPACKED_BYTES = 512 * 1024 * 1024


class UserError(Exception):
    pass


class JobCancelled(Exception):
    pass


def unpack_uploads(files):
    if not isinstance(files, list) or not 1 <= len(files) <= 10:
        raise UserError('Выберите от 1 до 10 файлов Excel (.xlsx).')
    result, seen, total = [], set(), 0
    for item in files:
        if not isinstance(item, dict):
            raise UserError('Не удалось прочитать список файлов.')
        name = str(item.get('name', '')).replace('\\', '/').rsplit('/', 1)[-1]
        if not name.lower().endswith('.xlsx') or name.startswith('~$') or len(name) > 180:
            raise UserError('Поддерживаются книги .xlsx. Закройте временные файлы Excel.')
        try:
            content = base64.b64decode(item.get('data', ''), validate=True)
        except (ValueError, TypeError, binascii.Error):
            raise UserError('Файл повреждён при загрузке.') from None
        total += len(content)
        if total > MAX_BYTES:
            raise UserError('Суммарный размер файлов не должен превышать 100 МБ.')
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                if '[Content_Types].xml' not in archive.namelist() or 'xl/workbook.xml' not in archive.namelist():
                    raise ValueError()
                if sum(info.file_size for info in archive.infolist()) > MAX_UNPACKED_BYTES:
                    raise UserError('Книга слишком большая после распаковки. Разделите её на части.')
        except (zipfile.BadZipFile, ValueError):
            raise UserError(f'«{name}» не является корректной книгой .xlsx или защищена паролем.') from None
        digest = hashlib.sha256(content).hexdigest()
        if digest not in seen:
            result.append((name, content))
            seen.add(digest)
    return result


class Application:
    def __init__(self, storage, registry, cache, token='', client_factory=DaDataClient,
                 token_store=None, registry_worker=None):
        self.storage, self.registry, self.cache = Path(storage), Path(registry), Path(cache)
        self.decisions_path = self.storage.parent / 'manual_decisions.json'
        self.initial_token = token
        self.token, self.client_factory = token, client_factory
        self.token_store = token_store
        self.token_saved = False
        if token_store:
            try:
                saved_token = token_store.load()
                self.token_saved = bool(saved_token)
                if not self.token:
                    self.token = saved_token
            except Exception:
                # Недоступное системное хранилище не мешает обработке с введённым ключом.
                pass
        self.registry_worker = registry_worker or RegistryWorker(self.registry)
        self.csrf = secrets.token_urlsafe(32)
        self.jobs = {}
        self.cancel_flags = {}
        self.lock = threading.Lock()
        self.active = None

    def settings(self):
        try:
            snapshot = current_snapshot(self.registry)
        except (ValueError, OSError, KeyError):
            snapshot = None
        date = snapshot[1].get('snapshot_date', '') if snapshot else ''
        return {'csrf': self.csrf, 'token_configured': bool(self.token), 'token_saved': self.token_saved,
                'registry_ready': bool(snapshot), 'registry_date': date,
                'registry_update': self.registry_worker.snapshot(), 'max_files': 10, 'max_mb': 100}

    def start_registry_update(self):
        return self.registry_worker.start()

    def forget_token(self):
        if self.token_store:
            try:
                self.token_store.delete()
            except Exception:
                raise UserError('Не удалось удалить ключ из системного хранилища.') from None
        self.token = self.initial_token
        self.token_saved = False

    def start(self, payload):
        if not isinstance(payload, dict):
            raise UserError('Некорректные параметры загрузки.')
        uploaded = unpack_uploads(payload.get('files'))
        token = payload.get('token') or self.token
        if not isinstance(token, str) or len(token) > 256 or '\n' in token or '\r' in token:
            raise UserError('Проверьте API-ключ DaData.')
        token = token.strip()
        offline = payload.get('offline') is not False
        if not token and not offline:
            raise UserError('Для отдельной проверки в DaData введите API-ключ.')
        with self.lock:
            if self.active:
                raise UserError('Уже идёт обработка. Дождитесь её завершения.')
            if payload.get('save_token') is True and payload.get('token'):
                if not self.token_store:
                    raise UserError('Системное хранилище ключей недоступно. Снимите флажок сохранения ключа.')
                try:
                    self.token_store.save(token)
                except Exception:
                    raise UserError('Не удалось сохранить ключ. Проверьте доступ к системному хранилищу или снимите флажок сохранения.') from None
                self.token_saved = True
            self.token = token
            job_id = secrets.token_hex(16)
            self.jobs[job_id] = {'id': job_id, 'status': 'running', 'phase': 'Читаем Excel',
                                 'progress': 2, 'done': 0, 'total': 0, 'error': ''}
            self.cancel_flags[job_id] = threading.Event()
            self.active = job_id
        thread = threading.Thread(target=self.run, args=(job_id, uploaded, token, offline), daemon=True)
        thread.start()
        return job_id

    def change(self, job_id, **fields):
        with self.lock:
            if self.jobs[job_id]['status'] == 'cancelling' and 'status' not in fields:
                return
            self.jobs[job_id].update(fields)

    def snapshot(self, job_id):
        with self.lock:
            job = self.jobs.get(job_id)
            if job:
                snapshot = dict(job)
                if snapshot.get('status') == 'done' and not (self.storage / job_id / 'result' / 'result.json').is_file():
                    snapshot.update(status='error', phase='Результат недоступен', error='Файлы этой обработки отсутствуют. Загрузите исходный Excel и запустите новую обработку.')
                return snapshot
        # Finished runs remain on disk after a server restart.
        path = self.storage / job_id / 'result' / 'result.json'
        if path.is_file():
            try:
                result = json.loads(path.read_text(encoding='utf-8'))
                return {'id': job_id, 'status': 'done', 'phase': 'Готово',
                        'progress': 100, 'done': 0, 'total': 0,
                        'review': result.get('review', [])[:200],
                        'selection_review': [{'name': r.get('name', ''),
                            'reason': r.get('selection_reason', ''),
                            'education': r.get('education_kind', ''),
                            'record_id': r.get('record_id', '')}
                            for r in result.get('records', [])
                            if r.get('selection') == 'ПРОВЕРИТЬ_ОТБОР'][:200],
                        'download': f'/api/jobs/{job_id}/download'}
            except (OSError, ValueError):
                pass
        return None

    def cancel(self, job_id):
        with self.lock:
            job = self.jobs.get(job_id)
            if not job:
                return None
            if job['status'] == 'running':
                self.cancel_flags.setdefault(job_id, threading.Event()).set()
                job.update(status='cancelling', phase='Отменяем обработку')
            return dict(job)

    def cancellation_requested(self, job_id):
        flag = self.cancel_flags.get(job_id)
        return bool(flag and flag.is_set())

    def check_cancel(self, job_id):
        if self.cancellation_requested(job_id):
            raise JobCancelled()

    def progress(self, job_id, done, total):
        self.check_cancel(job_id)
        self.change(job_id, done=done, total=total, progress=8 + round(82 * done / max(1, total)))

    def finish(self, job_id, **fields):
        with self.lock:
            if self.cancellation_requested(job_id):
                raise JobCancelled()
            self.jobs[job_id].update(status='done', phase='Готово', progress=100, **fields)

    @staticmethod
    def manual_key(record):
        from .normalization import normalize
        return '|'.join((normalize(record.get('name', '')), normalize(record.get('region', '')),
                         normalize(record.get('city', ''))))

    def load_manual_decisions(self):
        try:
            data = json.loads(self.decisions_path.read_text(encoding='utf-8'))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def apply_manual_decisions(self, result):
        decisions = self.load_manual_decisions()
        records = {r['record_id']: r for r in result.get('records', [])}
        by_group = {g['group_id']: g for g in result.get('resolution', {}).get('groups', [])}
        candidates = result.get('resolution', {}).get('candidates', [])
        for group_id, group in by_group.items():
            members = [records[rid] for rid in group.get('record_ids', []) if rid in records]
            if not members:
                continue
            key = self.manual_key(members[0])
            decision = decisions.get(key)
            if not decision:
                continue
            matches = [c for c in candidates if c.get('group_id') == group_id and
                       str(c.get('candidate_inn', '')) == str(decision.get('inn', ''))]
            saved_candidate = decision.get('candidate')
            selected = dict(saved_candidate) if isinstance(saved_candidate, dict) else (dict(matches[0]) if len(matches) == 1 else None)
            if selected:
                selected['group_id'] = group_id
                group['selected'] = selected
                group['status'] = 'ПРОВЕРЕНО'
                group['manual_comment'] = decision.get('comment', '')
                group['manual_reviewed_at'] = decision.get('reviewed_at', '')
                group['manual_decision'] = True

    @staticmethod
    def valid_inn(value):
        inn = str(value or '').strip()
        if not re.fullmatch(r'\d{10}|\d{12}', inn):
            return False
        digits = list(map(int, inn))
        def check(weights):
            return sum(a * b for a, b in zip(digits, weights)) % 11 % 10
        if len(digits) == 10:
            return check([2, 4, 10, 3, 5, 9, 4, 6, 8]) == digits[9]
        return (check([7, 2, 4, 10, 3, 5, 9, 4, 6, 8]) == digits[10]
                and check([3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]) == digits[11])

    def decisions_template(self, job_id):
        path = self.storage / job_id / 'result' / 'result.json'
        if not path.is_file():
            raise UserError('Файлы этой обработки отсутствуют. Загрузите Excel и запустите новую обработку на текущем сервере.')
        try:
            result = json.loads(path.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            raise UserError('Результат повреждён. Запустите обработку заново.') from None
        book = Workbook()
        sheet = book.active
        sheet.title = 'Решения'
        sheet.append(['Группа', 'Исходное название', 'Регион', 'Город', 'ИНН (выбор оператора)',
                      'Комментарий', 'Кандидаты (ИНН | название | адрес | балл | источник)', 'Причина проверки'])
        for row in result.get('review', []):
            choices = '\n'.join(f"{c.get('candidate_inn', '')} | {c.get('candidate_name', '')} | {c.get('address', '')} | {c.get('score', '')} | {c.get('source', '')}"
                                for c in row.get('candidates', [])[:3])
            sheet.append([row['group_id'], row.get('name',''), row.get('region',''), row.get('city',''),
                          '', '', choices, row.get('reason','')])
        for col, width in {'A':28,'B':55,'C':25,'D':25,'E':26,'F':45,'G':105,'H':55}.items():
            sheet.column_dimensions[col].width = width
        for cells in sheet.iter_rows(min_row=2):
            cells[4].number_format = '@'
        stream = io.BytesIO()
        book.save(stream)
        return stream.getvalue()

    def import_manual_decisions(self, job_id, payload):
        if not isinstance(payload, dict) or not isinstance(payload.get('data'), str):
            raise UserError('Загрузите заполненный файл решений .xlsx.')
        if len(payload['data']) > 8_000_000:
            raise UserError('Файл решений слишком большой.')
        try:
            content = base64.b64decode(payload['data'], validate=True)
            book = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
            sheet = book['Решения']
            rows = list(sheet.values)
            book.close()
        except Exception:
            raise UserError('Нужен файл .xlsx с листом «Решения».') from None
        expected = ('Группа', 'Исходное название', 'Регион', 'Город', 'ИНН (выбор оператора)', 'Комментарий')
        if not rows or tuple(rows[0][:6]) != expected:
            raise UserError('Заголовки файла решений изменены. Скачайте новый шаблон.')
        if len(rows) > 5001:
            raise UserError('В файле решений слишком много строк.')
        path = self.storage / job_id / 'result' / 'result.json'
        try:
            result = json.loads(path.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            raise UserError('Результат обработки не найден.') from None
        groups = {g['group_id']:g for g in result['resolution']['groups']}
        records = {r['record_id']:r for r in result['records']}
        candidates = result['resolution']['candidates']
        decisions = self.load_manual_decisions()
        staged = []
        seen = set()
        errors = []
        for number, row in enumerate(rows[1:], 2):
            if not row or not str(row[4] or '').strip():
                continue
            group_id = str(row[0] or '').strip()
            group = groups.get(group_id)
            if group_id in seen or not group:
                errors.append(f'Строка {number}: неизвестная или повторная группа')
                continue
            seen.add(group_id)
            member = next((records[rid] for rid in group['record_ids'] if rid in records), None)
            if not member or any(str(row[i] or '').strip() != str(member.get(field,'')).strip()
                                 for i,field in ((1,'name'),(2,'region'),(3,'city'))):
                errors.append(f'Строка {number}: изменены исходные данные')
                continue
            inn = str(row[4]).strip()
            if not self.valid_inn(inn):
                errors.append(f'Строка {number}: некорректный ИНН')
                continue
            options = [c for c in candidates if c['group_id'] == group_id and str(c.get('candidate_inn','')) == inn]
            if not options:
                errors.append(f'Строка {number}: ИНН отсутствует среди кандидатов группы')
                continue
            comment = str(row[5] or '').strip()
            if len(comment) > 1000:
                errors.append(f'Строка {number}: комментарий длиннее 1000 символов')
                continue
            staged.append((group, member, inn, comment, options[0]))
        if errors:
            raise UserError('Файл не импортирован: ' + '; '.join(errors[:12]))
        if not staged:
            raise UserError('В файле нет заполненных решений.')
        for group, member, inn, comment, candidate in staged:
            decisions[self.manual_key(member)] = {'inn':inn, 'comment':comment,
                'reviewed_at':datetime.now(timezone.utc).isoformat(), 'name':member.get('name',''),
                'region':member.get('region',''), 'city':member.get('city',''), 'candidate':dict(candidate)}
            group.update(selected=dict(candidate), status='ПРОВЕРЕНО', manual_comment=comment,
                         manual_decision=True)
        self.write_manual_decisions(decisions)
        aggregate(result)
        save(result, self.storage / job_id / 'result')
        self.change(job_id, review=result['review'][:200], preview=result['schools'][:20])
        return {'ok':True, 'imported':len(staged), 'review':result['review'][:200]}

    def write_manual_decisions(self, decisions):
        self.decisions_path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.decisions_path.with_suffix('.tmp')
        temp.write_text(json.dumps(decisions, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temp, self.decisions_path)

    def check_dadata(self, job_id, payload):
        if not isinstance(payload, dict):
            raise UserError('Некорректный запрос DaData.')
        token = str(payload.get('token', '')).strip()
        if not token or len(token) > 256 or '\\n' in token or '\\r' in token:
            raise UserError('Введите API-ключ для отдельной проверки в DaData.')
        result_path = self.storage / job_id / 'result' / 'result.json'
        try:
            result = json.loads(result_path.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            raise UserError('Сначала выполните обычную обработку.') from None
        group_id = str(payload.get('group_id', ''))
        groups = {g['group_id']: g for g in result.get('resolution', {}).get('groups', [])}
        group = groups.get(group_id)
        if not group:
            raise UserError('Группа для проверки не найдена.')
        records = {r['record_id']: r for r in result.get('records', [])}
        member = next((records[rid] for rid in group.get('record_ids', []) if rid in records), None)
        if not member:
            raise UserError('Не удалось определить исходное название.')
        query = ' '.join(str(x).strip() for x in (member.get('name'), member.get('region'), member.get('city')) if x)
        try:
            found = DaDataClient(token).suggest(query)
        except Exception:
            raise UserError('DaData не ответила. Проверьте ключ и попробуйте позже.') from None
        sample = {'group_id': group_id, 'query': query}
        converted = [{**candidate_row(sample, item, rank, group), 'source':'DaData', 'score':0, 'eligible':False, 'reasons':candidate_row(sample, item, rank, group).get('flags', [])} for rank, item in enumerate(found, 1)]
        all_candidates = result.setdefault('resolution', {}).setdefault('candidates', [])
        all_candidates[:] = [c for c in all_candidates if not (c.get('group_id') == group_id and c.get('source') == 'DaData')]
        all_candidates.extend(converted)
        selected_inn = str((group.get('selected') or {}).get('candidate_inn', ''))
        found_inns = sorted({str(c.get('candidate_inn', '')) for c in converted if c.get('candidate_inn')})
        discrepancies = []
        if selected_inn and selected_inn not in found_inns:
            discrepancies.append('Выбранный ИНН отсутствует в ответе DaData; решение не изменено.')
        if len(found_inns) > 1:
            discrepancies.append('DaData вернула разные ИНН; автоматическое подтверждение запрещено.')
        result.setdefault('manual_dadata_checks', []).append({
            'group_id':group_id, 'checked_at':datetime.now(timezone.utc).isoformat(),
            'inns':found_inns, 'discrepancies':discrepancies})
        aggregate(result)
        save(result, self.storage / job_id / 'result')
        self.change(job_id, review=result.get('review', [])[:200])
        return {'count': len(converted), 'review': result.get('review', [])[:200],
                'inns':found_inns, 'discrepancies':discrepancies}

    def save_manual_decision(self, job_id, payload):
        if not isinstance(payload, dict):
            raise UserError('Некорректное ручное решение.')
        job = self.snapshot(job_id)
        if not job or job.get('status') != 'done':
            raise UserError('Сначала дождитесь завершения обработки.')
        result_path = self.storage / job_id / 'result' / 'result.json'
        try:
            result = json.loads(result_path.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            raise UserError('Не удалось открыть результат обработки.') from None
        group_id = str(payload.get('group_id', ''))
        inn = str(payload.get('inn', '')).strip()
        comment = str(payload.get('comment', '')).strip()[:1000]
        groups = [g for g in result.get('resolution', {}).get('groups', []) if g.get('group_id') == group_id]
        if len(groups) != 1 or not self.valid_inn(inn):
            raise UserError('Выберите кандидата из списка.')
        group = groups[0]
        members = {r['record_id']: r for r in result.get('records', [])}
        member = next((members[rid] for rid in group.get('record_ids', []) if rid in members), None)
        if not member:
            raise UserError('Не удалось определить исходную запись.')
        options = [c for c in result.get('resolution', {}).get('candidates', [])
                   if c.get('group_id') == group_id and str(c.get('candidate_inn', '')) == inn]
        if len(options) != 1:
            raise UserError('ИНН должен точно соответствовать одному из показанных кандидатов. Конфликт ИНН требует дополнительной проверки.')
        decisions = self.load_manual_decisions()
        key = self.manual_key(member)
        decisions[key] = {'inn': inn, 'comment': comment, 'reviewed_at': datetime.now(timezone.utc).isoformat(),
                          'name': member.get('name', ''), 'region': member.get('region', ''),
                          'city': member.get('city', ''), 'candidate': dict(options[0])}
        self.write_manual_decisions(decisions)
        # Update current result and regenerate the workbook for immediate download.
        group['selected'] = dict(options[0])
        group['status'] = 'ПРОВЕРЕНО'
        group['manual_comment'] = comment
        group['manual_reviewed_at'] = decisions[key]['reviewed_at']
        group['manual_decision'] = True
        aggregate(result)
        save(result, self.storage / job_id / 'result')
        self.change(job_id, review=result['review'][:200], preview=result['schools'][:20])
        return {'ok': True, 'status': 'ПРОВЕРЕНО', 'inn': inn, 'comment': comment}

    def run(self, job_id, uploaded, token, offline):
        reserve = None
        was_cancelled = False
        directory = self.storage / job_id
        try:
            self.check_cancel(job_id)
            directory.mkdir(parents=True, mode=0o700)
            paths = []
            for index, (name, content) in enumerate(uploaded):
                self.check_cancel(job_id)
                folder = directory / 'inputs' / str(index)
                folder.mkdir(parents=True)
                path = folder / name
                path.write_bytes(content)
                paths.append(path)
            config = json.loads((ROOT / 'config.json').read_text())
            records, sources = read_records(paths, should_cancel=lambda: self.check_cancel(job_id))
            self.check_cancel(job_id)
            if len(records) > 5000:
                raise UserError('В прототипе можно обработать до 5 000 строк за запуск. Разделите таблицу.')
            result = process(records, config)
            self.check_cancel(job_id)
            result.update(sources=sources, created_at_utc=datetime.now(timezone.utc).isoformat())
            self.change(job_id, phase='Ищем организации и адреса', progress=8, total=len(result['groups']))
            api = SearchCache(self.cache, None if offline else self.client_factory(token))
            try:
                reserve = RegistrySearch(self.registry)
            except (ValueError, OSError, KeyError, sqlite3.Error):
                reserve = None
            stored_decisions = self.load_manual_decisions()
            resolve(result, api, reserve, manual_lookup=lambda record: stored_decisions.get(self.manual_key(record)), progress=False,
                on_progress=lambda done, total: self.progress(job_id, done, total),
                check_cancel=lambda: self.check_cancel(job_id))
            self.check_cancel(job_id)
            self.apply_manual_decisions(result)
            self.change(job_id, phase='Собираем итоговую таблицу', progress=93)
            aggregate(result)
            self.check_cancel(job_id)
            save(result, directory / 'result')
            self.check_cancel(job_id)
            summary = result['full_summary']
            errors = [r['status'] for r in result['resolution']['requests'] if r['status'] != 'OK']
            issues = summary['Исходных строк на проверку'] + result['summary']['selection_review_rows']
            warnings = []
            if errors:
                warnings.append('Не все запросы DaData выполнены. Проверьте лист «Запросы DaData»: ' + ', '.join(sorted(set(errors))))
            if offline:
                warnings.append('Использованы только сохранённые ответы DaData. Новых запросов не было.')
            if reserve is None:
                warnings.append('Резервный реестр недоступен. Ненайденные учреждения сохранены на листе «Проверить».')
            missing = summary['Строк без данных ученика']
            if missing:
                warnings.append(f'В {missing} исходных строках нет ФИО и/или номера личного дела.')
            self.finish(job_id,
                summary={'rows': summary['Входных строк'], 'schools': len(result['schools']),
                         'students': len(result['students']), 'review': issues,
                         'addresses': summary['Учреждений с адресом'],
                         'excluded': result['summary']['excluded_rows']},
                warnings=warnings, partial=bool(errors), preview=result['schools'][:20],
                review=result.get('review', [])[:200], selection_review=[{'name':r.get('name',''), 'reason':r.get('selection_reason',''), 'education':r.get('education_kind',''), 'record_id':r.get('record_id','')} for r in result['records'] if r.get('selection') == 'ПРОВЕРИТЬ_ОТБОР'][:200], download=f'/api/jobs/{job_id}/download')
        except JobCancelled:
            was_cancelled = True
        except (UserError, ValueError) as error:
            if self.cancellation_requested(job_id):
                was_cancelled = True
            else:
                self.change(job_id, status='error', phase='Обработка не завершена', error=str(error))
        except Exception:
            if self.cancellation_requested(job_id):
                was_cancelled = True
            else:
                # Содержимое файлов и ключ не попадают в HTTP-логи и сообщения об ошибках.
                self.change(job_id, status='error', phase='Обработка не завершена',
                            error='Не удалось обработать файл. Проверьте формат книги и доступность папки результатов.')
        finally:
            if reserve:
                reserve.connection.close()
            if was_cancelled:
                shutil.rmtree(directory, ignore_errors=True)
            with self.lock:
                if was_cancelled:
                    self.jobs[job_id].update(status='cancelled', phase='Обработка отменена')
                self.active = None
                self.cancel_flags.pop(job_id, None)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *_):
        pass

    @property
    def app(self):
        return self.server.app

    def trusted(self):
        host = self.headers.get('Host', '')
        allowed = {f'127.0.0.1:{self.server.server_port}', f'localhost:{self.server.server_port}'}
        return host in allowed

    def send(self, status, body, content_type='application/json; charset=utf-8', attachment=False, filename='schools.xlsx'):
        data = json.dumps(body, ensure_ascii=False).encode() if content_type.startswith('application/json') else body
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'")
        if attachment:
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if not self.trusted():
            return self.send(403, {'error': 'Доступ разрешён только с этого компьютера.'})
        static = {'/': ('index.html', 'text/html; charset=utf-8'),
                  '/app.js': ('app.js', 'text/javascript; charset=utf-8'),
                  '/style.css': ('style.css', 'text/css; charset=utf-8')}
        if self.path in static:
            name, content_type = static[self.path]
            return self.send(200, (ROOT / 'ou_processor' / 'static' / name).read_bytes(), content_type)
        if self.path == '/api/config':
            return self.send(200, self.app.settings())
        template_match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})/decisions-template', self.path)
        if template_match:
            try:
                data = self.app.decisions_template(template_match[1])
                return self.send(200, data,
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', attachment=True,
                    filename='manual_decisions.xlsx')
            except UserError as error:
                return self.send(404, {'error':str(error)})
            except Exception:
                return self.send(500, {'error':'Не удалось создать Excel решений. Перезапустите обработку и повторите скачивание.'})
        match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})(/download)?', self.path)
        if match:
            job = self.app.snapshot(match[1])
            if job:
                if not match[2]:
                    return self.send(200, job)
                if job['status'] != 'done':
                    return self.send(409, {'error': 'Результат ещё не готов.'})
                path = self.app.storage / match[1] / 'result' / 'schools.xlsx'
                return self.send(200, path.read_bytes(),
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', attachment=True)
        self.send(404, {'error': 'Не найдено. Возможно, сервер был перезапущен.'})

    def do_POST(self):
        if not self.trusted() or not secrets.compare_digest(self.headers.get('X-CSRF-Token', ''), self.app.csrf):
            return self.send(403, {'error': 'Обновите страницу перед загрузкой.'})
        origin = self.headers.get('Origin')
        if origin and origin != 'http://' + self.headers.get('Host', ''):
            return self.send(403, {'error': 'Недопустимый источник запроса.'})
        cancel_match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})/cancel', self.path)
        if cancel_match:
            job = self.app.cancel(cancel_match[1])
            return self.send(202 if job and job['status'] == 'cancelling' else 200 if job else 404,
                             job or {'error': 'Задача не найдена.'})
        dadata_match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})/dadata', self.path)
        if dadata_match:
            try:
                size = int(self.headers.get('Content-Length', '0'))
                if not 0 < size <= 100_000:
                    raise UserError('Некорректный размер запроса.')
                payload = json.loads(self.rfile.read(size))
                return self.send(200, self.app.check_dadata(dadata_match[1], payload))
            except (UserError, ValueError, UnicodeError) as error:
                return self.send(400, {'error': str(error) if isinstance(error, UserError) else 'Не удалось прочитать запрос.'})
        import_match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})/decisions-import', self.path)
        if import_match:
            try:
                size = int(self.headers.get('Content-Length', '0'))
                if not 0 < size <= 8_100_000:
                    raise UserError('Недопустимый размер файла решений.')
                return self.send(200, self.app.import_manual_decisions(import_match[1], json.loads(self.rfile.read(size))))
            except (UserError, ValueError) as error:
                return self.send(400, {'error':str(error) if isinstance(error, UserError) else 'Некорректный запрос.'})
        manual_match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})/decisions', self.path)
        if manual_match:
            try:
                size = int(self.headers.get('Content-Length', '0'))
                if not 0 < size <= 100_000:
                    raise UserError('Некорректный размер решения.')
                payload = json.loads(self.rfile.read(size))
                return self.send(200, self.app.save_manual_decision(manual_match[1], payload))
            except (UserError, ValueError, UnicodeError) as error:
                return self.send(400, {'error': str(error) if isinstance(error, UserError) else 'Не удалось прочитать решение.'})
        if self.path == '/api/registry/retry':
            return self.send(202, self.app.start_registry_update())
        if self.path == '/api/token/forget':
            try:
                self.app.forget_token()
                return self.send(200, {'token_configured': False, 'token_saved': False})
            except UserError as error:
                return self.send(400, {'error': str(error)})
        if self.path != '/api/jobs':
            return self.send(404, {'error': 'Не найдено.'})
        try:
            size = int(self.headers.get('Content-Length', '0'))
            if not 0 < size <= MAX_BODY:
                raise UserError('Суммарный размер файлов не должен превышать 100 МБ.')
            self.connection.settimeout(60)
            payload = json.loads(self.rfile.read(size))
            job_id = self.app.start(payload)
            self.send(202, {'id': job_id})
        except (UserError, ValueError, UnicodeError) as error:
            self.send(400, {'error': str(error) if isinstance(error, UserError) else 'Не удалось прочитать загрузку.'})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=8765)
    parser.add_argument('--storage', type=Path, default=ROOT.parent/'ui_runs')
    parser.add_argument('--registry', type=Path, default=ROOT.parent/'registry_cache')
    parser.add_argument('--cache', type=Path, default=ROOT.parent/'ui_cache'/'dadata.json')
    parser.add_argument('--ask-token', action='store_true')
    parser.add_argument('--no-browser', action='store_true')
    args = parser.parse_args()
    token = getpass.getpass('API-ключ DaData: ') if args.ask_token else os.environ.get('DADATA_API_KEY', '')
    if not args.cache.exists():
        existing = ROOT.parent/'final_schools'/'dadata_cache.json'
        if existing.exists():
            args.cache.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(existing, args.cache)
    app = Application(args.storage, args.registry, args.cache, token, token_store=TokenStore())
    try:
        server = ThreadingHTTPServer(('127.0.0.1', args.port), Handler)
    except OSError:
        parser.exit(1, 'Порт занят. Закройте старый экземпляр или укажите другой --port.\n')
    server.app = app
    app.start_registry_update()
    url = f'http://127.0.0.1:{server.server_port}'
    print(f'Откройте {url}. Для остановки нажмите Ctrl+C.', flush=True)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == '__main__':
    main()
