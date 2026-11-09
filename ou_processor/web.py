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
import threading
import webbrowser
import zipfile

from .dadata import DaDataClient
from .key_storage import TokenStore
from .full_run import aggregate, save
from .input import list_sheets, read_records
from .license_registry import auto_update as license_auto_update
from .license_registry import current_snapshot as current_license_snapshot
from .pipeline import process
from .registry import current_snapshot
from .registry_worker import RegistryWorker
from .workflow import search

ROOT = Path(__file__).resolve().parents[1]
MAX_BYTES = 100 * 1024 * 1024
MAX_BODY = 140 * 1024 * 1024
MAX_UNPACKED_BYTES = 512 * 1024 * 1024


class UserError(Exception):
    pass


class JobCancelled(Exception):
    pass


def unpack_uploads(files, deduplicate=True):
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
        if not deduplicate or digest not in seen:
            result.append((name, content))
            seen.add(digest)
    return result


class Application:
    def __init__(self, storage, registry, cache, token='', client_factory=DaDataClient,
                 token_store=None, registry_worker=None, licenses=None, licenses_worker=None):
        self.storage, self.registry, self.cache = Path(storage), Path(registry), Path(cache)
        self.licenses = Path(licenses) if licenses else ROOT.parent / 'license_registry_cache'
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
        self.licenses_worker = licenses_worker or RegistryWorker(self.licenses, updater=license_auto_update)
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
        try:
            licenses_snapshot = current_license_snapshot(self.licenses)
        except (ValueError, OSError, KeyError):
            licenses_snapshot = None
        licenses_date = licenses_snapshot[1].get('snapshot_date', '') if licenses_snapshot else ''
        return {'csrf': self.csrf, 'token_configured': bool(self.token), 'token_saved': self.token_saved,
                'registry_ready': bool(snapshot), 'registry_date': date,
                'registry_update': self.registry_worker.snapshot(),
                'licenses_ready': bool(licenses_snapshot), 'licenses_date': licenses_date,
                'licenses_update': self.licenses_worker.snapshot(), 'max_files': 10, 'max_mb': 100}

    def start_registry_update(self):
        return self.registry_worker.start()

    def start_licenses_update(self):
        return self.licenses_worker.start()

    def forget_token(self):
        if self.token_store:
            try:
                self.token_store.delete()
            except Exception:
                raise UserError('Не удалось удалить ключ из системного хранилища.') from None
        self.token = self.initial_token
        self.token_saved = False

    def inspect(self, payload):
        if not isinstance(payload, dict):
            raise UserError('Некорректные параметры загрузки.')
        result = []
        for name, content in unpack_uploads(payload.get('files'), deduplicate=False):
            try:
                sheets = list_sheets(io.BytesIO(content))
            except Exception:
                raise UserError(f'Не удалось прочитать листы «{name}». Проверьте, что книга открывается в Excel.') from None
            result.append({'name': name, 'sha256': hashlib.sha256(content).hexdigest(), 'sheets': sheets})
        return result

    def start(self, payload):
        if not isinstance(payload, dict):
            raise UserError('Некорректные параметры загрузки.')
        supplied = unpack_uploads(payload.get('files'), deduplicate=False)
        mode = payload.get('mode', 'registry')
        if not isinstance(mode, str) or mode not in {'registry', 'dadata_check'}:
            raise UserError('Выберите обработку по реестру или отдельную сверку DaData.')
        sheets = payload.get('sheets', [''] * len(supplied))
        if (not isinstance(sheets, list) or len(sheets) != len(supplied)
                or any(not isinstance(value, str) or len(value) > 31 for value in sheets)):
            raise UserError('Не удалось прочитать выбор листов. Выберите файлы повторно.')
        uploaded, selected_sheets, seen = [], [], set()
        for (name, content), sheet in zip(supplied, sheets):
            key = (hashlib.sha256(content).digest(), sheet)
            if key in seen:
                continue
            seen.add(key)
            uploaded.append((name, content))
            selected_sheets.append(sheet)
        token = (payload.get('token') or self.token) if mode == 'dadata_check' else ''
        if not isinstance(token, str) or len(token) > 256 or '\n' in token or '\r' in token:
            raise UserError('Проверьте API-ключ DaData.')
        token = token.strip()
        if mode == 'dadata_check' and not token:
            raise UserError('Для отдельной сверки введите API-ключ DaData. Обычная обработка по реестру работает без ключа.')
        with self.lock:
            if self.active:
                raise UserError('Уже идёт обработка. Дождитесь её завершения.')
            if mode == 'dadata_check' and payload.get('save_token') is True and payload.get('token'):
                if not self.token_store:
                    raise UserError('Системное хранилище ключей недоступно. Снимите флажок сохранения ключа.')
                try:
                    self.token_store.save(token)
                except Exception:
                    raise UserError('Не удалось сохранить ключ. Проверьте доступ к системному хранилищу или снимите флажок сохранения.') from None
                self.token_saved = True
            if mode == 'dadata_check':
                self.token = token
            job_id = secrets.token_hex(16)
            self.jobs[job_id] = {'id': job_id, 'status': 'running', 'phase': 'Читаем Excel',
                                 'progress': 2, 'done': 0, 'total': 0, 'error': ''}
            self.cancel_flags[job_id] = threading.Event()
            self.active = job_id
        thread = threading.Thread(target=self.run, args=(job_id, uploaded, token, mode, selected_sheets), daemon=True)
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
            return dict(job) if job else None

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

    def run(self, job_id, uploaded, token='', mode='registry', sheets=None):
        was_cancelled = False
        directory = self.storage / job_id
        try:
            self.check_cancel(job_id)
            directory.mkdir(parents=True, mode=0o700)
            paths, sheet_names = [], {}
            for index, (name, content) in enumerate(uploaded):
                self.check_cancel(job_id)
                folder = directory / 'inputs' / str(index)
                folder.mkdir(parents=True)
                path = folder / name
                path.write_bytes(content)
                paths.append(path)
                sheet_names[str(path.resolve())] = sheets[index] if sheets else ''
            config = json.loads((ROOT / 'config.json').read_text())
            records, sources = read_records(paths, should_cancel=lambda: self.check_cancel(job_id), sheet_names=sheet_names)
            self.check_cancel(job_id)
            if len(records) > 5000:
                raise UserError('В прототипе можно обработать до 5 000 строк за запуск. Разделите таблицу.')
            result = process(records, config)
            self.check_cancel(job_id)
            result.update(sources=sources, created_at_utc=datetime.now(timezone.utc).isoformat())
            phase = 'Ищем в реестрах и отдельно сверяем DaData' if mode == 'dadata_check' else 'Ищем организации в локальных реестрах'
            self.change(job_id, phase=phase, progress=8, total=len(result['groups']))
            client = self.client_factory(token) if mode == 'dadata_check' else None
            search(result, self.registry, config, licenses_path=self.licenses,
                cache=self.cache, client=client,
                on_progress=lambda done, total: self.progress(job_id, done, total),
                check_cancel=lambda: self.check_cancel(job_id))
            self.check_cancel(job_id)
            self.change(job_id, phase='Собираем итоговую таблицу', progress=93)
            aggregate(result)
            self.check_cancel(job_id)
            save(result, directory / 'result')
            self.check_cancel(job_id)
            summary = result['full_summary']
            errors = [r['status'] for r in result['resolution']['requests'] if r['status'] != 'OK']
            issues = summary['Исходных строк на проверку'] + result['summary']['selection_review_rows']
            warnings = list(result['warnings'])
            missing = summary['Строк без данных ученика']
            if missing:
                warnings.append(f'В {missing} исходных строках нет ФИО и/или номера личного дела.')
            self.finish(job_id,
                summary={'rows': summary['Входных строк'], 'schools': len(result['schools']),
                         'students': len(result['students']), 'review': issues,
                         'addresses': summary['Учреждений с адресом'],
                         'excluded': result['summary']['excluded_rows']},
                warnings=warnings, partial=bool(errors), preview=result['schools'][:20],
                download=f'/api/jobs/{job_id}/download', journal=f'/api/jobs/{job_id}/journal')
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

    def send(self, status, body, content_type='application/json; charset=utf-8', attachment=False):
        data = json.dumps(body, ensure_ascii=False).encode() if content_type.startswith('application/json') else body
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'")
        if attachment:
            filename = attachment if isinstance(attachment, str) else 'schools.xlsx'
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
        match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})(/download|/journal)?', self.path)
        if match:
            job = self.app.snapshot(match[1])
            if job:
                if not match[2]:
                    return self.send(200, job)
                if job['status'] != 'done':
                    return self.send(409, {'error': 'Результат ещё не готов.'})
                if match[2] == '/journal':
                    path = self.app.storage / match[1] / 'result' / 'processing_log.json'
                    return self.send(200, path.read_bytes(), 'application/octet-stream', attachment='processing_log.json')
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
        if self.path == '/api/registry/retry':
            return self.send(202, self.app.start_registry_update())
        if self.path == '/api/licenses/retry':
            return self.send(202, self.app.start_licenses_update())
        if self.path == '/api/token/forget':
            try:
                self.app.forget_token()
                return self.send(200, {'token_configured': False, 'token_saved': False})
            except UserError as error:
                return self.send(400, {'error': str(error)})
        if self.path not in {'/api/jobs', '/api/sheets'}:
            return self.send(404, {'error': 'Не найдено.'})
        try:
            size = int(self.headers.get('Content-Length', '0'))
            if not 0 < size <= MAX_BODY:
                raise UserError('Суммарный размер файлов не должен превышать 100 МБ.')
            self.connection.settimeout(60)
            payload = json.loads(self.rfile.read(size))
            if self.path == '/api/sheets':
                return self.send(200, self.app.inspect(payload))
            job_id = self.app.start(payload)
            self.send(202, {'id': job_id})
        except (UserError, ValueError, UnicodeError) as error:
            self.send(400, {'error': str(error) if isinstance(error, UserError) else 'Не удалось прочитать загрузку.'})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=8765)
    parser.add_argument('--storage', type=Path, default=ROOT.parent/'ui_runs')
    parser.add_argument('--registry', type=Path, default=ROOT.parent/'registry_cache')
    parser.add_argument('--licenses', type=Path, default=ROOT.parent/'license_registry_cache')
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
    app = Application(args.storage, args.registry, args.cache, token,
                      token_store=TokenStore(), licenses=args.licenses)
    try:
        server = ThreadingHTTPServer(('127.0.0.1', args.port), Handler)
    except OSError:
        parser.exit(1, 'Порт занят. Закройте старый экземпляр или укажите другой --port.\n')
    server.app = app
    app.start_registry_update()
    app.start_licenses_update()
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
