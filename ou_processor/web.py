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
from .full_run import aggregate, save
from .input import read_records
from .pipeline import process
from .registry import current_snapshot
from .resolution import RegistrySearch, SearchCache, resolve

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
    def __init__(self, storage, registry, cache, token='', client_factory=DaDataClient):
        self.storage, self.registry, self.cache = Path(storage), Path(registry), Path(cache)
        self.token, self.client_factory = token, client_factory
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
        return {'csrf': self.csrf, 'token_configured': bool(self.token), 'registry_ready': bool(snapshot),
                'registry_date': date, 'max_files': 10, 'max_mb': 100}

    def start(self, payload):
        if not isinstance(payload, dict):
            raise UserError('Некорректные параметры загрузки.')
        uploaded = unpack_uploads(payload.get('files'))
        token = payload.get('token') or self.token
        if not isinstance(token, str) or len(token) > 256 or '\n' in token or '\r' in token:
            raise UserError('Проверьте API-ключ DaData.')
        token = token.strip()
        offline = payload.get('offline') is True
        if not token and not offline:
            raise UserError('Введите API-ключ DaData или выберите обработку по сохранённым ответам.')
        if not self.settings()['registry_ready']:
            raise UserError('Резервный реестр не загружен. Выполните команду из инструкции запуска.')
        with self.lock:
            if self.active:
                raise UserError('Уже идёт обработка. Дождитесь её завершения.')
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
            reserve = RegistrySearch(self.registry)
            resolve(result, api, reserve, progress=False,
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
            warnings = []
            if errors:
                warnings.append('Не все запросы DaData выполнены. Проверьте лист «Запросы DaData»: ' + ', '.join(sorted(set(errors))))
            if offline:
                warnings.append('Использованы сохранённые ответы DaData и локальный реестр. Новых запросов не было.')
            missing = summary['Строк без данных ученика']
            if missing:
                warnings.append(f'В {missing} исходных строках нет ФИО и/или номера личного дела.')
            self.finish(job_id,
                summary={'rows': summary['Входных строк'], 'schools': len(result['schools']),
                         'students': len(result['students']), 'review': issues,
                         'addresses': summary['Учреждений с адресом'],
                         'excluded': result['summary']['excluded_rows']},
                warnings=warnings, partial=bool(errors), preview=result['schools'][:20],
                download=f'/api/jobs/{job_id}/download')
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
            self.send_header('Content-Disposition', 'attachment; filename="schools.xlsx"')
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
    app = Application(args.storage, args.registry, args.cache, token)
    try:
        server = ThreadingHTTPServer(('127.0.0.1', args.port), Handler)
    except OSError:
        parser.exit(1, 'Порт занят. Закройте старый экземпляр или укажите другой --port.\n')
    server.app = app
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
