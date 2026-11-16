"""Импорт UI/CLI, проверка чужого шаблона и повторное использование без поиска."""

import base64
from contextlib import redirect_stderr, redirect_stdout
from copy import deepcopy
from http.client import HTTPConnection
from http.server import ThreadingHTTPServer
from io import BytesIO, StringIO
import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from ou_processor.full_run import aggregate, main, save
from ou_processor.input import read_records
from ou_processor.manual_review import (ManualReviewError, apply_import, confirmed_for, load_store,
                                        save_store, template_bytes)
from ou_processor.pipeline import process
from ou_processor.resolution import resolve
from ou_processor.web import Application, Handler, UserError


CONFIG = json.loads((Path(__file__).resolve().parents[1] / 'config.json').read_text(encoding='utf-8'))
INN = '7707083893'


class AmbiguousRegistry:
    metadata = {'snapshot_date': '20261116'}
    def search(self, record):
        base = dict(candidate_name='Государственное бюджетное общеобразовательное учреждение «Школа № 1»',
                    candidate_ogrn='1027700132195', address='119019, г Москва, ул Тестовая, д 1',
                    postal_code='119019', region='Москва', city='Москва', legal_status='Действующее',
                    branch_type='MAIN', source='Реестр лицензий')
        return [dict(base, candidate_inn=inn) for inn in (INN, '7710140679')], False


def make_result(root):
    path = Path(root) / 'source.xlsx'
    book = Workbook()
    book.active.title = 'Данные'
    book.active.append(['Образовательное учреждение', 'Вид образования', 'ФИО', 'Номер личного дела',
                        'Регион', 'Город'])
    book.active.append(['ГБОУ школа № 1 г. Москва', 'Среднее общее', 'Тестовый Ученик', '0001', 'Москва', 'Москва'])
    book.create_sheet('Архив')['A1'] = 'Не менять'
    book.save(path)
    records, sources = read_records([path])
    result = process(records, CONFIG)
    result.update(sources=sources, created_at_utc='2026-11-16T12:00:00+00:00')
    resolve(result, None, AmbiguousRegistry(), progress=False, config=CONFIG)
    aggregate(result)
    assert len(result['review']) == 1
    return result


def decisions(result, **changes):
    book = load_workbook(BytesIO(template_bytes(result)))
    book.active['E2'] = INN
    book.active['F2'] = 'Проверено оператором'
    for cell, value in changes.items():
        book.active[cell] = value
    stream = BytesIO()
    book.save(stream)
    book.close()
    return stream.getvalue()


class ManualReviewTests(unittest.TestCase):
    def test_roundtrip_updates_report_log_and_filled_copies(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            original = make_result(root)
            source = root / 'source.xlsx'
            before = source.read_bytes()
            updated, entries = apply_import(original, decisions(original), {})
            self.assertIsNone(original['resolution']['groups'][0]['selected'])
            self.assertEqual(updated['resolution']['groups'][0]['status'], 'ПРОВЕРЕНО')
            aggregate(updated)
            save(updated, root / 'result')
            self.assertEqual(updated['schools'][0]['status'], 'ПРОВЕРЕНО')
            self.assertEqual(updated['review'], [])
            self.assertEqual(updated['records'][0]['manual_comment'], 'Проверено оператором')
            report = load_workbook(root / 'result' / 'schools.xlsx', read_only=True)
            try:
                rows = report['Письма'].values
                letter = dict(zip(next(rows), next(rows)))
                self.assertEqual(letter['Статус сопоставления'], 'ПРОВЕРЕНО')
                self.assertEqual(letter['Комментарий оператора'], 'Проверено оператором')
            finally:
                report.close()
            copy_path = root / 'result' / 'filled_sources' / updated['source_fill']['copies'][0]['file']
            book = load_workbook(copy_path)
            try:
                self.assertEqual(book['Архив']['A1'].value, 'Не менять')
                values = dict(zip([c.value for c in book['Данные'][1]], [c.value for c in book['Данные'][2]]))
                self.assertTrue(values['Полное название ОУ в именительном падеже'])
                self.assertTrue(values['Название ОУ в родительном падеже'])
                self.assertEqual(values['Индекс'], '119019')
            finally:
                book.close()
            self.assertEqual(source.read_bytes(), before)
            log = json.loads((root / 'result' / 'processing_log.json').read_text())
            self.assertTrue(log['groups'][0]['manual_decision'])
            self.assertEqual(log['groups'][0]['manual_comment'], 'Проверено оператором')
            self.assertNotIn('Тестовый Ученик', json.dumps(log, ensure_ascii=False))
            save_store(root / 'store.json', entries)
            self.assertEqual(load_store(root / 'store.json'), entries)

    def test_invalid_rows_never_mutate_result_or_store(self):
        bad = [dict(E2='1234567890'), dict(E2='0000000000'), dict(E2='7801000000'),
               dict(A2='unknown'), dict(B2='Другая школа'), dict(C2='Другой регион'),
               dict(I2='other-run'), dict(E2='=7707083893'), dict(F2='=HYPERLINK("x")'),
               dict(F2='x' * 1001), dict(A1='Изменён заголовок'), dict(G2='Изменён список кандидатов')]
        with tempfile.TemporaryDirectory() as tmp:
            result = make_result(tmp)
            initial = deepcopy(result)
            for changed in bad:
                with self.subTest(changed=changed):
                    entries = {}
                    with self.assertRaises(ManualReviewError):
                        apply_import(result, decisions(result, **changed), entries)
                    self.assertEqual(result, initial)
                    self.assertEqual(entries, {})

    def test_valid_inn_outside_candidates_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = make_result(tmp)
            with self.assertRaisesRegex(ManualReviewError, 'из списка'):
                apply_import(result, decisions(result, E2='7702070139'), {})

    def test_duplicate_group_and_partially_invalid_file_are_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = make_result(tmp)
            book = load_workbook(BytesIO(decisions(result)))
            book.active.append([c.value for c in book.active[2]])
            stream = BytesIO()
            book.save(stream)
            book.close()
            with self.assertRaisesRegex(ManualReviewError, 'повторная'):
                apply_import(result, stream.getvalue(), {})
            self.assertIsNone(result['resolution']['groups'][0]['selected'])

    def test_empty_and_broken_files_have_clear_errors(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = make_result(tmp)
            for content in (b'', b'not a workbook', template_bytes(result)):
                with self.subTest(content_size=len(content)), self.assertRaises(ManualReviewError):
                    apply_import(result, content, {})

    def test_other_run_and_already_confirmed_group_are_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = make_result(tmp)
            content = decisions(result)
            other = deepcopy(result)
            other['created_at_utc'] = '2026-11-16T12:01:00+00:00'
            with self.assertRaisesRegex(ManualReviewError, 'другому запуску'):
                apply_import(other, content, {})
            updated, entries = apply_import(result, content, {})
            with self.assertRaisesRegex(ManualReviewError, 'уже подтверждённая'):
                apply_import(updated, content, entries)

    def test_inactive_candidate_cannot_be_confirmed(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = make_result(tmp)
            for candidate in result['resolution']['candidates']:
                candidate['legal_status'] = 'LIQUIDATED'
            with self.assertRaisesRegex(ManualReviewError, 'действующего'):
                apply_import(result, decisions(result), {})

    def test_reuse_skips_all_searches_and_keeps_new_students(self):
        with tempfile.TemporaryDirectory() as tmp:
            initial = make_result(tmp)
            _, entries = apply_import(initial, decisions(initial), {})
            records, _ = read_records([Path(tmp) / 'source.xlsx'])
            records[0]['student_name'], records[0]['personal_file_number'] = 'Другой Ученик', '0002'
            rerun = process(records, CONFIG)
            class Forbidden:
                metadata = {}
                actual_requests = 0
                requests = []
                def search(self, *args):
                    raise AssertionError('Подтверждённая организация повторно не ищется')
            resolve(rerun, Forbidden(), Forbidden(), licenses=Forbidden(),
                    progress=False, config=CONFIG, manual_decisions=entries)
            aggregate(rerun)
            self.assertEqual(rerun['resolution']['groups'][0]['status'], 'ПРОВЕРЕНО')
            self.assertEqual(rerun['students'][0]['student'], 'Другой Ученик')
            self.assertEqual(rerun['students'][0]['file'], '0002')

    def test_changed_geography_year_or_risks_do_not_reuse_confirmation(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = make_result(tmp)
            _, entries = apply_import(result, decisions(result), {})
            for changed in ({'city': 'Калуга'}, {'year': '2000'}, {'flags': ['СТАРЫЙ_ВЫПУСК']},
                            {'inn_raw': '7702070139'}, {'school_address_raw': 'Другой адрес'}, {'postal_code_raw': '000001'}):
                record = {**result['records'][0], **changed}
                self.assertIsNone(confirmed_for([record], entries))

    def test_low_geography_or_risky_case_requires_confirmation_each_run(self):
        for fields in ({'region': ''}, {'city': ''}, {'flags': ['СТАРЫЙ_ВЫПУСК']},
                       {'flags': ['КОНФЛИКТ_РЕГИОНА']}, {'flags': ['ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ']}):
            with self.subTest(fields=fields), tempfile.TemporaryDirectory() as tmp:
                result = make_result(tmp)
                result['records'][0].update(fields)
                _, entries = apply_import(result, decisions(result), {})
                self.assertIsNone(confirmed_for(result['records'], entries))

    def test_corrupt_store_is_not_silently_overwritten(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'store.json'
            for data in ('broken', '{}', '{"version":2,"decisions":{}}'):
                path.write_text(data)
                with self.assertRaises(ManualReviewError):
                    load_store(path)
                self.assertEqual(path.read_text(), data)

    def test_cli_import_is_offline_and_writes_new_result(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            result = make_result(root)
            save(result, root / 'initial')
            (root / 'decisions.xlsx').write_bytes(decisions(result))
            output, errors = StringIO(), StringIO()
            with patch('ou_processor.full_run.search', side_effect=AssertionError('Импорт не запускает поиск')), redirect_stdout(output), redirect_stderr(errors):
                code = main(['--result', str(root / 'initial/result.json'), '--import-decisions', str(root / 'decisions.xlsx'),
                             '--output', str(root / 'reviewed'), '--decisions-store', str(root / 'store.json')])
            self.assertEqual(code, 0, errors.getvalue())
            updated = json.loads((root / 'reviewed/result.json').read_text())
            self.assertEqual(updated['schools'][0]['status'], 'ПРОВЕРЕНО')
            self.assertTrue(load_store(root / 'store.json'))
            self.assertTrue(json.loads((root / 'initial/result.json').read_text())['review'])

    def test_cli_normal_run_uses_shared_store_before_search(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            result = make_result(root)
            _, entries = apply_import(result, decisions(result), {})
            save_store(root / 'store.json', entries)
            class NoSearch:
                metadata = {}
                def __init__(self, *args, **kwargs):
                    self.connection = self
                def close(self):
                    pass
                def search(self, *args):
                    raise AssertionError('CLI не должен искать подтверждённую организацию')
            with patch('ou_processor.workflow.RegistrySearch', NoSearch), patch(
                    'ou_processor.workflow.LicenseSearch', NoSearch), redirect_stdout(StringIO()), redirect_stderr(StringIO()):
                code = main([str(root / 'source.xlsx'), '--output', str(root / 'out'), '--decisions-store', str(root / 'store.json')])
            self.assertEqual(code, 0)
            updated = json.loads((root / 'out/result.json').read_text())
            self.assertEqual(updated['resolution']['groups'][0]['status'], 'ПРОВЕРЕНО')

    def test_explicit_dadata_candidate_can_only_be_confirmed_by_operator(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = make_result(tmp)
            result['dadata_candidates'] = [dict(result['resolution']['candidates'][0], source='DaData',
                                               legal_status='ACTIVE', score=100, eligible=True, reasons=[])]
            result['resolution']['candidates'] = []
            self.assertIsNone(result['resolution']['groups'][0]['selected'])
            updated, _ = apply_import(result, decisions(result), {})
            self.assertEqual(updated['resolution']['groups'][0]['status'], 'ПРОВЕРЕНО')
            self.assertEqual(updated['resolution']['groups'][0]['selected']['source'], 'DaData')
            self.assertEqual(len(updated['resolution']['candidates']), 1)

    def test_cli_invalid_result_has_no_traceback_and_does_not_create_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'invalid.json').write_text('[]')
            (root / 'decisions.xlsx').write_bytes(b'broken')
            errors = StringIO()
            with redirect_stderr(errors), redirect_stdout(StringIO()):
                code = main(['--result', str(root / 'invalid.json'), '--import-decisions', str(root / 'decisions.xlsx'),
                             '--output', str(root / 'out'), '--decisions-store', str(root / 'store.json')])
            self.assertEqual(code, 1)
            self.assertIn('result.json', errors.getvalue())
            self.assertNotIn('Traceback', errors.getvalue())
            self.assertFalse((root / 'out').exists())

    def test_cli_missing_source_reports_partial_copy_but_keeps_reviewed_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            result = make_result(root)
            save(result, root / 'initial')
            (root / 'decisions.xlsx').write_bytes(decisions(result))
            (root / 'source.xlsx').unlink()
            errors = StringIO()
            with redirect_stderr(errors), redirect_stdout(StringIO()):
                code = main(['--result', str(root / 'initial/result.json'), '--import-decisions', str(root / 'decisions.xlsx'),
                             '--output', str(root / 'out'), '--decisions-store', str(root / 'store.json')])
            self.assertEqual(code, 2)
            self.assertIn('копию исходной книги', errors.getvalue())
            self.assertTrue((root / 'out' / 'schools.xlsx').exists())

    def test_ui_import_refreshes_metrics_and_rejects_invalid_before_writing(self):
        with tempfile.TemporaryDirectory() as tmp:
            root, job_id = Path(tmp), 'a' * 32
            app = Application(root / 'runs', root / 'registry', root / 'cache.json', decisions_store=root / 'store.json')
            result = make_result(root)
            folder = app.storage / job_id / 'result'
            save(result, folder)
            app.jobs[job_id] = {'id': job_id, 'status': 'done'}
            before = (folder / 'result.json').read_bytes()
            with self.assertRaises(UserError):
                app.import_decisions(job_id, {'data': base64.b64encode(decisions(result, I2='wrong')).decode()})
            self.assertEqual((folder / 'result.json').read_bytes(), before)
            self.assertFalse((root / 'store.json').exists())
            self.assertIsNone(app.active)
            payload = {'data': base64.b64encode(decisions(result)).decode()}
            with patch('ou_processor.web.save_store', side_effect=OSError('Справочник недоступен')):
                with self.assertRaises(UserError):
                    app.import_decisions(job_id, payload)
            self.assertEqual((folder / 'result.json').read_bytes(), before)
            self.assertIsNone(app.active)
            response = app.import_decisions(job_id, payload)
            self.assertEqual(response['imported'], 1)
            self.assertEqual(response['job']['manual_pending'], 0)
            self.assertEqual(response['job']['summary']['verified'], 1)
            self.assertEqual(response['job']['summary']['review'], 0)
            self.assertEqual(response['job']['summary']['schools'], 1)


class ManualHTTPTests(unittest.TestCase):
    def test_template_download_and_import_require_finished_job_and_csrf(self):
        with tempfile.TemporaryDirectory() as tmp:
            root, job_id = Path(tmp), 'b' * 32
            app = Application(root / 'runs', root / 'registry', root / 'cache.json', decisions_store=root / 'store.json')
            result = make_result(root)
            save(result, app.storage / job_id / 'result')
            app.jobs[job_id] = {'id': job_id, 'status': 'done'}
            server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
            server.app = app
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            def call(path, method='GET', body=None, headers=None):
                connection = HTTPConnection('127.0.0.1', server.server_port, timeout=10)
                connection.request(method, path, body, headers or {})
                response = connection.getresponse()
                status, data = response.status, response.read()
                connection.close()
                return status, data
            try:
                self.assertEqual(call(f'/api/jobs/{job_id}/manual-template')[0], 200)
                payload = json.dumps({'data': base64.b64encode(decisions(result)).decode()})
                path = f'/api/jobs/{job_id}/decisions-import'
                self.assertEqual(call(path, 'POST', payload)[0], 403)
                status, data = call(path, 'POST', payload, {'X-CSRF-Token': app.csrf})
                self.assertEqual(status, 200, data)
                self.assertEqual(json.loads(data)['job']['summary']['verified'], 1)
                app.jobs[job_id]['status'] = 'running'
                self.assertEqual(call(f'/api/jobs/{job_id}/manual-template')[0], 409)
            finally:
                server.shutdown()
                server.server_close()
                thread.join()
