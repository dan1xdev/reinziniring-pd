"""Общий сценарий UI/CLI на Python 3.12, синтетические данные, без внешней сети."""

import base64
from contextlib import redirect_stderr, redirect_stdout
from io import BytesIO, StringIO
import json
import os
from pathlib import Path
import sqlite3
import tempfile
import time
import unittest
import zipfile
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from ou_processor.full_run import main
from ou_processor.input import read_records
from ou_processor.pipeline import process
from ou_processor.web import Application, UserError
from ou_processor.workflow import search


CONFIG = json.loads((Path(__file__).resolve().parents[1] / 'config.json').read_text(encoding='utf-8'))


def upload(name='input.xlsx'):
    book = Workbook()
    book.active.title = 'Первый'
    for sheet in (book.active, book.create_sheet('Рабочий')):
        sheet.append(['Образовательное учреждение', 'Вид образования', 'ФИО', 'Личное дело.Номер'])
        sheet.append(['ГБОУ г. Москвы «Школа № 1»', 'Среднее общее', 'Синтетический Ученик', '0001'])
    stream = BytesIO()
    book.save(stream)
    return {'name': name, 'data': base64.b64encode(stream.getvalue()).decode()}


def result():
    return process([dict(record_id='s:Рабочий:2', source_file='input.xlsx', source_sheet='Рабочий',
                         source_row=2, name='ГБОУ г. Москвы «Школа № 1»', education='Среднее общее',
                         formula_fields=[], student_name='Синтетический Ученик', personal_file_number='0001')], CONFIG)


def candidate(inn='1111111111', source='Реестр аккредитации'):
    return dict(candidate_inn=inn, candidate_ogrn='1111111111111',
                candidate_name='ГБОУ г. Москвы «Школа № 1»', address='123456, г Москва, ул А, д 1',
                postal_code='123456', region='Москва', city='Москва', legal_status='ACTIVE',
                branch_type='MAIN', source=source)


class Registry:
    metadata = {'snapshot_date': '20261108', 'source_kind': 'accreditation_registry'}
    def __init__(self, *args, **kwargs):
        self.connection = self
    def close(self):
        pass
    def search(self, record):
        return [candidate()], False


def wait_job(app, job_id):
    for _ in range(500):
        job = app.snapshot(job_id)
        if job['status'] not in {'running', 'cancelling'}:
            return job
        time.sleep(.01)
    raise AssertionError('Обработка не завершилась за 5 секунд')


class WorkflowTests(unittest.TestCase):
    def test_registry_does_not_read_dadata_cache_or_create_client(self):
        with patch('ou_processor.workflow.RegistrySearch', Registry), patch(
                'ou_processor.workflow.SearchCache', side_effect=AssertionError('Кэш DaData не нужен')):
            calculated = search(result(), 'unused', CONFIG)
        self.assertEqual(calculated['resolution']['groups'][0]['selected']['candidate_inn'], '1111111111')
        self.assertEqual(calculated['resolution']['requests'], [])
        self.assertEqual(calculated['resolution']['actual_requests'], 0)

    def test_missing_registry_keeps_case_for_operator_without_dadata(self):
        with patch('ou_processor.workflow.RegistrySearch', side_effect=ValueError('нет снимка')), patch(
                'ou_processor.workflow.SearchCache', side_effect=AssertionError('Без автоматического перехода')):
            calculated = search(result(), 'missing', CONFIG)
        self.assertIsNone(calculated['resolution']['groups'][0]['selected'])
        self.assertTrue(calculated['warnings'])

    def test_corrupt_registry_keeps_case_for_operator(self):
        class Broken(Registry):
            def search(self, record):
                raise sqlite3.DatabaseError('broken')
        with patch('ou_processor.workflow.RegistrySearch', Broken):
            calculated = search(result(), 'broken', CONFIG)
        self.assertIsNone(calculated['resolution']['groups'][0]['selected'])
        self.assertIn('индекс', calculated['warnings'][0])

    def test_empty_registry_never_uses_dadata(self):
        class Empty(Registry):
            def search(self, record):
                return [], False
        with patch('ou_processor.workflow.RegistrySearch', Empty), patch(
                'ou_processor.workflow.SearchCache', side_effect=AssertionError('Нет резервного DaData')):
            calculated = search(result(), 'empty', CONFIG)
        self.assertEqual(calculated['resolution']['groups'][0]['status'], 'НЕ_НАЙДЕНО')

    def test_explicit_conflicting_check_preserves_registry_inn(self):
        class Cache:
            actual_requests, requests = 1, []
            def __init__(self, *args):
                pass
            def search(self, query, group):
                if 'Синтетический' in query or '0001' in query:
                    raise AssertionError('Персональные данные попали в запрос')
                return [candidate('2222222222', 'DaData')]
        with patch('ou_processor.workflow.RegistrySearch', Registry), patch('ou_processor.workflow.SearchCache', Cache):
            calculated = search(result(), 'unused', CONFIG, cache='unused', client=object())
        self.assertEqual(calculated['resolution']['groups'][0]['selected']['candidate_inn'], '1111111111')
        self.assertEqual(calculated['dadata_checks'][0]['status'], 'КОНФЛИКТ_ИНН')
        self.assertEqual(calculated['dadata_checks'][0]['dadata_inn'], '2222222222')

    def test_explicit_check_cannot_confirm_unmatched_case(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / 'dadata.json'
            cache.write_text(json.dumps({'ГБОУ г. Москвы «Школа № 1» Москва': {
                'status': 'OK', 'candidates': [candidate(source='DaData')]}}), encoding='utf-8')
            with patch('ou_processor.workflow.RegistrySearch', side_effect=ValueError('нет снимка')):
                calculated = search(result(), 'missing', CONFIG, cache=cache, client=object())
        self.assertIsNone(calculated['resolution']['groups'][0]['selected'])
        self.assertEqual(calculated['dadata_checks'][0]['dadata_inn'], '1111111111')
        self.assertEqual(calculated['dadata_checks'][0]['status'], 'ТРЕБУЕТ_ПРОВЕРКИ')

    def test_ui_default_with_saved_token_and_corrupt_cache_never_constructs_client(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'dadata.json').write_text('broken', encoding='utf-8')
            app = Application(root / 'runs', root / 'registry', root / 'dadata.json', token='secret',
                              client_factory=lambda _: self.fail('DaData client created'))
            data = upload()
            info = app.inspect({'files': [data]})[0]
            job_id = app.start({'files': [data], 'sheets': ['Рабочий']})
            job = wait_job(app, job_id)
            self.assertEqual(job['status'], 'done', job.get('error'))
            log = json.loads((root / 'runs' / job_id / 'result' / 'processing_log.json').read_text())
            self.assertEqual(log['mode'], 'registry')
            self.assertEqual(log['dadata_actual_requests'], 0)
            self.assertEqual(log['sources'][0]['selected_sheet'], 'Рабочий')
            self.assertEqual(log['selection'][0]['source_sheet'], 'Рабочий')
            self.assertNotIn('secret', json.dumps(log))
            self.assertNotIn('Синтетический Ученик', json.dumps(log, ensure_ascii=False))

    def test_sheet_inspection_keeps_duplicate_names_and_contents_in_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            app = Application(tmp, tmp, Path(tmp) / 'cache.json')
            data = upload()
            inspected = app.inspect({'files': [data, data, {**data, 'name': 'copy.xlsx'}]})
        self.assertEqual(len(inspected), 3)
        self.assertEqual(inspected[0]['sheets'], ['Первый', 'Рабочий'])
        self.assertEqual(inspected[2]['name'], 'copy.xlsx')

    def test_per_file_sheet_selection(self):
        with tempfile.TemporaryDirectory() as tmp:
            paths = [Path(tmp) / 'a.xlsx', Path(tmp) / 'b.xlsx']
            for index, path in enumerate(paths):
                path.write_bytes(base64.b64decode(upload()['data']))
                book = load_workbook(path)
                book['Рабочий']['D2'] = f'000{index + 2}'
                book.save(path)
                book.close()
            records, _ = read_records(paths, sheet_names={str(paths[0]): 'Первый', str(paths[1]): 'Рабочий'})
        self.assertEqual({r['source_sheet'] for r in records}, {'Первый', 'Рабочий'})

    def test_ui_keeps_identical_books_with_different_selected_sheets(self):
        book = Workbook()
        book.active.title = 'Первый'
        for sheet, school, file_number in (
                (book.active, 'Школа 1', '0001'),
                (book.create_sheet('Второй'), 'Школа 2', '0002')):
            sheet.append(['Образовательное учреждение', 'Вид образования', 'ФИО', 'Личное дело.Номер'])
            sheet.append([school, 'Среднее общее', 'Ученик', file_number])
        stream = BytesIO()
        book.save(stream)
        content = base64.b64encode(stream.getvalue()).decode()
        files = [{'name': 'a.xlsx', 'data': content}, {'name': 'b.xlsx', 'data': content}]

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            app = Application(root / 'runs', root / 'registry', root / 'cache.json')
            job_id = app.start({'files': files, 'sheets': ['Первый', 'Второй']})
            job = wait_job(app, job_id)
            self.assertEqual(job['status'], 'done', job.get('error'))
            result = json.loads((root / 'runs' / job_id / 'result' / 'result.json').read_text())

        self.assertEqual([(source['file'], source['selected_sheet']) for source in result['sources']],
                         [('a.xlsx', 'Первый'), ('b.xlsx', 'Второй')])
        self.assertEqual({record['name'] for record in result['records']}, {'Школа 1', 'Школа 2'})

    def test_invalid_mode_and_sheet_selection_are_explained(self):
        with tempfile.TemporaryDirectory() as tmp:
            app = Application(tmp, tmp, Path(tmp) / 'cache.json')
            for extra in ({'mode': []}, {'sheets': ['Рабочий', 'Лишний']}):
                with self.assertRaises(UserError):
                    app.start({'files': [upload()], **extra})

    def test_cli_lists_sheets_without_output_or_key(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'input.xlsx'
            path.write_bytes(base64.b64decode(upload()['data']))
            stream = StringIO()
            with redirect_stdout(stream):
                code = main([str(path), '--list-sheets'])
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(stream.getvalue())['sheets'], ['Первый', 'Рабочий'])

    def test_cli_corrupt_excel_is_explained_in_both_commands(self):
        for listing in (True, False):
            with self.subTest(listing=listing), tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                path = root / 'broken.xlsx'
                raw = BytesIO(base64.b64decode(upload()['data']))
                with zipfile.ZipFile(raw) as original, zipfile.ZipFile(path, 'w') as broken:
                    for member in original.namelist():
                        broken.writestr(member, b'<broken>' if member == 'xl/workbook.xml' else original.read(member))
                args = [str(path), '--list-sheets'] if listing else [str(path), '--output', str(root / 'result')]
                errors = StringIO()
                with redirect_stdout(StringIO()), redirect_stderr(errors):
                    code = main(args)
                self.assertEqual(code, 1)
                self.assertNotIn('Traceback', errors.getvalue())

    def test_cli_default_exports_review_and_journal_without_key(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            path = root / 'input.xlsx'
            path.write_bytes(base64.b64decode(upload()['data']))
            with patch.dict(os.environ, {'DADATA_API_KEY': 'ignored-secret'}), patch(
                    'ou_processor.full_run.DaDataClient', side_effect=AssertionError('No API')), redirect_stdout(StringIO()), redirect_stderr(StringIO()):
                code = main([str(path), '--sheet', 'Рабочий', '--registry', str(root / 'missing'), '--output', str(root / 'result')])
            self.assertEqual(code, 0)
            self.assertTrue((root / 'result' / 'processing_log.json').is_file())
            book = load_workbook(root / 'result' / 'schools.xlsx', read_only=True)
            try:
                self.assertEqual(book['Проверить'].max_row, 2)
                self.assertEqual(book['Запросы DaData'].max_row, 1)
            finally:
                book.close()

    def test_cli_missing_sheet_returns_clear_error_without_traceback(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            path = root / 'input.xlsx'
            path.write_bytes(base64.b64decode(upload()['data']))
            errors = StringIO()
            with redirect_stdout(StringIO()), redirect_stderr(errors):
                code = main([str(path), '--sheet', 'Нет такого', '--output', str(root / 'result')])
        self.assertEqual(code, 1)
        self.assertIn('Доступные листы', errors.getvalue())
        self.assertNotIn('Traceback', errors.getvalue())


if __name__ == '__main__':
    unittest.main()
