"""Проверка импорта ручных решений и повторного применения без поиска."""
import base64
import io
import json
from pathlib import Path
import tempfile
import unittest
from openpyxl import load_workbook
from ou_processor.web import Application, UserError
from ou_processor.resolution import resolve


class ManualImportTests(unittest.TestCase):
    def test_import_validates_and_updates_excel(self):
        with tempfile.TemporaryDirectory() as root:
            app = Application(Path(root)/'runs', Path(root)/'registry', Path(root)/'cache.json')
            job_id = 'a'*32
            folder = app.storage/job_id/'result'
            folder.mkdir(parents=True)
            candidate = {'group_id':'g1','candidate_inn':'7707083893', 'candidate_name':'Тестовая школа',
                         'address':'Москва','source':'Реестр', 'score':91, 'candidate_ogrn':'',
                         'postal_code':'', 'region':'Москва','reasons':[], 'eligible':True, 'legal_status':'ACTIVE'}
            record = {'record_id':'r1','name':'Школа','region':'Москва','city':'Москва',
                      'student_name':'Иванов Иван','personal_file_number':'001'}
            group = {'group_id':'g1','record_ids':['r1'],'selected':None,'status':'НЕОДНОЗНАЧНО',
                     'registry_used':True,'registry_truncated':False,'candidate_count':1}
            result = {'records':[record], 'resolution':{'groups':[group],'candidates':[candidate],
                       'requests':[], 'actual_requests':0,'registry':{}}, 'summary':{'selected_rows':1,
                       'selection_review_rows':0}, 'sources':[], 'created_at_utc':'2026-11-08'}
            from ou_processor.full_run import aggregate, save
            aggregate(result)
            save(result,folder)
            app.jobs[job_id] = {'status':'done'}
            template = load_workbook(io.BytesIO(app.decisions_template(job_id)))
            template.active['E2'] = '7707083893'
            template.active['F2'] = 'Подтверждено оператором'
            stream = io.BytesIO()
            template.save(stream)
            response = app.import_manual_decisions(job_id, {'data':base64.b64encode(stream.getvalue()).decode()})
            self.assertEqual(response['imported'],1)
            updated = json.loads((folder/'result.json').read_text())
            self.assertEqual(updated['schools'][0]['status'],'ПРОВЕРЕНО')
            self.assertEqual(updated['records'][0]['resolved_inn'],'7707083893')
            self.assertEqual(updated['records'][0]['manual_comment'],'Подтверждено оператором')
            self.assertEqual(updated['review'],[])
            self.assertTrue(app.load_manual_decisions())
            template.active['E2']='1234567890'
            stream=io.BytesIO(); template.save(stream)
            with self.assertRaises(UserError):
                app.import_manual_decisions(job_id, {'data':base64.b64encode(stream.getvalue()).decode()})

    def test_inn_checksum(self):
        self.assertTrue(Application.valid_inn('7707083893'))
        self.assertFalse(Application.valid_inn('1234567890'))
        self.assertFalse(Application.valid_inn('abc'))
