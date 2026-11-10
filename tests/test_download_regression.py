"""Regression: finished jobs must have a readable result, and the template must download."""
import base64
from http.client import HTTPConnection
from http.server import ThreadingHTTPServer
import io
import json
from pathlib import Path
import tempfile
import threading
import time
import unittest
from openpyxl import load_workbook
from ou_processor.web import Application, Handler

class DownloadRegression(unittest.TestCase):
    def test_real_processing_and_http_template(self):
        with tempfile.TemporaryDirectory() as root:
            app = Application(Path(root)/'runs', Path(root)/'registry', Path(root)/'cache.json')
            source = Path(__file__).resolve().parent / 'fixtures' / 'test_schools.xlsx'
            jid = app.start({'files':[{'name':source.name, 'data':base64.b64encode(source.read_bytes()).decode()}]})
            for _ in range(500):
                job = app.snapshot(jid)
                if job['status'] != 'running': break
                time.sleep(.02)
            self.assertEqual(job['status'], 'done', job.get('error'))
            self.assertEqual(job['summary']['rows'], 12)
            result = json.loads((app.storage/jid/'result'/'result.json').read_text())
            self.assertTrue(all(r['selection'] == 'ОТОБРАНО' for r in result['records']))
            server = ThreadingHTTPServer(('127.0.0.1',0),Handler)
            server.app=app
            thread=threading.Thread(target=server.serve_forever,daemon=True)
            thread.start()
            try:
                conn=HTTPConnection('127.0.0.1',server.server_port,timeout=10)
                conn.request('GET',f'/api/jobs/{jid}/decisions-template')
                response=conn.getresponse();data=response.read()
                self.assertEqual(response.status,200,data[:300])
                book=load_workbook(io.BytesIO(data),read_only=True)
                self.assertIn('Решения',book.sheetnames)
                book.close();conn.close()
            finally:
                server.shutdown();server.server_close();thread.join(timeout=3)
    def test_missing_result_not_reported_as_done(self):
        with tempfile.TemporaryDirectory() as root:
            app=Application(Path(root)/'runs',Path(root)/'registry',Path(root)/'cache.json')
            jid='a'*32
            app.jobs[jid]={'id':jid,'status':'done'}
            self.assertEqual(app.snapshot(jid)['status'],'error')
