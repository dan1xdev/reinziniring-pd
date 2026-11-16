"""Автоматическое подключение DaData в UI и CLI без внешней сети."""

import base64
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
import json
import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch

from ou_processor.full_run import main
from ou_processor.web import Application
from tests.test_web import school_upload


class Client:
    def __init__(self, token):
        self.token = token
        self.queries = []

    def suggest(self, query):
        self.queries.append(query)
        return []


class EntryPointTests(unittest.TestCase):
    # Ключ из любого поддержанного источника запускает каскад обычной кнопкой.
    def test_ui_default_uses_payload_environment_or_saved_key(self):
        for origin in ('payload', 'environment', 'saved'):
            with self.subTest(origin=origin), tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                clients = []
                def factory(token):
                    client = Client(token)
                    clients.append(client)
                    return client
                class Store:
                    def load(self):
                        return 'private-test-key'
                env = {'DADATA_API_KEY': 'private-test-key' if origin == 'environment' else ''}
                with patch.dict(os.environ, env):
                    app = Application(root / 'runs', root / 'registry', root / 'cache.json',
                                      licenses=root / 'licenses', client_factory=factory,
                                      token_store=Store() if origin == 'saved' else None)
                    payload = {'files': [school_upload()]}
                    if origin == 'payload':
                        payload['token'] = 'private-test-key'
                    job_id = app.start(payload)
                    for _ in range(1000):
                        job = app.snapshot(job_id)
                        if job['status'] != 'running':
                            break
                        time.sleep(.01)
                self.assertEqual(job['status'], 'done', job.get('error'))
                self.assertEqual(len(clients), 1)
                self.assertEqual(clients[0].token, 'private-test-key')
                self.assertTrue(clients[0].queries)
                self.assertNotIn('Иванов', ' '.join(clients[0].queries))
                self.assertNotIn('0001', ' '.join(clients[0].queries))
                self.assertNotIn('private-test-key', json.dumps(job))
                journal = root / 'runs' / job_id / 'result' / 'processing_log.json'
                self.assertNotIn('private-test-key', journal.read_text())

    # CLI читает env и запрашивает ключ без отдельного флага режима.
    def test_cli_uses_environment_key_and_standalone_prompt(self):
        for prompted in (False, True):
            with self.subTest(prompted=prompted), tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                source = root / 'input.xlsx'
                source.write_bytes(base64.b64decode(school_upload()['data']))
                client = Client('private-test-key')
                stdout, stderr = StringIO(), StringIO()
                arguments = [str(source), '--output', str(root / 'result'),
                             '--registry', str(root / 'registry'), '--licenses', str(root / 'licenses'),
                             '--cache', str(root / 'cache.json')]
                if prompted:
                    arguments.append('--ask-token')
                with patch.dict(os.environ, {'DADATA_API_KEY': '' if prompted else 'private-test-key'}), patch(
                        'ou_processor.full_run.getpass.getpass', return_value='private-test-key'), patch(
                        'ou_processor.full_run.DaDataClient', side_effect=lambda token: client if token == client.token else None), redirect_stdout(stdout), redirect_stderr(stderr):
                    code = main(arguments)
                self.assertEqual(code, 0)
                self.assertTrue(client.queries)
                self.assertNotIn('private-test-key', stdout.getvalue() + stderr.getvalue())
                self.assertNotIn('private-test-key', (root / 'result' / 'processing_log.json').read_text())

    def test_cli_without_key_finishes_with_clear_limitation(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'input.xlsx'
            source.write_bytes(base64.b64decode(school_upload()['data']))
            stderr = StringIO()
            with patch.dict(os.environ, {'DADATA_API_KEY': ''}), patch(
                    'ou_processor.full_run.DaDataClient', side_effect=AssertionError('Ключ не настроен')), redirect_stdout(StringIO()), redirect_stderr(stderr):
                code = main([str(source), '--output', str(root / 'result'),
                             '--registry', str(root / 'registry'), '--licenses', str(root / 'licenses')])
            self.assertEqual(code, 0)
            self.assertIn('ключ', stderr.getvalue().lower())
            self.assertIn('DaData', stderr.getvalue())
            self.assertIn('локальн', stderr.getvalue().lower())
