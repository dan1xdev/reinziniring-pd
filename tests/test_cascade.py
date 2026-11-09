"""Контракты каскада ТЗ: идентичность, происхождение полей и отсутствие ПДн."""
import unittest

from ou_processor.resolution import resolve
from tests.test_resolution import record, candidate


class Index:
    metadata = {}

    def __init__(self, rows):
        self.rows = rows

    def search(self, record):
        return self.rows, False


class API:
    actual_requests = 0

    def __init__(self, rows):
        self.rows, self.requests = rows, []

    def search(self, query, group):
        self.requests.append({'query': query, 'group_id': group, 'status': 'OK'})
        return self.rows


class CascadeTests(unittest.TestCase):
    def run_cascade(self, official, auxiliary):
        result = {'records': [record(record_id='1', student_name='Не отправлять')],
                  'groups': [{'group_id': 'g', 'record_ids': ['1']}]}
        api = API(auxiliary)
        resolve(result, api, None, licenses=Index(official), progress=False)
        self.assertNotIn('Не отправлять', str(api.requests))
        return result, api

    def test_address_enrichment_keeps_official_name_and_license(self):
        official = {**candidate(), 'source': 'Реестр лицензий', 'education_levels': ['Среднее общее образование']}
        auxiliary = candidate(name='Другое краткое название', address='г Москва, Новый адрес', candidate_kpp='77')
        result, api = self.run_cascade([official], [auxiliary])
        selected = result['resolution']['groups'][0]['selected']
        self.assertEqual(selected['candidate_name'], official['candidate_name'])
        self.assertEqual(selected['education_levels'], official['education_levels'])
        self.assertEqual(selected['address'], auxiliary['address'])
        self.assertEqual(selected['address_source'], 'DaData')
        self.assertEqual(api.requests[0]['query'], '123')

    def test_conflicting_inn_revokes_automatic_selection(self):
        result, _ = self.run_cascade([{**candidate(), 'source': 'Реестр лицензий'}], [candidate(inn='999')])
        group = result['resolution']['groups'][0]
        self.assertIsNone(group['selected'])
        self.assertEqual(group['status'], 'ВРУЧНАЯ')
        self.assertEqual(group['decision_reason'], 'КОНФЛИКТ_ИНН')

    def test_dadata_alone_is_review_proposal_not_official_confirmation(self):
        result, api = self.run_cascade([], [candidate()])
        group = result['resolution']['groups'][0]
        self.assertIsNone(group['selected'])
        self.assertEqual(group['status'], 'ПРОВЕРИТЬ')
        self.assertEqual(group['review_candidate']['candidate_inn'], '123')
        self.assertTrue(api.requests)

    def test_liquidated_legal_entity_cannot_use_active_license(self):
        result, _ = self.run_cascade([{**candidate(), 'source': 'Реестр лицензий'}], [candidate(legal_status='LIQUIDATED')])
        self.assertIsNone(result['resolution']['groups'][0]['selected'])
        self.assertEqual(result['resolution']['groups'][0]['status'], 'ВРУЧНАЯ')

    def test_missing_api_confirmation_is_review(self):
        for api in (None, API([])):
            with self.subTest(api=api):
                result = {'records': [record(record_id='1')], 'groups': [{'group_id': 'g', 'record_ids': ['1']}]}
                resolve(result, api, None, licenses=Index([{**candidate(), 'source': 'Реестр лицензий'}]), progress=False)
                group = result['resolution']['groups'][0]
                self.assertIsNone(group['selected'])
                self.assertEqual(group['status'], 'ПРОВЕРИТЬ')
                self.assertEqual(group['review_candidate']['candidate_inn'], '123')

    def test_different_address_region_revokes_selection(self):
        extra = {**candidate(address='г Тверь, улица Новая', region='Тверская область', city='Тверь'), 'postal_code': ''}
        result, _ = self.run_cascade([{**candidate(), 'source': 'Реестр лицензий'}], [extra])
        self.assertIsNone(result['resolution']['groups'][0]['selected'])
        self.assertEqual(result['resolution']['groups'][0]['status'], 'ВРУЧНАЯ')

    def test_new_address_without_postcode_never_inherits_old_postcode(self):
        extra = {**candidate(address='г Москва, улица Новая'), 'postal_code': ''}
        result, _ = self.run_cascade([{**candidate(), 'source': 'Реестр лицензий'}], [extra])
        self.assertEqual(result['resolution']['groups'][0]['selected']['postal_code'], '')

    def test_nullable_dadata_address_fields_do_not_abort_the_cascade(self):
        extra = {**candidate(), 'address': None, 'city': None, 'region': None, 'postal_code': None}
        result, _ = self.run_cascade([], [extra])
        self.assertIsNone(result['resolution']['groups'][0]['selected'])
        self.assertEqual(result['resolution']['groups'][0]['status'], 'ВРУЧНАЯ')
