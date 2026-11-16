"""Спорные названия и неизвестный регион запрещают автоматическое сопоставление."""

from copy import deepcopy
import json
from pathlib import Path
import unittest

from ou_processor.normalization import abbreviation_flags
from ou_processor.pipeline import process
from ou_processor.resolution import resolve


CONFIG = json.loads((Path(__file__).resolve().parents[1] / 'config.json').read_text(encoding='utf-8'))


def resolved(name, *, region='Москва', city='Москва', config=None):
    config = config or CONFIG
    result = process([{
        'record_id': '1', 'name': name, 'region_raw': region, 'city_raw': city,
        'education': 'Среднее общее', 'student_name': 'Тестовый Ученик',
        'personal_file_number': '0001', 'formula_fields': [],
    }], config)
    candidate = {
        'candidate_name': name, 'candidate_inn': '1111111111', 'candidate_ogrn': '1111111111111',
        'address': f'123456, {region or "Тверская область"}, г {city}, улица Тестовая, д 1',
        'postal_code': '123456', 'region': region or 'Тверская область', 'city': city,
        'legal_status': 'Действующее', 'branch_type': 'MAIN', 'source': 'Реестр лицензий',
    }

    class Licenses:
        metadata = {}

        def search(self, record):
            return [candidate], False

    class API:
        actual_requests = 0
        requests = []

        def search(self, query, group_id):
            return [{**candidate, 'source': 'DaData', 'legal_status': 'ACTIVE'}]

    return resolve(result, API(), None, licenses=Licenses(), config=config, progress=False)


class TZBlockerTests(unittest.TestCase):
    def assert_manual(self, result, flag):
        self.assertIn(flag, result['records'][0]['flags'])
        group = result['resolution']['groups'][0]
        self.assertIsNone(group['selected'])
        self.assertEqual(group['status'], 'ВРУЧНАЯ')

    def test_old_names_stay_manual_even_when_both_sources_match(self):
        for name in ('ГОУ СПО Тестовый техникум № 1', 'ФГОУ СПО Тестовый техникум № 1',
                     'ФГОУСПО Тестовый техникум № 1', 'МАДК Тестовый колледж № 1'):
            with self.subTest(name=name):
                self.assert_manual(resolved(name), 'УСТАРЕВШЕЕ_НАИМЕНОВАНИЕ')

    def test_unknown_abbreviation_stays_manual_even_when_sources_match(self):
        self.assert_manual(resolved('АБВГД школа № 1'), 'НЕИЗВЕСТНАЯ_АББРЕВИАТУРА')
        self.assertIn('НЕИЗВЕСТНАЯ_АББРЕВИАТУРА', abbreviation_flags('АБВГД', CONFIG))

    def test_missing_region_cannot_be_replaced_by_candidate_confidence(self):
        self.assert_manual(resolved('Школа № 1', region='', city='Неизвестноград'), 'РЕГИОН_НЕ_ОПРЕДЕЛЕН')

    def test_known_abbreviations_and_uppercase_full_names_are_not_unknown(self):
        for name in ('ГБОУ школа № 1', 'ГБПОУ колледж № 1', 'ФГБОУ СПО техникум № 1',
                     'СРЕДНЯЯ ШКОЛА ГОРОДА МОСКВЫ № 1', 'СРЕДНЯЯ ШКОЛА «СОСНЫ»',
                     'Школа № 1 имени А. С. Пушкина'):
            with self.subTest(name=name):
                flags = abbreviation_flags(name, CONFIG)
                self.assertNotIn('НЕИЗВЕСТНАЯ_АББРЕВИАТУРА', flags)
                self.assertNotIn('УСТАРЕВШЕЕ_НАИМЕНОВАНИЕ', flags)
        result = resolved('ГБОУ школа № 1')
        self.assertEqual(result['resolution']['groups'][0]['status'], 'АВТО')

    def test_unknown_abbreviation_exceptions_are_configurable(self):
        config = deepcopy(CONFIG)
        config.setdefault('abbreviation_exceptions', []).append('АБВГД')
        self.assertNotIn('НЕИЗВЕСТНАЯ_АББРЕВИАТУРА', abbreviation_flags('АБВГД школа № 1', config))

    def test_additional_old_names_are_configurable(self):
        config = deepcopy(CONFIG)
        config.setdefault('old_name_patterns', []).append(r'\bпрежний\s+колледж\b')
        self.assert_manual(resolved('Прежний колледж № 1', config=config), 'УСТАРЕВШЕЕ_НАИМЕНОВАНИЕ')


if __name__ == '__main__':
    unittest.main()
