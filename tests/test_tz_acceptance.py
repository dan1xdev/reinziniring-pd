"""Проверки выходного контракта ТЗ v2 без ручного импорта решений."""
import unittest
from tests.test_letters import letter_result
from ou_processor.letters import LETTER_COLUMNS
from ou_processor.resolution import acceptance_status

class AcceptanceTests(unittest.TestCase):
    def test_letter_has_print_columns_and_provenance(self):
        result = letter_result([('Петров Пётр', '02'), ('Иванов Иван', '01')], candidate_ogrn='1234567890123')
        letter = (result['letters_ready'] + result['letters_review'])[0]
        self.assertEqual(letter['letter_number'], 1)
        self.assertEqual(letter['ogrn'], '1234567890123')
        self.assertEqual(letter['region'], 'Москва')
        self.assertEqual(letter['source_names'], 'Школа № 1')
        self.assertEqual(letter['students'], 'Иванов Иван, Петров Пётр')
        self.assertEqual([x[0] for x in LETTER_COLUMNS[:8]], ['letter_number','official_name','genitive_name','postal_code','address','region','students','student_count'])

    def test_review_status_never_confirms_a_candidate(self):
        rows=[{'candidate_inn':'1','score':80,'reasons':[]},{'candidate_inn':'2','score':65,'reasons':[]}]
        self.assertEqual(acceptance_status(None, rows, {}, {}), 'ПРОВЕРИТЬ')
        rows[0]['reasons']=['ОБРЕЗАННОЕ_НАЗВАНИЕ']
        self.assertEqual(acceptance_status(None, rows, {}, {}), 'ВРУЧНАЯ')
        self.assertEqual(acceptance_status(None, [], {}, {}), 'НЕ_НАЙДЕНО')

    def test_manual_queue_separates_review_letters_and_keeps_input_conflicts(self):
        from ou_processor.full_run import _manual_review_rows
        result = {'review': [
            {'group_id': 'proposal', 'status': 'ПРОВЕРИТЬ'},
            {'group_id': 'manual', 'status': 'ВРУЧНАЯ'}],
            'records': [{'record_id': 'input', 'selection': 'ПРОВЕРИТЬ_ОТБОР',
                         'selection_reason': 'Конфликт страны', 'name': 'Тестовая школа'}]}
        rows = _manual_review_rows(result)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]['group_id'], 'manual')
        self.assertEqual(rows[1]['kind'], 'Отбор входной строки')
        self.assertEqual(rows[1]['record_ids'], 'input')
