"""Проверки готовности благодарственных писем после объединения по ИНН."""

import unittest

from ou_processor.full_run import aggregate
from ou_processor.letters import prepare_letters


def letter_result(people, **candidate_fields):
    candidate = dict(
        candidate_inn='7707083893',
        candidate_name='Средняя общеобразовательная школа № 1',
        candidate_ogrn='', postal_code='123456', region='Москва',
        address='г Москва, ул А, д 1', source='DaData',
    )
    candidate.update(candidate_fields)
    records = [dict(
        record_id=str(index), name='Школа № 1',
        student_name=name, personal_file_number=personal_file,
    ) for index, (name, personal_file) in enumerate(people)]
    result = dict(
        records=records, summary={'selected_rows': len(records)},
        resolution=dict(
            groups=[dict(
                group_id='g', selected=candidate,
                record_ids=[record['record_id'] for record in records],
                status='СОПОСТАВЛЕНО', registry_used=False,
            )],
            actual_requests=0, requests=[],
        ),
    )
    aggregate(result)
    prepare_letters(result)
    return result


class LetterTests(unittest.TestCase):
    # Противоречивые ФИО одного дела нельзя печатать как двух поступивших.
    def test_conflicting_names_for_same_file_require_review(self):
        result = letter_result([
            ('Иванов Иван Иванович', '0001'),
            ('Иванов Игорь Иванович', '0001'),
        ])
        self.assertEqual(result['letters_ready'], [])
        self.assertEqual(len(result['letters_review']), 1)
        self.assertIn('Один номер дела связан с разными ФИО',
                      result['letters_review'][0]['checks'])

    # Одиночное письмо сохраняет исходное ФИО без требования склонить его или определить пол.
    def test_single_person_uses_neutral_template(self):
        for name in ('Иванов Иван Иванович', 'Петрова Анна Ивановна'):
            with self.subTest(name=name):
                result = letter_result([(name, '0001')])
                self.assertEqual(result['letters_review'], [])
                letter, = result['letters_ready']
                self.assertEqual(letter['student_count'], 1)
                self.assertEqual(letter['letter_type'], 'Для одного поступившего')
                self.assertEqual(letter['letter_fragment'], (
                    'Благодарим коллектив Средней общеобразовательной школы № 1 '
                    'за подготовку к поступлению.\n'
                    f'ФИО: {name}.'
                ))

    # Неизвестное прилагательное требует проверки, даже если его окончание похоже на известное.
    def test_unknown_adjectives_do_not_become_verified_genitives(self):
        for name in ('Хорошая школа № 1', 'Высшая школа № 1', 'Зелёная школа № 1'):
            with self.subTest(name=name):
                result = letter_result([('Иванов Иван Иванович', '0001')],
                                       candidate_name=name)
                self.assertEqual(result['letters_ready'], [])
                letter, = result['letters_review']
                self.assertEqual(letter['genitive_name'], '')
                self.assertIn('Родительный падеж требует проверки', letter['checks'])

    # Без шестизначного индекса или при расхождении двух индексов письмо не готово к печати.
    def test_missing_invalid_or_conflicting_postal_code_requires_review(self):
        for code, address in (
            ('', 'г Москва, ул А, д 1'),
            ('12345', 'г Москва, ул А, д 1'),
            ('123456', '654321, г Москва, ул А, д 1'),
        ):
            with self.subTest(code=code, address=address):
                result = letter_result([('Иванов Иван Иванович', '0001')],
                                       postal_code=code, address=address)
                self.assertEqual(result['letters_ready'], [])
                letter, = result['letters_review']
                self.assertEqual(letter['postal_code'], '')
                self.assertEqual(letter['mailing_address'], '')
                self.assertIn('индекс', letter['checks'].lower())

    # Индекс из адреса подтверждённого источника позволяет подготовить письмо.
    def test_postal_code_from_registry_address_is_preserved(self):
        result = letter_result([('Иванов Иван Иванович', '0001')],
                               source='Реестр аккредитации', postal_code='',
                               address='Россия, 123456, г Москва, ул А, д 1')
        letter, = result['letters_ready']
        self.assertEqual(letter['postal_code'], '123456')
        self.assertEqual(letter['mailing_address'], '123456, г Москва, ул А, д 1')

    # Общий ИНН даёт одно письмо для нескольких человек с ФИО по алфавиту и без дубля пары.
    def test_multiple_people_are_sorted_and_repeated_pair_is_not_counted(self):
        result = letter_result([
            ('Петрова Анна Ивановна', '0002'),
            ('Иванов Иван Иванович', '0001'),
            ('Петрова Анна Ивановна', '0002'),
        ])
        self.assertEqual(result['letters_review'], [])
        letter, = result['letters_ready']
        self.assertEqual(letter['student_count'], 2)
        self.assertEqual(letter['letter_type'], 'Для нескольких поступивших')
        self.assertEqual(letter['students'], 'Иванов Иван Иванович\nПетрова Анна Ивановна')
        self.assertIn('поступивших:\nИванов Иван Иванович\nПетрова Анна Ивановна.',
                      letter['letter_fragment'])

    # Полное типовое название склоняется вместе с правовой формой, сохраняя название в кавычках.
    def test_known_legal_name_has_verified_genitive(self):
        result = letter_result([('Иванов Иван Иванович', '0001')], candidate_name=(
            'Муниципальное бюджетное общеобразовательное учреждение «Школа № 1»'
        ))
        letter, = result['letters_ready']
        self.assertEqual(letter['genitive_name'], (
            'Муниципального бюджетного общеобразовательного учреждения «Школа № 1»'
        ))
