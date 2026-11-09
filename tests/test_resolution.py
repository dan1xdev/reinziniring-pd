"""Проверки идентификации и сохранности связей учеников с делами."""
import tempfile
import unittest
import json
from pathlib import Path
from ou_processor.resolution import choose, SearchCache, resolve
from ou_processor.full_run import aggregate
from ou_processor.pipeline import process
from ou_processor.web import JobCancelled


def record(name='ГБОУ школа № 1', city='Москва', region='Москва', flags=None, **extra):
    return dict(name=name, city=city, region=region, flags=[] if flags is None else flags, **extra)


def candidate(inn='123', name='ГБОУ школа № 1', legal_status='ACTIVE', region='г Москва', city='г Москва', address='г Москва, ул А, д 1', **extra):
    return dict(candidate_inn=inn, candidate_name=name, candidate_ogrn='456',
        postal_code='123456', region=region, city=city, address=address,
        legal_status=legal_status, branch_type='MAIN', source='DaData', **extra)


class ResolutionTests(unittest.TestCase):
    def test_different_number_not_selected(self):
        selected,_ = choose(record(),[candidate(name='ГБОУ школа № 11')])
        self.assertIsNone(selected)

    def test_different_city_not_selected(self):
        selected,_ = choose(record(city='Калуга',region='Калужская область'),[candidate()])
        self.assertIsNone(selected)

    def test_generic_number_without_city_not_selected(self):
        selected,_ = choose(record(city=''),[candidate()])
        self.assertIsNone(selected)

    def test_several_inns_require_review(self):
        selected,_ = choose(record(),[candidate('123'),candidate('789')])
        self.assertIsNone(selected)

    def test_same_inn_different_spellings_merge_but_students_stay_paired(self):
        records = [dict(record_id=str(i),student_name=name,personal_file_number=file,name=label)
            for i,(name,file,label) in enumerate([
                ('Иванов Иван','0001','Школа 1'),('Иванов Иван','0001','СОШ 1'),
                ('Петров Пётр','0002','Школа 1'),('Иванов Иван','0003','Школа 1')])]
        groups = [dict(group_id=str(i),record_ids=[str(i)],selected=candidate(),status='СОПОСТАВЛЕНО',registry_used=False) for i in range(4)]
        result = dict(records=records,summary={'selected_rows':4},resolution=dict(groups=groups,actual_requests=0,requests=[]))
        aggregate(result)
        self.assertEqual(len(result['schools']),1)
        self.assertEqual(len(result['students']),3)
        self.assertEqual(result['schools'][0]['row_count'],4)
        self.assertEqual({s['file'] for s in result['students']},{'0001','0002','0003'})

    def test_cache_does_not_reuse_errors(self):
        from ou_processor.dadata import DaDataError
        class Client:
            def suggest(self,query):
                raise DaDataError('ОШИБКА_СЕТИ')
        with tempfile.TemporaryDirectory() as root:
            cache=SearchCache(Path(root)/'api.json',Client())
            cache.search('Школа Москва','g')
            self.assertEqual(cache.entries,{})
            self.assertEqual(cache.requests[0]['status'],'ОШИБКА_СЕТИ')

    # Реестр подтверждает учреждение, затем DaData проверяет действующее юридическое лицо по ИНН.
    def test_license_registry_is_primary_source(self):
        calls = []
        class API:
            actual_requests = 1
            requests = []
            def search(self, query, group_id):
                calls.append(('DaData', query))
                return [candidate()]
        class Licenses:
            metadata = {}
            def search(self, source_record):
                calls.append(('Реестр', source_record['record_id']))
                return [{**candidate(), 'source': 'Реестр лицензий'}], False
        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, API(), None, progress=False, licenses=Licenses())
        self.assertEqual(calls, [('Реестр', '1'), ('DaData', '123')])
        group = result['resolution']['groups'][0]
        self.assertTrue(group['license_used'])
        self.assertFalse(group['registry_used'])
        self.assertEqual(group['status'], 'АВТО')
        self.assertEqual(group['selected']['source'], 'Реестр лицензий')
        self.assertEqual(group['selected']['legal_status_source'], 'DaData')

    # Если в реестре лицензий совпадения нет, подключается DaData.
    def test_license_miss_falls_back_to_dadata(self):
        class API:
            actual_requests = 1
            requests = []
            def search(self, *args):
                return [candidate()]
        class Licenses:
            metadata = {}
            def search(self, record):
                return [], False
        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, API(), None, progress=False, licenses=Licenses())
        group = result['resolution']['groups'][0]
        self.assertTrue(group['license_used'])
        self.assertIsNone(group['selected'])
        self.assertEqual(group['status'], 'ПРОВЕРИТЬ')
        self.assertEqual(group['review_candidate']['candidate_inn'], '123')

    # Без ключа DaData отсутствие записи в обоих реестрах оставляет группу на проверке.
    def test_missing_license_with_no_dadata_client_finishes_without_requests(self):
        class Licenses:
            metadata = {}

            def search(self, source_record):
                return [], False

        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, None, None, progress=False, licenses=Licenses())
        self.assertEqual(result['resolution']['groups'][0]['status'], 'НЕ_НАЙДЕНО')
        self.assertEqual(result['resolution']['groups'][0]['decision_reason'], 'НЕ_НАЙДЕНО')
        self.assertEqual(result['resolution']['actual_requests'], 0)
        self.assertEqual(result['resolution']['requests'], [])

    # Неоднозначность основного реестра нельзя снять посторонним ИНН из DaData.
    def test_dadata_cannot_override_ambiguous_license_with_another_inn(self):
        class Licenses:
            metadata = {}

            def search(self, source_record):
                return [{**candidate(inn='111'), 'source': 'Реестр лицензий'},
                        {**candidate(inn='222'), 'source': 'Реестр лицензий'}], False

        class API:
            actual_requests = 1
            requests = []

            def search(self, query, group_id):
                return [candidate(inn='333')]

        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, API(), None, progress=False, licenses=Licenses())
        self.assertIsNone(result['resolution']['groups'][0]['selected'])
        self.assertEqual(result['resolution']['groups'][0]['decision_reason'], 'КОНФЛИКТ_ИНН')

    # Совпадение двух вспомогательных источников не должно скрыть конфликт с основным.
    def test_secondary_sources_cannot_hide_ambiguous_license_inns(self):
        class Licenses:
            metadata = {}

            def search(self, source_record):
                return [{**candidate(inn='111'), 'source': 'Реестр лицензий'},
                        {**candidate(inn='222'), 'source': 'Реестр лицензий'}], False

        class API:
            actual_requests = 1
            requests = []

            def search(self, query, group_id):
                return [candidate(inn='333')]

        class Registry:
            metadata = {}

            def search(self, source_record):
                return [{**candidate(inn='333'), 'source': 'Реестр аккредитации'}], False

        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, API(), Registry(), progress=False, licenses=Licenses())
        self.assertIsNone(result['resolution']['groups'][0]['selected'])

    def test_successful_license_skips_optional_registry(self):
        class API:
            actual_requests = 1
            requests = []
            def search(self, query, group_id):
                return [candidate()]
        class Licenses:
            metadata = {}
            def search(self, source_record):
                return [{**candidate(), 'source': 'Реестр лицензий'}], False
        class Registry:
            metadata = {}
            def search(self, source_record):
                raise AssertionError('После совпадения лицензии резерв не нужен')
        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, API(), Registry(), progress=False, licenses=Licenses())
        group = result['resolution']['groups'][0]
        self.assertFalse(group['registry_used'])
        self.assertEqual(group['selected']['source'], 'Реестр лицензий')

    def test_registry_match_without_dadata_stays_for_review(self):
        class Registry:
            metadata = {}
            def search(self, source_record):
                return [{**candidate(), 'source': 'Реестр аккредитации'}], False
        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, None, Registry(), progress=False)
        group = result['resolution']['groups'][0]
        self.assertTrue(group['registry_used'])
        self.assertIsNone(group['selected'])
        self.assertEqual(group['status'], 'ПРОВЕРИТЬ')
        self.assertEqual(group['decision_reason'], 'ЮРИДИЧЕСКИЙ_СТАТУС_НЕ_ПОДТВЕРЖДЁН')
        self.assertEqual(group['review_candidate']['candidate_inn'], '123')
        self.assertEqual(result['resolution']['requests'], [])

    def test_without_registry_keeps_unmatched_school_and_student_for_review(self):
        class API:
            actual_requests=1
            requests=[]
            def search(self,*args): return []
        r=record(record_id='1',student_name='Иванов Иван',personal_file_number='0001')
        result={'records':[r],'groups':[dict(group_id='g',record_ids=['1'])],
                'summary':{'selected_rows':1}}
        resolve(result,API(),None,progress=False)
        aggregate(result)
        self.assertEqual(result['resolution']['registry'],{})
        self.assertFalse(result['resolution']['groups'][0]['registry_used'])
        self.assertEqual(result['review'][0]['students'],'Иванов Иван')
        self.assertEqual(result['review'][0]['files'],'0001')

    def test_without_registry_dadata_result_is_only_review_proposal(self):
        class API:
            actual_requests = 1
            requests = []
            def search(self, query, group_id):
                return [candidate()]
        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, API(), None, progress=False)
        group = result['resolution']['groups'][0]
        self.assertIsNone(group['selected'])
        self.assertEqual(group['status'], 'ПРОВЕРИТЬ')
        self.assertEqual(group['review_candidate']['candidate_inn'], '123')
        self.assertEqual(group['review_candidate']['source'], 'DaData')

    def test_cancel_after_api_reply_skips_following_queries_and_groups(self):
        calls = []
        class API:
            def search(self, query, group_id):
                calls.append(('DaData', group_id))
                return []
        class Registry:
            metadata = {}
            def search(self, source_record):
                calls.append(('Реестр', source_record['record_id']))
                return [], False
        result = {'records': [record(record_id='1'), record(record_id='2')],
                  'groups': [dict(group_id='g1', record_ids=['1']), dict(group_id='g2', record_ids=['2'])]}
        def check_cancel():
            if any(source == 'DaData' for source, identifier in calls):
                raise JobCancelled()
        with self.assertRaises(JobCancelled):
            resolve(result, API(), Registry(), progress=False, check_cancel=check_cancel)
        self.assertEqual(calls, [('Реестр', '1'), ('DaData', 'g1')])

    def test_music_school_is_not_general_school(self):
        selected,_=choose(record(city='Можайск',region='Московская область'),[
            {**candidate(name='Детская музыкальная школа № 1'), 'region':'Московская обл',
             'city':'Можайск', 'address':'Можайск, Московская 17'}])
        self.assertIsNone(selected)

    def test_street_name_cannot_confirm_region(self):
        selected,_=choose(record(city='',region='Московская область'),[
            {**candidate(),'region':'Калужская обл','address':'г Калуга, Московская область проезд'}])
        self.assertIsNone(selected)

    def test_candidate_audit_keeps_distinct_group_ids(self):
        class API:
            actual_requests = 2
            requests = []
            def search(self, query, group_id):
                return [{**candidate(), 'group_id': 'чужая группа'}]
        class Licenses:
            metadata = {}
            def search(self, source_record):
                return [], False
        result = {'records': [record(record_id='1'), record(record_id='2')],
                  'groups': [dict(group_id='g1', record_ids=['1']), dict(group_id='g2', record_ids=['2'])]}
        resolve(result, API(), None, progress=False, licenses=Licenses())
        self.assertEqual([(item['group_id'], item['candidate_inn'])
                          for item in result['resolution']['candidates']], [('g1', '123'), ('g2', '123')])

    def test_source_without_student_column_is_retained(self):
        r=dict(record_id='1',name='Школа',personal_file_number='')
        result=dict(records=[r],summary={'selected_rows':1},resolution=dict(groups=[dict(
            group_id='g',record_ids=['1'],selected=candidate(),status='СОПОСТАВЛЕНО',registry_used=False)],
            actual_requests=0,requests=[]))
        aggregate(result)
        self.assertEqual(result['schools'][0]['student_count'],0)
        self.assertEqual(result['full_summary']['Строк без данных ученика'],1)

    def test_region_inflection_does_not_become_identity(self):
        from ou_processor.resolution import meaningful
        self.assertEqual(meaningful(record(name='ГАПОУ Московской области Губернский колледж',city='',region='Московская область')),{'губернский'})

    def test_trade_union_is_not_college(self):
        selected,_=choose(record(name='Московский колледж',city='Москва'),[
            candidate(name='Первичная профсоюзная организация Московский колледж')])
        self.assertIsNone(selected)

    def test_exact_title_beats_extended_title(self):
        r=record(name='ГАПОУ "Губернский колледж"',city='',region='Московская область')
        base={**candidate(),'region':'Московская обл','source':'Реестр аккредитации','legal_status':'Действующее'}
        exact={**base,'candidate_name':'ГАПОУ "Губернский колледж"','scope':'actual_organization'}
        other={**base,'candidate_inn':'789','candidate_name':'ГАПОУ "Московский губернский колледж искусств"','scope':'actual_organization'}
        selected,_=choose(r,[other,exact])
        self.assertEqual(selected['candidate_inn'],'123')
        self.assertEqual(selected['candidate_name'],exact['candidate_name'])

    def test_candidate_margin_triggers_review_when_small_difference(self):
        # Два кандидата с разными ИНН, один 95 баллов, второй 90 баллов (разница 5 < 15)
        r = record(name='Школа № 5 г. Москвы', city='Москва', region='Москва')
        c1 = candidate(inn='111', name='ГБОУ Школа № 5 города Москвы')
        c2 = candidate(inn='222', name='ГБОУ Школа № 5 имени Ломоносова города Москвы')
        selected, ranked = choose(r, [c1, c2])
        self.assertIsNone(selected)
        c1_ranked = next(c for c in ranked if c['candidate_inn'] == '111')
        self.assertIn('НЕДОСТАТОЧНЫЙ_ОТРЫВ_КАНДИДАТОВ', c1_ranked['reasons'])

    # Три и более равных кандидата не позволяют выбрать оставшегося после отказа лидерам.
    def test_insufficient_margin_never_selects_lower_ranked_candidate(self):
        for count in (3, 4):
            with self.subTest(candidate_count=count):
                candidates = [candidate(inn=str(index)) for index in range(count)]
                selected, ranked = choose(record(), candidates)
                self.assertIsNone(selected)
                self.assertFalse(any(c['eligible'] for c in ranked))
                self.assertTrue(all('НЕДОСТАТОЧНЫЙ_ОТРЫВ_КАНДИДАТОВ' in c['reasons'] for c in ranked))

    # Посторонние ИНН из DaData отменяют выбранное совпадение дополнительного реестра.
    def test_dadata_conflict_cancels_reserve_match(self):
        class API:
            actual_requests = 0
            requests = []

            def search(self, *args):
                return [candidate(inn='111'), candidate(inn='222')]

        class Registry:
            metadata = {}

            def search(self, source_record):
                return [{**candidate(inn='333'), 'source': 'Реестр аккредитации'}], False

        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, API(), Registry(), progress=False)
        self.assertIsNone(result['resolution']['groups'][0]['selected'])
        self.assertEqual(result['resolution']['groups'][0]['decision_reason'], 'КОНФЛИКТ_ИНН')

    # Старый выпуск остаётся на проверке независимо от положения рядом с новым выпуском.
    def test_old_graduation_is_not_confirmed_by_another_row_in_group(self):
        config = json.loads((Path(__file__).resolve().parents[1] / 'config.json').read_text(encoding='utf-8'))

        class API:
            actual_requests = 0
            requests = []

            def search(self, *args):
                return [candidate()]

        class Licenses:
            metadata = {}

            def search(self, source_record):
                return [{**candidate(), 'source': 'Реестр лицензий'}], False

        for years in [('2025', '2020'), ('2020', '2025')]:
            with self.subTest(years=years):
                records = [dict(record_id=year, name='ГБОУ школа № 1', education='Среднее общее',
                    city_raw='Москва', region_raw='Москва', year=year, formula_fields=[],
                    student_name='Ученик ' + year, personal_file_number=year) for year in years]
                result = process(records, config)
                resolve(result, API(), None, progress=False, config=config, licenses=Licenses())
                aggregate(result)
                by_year = {r['year']: r for r in result['records']}
                self.assertEqual(by_year['2020']['resolution_status'], 'ВРУЧНАЯ')
                self.assertEqual(by_year['2020']['resolved_inn'], '')
                self.assertEqual(by_year['2025']['resolution_status'], 'АВТО')
                self.assertEqual(result['schools'][0]['record_ids'], '2025')
                self.assertEqual(result['review'][0]['record_ids'], '2020')

    def test_truncated_name_vetoes_auto_match(self):
        r = record(name='ГБОУ города Москвы «Школа № 1284', flags=['ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ'])
        c = candidate(inn='123', name='ГБОУ города Москвы «Школа № 1284»')
        selected, ranked = choose(r, [c])
        self.assertIsNone(selected)
        self.assertIn('ОБРЕЗАННОЕ_НАЗВАНИЕ', ranked[0]['reasons'])

    def test_old_graduation_year_vetoes_auto_match(self):
        r = record(name='ГБОУ города Москвы «Школа № 1284»', flags=['СТАРЫЙ_ВЫПУСК'])
        c = candidate(inn='123', name='ГБОУ города Москвы «Школа № 1284»')
        selected, ranked = choose(r, [c])
        self.assertIsNone(selected)
        self.assertIn('НЕПОДТВЕРЖДЕННОЕ_ПРАВОПРЕЕМСТВО', ranked[0]['reasons'])
        self.assertIn('СТАРЫЙ_ВЫПУСК_ТРЕБУЕТ_ПРОВЕРКИ', ranked[0]['reasons'])

    def test_region_conflict_vetoes_auto_match(self):
        r = record(name='ГБОУ города Москвы «Школа № 1284»', city='Москва', region='Москва', flags=['КОНФЛИКТ_РЕГИОНА'])
        c = candidate(inn='123', name='ГБОУ города Москвы «Школа № 1284»', region='г Москва', city='г Москва')
        selected, ranked = choose(r, [c])
        self.assertIsNone(selected)
        self.assertIn('КОНФЛИКТ_РЕГИОНА', ranked[0]['reasons'])

    def test_liquidated_entity_vetoes_auto_match(self):
        r = record(name='ГБОУ города Москвы «Школа № 1284»', city='Москва', region='Москва')
        c = candidate(inn='123', name='ГБОУ города Москвы «Школа № 1284»', legal_status='LIQUIDATED')
        selected, ranked = choose(r, [c])
        self.assertIsNone(selected)
        self.assertIn('НЕПОДТВЕРЖДЕННОЕ_ПРАВОПРЕЕМСТВО', ranked[0]['reasons'])

    def test_resolve_status_assignment(self):
        class API:
            actual_requests = 1
            requests = []
            def search(self, *args):
                return [candidate(inn='123', name='ГБОУ города Москвы «Школа № 1284»')]

        # Проверка статуса при обрезке
        r1 = record(record_id='1', name='ГБОУ города Москвы «Школа № 1284', flags=['ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ'])
        res1 = {'records': [r1], 'groups': [dict(group_id='g1', record_ids=['1'])]}
        resolve(res1, API(), None, progress=False)
        self.assertEqual(res1['resolution']['groups'][0]['decision_reason'], 'ПРОВЕРИТЬ_ОБРЕЗКА')

        # Проверка статуса при конфликте региона
        r2 = record(record_id='2', name='ГБОУ города Москвы «Школа № 1284»', flags=['КОНФЛИКТ_РЕГИОНА'])
        res2 = {'records': [r2], 'groups': [dict(group_id='g2', record_ids=['2'])]}
        resolve(res2, API(), None, progress=False)
        self.assertEqual(res2['resolution']['groups'][0]['decision_reason'], 'ПРОВЕРИТЬ_КОНФЛИКТ_РЕГИОНА')

        # Проверка статуса при старом выпуске
        r3 = record(record_id='3', name='ГБОУ города Москвы «Школа № 1284»', flags=['СТАРЫЙ_ВЫПУСК'])
        res3 = {'records': [r3], 'groups': [dict(group_id='g3', record_ids=['3'])]}
        resolve(res3, API(), None, progress=False)
        self.assertEqual(res3['resolution']['groups'][0]['decision_reason'], 'ПРОВЕРИТЬ_ПРАВОПРЕЕМСТВО')

    # Несовпадение региона кандидата получает специальный статус проверки региона.
    def test_candidate_region_mismatch_has_region_review_status(self):
        class API:
            actual_requests = 0
            requests = []

            def search(self, *args):
                return [candidate(region='Калужская область')]

        result = {'records': [record(record_id='1')],
                  'groups': [dict(group_id='g', record_ids=['1'])]}
        resolve(result, API(), None, progress=False)
        self.assertEqual(result['resolution']['groups'][0]['decision_reason'], 'ПРОВЕРИТЬ_КОНФЛИКТ_РЕГИОНА')
        self.assertIsNone(result['resolution']['groups'][0]['selected'])

    def test_user_examples_match_across_all_variations(self):
        # 10 организаций из реального примера пользователя во всех 4 вариантах написания
        from ou_processor.geography import extract_geography
        from ou_processor.normalization import normalize, numbers, institution_type
        import json

        config = json.loads((Path(__file__).resolve().parents[1] / "config.json").read_text(encoding="utf-8"))

        test_cases = [
            (
                "Школа 1284",
                ("1284",),
                "ШКОЛА",
                "Москва",
                [
                    "Государственное бюджетное общеобразовательное учреждение города Москвы «Школа № 1284»",
                    "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ГОРОДА МОСКВЫ «ШКОЛА № 1284»",
                    "ГБОУ г. Москвы «Школа № 1284»",
                    'ГБОУ города Москвы ""Школа №1284""',
                ]
            ),
            (
                "Школа на Яузе",
                (),
                "ШКОЛА",
                "Москва",
                [
                    "Государственное бюджетное общеобразовательное учреждение города Москвы «Школа на Яузе»",
                    "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ГОРОДА МОСКВЫ «ШКОЛА НА ЯУЗЕ»",
                    "ГБОУ г. Москвы «Школа на Яузе»",
                    "школа на яузе, г. москва",
                ]
            ),
            (
                "Колледж 20",
                ("20",),
                "СПО",
                "Москва",
                [
                    "Государственное бюджетное профессиональное образовательное учреждение города Москвы «Колледж автоматизации и информационных технологий № 20»",
                    "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ГОРОДА МОСКВЫ «КОЛЛЕДЖ АВТОМАТИЗАЦИИ И ИНФОРМАЦИОННЫХ ТЕХНОЛОГИЙ № 20»",
                    "ГБПОУ г. Москвы «Колледж автоматизации и информационных технологий № 20»",
                    'ГБПОУ Г. МОСКВЫ ""КОЛЛЕДЖ АВТОМАТИЗАЦИИ И ИНФОРМАЦИОННЫХ ТЕХНОЛОГИЙ №20""',
                ]
            ),
            (
                "Школа 1508",
                ("1508",),
                "ШКОЛА",
                "Москва",
                [
                    "Государственное бюджетное общеобразовательное учреждение города Москвы «Измайловская школа № 1508»",
                    "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ГОРОДА МОСКВЫ «ИЗМАЙЛОВСКАЯ ШКОЛА № 1508»",
                    "ГБОУ г. Москвы «Измайловская школа № 1508»",
                    "  ГБОУ г. Москвы   «Измайловская школа № 1508»  ",
                ]
            ),
            (
                "Школа 1290",
                ("1290",),
                "ШКОЛА",
                "Москва",
                [
                    "Государственное бюджетное общеобразовательное учреждение города Москвы «Школа № 1290»",
                    "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ГОРОДА МОСКВЫ «ШКОЛА № 1290»",
                    "ГБОУ г. Москвы «Школа № 1290»",
                    "Государственное бюджетное общеобразовательное учреждение города Москвы Школа №1290",
                ]
            ),
            (
                "Школа 444",
                ("444",),
                "ШКОЛА",
                "Москва",
                [
                    "Государственное бюджетное общеобразовательное учреждение города Москвы «Школа № 444»",
                    "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ГОРОДА МОСКВЫ «ШКОЛА № 444»",
                    "ГБОУ г. Москвы «Школа № 444»",
                    "ГБОУ Москва школа с углубленным изучением математики, информатики, физики № 444",
                ]
            ),
            (
                "Покровский квартал",
                (),
                "ШКОЛА",
                "Москва",
                [
                    "Государственное бюджетное общеобразовательное учреждение города Москвы «Школа «Покровский квартал»»",
                    "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ГОРОДА МОСКВЫ «ШКОЛА «ПОКРОВСКИЙ КВАРТАЛ»»",
                    "ГБОУ г. Москвы «Школа «Покровский квартал»»",
                    'ГБОУ города Москвы ""Школа ""Покровский квартал"""',
                ]
            ),
            (
                "Колледж Николаева",
                (),
                "СПО",
                "Москва",
                [
                    "Государственное бюджетное профессиональное образовательное учреждение города Москвы «Московский автомобильно-дорожный колледж имени А. А. Николаева»",
                    "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ГОРОДА МОСКВЫ «МОСКОВСКИЙ АВТОМОБИЛЬНО-ДОРОЖНЫЙ КОЛЛЕДЖ ИМЕНИ А. А. НИКОЛАЕВА»",
                    "ГБПОУ г. Москвы «Московский автомобильно-дорожный колледж им. А. А. Николаева»",
                    "г. Москва Московский автомобильно-дорожный колледж им. А.А. Николаева",
                ]
            ),
            (
                "Школа 47 Калуга",
                ("47",),
                "ШКОЛА",
                "Калуга",
                [
                    "Муниципальное бюджетное общеобразовательное учреждение «Средняя общеобразовательная школа № 47» города Калуги",
                    "МУНИЦИПАЛЬНОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ «СРЕДНЯЯ ОБЩЕОБРАЗОВАТЕЛЬНАЯ ШКОЛА № 47» ГОРОДА КАЛУГИ",
                    "МБОУ СОШ № 47 г. Калуги",
                    'МБОУ ""СОШ №47"" г Калуга',
                ]
            ),
            (
                "Школа 1 Можайск",
                ("1",),
                "ШКОЛА",
                "Можайск",
                [
                    "Муниципальное общеобразовательное учреждение «Средняя общеобразовательная школа № 1» г. Можайска",
                    "МУНИЦИПАЛЬНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ «СРЕДНЯЯ ОБЩЕОБРАЗОВАТЕЛЬНАЯ ШКОЛА № 1» Г. МОЖАЙСКА",
                    "МОУ СОШ № 1 г. Можайска",
                    "МОУ Средняя общеобразовательная школа №1 г Можайск",
                ]
            ),
        ]

        for desc, expected_nums, expected_type, expected_city, variations in test_cases:
            for variant in variations:
                extracted_nums = numbers(variant)
                self.assertEqual(extracted_nums, expected_nums, f"{desc}: {variant} numbers mismatch")
                extracted_type = institution_type(variant, config)
                self.assertEqual(extracted_type, expected_type, f"{desc}: {variant} type mismatch")
                geo = extract_geography({"name": variant}, config)
                self.assertEqual(geo["city"], expected_city, f"{desc}: {variant} city mismatch")
                self.assertNotIn("КОНФЛИКТ_РЕГИОНА", geo["geo_flags"])
                self.assertNotIn("КОНФЛИКТ_ГОРОДА", geo["geo_flags"])
