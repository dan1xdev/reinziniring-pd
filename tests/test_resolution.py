"""Проверки идентификации и сохранности связей учеников с делами."""
import tempfile
import unittest
from pathlib import Path
from ou_processor.resolution import choose, SearchCache, resolve
from ou_processor.full_run import aggregate


def record(name='ГБОУ школа № 1', city='Москва', region='Москва', **extra):
    return dict(name=name, city=city, region=region, flags=[], **extra)


def candidate(inn='123', name='ГБОУ школа № 1', **extra):
    return dict(candidate_inn=inn,candidate_name=name, candidate_ogrn='456',
        postal_code='123456',region='г Москва',city='г Москва',address='г Москва, ул А, д 1',
        legal_status='ACTIVE',branch_type='MAIN',source='DaData',**extra)


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

    def test_successful_primary_skips_registry(self):
        class API:
            actual_requests=1
            requests=[]
            def search(self,*args): return [candidate()]
        class Registry:
            metadata={}
            def search(self,r): raise AssertionError('Реестр не нужен')
        r=record(record_id='1')
        result={'records':[r],'groups':[dict(group_id='g',record_ids=['1'])]}
        resolve(result,API(),Registry(),progress=False)
        self.assertFalse(result['resolution']['groups'][0]['registry_used'])

    def test_empty_primary_uses_registry(self):
        class API:
            actual_requests=1
            requests=[]
            def search(self,*args): return []
        class Registry:
            metadata={}
            def search(self,r): return [candidate()],False
        r=record(record_id='1')
        result={'records':[r],'groups':[dict(group_id='g',record_ids=['1'])]}
        resolve(result,API(),Registry(),progress=False)
        self.assertTrue(result['resolution']['groups'][0]['registry_used'])
        self.assertIsNotNone(result['resolution']['groups'][0]['selected'])

    def test_music_school_is_not_general_school(self):
        selected,_=choose(record(city='Можайск',region='Московская область'),[
            {**candidate(name='Детская музыкальная школа № 1'), 'region':'Московская обл',
             'city':'Можайск', 'address':'Можайск, Московская 17'}])
        self.assertIsNone(selected)

    def test_street_name_cannot_confirm_region(self):
        selected,_=choose(record(city='',region='Московская область'),[
            {**candidate(),'region':'Калужская обл','address':'г Калуга, Московская область проезд'}])
        self.assertIsNone(selected)

    def test_candidate_audit_keeps_group_id(self):
        class API:
            actual_requests=1
            requests=[]
            def search(self,*args): return [{**candidate(),'group_id':''}]
        class Registry:
            metadata={}
        result={'records':[record(record_id='1')],'groups':[dict(group_id='g',record_ids=['1'])]}
        resolve(result,API(),Registry(),progress=False)
        self.assertEqual(result['resolution']['candidates'][0]['group_id'],'g')

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
