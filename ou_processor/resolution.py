"""Полный поиск: DaData, затем резервный реестр; объединение только по ИНН."""

from collections import defaultdict
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import sqlite3
import time

from .dadata import DaDataError, candidate_row
from .normalization import expanded_key, institution_type, normalize, numbers
from .registry import current_snapshot

STOP = set('государственное государственный государственная бюджетное бюджетный бюджетная автономное автономная автономный муниципальное муниципальный муниципальная казенное казённое общеобразовательное образовательное профессиональное учреждение учреждения федеральное федерального государственногo государственного бюджетного образовательного профессионального среднего образования высшего некоммерческая организация краевое областное нетиповое средняя среднее общая общеобразовательная общеобразовательный основная основное школа школы школой лицей лицея гимназия гимназии колледж колледжа техникум техникума училище училища сош сш гбоу гбпоу гапоу мбоу моу маоу мкоу спо ноу ано оу фгбоу фгаоу г город города городского городской округа округ муниципального образования район района муниципальное область области обл республика республики респ край края имени им номер с и в на по при для об о д а к н п у гп оо лет углубленным изучением отдельных предметов'.split())


def terms(name):
    return {w for w in expanded_key(name).split() if w not in STOP and len(w) > 1 and not w.isdigit()}


def school_numbers(name):
    found = numbers(name)
    if found:
        return tuple(sorted(set(found)))
    return tuple(sorted(set(re.findall(r'\b(?:школа|сош|лицей|гимназия|колледж|техникум)\s+(\d+)\b', normalize(name)))))


def meaningful(record):
    words = terms(record['name'])
    def geo_stem(word):
        word = re.sub(r'ск(?:ая|ой|ую|ий|ого|ом|ие|их)$', 'ск', word)
        return re.sub(r'[аыуыeе]$', '', word)
    for field in ('region', 'city'):
        geo = {geo_stem(w) for w in terms(record.get(field, ''))}
        words = {w for w in words if geo_stem(w) not in geo}
    return words


def queries(record):
    original = ' '.join(record['name'].split())
    geography = ' '.join(dict.fromkeys(x for x in [record['region'], record['city']] if x))
    primary = original + (' ' + geography if geography else '')
    # Короткий запрос сохраняет номер и смысловое название, убирая длинную правовую форму.
    key = expanded_key(record['name'])
    shortened = ' '.join(w for w in key.split() if w not in STOP or w in {'школа', 'лицей', 'гимназия', 'колледж', 'техникум'})
    reduced = shortened + (' ' + geography if geography else '')
    variants = [primary, reduced]
    quoted = re.findall(r'[«"]([^»"]+)[»"]', original)
    if quoted:
        variants.append(max(quoted, key=len) + ' ' + geography)
    nums = school_numbers(original)
    kind = institution_type(original)
    if nums and record['city'] and kind in {'ШКОЛА', 'СПО'}:
        variants.append(('школа' if kind == 'ШКОЛА' else 'колледж') + ' ' + ' '.join(nums) + ' ' + record['city'])
    return list(dict.fromkeys(q.strip() for q in variants if q.strip()))


def convert(suggestion, query):
    row = candidate_row({'group_id': '', 'query': query}, suggestion, 1, {'name': ''})
    data = suggestion['data']
    row['short_name'] = (data.get('name') or {}).get('short_with_opf') or ''
    row['source'] = 'DaData'
    row['source_url'] = 'https://dadata.ru/api/suggest/party/'
    row['flags'] = []
    return row


class SearchCache:
    def __init__(self, path, client=None):
        self.path = Path(path)
        self.client = client
        self.entries = json.loads(self.path.read_text()) if self.path.exists() else {}
        self.requests = []
        self.actual_requests = 0
        self.stopped = ''

    def search(self, query, group_id):
        cached = query in self.entries
        entry = self.entries.get(query)
        if entry is None:
            if not self.client:
                entry = {'status': 'НЕТ_В_КЭШЕ', 'candidates': []}
            elif self.stopped:
                entry = {'status': 'ПРОПУЩЕНО_' + self.stopped, 'candidates': []}
            else:
                time.sleep(0.2)
                self.actual_requests += 1
                try:
                    found = self.client.suggest(query)
                    entry = {'status': 'OK', 'candidates': [convert(x, query) for x in found]}
                except DaDataError as error:
                    entry = {'status': error.code, 'candidates': []}
                    if error.code in {'ОШИБКА_АВТОРИЗАЦИИ', 'ДОСТУП_ЗАПРЕЩЕН', 'ЛИМИТ_ЗАПРОСОВ'}:
                        self.stopped = error.code
                entry['timestamp_utc'] = datetime.now(timezone.utc).isoformat()
                # Ошибки не кэшируем: следующий запуск сможет повторить запрос.
                if entry['status'] == 'OK':
                    self.entries[query] = entry
                    self.path.parent.mkdir(parents=True, exist_ok=True)
                    temporary = self.path.with_suffix('.tmp')
                    temporary.write_text(json.dumps(self.entries, ensure_ascii=False, indent=2))
                    temporary.replace(self.path)
        self.requests.append({'group_id': group_id, 'query': query, 'cached': cached,
                              'status': entry['status'], 'candidate_count': len(entry['candidates']),
                              'timestamp_utc': entry.get('timestamp_utc', '')})
        return entry['candidates']


def geography_matches(expected, candidate):
    # Административные суффиксы различаются между источниками.
    tokens = set(normalize(expected).split()) - {'г', 'область', 'обл', 'республика', 'респ', 'край', 'город'}
    actual = set(normalize(candidate).split())
    return bool(tokens) and tokens <= actual


def assess(record, candidate):
    reasons = []
    name = candidate['candidate_name']
    combined = name + ' ' + candidate.get('short_name', '')
    source_num, target_num = school_numbers(record['name']), school_numbers(name)
    if source_num and source_num != target_num:
        reasons.append('НОМЕР_НЕ_СОВПАЛ')
    source_type, target_type = institution_type(record['name']), institution_type(name)
    if source_type and target_type and source_type != target_type:
        reasons.append('ТИП_НЕ_СОВПАЛ')
    if source_type == 'ШКОЛА' and re.search(r'дополнительного образования|музыкальн|школа искусств|спортивн|художественн', normalize(name)):
        reasons.append('ДОПОЛНИТЕЛЬНОЕ_ОБРАЗОВАНИЕ_ВМЕСТО_ШКОЛЫ')
    if re.search(r'профсоюз|местная общественная организация', normalize(name)):
        reasons.append('ОБЩЕСТВЕННАЯ_ОРГАНИЗАЦИЯ_ВМЕСТО_УЧРЕЖДЕНИЯ')
    if record.get('inn_raw') and record['inn_raw'] != candidate['candidate_inn']:
        reasons.append('ИНН_НЕ_СОВПАЛ')
    if not candidate['candidate_inn']:
        reasons.append('НЕТ_ИНН')
    if candidate['legal_status'] not in {'ACTIVE', 'Действующее'}:
        reasons.append('НЕДЕЙСТВУЮЩАЯ_ЗАПИСЬ')
    if candidate.get('branch_type') not in {'MAIN', '', None}:
        reasons.append('ФИЛИАЛ_ТРЕБУЕТ_ПРОВЕРКИ')
    for field in ('region', 'city'):
        actual_geo = candidate.get(field, '')
        if field == 'city' or not actual_geo:
            actual_geo += ' ' + candidate.get('address', '')
        if record[field] and not geography_matches(record[field], actual_geo):
            reasons.append(field.upper() + '_НЕ_СОВПАЛ')
    if any('КОНФЛИКТ' in f for f in record['flags']):
        reasons.append('ПРОТИВОРЕЧИЕ_В_ИСТОЧНИКЕ')
    if 'ВОЗМОЖНОЕ_ПОДРАЗДЕЛЕНИЕ_ВУЗА' in record['flags']:
        reasons.append('ПОДРАЗДЕЛЕНИЕ_ВУЗА_ТРЕБУЕТ_ПРОВЕРКИ')
    key = meaningful(record)
    target = terms(combined)
    overlap = sum(w in target for w in key) / len(key) if key else 0
    # Номер школы без населённого пункта не идентифицирует организацию.
    numbered_city = bool(source_num and record['city'] and source_type)
    distinctive = bool(key) and overlap >= .9 and (
        len(key) >= 2 or max(map(len, key)) >= 8 or
        bool(record['city']) and max(map(len, key)) >= 4)
    exact_inn = bool(record.get('inn_raw') and record['inn_raw'] == candidate['candidate_inn'])
    identity = exact_inn or numbered_city or distinctive
    if not identity:
        reasons.append('НЕДОСТАТОЧНО_ПРИЗНАКОВ')
    if 'ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ' in record['flags']:
        reasons.append('ОБРЕЗАННОЕ_НАЗВАНИЕ')
    score = round(100 * overlap, 1)
    if numbered_city and not reasons:
        score = max(score, 95)
    if exact_inn and not reasons:
        score = 100
    return {**candidate, 'score': score, 'reasons': reasons, 'eligible': identity and not reasons}


def choose(record, candidates):
    ranked = [assess(record, c) for c in candidates]
    ranked.sort(key=lambda c: (c['eligible'], c['score'], c['legal_status'] in {'ACTIVE', 'Действующее'}), reverse=True)
    eligible = {c['candidate_inn'] for c in ranked if c['eligible']}
    # Полное совпадение смыслового названия внутри кавычек уточняет общий поиск.
    quoted = re.findall(r'[«"]([^»"]+)[»"]', record['name'])
    if len(eligible) > 1 and quoted:
        title = normalize(max(quoted, key=len))
        exact = [c for c in ranked if c['eligible'] and any(
            normalize(part) == title for part in re.findall(r'[«"]([^»"]+)[»"]', c['candidate_name']))]
        if len({c['candidate_inn'] for c in exact}) == 1:
            eligible = {exact[0]['candidate_inn']}
            ranked = exact + [c for c in ranked if c not in exact]
    # Несколько ИНН нельзя сводить к первому результату API.
    selected = next((c for c in ranked if c['eligible']), None) if len(eligible) == 1 else None
    if selected and selected['source'] == 'Реестр аккредитации':
        same = [c for c in ranked if c['eligible'] and c['candidate_inn'] == selected['candidate_inn']]
        current = [c for c in same if c.get('scope') == 'actual_organization'] or same
        selected = dict(current[0])
        addresses = {normalize(c['address']) for c in current if c['address']}
        if len(addresses) > 1:
            selected['address'] = ''
            selected['postal_code'] = ''
            selected['address_note'] = 'Несколько адресов в действующих записях реестра'
    return selected, ranked


class RegistrySearch:
    def __init__(self, cache):
        current = current_snapshot(Path(cache))
        if not current:
            raise ValueError('Нет скачанного резервного реестра')
        directory, self.metadata = current
        self.connection = sqlite3.connect(f'file:{directory / "index.sqlite"}?mode=ro', uri=True)
        self.connection.row_factory = sqlite3.Row
        self.cache = {}
        self.regions = [row[0] for row in self.connection.execute('SELECT DISTINCT region FROM organizations')]

    def search(self, record):
        # Достаточно редкого слова; окончательное сопоставление проверяет все признаки.
        significant = meaningful(record)
        nums = school_numbers(record['name'])
        tokens = [max(significant, key=lambda w: (len(w), w))] if significant else []
        if nums:
            tokens.append(nums[0])
        if not tokens:
            return [], False
        region_options = [r for r in self.regions if geography_matches(record['region'], r)] if record['region'] else []
        key = (tuple(tokens), tuple(region_options))
        if key not in self.cache:
            condition = ' AND '.join('instr(search_name, ?) > 0' for _ in tokens)
            parameters = list(tokens)
            if region_options:
                condition += ' AND region IN (' + ','.join('?' for _ in region_options) + ')'
                parameters += region_options
            rows = self.connection.execute('SELECT DISTINCT inn, ogrn, full_name, short_name, post_address, '
                'region, certificate_status, record_status, is_branch, scope, issue_date, certificate_id '
                f'FROM organizations WHERE {condition} ORDER BY issue_date DESC LIMIT 2001', parameters).fetchall()
            self.cache[key] = [dict(row) for row in rows]
        found = self.cache[key]
        result = []
        seen = set()
        for r in found[:2000]:
            duplicate = (r['inn'], r['full_name'], r['post_address'], r['certificate_status'], r['record_status'], r['scope'])
            if duplicate in seen:
                continue
            seen.add(duplicate)
            address = r['post_address'] or ''
            active = r['certificate_status'] == 'Действующее' and r['record_status'] in {'Действующее', ''}
            result.append({'candidate_inn': r['inn'], 'candidate_ogrn': r['ogrn'], 'candidate_name': r['full_name'],
                'short_name': r['short_name'], 'address': address, 'region': r['region'], 'city': '',
                'postal_code': (re.search(r'\b\d{6}\b', address) or [''])[0],
                'legal_status': 'Действующее' if active else r['certificate_status'] + '/' + r['record_status'],
                'branch_type': 'BRANCH' if str(r['is_branch']).lower() == 'true' else '',
                'source': 'Реестр аккредитации', 'source_url': 'https://isga.obrnadzor.gov.ru/accredreestr/opendata/',
                'certificate_id': r['certificate_id'], 'scope': r['scope'], 'issue_date': r['issue_date']})
        return result, len(found) > 2000


def resolve(result, api, registry, progress=True, on_progress=None, check_cancel=None):
    by_id = {r['record_id']: r for r in result['records']}
    resolutions, candidates = [], []
    for index, group in enumerate(result['groups'], 1):
        if check_cancel:
            check_cancel()
        record = by_id[group['record_ids'][0]]
        found = []
        selected = None
        for query in queries(record):
            if check_cancel:
                check_cancel()
            found.extend(api.search(query, group['group_id']))
            if check_cancel:
                check_cancel()
            selected, ranked = choose(record, found)
            if selected:
                break
        registry_used, truncated = False, False
        if not selected:
            if check_cancel:
                check_cancel()
            registry_used = True
            reserve, truncated = registry.search(record)
            chosen, reserve_ranked = choose(record, reserve)
            primary_inns = {c['candidate_inn'] for c in ranked if c['eligible']}
            if chosen and not truncated and (not primary_inns or chosen['candidate_inn'] in primary_inns):
                selected = chosen
            ranked.extend(reserve_ranked)
        # В аудите показываем все уникальные кандидаты, а не только выбранного.
        unique = {}
        for c in ranked:
            key = (c['source'], c['candidate_inn'], c['candidate_name'], c['address'], c['legal_status'])
            unique[key] = c
        candidates.extend({**c, 'group_id': group['group_id']} for c in unique.values())
        status = 'СОПОСТАВЛЕНО' if selected else ('НЕОДНОЗНАЧНО' if any(c['eligible'] for c in ranked) else 'НЕ_СОПОСТАВЛЕНО')
        resolutions.append({'group_id': group['group_id'], 'record_ids': group['record_ids'],
            'source_name': record['name'], 'selected': selected, 'status': status,
            'registry_used': registry_used, 'registry_truncated': truncated,
            'candidate_count': len(unique)})
        if progress:
            print(f"{index}/{len(result['groups'])}: {status}; DaData={api.actual_requests}", flush=True)
        if on_progress:
            on_progress(index, len(result['groups']))
    result['resolution'] = {'groups': resolutions, 'candidates': candidates, 'requests': api.requests,
                            'actual_requests': api.actual_requests, 'registry': registry.metadata}
    return result
