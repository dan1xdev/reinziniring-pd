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

DEFAULT_CONFIG_PATH = Path(__file__).resolve().parents[1] / "config.json"


def _load_default_config():
    if DEFAULT_CONFIG_PATH.exists():
        try:
            return json.loads(DEFAULT_CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


_DEFAULT_CONFIG = _load_default_config()


def get_stop_words(config=None):
    cfg = config or _DEFAULT_CONFIG
    stop_list = cfg.get("stop_words")
    if stop_list:
        return set(stop_list)
    return set('государственное государственный государственная бюджетное бюджетный бюджетная автономное автономная автономный муниципальное муниципальный муниципальная казенное казённое общеобразовательное образовательное профессиональное учреждение учреждения федеральное федерального государственногo государственного бюджетного образовательного профессионального среднего образования высшего некоммерческая организация краевое областное нетиповое средняя среднее общая общеобразовательная общеобразовательный основная основное школа школы школой лицей лицея гимназия гимназии колледж колледжа техникум техникума училище училища сош сш оош гбоу гбпоу гапоу мбоу моу маоу мкоу спо ноу ано оу фгбоу фгаоу г город города городского городской округа округ муниципального образования район района муниципальное область области обл республика республики респ край края имени им номер с и в на по при для об о д а к н п у гп оо лет углубленным изучением отдельных предметов'.split())


STOP = get_stop_words()


def terms(name, config=None):
    stop = get_stop_words(config)
    return {w for w in expanded_key(name, config).split() if w not in stop and len(w) > 1 and not w.isdigit()}


def school_numbers(name):
    found = numbers(name)
    if found:
        return tuple(sorted(set(found)))
    return tuple(sorted(set(re.findall(r'\b(?:школа|сош|лицей|гимназия|колледж|техникум)\s+(\d+)\b', normalize(name)))))


def meaningful(record, config=None):
    words = terms(record['name'], config)

    def geo_stem(word):
        word = re.sub(r'ск(?:ая|ой|ую|ий|ого|ом|ие|их)$', 'ск', word)
        return re.sub(r'[аыуыeе]$', '', word)

    for field in ('region', 'city'):
        geo = {geo_stem(w) for w in terms(record.get(field, ''), config)}
        words = {w for w in words if geo_stem(w) not in geo}
    return words


def queries(record, config=None):
    original = ' '.join(record['name'].split())
    geography = ' '.join(dict.fromkeys(x for x in [record['region'], record['city']] if x))
    primary = original + (' ' + geography if geography else '')
    key = expanded_key(record['name'], config)
    stop = get_stop_words(config)
    shortened = ' '.join(w for w in key.split() if w not in stop or w in {'школа', 'лицей', 'гимназия', 'колледж', 'техникум'})
    reduced = shortened + (' ' + geography if geography else '')
    variants = [primary, reduced]
    quoted = re.findall(r'[«"]([^»"]+)[»"]', original)
    if quoted:
        variants.append(max(quoted, key=len) + ' ' + geography)
    nums = school_numbers(original)
    kind = institution_type(original, config)
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
        self.entries = json.loads(self.path.read_text(encoding="utf-8")) if self.path.exists() else {}
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
                if entry['status'] == 'OK':
                    self.entries[query] = entry
                    self.path.parent.mkdir(parents=True, exist_ok=True)
                    temporary = self.path.with_suffix('.tmp')
                    temporary.write_text(json.dumps(self.entries, ensure_ascii=False, indent=2), encoding="utf-8")
                    temporary.replace(self.path)
        self.requests.append({'group_id': group_id, 'query': query, 'cached': cached,
                              'status': entry['status'], 'candidate_count': len(entry['candidates']),
                              'timestamp_utc': entry.get('timestamp_utc', '')})
        return entry['candidates']


def geography_matches(expected, candidate):
    tokens = set(normalize(expected).split()) - {'г', 'область', 'обл', 'республика', 'респ', 'край', 'город', 'ао'}
    actual = set(normalize(candidate).split())
    return bool(tokens) and tokens <= actual


def assess(record, candidate, config=None):
    cfg = config or _DEFAULT_CONFIG
    scoring_cfg = cfg.get("scoring", {})
    weights = scoring_cfg.get("weights", {
        "exact_inn": 100.0,
        "number_match": 35.0,
        "name_overlap_max": 40.0,
        "institution_type_match": 15.0,
        "city_match": 20.0,
        "region_match": 10.0,
        "quoted_match": 20.0,
        "active_status": 10.0
    })
    penalties = scoring_cfg.get("penalties", {
        "number_mismatch": 50.0,
        "type_mismatch": 35.0,
        "city_mismatch": 40.0,
        "region_mismatch": 40.0,
        "inactive_status": 30.0,
        "branch_risk": 20.0
    })
    thresholds = scoring_cfg.get("thresholds", {
        "auto_match_min_score": 85.0,
        "candidate_margin": 15.0,
        "numbered_city_min_score": 95.0,
        "distinctive_overlap_threshold": 0.85
    })

    reasons = []
    name = candidate['candidate_name']
    combined = name + ' ' + candidate.get('short_name', '')
    source_num = school_numbers(record['name'])
    target_num = school_numbers(name)

    if source_num and target_num and source_num != target_num:
        reasons.append('НОМЕР_НЕ_СОВПАЛ')
    elif source_num and not target_num:
        reasons.append('НОМЕР_НЕ_СОВПАЛ')

    source_type = institution_type(record['name'], cfg)
    target_type = institution_type(name, cfg)
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

    # Проверка правопреемства и ликвидации
    if candidate['legal_status'] not in {'ACTIVE', 'Действующее'}:
        reasons.append('НЕДЕЙСТВУЮЩАЯ_ЗАПИСЬ')
        reasons.append('НЕПОДТВЕРЖДЕННОЕ_ПРАВОПРЕЕМСТВО')
    if re.search(r'ликвидирован|в стадии ликвидации|в стадии реорганизации|правопреемник', normalize(name)):
        reasons.append('НЕПОДТВЕРЖДЕННОЕ_ПРАВОПРЕЕМСТВО')

    # Филиалы
    if candidate.get('branch_type') not in {'MAIN', '', None}:
        reasons.append('ФИЛИАЛ_ТРЕБУЕТ_ПРОВЕРКИ')

    # Географическое соответствие
    for field in ('region', 'city'):
        actual_geo = candidate.get(field, '')
        if field == 'city' or not actual_geo:
            actual_geo += ' ' + candidate.get('address', '')
        if record[field] and not geography_matches(record[field], actual_geo):
            reasons.append(field.upper() + '_НЕ_СОВПАЛ')

    # Противоречия и флаги риска из источника
    if any('КОНФЛИКТ' in f for f in record.get('flags', [])):
        reasons.append('ПРОТИВОРЕЧИЕ_В_ИСТОЧНИКЕ')
    if 'КОНФЛИКТ_РЕГИОНА' in record.get('flags', []):
        reasons.append('КОНФЛИКТ_РЕГИОНА')
    if 'КОНФЛИКТ_ГОРОДА' in record.get('flags', []):
        reasons.append('КОНФЛИКТ_ГОРОДА')
    if {'ВОЗМОЖНОЕ_ПОДРАЗДЕЛЕНИЕ_ВУЗА', 'ПОДРАЗДЕЛЕНИЕ_ВУЗА'} & set(record.get('flags', [])):
        reasons.append('ПОДРАЗДЕЛЕНИЕ_ВУЗА_ТРЕБУЕТ_ПРОВЕРКИ')
    if 'ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ' in record.get('flags', []):
        reasons.append('ОБРЕЗАННОЕ_НАЗВАНИЕ')
    if 'СТАРЫЙ_ВЫПУСК' in record.get('flags', []):
        reasons.append('СТАРЫЙ_ВЫПУСК_ТРЕБУЕТ_ПРОВЕРКИ')
        reasons.append('НЕПОДТВЕРЖДЕННОЕ_ПРАВОПРЕЕМСТВО')

    key = meaningful(record, cfg)
    target = terms(combined, cfg)
    overlap = sum(w in target for w in key) / len(key) if key else 0

    numbered_city = bool(source_num and record['city'] and source_type)
    distinctive_thresh = thresholds.get("distinctive_overlap_threshold", 0.85)
    distinctive = bool(key) and overlap >= distinctive_thresh and (
        len(key) >= 2 or max(map(len, key)) >= 8 or
        bool(record['city']) and max(map(len, key)) >= 4)
    exact_inn = bool(record.get('inn_raw') and record['inn_raw'] == candidate['candidate_inn'])
    identity = exact_inn or numbered_city or distinctive
    if not identity:
        reasons.append('НЕДОСТАТОЧНО_ПРИЗНАКОВ')

    # Расчёт баллов
    if exact_inn and not reasons:
        score = 100.0
    else:
        raw_score = 0.0
        # 1. Смысловое совпадение
        raw_score += overlap * weights.get("name_overlap_max", 40.0)

        # 2. Название внутри кавычек
        quoted = re.findall(r'[«"]([^»"]+)[»"]', record['name'])
        if quoted:
            q_text = normalize(max(quoted, key=len))
            if q_text and q_text in normalize(candidate['candidate_name']):
                raw_score += weights.get("quoted_match", 20.0)

        # 3. Номер
        if source_num and target_num:
            if source_num == target_num:
                raw_score += weights.get("number_match", 35.0)
            else:
                raw_score -= penalties.get("number_mismatch", 50.0)

        # 4. Тип
        if source_type and target_type:
            if source_type == target_type:
                raw_score += weights.get("institution_type_match", 15.0)
            else:
                raw_score -= penalties.get("type_mismatch", 35.0)

        # 5. Город
        if record['city']:
            actual_city = candidate.get('city', '') or candidate.get('address', '')
            if geography_matches(record['city'], actual_city):
                raw_score += weights.get("city_match", 20.0)
            else:
                raw_score -= penalties.get("city_mismatch", 40.0)

        # 6. Регион
        if record['region']:
            actual_reg = candidate.get('region', '') or candidate.get('address', '')
            if geography_matches(record['region'], actual_reg):
                raw_score += weights.get("region_match", 10.0)
            else:
                raw_score -= penalties.get("region_mismatch", 40.0)

        # 7. Действующий юридический статус
        if candidate['legal_status'] in {'ACTIVE', 'Действующее'}:
            raw_score += weights.get("active_status", 10.0)
        else:
            raw_score -= penalties.get("inactive_status", 30.0)

        if numbered_city and not reasons:
            raw_score = max(raw_score, thresholds.get("numbered_city_min_score", 95.0))

        score = round(max(0.0, min(100.0, raw_score)), 1)

    auto_min = thresholds.get("auto_match_min_score", 85.0)
    eligible = bool(identity and not reasons and score >= auto_min)

    return {**candidate, 'score': score, 'reasons': sorted(set(reasons)), 'eligible': eligible}


def choose(record, candidates, config=None):
    cfg = config or _DEFAULT_CONFIG
    thresholds = cfg.get("scoring", {}).get("thresholds", {
        "auto_match_min_score": 85.0,
        "candidate_margin": 15.0
    })
    ranked = [assess(record, c, cfg) for c in candidates]
    ranked.sort(key=lambda c: (c['eligible'], c['score'], c['legal_status'] in {'ACTIVE', 'Действующее'}), reverse=True)

    eligible_inns = {c['candidate_inn'] for c in ranked if c['eligible']}

    # Уточнение точным названием в кавычках при наличии нескольких подходящих
    quoted = re.findall(r'[«"]([^»"]+)[»"]', record['name'])
    exact_quoted = False
    if len(eligible_inns) > 1 and quoted:
        title = normalize(max(quoted, key=len))
        exact = [c for c in ranked if c['eligible'] and any(
            normalize(part) == title for part in re.findall(r'[«"]([^»"]+)[»"]', c['candidate_name']))]
        if len({c['candidate_inn'] for c in exact}) == 1:
            eligible_inns = {exact[0]['candidate_inn']}
            ranked = exact + [c for c in ranked if c not in exact]
            exact_quoted = True

    # Проверка отрыва кандидатов (margin)
    margin_threshold = thresholds.get("candidate_margin", 15.0)
    best_by_inn = {}
    for c in ranked:
        inn = c.get('candidate_inn')
        if inn:
            if inn not in best_by_inn or c['score'] > best_by_inn[inn]['score']:
                best_by_inn[inn] = c

    sorted_distinct = sorted(best_by_inn.values(), key=lambda c: (c['eligible'], c['score']), reverse=True)
    if len(sorted_distinct) >= 2:
        top1 = sorted_distinct[0]
        top2 = sorted_distinct[1]
        exact_inn = record.get('inn_raw') and record['inn_raw'] == top1['candidate_inn']
        if not exact_inn and not exact_quoted and (top1['score'] - top2['score'] < margin_threshold):
            # Неясность между лидерами запрещает выбирать и любого следующего кандидата.
            for c in ranked:
                if c['eligible'] or c['candidate_inn'] in (top1['candidate_inn'], top2['candidate_inn']):
                    c['reasons'].append('НЕДОСТАТОЧНЫЙ_ОТРЫВ_КАНДИДАТОВ')
                    c['eligible'] = False
            eligible_inns = {c['candidate_inn'] for c in ranked if c['eligible']}

    if len(eligible_inns) == 1:
        selected = next((c for c in ranked if c['eligible']), None)
    else:
        selected = None

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
    def __init__(self, cache, config=None):
        self.config = config or _DEFAULT_CONFIG
        current = current_snapshot(Path(cache))
        if not current:
            raise ValueError('Нет скачанного резервного реестра')
        directory, self.metadata = current
        self.connection = sqlite3.connect(f'file:{directory / "index.sqlite"}?mode=ro', uri=True)
        self.connection.row_factory = sqlite3.Row
        self.cache = {}
        self.regions = [row[0] for row in self.connection.execute('SELECT DISTINCT region FROM organizations')]

    def search(self, record):
        significant = meaningful(record, self.config)
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


def resolve(result, api, registry, progress=True, on_progress=None, check_cancel=None, config=None):
    cfg = config or _DEFAULT_CONFIG
    by_id = {r['record_id']: r for r in result['records']}
    resolutions, candidates = [], []
    for index, group in enumerate(result['groups'], 1):
        if check_cancel:
            check_cancel()
        record = by_id[group['record_ids'][0]]
        found = []
        selected, ranked = None, []
        for query in queries(record, cfg) if api is not None else []:
            if check_cancel:
                check_cancel()
            found.extend(api.search(query, group['group_id']))
            if check_cancel:
                check_cancel()
            selected, ranked = choose(record, found, cfg)
            if selected:
                break
        registry_used, truncated = False, False
        if not selected and registry is not None:
            if check_cancel:
                check_cancel()
            registry_used = True
            reserve, truncated = registry.search(record)
            chosen, reserve_ranked = choose(record, reserve, cfg)
            # Недостаточный отрыв сохраняет возможные ИНН основного поиска.
            primary_inns = {c['candidate_inn'] for c in ranked if c['eligible'] or
                            c['reasons'] == ['НЕДОСТАТОЧНЫЙ_ОТРЫВ_КАНДИДАТОВ']}
            if chosen and not truncated and (not primary_inns or chosen['candidate_inn'] in primary_inns):
                selected = chosen
            ranked.extend(reserve_ranked)

        unique = {}
        for c in ranked:
            key = (c['source'], c['candidate_inn'], c['candidate_name'], c['address'], c['legal_status'])
            unique[key] = c
        candidates.extend({**c, 'group_id': group['group_id']} for c in unique.values())

        if selected:
            status = 'СОПОСТАВЛЕНО'
        elif any('ОБРЕЗАННОЕ_НАЗВАНИЕ' in c['reasons'] for c in ranked) or 'ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ' in record.get('flags', []):
            status = 'ПРОВЕРИТЬ_ОБРЕЗКА'
        elif any('КОНФЛИКТ_РЕГИОНА' in c['reasons'] or 'REGION_НЕ_СОВПАЛ' in c['reasons'] for c in ranked) or 'КОНФЛИКТ_РЕГИОНА' in record.get('flags', []):
            status = 'ПРОВЕРИТЬ_КОНФЛИКТ_РЕГИОНА'
        elif any('НЕПОДТВЕРЖДЕННОЕ_ПРАВОПРЕЕМСТВО' in c['reasons'] or 'СТАРЫЙ_ВЫПУСК_ТРЕБУЕТ_ПРОВЕРКИ' in c['reasons'] for c in ranked) or 'СТАРЫЙ_ВЫПУСК' in record.get('flags', []):
            status = 'ПРОВЕРИТЬ_ПРАВОПРЕЕМСТВО'
        elif any('НЕДОСТАТОЧНЫЙ_ОТРЫВ_КАНДИДАТОВ' in c['reasons'] for c in ranked):
            status = 'ПРОВЕРИТЬ_ОТРЫВ'
        elif any(c['eligible'] for c in ranked):
            status = 'НЕОДНОЗНАЧНО'
        elif not ranked:
            status = 'НЕ_НАЙДЕНО'
        else:
            status = 'НЕ_СОПОСТАВЛЕНО'

        resolutions.append({'group_id': group['group_id'], 'record_ids': group['record_ids'],
            'source_name': record['name'], 'selected': selected, 'status': status,
            'registry_used': registry_used, 'registry_truncated': truncated,
            'candidate_count': len(unique)})
        if progress:
            print(f"{index}/{len(result['groups'])}: {status}; DaData={api.actual_requests if api else 0}", flush=True)
        if on_progress:
            on_progress(index, len(result['groups']))
    result['resolution'] = {'groups': resolutions, 'candidates': candidates, 'requests': api.requests if api else [],
                            'actual_requests': api.actual_requests if api else 0,
                            'registry': registry.metadata if registry is not None else {}}
    return result
