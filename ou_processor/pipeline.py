"""Отбор, признаки риска, строгие группы и кандидаты на сопоставление."""

from collections import Counter, defaultdict
from difflib import SequenceMatcher
import hashlib
import re

from .geography import extract_geography
from .normalization import ABBREVIATIONS, abbreviation_flags, expanded_key, institution_type, normalize, numbers, text


# --- Определение страны ОУ ---
RUSSIA = "РОССИЯ"                      # валидный ИНН
RUSSIA_PROBABLE = "РОССИЯ_ВЕРОЯТНО"    # только слабые российские маркеры
UNKNOWN = "НЕИЗВЕСТНО"                 # нет данных
FOREIGN = "ЗА_РУБЕЖОМ"                 # иностранный признак, нет ИНН -> исключаем
CONFLICT = "КОНФЛИКТ_СТРАНЫ"           # иностранный признак + валидный ИНН -> проверка


def _words(*items):
    return re.compile(r"(?<!\w)(?:" + "|".join(re.escape(i) for i in items) + r")(?!\w)")


FOREIGN_RE = _words(
    "казахстан", "беларусь", "белоруссия", "украина", "узбекистан", "кыргызстан",
    "киргизия", "таджикистан", "туркменистан", "армения", "азербайджан", "грузия",
    "молдова", "латвия", "литва", "эстония",
    "алматы", "астана", "нур-султан", "караганда", "шымкент", "минск", "гомель",
    "брест", "витебск", "бишкек", "ташкент", "душанбе", "ашхабад", "баку",
    "ереван", "тбилиси", "кишинев", "рига", "вильнюс", "таллин", "киев", "харьков",
)

RUSSIAN_SCHOOL_RE = re.compile(
    r"(?<!\w)(?:\w*(?:боу|поу|аоу|коу)|оу|сош|оош|нош|спо|нпо|"
    r"гимназия|лицей|школа|техникум|колледж|училище)(?!\w)"
    r"|организаци\w*|образовательн\w*|учреждени\w*")

RUSSIAN_ADDR_RE = re.compile(
    r"(?<!\w)(?:обл|область|край|респ|республика|р-н|район|город|г)\.?(?!\w)"
    r"|(?<!\w)(?:крым|симферополь|донецк|луганск|херсон|мелитополь)(?!\w)")

TEXT_FIELDS = ("name", "school_address_raw", "location_raw", "city_raw", "region_raw")
EXPLICIT_FOREIGN_COUNTRIES = {"германия", "китай"}


def valid_inn(s):
    """Проверка контрольных цифр ИНН (10 или 12 цифр)."""
    if not s.isdigit():
        return False
    d = [int(c) for c in s]

    def check(weights, index):
        return sum(w * x for w, x in zip(weights, d)) % 11 % 10 == d[index]

    if len(s) == 10:
        return check([2, 4, 10, 3, 5, 9, 4, 6, 8], 9)
    if len(s) == 12:
        return (check([7, 2, 4, 10, 3, 5, 9, 4, 6, 8], 10)
                and check([3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8], 11))
    return False


def classify_country(record):
    """Возвращает (метка_страны, флаги). Исключается только ЗА_РУБЕЖОМ."""
    flags = []

    inn_raw = str(record.get("inn_raw") or "")
    inn = "" if re.search(r"[eE]\+", inn_raw) else re.sub(r"\D", "", inn_raw)
    russian_inn = valid_inn(inn)
    if inn and not russian_inn:
        flags.append("ИНН_НЕКОРРЕКТЕН")

    blob = " ".join(str(record.get(k) or "") for k in TEXT_FIELDS).lower().replace("ё", "е")

    # Дополнительные страны распознаём как отдельную часть географии, не как имя школы или улицы.
    address_parts = {normalize(part) for field in TEXT_FIELDS if field != "name"
                     for part in re.split(r"[,;\n]", str(record.get(field) or ""))}
    foreign_hits = sorted(set(FOREIGN_RE.findall(blob)) | (address_parts & EXPLICIT_FOREIGN_COUNTRIES))

    postal = re.sub(r"\D", "", str(record.get("postal_code_raw") or ""))
    if len(postal) == 6 and postal[0] in "0789":
        flags.append("ИНДЕКС_НЕ_РОССИЙСКИЙ")

    if foreign_hits and russian_inn:
        return CONFLICT, flags + ["КОНФЛИКТ_СТРАНЫ"]
    if foreign_hits:
        return FOREIGN, flags + ["ЗА_РУБЕЖОМ:" + ",".join(foreign_hits)]
    if russian_inn:
        return RUSSIA, flags
        
    if RUSSIAN_ADDR_RE.search(blob) or RUSSIAN_SCHOOL_RE.search(blob):
        return RUSSIA_PROBABLE, flags
        
    return UNKNOWN, flags
# ---------------------------------------------------


def education_kind(record, config):
    explicit = normalize(record.get("education"))
    document = normalize(record.get("document"))
    aliases = {normalize(alias): kind for kind, values in config["education"].items()
               for alias in [kind, *values]}
    document_kind = next((kind for fragment, kind in config["documents"].items()
                          if normalize(fragment) in document), "")
    if explicit:
        kind = aliases.get(explicit, "НЕИЗВЕСТНО")
        conflict = bool(document_kind and document_kind != kind)
        return kind, "Вид образования", conflict
    return document_kind or "НЕИЗВЕСТНО", "Тип документа об образовании" if document_kind else "нет данных", False


def process(records, config):
    for record in records:
        name = record.get("name", "")
        record["personal_file_number"] = text(record.get("personal_file_number"))
        kind, source, conflict = education_kind(record, config)
        record.update({"education_kind": kind, "education_source": source,
                       "normalized_name": normalize(name), "search_variant": expanded_key(name, config=config),
                       "institution_type": institution_type(name, config=config), "numbers": list(numbers(name))})
        record.update(extract_geography(record, config))
        flags = list(record.pop("geo_flags")) + abbreviation_flags(name, config=config)
        flags.extend(record.get("input_flags", []))
        if any(re.search(pattern, record["normalized_name"], flags=re.IGNORECASE)
               for pattern in config.get("old_name_patterns", [])):
            flags.append("УСТАРЕВШЕЕ_НАИМЕНОВАНИЕ")
        
        # --- ОПРЕДЕЛЕНИЕ СТРАНЫ ---
        country_label, country_flags = classify_country(record)
        record["country"] = country_label
        flags.extend(country_flags)
        if country_label == FOREIGN:
            flags.append("ЗАРУБЕЖНОЕ_ОУ")
        elif country_label == CONFLICT:
            flags.append("КОНФЛИКТ_СТРАНЫ")
        # --------------------------
        
        if conflict:
            flags.append("КОНФЛИКТ_ОБРАЗОВАНИЯ")
        if len(name) == config["truncation_length"]:
            flags.append("ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ")
        year = record.get("year", "")
        if year:
            try:
                if int(year) <= config["old_graduation_year"]:
                    flags.append("СТАРЫЙ_ВЫПУСК")
            except ValueError:
                flags.append("НЕВЕРНЫЙ_ГОД")
        if record.get("input_conflict"):
            flags.append("КОНФЛИКТ_ЛИЧНОГО_ДЕЛА")
        if re.search(r"университет|институт|высшего образова|филиал|академи", normalize(name)):
            flags.append("ПОДРАЗДЕЛЕНИЕ_ВУЗА")
        if record["formula_fields"]:
            flags.append("ФОРМУЛА_ВО_ВХОДНЫХ_ДАННЫХ")
        if not name:
            flags.append("ПУСТОЕ_НАЗВАНИЕ")
            
        # --- ЛОГИКА ОТБОРА ---
        if country_label == FOREIGN:
            selection, reason = "ИСКЛЮЧЕНО", "Зарубежное ОУ"
        elif kind == "Высшее":
            selection, reason = "ИСКЛЮЧЕНО", "Высшее образование"
        elif country_label == CONFLICT:
            selection, reason = "ПРОВЕРИТЬ_ОТБОР", "Конфликт страны"
        elif record.get("input_conflict"):
            selection, reason = "ПРОВЕРИТЬ_ОТБОР", "Один номер личного дела у разных записей"
        elif "year" in record.get("damaged_fields", []):
            selection, reason = "ПРОВЕРИТЬ_ОТБОР", "Повреждён год завершения в Excel"
        elif kind not in config["allowed_education"] or conflict or not name or record["formula_fields"]:
            selection, reason = "ПРОВЕРИТЬ_ОТБОР", "Не хватает данных или есть противоречие"
        else:
            selection, reason = "ОТОБРАНО", "Допустимый вид образования"
        # ----------------------

        record.update({"selection": selection, "selection_reason": reason, "flags": sorted(set(flags)),
                       "status": "ПРОВЕРИТЬ" if selection != "ИСКЛЮЧЕНО" else "ИСКЛЮЧЕНО",
                       "confirmed_inn": "", "official_name": "", "group_id": ""})
                       
    add_region_hypotheses(records)
    groups = group_records(records)
    pairs = possible_matches(groups, records, config)
    
    summary = {
        "input_rows": len(records), 
        "selected_rows": sum(r["selection"] == "ОТОБРАНО" for r in records),
        "excluded_rows": sum(r["selection"] == "ИСКЛЮЧЕНО" for r in records),
        "selection_review_rows": sum(r["selection"] == "ПРОВЕРИТЬ_ОТБОР" for r in records),
        
        "foreign_rows": sum(r.get("country") == FOREIGN for r in records),
        "conflict_rows": sum(r.get("country") == CONFLICT for r in records),
        "russian_rows": sum(r.get("country") == RUSSIA for r in records),
        "russian_probable_rows": sum(r.get("country") == RUSSIA_PROBABLE for r in records),
        "unknown_country_rows": sum(r.get("country") == UNKNOWN for r in records),
        
        "preliminary_groups": len(groups), "possible_matches": len(pairs),
        "confirmed_organizations": 0,
        "selected_rows_with_personal_file": sum(r["selection"] == "ОТОБРАНО" and bool(r["personal_file_number"]) for r in records),
        "selected_rows_without_personal_file": sum(r["selection"] == "ОТОБРАНО" and not r["personal_file_number"] for r in records),
        "distinct_personal_file_numbers": len({r["personal_file_number"] for r in records
                                               if r["selection"] == "ОТОБРАНО" and r["personal_file_number"]}),
        "flags": dict(Counter(flag for r in records for flag in r["flags"]))
    }
    return {"summary": summary, "records": records, "groups": groups, "possible_matches": pairs}


def add_region_hypotheses(records):
    known = defaultdict(set)
    for record in records:
        if record["selection"] == "ОТОБРАНО" and record["region"]:
            known[record["normalized_name"]].add(record["region"])
    for record in records:
        record["region_hypothesis"] = ""
        options = known[record["normalized_name"]]
        if not record["region"] and len(options) == 1 and "КОНФЛИКТ_РЕГИОНА" not in record["flags"]:
            record["region_hypothesis"] = next(iter(options))
            record["flags"].append("РЕГИОН_ГИПОТЕЗА_ИЗ_ДРУГОЙ_СТРОКИ")


def group_records(records):
    buckets = defaultdict(list)
    for record in records:
        if record["selection"] != "ОТОБРАНО":
            continue
        key = (record["normalized_name"], record["region"], record["city"])
        isolate = not record["region"] or not record["city"] or any(
            flag in record["flags"] for flag in ["КОНФЛИКТ_РЕГИОНА", "КОНФЛИКТ_ГОРОДА",
                                                "ПОДОЗРЕНИЕ_НА_ОБРЕЗКУ", "СТАРЫЙ_ВЫПУСК"])
        if isolate:
            key += (record["record_id"],)
        buckets[key].append(record)
    groups = []
    for key, members in sorted(buckets.items()):
        group_id = "G-" + hashlib.sha256("\x1f".join(key).encode()).hexdigest()[:12]
        for member in members:
            member["group_id"] = group_id
        personal_files = sorted({m["personal_file_number"] for m in members if m["personal_file_number"]})
        missing = sum(not m["personal_file_number"] for m in members)
        personal_file_status = "НОМЕРА_ЕСТЬ_ВО_ВСЕХ_СТРОКАХ"
        if missing:
            personal_file_status = "НОМЕРА_ЕСТЬ_ЧАСТИЧНО" if personal_files else "НЕТ_НОМЕРОВ_В_ИСТОЧНИКЕ"
        groups.append({"group_id": group_id, "name": members[0]["name"],
                       "personal_file_numbers": personal_files, "personal_file_count": len(personal_files),
                       "rows_without_personal_file": missing, "personal_file_status": personal_file_status,
                       "normalized_name": key[0], "region": key[1], "city": key[2],
                       "row_count": len(members), "record_ids": [m["record_id"] for m in members],
                       "basis": "ИМЯ_РЕГИОН_ГОРОД" if len(key) == 3 else "ОТДЕЛЬНАЯ_СТРОКА",
                       "status": "ПРЕДВАРИТЕЛЬНАЯ", "confirmed_inn": ""})
    return groups


def possible_matches(groups, records, config):
    by_id = {r["record_id"]: r for r in records}
    pairs = []
    blocks = defaultdict(list)
    for group in groups:
        record = by_id[group["record_ids"][0]]
        if record["region"] and not any("КОНФЛИКТ" in flag for flag in record["flags"]):
            blocks[(record["region"], record["institution_type"])].append((group, record))
    for block in blocks.values():
        for i, (left, a) in enumerate(block):
            for right, b in block[i + 1:]:
                if a["city"] and b["city"] and a["city"] != b["city"]:
                    continue
                if a["numbers"] != b["numbers"]:
                    continue
                ka, kb = comparison_key(a, config), comparison_key(b, config)
                score = round(100 * SequenceMatcher(None, ka, kb, autojunk=False).ratio(), 1)
                same_school_number = (a["institution_type"] == "ШКОЛА" and a["numbers"]
                                      and a["city"] and a["city"] == b["city"])
                if ka and kb and (score >= config["possible_match_threshold"] or same_school_number):
                    reason = "Поисковые варианты похожи; требуется подтверждение ИНН"
                    if score < config["possible_match_threshold"]:
                        reason = "Совпали номер, тип и город; названия различаются, требуется проверка"
                    pairs.append({"left_group": left["group_id"], "right_group": right["group_id"],
                                  "left_name": a["name"], "right_name": b["name"],
                                  "text_similarity": score, "region": a["region"],
                                  "reason": reason,
                                  "status": "ПРОВЕРИТЬ"})
    return sorted(pairs, key=lambda p: (-p["text_similarity"], p["left_group"], p["right_group"]))


def comparison_key(record, config):
    key = record["search_variant"]
    # Удаление формы и географии допустимо только в очереди кандидатов.
    abbr_map = config.get("abbreviations", ABBREVIATIONS)
    removals = [phrase for phrase in abbr_map.values() if "учреждение" in phrase]
    city = record["city"]
    if city in config.get("cities", {}):
        removals += [normalize(alias) for alias in [city, *config["cities"][city].get("explicit_aliases", [])]]
    for phrase in sorted(removals, key=len, reverse=True):
        key = re.sub(r"(?<!\w)" + re.escape(phrase) + r"(?!\w)", " ", key)
    key = key.replace("средняя общеобразовательная школа", "школа")
    key = re.sub(r"\b(?:г|город|города)\b", " ", key)
    return " ".join(sorted(key.split()))
