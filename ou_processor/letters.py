
from collections import defaultdict
import re
import unicodedata


LETTER_COLUMNS = [
    ("letter_number", "Номер письма"),
    ("official_name", "Официальное название"),
    ("genitive_name", "Название в родительном падеже"),
    ("postal_code", "Почтовый индекс"),
    ("address", "Адрес без индекса"),
    ("region", "Регион"),
    ("students", "Поступившие"),
    ("student_count", "Количество поступивших"),
    ("match_status", "Статус сопоставления"),
    ("status", "Готовность"),
    ("inn", "ИНН"),
    ("ogrn", "ОГРН"),
    ("source_names", "Варианты исходного названия"),
    ("flags", "Признаки риска"),
    ("operator_comment", "Комментарий оператора"),
    ("letter_type", "Вид письма"),
    ("mailing_address", "Адрес для письма"),
    ("source", "Источник"),
    ("source_url", "Ссылка на источник"),
    ("letter_fragment", "Текст для письма"),
    ("checks", "Причины проверки"),
    ("source_record_ids", "Исходные записи"),
]


def _value(value):
    return "" if value is None else " ".join(str(value).split())


def _normalized(value):
    return " ".join(
        re.findall(r"\w+", _value(value).casefold().replace("ё", "е"))
    )


def valid_inn(value):
    inn = _value(value)

    if not inn.isascii() or not inn.isdecimal():
        return False

    if len(inn) not in (10, 12):
        return False

    digits = list(map(int, inn))

    def check_digit(weights):
        return sum(
            digit * weight
            for digit, weight in zip(digits, weights)
        ) % 11 % 10

    if len(inn) == 10:
        return check_digit(
            (2, 4, 10, 3, 5, 9, 4, 6, 8)
        ) == digits[9]

    return (
        check_digit(
            (7, 2, 4, 10, 3, 5, 9, 4, 6, 8)
        ) == digits[10]
        and check_digit(
            (3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8)
        ) == digits[11]
    )


def valid_postal_code(value):
    code = _value(value)
    return bool(re.fullmatch(r"[0-9]{6}", code))


def postal_details(candidate):
    address = _value(candidate.get("address"))
    explicit = _value(candidate.get("postal_code"))
    notes = []

    indices = re.findall(
        r"(?<!\d)[0-9]{6}(?!\d)", address
    )
    unique_indices = set(indices)

    if explicit and not valid_postal_code(explicit):
        notes.append("Некорректный индекс в источнике")

    if len(unique_indices) > 1:
        notes.append("В адресе обнаружены разные индексы")

    if (
        valid_postal_code(explicit)
        and unique_indices
        and unique_indices != {explicit}
    ):
        notes.append("Индексы в источниках не совпадают")

    code = ""

    if valid_postal_code(explicit):
        code = explicit
    elif len(unique_indices) == 1:
        code = next(iter(unique_indices))

    if notes:
        code = ""

    cleaned = re.sub(
        r"^\s*(?:Россия|Российская Федерация)\s*,?\s*",
        "",
        address,
        flags=re.I,
    )

    cleaned = re.sub(
        r"^\s*[0-9]{6}\s*,?\s*",
        "",
        cleaned,
    )

    cleaned = re.sub(
        r"^\s*(?:Россия|Российская Федерация)\s*,?\s*",
        "",
        cleaned,
        flags=re.I,
    )

    cleaned = re.sub(r"\s*,\s*", ", ", cleaned)
    cleaned = cleaned.strip(" ,")

    if re.match(r"^[0-9]{6}(?:\s|,|$)", cleaned):
        cleaned = re.sub(
            r"^[0-9]{6}\s*,?\s*",
            "",
            cleaned,
        )

    if not cleaned:
        notes.append("Отсутствует адрес организации")

    if not valid_postal_code(code):
        notes.append("Нет подтверждённого шестизначного индекса")

    return code, cleaned, sorted(set(notes))


_ABBREVIATIONS = {
    "ГБОУ", "ГБПОУ", "МБОУ", "МАОУ",
    "МКОУ", "МОУ", "ФГБОУ", "ФГАОУ",
    "АНО", "НОУ", "ОУ", "СПО",
    "РФ", "РАН", "МГУ", "СПБГУ",
}

_GEOGRAPHY = {
    "москвы": "Москвы",
    "москва": "Москва",
    "россии": "России",
    "российской": "Российской",
    "федерации": "Федерации",
    "санкт-петербурга": "Санкт-Петербурга",
    "тверской": "Тверской",
    "орловской": "Орловской",
    "свердловской": "Свердловской",
}


def _repair_uppercase(name):
    words = re.findall(
        r"[А-Яа-яЁёA-Za-z]+(?:-[А-Яа-яЁёA-Za-z]+)*",
        name,
    )

    ordinary = [
        word for word in words
        if word.upper() not in _ABBREVIATIONS
    ]

    all_caps = (
        len(words) >= 2
        and bool(ordinary)
        and all(word.isupper() for word in words)
    )

    if not all_caps:
        return name, False

    name = name.lower()
    name = name[0].upper() + name[1:]

    for word, replacement in _GEOGRAPHY.items():
        name = re.sub(
            rf"(?<!\w){word}(?!\w)",
            replacement,
            name,
            flags=re.I,
        )

    for abbr in _ABBREVIATIONS:
        name = re.sub(
            rf"(?<!\w){abbr.casefold()}(?!\w)",
            abbr,
            name,
            flags=re.I,
        )

    name = re.sub(
        r"«\s*([а-яёa-z])",
        lambda match: "«" + match.group(1).upper(),
        name,
    )

    return name, True


def _repair_quotes(name):
    # Внешние кавычки — ёлочки, вложенные — лапки; незакрытые пары закрываем в конце.
    result, opened = [], []
    for index, char in enumerate(name):
        if char not in '«»„“”"':
            result.append(char)
            continue
        if char in '«„':
            opening = True
        elif char in '»”' or (char == '“' and opened and opened[-1] == '„'):
            opening = False
        else:
            following = name[index + 1:index + 2]
            opening = not opened or (index > 0 and name[index - 1].isspace()
                                     and bool(following) and following.isalnum())
        if opening:
            result.append('«' if len(opened) % 2 == 0 else '„')
            opened.append(char)
        elif opened:
            opened.pop()
            result.append('»' if len(opened) % 2 == 0 else '“')
        else:
            result.append(char)
    while opened:
        opened.pop()
        result.append('»' if len(opened) % 2 == 0 else '“')
    return ''.join(result)


def official_typography(value):
    name = unicodedata.normalize("NFC", _value(value))

    name = _repair_quotes(name)

    name = re.sub(r"\s*№\s*(?=\d)", " № ", name)
    name = re.sub(r"\s+([,.;:»“])", r"\1", name)
    name = re.sub(r"([«„])\s+", r"\1", name)
    name = re.sub(r"\s+", " ", name).strip()

    name, adjusted = _repair_uppercase(name)

    if name and name[0].islower():
        name = name[0].upper() + name[1:]

    name = re.sub(
        r"«([а-яё])",
        lambda match: "«" + match.group(1).upper(),
        name,
    )

    return name, adjusted


# Совпадение окончания не подтверждает склонение: например, «хорошая» → «хорошей».
# Автоматически используем только проверенные формы типовых названий.
_GENITIVE_ADJECTIVES = {
    "муниципальное": "муниципального",
    "государственное": "государственного",
    "федеральное": "федерального",
    "бюджетное": "бюджетного",
    "автономное": "автономного",
    "казенное": "казенного",
    "казённое": "казённого",
    "частное": "частного",
    "негосударственное": "негосударственного",
    "образовательное": "образовательного",
    "общеобразовательное": "общеобразовательного",
    "профессиональное": "профессионального",
    "дошкольное": "дошкольного",
    "средняя": "средней",
    "основная": "основной",
    "начальная": "начальной",
    "общеобразовательная": "общеобразовательной",
    "специальная": "специальной",
    "коррекционная": "коррекционной",
}

_NOUNS = {
    "учреждение": "учреждения",
    "организация": "организации",
    "школа": "школы",
    "гимназия": "гимназии",
    "академия": "академии",
    "университет": "университета",
    "институт": "института",
    "колледж": "колледжа",
    "техникум": "техникума",
    "лицей": "лицея",
}


_GENITIVE_FORMS = {**_GENITIVE_ADJECTIVES, **_NOUNS}
_NOMINATIVE_FORMS = {genitive: nominative for nominative, genitive in _GENITIVE_FORMS.items()}
_PROPER_WORD = r"[А-ЯЁ][а-яё]+(?:-[А-ЯЁа-яё]+)*"
_PRESERVED_GENITIVE_TAIL = re.compile(
    rf"(?:города|[Рр]еспублики)\s+{_PROPER_WORD}(?:\s+{_PROPER_WORD})*"
    r"|[А-ЯЁ][а-яё-]*(?:ой|ого|его)\s+(?:области|края|[Рр]еспублики)"
    r"|(?:(?:среднего|основного|начального|высшего|дополнительного|"
    r"профессионального|общего|дошкольного)\s+)*образования"
    r"|№\s*\d+"
)


def _preserved_tail_length(words, start):
    # Не захватываем следующую именную часть, даже если она начинается с прописной буквы.
    for length in range(min(5, len(words) - start), 0, -1):
        tail = words[start:start + length]
        if any(word.casefold() in _GENITIVE_FORMS for word in tail):
            continue
        if _PRESERVED_GENITIVE_TAIL.fullmatch(' '.join(tail)):
            return length
    return 0


def _initial_case(form, original):
    return form[0].upper() + form[1:] if original[0].isupper() else form


def genitive_official_name(name):
    if not name:
        return "", False

    outside, separator, inside = name.partition("«")
    words = outside.strip().split()
    converted = list(words)
    changed = []
    index, has_head = 0, False

    while index < len(words):
        preserved = _preserved_tail_length(words, index) if has_head else 0
        if preserved:
            index += preserved
            continue
        # Склоняем только полную именную часть: известные определения и головное существительное.
        while index < len(words) and (words[index] in _ABBREVIATIONS
                                      or words[index].casefold() in _GENITIVE_ADJECTIVES):
            word = words[index]
            if word not in _ABBREVIATIONS:
                if word.isupper():
                    return "", False
                converted[index] = _initial_case(_GENITIVE_FORMS[word.casefold()], word)
                changed.append(index)
            index += 1
        if index == len(words) or words[index].casefold() not in _NOUNS or words[index].isupper():
            return "", False
        converted[index] = _initial_case(_NOUNS[words[index].casefold()], words[index])
        changed.append(index)
        has_head = True
        index += 1

    # Обратная проверка изменённых слов по словарю; хвосты и цитируемое название не склоняются.
    restored = list(converted)
    for index in changed:
        original_form = _NOMINATIVE_FORMS.get(converted[index].casefold())
        if original_form is None:
            return "", False
        restored[index] = _initial_case(original_form, converted[index])
    if not has_head or restored != words:
        return "", False
    result = ' '.join(converted)
    if separator:
        result += " «" + inside
    return result, True


def _person_ok(name):
    parts = _value(name).split()

    return (
        2 <= len(parts) <= 4
        and all(
            re.fullmatch(
                r"[А-Яа-яЁёA-Za-z-]+", part
            )
            for part in parts
        )
    )


def _students(students):
    names = defaultdict(set)
    names_per_file = defaultdict(set)
    records = set()
    notes = []

    for entry in students:
        full = _value(entry.get("student"))
        personal_file = _value(entry.get("file"))
        source_ids = entry.get("record_ids") or ""

        records.update(str(source_ids).splitlines())

        if not full:
            notes.append("Отсутствует ФИО")
            continue

        if not _person_ok(full):
            notes.append(
                f"Некорректное или неполное ФИО: {full}"
            )

        names[full].add(personal_file)

        if personal_file:
            names_per_file[personal_file].add(_normalized(full))

    if not names:
        notes.append("Нет сведений о поступивших")

    if any(len(file_names) > 1 for file_names in names_per_file.values()):
        notes.append("Один номер дела связан с разными ФИО")

    if any(
        len(files - {""}) > 1
        for files in names.values()
    ):
        notes.append(
            "Одно ФИО связано с разными личными делами"
        )

    counters = defaultdict(int)

    for entry in students:
        name = _value(entry.get("student"))
        personal_file = _value(entry.get("file"))

        if name and not personal_file:
            counters[name] += 1

    if any(count > 1 for count in counters.values()):
        notes.append("Повтор ФИО без номера личного дела")

    ordered = sorted(
        names,
        key=lambda name: (
            name.casefold().replace("ё", "е"),
            name,
        ),
    )

    return ordered, sorted(records), sorted(set(notes))


def _preferred_candidates(groups):
    candidates = [
        group.get("selected")
        for group in groups
        if group.get("selected")
    ]

    primary = [
        candidate for candidate in candidates
        if candidate.get("source") == "DaData"
    ]

    return primary or candidates


def _prepare_one(inn, groups, students, override=None):
    checks = []
    candidates = _preferred_candidates(groups)
    chosen = candidates[0] if candidates else {}

    if not valid_inn(inn):
        checks.append("ИНН не прошёл проверку контрольной суммы")

    if not candidates:
        checks.append(
            "Нет выбранной записи с подтверждённым ИНН"
        )

    official_names = {
        _normalized(candidate.get("candidate_name"))
        for candidate in candidates
        if candidate.get("candidate_name")
    }

    if len(official_names) != 1:
        checks.append("Официальные названия не совпадают")

    official_name, restored = official_typography(
        chosen.get("candidate_name", "")
    )

    if not official_name:
        checks.append("Отсутствует официальное название")

    if official_name.count("«") != official_name.count("»"):
        checks.append("Несбалансированные кавычки в названии")

    if restored:
        checks.append(
            "Регистр восстановлен автоматически — проверить"
        )

    suspicious_words = [
        word
        for word in re.findall(
            r"[А-Яа-яЁёA-Za-z]+", official_name
        )
        if (
            len(word) > 3
            and word.isupper()
            and word not in _ABBREVIATIONS
        )
    ]

    if suspicious_words:
        checks.append(
            "Регистр отдельных слов требует проверки"
        )

    genitive, template_ok = genitive_official_name(
        official_name
    )

    if not template_ok:
        checks.append("Родительный падеж требует проверки")

    if override and override.get("verified") is True:
        manual_name = _value(override.get("official_name"))
        manual_genitive = _value(
            override.get("genitive_name")
        )

        corrected, _ = official_typography(manual_name)

        if (
            corrected
            and manual_genitive
            and corrected.count("«") == corrected.count("»")
            and manual_genitive.count("«")
                == manual_genitive.count("»")
        ):
            official_name = corrected
            genitive = manual_genitive

            checks = [
                check for check in checks
                if not check.startswith((
                    "Регистр восстановлен",
                    "Регистр отдельных слов",
                    "Родительный падеж",
                ))
                and check not in {
                    "Отсутствует официальное название",
                    "Несбалансированные кавычки в названии",
                }
            ]
        else:
            checks.append(
                "Подтверждённые исправления заполнены неверно"
            )

    addresses = {
        _normalized(candidate.get("address"))
        for candidate in candidates
        if candidate.get("address")
    }

    postcodes = {
        _value(candidate.get("postal_code"))
        for candidate in candidates
        if candidate.get("postal_code")
    }

    if len(addresses) > 1:
        checks.append("Адреса в источниках не совпадают")

    if len(postcodes) > 1:
        checks.append("Индексы в источниках не совпадают")

    code, address, postal_notes = postal_details(chosen)
    checks.extend(postal_notes)

    people, source_ids, person_notes = _students(students)
    checks.extend(person_notes)

    if len(people) == 1:
        kind = "Для одного поступившего"
    elif len(people) > 1:
        kind = "Для нескольких поступивших"
    else:
        kind = "Не определён"

    fragment = ""

    if people and genitive:
        if len(people) == 1:
            fragment = (
                f"Благодарим коллектив {genitive} "
                "за подготовку к поступлению.\n"
                f"ФИО: {people[0]}."
            )
        else:
            fragment = (
                f"Благодарим коллектив {genitive} "
                "за подготовку поступивших:\n"
                + "\n".join(people)
                + "."
            )

    checks = sorted(set(checks))

    return {
        "status": (
            "Готово к печати"
            if not checks else "Проверить"
        ),
        "letter_type": kind,
        "inn": inn,
        "official_name": official_name,
        "genitive_name": genitive,
        "students": ", ".join(people),
        "student_count": len(people),
        "postal_code": code,
        "address": address,
        "mailing_address": (
            f"{code}, {address}"
            if code and address else ""
        ),
        "source": chosen.get("source", ""),
        "source_url": chosen.get("source_url", ""),
        "letter_fragment": fragment,
        "checks": "\n".join(checks),
        "source_record_ids": "\n".join(source_ids),
    }


def prepare_letters(result, overrides=None):
    overrides = overrides or {}

    groups_by_inn = defaultdict(list)
    unresolved = []

    for group in result.get(
        "resolution", {}
    ).get("groups", []):

        selected = group.get("selected")

        if selected and _value(
            selected.get("candidate_inn")
        ):
            inn = _value(selected["candidate_inn"])
            groups_by_inn[inn].append(group)
        else:
            unresolved.append(group)

    by_id = {
        record.get("record_id"): record
        for record in result.get("records", [])
    }

    students_by_inn = defaultdict(list)

    for entry in result.get("students", []):
        inn = _value(entry.get("inn"))
        students_by_inn[inn].append(entry)

    rows = []

    for inn, groups in sorted(groups_by_inn.items()):
        row = _prepare_one(
            inn,
            groups,
            students_by_inn.get(inn, []),
            overrides.get(inn),
        )

        member_ids = {
            record_id
            for group in groups
            for record_id in group.get("record_ids", [])
        }

        if any(
            not _value(by_id[record_id].get("student_name"))
            for record_id in member_ids
            if record_id in by_id
        ):
            row["status"] = "Проверить"
            row["checks"] += (
                "\n" if row["checks"] else ""
            ) + "В исходных записях отсутствует ФИО"

        rows.append(row)

    for group in unresolved:
        members = [
            by_id[record_id]
            for record_id in group.get("record_ids", [])
            if record_id in by_id
        ]

        students = [
            {
                "student": member.get("student_name", ""),
                "file": member.get(
                    "personal_file_number", ""
                ),
                "record_ids": member.get("record_id", ""),
            }
            for member in members
        ]

        proposal = group.get('review_candidate')
        proposed_group = {**group, 'selected': proposal} if proposal else group
        row = _prepare_one(proposal['candidate_inn'] if proposal else '', [proposed_group], students)
        row['status'] = 'Проверить'
        if not proposal:
            row['official_name'] = _value(group.get('source_name'))
        row['checks'] += '\n' + group.get('decision_reason', 'ИНН требует подтверждения')

        row["checks"] += (
            "\nНельзя объединять записи без "
            "подтверждённого ИНН"
        )

        rows.append(row)

    ready = [
        row for row in rows
        if row["status"] == "Готово к печати"
    ]

    review = [
        row for row in rows
        if row["status"] != "Готово к печати"
    ]

    def sort_key(row):
        return (
            row["official_name"].casefold(),
            row["inn"],
        )

    ready.sort(key=sort_key)
    review.sort(key=sort_key)

    groups_by_record = {
        record_id: group for group in result.get('resolution', {}).get('groups', [])
        for record_id in group.get('record_ids', [])
    }
    for number, row in enumerate(ready + review, 1):
        ids = row['source_record_ids'].splitlines()
        members = [by_id[i] for i in ids if i in by_id]
        groups = [groups_by_record[i] for i in ids if i in groups_by_record]
        selected = [g.get('selected') or g.get('review_candidate') for g in groups if g.get('selected') or g.get('review_candidate')]
        regions = {c.get('region') for c in selected if c.get('region')}
        ogrns = {c.get('candidate_ogrn') for c in selected if c.get('candidate_ogrn')}
        statuses = {g.get('status', '') for g in groups}
        match_status = ('АВТО' if row['inn'] and statuses <= {'АВТО', 'СОПОСТАВЛЕНО'}
                        else 'ПРОВЕРИТЬ' if statuses == {'ПРОВЕРИТЬ'} else 'ВРУЧНАЯ')
        flags = {flag for r in members for flag in r.get('flags', [])}
        if not row['postal_code']:
            flags.add('НЕТ_ИНДЕКСА')
        if 'Регистр восстановлен' in row['checks']:
            flags.add('РЕГИСТР_ВОССТАНОВЛЕН')
        row.update(letter_number=number, region=next(iter(regions)) if len(regions) == 1 else '',
                   ogrn=next(iter(ogrns)) if len(ogrns) == 1 else '',
                   match_status=match_status,
                   source_names=' || '.join(sorted({r['name'] for r in members})),
                   flags='; '.join(sorted(flags)),
                   operator_comment='')

    result["letters_ready"] = ready
    result["letters_review"] = review

    if "full_summary" in result:
        result["full_summary"].update({
            "Благодарственных писем готово к печати":
                len(ready),
            "Благодарственных писем требуют проверки":
                len(review),
            "Писем для одного поступившего":
                sum(
                    row["letter_type"] == "Для одного поступившего"
                    for row in rows
                ),
            "Писем для нескольких поступивших":
                sum(
                    row["letter_type"] == "Для нескольких поступивших"
                    for row in rows
                ),
        })

    return ready, review
