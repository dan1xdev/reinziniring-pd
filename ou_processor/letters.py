
from collections import defaultdict
import re
import unicodedata


LETTER_COLUMNS = [
    ("status", "Готовность"),
    ("letter_type", "Вид письма"),
    ("inn", "Подтверждённый ИНН"),
    ("official_name", "Официальное название"),
    ("genitive_name", "Название в родительном падеже"),
    ("students", "Поступившие"),
    ("student_count", "Количество поступивших"),
    ("postal_code", "Почтовый индекс"),
    ("address", "Адрес без индекса"),
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
    return (
        len(code) == 6
        and code.isascii()
        and code.isdecimal()
    )


def postal_details(candidate):
    address = _value(candidate.get("address"))
    explicit = _value(candidate.get("postal_code"))

    found = set(
        re.findall(r"(?<!\d)[0-9]{6}(?!\d)", address)
    )

    notes = []

    if explicit and not valid_postal_code(explicit):
        notes.append("Некорректный индекс в источнике")

    if len(found) > 1:
        notes.append("В адресе найдено несколько индексов")

    if explicit and found and explicit not in found:
        notes.append("Индексы в источнике не совпадают")

    code = ""

    if valid_postal_code(explicit):
        code = explicit
    elif len(found) == 1:
        code = next(iter(found))

    if notes:
        code = ""

    cleaned = re.sub(
        r"^\s*[0-9]{6}\s*,?\s*", "", address
    )
    cleaned = re.sub(
        r"^Россия\s*,\s*", "", cleaned, flags=re.I
    )

    if not cleaned:
        notes.append("Отсутствует адрес организации")

    if not valid_postal_code(code):
        notes.append(
            "Нет подтверждённого шестизначного индекса"
        )

    return code, cleaned, notes


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
        r"[А-ЯЁA-Z][А-ЯЁA-Z-]*", name
    )

    all_caps = (
        len(words) >= 3
        and all(word == word.upper() for word in words)
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


def official_typography(value):
    name = unicodedata.normalize(
        "NFC", _value(value)
    )

    name = (
        name.replace("„", "«")
        .replace("“", "»")
        .replace("”", "»")
    )

    name = re.sub(
        r'"([^"\n]+)"',
        lambda match: "«" + match.group(1).strip() + "»",
        name,
    )

    name = re.sub(r"\s*№\s*(?=\d)", " № ", name)
    name = re.sub(r"\s+([,.;:»])", r"\1", name)
    name = re.sub(r"«\s+", "«", name)
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


_ADJECTIVE_ENDINGS = [
    ("ое", "ого"),
    ("ее", "его"),
    ("ая", "ой"),
    ("яя", "ей"),
    ("ый", "ого"),
    ("ий", "его"),
    ("ой", "ого"),
]

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


def genitive_official_name(name):
    if not name:
        return "", False

    outside, separator, inside = name.partition("«")
    words = outside.strip().split()

    head_idx = next(
        (
            index
            for index, word in enumerate(words)
            if word.casefold() in _NOUNS
        ),
        None,
    )

    if head_idx is None or head_idx == 0 or head_idx > 9:
        return "", False

    converted = []

    for word in words[:head_idx]:
        lower = word.casefold()

        ending = next(
            (
                (old, new)
                for old, new in _ADJECTIVE_ENDINGS
                if lower.endswith(old)
                and len(word) > len(old) + 2
            ),
            None,
        )

        if not ending:
            return "", False

        old, new = ending
        changed = word[:-len(old)] + new

        if word[0].isupper():
            changed = changed[0].upper() + changed[1:]

        converted.append(changed)

    head = words[head_idx]
    new_head = _NOUNS[head.casefold()]

    if head[0].isupper():
        new_head = new_head.capitalize()

    converted.append(new_head)

    tail = " ".join(words[head_idx + 1:])

    if re.search(
        r"\b(?:организация|школа|гимназия|лицей|"
        r"университет|институт|колледж|техникум)\b",
        tail,
        re.I,
    ):
        return "", False

    result = " ".join(
        converted + ([tail] if tail else [])
    )

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

    if not names:
        notes.append("Нет сведений о поступивших")

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
        notes.append(
            "Повтор ФИО без номера личного дела"
        )

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
        candidate
        for candidate in candidates
        if candidate.get("source") == "DaData"
    ]

    return primary or candidates


def _prepare_one(inn, groups, students, override=None):
    checks = []

    candidates = _preferred_candidates(groups)
    chosen = candidates[0] if candidates else {}

    if not valid_inn(inn):
        checks.append(
            "ИНН не прошёл проверку контрольной суммы"
        )

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
        checks.append(
            "Официальные названия не совпадают"
        )

    official_name, restored = official_typography(
        chosen.get("candidate_name", "")
    )

    if not official_name:
        checks.append(
            "Отсутствует официальное название"
        )

    if official_name.count("«") != official_name.count("»"):
        checks.append(
            "Несбалансированные кавычки в названии"
        )

    if restored:
        checks.append(
            "Регистр восстановлен автоматически — проверить"
        )

    genitive, template_ok = genitive_official_name(
        official_name
    )

    if not template_ok:
        checks.append(
            "Родительный падеж требует проверки"
        )

    if override and override.get("verified") is True:
        if override.get("official_name"):
            official_name, restored = official_typography(
                override["official_name"]
            )

        if override.get("genitive_name"):
            genitive = _value(
                override["genitive_name"]
            )

        if not restored and genitive and official_name:
            checks = [
                check for check in checks
                if not check.startswith((
                    "Регистр восстановлен",
                    "Родительный падеж",
                ))
            ]

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
        checks.append(
            "Адреса в источниках не совпадают"
        )

    if len(postcodes) > 1:
        checks.append(
            "Индексы в источниках не совпадают"
        )

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
                f"за подготовку поступившего {people[0]}."
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
            if not checks
            else "Проверить"
        ),
        "letter_type": kind,
        "inn": inn,
        "official_name": official_name,
        "genitive_name": genitive,
        "students": "\n".join(people),
        "student_count": len(people),
        "postal_code": code,
        "address": address,
        "mailing_address": (
            f"{code}, {address}"
            if code and address
            else ""
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
            inn = _value(
                selected["candidate_inn"]
            )
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
            not _value(
                by_id[record_id].get("student_name")
            )
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
                "student": member.get(
                    "student_name", ""
                ),
                "file": member.get(
                    "personal_file_number", ""
                ),
                "record_ids": member.get(
                    "record_id", ""
                ),
            }
            for member in members
        ]

        row = _prepare_one(
            "", [group], students
        )

        row["official_name"] = _value(
            group.get("source_name")
        )

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
