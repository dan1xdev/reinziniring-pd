"""Поисковая нормализация без изменения исходных и официальных названий."""

import re
import unicodedata


ABBREVIATIONS = {
    "гбоу": "государственное бюджетное общеобразовательное учреждение",
    "гбпоу": "государственное бюджетное профессиональное образовательное учреждение",
    "гапоу": "государственное автономное профессиональное образовательное учреждение",
    "мбоу": "муниципальное бюджетное общеобразовательное учреждение",
    "маоу": "муниципальное автономное общеобразовательное учреждение",
    "мкоу": "муниципальное казенное общеобразовательное учреждение",
    "моу": "муниципальное общеобразовательное учреждение",
    "сош": "средняя общеобразовательная школа",
    "сш": "средняя школа",
}


def text(value):
    return "" if value is None else str(value).strip()


def normalize(value):
    # NFKC превращает № в No, поэтому защищаем обозначение до преобразования.
    value = unicodedata.normalize("NFKC", text(value).replace("№", " номер "))
    value = value.casefold()
    value = re.sub(r"\b(?:no\.?|n[°º])\s*(?=\d)", "номер ", value)
    value = re.sub(r"[\W_]+", " ", value, flags=re.UNICODE)
    return " ".join(value.split())


def expanded_key(value):
    words = normalize(value).split()
    words = [ABBREVIATIONS.get(word, "имени" if word == "им" else word) for word in words]
    return " ".join(words)


def numbers(value):
    return tuple(re.findall(r"\bномер\s*(\d+[а-яa-z]?)\b", normalize(value)))


def institution_type(value):
    value = expanded_key(value)
    for pattern, label in [
        (r"\b(?:колледж\w*|техникум\w*|училищ\w*)", "СПО"),
        (r"\b(?:университет\w*|институт\w*|академи\w*)", "ВУЗ"),
        (r"\b(?:школ\w*|лицей\w*|лицея|гимнази\w*)", "ШКОЛА"),
    ]:
        if re.search(pattern, value):
            return label
    return ""


def abbreviation_flags(value):
    key = normalize(value)
    flags = []
    if any(word in ABBREVIATIONS for word in key.split()):
        flags.append("СОКРАЩЕНИЯ")
    if not institution_type(value) and len(key.split()) <= 5:
        flags.append("НЕЯСНОЕ_НАЗВАНИЕ")
    return flags
