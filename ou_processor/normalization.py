"""Поисковая нормализация без изменения исходных и официальных названий."""

import json
from pathlib import Path
import re
import unicodedata

DEFAULT_CONFIG_PATH = Path(__file__).resolve().parents[1] / "config.json"


def _load_default_config():
    if DEFAULT_CONFIG_PATH.exists():
        try:
            return json.loads(DEFAULT_CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


_DEFAULT_CONFIG = _load_default_config()

ABBREVIATIONS = _DEFAULT_CONFIG.get("abbreviations", {
    "гбоу": "государственное бюджетное общеобразовательное учреждение",
    "гбпоу": "государственное бюджетное профессиональное образовательное учреждение",
    "гапоу": "государственное автономное профессиональное образовательное учреждение",
    "мбоу": "муниципальное бюджетное общеобразовательное учреждение",
    "маоу": "муниципальное автономное общеобразовательное учреждение",
    "мкоу": "муниципальное казенное общеобразовательное учреждение",
    "моу": "муниципальное общеобразовательное учреждение",
    "гоу": "государственное общеобразовательное учреждение",
    "гаоу": "государственное автономное общеобразовательное учреждение",
    "чоу": "частное общеобразовательное учреждение",
    "чпоу": "частное профессиональное образовательное учреждение",
    "аноо": "автономная некоммерческая общеобразовательная организация",
    "ано": "автономная некоммерческая организация",
    "ноу": "негосударственное образовательное учреждение",
    "сош": "средняя общеобразовательная школа",
    "сш": "средняя школа",
    "оош": "основная общеобразовательная школа",
    "цо": "центр образования",
    "оц": "образовательный центр",
    "сп": "структурное подразделение",
    "им": "имени",
    "им.": "имени",
})

INSTITUTION_TYPES = _DEFAULT_CONFIG.get("institution_types", [
    {"type": "СПО", "pattern": r"\b(?:колледж\w*|техникум\w*|училищ\w*)"},
    {"type": "ВУЗ", "pattern": r"\b(?:университет\w*|институт\w*|академи\w*)"},
    {"type": "ШКОЛА", "pattern": r"\b(?:школ\w*|лицей\w*|лицея|гимнази\w*|цо\b|центр(?:а)?\s+образования|образовательн\w*\s+центр\w*)"},
])


def text(value):
    return "" if value is None else str(value).strip()


def normalize(value):
    val = text(value)
    # Разделяем склеенные инициалы (например "А.А." -> "А. А.")
    val = re.sub(r"(?<=[а-яa-zА-ЯA-Z])\.(?=[а-яa-zА-ЯA-Z]\.)", ". ", val)
    # NFKC превращает № в No, поэтому защищаем обозначение до преобразования.
    val = unicodedata.normalize("NFKC", val.replace("№", " номер "))
    val = val.casefold()
    val = re.sub(r"\b(?:no\.?|n[°º])\s*(?=\d)", "номер ", val)
    val = re.sub(r"[\W_]+", " ", val, flags=re.UNICODE)
    return " ".join(val.split())


def expanded_key(value, config=None):
    abbr_map = config.get("abbreviations", ABBREVIATIONS) if config else ABBREVIATIONS
    words = normalize(value).split()
    words = [abbr_map.get(word, "имени" if word == "им" else word) for word in words]
    return " ".join(words)


def numbers(value):
    return tuple(re.findall(r"\bномер\s*(\d+[а-яa-z]?)\b", normalize(value)))


def institution_type(value, config=None):
    types_list = config.get("institution_types", INSTITUTION_TYPES) if config else INSTITUTION_TYPES
    val = expanded_key(value, config=config)
    for item in types_list:
        label = item["type"] if isinstance(item, dict) else item[1]
        pattern = item["pattern"] if isinstance(item, dict) else item[0]
        if re.search(pattern, val):
            return label
    return ""


def abbreviation_flags(value, config=None):
    abbr_map = config.get("abbreviations", ABBREVIATIONS) if config else ABBREVIATIONS
    key = normalize(value)
    flags = []
    if any(word in abbr_map for word in key.split()):
        flags.append("СОКРАЩЕНИЯ")
    if not institution_type(value, config=config) and len(key.split()) <= 5:
        flags.append("НЕЯСНОЕ_НАЗВАНИЕ")
    return flags
