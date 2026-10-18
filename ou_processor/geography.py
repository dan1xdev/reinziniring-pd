"""Консервативная география: явные названия и небольшой расширяемый словарь."""

import re

from .normalization import normalize, text


def matches(value, entries):
    value = " " + normalize(value) + " "
    return {name for name, aliases in entries.items()
            if any(" " + normalize(alias) + " " in value for alias in [name, *aliases])}


def extract_geography(record, config):
    region_evidence, city_evidence = [], []
    regions = config["regions"]
    cities = config["cities"]
    for field, source in [("region_raw", "поле"), ("location_raw", "местонахождение"),
                          ("name", "название")]:
        raw = record.get(field, "")
        for region in sorted(matches(raw, regions)):
            region_evidence.append({"value": region, "source": source})
        for city, details in cities.items():
            # В названии ищем явный город, а не прилагательное «Московский».
            aliases = details["explicit_aliases"]
            if field != "name":
                aliases = [*aliases, city]
            if matches(raw, {city: aliases}) and (
                field != "name" or any(re.search(pattern, normalize(raw))
                                       for pattern in details["name_patterns"])
            ):
                city_evidence.append({"value": city, "source": source})
                region_evidence.append({"value": details["region"],
                                        "source": source + ":город из словаря"})
    explicit_city = text(record.get("city_raw"))
    if explicit_city:
        matched = matches(explicit_city, {city: data["explicit_aliases"] for city, data in cities.items()})
        city_evidence.append({"value": next(iter(matched)) if len(matched) == 1 else normalize(explicit_city),
                              "source": "поле"})
        if len(matched) == 1:
            region_evidence.append({"value": cities[next(iter(matched))]["region"], "source": "поле:город из словаря"})
    region_values = {item["value"] for item in region_evidence}
    city_values = {item["value"] for item in city_evidence}
    flags = []
    if len(region_values) > 1:
        flags.append("КОНФЛИКТ_РЕГИОНА")
    if len(city_values) > 1:
        flags.append("КОНФЛИКТ_ГОРОДА")
    region = next(iter(region_values)) if len(region_values) == 1 else ""
    city = next(iter(city_values)) if len(city_values) == 1 else ""
    if not region:
        flags.append("РЕГИОН_НЕ_ОПРЕДЕЛЕН")
    if not city:
        flags.append("ГОРОД_НЕ_ОПРЕДЕЛЕН")
    return {"region": region, "city": city, "region_evidence": region_evidence,
            "city_evidence": city_evidence, "geo_flags": flags}
