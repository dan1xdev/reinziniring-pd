"""Консервативная география: явные названия и расширяемый справочник регионов и городов."""

import re

from .normalization import normalize, text


def matches(value, entries):
    value = " " + normalize(value) + " "
    return {name for name, aliases in entries.items()
            if any(" " + normalize(alias) + " " in value for alias in [name, *aliases])}


def extract_geography(record, config):
    region_evidence, city_evidence = [], []
    regions = config.get("regions", {})
    cities = config.get("cities", {})

    fields_to_check = [
        ("region_raw", "поле"),
        ("location_raw", "местонахождение"),
        ("school_address_raw", "адрес"),
        ("name", "название")
    ]

    for field, source in fields_to_check:
        raw = record.get(field, "")
        if not raw:
            continue
        for region in sorted(matches(raw, regions)):
            region_evidence.append({"value": region, "source": source})
        for city, details in cities.items():
            # В названии и адресе ищем явный город, а не случайное прилагательное
            aliases = details.get("explicit_aliases", [])
            all_aliases = [city, *aliases] if field != "name" else aliases
            if matches(raw, {city: all_aliases}) and (
                field != "name" or any(re.search(pattern, normalize(raw))
                                       for pattern in details.get("name_patterns", []))
            ):
                city_evidence.append({"value": city, "source": source})
                if details.get("region"):
                    region_evidence.append({"value": details["region"],
                                            "source": source + ":город из словаря"})

    explicit_city = text(record.get("city_raw"))
    if explicit_city:
        matched = matches(explicit_city, {city: [city, *data.get("explicit_aliases", [])]
                                          for city, data in cities.items()})
        if len(matched) == 1:
            matched_city = next(iter(matched))
            city_evidence.append({"value": matched_city, "source": "поле"})
            if cities[matched_city].get("region"):
                region_evidence.append({"value": cities[matched_city]["region"],
                                        "source": "поле:город из словаря"})
        else:
            city_evidence.append({"value": normalize(explicit_city), "source": "поле"})

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

    return {
        "region": region,
        "city": city,
        "region_evidence": region_evidence,
        "city_evidence": city_evidence,
        "geo_flags": flags
    }
