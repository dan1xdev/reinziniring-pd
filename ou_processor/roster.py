"""Связь учреждения, адреса и ученика без подмены подтверждения кандидатом."""

from collections import defaultdict

from .normalization import text


def address_for_group(records, candidates):
    addresses = sorted({text(r.get("school_address_raw")) for r in records if text(r.get("school_address_raw"))})
    postcodes = sorted({text(r.get("postal_code_raw")) for r in records if text(r.get("postal_code_raw"))})
    base = {"school_address": "", "postal_code": "", "address_source": "",
            "address_status": "АДРЕС_НЕ_ОПРЕДЕЛЕН", "candidate_inn": "",
            "candidate_name": "", "source_addresses": addresses}
    if len(addresses) > 1:
        return {**base, "address_source": "Исходные таблицы", "address_status": "КОНФЛИКТ_АДРЕСОВ_В_ИСТОЧНИКЕ"}
    if addresses:
        return {**base, "school_address": addresses[0], "postal_code": postcodes[0] if len(postcodes) == 1 else "",
                "address_source": "Исходная таблица; адрес не проверен",
                "address_status": "АДРЕС_ИЗ_ИСТОЧНИКА_ПРОВЕРИТЬ"}
    if len(candidates) > 1:
        return {**base, "address_source": "DaData", "address_status": "НЕСКОЛЬКО_КАНДИДАТОВ_ВЫБРАТЬ_ВРУЧНУЮ"}
    if len(candidates) == 1 and text(candidates[0].get("address")):
        candidate = candidates[0]
        if candidate.get("flags") or candidate.get("legal_status") != "ACTIVE":
            return {**base, "address_source": "DaData", "address_status": "КАНДИДАТ_С_ПРОТИВОРЕЧИЯМИ_ПРОВЕРИТЬ"}
        return {**base, "school_address": text(candidate["address"]),
                "postal_code": text(candidate.get("postal_code")),
                "candidate_inn": text(candidate.get("candidate_inn")),
                "candidate_name": text(candidate.get("candidate_name")),
                "address_source": "Юридический адрес кандидата DaData",
                "address_status": "АДРЕС_КАНДИДАТА_ПРОВЕРИТЬ"}
    return base


def attach_roster(result):
    records_by_id = {record["record_id"]: record for record in result["records"]}
    candidates_by_group = defaultdict(list)
    for candidate in result.get("dadata", {}).get("candidates", []):
        candidates_by_group[candidate.get("group_id", "")].append(candidate)
    roster = []
    for group in result["groups"]:
        members = [records_by_id[identifier] for identifier in group["record_ids"]]
        address = address_for_group(members, candidates_by_group[group["group_id"]])
        students = defaultdict(list)
        names_by_number = defaultdict(set)
        for record in members:
            name = text(record.get("student_name"))
            number = text(record.get("personal_file_number"))
            if name and number:
                names_by_number[number].add(name)
            # Без полной пары ФИО + личное дело повтор не доказывает того же ученика.
            key = (name, number) if name and number else (name, number, record["record_id"])
            students[key].append(record)
        for key, student_records in students.items():
            name, number = key[:2]
            flags = []
            if not name:
                flags.append("НЕТ_ФИО_В_ИСТОЧНИКЕ")
            if not number:
                flags.append("НЕТ_НОМЕРА_ЛИЧНОГО_ДЕЛА")
            if len(names_by_number[number]) > 1:
                flags.append("РАЗНЫЕ_ФИО_ДЛЯ_ОДНОГО_ДЕЛА")
            roster.append({"group_id": group["group_id"], "school_name": group["name"],
                           **address, "student_name": name, "personal_file_number": number,
                           "student_status": "; ".join(flags) or "ФИО_И_НОМЕР_ЕСТЬ",
                           "group_status": group["status"],
                           "source_record_ids": [r["record_id"] for r in student_records]})
    result["school_students"] = sorted(roster, key=lambda r: (not bool(r["student_name"]),
                                                              not bool(r["school_address"]),
                                                              r["school_name"], r["personal_file_number"]))
    selected = [r for r in result["records"] if r["selection"] == "ОТОБРАНО"]
    result["summary"].update({
        "selected_rows_with_student_name": sum(bool(text(r.get("student_name"))) for r in selected),
        "selected_rows_without_student_name": sum(not text(r.get("student_name")) for r in selected),
        "groups_with_address_available": len({r["group_id"] for r in roster if r["school_address"]}),
        "groups_without_address": len({r["group_id"] for r in roster if not r["school_address"]}),
        "school_student_rows": len(roster),
    })
