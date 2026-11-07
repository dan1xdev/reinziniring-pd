"""Чтение двух предоставленных форматов без использования аналитической разметки."""

import hashlib
import datetime
from pathlib import Path

from openpyxl import load_workbook

from .normalization import normalize, text


ALIASES = {
    "name": ["Образовательное учреждение", "Название ОУ в исходном Excel"],
    "education": ["Вид образования"],
    "document": ["Тип документа об образовании"],
    "region_raw": ["Регион (как в файле)", "Регион", "Регион ОУ"],
    "location_raw": ["Местонахождение ОУ", "Местонахождение"],
    "city_raw": ["Населённый пункт", "Населенный пункт", "Город"],
    "year": ["Год завершения", "Год окончания"],
    "case_id": ["ID"],
    "original_row": ["Строка в Excel"],
    "inn_raw": ["ИНН", "ИНН ОУ"],
    "personal_file_number": ["Личное дело.Номер", "Номер личного дела", "Личное дело"],
    "student_name": ["ФИО", "ФИО ученика", "ФИО поступающего"],
    "school_address_raw": ["Адрес", "Адрес ОУ", "Адрес школы"],
    "postal_code_raw": ["Индекс", "Почтовый индекс"],
}

REQUIRED_GROUPS = [("name",), ("education", "document")]

REQUIRED_LABELS = {
    "name": "Образовательное учреждение",
    "education": "Вид образования",
    "document": "Тип документа об образовании",
    "student_name": "ФИО",
}

SERVICE_SHEETS = ["ВУЗы", "Не нашли"]


def personal_file_number(cell):
    # Сохраняем ведущие нули, в том числе заданные простым числовым форматом Excel.
    value = cell.value
    if isinstance(value, (int, float)) and not isinstance(value, bool) and value == int(value):
        number_format = cell.number_format
        if number_format and set(number_format) == {"0"}:
            return str(int(value)).zfill(len(number_format))
    return text(value)


def _missing_message(filename, sheet, seen, header_found):
    found = [REQUIRED_LABELS[f] for f in seen if f in REQUIRED_LABELS]
    missing = []
    for group in REQUIRED_GROUPS:
        if not any(f in seen for f in group):
            missing.append(" или ".join(f"«{REQUIRED_LABELS[f]}»" for f in group))
            
    parts = [f"{filename}/«{sheet}»:"]
    if not header_found:
        parts.append("не найдена строка заголовков.")
    if found:
        parts.append("найдены колонки: " + ", ".join(f"«{f}»" for f in found) + ".")
    if missing:
        parts.append("не найдены обязательные колонки: " + "; ".join(missing) + ".")
    return " ".join(parts)


def _check_damaged_cell(cell, value):
    flags = []
    if value is None:
        return flags
    
    if isinstance(value, (datetime.datetime, datetime.date)):
        flags.append("ПОВРЕЖДЁН_КОД_EXCEL")
        
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        if abs(value) > 10000:
            flags.append("ПОВРЕЖДЁН_КОД_EXCEL")
            
    elif isinstance(value, str) and ("E+" in value or "e+" in value):
        flags.append("ПОВРЕЖДЁН_КОД_EXCEL")
        
    elif cell.number_format and any(char in cell.number_format.upper() for char in ("Y", "M", "D", "H")):
        flags.append("ПОВРЕЖДЁН_КОД_EXCEL")
        
    return flags


def dedupe_personal_files(records, sources):
    # Удаляет полные повторы, а конфликты помечает флагом input_conflict
    by_number = {}
    unique_records = []
    removed_count = 0

    for r in records:
        r["input_conflict"] = False 
        num = r.get("personal_file_number")
        
        if not num:
            unique_records.append(r)
            continue
            
        if num not in by_number:
            by_number[num] = r
            unique_records.append(r)
        else:
            existing = by_number[num]
            is_exact_duplicate = (
                r.get("student_name") == existing.get("student_name") and
                r.get("name") == existing.get("name") and
                r.get("education") == existing.get("education") and
                r.get("document") == existing.get("document")
            )

            if is_exact_duplicate:
                removed_count += 1
            else:
                r["input_conflict"] = True
                existing["input_conflict"] = True
                unique_records.append(r)

    if sources:
        sources[-1]["duplicates_removed"] = sources[-1].get("duplicates_removed", 0) + removed_count

    return unique_records


def read_records(paths, should_cancel=None, sheet_name=None):
    records, sources = [], []
    seen_paths = set()
    
    for filename in paths:
        if should_cancel:
            should_cancel()
        path = Path(filename).expanduser().resolve()
        if path in seen_paths:
            continue
        seen_paths.add(path)
        
        if path.suffix.lower() != ".xlsx":
            raise ValueError(f"Поддерживается .xlsx: {path.name}")
            
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        source_id = hashlib.sha256(str(path).encode()).hexdigest()[:8]
        workbook = load_workbook(path, read_only=True, data_only=False)
        
        count = 0
        skipped = []
        diagnostics = []
        requested_sheet_found = False
        
        try:
            for sheet in workbook:
                if sheet_name:
                    if sheet.title != sheet_name:
                        skipped.append(sheet.title)
                        continue
                    requested_sheet_found = True
                else:
                    if "100 записей" in workbook.sheetnames and sheet.title != "100 записей":
                        skipped.append(sheet.title)
                        continue
                    if sheet.title in SERVICE_SHEETS:
                        skipped.append(sheet.title)
                        continue
                    
                    title_lower = sheet.title.lower()
                    if "копия" in title_lower or "коллективу" in title_lower:
                        skipped.append(sheet.title)
                        continue

                found, info = read_sheet(sheet, path.name, source_id, should_cancel)
                if found is None:
                    skipped.append(sheet.title)
                    diagnostics.append(info)
                    continue
                records.extend(found)
                count += len(found)
        finally:
            workbook.close()
            
        if sheet_name and not requested_sheet_found:
            available = ", ".join(f"«{s}»" for s in workbook.sheetnames)
            raise ValueError(f"{path.name}: лист «{sheet_name}» не найден. Доступные листы: {available}")

        if not count:
            if diagnostics:
                info = diagnostics[0]
                raise ValueError(_missing_message(path.name, info["sheet"], info["seen"], info["header"]))
            raise ValueError(f"{path.name}: заголовки найдены, но нет строк с данными")
            
        sources.append({"file": path.name, "source_id": source_id, "sha256": digest,
                        "rows": count, "skipped_sheets": skipped})

        records = dedupe_personal_files(records, sources)

    return records, sources


def read_sheet(sheet, filename, source_id, should_cancel=None):
    mapping = None
    records = []
    seen = {}
    header_found = False
    
    for row_number, cells in enumerate(sheet.iter_rows(), 1):
        if should_cancel and row_number % 100 == 1:
            should_cancel()
        values = [cell.value for cell in cells]
        
        if mapping is None:
            if row_number > 30:
                return None, {"sheet": sheet.title, "seen": seen, "header": header_found}
                
            headers = [normalize(value) for value in values]
            candidate = {}
            for field, aliases in ALIASES.items():
                indices = [i for i, header in enumerate(headers)
                           if header in {normalize(alias) for alias in aliases}]
                if len(indices) > 1:
                    raise ValueError(
                        f"{filename}/«{sheet.title}»: дубли колонок "
                        f"«{REQUIRED_LABELS.get(field, field)}»"
                    )
                if indices:
                    candidate[field] = indices[0]
                    seen[field] = values[indices[0]]
                    
            if "name" in candidate and ("education" in candidate or "document" in candidate):
                mapping = candidate
                header_found = True
            continue
            
        if all(value is None or text(value) == "" for value in values):
            continue
            
        # Очищаем целочисленные float (например, 2020.0 -> "2020")
        cleaned_values = {}
        for field, index in mapping.items():
            val = values[index]
            if isinstance(val, float) and val == int(val):
                cleaned_values[field] = str(int(val))
            else:
                cleaned_values[field] = val

        record = {field: text(val) for field, val in cleaned_values.items()}
        record["name"] = "" if values[mapping["name"]] is None else str(values[mapping["name"]])
        record["personal_file_number"] = (
            personal_file_number(cells[mapping["personal_file_number"]])
            if "personal_file_number" in mapping else ""
        )
        if "postal_code_raw" in mapping:
            record["postal_code_raw"] = personal_file_number(cells[mapping["postal_code_raw"]])
            
        formula_fields = [field for field, index in mapping.items() if cells[index].data_type == "f"]
        for field in formula_fields:
            record[field] = ""

        # Проверяем ячейки на повреждение Excel
        input_flags = []
        for index in mapping.values():
            input_flags.extend(_check_damaged_cell(cells[index], cells[index].value))
            
        record.update({"record_id": f"{source_id}:{sheet.title}:{row_number}",
                       "source_file": filename, "source_sheet": sheet.title,
                       "source_row": row_number, "formula_fields": formula_fields,
                       "input_flags": input_flags})
        records.append(record)
        
    if mapping is None:
        return None, {"sheet": sheet.title, "seen": seen, "header": header_found}
    return records, {"sheet": sheet.title, "seen": seen, "header": header_found}
