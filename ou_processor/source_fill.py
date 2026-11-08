"""Заполнение целевых полей в отдельной копии исходных книг Excel."""

from collections import Counter, defaultdict
from copy import copy
from pathlib import Path
import re
import shutil

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .normalization import normalize


TARGET_FIELDS = {
    "official_name": {
        "header": "Полное название ОУ в именительном падеже",
        "aliases": (
            "Полное название ОУ в именительном падеже",
            "Название ОУ в именительном падеже",
            "Официальное название ОУ",
        ),
    },
    "genitive_name": {
        "header": "Название ОУ в родительном падеже",
        "aliases": (
            "Название ОУ в родительном падеже",
            "Полное название ОУ в родительном падеже",
        ),
    },
    "address": {
        "header": "Адрес",
        "aliases": (
            "Адрес",
            "Адрес ОУ",
            "Адрес школы",
            "Адрес образовательного учреждения",
            "Юридический адрес",
        ),
    },
    "postal_code": {
        "header": "Индекс",
        "aliases": (
            "Индекс",
            "Почтовый индекс",
        ),
    },
}

HEADER_IDENTIFIERS = {
    normalize(value)
    for value in (
        "Образовательное учреждение",
        "Название ОУ в исходном Excel",
        "ФИО",
        "ФИО ученика",
    )
}


def _letter_values(letter):
    """Возвращает только значения, которые не отмечены как конфликтные."""
    if not letter:
        return {}

    checks = set(str(letter.get("checks") or "").splitlines())
    official_conflict = "Официальные названия не совпадают" in checks
    address_conflict = "Адреса в источниках не совпадают" in checks

    official_name = "" if official_conflict else str(letter.get("official_name") or "").strip()
    genitive_name = ""
    if official_name and "Родительный падеж требует проверки" not in checks:
        genitive_name = str(letter.get("genitive_name") or "").strip()

    address = "" if address_conflict else str(letter.get("address") or "").strip()
    postal_code = str(letter.get("postal_code") or "").strip()
    if not re.fullmatch(r"\d{6}", postal_code):
        postal_code = ""

    return {
        "official_name": official_name,
        "genitive_name": genitive_name,
        "address": address,
        "postal_code": postal_code,
    }


def _record_values(result):
    """Связывает подтверждённые группы с результатами подготовки писем."""
    letters = result.get("letters_ready", []) + result.get("letters_review", [])
    letters_by_inn = {
        str(row.get("inn") or "").strip(): row
        for row in letters
        if str(row.get("inn") or "").strip()
    }

    values_by_record = {}
    for group in result.get("resolution", {}).get("groups", []):
        selected = group.get("selected")
        if not selected:
            continue

        inn = str(selected.get("candidate_inn") or "").strip()
        letter = letters_by_inn.get(inn)
        if not letter:
            continue

        values = _letter_values(letter)
        for record_id in group.get("record_ids", []):
            values_by_record[record_id] = values

    return values_by_record


def _source_row_entries(result, values_by_record):
    """Разворачивает записи и удалённые дубли в координаты исходных книг."""
    entries = defaultdict(dict)
    records_by_id = {
        record.get("record_id"): record
        for record in result.get("records", [])
    }

    for record in result.get("records", []):
        values = values_by_record.get(record.get("record_id"), {})
        row_meta = {
            "source_id": record.get("source_id", ""),
            "source_file": record.get("source_file", ""),
            "source_path": record.get("source_path", ""),
            "source_sheet": record.get("source_sheet", ""),
            "source_row": record.get("source_row", ""),
            "header_row": record.get("header_row", ""),
        }
        entries[record.get("source_id", "")][(
            row_meta["source_sheet"], row_meta["source_row"]
        )] = {"meta": row_meta, "values": values}

        for duplicate in record.get("duplicate_source_rows", []):
            duplicate_meta = dict(duplicate)
            entries[duplicate_meta.get("source_id", "")][(
                duplicate_meta.get("source_sheet", ""),
                duplicate_meta.get("source_row", ""),
            )] = {"meta": duplicate_meta, "values": values}

    source_by_id = {
        source.get("source_id", ""): source
        for source in result.get("sources", [])
    }
    return entries, source_by_id


def _find_header_row(sheet, preferred=None):
    if preferred and 1 <= int(preferred) <= min(sheet.max_row, 30):
        return int(preferred)

    best_row = None
    best_score = 0
    for row_number in range(1, min(sheet.max_row, 30) + 1):
        headers = {normalize(cell.value) for cell in sheet[row_number] if cell.value is not None}
        score = len(headers & HEADER_IDENTIFIERS)
        if score > best_score:
            best_row, best_score = row_number, score
    if best_row is None:
        raise ValueError(f"Не найдена строка заголовков на листе «{sheet.title}».")
    return best_row


def _target_columns(sheet, header_row):
    header_values = {
        normalize(cell.value): cell.column
        for cell in sheet[header_row]
        if cell.value is not None
    }
    columns = {}

    for field, spec in TARGET_FIELDS.items():
        found = None
        for alias in spec["aliases"]:
            index = header_values.get(normalize(alias))
            if index:
                found = index
                break

        if found is None:
            found = sheet.max_column + 1
            cell = sheet.cell(header_row, found, spec["header"])
            if found > 1:
                previous = sheet.cell(header_row, found - 1)
                if previous.has_style:
                    cell._style = copy(previous._style)
                if previous.alignment:
                    cell.alignment = copy(previous.alignment)
            cell.number_format = "@"
            sheet.column_dimensions[get_column_letter(found)].width = (
                48 if field in {"official_name", "genitive_name", "address"} else 14
            )
            header_values[normalize(spec["header"])] = found

        columns[field] = found

    return columns


def fill_source_copies(result, output):
    """Создаёт отдельные заполненные копии исходных книг, не меняя оригиналы."""
    output = Path(output)
    filled_dir = output / "filled_sources"
    filled_dir.mkdir(parents=True, exist_ok=True)

    values_by_record = _record_values(result)
    entries_by_source, source_by_id = _source_row_entries(result, values_by_record)

    stats = {
        "copies": [],
        "source_rows": 0,
        "written": Counter(),
        "present": Counter(),
        "errors": [],
    }

    for source_id, rows in entries_by_source.items():
        source = source_by_id.get(source_id, {})
        source_path = Path(source.get("path") or next(
            (item["meta"].get("source_path") for item in rows.values()
             if item["meta"].get("source_path")),
            "",
        ))
        if not source_path or not source_path.is_file():
            stats["errors"].append(
                f"Не найдена исходная книга для источника {source.get('file') or source_id}."
            )
            continue

        destination = filled_dir / (
            f"{source_id}_{source_path.stem}_заполнен.xlsx"
        )
        shutil.copy2(source_path, destination)

        book = load_workbook(destination)
        file_rows = 0
        field_present = Counter()
        field_written = Counter()

        try:
            for (sheet_name, row_number), item in rows.items():
                if not sheet_name or not row_number:
                    continue
                if sheet_name not in book.sheetnames:
                    stats["errors"].append(
                        f"{source_path.name}: лист «{sheet_name}» не найден в копии."
                    )
                    continue

                sheet = book[sheet_name]
                header_row = _find_header_row(
                    sheet, item["meta"].get("header_row")
                )
                columns = _target_columns(sheet, header_row)
                row_number = int(row_number)
                file_rows += 1

                for field, column in columns.items():
                    cell = sheet.cell(row_number, column)
                    value = str(item["values"].get(field) or "").strip()
                    # Пустое или неподтверждённое значение не стирает исходные данные.
                    if value:
                        cell.value = value
                        cell.data_type = "s"
                        cell.number_format = "@"
                        field_written[field] += 1
                    if cell.value is not None and str(cell.value).strip():
                        field_present[field] += 1

            book.save(destination)
        finally:
            book.close()

        stats["source_rows"] += file_rows
        stats["written"].update(field_written)
        stats["present"].update(field_present)
        stats["copies"].append({
            "source": source_path.name,
            "file": destination.name,
            "rows": file_rows,
        })

    stats["copy_count"] = len(stats["copies"])
    stats["written"] = dict(stats["written"])
    stats["present"] = dict(stats["present"])
    return stats
