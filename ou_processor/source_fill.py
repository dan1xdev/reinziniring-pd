"""Заполнение целевых полей в отдельной копии исходных книг Excel."""

from collections import Counter, defaultdict
from pathlib import Path
import re
import posixpath
from xml.etree import ElementTree
from xml.parsers import expat
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, coordinate_to_tuple, range_boundaries

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
    official_conflict = bool(checks & {
        "Официальные названия не совпадают",
        "Регистр восстановлен автоматически — проверить",
        "Регистр отдельных слов требует проверки",
        "Несбалансированные кавычки в названии",
    })
    address_conflict = "Адреса в источниках не совпадают" in checks

    official_name = "" if official_conflict else str(letter.get("official_name") or "").strip()
    genitive_name = ""
    if official_name and "Родительный падеж требует проверки" not in checks:
        genitive_name = str(letter.get("genitive_name") or "").strip()

    address = "" if address_conflict else str(letter.get("address") or "").strip()
    postal_code = "" if "Индексы в источниках не совпадают" in checks else str(
        letter.get("postal_code") or ""
    ).strip()
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
        and row.get("match_status", "АВТО") in {"АВТО", "ПРОВЕРЕНО"}
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


def _target_columns(sheet, header_row, updates):
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
            updates[cell.coordinate] = spec["header"]
            header_values[normalize(spec["header"])] = found

        columns[field] = found

    return columns


def _sheet_parts(archive):
    """Разрешает имена листов через связи, не полагаясь на их порядок."""
    relations = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        item.attrib["Id"]: posixpath.normpath(
            posixpath.join("xl", item.attrib["Target"])
        ).lstrip("/")
        for item in relations
    }
    workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    relation = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    return {
        sheet.attrib["name"]: targets[sheet.attrib[relation]]
        for sheet in workbook.findall(f"{{{namespace}}}sheets/{{{namespace}}}sheet")
    }


def _sheet_spans(xml):
    """Находит байтовые границы элементов без повторной сериализации XML."""
    parser = expat.ParserCreate()
    stack = []
    root = None
    opening_tag = re.compile(br"(?:[^>\"']|\"[^\"]*\"|'[^']*')*>")

    def start(tag, attributes):
        nonlocal root
        local = tag.rsplit(":", 1)[-1]
        parent = stack[-1] if stack else None
        allowed = {
            "worksheet": {"sheetData", "dimension"},
            "sheetData": {"row"},
            "row": {"c"},
        }
        if (not stack and local == "worksheet") or (
            parent and local in allowed.get(parent["local"], set())
        ):
            offset = parser.CurrentByteIndex
            opening = opening_tag.match(xml, offset)
            node = {
                "tag": tag, "local": local, "attrs": attributes, "start": offset,
                "open_end": opening.end(), "children": [],
            }
            if parent:
                parent["children"].append(node)
            else:
                root = node
            stack.append(node)
        else:
            stack.append(None)

    def end(tag):
        node = stack.pop()
        if node is None:
            return
        self_closing = xml[node["start"]:node["open_end"]].rstrip().endswith(b"/>")
        node["close_start"] = node["open_end"] if self_closing else parser.CurrentByteIndex
        node["end"] = node["open_end"] if self_closing else xml.index(b">", parser.CurrentByteIndex) + 1
        node["self_closing"] = self_closing

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(xml, True)
    return root


def _xml_tag(tag, attributes):
    return "<" + tag + "".join(f" {key}={quoteattr(value)}" for key, value in attributes.items()) + ">"


def _inline_cell(coordinate, value, prefix, attributes=None):
    attributes = {**(attributes or {}), "r": coordinate, "t": "inlineStr"}
    text = escape(value).replace("\r", "&#13;")
    return (
        _xml_tag(prefix + "c", attributes)
        + f'<{prefix}is><{prefix}t xml:space="preserve">{text}</{prefix}t></{prefix}is></{prefix}c>'
    ).encode("utf-8")


def _apply_xml_edits(xml, edits):
    # Объединяем вставки в одну позицию, чтобы сохранить порядок новых ячеек.
    combined = defaultdict(bytes)
    for start, end, content in edits:
        combined[start, end] += content
    chunks = []
    position = 0
    for (start, end), content in sorted(combined.items()):
        chunks.extend((xml[position:start], content))
        position = end
    chunks.append(xml[position:])
    return b"".join(chunks)


def _patch_sheet(xml, updates, dimension):
    root = _sheet_spans(xml)
    sheet_data = next(node for node in root["children"] if node["local"] == "sheetData")
    prefix = sheet_data["tag"][:-len("sheetData")]
    rows = {int(row["attrs"]["r"]): row for row in sheet_data["children"]}
    grouped = defaultdict(dict)
    for coordinate, value in updates.items():
        grouped[coordinate_to_tuple(coordinate)[0]][coordinate] = value
    edits = []
    for number, values in sorted(grouped.items()):
        row = rows.get(number)
        cells = {cell["attrs"]["r"]: cell for cell in row["children"]} if row else {}
        additions = []
        for coordinate, value in sorted(values.items(), key=lambda item: coordinate_to_tuple(item[0])):
            cell = cells.get(coordinate)
            content = _inline_cell(coordinate, value, prefix, cell["attrs"] if cell else None)
            if cell:
                edits.append((cell["start"], cell["end"], content))
            elif row and not row["self_closing"]:
                later = [node["start"] for ref, node in cells.items()
                         if coordinate_to_tuple(ref)[1] > coordinate_to_tuple(coordinate)[1]]
                offset = min(later) if later else row["close_start"]
                edits.append((offset, offset, content))
            else:
                additions.append(content)
        if additions:
            tag = row["tag"] if row else prefix + "row"
            attributes = row["attrs"] if row else {"r": str(number)}
            content = _xml_tag(tag, attributes).encode() + b"".join(additions) + f"</{tag}>".encode()
            if row:
                edits.append((row["start"], row["end"], content))
            else:
                later = [node["start"] for index, node in rows.items() if index > number]
                offset = min(later) if later else sheet_data["close_start"]
                edits.append((offset, offset, content))

    existing_dimension = next((node for node in root["children"] if node["local"] == "dimension"), None)
    bounds = range_boundaries(dimension)
    if existing_dimension:
        previous = range_boundaries(existing_dimension["attrs"]["ref"])
        bounds = (min(bounds[0], previous[0]), min(bounds[1], previous[1]),
                  max(bounds[2], previous[2]), max(bounds[3], previous[3]))
    reference = f"{get_column_letter(bounds[0])}{bounds[1]}:{get_column_letter(bounds[2])}{bounds[3]}"
    if existing_dimension:
        attributes = {**existing_dimension["attrs"], "ref": reference}
        content = (_xml_tag(existing_dimension["tag"], attributes)[:-1] + "/>").encode()
        edits.append((existing_dimension["start"], existing_dimension["end"], content))
    # dimension необязателен: при отсутствии Excel вычислит границы по ячейкам.
    return _apply_xml_edits(xml, edits)


def _write_copy(source, destination, updates, book):
    with ZipFile(source) as original, ZipFile(destination, "w") as copied:
        parts = _sheet_parts(original)
        changes = {parts[name]: (values, book[name].calculate_dimension())
                   for name, values in updates.items() if values}
        copied.comment = original.comment
        for item in original.infolist():
            content = original.read(item)
            if item.filename in changes:
                values, dimension = changes[item.filename]
                content = _patch_sheet(content, values, dimension)
            copied.writestr(item, content)


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
        book = load_workbook(source_path)
        updates = defaultdict(dict)
        column_cache = {}
        array_formula_ranges = {}
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
                if sheet_name not in array_formula_ranges:
                    array_formula_ranges[sheet_name] = [
                        range_boundaries(reference) for reference in sheet.array_formulae.values()
                    ]
                header_key = (sheet_name, item["meta"].get("header_row"))
                if header_key not in column_cache:
                    header_row = _find_header_row(sheet, header_key[1])
                    column_cache[header_key] = _target_columns(sheet, header_row, updates[sheet_name])
                columns = column_cache[header_key]
                row_number = int(row_number)
                file_rows += 1

                for field, column in columns.items():
                    cell = sheet.cell(row_number, column)
                    value = str(item["values"].get(field) or "").strip()
                    # Защищаем и формулу, и все ячейки результата формулы массива.
                    in_array_formula = any(
                        first_col <= column <= last_col and first_row <= row_number <= last_row
                        for first_col, first_row, last_col, last_row in array_formula_ranges[sheet_name]
                    )
                    if cell.data_type == "f" or in_array_formula:
                        field_present[field] += 1
                        continue
                    # Пустое или неподтверждённое значение не стирает исходные данные.
                    if value:
                        if field == "address" and cell.value is not None and str(cell.value).strip():
                            if normalize(cell.value) != normalize(value):
                                field_present[field] += 1
                                continue
                        cell.value = value
                        cell.data_type = "s"
                        updates[sheet_name][cell.coordinate] = value
                        field_written[field] += 1
                    if cell.value is not None and str(cell.value).strip():
                        field_present[field] += 1

            _write_copy(source_path, destination, updates, book)
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
