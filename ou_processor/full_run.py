
"""Полный прогон и одна итоговая строка на ИНН."""

import argparse
from collections import defaultdict
from datetime import datetime, timezone
import getpass
import json
import os
from pathlib import Path
import re
import sys
import zipfile
from xml.etree.ElementTree import ParseError

from openpyxl import Workbook
from openpyxl.utils.exceptions import InvalidFileException

from .dadata import DaDataClient
from .export import add_table, RECORD_COLUMNS
from .letters import (
    prepare_letters,
    postal_details,
    LETTER_COLUMNS,
)
from .input import list_sheets, read_records
from .normalization import normalize
from .pipeline import process
from .workflow import search, write_log


def aggregate(result):
    records = {
        r["record_id"]: r
        for r in result["records"]
    }

    buckets = defaultdict(list)
    review = []
    candidates_by_group = defaultdict(list)

    for candidate in result["resolution"].get("candidates", []):
        candidates_by_group[candidate["group_id"]].append(candidate)

    for resolution in result["resolution"]["groups"]:
        selected = resolution["selected"]
        members = [
            records[i] for i in resolution["record_ids"]
        ]

        if selected:
            buckets[selected["candidate_inn"]].append(
                (resolution, members)
            )
        else:
            options = sorted(
                candidates_by_group[resolution["group_id"]],
                key=lambda c: (
                    c["eligible"],
                    c["score"],
                ),
                reverse=True,
            )

            top = {}

            for option in options:
                top.setdefault(
                    option["candidate_inn"], option
                )

            top = list(top.values())[:3]

            reasons = sorted({
                reason
                for option in top
                for reason in option["reasons"]
            })

            review.append({
                **resolution,
                "region": members[0]["region"],
                "city": members[0]["city"],
                "students": "\n".join(
                    dict.fromkeys(
                        r.get("student_name", "")
                        for r in members
                        if r.get("student_name", "")
                    )
                ),
                "files": "\n".join(
                    dict.fromkeys(
                        r["personal_file_number"]
                        for r in members
                        if r["personal_file_number"]
                    )
                ),
                "reason": "\n".join(reasons) or (
                    "Несколько подходящих ИНН"
                    if options else "Кандидаты не найдены"
                ),
                "top_candidates": "\n\n".join(
                    f"{c['candidate_inn']} | "
                    f"{c['candidate_name']} | "
                    f"{c['address']} | {c['source']}"
                    for c in top
                ),
                "record_ids": "\n".join(
                    resolution["record_ids"]
                ),
            })

        for record in members:
            record["resolved_inn"] = (
                selected["candidate_inn"]
                if selected else ""
            )
            record["resolution_status"] = resolution["status"]
            record["resolution_source"] = (
                selected["source"]
                if selected else ""
            )

    schools = []
    students = []

    for inn, groups in buckets.items():
        members = [
            record
            for _, group in groups
            for record in group
        ]

        selected = [
            group["selected"]
            for group, _ in groups
        ]

        preferred = [
            candidate
            for candidate in selected
            if candidate["source"] == "DaData"
        ] or selected

        candidate = preferred[0]

        addresses = {
            c["address"]
            for c in preferred
            if c["address"]
        }

        address = ""

        if len(addresses) == 1:
            source_address = next(iter(addresses))
            address_candidate = next(
                c for c in preferred
                if c["address"] == source_address
            )
            _, address, _ = postal_details(
                address_candidate
            )

        pairs = {}

        for record in members:
            name = record.get("student_name", "")
            personal_file = record["personal_file_number"]

            if name or personal_file:
                key = (
                    (normalize(name), personal_file)
                    if name and personal_file
                    else (record["record_id"],)
                )

                pairs.setdefault(
                    key,
                    {
                        "name": name,
                        "file": personal_file,
                        "record_ids": [],
                    },
                )["record_ids"].append(
                    record["record_id"]
                )

        people = sorted(
            pairs.values(),
            key=lambda p: (
                p["name"].casefold(),
                p["file"],
            ),
        )

        names_per_file = defaultdict(set)

        for person in people:
            if person["file"]:
                names_per_file[person["file"]].add(
                    normalize(person["name"])
                )

            students.append({
                "school_name": candidate["candidate_name"],
                "inn": inn,
                "address": address,
                "student": person["name"],
                "file": person["file"],
                "record_ids": "\n".join(
                    person["record_ids"]
                ),
            })

        notes = []

        if any(
            len(names) > 1
            for names in names_per_file.values()
        ):
            notes.append(
                "Один номер дела связан с разными ФИО — проверить"
            )

        if len(addresses) != 1:
            notes.append(
                "Адрес отсутствует или есть противоречие — проверить"
            )

        if candidate["source"] != "DaData":
            notes.append(
                "Адрес реестра аккредитации; "
                "актуальность юридического адреса "
                "не подтверждена"
            )

        missing = sum(
            not record.get("student_name", "")
            or not record["personal_file_number"]
            for record in members
        )

        if missing:
            notes.append(
                f"В {missing} исходных строках "
                "нет ФИО и/или номера дела"
            )

        schools.append({
            "school_name": candidate["candidate_name"],
            "address": address,
            "students": "\n".join(
                person["name"] or "ФИО отсутствует"
                for person in people
            ),
            "files": "\n".join(
                person["file"] or "Номер отсутствует"
                for person in people
            ),
            "student_count": len(people),
            "inn": inn,
            "ogrn": candidate["candidate_ogrn"],
            "postal_code": candidate["postal_code"],
            "region": candidate["region"],
            "source": candidate["source"],
            "status": "Сопоставлено автоматически",
            "notes": "\n".join(notes),
            "aliases": "\n".join(
                sorted({
                    record["name"]
                    for record in members
                })
            ),
            "row_count": len(members),
            "record_ids": "\n".join(
                record["record_id"]
                for record in members
            ),
        })

    schools.sort(
        key=lambda row: (
            not bool(row["student_count"]),
            row["school_name"].casefold(),
        )
    )

    result.update(
        schools=schools,
        students=students,
        review=review,
    )

    groups = result["resolution"]["groups"]

    result["full_summary"] = {
        "Входных строк": len(result["records"]),
        "Отобрано строк": result["summary"]["selected_rows"],
        "Сопоставлено исходных строк": sum(
            row["row_count"] for row in schools
        ),
        "Исходных строк на проверку": sum(
            len(group["record_ids"])
            for group in groups
            if not group["selected"]
        ),
        "Учреждений без повторов по ИНН": len(schools),
        "Учреждений с адресом": sum(
            bool(row["address"]) for row in schools
        ),
        "Учреждений с данными учеников": sum(
            bool(row["student_count"]) for row in schools
        ),
        "Пар ФИО и номер дела после объединения": len(
            students
        ),
        "Групп на проверку": len(review),
        "Групп сопоставлено через DaData": sum(
            bool(group["selected"])
            and group["selected"]["source"] == "DaData"
            for group in groups
        ),
        "Групп сопоставлено через реестр": sum(
            bool(group["selected"])
            and group["selected"]["source"] != "DaData"
            for group in groups
        ),
        "Групп проверено по реестру": sum(
            group["registry_used"]
            for group in groups
        ),
        "Запросов DaData в этом запуске":
            result["resolution"]["actual_requests"],
        "Успешных уникальных запросов с учётом кэша": len({
            row["query"]
            for row in result["resolution"]["requests"]
            if row["status"] == "OK"
        }),
        "Ошибка или пропуск запроса": sum(
            row["status"] != "OK"
            for row in result["resolution"]["requests"]
        ),
        "Строк без данных ученика": sum(
            not row.get("student_name", "")
            or not row["personal_file_number"]
            for row in result["records"]
        ),
    }

    return result


def save(result, output, overrides=None):
    prepare_letters(result, overrides=overrides)

    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)

    book = Workbook()
    book.remove(book.active)

    main = add_table(
        book,
        "Учреждения",
        result["schools"],
        [
            ("school_name", "Школа / учреждение"),
            ("address", "Адрес организации"),
            ("students", "ФИО учеников"),
            ("files", "Номера личных дел"),
            ("student_count", "Учеников / дел"),
            ("postal_code", "Индекс"),
            ("inn", "ИНН"),
            ("source", "Источник адреса"),
            ("notes", "Примечания"),
            ("region", "Регион"),
            ("ogrn", "ОГРН"),
            ("status", "Статус сопоставления"),
            ("row_count", "Исходных строк"),
        ],
    )

    for col, width in {
        "A": 65,
        "B": 55,
        "C": 35,
        "D": 23,
        "E": 15,
        "F": 12,
        "G": 18,
        "H": 28,
        "I": 65,
    }.items():
        main.column_dimensions[col].width = width

    add_table(
        book,
        "Ученики и дела",
        result["students"],
        [
            ("school_name", "Школа / учреждение"),
            ("address", "Адрес"),
            ("student", "ФИО ученика"),
            ("file", "Номер личного дела"),
            ("inn", "ИНН"),
            ("record_ids", "Исходные записи"),
        ],
    )

    add_table(
        book,
        "Письма к печати",
        result["letters_ready"],
        LETTER_COLUMNS,
    )

    add_table(
        book,
        "Письма проверить",
        result["letters_review"],
        LETTER_COLUMNS,
    )

    review_sheet = add_table(
        book,
        "Проверить",
        result["review"],
        [
            ("source_name", "Исходное название"),
            ("region", "Регион"),
            ("city", "Город"),
            ("status", "Статус"),
            ("students", "ФИО"),
            ("files", "Номер дела"),
            ("reason", "Что уточнить"),
            ("top_candidates", "Первые три кандидата"),
            ("candidate_count", "Кандидатов"),
            ("registry_used", "Проверен реестр"),
            ("registry_truncated", "Выдача ограничена"),
            ("group_id", "Группа"),
            ("record_ids", "Исходные записи"),
        ],
    )

    review_sheet.column_dimensions["H"].width = 110

    for row in range(2, review_sheet.max_row + 1):
        review_sheet.row_dimensions[row].height = (
            240
            if result["review"][row - 2]["top_candidates"]
            else 65
        )

    summary = [
        {"metric": key, "value": value}
        for key, value in result["full_summary"].items()
    ]

    summary += [
        {
            "metric": "Единица группировки",
            "value": (
                "Одно юридическое лицо (ИНН), "
                "независимо от написания. "
                "Корпуса не разделены."
            ),
        },
        {
            "metric": "Адрес",
            "value": (
                "DaData: юридический адрес. "
                "Реестр: адрес организации из реестра "
                "аккредитации, требует сверки."
            ),
        },
        {
            "metric": "ФИО и номера дел",
            "value": (
                "Строки в соседних ячейках совпадают "
                "по порядку. Полные пары удалены "
                "как дубли только внутри одного ИНН."
            ),
        },
        {
            "metric": "Дата обработки UTC",
            "value": result["created_at_utc"],
        },
        {
            "metric": "Дата снимка реестра",
            "value": result["resolution"]["registry"].get(
                "snapshot_date", ""
            ),
        },
        {
            "metric": "DaData",
            "value": "https://dadata.ru/api/suggest/party/",
        },
        {
            "metric": "Резервный реестр",
            "value": (
                "https://isga.obrnadzor.gov.ru/"
                "accredreestr/opendata/"
            ),
        },
    ]

    add_table(
        book,
        "Сводка",
        summary,
        [
            ("metric", "Показатель"),
            ("value", "Значение"),
        ],
    )

    add_table(
        book,
        "Исходные записи",
        result["records"],
        [
            ("resolved_inn", "Сопоставленный ИНН"),
            ("resolution_status", "Результат сопоставления"),
            ("resolution_source", "Источник сопоставления"),
            *RECORD_COLUMNS,
        ],
    )

    add_table(
        book,
        "Запросы DaData",
        result["resolution"]["requests"],
        [
            ("group_id", "Группа"),
            ("query", "Запрос"),
            ("status", "Статус"),
            ("candidate_count", "Кандидатов"),
            ("cached", "Из кэша"),
            ("timestamp_utc", "Время UTC"),
        ],
    )

    if 'dadata_checks' in result:
        add_table(book, 'Сверка DaData', result['dadata_checks'], [
            ('group_id', 'Группа'), ('registry_inn', 'ИНН из реестра'),
            ('dadata_inn', 'ИНН DaData'), ('status', 'Результат сверки')])
        add_table(book, 'Кандидаты сверки DaData', result['dadata_candidates'], [
            ('group_id', 'Группа'), ('candidate_inn', 'ИНН'), ('candidate_name', 'Название'),
            ('address', 'Адрес'), ('score', 'Баллы'), ('eligible', 'Прошёл правила'),
            ('reasons', 'Причины проверки'), ('source_url', 'Источник')])

    candidate_groups = defaultdict(list)

    for candidate in result["resolution"]["candidates"]:
        candidate_groups[
            candidate["group_id"]
        ].append(candidate)

    shortlist = []

    for group in candidate_groups.values():
        group.sort(
            key=lambda c: (
                c["eligible"],
                c["score"],
                c["legal_status"] in {
                    "ACTIVE",
                    "Действующее",
                },
            ),
            reverse=True,
        )

        distinct = {}

        for candidate in group:
            distinct.setdefault(
                candidate["candidate_inn"],
                candidate,
            )

        shortlist.extend(
            list(distinct.values())[:5]
        )

    add_table(
        book,
        "Кандидаты",
        shortlist,
        [
            ("group_id", "Группа"),
            ("source", "Источник"),
            ("candidate_name", "Название"),
            ("candidate_inn", "ИНН"),
            ("address", "Адрес"),
            ("legal_status", "Статус источника"),
            ("eligible", "Прошёл правила"),
            ("score", "Совпадение слов, %"),
            ("reasons", "Причины отказа"),
            ("source_url", "Источник в интернете"),
            ("certificate_id", "ID записи реестра"),
        ],
    )

    add_table(
        book,
        "Источники",
        result["sources"],
        [
            ("file", "Файл"),
            ("rows", "Строк"),
            ("sha256", "SHA-256"),
            ("source_id", "Код источника"),
        ],
    )

    temporary_book = output / "schools.tmp.xlsx"
    book.save(temporary_book)
    temporary_book.replace(output / "schools.xlsx")

    temporary_json = output / "result.tmp.json"
    temporary_json.write_text(
        json.dumps(
            result,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    temporary_json.replace(output / "result.json")
    write_log(result, output)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", type=Path)
    parser.add_argument(
        "--output",
        type=Path,
        help='Новая или пустая папка результата',
    )
    parser.add_argument(
        "--cache",
        type=Path,
        default=Path(__file__).resolve().parents[2] / 'ui_cache' / 'dadata.json',
    )
    parser.add_argument(
        "--registry",
        type=Path,
        default=(
            Path(__file__).resolve().parents[2]
            / "registry_cache"
        ),
    )
    parser.add_argument('--config', type=Path, default=Path(__file__).resolve().parents[1] / 'config.json')
    parser.add_argument('--sheet', help='Название рабочего листа во всех входных книгах')
    parser.add_argument('--list-sheets', action='store_true', help='Показать листы и завершить без обработки')
    parser.add_argument('--check-dadata', '--online', dest='check_dadata', action='store_true',
                        help='Отдельно сверить результат в DaData; не менять выбранные ИНН')
    parser.add_argument("--ask-token", action="store_true")

    args = parser.parse_args(argv)

    if args.ask_token and not args.check_dadata:
        parser.error('--ask-token используется только с --check-dadata')
    if args.list_sheets:
        try:
            for path in args.inputs:
                print(json.dumps({'file': path.name, 'sheets': list_sheets(path)}, ensure_ascii=False))
        except (OSError, ValueError, KeyError, zipfile.BadZipFile, ParseError, InvalidFileException) as error:
            print(f'Не удалось прочитать список листов: {error}', file=sys.stderr)
            return 1
        return 0
    if args.output is None:
        parser.error('Укажите --output для обработки')
    if args.output.exists() and not args.output.is_dir():
        parser.error('--output должен указывать на папку')

    if args.output.exists() and any(args.output.iterdir()):
        parser.error(
            "Папка результата должна быть новой или пустой"
        )

    try:
        config = json.loads(args.config.read_text(encoding='utf-8'))
        records, sources = read_records(args.inputs, sheet_name=args.sheet)
        result = process(records, config)
        result.update(sources=sources, created_at_utc=datetime.now(timezone.utc).isoformat())
        client = None
        if args.check_dadata:
            token = getpass.getpass('API-ключ DaData: ') if args.ask_token else os.environ.get('DADATA_API_KEY', '')
            if not token.strip():
                parser.error('Для --check-dadata нужен DADATA_API_KEY или --ask-token')
            client = DaDataClient(token.strip())

        def progress(done, total):
            print(f'Обработано групп: {done}/{total}', file=sys.stderr, flush=True)

        search(result, args.registry, config, cache=args.cache, client=client, on_progress=progress)
        aggregate(result)
        save(result, args.output)
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ParseError, InvalidFileException) as error:
        print(f'Ошибка обработки: {error}', file=sys.stderr)
        return 1

    for warning in result['warnings']:
        print(warning, file=sys.stderr)
    print(f'Excel: {args.output / "schools.xlsx"}')
    print(f'Журнал: {args.output / "processing_log.json"}')

    print(
        json.dumps(
            result["full_summary"],
            ensure_ascii=False,
            indent=2,
        )
    )

    return (
        2
        if result["full_summary"]["Ошибка или пропуск запроса"]
        else 0
    )


if __name__ == "__main__":
    raise SystemExit(main())
