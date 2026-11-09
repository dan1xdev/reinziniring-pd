"""Сохранность входных данных до отбора и сопоставления учреждений."""

import datetime
import json
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

from ou_processor.export import write_result
from ou_processor.input import dedupe_personal_files, read_records
from ou_processor.pipeline import process
from ou_processor.resolution import choose


CONFIG = json.loads((Path(__file__).resolve().parents[1] / "config.json").read_text())


def record(**changes):
    result = {
        "record_id": "a:Данные:2", "source_file": "a.xlsx",
        "source_sheet": "Данные", "source_row": 2,
        "name": "Школа № 1", "student_name": "Иванов Иван",
        "personal_file_number": "001", "education": "Среднее общее",
        "document": "Аттестат о среднем общем образовании",
        "formula_fields": [], "input_flags": [],
    }
    result.update(changes)
    return result


class InputIntegrityTests(unittest.TestCase):
    # Одного совпадения ФИО и названия недостаточно: различия влияют на отбор и адрес.
    def test_duplicate_number_preserves_conflicting_input_fields(self):
        for field, first, second in [
            ("region_raw", "Москва", "Московская область"),
            ("city_raw", "Москва", "Минск"),
            ("location_raw", "Москва", "Минск"),
            ("school_address_raw", "ул. Первая, 1", "ул. Вторая, 2"),
            ("inn_raw", "7707083893", "7736050003"),
            ("postal_code_raw", "101000", "119019"),
            ("year", "2025", "2000"),
            ("document", "Аттестат о среднем общем образовании", "Диплом бакалавра"),
            ("specialty_code", "09.03.01", "09.03.02"),
            ("formula_fields", [], ["year"]),
            ("input_flags", [], ["ПОВРЕЖДЁН_КОД_EXCEL"]),
        ]:
            with self.subTest(field=field):
                rows = [record(**{field: first}), record(record_id="a:Данные:3", **{field: second})]
                kept = dedupe_personal_files(rows, [])
                self.assertEqual(len(kept), 2)
                self.assertTrue(all(r["input_conflict"] for r in kept))

    # При повторе конфликтного варианта сохраняется одна копия каждого варианта.
    def test_repeated_conflicting_variant_is_deduplicated(self):
        rows = [record(name="Школа № 1"), record(name="Школа № 2"), record(name="Школа № 2")]
        kept = dedupe_personal_files(rows, [])
        self.assertEqual([r["name"] for r in kept], ["Школа № 1", "Школа № 2"])
        self.assertTrue(all(r["input_conflict"] for r in kept))

    # Счётчик относится к файлу удалённой строки, а не к последней загруженной книге.
    def test_duplicate_counts_belong_to_their_sources(self):
        sources = [{"source_id": "a", "file": "a.xlsx"}, {"source_id": "b", "file": "b.xlsx"}]
        rows = [record(), record(record_id="a:Другой лист:3", source_row=3),
                record(record_id="b:Данные:2", source_file="b.xlsx")]
        kept = dedupe_personal_files(rows, sources)
        self.assertEqual(len(kept), 1)
        self.assertEqual([s.get("duplicates_removed", 0) for s in sources], [1, 1])

    # СПО при вузе сохраняется в отборе, но совпадение кандидата не отменяет проверку подразделения.
    def test_university_department_cannot_be_resolved_automatically(self):
        item = process([record(name="Колледж Московского университета", education="СПО", document="",
                               inn_raw="7707083893", region_raw="Москва", city_raw="Москва")], CONFIG)["records"][0]
        candidate = {"candidate_name": item["name"], "candidate_inn": "7707083893",
                     "legal_status": "ACTIVE", "branch_type": "MAIN", "source": "DaData",
                     "city": "Москва", "region": "Москва", "address": "Москва, ул. Первая, 1"}
        selected, ranked = choose(item, [candidate])
        self.assertEqual(item["selection"], "ОТОБРАНО")
        self.assertIsNone(selected)
        self.assertIn("ПОДРАЗДЕЛЕНИЕ_ВУЗА_ТРЕБУЕТ_ПРОВЕРКИ", ranked[0]["reasons"])

    # Повреждённый год виден в выгрузке и не отправляет строку в автоматическое сопоставление.
    def test_damaged_excel_year_is_held_and_exported(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Образовательное учреждение", "Вид образования", "Год завершения", "ФИО"])
            sheet.append(["Школа № 1", "Среднее общее", datetime.datetime(2001, 3, 9)])
            workbook.save(path)
            records, sources = read_records([path])
            result = process(records, CONFIG)
            item = result["records"][0]
            self.assertIn("ПОВРЕЖДЁН_КОД_EXCEL", item["flags"])
            self.assertEqual(item["selection"], "ПРОВЕРИТЬ_ОТБОР")
            self.assertEqual(result["groups"], [])
            result.update(sources=sources, rules=CONFIG,
                          dadata={"mode": "OFFLINE", "actual_requests": 0, "requests": [], "candidates": []})
            output = write_result(result, Path(tmp) / "out")
            exported = load_workbook(output, read_only=True)
            try:
                rows = list(exported["Записи"].values)
                flags_column = rows[0].index("Признаки риска")
                self.assertIn("ПОВРЕЖДЁН_КОД_EXCEL", rows[1][flags_column])
            finally:
                exported.close()

    # Код поступления не участвует в идентификации школы: повреждение только предупреждает.
    def test_damaged_specialty_code_does_not_exclude_school(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.xlsx"
            workbook = Workbook()
            workbook.active.append(["Образовательное учреждение", "Вид образования", "Код специальности", "ФИО"])
            workbook.active.append(["Школа № 1", "Среднее общее", datetime.datetime(2001, 3, 9)])
            workbook.save(path)
            records, _ = read_records([path])
            result = process(records, CONFIG)
            self.assertIn("ПОВРЕЖДЁН_КОД_EXCEL", result["records"][0]["flags"])
            self.assertEqual(result["records"][0]["selection"], "ОТОБРАНО")
            self.assertEqual(len(result["groups"]), 1)

    # ИНН с нулём в формате Excel должен участвовать в проверке страны без потери цифр.
    def test_numeric_inn_retains_displayed_leading_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Образовательное учреждение", "Вид образования", "ИНН", "ФИО"])
            sheet.append(["Школа № 1, г. Минск", "Среднее общее", 123456788])
            sheet["C2"].number_format = "0000000000"
            workbook.save(path)
            records, _ = read_records([path])
            item = process(records, CONFIG)["records"][0]
            self.assertEqual(item["inn_raw"], "0123456788")
            self.assertEqual(item["country"], "КОНФЛИКТ_СТРАНЫ")
            self.assertEqual(item["selection"], "ПРОВЕРИТЬ_ОТБОР")

    # Явно указанная страна в географическом поле не становится российской из-за слова «школа».
    def test_explicit_foreign_countries_outside_cis_are_excluded(self):
        for fields in [{"school_address_raw": "Германия, г. Берлин"},
                       {"location_raw": "Китай; г. Пекин"},
                       {"region_raw": "Германия"}]:
            with self.subTest(fields=fields):
                item = process([record(**fields)], CONFIG)["records"][0]
                self.assertEqual(item["country"], "ЗА_РУБЕЖОМ")
                self.assertEqual(item["selection"], "ИСКЛЮЧЕНО")

    # Страна и российский ИНН противоречат друг другу: исходную строку нельзя исключать молча.
    def test_explicit_foreign_country_with_russian_inn_is_held(self):
        item = process([record(school_address_raw="Китай, г. Пекин", inn_raw="7707083893")], CONFIG)["records"][0]
        self.assertEqual(item["country"], "КОНФЛИКТ_СТРАНЫ")
        self.assertEqual(item["selection"], "ПРОВЕРИТЬ_ОТБОР")

    # Название школы и топоним «Китай-город» не доказывают иностранную страну.
    def test_country_words_in_school_name_and_street_do_not_exclude(self):
        for fields in [{"name": "Школа «Германия»", "city_raw": "Москва"},
                       {"school_address_raw": "Москва, Китай-город"}]:
            with self.subTest(fields=fields):
                item = process([record(**fields)], CONFIG)["records"][0]
                self.assertEqual(item["selection"], "ОТОБРАНО")


if __name__ == "__main__":
    unittest.main()
