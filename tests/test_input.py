from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from ou_processor.input import read_records


def make_book(path, rows):
    book = Workbook()
    sheet = book.active
    sheet.title = "Данные"
    for row in rows:
        sheet.append(row)
    book.save(path)


def read_error(header, rows=()):
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "in.xlsx"
        make_book(path, [header, *rows])
        with tempfile.TemporaryDirectory():
            pass
        try:
            read_records([path])
        except ValueError as error:
            return str(error)
        raise AssertionError("ValueError не было")


class RequiredFieldMessageTests(unittest.TestCase):
    def test_message_lists_found_columns(self):
        message = read_error(
            ["ФИО", "Образовательное учреждение", "Год завершения"],
            [["Иванов", "Школа 1", 2020]],
        )
        self.assertIn("найдены", message)
        self.assertIn("«ФИО»", message)
        self.assertIn("«Образовательное учреждение»", message)

    def test_message_lists_missing_columns(self):
        message = read_error(["ФИО", "Образовательное учреждение"], [["Иванов", "Школа 1"]])
        self.assertIn("не найдены", message)
        self.assertIn("«Вид образования»", message)
        self.assertIn("«Тип документа об образовании»", message)

    def test_found_columns_not_repeated_as_missing(self):
        message = read_error(
            ["ФИО", "Образовательное учреждение", "Год завершения"],
            [["Иванов", "Школа 1", 2020]],
        )
        missing_part = message.split("не найдены")[-1]
        self.assertNotIn("«ФИО»", missing_part)
        self.assertNotIn("«Образовательное учреждение»", missing_part)

    def test_document_replaces_education_no_error(self):
        header = ["ФИО", "Образовательное учреждение", "Тип документа об образовании"]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            make_book(path, [header, ["Иванов", "Школа 1", "Аттестат о среднем общем образовании"]])
            records, _ = read_records([path])
        self.assertEqual(len(records), 1)

    def test_duplicate_header_uses_russian_label(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            make_book(path, [
                ["ФИО", "Образовательное учреждение", "Вид образования", "ФИО"],
                ["Иванов", "Школа 1", "Среднее общее", "Петров"],
            ])
            with self.assertRaises(ValueError) as ctx:
                read_records([path])
        self.assertIn("«ФИО»", str(ctx.exception))

class DuplicateTests(unittest.TestCase):
    def read(self, rows):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            make_book(path, [
                ["ФИО", "Образовательное учреждение", "Вид образования", "Личное дело.Номер"],
                *rows
            ])
            return read_records([path])

    def test_exact_duplicate_removed(self):
        rows = [
            ["Иванов", "Школа 1", "Среднее общее", "001"],
            ["Иванов", "Школа 1", "Среднее общее", "001"],
            ["Петров", "Школа 1", "Среднее общее", "002"]
        ]
        records, sources = self.read(rows)
        self.assertEqual(len(records), 2)
        self.assertEqual(sources[0]["duplicates_removed"], 1)

    def test_conflicting_number_kept_and_flagged(self):
        rows = [
            ["Иванов", "Школа 1", "Среднее общее", "001"],
            ["Сидоров", "Школа 1", "Среднее общее", "001"]
        ]
        records, sources = self.read(rows)
        self.assertEqual(len(records), 2)
        self.assertTrue(records[0]["input_conflict"])
        self.assertTrue(records[1]["input_conflict"])

    def test_empty_numbers_not_deduplicated(self):
        rows = [
            ["Иванов", "Школа 1", "Среднее общее", ""],
            ["Петров", "Школа 1", "Среднее общее", ""]
        ]
        records, _ = self.read(rows)
        self.assertEqual(len(records), 2)

    def test_same_number_and_name_but_different_school_is_conflict(self):
        rows = [
            ["Иванов", "Школа 1", "Среднее общее", "001"],
            ["Иванов", "Школа 2", "Среднее общее", "001"]
        ]
        records, sources = self.read(rows)
        
        self.assertEqual(len(records), 2)
        self.assertTrue(records[0]["input_conflict"])
        self.assertTrue(records[1]["input_conflict"])

class SheetSelectionTests(unittest.TestCase):
    def read_multi(self, sheets, **kwargs):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            book = Workbook()
            book.remove(book.active)
            for title, rows in sheets:
                sheet = book.create_sheet(title)
                for row in rows:
                    sheet.append(row)
            book.save(path)
            return read_records([path], **kwargs)

    def test_service_lists_ignored_by_default(self):
        header = ["ФИО", "Образовательное учреждение", "Вид образования"]
        sheets = [
            ("Данные", [header, ["Иванов", "Школа 1", "Среднее общее"]]),
            ("ВУЗы", [header, ["Петров", "Вуз 1", "Высшее"]]),
            ("Не нашли", [header, ["Сидоров", "Школа 2", "Среднее общее"]]),
            ("Коллективу (1 чел)", [header, ["Смирнов", "Школа 3", "Среднее общее"]])
        ]
        records, sources = self.read_multi(sheets)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["student_name"], "Иванов")
        self.assertIn("ВУЗы", sources[0]["skipped_sheets"])
        self.assertIn("Не нашли", sources[0]["skipped_sheets"])
        self.assertIn("Коллективу (1 чел)", sources[0]["skipped_sheets"])

    def test_explicit_sheet_selection(self):
        header = ["ФИО", "Образовательное учреждение", "Вид образования"]
        sheets = [
            ("Лист А", [header, ["Иванов", "Школа 1", "Среднее общее"]]),
            ("Лист Б", [header, ["Петров", "Школа 2", "Среднее общее"]])
        ]
        records, sources = self.read_multi(sheets, sheet_name="Лист Б")
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["student_name"], "Петров")
        self.assertIn("Лист А", sources[0]["skipped_sheets"])

    def test_copy_of_working_sheet_ignored(self):
        header = ["ФИО", "Образовательное учреждение", "Вид образования"]
        sheets = [
            ("Лист рабочий", [header, ["Иванов", "Школа 1", "Среднее общее"]]),
            ("Лист рабочий (копия)", [header, ["Петров", "Школа 1", "Среднее общее"]])
        ]
        records, sources = self.read_multi(sheets)
        self.assertEqual(len(records), 1)
        self.assertIn("Лист рабочий (копия)", sources[0]["skipped_sheets"])


class DamagedCodeTests(unittest.TestCase):
    def read_cell(self, value, number_format="General"):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "Данные"
            sheet.append(["ФИО", "Образовательное учреждение", "Вид образования", "Год завершения"])
            sheet.append(["Иванов", "Школа 1", "Среднее общее", value])
            sheet["D2"].number_format = number_format
            book.save(path)
            return read_records([path])

    def test_integer_float_is_clean_year(self):
        records, _ = self.read_cell(2020.0)
        self.assertEqual(records[0]["year"], "2020")
        self.assertNotIn("ПОВРЕЖДЁН_КОД_EXCEL", records[0].get("input_flags", []))

    def test_date_is_flagged(self):
        import datetime
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "Данные"
            sheet.append(["ФИО", "Образовательное учреждение", "Вид образования", "Год завершения"])
            sheet.append(["Иванов", "Школа 1", "Среднее общее", datetime.datetime(2026, 3, 12)])
            sheet["D2"].number_format = "DD.MM.YYYY" 
            book.save(path)
            records, _ = read_records([path])
        self.assertIn("ПОВРЕЖДЁН_КОД_EXCEL", records[0]["input_flags"])

    def test_scientific_notation_is_flagged(self):
        records, _ = self.read_cell(1.5e10)
        self.assertIn("ПОВРЕЖДЁН_КОД_EXCEL", records[0]["input_flags"])


class FormatSupportTests(unittest.TestCase):
    def test_region_aliases(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "Данные"
            # Тестируем длинное название из ТЗ
            sheet.append(["ФИО", "Образовательное учреждение", "Вид образования", "Регион образовательного учреждения"])
            sheet.append(["Иванов", "Школа 1", "Среднее общее", "Московская область"])
            book.save(path)

            records, _ = read_records([path])
            self.assertEqual(records[0]["region_raw"], "Московская область")

    def test_reads_department_export_columns(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "Лист рабочий"
            # Заголовки из реальной выгрузки отдела
            header = [
                "Форма обучения", "Код специальности", "Направление\\специальность", "Профиль",
                "Профиль.Сокращенное наименование", "Конкурсная группа", "Личное дело.Номер",
                "Уровень подготовки", "ФИО", "Тип документа об образовании",
                "Образовательное учреждение", "Полное название ОУ в именительном падеже",
                "Название ОУ в родительном падеже", "Адрес", "Индекс"
            ]
            sheet.append(header)
            sheet.append([
                "Очная", "09.03.01", "Информатика", "ИБС", "4.0",
                "09.03.01.04 (Б-О, Москва)", "T-1001", "Бакалавриат",
                "Иванов Иван", "Аттестат о среднем общем образовании",
                "ГБОУ Школа № 1284", "", "", "", ""
            ])
            book.save(path)

            records, _ = read_records([path])

            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["student_name"], "Иванов Иван")
            self.assertEqual(records[0]["personal_file_number"], "T-1001")
            self.assertEqual(records[0]["name"], "ГБОУ Школа № 1284")
            self.assertEqual(records[0]["document"], "Аттестат о среднем общем образовании")

class ExclusionLogicTests(unittest.TestCase):
    def process_row(self, name, inn="", address="", education="Среднее общее", **extra):
        import json
        from pathlib import Path
        from ou_processor.pipeline import process
        config = json.loads((Path(__file__).resolve().parents[1] / "config.json").read_text(encoding="utf-8"))
        
        record = {
            "record_id": "1", 
            "name": name, 
            "inn_raw": inn, 
            "school_address_raw": address, 
            "education": education, 
            "formula_fields": []
        }
        record.update(extra)
        return process([record], config)["records"][0]

    def test_foreign_school_is_excluded(self):
        record = self.process_row(name="СОШ № 5, г. Алматы")
        self.assertEqual(record["country"], "ЗА_РУБЕЖОМ")
        self.assertEqual(record["selection"], "ИСКЛЮЧЕНО")
        self.assertIn("ЗАРУБЕЖНОЕ_ОУ", record["flags"])

    def test_foreign_school_with_russian_inn_goes_to_review(self):
        record = self.process_row(
            name="Школа № 1, г. Минск", 
            inn="7707083893" # Валидный ИНН Сбербанка
        )
        self.assertEqual(record["country"], "КОНФЛИКТ_СТРАНЫ")
        self.assertEqual(record["selection"], "ПРОВЕРИТЬ_ОТБОР")
        self.assertIn("КОНФЛИКТ_СТРАНЫ", record["flags"])

    def test_higher_education_is_excluded(self):
        record = self.process_row(name="МГУ имени Ломоносова", education="Высшее")
        self.assertEqual(record["selection"], "ИСКЛЮЧЕНО")
        self.assertNotIn("ЗАРУБЕЖНОЕ_ОУ", record["flags"])

    def test_unknown_country_is_not_excluded(self):
        record = self.process_row(name="ООО Ромашка")
        self.assertEqual(record["country"], "НЕИЗВЕСТНО")
        self.assertNotEqual(record["selection"], "ИСКЛЮЧЕНО")



    def test_personal_file_conflict_goes_to_review(self):
        record = self.process_row(name="ГБОУ Школа № 1", input_conflict=True)
        self.assertEqual(record["selection"], "ПРОВЕРИТЬ_ОТБОР")
        self.assertIn("КОНФЛИКТ_ЛИЧНОГО_ДЕЛА", record["flags"])

    def test_country_counters_add_up(self):
        import json
        from ou_processor.pipeline import process
        config = json.loads((Path(__file__).resolve().parents[1] / "config.json").read_text(encoding="utf-8"))
        names = ["ГБОУ Школа № 1", "ООО Ромашка", "СОШ № 5, г. Алматы"]
        records = [{"record_id": str(i), "name": n, "inn_raw": "", "school_address_raw": "",
                    "education": "Среднее общее", "formula_fields": []} for i, n in enumerate(names)]
        s = process(records, config)["summary"]
        self.assertEqual(s["russian_probable_rows"], 1)
        self.assertEqual(s["unknown_country_rows"], 1)
        self.assertEqual(s["foreign_rows"], 1)
        total = (s["russian_rows"] + s["russian_probable_rows"] + s["unknown_country_rows"]
                 + s["foreign_rows"] + s["conflict_rows"])
        self.assertEqual(total, s["input_rows"])


class DamagedScopeTests(unittest.TestCase):
    def read_row(self, header, row):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "in.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "Данные"
            sheet.append(header)
            sheet.append(row)
            book.save(path)
            records, _ = read_records([path])
        return records[0]

    def test_specialty_code_date_flagged_once(self):
        import datetime
        record = self.read_row(
            ["ФИО", "Образовательное учреждение", "Вид образования", "Код специальности"],
            ["Иванов", "Школа 1", "Среднее общее", datetime.datetime(2001, 3, 9)])
        self.assertEqual(record["input_flags"].count("ПОВРЕЖДЁН_КОД_EXCEL"), 1)

    def test_numeric_inn_index_and_file_number_not_flagged(self):
        record = self.read_row(
            ["ФИО", "Образовательное учреждение", "Вид образования", "ИНН", "Индекс", "Личное дело.Номер"],
            ["Иванов", "Школа 1", "Среднее общее", 7707083893, 119019, 123456])
        self.assertEqual(record["input_flags"], [])


if __name__ == "__main__":
    unittest.main()
