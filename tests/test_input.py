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


if __name__ == "__main__":
    unittest.main()
