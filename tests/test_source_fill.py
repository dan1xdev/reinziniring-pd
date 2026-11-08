from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

from ou_processor.input import read_records
from ou_processor.letters import prepare_letters
from ou_processor.source_fill import fill_source_copies
from ou_processor.full_run import save


class SourceFillTests(unittest.TestCase):
    def make_input(self, path):
        book = Workbook()
        sheet = book.active
        sheet.title = "Данные"
        sheet.append([
            "ФИО", "Образовательное учреждение", "Вид образования",
            "Личное дело.Номер", "Полное название ОУ в именительном падеже",
            "Название ОУ в родительном падеже", "Адрес", "Индекс",
        ])
        sheet.append([
            "Иванов Иван", "ГБОУ Школа № 123", "Среднее общее",
            "00123", "", "", "", "",
        ])
        other = book.create_sheet("Другой лист")
        other["A1"] = "Не менять"
        other["B2"] = 42
        book.save(path)

    def make_result(self, path):
        records, sources = read_records([path])
        record = records[0]
        candidate = {
            "candidate_inn": "7707083893",
            "candidate_name": "Государственное бюджетное общеобразовательное учреждение № 123",
            "candidate_ogrn": "1027700132195",
            "address": "119019, Москва, улица Арбат, дом 1",
            "postal_code": "119019",
            "source": "DaData",
            "source_url": "https://dadata.ru/",
            "legal_status": "ACTIVE",
            "eligible": True,
            "score": 100.0,
            "reasons": [],
            "branch_type": "",
        }
        group = {
            "group_id": "G-1",
            "record_ids": [record["record_id"]],
            "source_name": record["name"],
            "selected": candidate,
            "status": "СОПОСТАВЛЕНО",
            "registry_used": False,
            "registry_truncated": False,
            "candidate_count": 1,
        }
        result = {
            "records": records,
            "sources": sources,
            "groups": [],
            "resolution": {
                "groups": [group],
                "candidates": [{**candidate, "group_id": "G-1"}],
                "requests": [],
                "actual_requests": 0,
                "registry": {},
            },
            "students": [{
                "school_name": candidate["candidate_name"],
                "inn": candidate["candidate_inn"],
                "address": candidate["address"],
                "student": record["student_name"],
                "file": record["personal_file_number"],
                "record_ids": record["record_id"],
            }],
            "schools": [],
            "review": [],
            "full_summary": {},
            "summary": {"excluded_rows": 0, "selection_review_rows": 0},
            "created_at_utc": "2026-11-08T00:00:00+00:00",
        }
        prepare_letters(result)
        return result, record, candidate

    def test_fills_four_columns_in_separate_copy_and_preserves_other_sheet(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            output = Path(tmp) / "out"
            self.make_input(source)
            result, record, candidate = self.make_result(source)

            stats = fill_source_copies(result, output)

            self.assertEqual(stats["copy_count"], 1)
            self.assertEqual(stats["source_rows"], 1)
            self.assertEqual(stats["written"]["official_name"], 1)
            self.assertEqual(stats["written"]["genitive_name"], 1)
            self.assertEqual(stats["written"]["address"], 1)
            self.assertEqual(stats["written"]["postal_code"], 1)

            copy_path = output / "filled_sources" / stats["copies"][0]["file"]
            self.assertTrue(copy_path.exists())
            self.assertTrue(source.exists())

            original = load_workbook(source, data_only=False)
            copied = load_workbook(copy_path, data_only=False)
            self.assertIsNone(original["Данные"]["E2"].value)
            self.assertEqual(copied["Данные"]["E2"].value, candidate["candidate_name"])
            self.assertEqual(copied["Данные"]["F2"].value, "Государственного бюджетного общеобразовательного учреждения № 123")
            self.assertEqual(copied["Данные"]["G2"].value, "Москва, улица Арбат, дом 1")
            self.assertEqual(copied["Данные"]["H2"].value, "119019")
            self.assertEqual(copied["Другой лист"]["A1"].value, "Не менять")
            self.assertEqual(copied["Другой лист"]["B2"].value, 42)
            original.close()
            copied.close()

    def test_adds_missing_target_columns_without_touching_existing_cells(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            output = Path(tmp) / "out"
            book = Workbook()
            sheet = book.active
            sheet.title = "Данные"
            sheet.append(["ФИО", "Образовательное учреждение", "Вид образования", "Личное дело.Номер"])
            sheet.append(["Иванов Иван", "ГБОУ Школа № 123", "Среднее общее", "00123"])
            sheet["D2"].number_format = "@"
            book.create_sheet("Архив")["A1"] = "Сохранить"
            book.save(source)

            result, record, candidate = self.make_result_from_minimal(source)
            stats = fill_source_copies(result, output)
            copy_path = output / "filled_sources" / stats["copies"][0]["file"]
            copied = load_workbook(copy_path)
            headers = [cell.value for cell in copied["Данные"][1]]
            self.assertIn("Полное название ОУ в именительном падеже", headers)
            self.assertIn("Название ОУ в родительном падеже", headers)
            self.assertIn("Адрес", headers)
            self.assertIn("Индекс", headers)
            self.assertEqual(copied["Архив"]["A1"].value, "Сохранить")
            copied.close()

    def test_exact_duplicate_source_row_is_also_filled(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            output = Path(tmp) / "out"
            self.make_input(source)
            book = load_workbook(source)
            sheet = book["Данные"]
            sheet.append([cell.value for cell in sheet[2]])
            book.save(source)
            book.close()

            result, record, candidate = self.make_result(source)
            self.assertEqual(len(result["records"]), 1)
            self.assertEqual(len(result["records"][0]["duplicate_source_rows"]), 1)

            stats = fill_source_copies(result, output)
            self.assertEqual(stats["source_rows"], 2)
            self.assertEqual(stats["written"]["official_name"], 2)

            copy_path = output / "filled_sources" / stats["copies"][0]["file"]
            copied = load_workbook(copy_path, data_only=True)
            self.assertEqual(copied["Данные"]["E2"].value, candidate["candidate_name"])
            self.assertEqual(copied["Данные"]["E3"].value, candidate["candidate_name"])
            copied.close()

    def test_full_export_has_required_sheets_and_summary(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            output = Path(tmp) / "out"
            self.make_input(source)
            result, record, candidate = self.make_result(source)
            result["schools"] = []
            result["review"] = []
            result["summary"] = {
                "excluded_rows": 0,
                "selection_review_rows": 0,
            }
            result["full_summary"] = {}
            result["resolution"]["registry"] = {}
            result["resolution"]["requests"] = [{
                "group_id": "G-1", "query": "ГБОУ Школа 123",
                "status": "ОШИБКА_СЕТИ", "candidate_count": 0,
                "cached": False, "timestamp_utc": "2026-11-08T00:00:00Z",
            }]
            result["resolution"]["actual_requests"] = 1
            save(result, output)

            report = load_workbook(output / "schools.xlsx", data_only=True)
            self.assertIn("Письма", report.sheetnames)
            self.assertIn("Ручная проверка", report.sheetnames)
            self.assertIn("Сводка", report.sheetnames)
            self.assertNotIn("Письма к печати", report.sheetnames)
            self.assertNotIn("Письма проверить", report.sheetnames)
            self.assertEqual(report["Письма"].max_row, 2)
            summary = {row[0]: row[1] for row in report["Сводка"].iter_rows(min_row=2, values_only=True)}
            self.assertEqual(summary["Создано копий исходных книг"], 1)
            self.assertEqual(summary["DaData — записей запросов"], 1)
            self.assertEqual(summary["DaData — ошибок и пропусков"], 1)
            self.assertEqual(summary["Дубли удалены до обработки"], 0)
            self.assertIn("Заполнено в копиях — Индекс", summary)
            manual_types = [row[0] for row in report["Ручная проверка"].iter_rows(min_row=2, values_only=True)]
            self.assertIn("Проверка DaData", manual_types)
            report.close()

    def make_result_from_minimal(self, source):
        records, sources = read_records([source])
        record = records[0]
        candidate = {
            "candidate_inn": "7707083893",
            "candidate_name": "Государственное бюджетное общеобразовательное учреждение № 123",
            "candidate_ogrn": "",
            "address": "119019, Москва, улица Арбат, дом 1",
            "postal_code": "119019",
            "source": "DaData",
            "source_url": "",
            "legal_status": "ACTIVE",
            "eligible": True,
            "score": 100.0,
            "reasons": [],
            "branch_type": "",
        }
        group = {
            "group_id": "G-1", "record_ids": [record["record_id"]],
            "source_name": record["name"], "selected": candidate,
            "status": "СОПОСТАВЛЕНО", "registry_used": False,
            "registry_truncated": False, "candidate_count": 1,
        }
        result = {
            "records": records, "sources": sources,
            "resolution": {"groups": [group], "candidates": [], "requests": [],
                           "actual_requests": 0, "registry": {}},
            "students": [{"inn": candidate["candidate_inn"], "student": record["student_name"],
                          "file": record["personal_file_number"], "record_ids": record["record_id"]}],
            "letters_ready": [], "letters_review": [],
        }
        prepare_letters(result)
        return result, record, candidate


if __name__ == "__main__":
    unittest.main()
