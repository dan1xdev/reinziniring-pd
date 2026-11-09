from pathlib import Path
import re
import tempfile
import unittest
from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.formula import ArrayFormula

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

            with ZipFile(source) as archive:
                original_parts = {name: archive.read(name) for name in archive.namelist()}
            result, record, candidate = self.make_result_from_minimal(source)
            result["letters_ready"][0]["official_name"] = '=Название & "филиал" <1>\nВторая строка'
            stats = fill_source_copies(result, output)
            copy_path = output / "filled_sources" / stats["copies"][0]["file"]
            copied = load_workbook(copy_path)
            headers = [cell.value for cell in copied["Данные"][1]]
            self.assertIn("Полное название ОУ в именительном падеже", headers)
            self.assertIn("Название ОУ в родительном падеже", headers)
            self.assertIn("Адрес", headers)
            self.assertIn("Индекс", headers)
            self.assertEqual(copied["Архив"]["A1"].value, "Сохранить")
            self.assertEqual(copied["Данные"]["E2"].value, '=Название & "филиал" <1>\nВторая строка')
            self.assertEqual(copied["Данные"]["E2"].data_type, "s")
            self.assertEqual(copied["Данные"]["D2"].value, "00123")
            self.assertEqual(copied["Данные"]["D2"].number_format, "@")
            self.assertEqual(copied["Данные"].calculate_dimension(), "A1:H2")
            copied.close()
            with ZipFile(copy_path) as archive:
                for name, content in original_parts.items():
                    if name != "xl/worksheets/sheet1.xml":
                        self.assertEqual(archive.read(name), content, name)
                self.assertIn(b'ref="A1:H2"', archive.read("xl/worksheets/sheet1.xml"))

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

    def test_unconfirmed_letter_with_same_inn_cannot_replace_confirmed_values(self):
        for confirmed_status in ("АВТО", "ПРОВЕРЕНО", None):
            for proposal_status in ("ПРОВЕРИТЬ", "ВРУЧНАЯ"):
                with self.subTest(confirmed=confirmed_status, proposal=proposal_status):
                    with tempfile.TemporaryDirectory() as tmp:
                        source = Path(tmp) / "source.xlsx"
                        self.make_input(source)
                        result, _, candidate = self.make_result(source)
                        confirmed = {**(result["letters_ready"] + result["letters_review"])[0]}
                        if confirmed_status is None:
                            confirmed.pop("match_status", None)
                        else:
                            confirmed["match_status"] = confirmed_status
                        result["letters_ready"] = [confirmed]
                        result["letters_review"] = [{
                            **confirmed,
                            "match_status": proposal_status,
                            "official_name": "Неподтверждённое учреждение",
                            "address": "Неподтверждённый адрес",
                        }]

                        stats = fill_source_copies(result, Path(tmp) / "out")

                        copied = load_workbook(
                            Path(tmp) / "out" / "filled_sources" / stats["copies"][0]["file"]
                        )
                        self.assertEqual(copied["Данные"]["E2"].value, candidate["candidate_name"])
                        self.assertEqual(copied["Данные"]["G2"].value, "Москва, улица Арбат, дом 1")
                        copied.close()

    def test_conflicting_postal_codes_are_not_written(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            self.make_input(source)
            original = load_workbook(source)
            original["Данные"].append([
                "Петров Пётр", "ГБОУ Школа № 123", "Среднее общее", "00234", "", "", "", ""
            ])
            original.save(source)
            original.close()
            result, _, _ = self.make_result(source)
            other = dict(result["resolution"]["groups"][0])
            other["record_ids"] = [result["records"][1]["record_id"]]
            other["selected"] = {**other["selected"], "postal_code": "119020"}
            result["resolution"]["groups"].append(other)
            prepare_letters(result)
            self.assertIn("Индексы в источниках не совпадают", result["letters_review"][0]["checks"])

            stats = fill_source_copies(result, Path(tmp) / "out")
            copied = load_workbook(Path(tmp) / "out" / "filled_sources" / stats["copies"][0]["file"])
            self.assertIsNone(copied["Данные"]["H2"].value)
            self.assertIsNone(copied["Данные"]["H3"].value)
            copied.close()

    def test_unverified_restored_name_is_not_written(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            self.make_input(source)
            result, _, _ = self.make_result(source)
            result["resolution"]["groups"][0]["selected"]["candidate_name"] = (
                "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ № 123"
            )
            prepare_letters(result)
            self.assertIn("Регистр восстановлен автоматически", result["letters_review"][0]["checks"])

            stats = fill_source_copies(result, Path(tmp) / "out")
            copied = load_workbook(Path(tmp) / "out" / "filled_sources" / stats["copies"][0]["file"])
            self.assertIsNone(copied["Данные"]["E2"].value)
            copied.close()

    def test_conflicting_existing_address_is_preserved(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            self.make_input(source)
            original = load_workbook(source)
            original["Данные"]["G2"] = "Москва, Тверская улица, дом 1"
            original.save(source)
            original.close()
            result, _, _ = self.make_result(source)

            stats = fill_source_copies(result, Path(tmp) / "out")
            copied = load_workbook(Path(tmp) / "out" / "filled_sources" / stats["copies"][0]["file"])
            self.assertEqual(copied["Данные"]["G2"].value, "Москва, Тверская улица, дом 1")
            copied.close()

    def test_array_formula_result_cells_are_preserved_with_the_anchor(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            self.make_input(source)
            result, _, candidate = self.make_result(source)
            original = load_workbook(source)
            original["Данные"]["G2"] = ArrayFormula(ref="G2:H2", text="=TRANSPOSE({1;101000})")
            original["Данные"]["H2"] = 101000
            original.save(source)
            original.close()
            with ZipFile(source) as archive:
                before = archive.read("xl/worksheets/sheet1.xml")
            original_cells = [
                re.search(rb'<c\b[^>]*r="' + coordinate + rb'"[^>]*>.*?</c>', before).group()
                for coordinate in (b"G2", b"H2")
            ]

            stats = fill_source_copies(result, Path(tmp) / "out")

            destination = Path(tmp) / "out" / "filled_sources" / stats["copies"][0]["file"]
            copied = load_workbook(destination, data_only=True)
            self.assertEqual(copied["Данные"]["H2"].value, 101000)
            self.assertEqual(copied["Данные"]["E2"].value, candidate["candidate_name"])
            copied.close()
            self.assertEqual(stats["written"].get("postal_code", 0), 0)
            with ZipFile(destination) as archive:
                after = archive.read("xl/worksheets/sheet1.xml")
            for original_cell in original_cells:
                self.assertIn(original_cell, after)

    def test_fills_formula_book_preserving_all_other_parts_and_cached_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            self.make_input(source)
            original = load_workbook(source)
            original["Другой лист"]["C3"] = "=1+1"
            original["Данные"]["I2"] = "=1+1"
            original["Данные"]["H2"] = "=100000+19019"
            original.save(source)
            original.close()
            with ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            for name in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
                parts[name] = re.sub(br"<f>1\+1</f><v(?:></v>|\s*/>)", b"<f>1+1</f><v>2</v>", parts[name])
                parts[name] = re.sub(
                    br"<f>100000\+19019</f><v(?:></v>|\s*/>)",
                    b"<f>100000+19019</f><v>119019</v>", parts[name]
                )
            with ZipFile(source, "w", ZIP_DEFLATED) as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
            original = load_workbook(source, data_only=True)
            self.assertEqual(original["Другой лист"]["C3"].value, 2)
            original.close()

            before = source.read_bytes()
            stats = fill_source_copies(self.make_result(source)[0], Path(tmp) / "out")
            self.assertEqual(stats["copy_count"], 1)
            self.assertEqual(stats["errors"], [])
            self.assertEqual(source.read_bytes(), before)
            destination = Path(tmp) / "out" / "filled_sources" / stats["copies"][0]["file"]
            with ZipFile(destination) as archive:
                self.assertEqual(set(archive.namelist()), set(parts))
                for name, content in parts.items():
                    if name != "xl/worksheets/sheet1.xml":
                        self.assertEqual(archive.read(name), content, name)
                selected_xml = archive.read("xl/worksheets/sheet1.xml")
                self.assertIn(b"<f>1+1</f><v>2</v>", selected_xml)
                self.assertIn(b"<f>100000+19019</f><v>119019</v>", selected_xml)
            copied = load_workbook(destination, data_only=True)
            self.assertEqual(copied["Данные"]["I2"].value, 2)
            self.assertEqual(copied["Данные"]["H2"].value, 119019)
            self.assertTrue(copied["Данные"]["E2"].value)
            copied.close()
            original = load_workbook(source, data_only=True)
            self.assertEqual(original["Другой лист"]["C3"].value, 2)
            original.close()

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
            result["dadata_checks"] = [
                {"status": "КОНФЛИКТ_ИНН", "group_id": "G-1"},
                {"status": "СОВПАДАЕТ", "group_id": "G-2"},
            ]
            result["dadata_candidates"] = []
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
            self.assertEqual(summary["DaData — кандидатов получено"], 0)
            self.assertEqual(summary["DaData — сверок КОНФЛИКТ_ИНН"], 1)
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
