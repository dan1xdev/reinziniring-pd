"""Адреса и ученики связываются только внутри соответствующей группы."""

import json
from pathlib import Path
import unittest

from ou_processor.dadata import run_sample
from ou_processor.pipeline import process
from ou_processor.roster import attach_roster


CONFIG = json.loads((Path(__file__).resolve().parents[1] / "config.json").read_text(encoding="utf-8"))


def row(index, **values):
    return {"record_id": str(index), "name": "ГБОУ города Москвы Школа № 1",
            "student_name": "Тестов Первый Ученик", "personal_file_number": "00001",
            "education": "Среднее общее", "formula_fields": [], **values}


def build(rows, candidates=None):
    result = process(rows, CONFIG)
    result["dadata"] = {"candidates": candidates or []}
    attach_roster(result)
    return result


class RosterTests(unittest.TestCase):
    def test_names_remain_paired_with_numbers(self):
        result = build([row(1), row(2, student_name="Тестов Второй Ученик", personal_file_number="00002"), row(3)])
        roster = result["school_students"]
        self.assertEqual(len(roster), 2)
        self.assertEqual({(r["student_name"], r["personal_file_number"]) for r in roster},
                         {("Тестов Первый Ученик", "00001"), ("Тестов Второй Ученик", "00002")})
        self.assertEqual(sorted(roster[0]["source_record_ids"]), ["1", "3"])

    def test_partial_identity_not_deduplicated(self):
        result = build([row(1, personal_file_number=""), row(2, personal_file_number="")])
        self.assertEqual(len(result["school_students"]), 2)

    def test_name_conflict_for_same_number(self):
        result = build([row(1), row(2, student_name="Другое ФИО")])
        self.assertEqual(len(result["school_students"]), 2)
        self.assertTrue(all("РАЗНЫЕ_ФИО" in r["student_status"] for r in result["school_students"]))

    def test_missing_student_not_invented(self):
        result = build([row(1, student_name="", personal_file_number="", case_id="CASE-001")])
        roster = result["school_students"][0]
        self.assertEqual(roster["student_name"], "")
        self.assertIn("НЕТ_ФИО", roster["student_status"])
        self.assertEqual(roster["school_address"], "")

    def test_source_address_preserved_and_conflict_not_overwritten(self):
        result = build([row(1, school_address_raw="Адрес 1", postal_code_raw="001234")])
        roster = result["school_students"][0]
        self.assertEqual(roster["school_address"], "Адрес 1")
        self.assertEqual(roster["postal_code"], "001234")
        result = build([row(1, school_address_raw="Адрес 1"), row(2, school_address_raw="Адрес 2")])
        self.assertEqual(result["school_students"][0]["school_address"], "")
        self.assertIn("КОНФЛИКТ", result["school_students"][0]["address_status"])

    def test_candidate_address_scoped_and_unconfirmed(self):
        rows = [row(1), row(2, name="ГБОУ города Москвы Школа № 2")]
        result = build(rows)
        group_id = result["records"][0]["group_id"]
        result["dadata"]["candidates"] = [{"group_id": group_id, "address": "Адрес кандидата",
                                           "legal_status": "ACTIVE", "candidate_inn": "0000000000"}]
        attach_roster(result)
        matched = next(r for r in result["school_students"] if r["group_id"] == group_id)
        other = next(r for r in result["school_students"] if r["group_id"] != group_id)
        self.assertEqual(matched["school_address"], "Адрес кандидата")
        self.assertEqual(matched["address_status"], "АДРЕС_КАНДИДАТА_ПРОВЕРИТЬ")
        self.assertEqual(other["school_address"], "")
        self.assertTrue(all(r["confirmed_inn"] == "" for r in result["records"]))

    def test_ambiguous_or_conflicting_candidate_not_selected(self):
        result = build([row(1)])
        candidate = {"group_id": result["groups"][0]["group_id"], "address": "Адрес",
                     "legal_status": "ACTIVE"}
        for candidates in [[candidate, candidate], [{**candidate, "flags": ["КОНФЛИКТ_НОМЕРА"]}]]:
            result["dadata"]["candidates"] = candidates
            attach_roster(result)
            self.assertEqual(result["school_students"][0]["school_address"], "")

    def test_student_data_never_in_api_payload(self):
        result = build([row(1, student_name="ЛОКАЛЬНОЕ ФИО", personal_file_number="PRIVATE-999")])
        report = json.dumps(run_sample(result), ensure_ascii=False)
        self.assertNotIn("ЛОКАЛЬНОЕ ФИО", report)
        self.assertNotIn("PRIVATE-999", report)


if __name__ == "__main__":
    unittest.main()
