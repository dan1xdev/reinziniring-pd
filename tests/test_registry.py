"""Проверки обновления и поиска реестра без внешней сети."""

from io import BytesIO
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from ou_processor.registry import build_index, cache_lock, current_snapshot, lookup, update


XML = '''<?xml version="1.0" encoding="UTF-8"?>
<OpenData><Certificates><Certificate><Id>certificate-1</Id>
<StatusName>Действующее</StatusName><IssueDate>2026-01-01</IssueDate>
<RegionName>Москва</RegionName><EduOrgFullName>Тестовая школа № 1</EduOrgFullName>
<EduOrgINN>0000000001</EduOrgINN><EduOrgOGRN>0000000000001</EduOrgOGRN>
<PostAddress>Тестовый адрес</PostAddress>
<ActualEducationOrganization><FullName>Тестовая школа № 1</FullName>
<INN>0000000001</INN><PostAddress>Новый адрес</PostAddress></ActualEducationOrganization>
<Supplements><Supplement><Id>supplement-1</Id><StatusName>Действующее</StatusName>
<IsForBranch>1</IsForBranch><EduOrgFullName>Филиал школы</EduOrgFullName>
<EduOrgAddress>Адрес филиала</EduOrgAddress><EducationalPrograms><EducationalProgram>
<EduLevelName>Среднее общее образование</EduLevelName>
</EducationalProgram></EducationalPrograms></Supplement></Supplements>
</Certificate></Certificates></OpenData>'''


def archive_bytes(xml=XML, date="20261018"):
    stream = BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"data-{date}-structure-20160713.xml", xml.encode("utf-8"))
    return stream.getvalue()


class Response(BytesIO):
    def __init__(self, data, headers):
        super().__init__(data)
        self.headers = headers


class Server:
    def __init__(self, data=None, date="20261018"):
        self.data = archive_bytes(date=date) if data is None else data
        self.calls = []
        self.headers = {"Content-Length": str(len(self.data)), "Last-Modified": date,
                        "Content-Disposition": f"attachment; filename=data-{date}-structure-20160713.zip"}

    def __call__(self, request, timeout):
        self.calls.append(request.method)
        return Response(b"" if request.method == "HEAD" else self.data, self.headers)


class RegistryTests(unittest.TestCase):
    def test_update_then_head_only_and_lookup(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp)
            server = Server()
            report = update(cache, opener=server)
            self.assertEqual(report["status"], "updated")
            self.assertEqual(report["manifest"]["certificates"], 1)
            self.assertEqual(report["manifest"]["organization_records"], 3)
            report = update(cache, opener=server)
            self.assertEqual(report["status"], "unchanged")
            self.assertEqual(server.calls, ["HEAD", "GET", "HEAD"])
            found = lookup(cache, inn="0000000001")
            self.assertEqual(len(found["matches"]), 2)
            self.assertEqual(found["source_kind"], "accreditation_registry")

    def test_branch_does_not_inherit_inn_or_legal_address(self):
        with tempfile.TemporaryDirectory() as tmp:
            update(tmp, opener=Server())
            found = lookup(tmp, name="Филиал")
            row = found["matches"][0]
            self.assertEqual(row["inn"], "")
            self.assertEqual(row["parent_inn"], "0000000001")
            self.assertEqual(row["post_address"], "Адрес филиала")
            self.assertEqual(row["education_levels"], ["Среднее общее образование"])

    def test_corrupt_download_preserves_current_snapshot(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp)
            update(cache, opener=Server())
            original = (cache / "current.json").read_bytes()
            for broken in [b"<html>error</html>", archive_bytes(XML[:-30], "20261019")]:
                with self.assertRaises(Exception):
                    update(cache, opener=Server(broken, "20261019"))
                self.assertEqual((cache / "current.json").read_bytes(), original)
                self.assertEqual(json.loads((cache / "last_check.json").read_text())["status"], "error")
                self.assertTrue(lookup(cache, inn="0000000001")["matches"])

    def test_content_length_mismatch_not_published(self):
        with tempfile.TemporaryDirectory() as tmp:
            server = Server()
            server.headers["Content-Length"] = str(len(server.data) + 1)
            with self.assertRaisesRegex(ValueError, "не полностью"):
                update(tmp, opener=server)
            self.assertFalse((Path(tmp) / "current.json").exists())

    def test_dtd_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            xml = '<?xml version="1.0"?><!DOCTYPE OpenData [<!ENTITY x "bad">]><OpenData>&x;</OpenData>'
            with self.assertRaises(Exception):
                update(tmp, opener=Server(archive_bytes(xml)))
            self.assertFalse((Path(tmp) / "current.json").exists())

    def test_force_same_bytes_is_validated_without_replacing_snapshot(self):
        with tempfile.TemporaryDirectory() as tmp:
            server = Server()
            before = update(tmp, opener=server)
            after = update(tmp, force=True, opener=server)
            self.assertEqual(after["status"], "unchanged")
            self.assertTrue(after["verified_download"])
            self.assertEqual(before["snapshot"], after["snapshot"])

    def test_two_versions_retained(self):
        with tempfile.TemporaryDirectory() as tmp:
            for date in ["20261018", "20261019", "20261020"]:
                update(tmp, opener=Server(date=date))
            snapshots = list(Path(tmp).glob("snapshot-*"))
            self.assertEqual(len(snapshots), 2)
            self.assertEqual(current_snapshot(Path(tmp))[1]["snapshot_date"], "20261020")

    def test_lock_prevents_concurrent_update(self):
        with tempfile.TemporaryDirectory() as tmp:
            with cache_lock(Path(tmp)):
                with self.assertRaisesRegex(ValueError, "уже запущено"):
                    update(tmp, opener=Server())
            self.assertFalse((Path(tmp) / "update.lock").exists())

    def test_multiple_xml_files_and_empty_dataset_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.zip"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("one.xml", XML)
                archive.writestr("two.xml", XML)
            with self.assertRaisesRegex(ValueError, "ровно один"):
                build_index(path, Path(tmp) / "index.sqlite")
            with self.assertRaisesRegex(ValueError, "не найдены"):
                update(Path(tmp) / "cache", opener=Server(archive_bytes("<OpenData><Certificates/></OpenData>")))


if __name__ == "__main__":
    unittest.main()
