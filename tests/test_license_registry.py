"""Проверки реестра лицензий: обновление, снимок при сбое, поиск, настройки."""

from io import BytesIO
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from ou_processor.license_registry import build_index, current_snapshot, lookup, update, auto_update
from ou_processor.resolution import LicenseSearch


XML = '''<?xml version="1.0" encoding="UTF-8"?>
<OpenData><Licenses><License>
<LicenseRegNumber>Л035-00001-77/00000001</LicenseRegNumber>
<FullName>Тестовая школа № 1</FullName><ShortName>ТШ № 1</ShortName>
<INN>0000000001</INN><OGRN>0000000000001</OGRN><KPP>000001001</KPP>
<DateStart>2026-01-15</DateStart><StatusName>Действующая</StatusName>
<RegionName>Москва</RegionName><PostAddress>г Москва, Тестовая улица, д 1</PostAddress>
</License>
<License>
<LicenseRegNumber>Л035-00002-77/00000002</LicenseRegNumber>
<FullName>Тестовый колледж № 2</FullName><INN>0000000002</INN>
<DateStart>2026-02-20</DateStart><StatusName>Прекращена</StatusName>
<RegionName>Санкт-Петербург</RegionName>
<PostAddress>г Санкт-Петербург, Новая улица, д 2</PostAddress>
</License>
</Licenses></OpenData>'''

CSV = '''Наименование лицензиата;ИНН;ОГРН;Номер лицензии;Статус лицензии;Регион;Адрес
Тестовая школа № 1;0000000001;0000000000001;Л035-00001-77/00000001;Действующая;Москва;г Москва, Тестовая улица, д 1
'''

PAGE = '''<html><body><a href="/opendata/7701537808-FBDRL/data-{date}-structure-20150421.zip">выгрузка</a></body></html>'''


def archive_bytes(xml=XML, date="20261107"):
    stream = BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"data-{date}-structure-20150421.xml", xml.encode("utf-8"))
    return stream.getvalue()


class Response(BytesIO):
    def __init__(self, data, headers):
        super().__init__(data)
        self.headers = headers


class Server:
    """Фейковый источник: страница открытых данных и файл выгрузки."""

    def __init__(self, data=None, date="20261107"):
        self.date = date
        self.data = archive_bytes(date=date) if data is None else data
        self.calls = []
        self.headers = {"Content-Length": str(len(self.data)), "ETag": '"tag-' + date + '"'}

    def __call__(self, request, timeout):
        self.calls.append(request.get_full_url())
        if "opendata_licenses" in request.get_full_url():
            return Response(PAGE.format(date=self.date).encode("utf-8"),
                            {"Content-Type": "text/html; charset=utf-8"})
        return Response(b"" if request.get_method() == "HEAD" else self.data, self.headers)


class LicenseRegistryTests(unittest.TestCase):
    def test_update_then_unchanged_and_lookup(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp)
            server = Server()
            report = update(cache, settings={"update_days": 0}, opener=server)
            self.assertEqual(report["status"], "updated")
            self.assertEqual(report["manifest"]["records"], 2)
            self.assertEqual(report["manifest"]["source_kind"], "license_registry")
            self.assertEqual(report["manifest"]["snapshot_date"], "20261107")
            report = update(cache, settings={"update_days": 0}, opener=server)
            self.assertEqual(report["status"], "unchanged")
            found = lookup(cache, inn="0000000001")
            self.assertEqual(found["source_kind"], "license_registry")
            self.assertEqual(len(found["matches"]), 1)
            self.assertEqual(found["matches"][0]["full_name"], "Тестовая школа № 1")
            self.assertEqual(found["matches"][0]["license_number"], "Л035-00001-77/00000001")
            by_name = lookup(cache, name="колледж")
            self.assertEqual(by_name["matches"][0]["inn"], "0000000002")

    def test_csv_export_is_supported(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "vygruzka.csv"
            path.write_text(CSV, encoding="utf-8-sig")
            base = build_index(path, Path(tmp) / "index.sqlite")
            self.assertEqual(base["records"], 1)

    def test_corrupt_download_preserves_current_snapshot(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp)
            update(cache, settings={"update_days": 0}, opener=Server())
            original = (cache / "current.json").read_bytes()
            for broken in [b"<html>server error</html>", archive_bytes(XML[:-30], "20261108")]:
                with self.assertRaises(Exception):
                    update(cache, settings={"update_days": 0}, opener=Server(broken, "20261108"))
                self.assertEqual((cache / "current.json").read_bytes(), original)
                self.assertEqual(json.loads((cache / "last_check.json").read_text())["status"], "error")
                self.assertTrue(lookup(cache, inn="0000000001")["matches"])

    def test_two_versions_retained(self):
        with tempfile.TemporaryDirectory() as tmp:
            for date in ["20261106", "20261107", "20261108"]:
                update(tmp, settings={"update_days": 0}, opener=Server(date=date))
            snapshots = list(Path(tmp).glob("snapshot-*"))
            self.assertEqual(len(snapshots), 2)
            self.assertEqual(current_snapshot(Path(tmp))[1]["snapshot_date"], "20261108")

    def test_update_days_skips_network_until_forced(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp)
            server = Server()
            update(cache, settings={"update_days": 0}, opener=server)
            calls_before = len(server.calls)
            report = update(cache, settings={"update_days": 7}, opener=server)
            self.assertEqual(report["status"], "fresh")
            self.assertEqual(len(server.calls), calls_before)
            forced = update(cache, settings={"update_days": 7}, force=True, opener=server)
            # Одинаковый SHA-256 не создаёт новую копию снимка, только проверяется.
            self.assertEqual(forced["status"], "unchanged")
            self.assertTrue(forced["verified_download"])

    def test_auto_update_reads_config_settings(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            cache = root / "cache"
            server = Server()
            update(cache, settings={"update_days": 0}, opener=server)
            config = root / "config.json"
            config.write_text(json.dumps({"license_registry": {"update_days": 3}}), encoding="utf-8")
            report = auto_update(cache, config_path=config)
            self.assertEqual(report["status"], "fresh")

    def test_local_source_file_works_without_network(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            export = root / "vygruzka.zip"
            export.write_bytes(archive_bytes())
            settings = {"update_days": 0, "source_file": str(export), "source_page": ""}
            report = update(root / "cache", settings=settings, opener=Server())
            self.assertEqual(report["status"], "updated")
            self.assertTrue(report["source_url"].startswith("file:"))
            repeat = update(root / "cache", settings=settings, opener=Server())
            self.assertEqual(repeat["status"], "unchanged")
            self.assertTrue(repeat["verified_download"])

    def test_missing_local_file_is_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            settings = {"update_days": 0, "source_file": str(Path(tmp) / "net.xml"), "source_page": ""}
            with self.assertRaisesRegex(ValueError, "не найден"):
                update(Path(tmp) / "cache", settings=settings, opener=Server())

    def test_multiple_data_files_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.zip"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("one.xml", XML)
                archive.writestr("two.xml", XML)
            with self.assertRaisesRegex(ValueError, "ровно один"):
                build_index(path, Path(tmp) / "index.sqlite")

    def test_license_search_is_primary_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp)
            update(cache, settings={"update_days": 0}, opener=Server())
            search = LicenseSearch(cache)
            try:
                candidates, truncated = search.search(dict(
                    name='Тестовая школа № 1', city='Москва', region='Москва'))
                self.assertFalse(truncated)
                self.assertEqual({c['candidate_inn'] for c in candidates}, {'0000000001'})
                self.assertEqual(candidates[0]['source'], 'Реестр лицензий')
                self.assertEqual(candidates[0]['legal_status'], 'Действующее')
                stopped = search.search(dict(name='Тестовый колледж № 2',
                                             city='Санкт-Петербург', region='Санкт-Петербург'))
                self.assertEqual(stopped[0][0]['legal_status'], 'Прекращена')
            finally:
                search.connection.close()


if __name__ == "__main__":
    unittest.main()
