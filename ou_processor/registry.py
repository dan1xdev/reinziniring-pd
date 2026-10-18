"""Резервный реестр аккредитации: обновление снимка и локальный поиск."""

import argparse
from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import sys
import tempfile
import time
from urllib.error import HTTPError
from urllib.request import Request, urlopen
import zipfile

from defusedxml.ElementTree import iterparse

from .normalization import normalize


SOURCE_URL = "https://isga.obrnadzor.gov.ru/api/spa/accredreestr/opendata/filezip"
SOURCE_PAGE = "https://isga.obrnadzor.gov.ru/accredreestr/opendata/"
DEFAULT_CACHE = Path(__file__).resolve().parents[2] / "registry_cache"
MAX_ARCHIVE_BYTES = 2 * 1024 ** 3
MAX_XML_BYTES = 4 * 1024 ** 3
SNAPSHOT_PATTERN = r"snapshot-\d{8}-[0-9a-f]{16}"


def utc_now():
    return datetime.now(timezone.utc).isoformat()


def write_json_atomic(path, value):
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


@contextmanager
def cache_lock(cache):
    cache.mkdir(parents=True, exist_ok=True)
    lock = cache / "update.lock"
    try:
        descriptor = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    except FileExistsError:
        raise ValueError("Обновление уже запущено: update.lock существует. После аварии проверьте PID перед удалением блокировки.") from None
    try:
        with os.fdopen(descriptor, "w") as stream:
            stream.write(json.dumps({"pid": os.getpid(), "started_at": utc_now()}))
        yield
    finally:
        lock.unlink(missing_ok=True)


def current_snapshot(cache):
    pointer = cache / "current.json"
    if not pointer.exists():
        return None
    name = json.loads(pointer.read_text(encoding="utf-8"))["snapshot"]
    if not re.fullmatch(SNAPSHOT_PATTERN, name):
        raise ValueError("Некорректный указатель снимка реестра")
    directory = cache / name
    metadata = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
    if not (directory / "index.sqlite").is_file() or not (directory / "source.zip").is_file():
        raise ValueError("Текущий снимок неполный")
    if (directory / "source.zip").stat().st_size != metadata["archive_bytes"]:
        raise ValueError("Размер текущего архива отличается от сохраненного")
    return directory, metadata


def request(url, method="GET", headers=None):
    return Request(url, method=method, headers={"User-Agent": "OURegistryPrototype/0.1",
                                               "Accept-Encoding": "identity", **(headers or {})})


def remote_metadata(opener=urlopen):
    try:
        with opener(request(SOURCE_URL, "HEAD"), timeout=60) as response:
            return dict(response.headers)
    except HTTPError as error:
        if error.code in {405, 501}:
            return {}
        raise


def header(headers, name):
    return next((str(v) for k, v in headers.items() if k.casefold() == name.casefold()), "")


def unchanged(remote, metadata):
    length = header(remote, "Content-Length")
    if not length or int(length) != metadata["archive_bytes"]:
        return False
    etag = header(remote, "ETag")
    if etag and metadata.get("etag"):
        return etag == metadata["etag"]
    modified = header(remote, "Last-Modified")
    return bool(modified and modified == metadata.get("last_modified")
                and header(remote, "Content-Disposition") == metadata.get("content_disposition"))


def download_archive(destination, opener=urlopen):
    digest = hashlib.sha256()
    size = 0
    started = time.monotonic()
    with opener(request(SOURCE_URL), timeout=60) as response, destination.open("wb") as stream:
        headers = dict(response.headers)
        expected = int(header(headers, "Content-Length") or 0)
        if expected > MAX_ARCHIVE_BYTES:
            raise ValueError("Архив превысил допустимый размер 2 ГиБ")
        next_progress = 64 * 1024 ** 2
        while chunk := response.read(1024 ** 2):
            size += len(chunk)
            if size > MAX_ARCHIVE_BYTES or time.monotonic() - started > 1800:
                raise ValueError("Превышен лимит размера или времени скачивания")
            stream.write(chunk)
            digest.update(chunk)
            if size >= next_progress:
                print(f"Скачано {size // (1024 ** 2)} МиБ", file=sys.stderr, flush=True)
                next_progress += 64 * 1024 ** 2
        if not size or (expected and expected != size):
            raise ValueError("Архив скачан не полностью")
    return headers, size, digest.hexdigest()


def child_text(element, name):
    value = element.findtext(name)
    return "" if value is None else value.strip()


def organization_rows(certificate):
    cert_id = child_text(certificate, "Id")
    cert_status = child_text(certificate, "StatusName")
    issue_date = child_text(certificate, "IssueDate")
    region = child_text(certificate, "RegionName")
    parent_inn = child_text(certificate, "EduOrgINN")
    levels = sorted({child_text(p, "EduLevelName") for p in certificate.findall(".//EducationalProgram")
                     if child_text(p, "EduLevelName")})
    scopes = [(certificate, "certificate", cert_status, "", levels)]
    actual = certificate.find("ActualEducationOrganization")
    if actual is not None:
        scopes.append((actual, "actual_organization", cert_status, "", levels))
    for supplement in certificate.findall("./Supplements/Supplement"):
        supplement_id = child_text(supplement, "Id")
        supplement_status = child_text(supplement, "StatusName")
        own_levels = sorted({child_text(p, "EduLevelName") for p in supplement.findall(".//EducationalProgram")
                             if child_text(p, "EduLevelName")})
        scopes.append((supplement, "supplement", supplement_status, supplement_id, own_levels))
        actual = supplement.find("ActualEducationOrganization")
        if actual is not None:
            scopes.append((actual, "supplement_actual_organization", supplement_status, supplement_id, own_levels))
    for element, scope, status, supplement_id, education_levels in scopes:
        actual = "actual_organization" in scope
        prefix = "" if actual else "EduOrg"
        full_name = child_text(element, "FullName" if actual else "EduOrgFullName")
        short_name = child_text(element, "ShortName" if actual else "EduOrgShortName")
        inn = child_text(element, prefix + "INN")
        ogrn = child_text(element, prefix + "OGRN")
        if not full_name and not inn:
            continue
        # Отсутствующий ИНН филиала не заменяется ИНН головной организации.
        yield (cert_id, supplement_id, scope, cert_status, status, issue_date,
               child_text(element, "RegionName") or region, inn, ogrn,
               child_text(element, prefix + "KPP"), full_name, short_name,
               child_text(element, "PostAddress") or child_text(element, "EduOrgAddress"),
               child_text(element, "IsBranch") or child_text(element, "IsForBranch"),
               parent_inn, json.dumps(education_levels, ensure_ascii=False),
               normalize(full_name + " " + short_name))


def build_index(archive_path, database_path):
    connection = sqlite3.connect(database_path)
    certificates = organizations = 0
    try:
        connection.execute("""CREATE TABLE organizations (
            certificate_id TEXT, supplement_id TEXT, scope TEXT, certificate_status TEXT,
            record_status TEXT, issue_date TEXT, region TEXT, inn TEXT, ogrn TEXT, kpp TEXT,
            full_name TEXT, short_name TEXT, post_address TEXT, is_branch TEXT, parent_inn TEXT,
            education_levels TEXT, search_name TEXT)""")
        with zipfile.ZipFile(archive_path) as archive:
            members = [member for member in archive.infolist() if not member.is_dir() and member.filename.lower().endswith(".xml")]
            if len(members) != 1:
                raise ValueError("Ожидался ровно один XML в архиве")
            member = members[0]
            if member.file_size <= 0 or member.file_size > MAX_XML_BYTES:
                raise ValueError("Недопустимый размер XML")
            # Не извлекаем архив на диск, поэтому пути внутри ZIP не используются как пути файлов.
            with archive.open(member) as stream:
                stack = []
                for event, element in iterparse(stream, events=("start", "end"), forbid_dtd=True,
                                                forbid_entities=True, forbid_external=True):
                    if event == "start":
                        if not stack and element.tag != "OpenData":
                            raise ValueError("Неизвестный корневой элемент XML: ожидается OpenData")
                        stack.append(element)
                        continue
                    if element.tag == "Certificate" and len(stack) == 3 and stack[-2].tag == "Certificates":
                        rows = list(organization_rows(element))
                        connection.executemany("INSERT INTO organizations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", rows)
                        organizations += len(rows)
                        certificates += 1
                        stack[-2].remove(element)
                        element.clear()
                        if certificates % 10000 == 0:
                            connection.commit()
                            print(f"Прочитано {certificates} свидетельств", file=sys.stderr, flush=True)
                    stack.pop()
            if not certificates or not organizations:
                raise ValueError("В XML не найдены свидетельства и организации")
            xml_name, xml_size = member.filename, member.file_size
        connection.execute("CREATE INDEX inn_idx ON organizations(inn)")
        connection.execute("CREATE INDEX ogrn_idx ON organizations(ogrn)")
        connection.commit()
        if connection.execute("PRAGMA quick_check").fetchone()[0] != "ok":
            raise ValueError("Индекс SQLite не прошел проверку")
    finally:
        connection.close()
    return {"certificates": certificates, "organization_records": organizations,
            "xml_member": xml_name, "xml_bytes": xml_size}


def prune_snapshots(cache, current_name, keep):
    snapshots = [p for p in cache.iterdir() if p.is_dir() and re.fullmatch(SNAPSHOT_PATTERN, p.name)]
    older = sorted((p for p in snapshots if p.name != current_name), key=lambda p: p.stat().st_mtime, reverse=True)
    for path in older[max(0, keep - 1):]:
        shutil.rmtree(path)


def update(cache, keep=2, force=False, opener=urlopen):
    cache = Path(cache)
    if keep < 2:
        raise ValueError("Нужно сохранять минимум два снимка")
    with cache_lock(cache):
        try:
            current = current_snapshot(cache)
            remote = remote_metadata(opener)
            if current and not force and unchanged(remote, current[1]):
                report = {"status": "unchanged", "checked_at": utc_now(), "snapshot": current[0].name,
                          "source_url": SOURCE_URL, "manifest": current[1]}
                write_json_atomic(cache / "last_check.json", report)
                return report
            with tempfile.TemporaryDirectory(prefix=".building-", dir=cache) as temporary:
                directory = Path(temporary)
                headers, size, digest = download_archive(directory / "source.zip", opener)
                if current and digest == current[1]["sha256"]:
                    report = {"status": "unchanged", "checked_at": utc_now(), "snapshot": current[0].name,
                              "source_url": SOURCE_URL, "manifest": current[1], "verified_download": True}
                    write_json_atomic(cache / "last_check.json", report)
                    return report
                details = build_index(directory / "source.zip", directory / "index.sqlite")
                date_match = re.search(r"data-(\d{8})-structure-\d{8}\.xml", details["xml_member"])
                if not date_match:
                    raise ValueError("Дата выгрузки не найдена в имени XML")
                snapshot_date = date_match.group(1)
                metadata = {"source_kind": "accreditation_registry", "source_page": SOURCE_PAGE,
                            "source_url": SOURCE_URL, "snapshot_date": snapshot_date, "downloaded_at": utc_now(),
                            "archive_bytes": size, "sha256": digest, "etag": header(headers, "ETag"),
                            "last_modified": header(headers, "Last-Modified"),
                            "content_disposition": header(headers, "Content-Disposition"), **details}
                write_json_atomic(directory / "manifest.json", metadata)
                name = f"snapshot-{snapshot_date}-{digest[:16]}"
                destination = cache / name
                if destination.exists():
                    raise ValueError("Такой снимок уже существует; сохраненная версия не перезаписана")
                os.replace(directory, destination)
                write_json_atomic(cache / "current.json", {"snapshot": name})
            report = {"status": "updated", "checked_at": utc_now(), "snapshot": name,
                      "source_url": SOURCE_URL, "manifest": metadata}
            write_json_atomic(cache / "last_check.json", report)
            prune_snapshots(cache, name, keep)
            return report
        except Exception as error:
            write_json_atomic(cache / "last_check.json", {"status": "error", "checked_at": utc_now(),
                              "error_type": type(error).__name__, "error": str(error), "source_url": SOURCE_URL})
            raise


def lookup(cache, inn=None, name=None, limit=20):
    if not 1 <= limit <= 100:
        raise ValueError("Лимит результатов: 1–100")
    snapshot = current_snapshot(Path(cache))
    if snapshot is None:
        raise ValueError("Нет локального снимка; сначала выполните update")
    if bool(inn) == bool(name):
        raise ValueError("Укажите только один ключ: ИНН или название")
    conditions, params = [], []
    if inn:
        if not re.fullmatch(r"\d{10}|\d{12}", inn):
            raise ValueError("ИНН должен содержать 10 или 12 цифр")
        conditions.append("inn = ?")
        params.append(inn)
    else:
        tokens = normalize(name).split()
        if not tokens:
            raise ValueError("Пустое название для поиска")
        for token in tokens[:20]:
            conditions.append("instr(search_name, ?) > 0")
            params.append(token)
    sql = "SELECT * FROM organizations WHERE " + " AND ".join(conditions)
    sql += " ORDER BY (certificate_status = 'Действующее') DESC, issue_date DESC LIMIT ?"
    with sqlite3.connect((snapshot[0] / "index.sqlite").resolve().as_uri() + "?mode=ro", uri=True) as db:
        db.row_factory = sqlite3.Row
        rows = [dict(row) for row in db.execute(sql, [*params, limit + 1])]
    for row in rows:
        row["education_levels"] = json.loads(row["education_levels"])
        row.pop("search_name", None)
    return {"source_kind": "accreditation_registry", "snapshot_date": snapshot[1]["snapshot_date"],
            "source_url": SOURCE_URL, "truncated": len(rows) > limit,
            "notice": "Статус свидетельства не является статусом юрлица. PostAddress не подтверждает юридический адрес.",
            "matches": rows[:limit]}


def main(argv=None):
    parser = argparse.ArgumentParser(description="Резервная локальная копия реестра аккредитации Рособрнадзора")
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    commands = parser.add_subparsers(dest="command", required=True)
    updating = commands.add_parser("update", help="Проверить и при необходимости скачать новый снимок")
    updating.add_argument("--keep", type=int, default=2)
    updating.add_argument("--force", action="store_true")
    finding = commands.add_parser("lookup", help="Поиск по локальному индексу без сети")
    keys = finding.add_mutually_exclusive_group(required=True)
    keys.add_argument("--inn")
    keys.add_argument("--name")
    finding.add_argument("--limit", type=int, default=20)
    commands.add_parser("status", help="Сведения о текущем снимке и последнем обновлении")
    args = parser.parse_args(argv)
    try:
        if args.command == "update":
            result = update(args.cache, args.keep, args.force)
        elif args.command == "lookup":
            result = lookup(args.cache, args.inn, args.name, args.limit)
        else:
            snapshot = current_snapshot(args.cache)
            result = {"snapshot": snapshot[0].name if snapshot else None,
                      "manifest": snapshot[1] if snapshot else None}
            check = args.cache / "last_check.json"
            result["last_check"] = json.loads(check.read_text(encoding="utf-8")) if check.exists() else None
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as error:
        print(f"Ошибка реестра: {type(error).__name__}: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
