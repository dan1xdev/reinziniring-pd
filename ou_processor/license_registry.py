"""Реестр лицензий Рособрнадзора: основной источник сопоставления.

Официальная выгрузка (набор 7701537808-FBDRL «Сводный реестр лицензий на
осуществление образовательной деятельности») скачивается с сайта Рособрнадзора
и превращается в локальный SQLite-индекс. Периодичность обновления и источник
задаются в config.json (раздел license_registry). Пока новое обновление
не скачано полностью, работает предыдущий снимок; при сбое он же сохраняется.
"""

import argparse
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
import hashlib
import io
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import sys
import tempfile
import time
import csv
from urllib.error import HTTPError
from urllib.parse import urljoin
from urllib.request import Request, urlopen
import zipfile

from defusedxml.ElementTree import iterparse

from .normalization import normalize


SOURCE_PAGE = "https://islod.obrnadzor.gov.ru/opendata_licenses.html"
DATA_LINK_RE = re.compile(r"[^\s\"'<>]*data-(\d{8})-structure-\d{8}\.zip")
DEFAULT_CACHE = Path(__file__).resolve().parents[2] / "license_registry_cache"
DEFAULT_CONFIG = Path(__file__).resolve().parents[1] / "config.json"
MAX_ARCHIVE_BYTES = 2 * 1024 ** 3
SNAPSHOT_PATTERN = r"snapshot-\d{8}-[0-9a-f]{16}"
DATA_FILE = "source.data"

# Названия полей различаются между версиями выгрузки, поэтому поля
# распознаются по известным вариантам имени тега XML или заголовка колонки.
FIELD_SYNONYMS = {
    "full_name": ("fullname", "полноенаименование", "eduorgfullname", "наименование"),
    "short_name": ("shortname", "краткоенаименование", "eduorgshortname"),
    "inn": ("inn", "инн", "eduorginn"),
    "ogrn": ("ogrn", "огрн", "eduorgogrn"),
    "kpp": ("kpp", "кпп", "eduorgkpp"),
    "license_number": ("licenseregnumber", "licensenumber", "регномер", "номерлицензии",
                       "реестровыйномер", "regnumber"),
    "status": ("statusname", "статуслицензии", "статус", "status"),
    "date_start": ("datestart", "датаначала", "dateresolution", "датарешения", "date", "дата"),
    "region": ("regionname", "субъектрф", "регион", "region"),
    "address": ("postaddress", "юридическийадрес", "eduorgaddress", "адрес", "address"),
}


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
    """Возвращает (папка, manifest) последнего рабочего снимка или None."""
    pointer = cache / "current.json"
    if not pointer.exists():
        return None
    name = json.loads(pointer.read_text(encoding="utf-8"))["snapshot"]
    if not re.fullmatch(SNAPSHOT_PATTERN, name):
        raise ValueError("Некорректный указатель снимка реестра лицензий")
    directory = cache / name
    metadata = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
    if not (directory / "index.sqlite").is_file() or not (directory / DATA_FILE).is_file():
        raise ValueError("Текущий снимок неполный")
    if (directory / DATA_FILE).stat().st_size != metadata["archive_bytes"]:
        raise ValueError("Размер текущего файла выгрузки отличается от сохраненного")
    return directory, metadata


def field_key(tag):
    """Сопоставляет тег XML или заголовок колонки с полем индекса."""
    key = re.sub(r"[^0-9a-zа-яё]", "", str(tag).casefold())
    for field, variants in FIELD_SYNONYMS.items():
        if key in variants:
            return field
    if "licens" in key and ("номер" in key or "рег" in key or "number" in key):
        return "license_number"
    if "статус" in key or "status" in key:
        return "status"
    if key.startswith(("date", "дата")):
        return "date_start"
    if "краткое" in key:
        return "short_name"
    if "наименование" in key or "название" in key or "fullname" in key:
        return "full_name"
    if "адрес" in key or "address" in key:
        return "address"
    if "регион" in key or "субъект" in key:
        return "region"
    return ""


def records_from_xml(stream):
    """Выдаёт записи выгрузки из XML. Запись — элемент, у которого все
    вложенные элементы простые (текстовые) и их не меньше трёх."""
    stack = []
    for event, element in iterparse(stream, events=("start", "end"), forbid_dtd=True,
                                    forbid_entities=True, forbid_external=True):
        if event == "start":
            stack.append(element)
            continue
        stack.pop()
        if len(element) < 3 or any(len(child) for child in element):
            continue
        fields = {}
        for child in element:
            value = (child.text or "").strip()
            key = field_key(child.tag)
            if key and value and key not in fields:
                fields[key] = value
        if stack:
            stack[-1].remove(element)
        element.clear()
        if fields.get("full_name") or fields.get("inn") or fields.get("license_number"):
            yield fields


def records_from_csv(stream):
    """Выдаёт записи из CSV: первая строка считается заголовком колонок."""
    text_stream = io.TextIOWrapper(stream, encoding="utf-8-sig", newline="")
    try:
        try:
            first = text_stream.readline()
        except UnicodeDecodeError:
            text_stream.detach()
            stream.seek(0)
            text_stream = io.TextIOWrapper(stream, encoding="cp1251", newline="")
            first = text_stream.readline()
        text_stream.seek(0)
        delimiter = ";" if first.count(";") >= first.count(",") else ","
        rows = csv.reader(text_stream, delimiter=delimiter)
        keys = [field_key(header) for header in next(rows, [])]
        for row in rows:
            if not any(cell.strip() for cell in row):
                continue
            fields = {}
            for key, value in zip(keys, row):
                value = value.strip()
                if key and value and key not in fields:
                    fields[key] = value
            if fields.get("full_name") or fields.get("inn") or fields.get("license_number"):
                yield fields
    finally:
        text_stream.detach()


def records_from_xlsx(stream):
    """Выдаёт записи из книги Excel: первая строка считается заголовком."""
    from openpyxl import load_workbook

    book = load_workbook(stream, read_only=True, data_only=True)
    try:
        sheet = book[book.sheetnames[0]]
        rows = sheet.iter_rows(values_only=True)
        keys = [field_key(header) for header in next(rows, ())]
        for row in rows:
            fields = {}
            for key, value in zip(keys, row):
                text_value = "" if value is None else str(value).strip()
                if key and text_value and key not in fields:
                    fields[key] = text_value
            if fields.get("full_name") or fields.get("inn") or fields.get("license_number"):
                yield fields
    finally:
        book.close()


RECORD_READERS = {"xml": records_from_xml, "csv": records_from_csv, "xlsx": records_from_xlsx}


def records_from_export(path):
    """Определяет формат официальной выгрузки и выдаёт из неё записи."""
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            data = [member for member in archive.infolist() if not member.is_dir()
                    and member.filename.lower().endswith((".xml", ".csv", ".xlsx"))]
            if len(data) != 1:
                raise ValueError("Ожидался ровно один файл данных в архиве")
            member = data[0]
            kind = member.filename.lower().rsplit(".", 1)[-1]
            with archive.open(member) as stream:
                yield from RECORD_READERS[kind](stream)
        return
    with open(path, "rb") as stream:
        reader = RECORD_READERS.get(path.suffix.lower().lstrip("."), records_from_xml)
        yield from reader(stream)


def build_index(source_path, database_path):
    """Собирает SQLite-индекс лицензий из официальной выгрузки."""
    connection = sqlite3.connect(database_path)
    records = 0
    try:
        connection.execute("""CREATE TABLE licenses (
            inn TEXT, ogrn TEXT, kpp TEXT, full_name TEXT, short_name TEXT,
            license_number TEXT, date_start TEXT, status TEXT, region TEXT,
            address TEXT, search_name TEXT)""")
        for fields in records_from_export(source_path):
            full_name = fields.get("full_name", "")
            if not full_name and not fields.get("inn"):
                continue
            connection.execute("INSERT INTO licenses VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (fields.get("inn", ""), fields.get("ogrn", ""), fields.get("kpp", ""),
                 full_name, fields.get("short_name", ""), fields.get("license_number", ""),
                 fields.get("date_start", ""), fields.get("status", ""),
                 fields.get("region", ""), fields.get("address", ""),
                 normalize(full_name + " " + fields.get("short_name", ""))))
            records += 1
            if records % 10000 == 0:
                connection.commit()
                print(f"Прочитано {records} записей", file=sys.stderr, flush=True)
        if not records:
            raise ValueError("В выгрузке не найдено ни одной записи о лицензии")
        connection.execute("CREATE INDEX inn_idx ON licenses(inn)")
        connection.commit()
        if connection.execute("PRAGMA quick_check").fetchone()[0] != "ok":
            raise ValueError("Индекс SQLite не прошел проверку")
    finally:
        connection.close()
    return {"records": records}


def load_settings(config_path=DEFAULT_CONFIG):
    """Читает раздел license_registry из config.json."""
    try:
        config = json.loads(Path(config_path).read_text(encoding="utf-8"))
    except Exception:
        return {}
    settings = config.get("license_registry")
    return settings if isinstance(settings, dict) else {}


def resolve_data_url(settings, opener=urlopen):
    """Определяет ссылку на актуальную официальную выгрузку.

    Сначала берётся прямая ссылка из настроек, затем страница открытых данных,
    с которой снимается самый свежий файл data-ДДДДММГГ-structure-*.zip."""
    direct = str(settings.get("source_url", "") or "").strip()
    if direct:
        return direct
    # Пустая source_page в настройках отключает чтение страницы:
    # источник задаётся напрямую или локальным файлом.
    page = str(settings.get("source_page") if settings.get("source_page") is not None else SOURCE_PAGE).strip()
    if not page:
        return ""
    request = Request(page, headers={"User-Agent": "OURegistryPrototype/0.1",
                                     "Accept-Encoding": "identity"})
    with opener(request, timeout=60) as response:
        html = response.read(5 * 1024 * 1024).decode("utf-8", "replace")
    candidates = [(match.group(1), match.group(0)) for match in DATA_LINK_RE.finditer(html)]
    if not candidates:
        raise ValueError("На странице открытых данных не найдена ссылка на файл выгрузки")
    date, link = max(candidates)
    return urljoin(page, link)


def resolve_source(settings, opener=urlopen):
    """Выбирает источник выгрузки: сеть или локальный файл из настроек.

    Локальный файл (source_file в config.json) используется, когда сайт
    недоступен, — так проверка модуля возможна без интернета."""
    local = str(settings.get("source_file", "") or "").strip()
    try:
        url = resolve_data_url(settings, opener)
        if url:
            return {"url": url}
    except Exception:
        if not local:
            raise
        print("Сайт недоступен, использую локальный файл выгрузки из настроек",
              file=sys.stderr, flush=True)
    if local:
        path = Path(local)
        if not path.is_absolute():
            path = Path(__file__).resolve().parents[2] / path
        if not path.is_file():
            raise ValueError("Локальный файл выгрузки не найден: " + str(path))
        return {"file": str(path)}
    raise ValueError("Источник выгрузки не задан в config.json")


def remote_metadata(url, opener=urlopen):
    try:
        request = Request(url, method="HEAD", headers={"User-Agent": "OURegistryPrototype/0.1"})
        with opener(request, timeout=60) as response:
            return dict(response.headers)
    except HTTPError as error:
        if error.code in {405, 501}:
            return {}
        raise


def header(headers, name):
    return next((str(v) for k, v in (headers or {}).items() if k.casefold() == name.casefold()), "")


def unchanged(remote, metadata):
    length = header(remote, "Content-Length")
    if not length or int(length) != metadata["archive_bytes"]:
        return False
    etag = header(remote, "ETag")
    if etag and metadata.get("etag"):
        return etag == metadata["etag"]
    modified = header(remote, "Last-Modified")
    return bool(modified and modified == metadata.get("last_modified"))


def _save_archive(source, destination):
    """Записывает выгрузку в снимок (из сети или из локального файла)."""
    digest = hashlib.sha256()
    size = 0
    started = time.monotonic()
    headers = dict(getattr(source, "headers", {}) or {})
    expected = int(header(headers, "Content-Length") or 0)
    if expected > MAX_ARCHIVE_BYTES:
        raise ValueError("Файл выгрузки превысил допустимый размер 2 ГиБ")
    next_progress = 64 * 1024 ** 2
    with destination.open("wb") as writer:
        while chunk := source.read(1024 ** 2):
            size += len(chunk)
            if size > MAX_ARCHIVE_BYTES or time.monotonic() - started > 1800:
                raise ValueError("Превышен лимит размера или времени скачивания")
            writer.write(chunk)
            digest.update(chunk)
            if size >= next_progress:
                print(f"Скачано {size // (1024 ** 2)} МиБ", file=sys.stderr, flush=True)
                next_progress += 64 * 1024 ** 2
    if not size or (expected and expected != size):
        raise ValueError("Файл выгрузки скачан не полностью")
    return headers, size, digest.hexdigest()


def prune_snapshots(cache, current_name, keep):
    snapshots = [p for p in cache.iterdir() if p.is_dir() and re.fullmatch(SNAPSHOT_PATTERN, p.name)]
    older = sorted((p for p in snapshots if p.name != current_name), key=lambda p: p.stat().st_mtime, reverse=True)
    for path in older[max(0, keep - 1):]:
        shutil.rmtree(path)


def update(cache, keep=None, force=False, opener=urlopen, settings=None):
    """Проверяет источник и при необходимости собирает новый снимок.

    Новый снимок собирается во временной папке; указатель current.json
    переключается только после успешной сборки. Во время обновления
    и при любом сбое предыдущий снимок остаётся рабочим."""
    cache = Path(cache)
    settings = settings if settings is not None else load_settings()
    if keep is None:
        keep = int(settings.get("keep_snapshots", 2) or 2)
    if keep < 2:
        raise ValueError("Нужно сохранять минимум два снимка")
    update_days = int(settings.get("update_days", 0) or 0)
    with cache_lock(cache):
        try:
            current = current_snapshot(cache)
            # 1. По настройкам свежий снимок можно не проверять в сети.
            if current and not force and update_days:
                try:
                    downloaded = datetime.fromisoformat(current[1]["downloaded_at"])
                except (KeyError, ValueError):
                    downloaded = None
                if downloaded is not None and datetime.now(timezone.utc) - downloaded < timedelta(days=update_days):
                    report = {"status": "fresh", "checked_at": utc_now(), "snapshot": current[0].name,
                              "update_days": update_days, "manifest": current[1]}
                    write_json_atomic(cache / "last_check.json", report)
                    return report
            source = resolve_source(settings, opener)
            is_local = "file" in source
            url = source.get("file", source.get("url", ""))
            remote = {} if is_local else remote_metadata(source["url"], opener)
            if current and not force and not is_local and unchanged(remote, current[1]) \
                    and current[1].get("source_url") == source["url"]:
                report = {"status": "unchanged", "checked_at": utc_now(), "snapshot": current[0].name,
                          "source_url": source["url"], "manifest": current[1]}
                write_json_atomic(cache / "last_check.json", report)
                return report
            with tempfile.TemporaryDirectory(prefix=".building-", dir=cache) as temporary:
                directory = Path(temporary)
                if is_local:
                    with open(url, "rb") as reader:
                        headers, size, digest = _save_archive(reader, directory / DATA_FILE)
                    url = "file:" + source["file"]
                else:
                    request = Request(source["url"], headers={"User-Agent": "OURegistryPrototype/0.1",
                                                              "Accept-Encoding": "identity"})
                    with opener(request, timeout=120) as response:
                        headers, size, digest = _save_archive(response, directory / DATA_FILE)
                if current and digest == current[1]["sha256"]:
                    report = {"status": "unchanged", "checked_at": utc_now(), "snapshot": current[0].name,
                              "source_url": url, "manifest": current[1], "verified_download": True}
                    write_json_atomic(cache / "last_check.json", report)
                    return report
                details = build_index(directory / DATA_FILE, directory / "index.sqlite")
                date_match = re.search(r"data-(\d{8})-", url)
                snapshot_date = date_match.group(1) if date_match else time.strftime("%Y%m%d")
                metadata = {"source_kind": "license_registry",
                            "source_page": str(settings.get("source_page", "") or SOURCE_PAGE),
                            "source_url": url, "snapshot_date": snapshot_date, "downloaded_at": utc_now(),
                            "archive_bytes": size, "sha256": digest,
                            "etag": header(headers, "ETag"), "last_modified": header(headers, "Last-Modified"),
                            **details}
                write_json_atomic(directory / "manifest.json", metadata)
                name = f"snapshot-{snapshot_date}-{digest[:16]}"
                destination = cache / name
                if destination.exists():
                    raise ValueError("Такой снимок уже существует; сохраненная версия не перезаписана")
                os.replace(directory, destination)
                write_json_atomic(cache / "current.json", {"snapshot": name})
            report = {"status": "updated", "checked_at": utc_now(), "snapshot": name,
                      "source_url": url, "manifest": metadata}
            write_json_atomic(cache / "last_check.json", report)
            prune_snapshots(cache, name, keep)
            return report
        except Exception as error:
            write_json_atomic(cache / "last_check.json", {"status": "error", "checked_at": utc_now(),
                              "error_type": type(error).__name__, "error": str(error),
                              "source_page": str(settings.get("source_page", "") or SOURCE_PAGE)})
            raise


def auto_update(cache, config_path=DEFAULT_CONFIG):
    """Обновление по настройкам из config.json (для интерфейса и планировщика)."""
    return update(cache, keep=None, settings=load_settings(config_path))


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
    sql = "SELECT * FROM licenses WHERE " + " AND ".join(conditions)
    sql += " ORDER BY (status = 'Действующая') DESC, date_start DESC LIMIT ?"
    with sqlite3.connect((snapshot[0] / "index.sqlite").resolve().as_uri() + "?mode=ro", uri=True) as db:
        db.row_factory = sqlite3.Row
        rows = [dict(row) for row in db.execute(sql, [*params, limit + 1])]
    for row in rows:
        row.pop("search_name", None)
    return {"source_kind": "license_registry", "snapshot_date": snapshot[1]["snapshot_date"],
            "source_url": snapshot[1].get("source_url", ""),
            "matches": rows[:limit]}


def main(argv=None):
    parser = argparse.ArgumentParser(description="Локальный индекс реестра лицензий Рособрнадзора")
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    commands = parser.add_subparsers(dest="command", required=True)
    updating = commands.add_parser("update", help="Проверить источник и при необходимости собрать новый снимок")
    updating.add_argument("--keep", type=int, default=None)
    updating.add_argument("--force", action="store_true")
    finding = commands.add_parser("lookup", help="Поиск по локальному индексу без сети")
    keys = finding.add_mutually_exclusive_group(required=True)
    keys.add_argument("--inn")
    keys.add_argument("--name")
    finding.add_argument("--limit", type=int, default=20)
    commands.add_parser("status", help="Сведения о текущем снимке и последней проверке")
    args = parser.parse_args(argv)
    try:
        settings = load_settings(args.config)
        if args.command == "update":
            result = update(args.cache, keep=args.keep, force=args.force, settings=settings)
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
        print(f"Ошибка реестра лицензий: {type(error).__name__}: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
