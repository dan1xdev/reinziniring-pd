"""Ограниченный пробный поиск организаций. Кандидат не равен подтвержденному ИНН."""

from datetime import datetime, timezone
import json
import socket
import time
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from .normalization import numbers, text


ENDPOINT = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/party"


class DaDataError(Exception):
    def __init__(self, code):
        self.code = code
        super().__init__(code)


class DaDataClient:
    def __init__(self, token, timeout=15, opener=urlopen):
        if not token or "\n" in token or "\r" in token:
            raise ValueError("Некорректный API-токен")
        self.token = token
        self.timeout = timeout
        self.opener = opener

    def suggest(self, query):
        payload = {"query": query, "count": 5, "type": "LEGAL"}
        request = Request(ENDPOINT, data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                          headers={"Authorization": "Token " + self.token,
                                   "Content-Type": "application/json", "Accept": "application/json"},
                          method="POST")
        try:
            with self.opener(request, timeout=self.timeout) as response:
                raw = response.read(2_000_001)
            if len(raw) > 2_000_000:
                raise DaDataError("ОТВЕТ_СЛИШКОМ_БОЛЬШОЙ")
            body = json.loads(raw)
            suggestions = body.get("suggestions") if isinstance(body, dict) else None
            if not isinstance(suggestions, list) or any(
                not isinstance(item, dict) or not isinstance(item.get("data"), dict)
                for item in suggestions
            ):
                raise DaDataError("НЕВЕРНЫЙ_ФОРМАТ_ОТВЕТА")
            return suggestions
        except HTTPError as error:
            codes = {401: "ОШИБКА_АВТОРИЗАЦИИ", 403: "ДОСТУП_ЗАПРЕЩЕН", 429: "ЛИМИТ_ЗАПРОСОВ"}
            raise DaDataError(codes.get(error.code, f"HTTP_{error.code}")) from None
        except (URLError, TimeoutError, socket.timeout):
            raise DaDataError("ОШИБКА_СЕТИ") from None
        except (ValueError, UnicodeDecodeError):
            raise DaDataError("НЕВЕРНЫЙ_ФОРМАТ_ОТВЕТА") from None


def sample_groups(result, limit):
    # Сначала полные географические данные; затем стабильный порядок групп.
    by_id = {r["record_id"]: r for r in result["records"]}
    groups = sorted(result["groups"], key=lambda g: (
        not bool(g["region"] and g["city"]),
        len(by_id[g["record_ids"][0]]["flags"]), g["group_id"]))
    samples, seen = [], set()
    for group in groups:
        query = " ".join(text(group["name"]).split())
        for field in ["region", "city"]:
            if group[field] and group[field].casefold() not in query.casefold():
                query += " " + group[field]
        if query in seen:
            continue
        seen.add(query)
        samples.append({"group_id": group["group_id"], "query": query})
        if len(samples) >= limit:
            break
    return samples if limit > 0 else []


def run_sample(result, limit=5, client=None, online=False, delay=0.2):
    samples = sample_groups(result, limit)
    requests, candidates = [], []
    stopped = False
    request_count = 0
    by_group = {group["group_id"]: group for group in result["groups"]}
    for sample in samples:
        entry = {**sample, "timestamp_utc": datetime.now(timezone.utc).isoformat(),
                 "status": "ПОДГОТОВЛЕН_ОФЛАЙН", "candidate_count": 0}
        if not online:
            pass
        elif client is None:
            entry["status"] = "НЕ_ВЫПОЛНЕНО_НЕТ_КЛЮЧА"
        elif stopped:
            entry["status"] = "ПРОПУЩЕНО_ПОСЛЕ_ОШИБКИ_ДОСТУПА"
        else:
            try:
                if request_count and delay:
                    time.sleep(delay)
                request_count += 1
                found = client.suggest(sample["query"])
                converted = [candidate_row(sample, item, rank, by_group[sample["group_id"]])
                             for rank, item in enumerate(found, 1)]
                candidates.extend(converted)
                entry.update(status="КАНДИДАТЫ_ДЛЯ_ПРОВЕРКИ" if found else "НЕ_НАЙДЕНО", candidate_count=len(found))
            except DaDataError as error:
                entry["status"] = error.code
                stopped = error.code in {"ОШИБКА_АВТОРИЗАЦИИ", "ДОСТУП_ЗАПРЕЩЕН", "ЛИМИТ_ЗАПРОСОВ"}
            except (AttributeError, TypeError, ValueError):
                entry["status"] = "НЕВЕРНЫЙ_ФОРМАТ_ОТВЕТА"
        requests.append(entry)
    return {"mode": "online" if online else "offline", "endpoint": ENDPOINT,
            "actual_requests": request_count, "requests": requests, "candidates": candidates}


def candidate_row(sample, suggestion, rank, group):
    data = suggestion["data"]
    name = data.get("name") or {}
    address = data.get("address") or {}
    address_data = address.get("data") or {}
    state = data.get("state") or {}
    full_name = name.get("full_with_opf") or suggestion.get("value") or ""
    candidate_numbers = numbers(full_name)
    source_numbers = numbers(group["name"])
    flags = []
    if source_numbers and candidate_numbers != source_numbers:
        flags.append("НОМЕР_НЕ_ПОДТВЕРЖДЕН" if not candidate_numbers else "КОНФЛИКТ_НОМЕРА")
    if state.get("status") != "ACTIVE":
        flags.append("СТАТУС_НЕ_ПОДТВЕРЖДЕН_ДЕЙСТВУЮЩИМ")
    if not data.get("inn"):
        flags.append("НЕТ_ИНН")
    return {"group_id": sample["group_id"], "rank": rank, "query": sample["query"],
            "candidate_name": full_name, "candidate_inn": data.get("inn", ""),
            "candidate_kpp": data.get("kpp", ""), "candidate_ogrn": data.get("ogrn", ""),
            "legal_status": state.get("status", ""), "branch_type": data.get("branch_type", ""),
            "address": address.get("unrestricted_value") or address.get("value", ""),
            "postal_code": address_data.get("postal_code", ""),
            "region": address_data.get("region_with_type", ""),
            "city": address_data.get("city_with_type") or address_data.get("settlement_with_type", ""),
            "flags": flags, "decision": "ПРОВЕРИТЬ", "confirmed": False}
