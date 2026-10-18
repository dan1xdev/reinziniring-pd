"""Запуск: python -m ou_processor INPUT.xlsx --output result."""

import argparse
from datetime import datetime, timezone
import getpass
import json
import os
from pathlib import Path
import sys

from .dadata import DaDataClient, run_sample
from .export import write_result
from .input import read_records
from .pipeline import process


def main(argv=None):
    parser = argparse.ArgumentParser(description="Предварительная обработка образовательных учреждений")
    parser.add_argument("inputs", nargs="+", type=Path, help="Один или несколько Excel .xlsx")
    parser.add_argument("--config", type=Path, default=Path(__file__).resolve().parent.parent / "config.json")
    parser.add_argument("--output", type=Path, required=True, help="Новая папка результата")
    parser.add_argument("--online", action="store_true", help="Выполнить пробный поиск через DaData")
    parser.add_argument("--dadata-limit", type=int, default=5, help="Максимум запросов (0–20, по умолчанию 5)")
    parser.add_argument("--token-env", default="DADATA_API_KEY", help="Имя переменной с токеном")
    parser.add_argument("--ask-token", action="store_true", help="Ввести API-токен без отображения на экране")
    args = parser.parse_args(argv)
    if not 0 <= args.dadata_limit <= 20:
        parser.error("--dadata-limit должен быть от 0 до 20")
    if args.ask_token and not args.online:
        parser.error("--ask-token используется только вместе с --online")
    if args.output.exists() and not args.output.is_dir():
        parser.error("--output должен указывать на папку")
    if args.output.exists() and any(args.output.iterdir()):
        parser.error("Папка результата не пуста. Укажите новую папку, чтобы сохранить предыдущий запуск.")
    try:
        config = json.loads(args.config.read_text(encoding="utf-8"))
        records, sources = read_records(args.inputs)
        result = process(records, config)
        result["sources"] = sources
        result["rules"] = config
        result["created_at_utc"] = datetime.now(timezone.utc).isoformat()
        token = os.environ.get(args.token_env, "").strip()
        if args.ask_token:
            token = getpass.getpass("API-токен DaData: ").strip()
        client = DaDataClient(token) if args.online and token else None
        result["dadata"] = run_sample(result, args.dadata_limit, client, args.online)
        output = write_result(result, args.output)
    except (OSError, ValueError, KeyError) as error:
        print(f"Ошибка: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result["summary"], ensure_ascii=False, indent=2))
    print(f"Результат: {output}")
    print(f"DaData: {result['dadata']['mode']}, запросов: {result['dadata']['actual_requests']}")
    success = {"КАНДИДАТЫ_ДЛЯ_ПРОВЕРКИ", "НЕ_НАЙДЕНО"}
    if args.online and any(r["status"] not in success for r in result["dadata"]["requests"]):
        print("API-прогон завершен не полностью. Подробности на листе «Запросы DaData».", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
