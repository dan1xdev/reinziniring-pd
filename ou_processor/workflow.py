"""Общий сценарий UI и CLI: реестр и отдельная диагностическая сверка DaData."""

from copy import deepcopy
from datetime import datetime, timezone
import json
from pathlib import Path
import sqlite3

from .resolution import RegistrySearch, SearchCache, resolve


def search(result, registry_path, config, *, cache=None, client=None,
           on_progress=None, check_cancel=None):
    registry = None
    warnings = []
    result['processing_mode'] = 'dadata_check' if client is not None else 'registry'
    total = len(result['groups']) * (2 if client is not None else 1)

    def report(offset):
        return lambda done, _: on_progress(done + offset, total) if on_progress else None

    try:
        try:
            registry = RegistrySearch(registry_path, config=config)
        except (ValueError, OSError, KeyError, sqlite3.Error):
            warnings.append('Локальный реестр недоступен. Несопоставленные учреждения переданы на проверку; автоматического обращения к DaData нет.')
        try:
            resolve(result, None, registry, progress=False, config=config,
                    on_progress=report(0), check_cancel=check_cancel)
        except sqlite3.Error:
            warnings.append('Не удалось прочитать индекс реестра. Учреждения переданы на проверку; обновите локальный реестр.')
            resolve(result, None, None, progress=False, config=config,
                    on_progress=report(0), check_cancel=check_cancel)
    finally:
        if registry is not None:
            registry.connection.close()

    if client is not None:
        # Сверка не меняет выбранный реестром ИНН и не подтверждает новые случаи.
        diagnostic = {'records': deepcopy(result['records']), 'groups': deepcopy(result['groups'])}
        api = SearchCache(cache, client)
        resolve(diagnostic, api, None, progress=False, config=config,
                on_progress=report(len(result['groups'])), check_cancel=check_cancel)
        checks = []
        for primary, checked in zip(result['resolution']['groups'], diagnostic['resolution']['groups']):
            original = primary['selected']
            suggested = checked['selected']
            original_inn = original['candidate_inn'] if original else ''
            suggested_inn = suggested['candidate_inn'] if suggested else ''
            status = ('СОВПАДАЕТ' if original_inn == suggested_inn else 'КОНФЛИКТ_ИНН') if original_inn and suggested_inn else 'ТРЕБУЕТ_ПРОВЕРКИ'
            checks.append({'group_id': primary['group_id'], 'registry_inn': original_inn,
                           'dadata_inn': suggested_inn, 'status': status})
        result['dadata_checks'] = checks
        result['dadata_candidates'] = diagnostic['resolution']['candidates']
        result['resolution']['requests'] = api.requests
        result['resolution']['actual_requests'] = api.actual_requests
        if any(check['status'] == 'КОНФЛИКТ_ИНН' for check in checks):
            warnings.append('Сверка DaData выявила конфликт ИНН. Итог реестра сохранён; расхождения требуют ручного решения.')
        errors = sorted({r['status'] for r in api.requests if r['status'] != 'OK'})
        if errors:
            warnings.append('Сверка DaData выполнена не полностью: ' + ', '.join(errors))
    result['warnings'] = warnings
    return result


def write_log(result, destination):
    """Сохранять причины и происхождение решений отдельно от персональных данных."""
    candidates = {}
    for candidate in result['resolution'].get('candidates', []):
        candidates.setdefault(candidate['group_id'], []).append(candidate)
    groups = []
    for group in result['resolution']['groups']:
        selected = group['selected'] or {}
        groups.append({
            'group_id': group['group_id'], 'record_ids': group['record_ids'],
            'status': group['status'], 'selected_inn': selected.get('candidate_inn', ''),
            'source': selected.get('source', ''), 'source_url': selected.get('source_url', ''),
            'score': selected.get('score'), 'registry_used': group['registry_used'],
            'registry_truncated': group.get('registry_truncated', False),
            'candidates': [{k: c.get(k) for k in ('candidate_inn', 'source', 'source_url', 'score', 'eligible', 'reasons')}
                           for c in candidates.get(group['group_id'], [])],
        })
    log = {
        'created_at_utc': result.get('created_at_utc', ''),
        'finished_at_utc': datetime.now(timezone.utc).isoformat(),
        'mode': result.get('processing_mode', 'registry'),
        'registry': result['resolution'].get('registry', {}),
        'sources': result.get('sources', []), 'summary': result.get('full_summary', {}),
        'warnings': result.get('warnings', []), 'groups': groups,
        'selection': [{k: r.get(k) for k in ('record_id', 'source_file', 'source_sheet', 'source_row',
                                            'selection', 'selection_reason', 'flags', 'region_evidence', 'city_evidence')}
                      for r in result['records']],
        'dadata_checks': result.get('dadata_checks', []),
        'dadata_requests': result['resolution']['requests'],
        'dadata_actual_requests': result['resolution']['actual_requests'],
    }
    path = Path(destination) / 'processing_log.json'
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary.replace(path)
    return path
