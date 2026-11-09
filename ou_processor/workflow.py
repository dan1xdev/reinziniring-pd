"""Общий сценарий UI и CLI: каскад официального реестра и DaData."""

from datetime import datetime, timezone
import json
from pathlib import Path
import sqlite3

from .resolution import LicenseSearch, RegistrySearch, SearchCache, resolve


class _AvailableIndex:
    """Ошибка одного индекса не отменяет уже найденные группы."""
    def __init__(self, index, warnings, name):
        self.index, self.warnings, self.name = index, warnings, name
        self.metadata = index.metadata
        self.failed = False

    def search(self, record):
        if self.failed:
            return [], False
        try:
            return self.index.search(record)
        except sqlite3.Error:
            self.failed = True
            self.warnings.append(f'Не удалось прочитать индекс {self.name}; продолжен поиск в остальных источниках.')
            return [], False


def search(result, registry_path, config, *, licenses_path=None, cache=None, client=None,
           on_progress=None, check_cancel=None):
    warnings, opened = [], []
    registry = licenses = None
    result['processing_mode'] = 'cascade' if client is not None else 'registry'

    def open_index(factory, path, label):
        try:
            index = factory(path, config=config)
            opened.append(index)
            return _AvailableIndex(index, warnings, label)
        except (ValueError, OSError, KeyError, sqlite3.Error):
            warnings.append(f'Локальный реестр недоступен: {label}. Обновите снимок реестра.')
            return None

    try:
        if licenses_path is not None:
            licenses = open_index(LicenseSearch, licenses_path, 'реестра лицензий')
        # Резерв первого семестра доступен явно, но не входит в стандартный каскад ТЗ v2.
        if config.get('allow_accreditation_fallback', False):
            registry = open_index(RegistrySearch, registry_path, 'резервного реестра аккредитации')
        api = SearchCache(cache, client) if client is not None else None
        if api is not None:
            api.max_requests = config.get('dadata_max_requests', 10000)
            if getattr(api, 'warning', ''):
                warnings.append(api.warning)
        resolve(result, api, registry, progress=False, config=config, licenses=licenses,
                on_progress=on_progress, check_cancel=check_cancel)
        if client is None:
            warnings.append('Ключ DaData не настроен: проверка ограничена локальными источниками; адрес и юридический статус не уточнены через DaData.')
        elif any(g['decision_reason'] == 'КОНФЛИКТ_ИНН' for g in result['resolution']['groups']):
            warnings.append('Конфликт ИНН между источниками: автоматический выбор отменён, требуется ручная проверка.')
        errors = sorted({r['status'] for r in result['resolution']['requests'] if r['status'] != 'OK'})
        if errors:
            warnings.append('DaData выполнен не полностью: ' + ', '.join(errors))
    finally:
        for index in opened:
            index.connection.close()
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
            'status': group['status'], 'decision_reason': group.get('decision_reason', ''), 'selected_inn': selected.get('candidate_inn', ''),
            'source': selected.get('source', ''), 'source_url': selected.get('source_url', ''),
            'score': selected.get('score'), 'license_used': group.get('license_used', False),
            'license_truncated': group.get('license_truncated', False),
            'registry_used': group['registry_used'],
            'registry_truncated': group.get('registry_truncated', False),
            'candidates': [{k: c.get(k) for k in ('candidate_inn', 'source', 'source_url', 'score', 'eligible', 'reasons')}
                           for c in candidates.get(group['group_id'], [])],
        })
    log = {
        'created_at_utc': result.get('created_at_utc', ''),
        'finished_at_utc': datetime.now(timezone.utc).isoformat(),
        'mode': result.get('processing_mode', 'registry'),
        'registry': result['resolution'].get('registry', {}),
        'license': result['resolution'].get('license', {}),
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
