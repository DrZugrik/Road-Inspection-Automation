"""Формирование и сохранение сводного манифеста пакетного запуска."""

import json
from datetime import datetime
from pathlib import Path


def build_run_manifest(config, folder_results: list[dict]) -> dict:
    status_counts = {
        'success': sum(result['status'] == 'success' for result in folder_results),
        'completed_with_warnings': sum(result['status'] == 'completed_with_warnings' for result in folder_results),
        'failed': sum(result['status'] == 'failed' for result in folder_results),
    }
    return {
        'created_at': datetime.now().astimezone().isoformat(timespec='seconds'),
        'configuration': {
            'template_path': str(config.template_path),
            'catalog_path': str(config.catalog_path),
            'output_dir': str(config.output_dir),
            'update_word_fields': config.update_word_fields,
        },
        'summary': {'folders_total': len(folder_results), **status_counts},
        'folders': folder_results,
    }


def write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8') as output_file:
        json.dump(data, output_file, ensure_ascii=False, indent=2)