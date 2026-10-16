"""Загрузка параметров запуска генератора из CLI, TOML и переменных окружения."""

from __future__ import annotations

import os
import tomllib
from dataclasses import dataclass
from pathlib import Path


class ConfigurationError(ValueError):
    """Конфигурация не позволяет безопасно начать обработку."""


@dataclass(frozen=True)
class RuntimeConfig:
    template_path: Path
    catalog_path: Path
    output_dir: Path
    folders_to_process: list[Path]
    update_word_fields: bool
    log_level: str


def _read_toml(path: Path | None) -> dict:
    if path is None:
        return {}
    if not path.exists():
        raise ConfigurationError(f'Файл конфигурации не найден: {path}')
    try:
        with path.open('rb') as config_file:
            return tomllib.load(config_file)
    except tomllib.TOMLDecodeError as error:
        raise ConfigurationError(f'Некорректный TOML в {path}: {error}') from error


def _get_value(cli_value, config: dict, key: str, env_name: str, default=None):
    if cli_value is not None:
        return cli_value
    if key in config and config[key] not in (None, ''):
        return config[key]
    if os.getenv(env_name):
        return os.environ[env_name]
    return default


def _as_path(value, name: str) -> Path:
    if value in (None, ''):
        raise ConfigurationError(f'Не задан обязательный параметр: {name}.')
    return Path(value).expanduser()


def load_runtime_config(args) -> RuntimeConfig:
    """Собирает настройки с приоритетом CLI → TOML → окружение."""
    config_path = args.config or (Path(os.environ['ROAD_INSPECTION_CONFIG']) if os.getenv('ROAD_INSPECTION_CONFIG') else None)
    config = _read_toml(config_path)

    template_path = _as_path(_get_value(args.template, config, 'template_path', 'ROAD_INSPECTION_TEMPLATE'), 'template_path')
    catalog_path = _as_path(_get_value(args.catalog, config, 'catalog_path', 'ROAD_INSPECTION_CATALOG'), 'catalog_path')
    output_dir = _as_path(_get_value(args.output_dir, config, 'output_dir', 'ROAD_INSPECTION_OUTPUT_DIR', 'output'), 'output_dir')

    selected_folders = args.folders or []
    if selected_folders:
        folders_to_process = [Path(folder).expanduser() for folder in selected_folders]
    else:
        source_dir = _get_value(args.source, config, 'source_dir', 'ROAD_INSPECTION_SOURCE')
        if source_dir in (None, ''):
            configured_folders = config.get('folders_to_process', [])
            if not configured_folders:
                raise ConfigurationError('Укажите --source, один или несколько --folder или folders_to_process в TOML.')
            folders_to_process = [Path(folder).expanduser() for folder in configured_folders]
        else:
            source_path = Path(source_dir).expanduser()
            process_subfolders = args.process_subfolders or bool(config.get('process_subfolders', False))
            if process_subfolders and source_path.exists():
                folders_to_process = sorted((item for item in source_path.iterdir() if item.is_dir()), key=lambda item: item.name.lower())
            else:
                folders_to_process = [source_path]

    update_word_fields = not args.skip_com_update and bool(config.get('update_word_fields', True))
    log_level = (args.log_level or config.get('log_level') or os.getenv('ROAD_INSPECTION_LOG_LEVEL') or 'INFO').upper()
    return RuntimeConfig(template_path, catalog_path, output_dir, folders_to_process, update_word_fields, log_level)