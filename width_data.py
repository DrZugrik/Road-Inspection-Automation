"""Единое чтение и анализ Excel-данных ширины проезжей части."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd


class WidthDataError(ValueError):
    """Excel-файл ширины не соответствует поддерживаемому формату."""


@dataclass(frozen=True)
class WidthStats:
    min_width: str
    max_width: str
    min_piket: str
    max_piket: str
    min_coordinate: str
    max_coordinate: str


def _find_column(columns, *keywords):
    for column in columns:
        normalized = str(column).lower()
        if any(keyword.lower() in normalized for keyword in keywords):
            return column
    return None


def _format_width(value) -> str:
    return f'{value:.1f}'.replace('.', ',')


def parse_width_dataframe(raw_data: pd.DataFrame) -> tuple[pd.DataFrame, str, str | None]:
    """Преобразует сырой лист ширины в данные, название колонки ширины и координат."""
    header_row_idx = next(
        (
            index for index in range(min(20, len(raw_data)))
            if _find_column(raw_data.iloc[index].fillna('').astype(str).str.strip().tolist(), 'ширина')
        ),
        None,
    )

    if header_row_idx is None:
        if len(raw_data) <= 14:
            raise WidthDataError('Строка заголовков не найдена, а файл слишком короткий для старого формата.')
        headers = raw_data.iloc[12, :4].fillna('').tolist()
        data = raw_data.iloc[14:, :4].copy()
    else:
        headers = raw_data.iloc[header_row_idx, :4].fillna('').tolist()
        data = raw_data.iloc[header_row_idx + 1:, :4].copy()

    data.columns = headers
    data = data.dropna(how='all')
    width_column = _find_column(data.columns, 'ширина')
    if width_column is None:
        raise WidthDataError('Колонка с шириной проезжей части не найдена.')

    data['_width_numeric'] = pd.to_numeric(
        data[width_column].astype(str).str.replace(',', '.', regex=False).str.replace(' ', '', regex=False),
        errors='coerce',
    )
    location_column = _find_column(data.columns, 'местоположение', 'км+м')
    if location_column is None:
        location_column = _find_column(data.columns, '№')

    date_column = _find_column(data.columns, 'дата')
    if date_column is not None:
        data[date_column] = data[date_column].apply(
            lambda value: value.strftime('%d.%m.%Y') if pd.notna(value) and hasattr(value, 'strftime') else str(value)
        )
    return data, width_column, location_column


def load_width_data(file_path: Path | str) -> pd.DataFrame:
    """Загружает ширину для таблицы Word; техническая числовая колонка наружу не передаётся."""
    raw_data = pd.read_excel(file_path, header=None)
    data, width_column, _ = parse_width_dataframe(raw_data)
    result = data.drop(columns=['_width_numeric']).copy()
    result[width_column] = data['_width_numeric'].apply(lambda value: _format_width(value) if pd.notna(value) else '')
    return result


def get_width_stats(file_path: Path | str) -> WidthStats:
    """Возвращает минимальную и максимальную ширину и связанные с ними координаты."""
    raw_data = pd.read_excel(file_path, header=None)
    data, _, location_column = parse_width_dataframe(raw_data)
    valid_data = data.dropna(subset=['_width_numeric'])
    if valid_data.empty:
        raise WidthDataError('В файле отсутствуют корректные числовые значения ширины.')

    min_row = valid_data.loc[valid_data['_width_numeric'].idxmin()]
    max_row = valid_data.loc[valid_data['_width_numeric'].idxmax()]
    min_coordinate = str(min_row[location_column]).strip() if location_column else ''
    max_coordinate = str(max_row[location_column]).strip() if location_column else ''
    return WidthStats(
        min_width=_format_width(min_row['_width_numeric']),
        max_width=_format_width(max_row['_width_numeric']),
        min_piket=min_coordinate or 'ПК0+000',
        max_piket=max_coordinate or 'ПК0+000',
        min_coordinate=min_coordinate,
        max_coordinate=max_coordinate,
    )