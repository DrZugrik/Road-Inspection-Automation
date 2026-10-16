import unittest
from unittest import mock
import tempfile
from pathlib import Path

import pandas as pd

from width_data import WidthDataError, get_width_stats, load_width_data, parse_width_dataframe


class WidthDataTests(unittest.TestCase):
    def test_new_format_parses_decimal_commas_and_selects_min_max(self):
        raw = pd.DataFrame([
            ['Служебная строка', None, None, None],
            ['№', 'Местоположение, км+м', 'Ширина проезжей части, м', 'Дата проведения обследования'],
            [1, 'ПК0+100', '7,25', pd.Timestamp('2026-10-01')],
            [2, 'ПК0+200', '6,5', pd.Timestamp('2026-10-02')],
            [3, 'ПК0+300', '8.0', pd.Timestamp('2026-10-03')],
        ])
        data, width_column, location_column = parse_width_dataframe(raw)
        self.assertEqual(width_column, 'Ширина проезжей части, м')
        self.assertEqual(location_column, 'Местоположение, км+м')
        self.assertEqual(data['_width_numeric'].tolist(), [7.25, 6.5, 8.0])
        self.assertEqual(data.iloc[0, 3], '01.10.2026')

    def test_legacy_format_is_supported(self):
        rows = [[None, None, None, None] for _ in range(15)]
        rows[12] = ['№', 'Местоположение, км+м', 'Ширина проезжей части, м', 'Дата']
        rows[14] = [1, 'ПК1+000', '6,0', '01.10.2026']
        data, _, _ = parse_width_dataframe(pd.DataFrame(rows))
        self.assertEqual(data['_width_numeric'].tolist(), [6.0])

    def test_missing_width_column_is_rejected(self):
        raw = pd.DataFrame([['№', 'Местоположение, км+м'], [1, 'ПК0+100']])
        with self.assertRaises(WidthDataError):
            parse_width_dataframe(raw)

    def test_excel_file_is_loaded_with_display_formatting(self):
        raw = pd.DataFrame([
            ['№', 'Местоположение, км+м', 'Ширина проезжей части, м', 'Дата'],
            [1, 'ПК0+100', '7,25', pd.Timestamp('2026-10-01')],
        ])
        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = Path(temp_dir) / 'Ширина проезжей части.xlsx'
            raw.to_excel(excel_path, index=False, header=False)
            data = load_width_data(excel_path)
        self.assertEqual(data.iloc[0]['Ширина проезжей части, м'], '7,2')
        self.assertEqual(data.iloc[0]['Дата'], '01.10.2026')

    def test_stats_use_coordinates_instead_of_sequence_numbers(self):
        raw = pd.DataFrame([
            ['№', 'Местоположение, км+м', 'Ширина проезжей части, м', None],
            [1, 'ПК2+000', '7,0', None],
            [2, 'ПК2+100', '6,0', None],
        ])
        with mock.patch('width_data.pd.read_excel', return_value=raw):
            stats = get_width_stats('width.xlsx')
        self.assertEqual(stats.min_width, '6,0')
        self.assertEqual(stats.max_width, '7,0')
        self.assertEqual(stats.min_coordinate, 'ПК2+100')
        self.assertEqual(stats.max_coordinate, 'ПК2+000')