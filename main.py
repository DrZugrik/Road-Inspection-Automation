"""
main.py - Генератор отчетов из данных Excel для нескольких папок.
Без COM-объектов, через Pandas.

============================================================
ПЛЕЙСХОЛДЕРЫ В ШАБЛОНЕ:
============================================================

МЕТАДАННЫЕ ОБЪЕКТА (из каталога Excel):
------------------------------------------------------------
{{НОМЕР_ПП}}           - номер п/п из каталога
{{ДОРОГА_ИД}}          - идентификационный номер дороги из каталога
{{ДОРОГА_НАЗВАНИЕ}}    - название дороги (основное, берется из Excel если найден,
                         иначе из названия папки)
{{ДОРОГА_НАЗВАНИЕ_ПАПКА}} - название дороги из названия папки
{{ДОРОГА_НАЗВАНИЕ_EXCEL}} - название дороги из Excel (как указано в каталоге)
{{БАЛАНСОДЕРЖАТЕЛЬ}}   - балансодержатель дороги (если пусто - прочерк "-")

ТАБЛИЦЫ ДАННЫХ (из Excel-файлов в папке объекта):
------------------------------------------------------------
{{ТАБЛИЦА_ПРОДОЛЬНАЯ_ДАННЫЕ}}  - данные продольной ровности
{{ТАБЛИЦА_ПОПЕРЕЧНАЯ_ДАННЫЕ}}  - данные поперечной ровности
{{ТАБЛИЦА_ДЕФЕКТЫ_ДАННЫЕ}}     - данные дефектов покрытия
{{ТАБЛИЦА_ПРОЧНОСТЬ_ДАННЫЕ}}   - данные прочности дорожной одежды
{{ТАБЛИЦА_ШИРИНА_ДАННЫЕ}}      - данные ширины проезжей части
{{ТАБЛИЦА_ШИРИНА_СХЕМА}}       - схема (картинка) из файла ширины

ФОТОГРАФИИ ОБЩИХ ВИДОВ (из папки "Фотоматериалы <название>"):
------------------------------------------------------------
{{ФОТО_ВИД_ДОРОГА_1}} ... {{ФОТО_ВИД_ДОРОГА_7}}  - фотографии общих видов
{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_1}} ... {{ФОТО_ВИД_ДОРОГА_ПИКЕТ_7}} - подписи пикетов
"""

import os
import shutil
import json
import zipfile
import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pandas as pd
import numpy as np
from PIL import Image
import photos_processor

# Конфигурация
TEMPLATE_PATH = r"D:\ProgrammEngine\road_inspection_automation\template\template_6.docx"
CATALOG_PATH = r"D:\ProgrammEngine\road_inspection_automation\template\Восстановление АД (форма ЦУКС)_2 этап_12.08.2026.xlsx"
#CATALOG_PATH = r"D:\ProgrammEngine\road_inspection_automation\template\Перечень дорог 2 этап.xlsx"
BASE_DIR = Path(r"D:\ProgrammEngine\road_inspection_automation")
OUTPUT_DIR = BASE_DIR / "output"

# Конфигурация источника данных
SOURCE_DIR = r"C:\Users\user\Downloads\4 этап - доделки 01,10,2026"

# Список папок для обработки
FOLDERS_TO_PROCESS = [
    SOURCE_DIR,
]

# Маппинг плейсхолдеров таблиц
PLACEHOLDER_MAPPING = {
    '{{ТАБЛИЦА_ПРОДОЛЬНАЯ_ДАННЫЕ}}': {
        'file_pattern': 'Продольная ровность покрытия',
        'skip_rows': 4,
        'excel_mapping': {
            0: 'Начало участка, км+м',
            1: 'Конец участка, км+м',
            2: 'Протяженность, м',
            3: 'Прямое направление_1',
            11: 'Обратное направление_2',
            19: 'Наихудшее значение',
            20: 'Норматив',
            21: 'Соответствие нормативу'
        }
    },
    '{{ТАБЛИЦА_ПОПЕРЕЧНАЯ_ДАННЫЕ}}': {
        'file_pattern': 'Поперечная ровность покрытия',
        'skip_rows': 4,
        'excel_mapping': {
            0: 'Начало участка, км+м',
            1: 'Конец участка, км+м',
            2: 'Протяженность, м',
            3: 'Прямое направление_1',
            11: 'Обратное направление_2',
            19: 'Наихудшее значение',
            20: 'Норматив',
            21: 'Соответствие нормативу'
        }
    },
    '{{ТАБЛИЦА_ДЕФЕКТЫ_ДАННЫЕ}}': {
        'file_pattern': 'Дефекты покрытия ОДМ 218',
        'skip_rows': 3,
        'excel_mapping': {
            0: 'Начало участка, км+м',
            1: 'Конец участка, км+м',
            2: 'Протяженность, м',
            3: 'Прямое направление_коды',
            4: 'Обратное направление_коды',
            5: 'Прямое направление_баллы',
            6: 'Обратное направление_баллы'
        }
    },
    '{{ТАБЛИЦА_ПРОЧНОСТЬ_ДАННЫЕ}}': {
        'file_pattern': 'Прочность дорожной одежды',
        'skip_rows': 3,
        'excel_mapping': {
            0: 'Начало участка, км+м',
            1: 'Конец участка, км+м',
            2: 'Протяженность, м',
            3: 'Прогиб_прямое',
            4: 'Прогиб_обратное',
            5: 'Модуль_упругости_прямое',
            6: 'Модуль_упругости_обратное',
            7: 'Коэф_прочности_прямое',
            8: 'Коэф_прочности_обратное',
            9: 'Коэф_прочности_требуемый',
            10: 'Соответствие нормативу',
            11: 'Дата проведения'
        },
        'font_size': 9
    },
    '{{ТАБЛИЦА_ШИРИНА_СХЕМА}}': {
        'file_pattern': 'Ширина проезжей части'
    },
    '{{ТАБЛИЦА_ШИРИНА_ДАННЫЕ}}': {
        'file_pattern': 'Ширина проезжей части',
        'font_size': 9
    }
}

# ==============================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ РАБОТЫ С ШРИФТАМИ
# ==============================================
def set_font_for_paragraph(paragraph, font_name='Times New Roman', font_size=10, bold=False):
    """Устанавливает шрифт для всех прогонов в параграфе"""
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        # Устанавливаем кодировку для корректного отображения кириллицы
        run.font.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def set_font_for_cell(cell, font_name='Times New Roman', font_size=10, bold=False):
    """Устанавливает шрифт для всех параграфов в ячейке"""
    for paragraph in cell.paragraphs:
        set_font_for_paragraph(paragraph, font_name, font_size, bold)

def set_font_for_table(table, font_name='Times New Roman', font_size=10):
    """Устанавливает шрифт для всей таблицы"""
    for row in table.rows:
        for cell in row.cells:
            set_font_for_cell(cell, font_name, font_size)

# ==============================================
# ФУНКЦИЯ ДЛЯ ЗАМЕНЫ ТЕКСТА С СОХРАНЕНИЕМ СТИЛЕЙ
# ==============================================
def replace_text_preserve_style(paragraph, old_text, new_text):
    """
    Заменяет текст в параграфе, сохраняя стили и форматирование.
    Работает с существующими прогонами (runs), не сбрасывая их свойства.
    """
    if old_text not in paragraph.text:
        return False
    
    # Получаем все прогоны
    runs = paragraph.runs
    if not runs:
        # Если нет прогонов, создаем новый
        paragraph.text = paragraph.text.replace(old_text, new_text)
        return True
    
    # Собираем весь текст параграфа
    full_text = ''.join(run.text for run in runs)
    
    # Если old_text не найден в полном тексте, пробуем заменить напрямую
    if old_text not in full_text:
        paragraph.text = paragraph.text.replace(old_text, new_text)
        return True
    
    # Находим позицию old_text в полном тексте
    old_pos = full_text.find(old_text)
    old_len = len(old_text)
    
    # Пересобираем текст с заменой
    current_pos = 0
    for run in runs:
        run_text = run.text
        run_start = current_pos
        run_end = current_pos + len(run_text)
        
        # Проверяем, попадает ли замена в этот прогон
        if old_pos >= run_start and old_pos < run_end:
            # Замена в этом прогоне
            before = run_text[:old_pos - run_start]
            after = run_text[old_pos - run_start + old_len:]
            run.text = before + new_text + after
        elif old_pos < run_start and old_pos + old_len > run_start:
            # Замена начинается до прогона и заканчивается внутри
            overlap_start = run_start - old_pos
            overlap_len = old_len - overlap_start
            if overlap_len > 0:
                after = run_text[overlap_len:]
                run.text = after
        elif old_pos < run_end and old_pos + old_len > run_end:
            # Замена начинается внутри прогона и заканчивается после
            before = run_text[:old_pos - run_start]
            run.text = before
        elif old_pos >= run_start and old_pos + old_len <= run_end:
            # Полная замена внутри прогона - уже обработано выше
            pass
        
        current_pos += len(run_text)
    
    return True

def replace_text_placeholders_preserve_style(doc, metadata):
    """
    Заменяет текстовые плейсхолдеры в документе с сохранением форматирования
    """
    replaced_count = 0
    
    # Обрабатываем параграфы
    for paragraph in doc.paragraphs:
        for key, value in metadata.items():
            if key in paragraph.text:
                if replace_text_preserve_style(paragraph, key, str(value)):
                    replaced_count += 1
    
    # Обрабатываем таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in metadata.items():
                    if key in cell.text:
                        # В ячейках проще - заменяем текст напрямую
                        cell.text = cell.text.replace(key, str(value))
                        replaced_count += 1
    
    print(f"   ✅ Заменено {replaced_count} текстовых плейсхолдеров (с сохранением стилей)")
    return replaced_count

# ==============================================
# ФУНКЦИИ ДЛЯ НЕСТРОГОГО ПОИСКА СОВПАДЕНИЙ
# ==============================================
def normalize_string(text):
    """
    Нормализует строку для сравнения:
    - приводит к нижнему регистру
    - удаляет лишние пробелы
    - заменяет разные варианты написания на единый формат
    """
    if not isinstance(text, str):
        return ""
    
    # Приводим к нижнему регистру
    text = text.lower().strip()
    
    # Заменяем различные варианты написания
    replacements = {
        'а/д': 'ад',
        'а_д': 'ад',
        'ж/д': 'жд',
        'ж_д': 'жд',
        'ул': 'улица',
        'ул.': 'улица',
        'д.': 'дом',
        'ш.': 'шоссе',
        'ш': 'шоссе',
        'пр-д': 'проезд',
        'пр': 'проезд',
        'проезд': 'пр',
        'шоссе': 'ш',
    }
    
    for old, new in replacements.items():
        text = text.replace(old, new)
    
    # Удаляем все знаки препинания и лишние символы
    text = re.sub(r'[^\w\s]', ' ', text)
    
    # Удаляем множественные пробелы
    text = re.sub(r'\s+', ' ', text)
    
    return text.strip()

def find_matching_road(folder_name, df, col_name='Наименование дороги'):
    """
    Находит соответствие между названием папки и названием дороги в Excel
    с использованием нестрогого сравнения
    """
    folder_normalized = normalize_string(folder_name)
    
    best_match = None
    best_score = 0
    
    for idx, row in df.iterrows():
        excel_name = row[col_name]
        if pd.isna(excel_name):
            continue
        
        excel_name_str = str(excel_name)
        excel_normalized = normalize_string(excel_name_str)
        
        # Проверяем точное совпадение после нормализации
        if folder_normalized == excel_normalized:
            return idx, excel_name_str  # Точное совпадение!
        
        # Проверяем, содержит ли одно название другое
        if folder_normalized in excel_normalized or excel_normalized in folder_normalized:
            # Чем длиннее совпадающая часть, тем лучше
            match_len = min(len(folder_normalized), len(excel_normalized))
            if match_len > best_score:
                best_score = match_len
                best_match = (idx, excel_name_str)
        
        # Проверяем по словам (пересечение множеств слов)
        folder_words = set(folder_normalized.split())
        excel_words = set(excel_normalized.split())
        
        if folder_words and excel_words:
            common_words = folder_words.intersection(excel_words)
            # Вычисляем коэффициент Жаккара (чем ближе к 1, тем лучше)
            jaccard = len(common_words) / len(folder_words.union(excel_words))
            
            if jaccard > 0.5 and jaccard > best_score:
                best_score = jaccard
                best_match = (idx, excel_name_str)
    
    return best_match if best_match else (None, None)
# ==============================================

def find_excel_file(directory, pattern):
    for file in directory.glob("*.xlsx"):
        if pattern in file.name:
            return file
    return None

def find_placeholder_in_document(doc, placeholder):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            return {'type': 'paragraph', 'element': paragraph}
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    return {'type': 'table', 'element': table, 'cell': cell, 'row': row}
    return None

def set_cell_border(cell, border_color='000000', border_size=1):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), str(border_size * 4))
        border.set(qn('w:color'), border_color)
        borders.append(border)
    tcPr.append(borders)

def set_table_borders(table, border_color='000000', border_size=1):
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, border_color, border_size)

def merge_cells(table, row_idx, col_start, col_end):
    if col_start < col_end:
        table.rows[row_idx].cells[col_start].merge(table.rows[row_idx].cells[col_end])

def format_cell_text(cell, bold=False, align='center', font_name='Times New Roman', font_size=10):
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == 'center' else WD_ALIGN_PARAGRAPH.LEFT
        set_font_for_paragraph(paragraph, font_name, font_size, bold)

# ==========================================================
# ЗАМЕНА: Pandas вместо win32com
# ==========================================================
def load_excel_with_pandas(file_path, skip_rows, mapping):
    try:
        df = pd.read_excel(file_path, header=None, skiprows=skip_rows)
        df = df.iloc[:, list(mapping.keys())].copy()
        df.columns = list(mapping.values())
        
        # Удаляем пустые строки
        df = df.dropna(how='all')
        
        # Конвертируем числа в строки с запятой (как в Excel)
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                # ====================================================
                # ФОРМАТИРОВАНИЕ ДЛЯ ТАБЛИЦЫ ПРОЧНОСТИ
                # ====================================================
                if any(keyword in col for keyword in ['Модуль_упругости', 'Коэф_прочности', 'Прогиб']):
                    df[col] = df[col].apply(lambda x: 
                        f"{x:.1f}".replace('.', ',') if x >= 100 else 
                        f"{x:.2f}".replace('.', ',')
                    )
                # ====================================================
                # ФОРМАТИРОВАНИЕ ДЛЯ ТАБЛИЦЫ ДЕФЕКТОВ (баллы)
                # ====================================================
                elif any(keyword in col for keyword in ['_баллы']):
                    df[col] = df[col].apply(lambda x: 
                        f"{x:.1f}".replace('.', ',')  # Всегда 1 знак после запятой
                    )
                # ====================================================
                # УНИВЕРСАЛЬНОЕ ФОРМАТИРОВАНИЕ ДЛЯ ВСЕХ ОСТАЛЬНЫХ ЧИСЕЛ
                # (Продольная ровность, Поперечная ровность, метры, нормативы, глубины)
                # ====================================================
                else:
                    df[col] = df[col].apply(lambda x: 
                        f"{x:.1f}".replace('.', ',') if pd.notna(x) else ""
                    )

        # Обработка даты (если есть)
        if 'Дата проведения' in df.columns:
            df['Дата проведения'] = df['Дата проведения'].apply(
                lambda x: x.strftime('%d.%m.%Y') if pd.notna(x) and hasattr(x, 'strftime') else str(x)
            )

        return df
    except Exception as e:
        print(f"   ❌ Ошибка Pandas: {e}")
        return pd.DataFrame()

def load_width_data_with_pandas(file_path):
    """
    Загружает данные ширины проезжей части из Excel-файла.
    Поддерживает два формата:
    1. С заголовками на строке 13 (индекс 12) и данными с 15 (индекс 14)
    2. С заголовками "№", "Местоположение, км+м", "Ширина проезжей части, м", "Дата проведения обследования"
    """
    try:
        # Читаем весь файл без заголовков
        df_raw = pd.read_excel(file_path, header=None)
        
        # Пытаемся найти строку с заголовками
        header_row_idx = None
        possible_headers = ['№', 'Местоположение, км+м', 'Ширина проезжей части, м', 'Дата проведения обследования']
        
        for i in range(min(20, len(df_raw))):
            row_values = df_raw.iloc[i].astype(str).str.strip().tolist()
            # Проверяем, содержит ли строка основные заголовки
            if any('Местоположение' in str(v) or 'Ширина' in str(v) for v in row_values):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            # Если не нашли - используем старый формат (заголовки на строке 13)
            print("   ℹ️ Использую старый формат (заголовки на строке 13)")
            headers = df_raw.iloc[12, :4].fillna('').tolist()
            data = df_raw.iloc[14:, :4].copy()
            data.columns = headers
        else:
            # Используем найденные заголовки
            print(f"   ℹ️ Найдены заголовки на строке {header_row_idx + 1}")
            headers = df_raw.iloc[header_row_idx, :4].fillna('').tolist()
            # Данные начинаются со следующей строки после заголовка
            data_start = header_row_idx + 1
            data = df_raw.iloc[data_start:, :4].copy()
            data.columns = headers
        
        # Удаляем пустые строки
        data = data.dropna(how='all')
        
        # Ищем колонку с шириной
        width_col = None
        for col in data.columns:
            if 'Ширина' in str(col):
                width_col = col
                break
        
        if width_col is None:
            print("   ⚠️ Колонка с шириной не найдена")
            return pd.DataFrame()
        
        # Преобразуем ширину в числа (запятая -> точка)
        data[width_col] = data[width_col].astype(str).str.replace(',', '.').str.replace(' ', '')
        data[width_col] = pd.to_numeric(data[width_col], errors='coerce')
        
        # Форматируем числа обратно с запятой для отображения
        data[width_col] = data[width_col].apply(
            lambda x: f"{x:.1f}".replace('.', ',') if pd.notna(x) else ""
        )
        
        # Ищем колонку с местоположением/пикетом
        loc_col = None
        for col in data.columns:
            if 'Местоположение' in str(col) or 'км+м' in str(col):
                loc_col = col
                break
        
        if loc_col is None:
            # Пробуем найти колонку с номером и пикетом
            for col in data.columns:
                if '№' in str(col):
                    loc_col = col
                    break
        
        # Обработка даты
        date_col = None
        for col in data.columns:
            if 'Дата' in str(col) or 'дата' in str(col).lower():
                date_col = col
                break
        
        if date_col is not None:
            data[date_col] = data[date_col].apply(
                lambda x: x.strftime('%d.%m.%Y') if pd.notna(x) and hasattr(x, 'strftime') else str(x)
            )
        
        # Переименовываем колонки для единообразия
        new_columns = []
        for col in data.columns:
            col_str = str(col).strip()
            if '№' in col_str:
                new_columns.append('№')
            elif 'Местоположение' in col_str or 'км+м' in col_str:
                new_columns.append('Местоположение, км+м')
            elif 'Ширина' in col_str:
                new_columns.append('Ширина проезжей части, м')
            elif 'Дата' in col_str:
                new_columns.append('Дата проведения обследования')
            else:
                new_columns.append(col_str)
        
        data.columns = new_columns
        
        return data
        
    except Exception as e:
        print(f"   ❌ Ошибка загрузки ширины: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


def get_width_stats(object_folder_path):
    """
    Извлекает минимальную и максимальную ширину проезжей части
    с соответствующими пикетами из Excel-файла.
    Возвращает словарь с данными.
    """
    try:
        # Ищем файл с шириной
        excel_file = None
        for file in Path(object_folder_path).glob("*.xlsx"):
            if "Ширина" in file.name:
                excel_file = file
                break
        
        if not excel_file:
            print("   ⚠️ Excel-файл с 'Ширина' не найден")
            return None
        
        # Загружаем данные
        data = load_width_data_with_pandas(excel_file)
        if data.empty:
            print("   ⚠️ Нет данных в файле ширины")
            return None
        
        # Ищем колонку с шириной
        width_col = None
        for col in data.columns:
            if 'Ширина' in col:
                width_col = col
                break
        
        if width_col is None:
            print("   ⚠️ Колонка с шириной не найдена")
            return None
        
        # Преобразуем в числа для вычислений
        data['_width_num'] = data[width_col].astype(str).str.replace(',', '.').str.replace(' ', '')
        data['_width_num'] = pd.to_numeric(data['_width_num'], errors='coerce')
        
        # Удаляем строки с NaN
        data_clean = data.dropna(subset=['_width_num'])
        
        if data_clean.empty:
            print("   ⚠️ Нет валидных данных о ширине")
            return None
        
        # Находим min и max
        min_row = data_clean.loc[data_clean['_width_num'].idxmin()]
        max_row = data_clean.loc[data_clean['_width_num'].idxmax()]
        
        min_width = min_row['_width_num']
        max_width = max_row['_width_num']
        
        # Ищем колонку с местоположением
        loc_col = None
        for col in data_clean.columns:
            if 'Местоположение' in col or 'км+м' in col:
                loc_col = col
                break
        
        if loc_col is None:
            # Пробуем найти по номеру
            for col in data_clean.columns:
                if '№' in col:
                    loc_col = col
                    break
        
        # Получаем пикеты
        min_piket = str(min_row[loc_col]).strip() if loc_col else "ПК0+000"
        max_piket = str(max_row[loc_col]).strip() if loc_col else "ПК0+000"
        
        # Приводим пикеты к формату ПКx+xxx
        if '+' not in min_piket:
            # Если нет плюса - возможно это просто номер
            if loc_col and '№' in loc_col:
                # Ищем другую колонку с местоположением
                for col in data_clean.columns:
                    if 'км+м' in col:
                        min_piket = str(min_row[col]).strip()
                        max_piket = str(max_row[col]).strip()
                        break
        
        # Форматируем ширину с запятой
        min_width_str = f"{min_width:.1f}".replace('.', ',')
        max_width_str = f"{max_width:.1f}".replace('.', ',')
        
        result = {
            'min_width': min_width_str,
            'max_width': max_width_str,
            'min_piket': min_piket,
            'max_piket': max_piket,
            'min_row': min_row,
            'max_row': max_row
        }
        
        print(f"   📏 Ширина: MIN={min_width_str} м (пикет {min_piket}), MAX={max_width_str} м (пикет {max_piket})")
        return result
        
    except Exception as e:
        print(f"   ❌ Ошибка при получении ширины: {e}")
        import traceback
        traceback.print_exc()
        return None

# ==========================================================

def extract_and_insert_image_from_excel(doc, excel_file_path, placeholder):
    found = find_placeholder_in_document(doc, placeholder)
    if not found:
        print(f"   ⚠️ Плейсхолдер {placeholder} не найден")
        return False

    temp_dir = Path(BASE_DIR / "temp_extracted_images")
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(excel_file_path, 'r') as zf:
            image_files = [f for f in zf.namelist() if f.startswith('xl/media/')]
            if not image_files:
                print("   ⚠️ В файле Excel не найдено встроенных картинок.")
                return False

            saved_paths = []
            for file in image_files:
                target_path = temp_dir / Path(file).name
                with zf.open(file) as source, open(target_path, 'wb') as target:
                    shutil.copyfileobj(source, target)
                saved_paths.append(target_path)

        best_img_path = None
        best_img_size = 0
        for img_path in saved_paths:
            try:
                with Image.open(img_path) as pil_img:
                    w, h = pil_img.size
                    if w * h > best_img_size:
                        best_img_size = w * h
                        best_img_path = img_path
            except:
                continue

        if best_img_path is None:
            shutil.rmtree(temp_dir)
            return False

        if found['type'] == 'paragraph':
            paragraph = found['element']
            full_text = paragraph.text
            if placeholder in full_text:
                parts = full_text.split(placeholder)
                prefix = parts[0]
                suffix = parts[1] if len(parts) > 1 else ""

                paragraph.clear()
                if prefix:
                    run = paragraph.add_run(prefix)
                    set_font_for_paragraph(paragraph, 'Times New Roman', 10)

                run = paragraph.add_run()
                run.add_picture(str(best_img_path), width=Cm(14), height=Cm(12))

                if suffix:
                    suffix_para = doc.add_paragraph(suffix)
                    suffix_para.alignment = paragraph.alignment
                    set_font_for_paragraph(suffix_para, 'Times New Roman', 10)

        elif found['type'] == 'table':
            cell = found['cell']
            cell.text = ""
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(str(best_img_path), width=Cm(14), height=Cm(12))
            set_font_for_cell(cell, 'Times New Roman', 10)

        print(f"   ✅ Картинка успешно вставлена вместо {placeholder}")
        shutil.rmtree(temp_dir)
        return True

    except Exception as e:
        print(f"   ❌ Ошибка при извлечении картинки: {e}")
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        return False

def fill_table_data_keep_structure(table, data_df, data_columns, placeholder, config):
    placeholder_row_idx = None
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            if placeholder in cell.text:
                placeholder_row_idx = row_idx
                break
        if placeholder_row_idx is not None:
            break
    if placeholder_row_idx is None: return False
    while len(table.rows) > placeholder_row_idx:
        table._element.remove(table.rows[-1]._element)
    if len(table.rows) == 0:
        header_row = table.add_row()
        for idx, col_name in enumerate(data_columns):
            if idx < len(header_row.cells):
                header_row.cells[idx].text = col_name
    font_size = config.get('font_size', 10)
    for _, row in data_df.iterrows():
        new_row = table.add_row()
        for col_idx, col_name in enumerate(data_columns):
            if col_idx < len(new_row.cells):
                value = row[col_name]
                new_row.cells[col_idx].text = str(value) if value is not None else ""
    for col in table.columns:
        col.width = Inches(1.5)
    set_table_borders(table, '000000', 1)
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_font_for_cell(cell, 'Times New Roman', font_size)
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return True

def configure_table_style(table):
    if len(table.rows) > 0:
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), 'true')
        trPr.append(tblHeader)

def process_table(doc, data_df, placeholder, config):
    print(f"\n   📊 Обработка таблицы: {placeholder}")
    found = find_placeholder_in_document(doc, placeholder)
    if not found:
        print(f"   ⚠️ Плейсхолдер {placeholder} не найден")
        return False
    col_names = list(data_df.columns)
    font_size = config.get('font_size', 10)
    if found['type'] == 'paragraph':
        paragraph = found['element']
        full_text = paragraph.text
        if placeholder in full_text:
            parts = full_text.split(placeholder)
            prefix = parts[0]
            suffix = parts[1] if len(parts) > 1 else ""

            paragraph.clear()

            if prefix:
                run = paragraph.add_run(prefix)
                run.font.size = Pt(10)
                run.font.name = 'Times New Roman'

            table = doc.add_table(rows=1, cols=len(col_names))
            table.style = 'Table Grid'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            if placeholder == '{{ТАБЛИЦА_ШИРИНА_ДАННЫЕ}}':
                table.columns[0].width = Inches(0.8)
                table.columns[1].width = Inches(2.5)
                table.columns[2].width = Inches(1.8)
                table.columns[3].width = Inches(1.8)

                new_row = table.add_row()
                row_idx = 0
                table.rows[row_idx].cells[0].text = "Ведомость ширины проезжей части"
                merge_cells(table, row_idx, 0, len(col_names) - 1)
                format_cell_text(table.rows[row_idx].cells[0], bold=True, align='center', font_name='Times New Roman', font_size=font_size)

                header_row = table.add_row()
                for idx, col_name in enumerate(col_names):
                    header_row.cells[idx].text = col_name
                    format_cell_text(header_row.cells[idx], bold=True, align='center', font_name='Times New Roman', font_size=font_size)

                for _, row in data_df.iterrows():
                    new_row = table.add_row()
                    for col_idx, col_name in enumerate(col_names):
                        value = row[col_name]
                        new_row.cells[col_idx].text = str(value) if value is not None else ""
                        if col_name != "Географическая координата":
                            format_cell_text(new_row.cells[col_idx], bold=False, align='center', font_name='Times New Roman', font_size=font_size)
                        else:
                            format_cell_text(new_row.cells[col_idx], bold=False, align='left', font_name='Times New Roman', font_size=font_size)
            else:
                header_row = table.rows[0]
                for idx, col_name in enumerate(col_names):
                    if idx < len(header_row.cells):
                        header_row.cells[idx].text = col_name
                        format_cell_text(header_row.cells[idx], bold=True, align='center', font_name='Times New Roman', font_size=font_size)
                for _, row in data_df.iterrows():
                    new_row = table.add_row()
                    for col_idx, col_name in enumerate(col_names):
                        if col_idx < len(new_row.cells):
                            value = row[col_name]
                            new_row.cells[col_idx].text = str(value) if value is not None else ""
                            format_cell_text(new_row.cells[col_idx], bold=False, align='center', font_name='Times New Roman', font_size=font_size)
                for col in table.columns:
                    col.width = Inches(1.5)

            set_table_borders(table, '000000', 1)
            for row in table.rows:
                for cell in row.cells:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    set_font_for_cell(cell, 'Times New Roman', font_size)

            if suffix:
                suffix_para = doc.add_paragraph(suffix)
                suffix_para.alignment = paragraph.alignment
                set_font_for_paragraph(suffix_para, 'Times New Roman', 10)

            print(f"   ✅ Таблица создана ({len(data_df)} строк, шрифт {font_size})")
        
    else:
        table = found['element']
        success = fill_table_data_keep_structure(table, data_df, col_names, placeholder, config)
        if success:
            print(f"   ✅ Таблица заполнена ({len(data_df)} строк, шрифт {font_size})")
        else:
            print(f"   ❌ Ошибка при заполнении таблицы")
            return False
    
    configure_table_style(table)
    return True

# =========================================================================
# ЗАГРУЗКА МЕТАДАННЫХ ИЗ КАТАЛОГА
# =========================================================================
def clean_value(value):
    if pd.isna(value):
        return "-"  # Возвращаем прочерк вместо "(не указано)"
    if isinstance(value, (int, float, np.integer, np.floating)):
        if float(value).is_integer():
            return int(value)
        else:
            return value
    return str(value)

def load_object_metadata(catalog_path, target_dir_name):
    print(f"\n🔍 Поиск объекта по папке: {target_dir_name}")
    if not Path(catalog_path).exists():
        print(f"❌ Файл каталога не найден: {catalog_path}")
        return {}

    try:
        df = pd.read_excel(catalog_path)
        
        # ============================================================
        # ОПРЕДЕЛЯЕМ НАЗВАНИЯ КОЛОНОК С ГИБКИМ ПОИСКОМ
        # ============================================================
        COL_NUM = '№ п/п'
        COL_SECTION = 'Участок'
        COL_NAME = 'Наименование дороги'
        COL_BALANCE = 'Балансодержатель автомобильной дороги'
        COL_ID = 'Идентификационный номер'
        
        # ГИБКИЙ ПОИСК для длины - пробуем разные варианты
        length_col = None
        possible_length_names = [
            'Протяженость(Линейная)',
            'Протяженость',
            'Протяженность',
            'Длина',
            'Протяженность, м',
            'Протяженность (м)'
        ]
        for possible in possible_length_names:
            if possible in df.columns:
                length_col = possible
                break
        
        if length_col is None:
            print(f"⚠️ Столбец с протяженностью не найден. Будет использован '-'")
        else:
            print(f"   ✅ Найден столбец длины: '{length_col}'")

        # Проверяем обязательные столбцы
        missing_cols = []
        if COL_NUM not in df.columns: missing_cols.append(COL_NUM)
        if COL_NAME not in df.columns: missing_cols.append(COL_NAME)
        if COL_BALANCE not in df.columns: missing_cols.append(COL_BALANCE)
        if COL_ID not in df.columns: missing_cols.append(COL_ID)
        
        # Для номера участка - гибкий поиск
        if COL_SECTION not in df.columns:
            possible_names = ['Участок', 'Номер участка', '№ участка', 'Уч-к', 'Секция']
            found = False
            for possible in possible_names:
                if possible in df.columns:
                    COL_SECTION = possible
                    found = True
                    break
            if not found:
                print(f"⚠️ Столбец с номером участка не найден. Будет использован '-'")
                COL_SECTION = None

        if missing_cols:
            print(f"❌ В файле отсутствуют колонки: {missing_cols}")
            return {}

        # Используем нестрогий поиск
        idx, excel_road_name = find_matching_road(target_dir_name, df, COL_NAME)
        
        if idx is None:
            print(f"❌ Объект '{target_dir_name}' не найден в каталоге.")
            return {}
        
        match_row = df.iloc[idx]
        
        # Получаем название дороги из Excel
        road_name_excel = excel_road_name
        road_name_main = road_name_excel if road_name_excel else target_dir_name
        
        # Номер участка
        section_number = "-"
        if COL_SECTION is not None and COL_SECTION in df.columns:
            raw_value = match_row[COL_SECTION]
            if pd.notna(raw_value):
                section_number = clean_value(raw_value)
        
        # ============================================================
        # ДЛИНА ДОРОГИ (с конвертацией в км и 3 знаками после запятой)
        # ============================================================
        road_length = "-"
        if length_col is not None and length_col in df.columns:
            raw_value = match_row[length_col]
            if pd.notna(raw_value):
                try:
                    val = float(raw_value)
                    # Если значение > 1000 - скорее всего в метрах, конвертируем в км
                    if val > 1000:
                        road_length = f"{val/1000:.3f}".replace('.', ',')
                    else:
                        road_length = f"{val:.3f}".replace('.', ',')
                except:
                    road_length = str(raw_value)
        # ============================================================
        # ============================================================
        
        metadata = {
            '{{НОМЕР_ПП}}': clean_value(match_row[COL_NUM]),
            '{{НОМЕР_УЧСТОК}}': section_number,
            '{{ДОРОГА_НАЗВАНИЕ}}': road_name_main,
            '{{ДОРОГА_НАЗВАНИЕ_ПАПКА}}': target_dir_name,
            '{{ДОРОГА_НАЗВАНИЕ_EXCEL}}': road_name_excel,
            '{{БАЛАНСОДЕРЖАТЕЛЬ}}': clean_value(match_row[COL_BALANCE]),
            '{{ДОРОГА_ИД}}': clean_value(match_row[COL_ID]),
            '{{ДОРОГА_ДЛИНА}}': road_length,  # <-- Исправлено
            '{{ДОРОГА_ШИРИНА_MIN}}': '-',  # Будет заполнено позже
            '{{ДОРОГА_ШИРИНА_MAX}}': '-',  # Будет заполнено позже
        }
        
        print("✅ Найдены метаданные для объекта:")
        for k, v in metadata.items():
            print(f"   {k} -> {v}")
        return metadata

    except Exception as e:
        print(f"❌ Ошибка при чтении каталога: {e}")
        import traceback
        traceback.print_exc()
        return {}

# =========================================================================

def process_all_tables(doc, data_dir):
    results = []
    for placeholder, config in PLACEHOLDER_MAPPING.items():
        excel_file = find_excel_file(data_dir, config['file_pattern'])
        if not excel_file:
            print(f"   ⚠️ Excel-файл не найден для {placeholder}")
            results.append({'placeholder': placeholder, 'status': 'not_found', 'rows': 0})
            continue
        try:
            print(f"   📖 Чтение данных из: {excel_file.name}")
            if placeholder == '{{ТАБЛИЦА_ШИРИНА_СХЕМА}}':
                success = extract_and_insert_image_from_excel(doc, excel_file, placeholder)
                if success:
                    results.append({'placeholder': placeholder, 'status': 'success', 'rows': 'Image', 'file': excel_file.name})
                else:
                    results.append({'placeholder': placeholder, 'status': 'failed', 'rows': 0, 'file': excel_file.name})
                continue

            if placeholder == '{{ТАБЛИЦА_ШИРИНА_ДАННЫЕ}}':
                df_filtered = load_width_data_with_pandas(excel_file)
            else:
                skip_rows = config.get('skip_rows', 4)
                mapping = config['excel_mapping']
                df_filtered = load_excel_with_pandas(excel_file, skip_rows, mapping)
            
            if 'Соответствие нормативу' in df_filtered.columns:
                df_filtered['Соответствие нормативу'] = df_filtered['Соответствие нормативу'].astype(str).str.strip()
                df_filtered['Соответствие нормативу'] = df_filtered['Соответствие нормативу'].replace({
                    'да': 'Да', 'нет': 'Нет', 'nan': '', 'NaN': '', '': ''
                })
            if placeholder == '{{ТАБЛИЦА_ПРОЧНОСТЬ_ДАННЫЕ}}':
                df_filtered['Температура, °C'] = '+24'
            
            print(f"   📊 Обработано {len(df_filtered)} строк данных")
            if len(df_filtered) > 0:
                print(f"   📋 Первые строки данных:")
                print(df_filtered.head(3).to_string())
            
            success = process_table(doc, df_filtered, placeholder, config)
            results.append({
                'placeholder': placeholder,
                'status': 'success' if success else 'failed',
                'rows': len(df_filtered),
                'file': excel_file.name
            })
        except Exception as e:
            print(f"   ❌ Ошибка при обработке {placeholder}: {e}")
            import traceback
            traceback.print_exc()
            results.append({'placeholder': placeholder, 'status': 'error', 'error': str(e)})
    return results

def generate_report_for_single_folder(folder_path, template_path):
    folder_name = folder_path.name
    print(f"\n📁 Обработка папки: {folder_name}")
    print("-" * 60)
    
    metadata = load_object_metadata(CATALOG_PATH, folder_name)

    # ============================================================
    # ДОПОЛНЯЕМ МЕТАДАННЫЕ ШИРИНОЙ И ПИКЕТАМИ
    # ============================================================
    if metadata:
        width_stats = get_width_stats(folder_path)
        if width_stats:
            metadata['{{ДОРОГА_ШИРИНА_MIN}}'] = width_stats['min_width']
            metadata['{{ДОРОГА_ШИРИНА_MAX}}'] = width_stats['max_width']
            metadata['{{ПИКЕТ_ШИРИНА_MIN}}'] = width_stats['min_piket']
            metadata['{{ПИКЕТ_ШИРИНА_MAX}}'] = width_stats['max_piket']
            print(f"   📏 Обновлены метаданные: MIN={width_stats['min_width']} ({width_stats['min_piket']}), MAX={width_stats['max_width']} ({width_stats['max_piket']})")
    # ============================================================

    pp_number = str(metadata.get('{{НОМЕР_ПП}}', '')).strip() if metadata else ''
    sec_number = str(metadata.get('{{НОМЕР_УЧСТОК}}', '')).strip() if metadata else ''
    
    if pp_number and pp_number != '-':
        # Используем название из Excel для имени файла
        road_name = metadata.get('{{ДОРОГА_НАЗВАНИЕ}}', folder_name) if metadata else folder_name
        # Очищаем название от недопустимых символов для имени файла
        road_name_clean = re.sub(r'[<>:"/\\|?*]', '_', road_name)
        output_filename = f"{pp_number}_{sec_number}_{road_name_clean}_report.docx"
    else:
        output_filename = f"{folder_name}_report.docx"
        
    output_file = OUTPUT_DIR / output_filename
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # ВАЖНОЕ ИЗМЕНЕНИЕ: Копируем шаблон в выходной файл
    shutil.copy2(template_path, output_file)
    
    # ОТКРЫВАЕМ ФАЙЛ, КОТОРЫЙ БУДЕМ РЕДАКТИРОВАТЬ
    doc = Document(output_file)
    
    if metadata:
        # Используем функцию с сохранением стилей для текстовых плейсхолдеров
        replace_text_placeholders_preserve_style(doc, metadata)
        print("✅ Текстовые плейсхолдеры заменены с сохранением стилей.")
    else:
        print("⚠️ Метаданные не загружены.")
    
    # 1. Сначала вставляем ТАБЛИЦЫ
    results = process_all_tables(doc, folder_path)
    
    # 2. Сохраняем файл и ЗАКРЫВАЕМ его в python-docx
    doc.save(output_file)
    
    # 3. ОТКРЫВАЕМ ЭТОТ ЖЕ ФАЙЛ ЗАНОВО для вставки фото
    # Это сбрасывает все конфликты ID картинок и даёт Word "чистый лист"
    doc = Document(output_file)
    
    # 4. Вставляем ФОТОГРАФИИ в только что открытый документ
    photos_processor.process_kern_photos(doc, folder_path)
    photos_processor.process_strength_photos(doc, folder_path)
    photos_processor.process_width_photos(doc, folder_path)
    photos_processor.distribute_photos_and_pikets(doc, folder_path)
    
    # 5. Сохраняем финальную версию
    doc.save(output_file)
    print(f"\n✅ Отчет сохранен: {output_file}")
    
    # ====================================================
    # РАСШИРЕННЫЙ JSON С МЕТАДАННЫМИ
    # ====================================================
    result_file = output_file.with_suffix('.json')
    
    # Подготавливаем данные для JSON
    json_data = {
        'folder': folder_name,
        'output_file': str(output_file),
        'processed_tables': results
    }
    
    # Добавляем метаданные, если они есть
    if metadata:
        json_data['metadata'] = {
            'номер_пп': metadata.get('{{НОМЕР_ПП}}', ''),
            'номер_участка': metadata.get('{{НОМЕР_УЧСТОК}}', ''),
            'идентификационный_номер': metadata.get('{{ДОРОГА_ИД}}', ''),
            'название_дороги_из_папки': metadata.get('{{ДОРОГА_НАЗВАНИЕ_ПАПКА}}', ''),
            'название_дороги_из_excel': metadata.get('{{ДОРОГА_НАЗВАНИЕ_EXCEL}}', ''),
            'название_дороги_в_отчете': metadata.get('{{ДОРОГА_НАЗВАНИЕ}}', ''),
            'балансодержатель': metadata.get('{{БАЛАНСОДЕРЖАТЕЛЬ}}', '')
        }
    
    with open(result_file, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    # ====================================================

    # =========================
    # Обновление полей через COM
    # =========================
    try:
        import win32com.client as win32
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        
        doc_com = word.Documents.Open(str(output_file))
        doc_com.Fields.Update()  # ВАЖНО: Обновляем все поля!
        doc_com.Save()
        doc_com.Close()
        word.Quit()
#        print("✅ Поля SEQ успешно обновлены перед сохранением.")
    except Exception as e:
        print(f"⚠️ Не удалось обновить поля через COM: {e}")
    # =========================
    #     
    return results

def main():
    print("=" * 80)
    print("🚀 ПАКЕТНАЯ ГЕНЕРАЦИЯ ОТЧЕТОВ (PANDAS)")
    print("=" * 80)
    
    if not Path(TEMPLATE_PATH).exists():
        print(f"❌ Шаблон не найден: {TEMPLATE_PATH}")
        return
    if not Path(CATALOG_PATH).exists():
        print(f"❌ Каталог объектов не найден: {CATALOG_PATH}")
        return
    
    print(f"📄 Шаблон: {TEMPLATE_PATH}")
    print(f"📁 Кол-во папок для обработки: {len(FOLDERS_TO_PROCESS)}")
    print()
    
    all_results = {}
    for folder_path_str in FOLDERS_TO_PROCESS:
        folder_path = Path(folder_path_str)
        if not folder_path.exists():
            print(f"⚠️ Папка не найдена, пропуск: {folder_path_str}")
            continue
        
        results = generate_report_for_single_folder(folder_path, TEMPLATE_PATH)
        all_results[folder_path.name] = results
    
    print("\n" + "=" * 80)
    print("📊 ИТОГОВЫЙ ОТЧЕТ ПО ВСЕМ ПАПКАМ")
    print("=" * 80)
    
    total_success = 0
    total_errors = 0
    for folder_name, res_list in all_results.items():
        success = sum(1 for r in res_list if r['status'] == 'success')
        errors = sum(1 for r in res_list if r['status'] in ['error', 'failed'])
        total_success += success
        total_errors += errors
        print(f"📁 {folder_name}: ✅ {success} / ❌ {errors}")
    
    print(f"\n✅ Всего успешно: {total_success}")
    print(f"❌ Всего ошибок: {total_errors}")

if __name__ == "__main__":
    main()