"""
photos_processor.py
Модуль для обработки фотографий.
Вставка фото и подписей кернов прямо перед плейсхолдером.
ПОРЯДОК ИСПРАВЛЕН + удаление шаблонного хвоста.
"""

import os
import shutil
import zipfile
import re
from pathlib import Path
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pandas as pd
import numpy as np
from PIL import Image

# ================= НАСТРОЙКИ =================
PHOTO_WIDTH_CM = 10.0
# ==============================================

def convert_piket_to_gost(piket_string):
    piket_string = piket_string.strip()
    match = re.search(r'ПК\s*(\d+)\+(\d+)', piket_string)
    if not match:
        return piket_string
    
    pk_part = match.group(1)
    meter_part = match.group(2)
    
    if len(pk_part) > 1:
        new_pk = pk_part[:-1]
        last_digit = pk_part[-1]
        result = f"ПК{new_pk}+{last_digit}{meter_part}"
    else:
        result = f"ПК{pk_part}+{meter_part}"
    
    return result

def find_placeholder_in_document(doc, placeholder):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            return {'type': 'paragraph', 'element': paragraph}
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    return {'type': 'table_cell', 'element': cell}
    return None




# ====================================================================
# КЕРНЫ - ВЕРСИЯ С ИСПОЛЬЗОВАНИЕМ СТИЛЯ
# ====================================================================

# ====================================================================
# КЕРНЫ - ВЕРСИЯ С ПРАВИЛЬНЫМ КОПИРОВАНИЕМ НУМЕРАЦИИ ID 19
# ====================================================================

def process_kern_photos(doc, object_folder_path):
    print("\n📸 [ФОТО-ПРОЦЕССОР] Запуск обработки кернов и шурфов...")
    
    # 1. Находим плейсхолдер
    anchor = None
    for p in doc.paragraphs:
        if "{{ФОТО_ШУРФ_КЕРН}}" in p.text:
            anchor = p
            break

    if not anchor:
        print(f"   ❌ В шаблоне не найден плейсхолдер {{ФОТО_ШУРФ_КЕРН}}")
        return

    # 2. Ищем фотоматериалы
    photos_base_dir = Path(object_folder_path) / "Фотоматериалы кернов и шурфов"
    if not photos_base_dir.exists():
        print(f"   ⚠️ Папка не найдена: {photos_base_dir}")
        return

    sample_folders = [f for f in photos_base_dir.iterdir() if f.is_dir()]
    if not sample_folders:
        print(f"   ⚠️ В папке '{photos_base_dir.name}' нет подпапок с образцами.")
        return

    sample_folders.sort(key=lambda x: int(re.search(r'Образец №\s*(\d+)', x.name).group(1)) if re.search(r'Образец №\s*(\d+)', x.name) else 0)
    print(f"   📁 Найдено папок: {len(sample_folders)}")

    # ==========================================================
    # НАХОДИМ ЭТАЛОННЫЙ ЭЛЕМЕНТ СПИСКА ID 19
    # ==========================================================
    template_paragraph = None
    for paragraph in doc.paragraphs:
        p = paragraph._element
        p_pr = p.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if p_pr is not None:
            num_pr = p_pr.find('.//w:numPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if num_pr is not None:
                num_id_elem = num_pr.find('.//w:numId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if num_id_elem is not None and num_id_elem.get(qn('w:val')) == '19':
                    template_paragraph = paragraph
                    break

    # 3. Вставляем фото и подписи
    for folder in sample_folders:
        match = re.search(r'Образец №\s*(\d+)\s+(ПК\s*\d+\+\d+)', folder.name)
        if not match:
            continue
        
        sample_number = match.group(1)
        raw_piket = match.group(2)
        converted_piket = convert_piket_to_gost(raw_piket)

        # ФОТО
        all_files = list(folder.iterdir())
        image_files = [f for f in all_files if f.suffix.lower() in ['.jpg', '.jpeg', '.png']]
        if image_files:
            print(f"      📸 Вставляем {len(image_files)} фото для Образца №{sample_number}")
            for img_file in image_files:
                try:
                    pil_img = Image.open(img_file)
                    if pil_img.mode in ('RGBA', 'LA') or (pil_img.mode == 'P' and 'transparency' in pil_img.info):
                        background = Image.new('RGB', pil_img.size, (255, 255, 255))
                        background.paste(pil_img, mask=pil_img.split()[-1])
                        pil_img = background
                    img_bytes = BytesIO()
                    pil_img.save(img_bytes, format='JPEG', quality=95)
                    img_bytes.seek(0)
                    
                    photo_para = anchor.insert_paragraph_before()
                    photo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = photo_para.add_run()
                    run.add_picture(img_bytes, width=Cm(PHOTO_WIDTH_CM))
                except Exception as e:
                    print(f"      ⚠️ Ошибка при обработке {img_file.name}: {e}")
        else:
            print(f"      ⚠️ В папке {folder.name} нет фото.")

        # ==========================================================
        # ПОДПИСЬ - КОПИРУЕМ ВЕСЬ АБЗАЦ С ЭТАЛОНА
        # ==========================================================
        caption_para = anchor.insert_paragraph_before()
        
        if template_paragraph is not None:
            # КОПИРУЕМ ВСЁ СОДЕРЖИМОЕ АБЗАЦА
            # 1. Копируем XML элемента абзаца
            template_p = template_paragraph._element
            new_p = caption_para._element
            
            # Копируем все свойства абзаца
            template_p_pr = template_p.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if template_p_pr is not None:
                # Удаляем старый pPr
                old_p_pr = new_p.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if old_p_pr is not None:
                    new_p.remove(old_p_pr)
                
                # Создаем копию pPr
                new_p_pr = OxmlElement('w:pPr')
                for child in template_p_pr:
                    new_p_pr.append(child)
                new_p.insert(0, new_p_pr)
            
            # Копируем стиль
            if template_paragraph.style:
                caption_para.style = template_paragraph.style
            
            # Копируем форматирование
            caption_para.paragraph_format.alignment = template_paragraph.paragraph_format.alignment
            if template_paragraph.paragraph_format.left_indent:
                caption_para.paragraph_format.left_indent = template_paragraph.paragraph_format.left_indent
            if template_paragraph.paragraph_format.first_line_indent:
                caption_para.paragraph_format.first_line_indent = template_paragraph.paragraph_format.first_line_indent
        
        # Устанавливаем текст
        caption_para.text = f"– Керн №{sample_number} на {converted_piket}"
        
        # Устанавливаем шрифт для всех прогонов
        for run in caption_para.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
        
        print(f"   ✅ Вставлена подпись: – Керн №{sample_number} на {converted_piket}")

        # Пустая строка
        anchor.insert_paragraph_before()

    # 4. Удаляем якорь
    anchor._element.getparent().remove(anchor._element)

    print(f"\n   ✅ Обработка фото кернов завершена!")


# ====================================================================
# ПРОЧНОСТЬ 
# ====================================================================

def process_strength_photos(doc, object_folder_path):
    print("\n🔧 [ФОТО-ПРОЦЕССОР] Обработка фото прочности...")
    
    found_photo = find_placeholder_in_document(doc, "{{ФОТО_ПРОЧНОСТЬ}}")
    if not found_photo:
        print("   ⚠️ Плейсхолдер {{ФОТО_ПРОЧНОСТЬ}} не найден.")
        return

    photos_base_dir = Path(object_folder_path)
    strength_folders = [f for f in photos_base_dir.iterdir() if f.is_dir() and "измерения прочности" in f.name.lower()]
    if not strength_folders:
        print(f"   ⚠️ Папка с 'измерения прочности' не найдена.")
        return
    
    target_folder = strength_folders[0]
    print(f"   📁 Найдена папка: {target_folder.name}")

    all_files = list(target_folder.iterdir())
    image_files = [f for f in all_files if f.suffix.lower() in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']]
    if not image_files:
        print("   ⚠️ В папке нет изображений.")
        return

    target_image = None
    for f in image_files:
        if "1 замер" in f.stem.lower():
            target_image = f
            break
    if target_image is None:
        target_image = image_files[0]
        print(f"   ⚠️ Файл с '1 замер' не найден, использую: {target_image.name}")
    else:
        print(f"   ✅ Найден файл: {target_image.name}")

    if found_photo['type'] == 'paragraph':
        para = found_photo['element']
        para.clear()
        run = para.add_run()
        run.add_picture(str(target_image), width=Cm(PHOTO_WIDTH_CM))
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        cell = found_photo['element']
        cell.text = ""
        para = cell.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(str(target_image), width=Cm(PHOTO_WIDTH_CM))
    print("   ✅ Фото вставлено на место {{ФОТО_ПРОЧНОСТЬ}}")

    found_piket = find_placeholder_in_document(doc, "{{ПИКЕТ_ПРОЧНОСТЬ}}")
    if found_piket:
        if found_piket['type'] == 'paragraph':
            para = found_piket['element']
            for run in para.runs:
                if 'SEQ' not in run.text and '{{' in run.text:
                    run.text = run.text.replace("{{ПИКЕТ_ПРОЧНОСТЬ}}", "ПК0+000")
                    break
        else:
            found_piket['element'].text = found_piket['element'].text.replace("{{ПИКЕТ_ПРОЧНОСТЬ}}", "ПК0+000")
        print("   ✅ {{ПИКЕТ_ПРОЧНОСТЬ}} заменён на ПК0+000")
    else:
        print("   ⚠️ Плейсхолдер {{ПИКЕТ_ПРОЧНОСТЬ}} не найден.")

    print("   ✅ Обработка фото прочности завершена!")


# ====================================================================
# ШИРИНА
# ====================================================================

def process_width_photos(doc, object_folder_path):
    print("\n📏 [ФОТО-ПРОЦЕССОР] Обработка фото ширины...")
    
    found_min_ph = find_placeholder_in_document(doc, "{{ФОТО_ШИРИНА_MIN}}")
    found_max_ph = find_placeholder_in_document(doc, "{{ФОТО_ШИРИНА_MAX}}")
    if not found_min_ph or not found_max_ph:
        print("   ⚠️ Плейсхолдеры {{ФОТО_ШИРИНА_MIN/MAX}} не найдены.")
        return

    # Получаем данные о ширине
    excel_file = None
    for file in Path(object_folder_path).glob("*.xlsx"):
        if "Ширина" in file.name:
            excel_file = file
            break
    
    if not excel_file:
        print("   ⚠️ Excel-файл с 'Ширина' не найден.")
        return

    try:
        df_raw = pd.read_excel(excel_file, header=None)
        
        # Ищем заголовки
        header_row_idx = None
        for i in range(min(20, len(df_raw))):
            row_values = df_raw.iloc[i].astype(str).str.strip().tolist()
            if any('Местоположение' in str(v) or 'Ширина' in str(v) for v in row_values):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            headers = df_raw.iloc[12, :4].fillna('').tolist()
            data = df_raw.iloc[14:, :4].copy()
            data.columns = headers
        else:
            headers = df_raw.iloc[header_row_idx, :4].fillna('').tolist()
            data = df_raw.iloc[header_row_idx + 1:, :4].copy()
            data.columns = headers
        
        data = data.dropna(how='all')
        
        # Ищем колонку с шириной
        width_col = None
        for col in data.columns:
            if 'Ширина' in str(col):
                width_col = col
                break
        
        if width_col is None:
            print("   ⚠️ Колонка с шириной не найдена")
            return
        
        # Преобразуем в числа
        data['_width_num'] = data[width_col].astype(str).str.replace(',', '.').str.replace(' ', '')
        data['_width_num'] = pd.to_numeric(data['_width_num'], errors='coerce')
        data_clean = data.dropna(subset=['_width_num'])
        
        if data_clean.empty:
            print("   ⚠️ Нет валидных данных о ширине")
            return
        
        # Находим min и max
        min_row = data_clean.loc[data_clean['_width_num'].idxmin()]
        max_row = data_clean.loc[data_clean['_width_num'].idxmax()]
        
        # Ищем колонку с местоположением
        loc_col = None
        for col in data_clean.columns:
            if 'Местоположение' in str(col) or 'км+м' in str(col):
                loc_col = col
                break
        
        if loc_col is None:
            for col in data_clean.columns:
                if '№' in str(col):
                    loc_col = col
                    break
        
        min_coord = str(min_row[loc_col]).strip() if loc_col else ""
        max_coord = str(max_row[loc_col]).strip() if loc_col else ""
        
        print(f"   ✅ Найдена МИН ширина: {min_row['_width_num']} м, координата: {min_coord}")
        print(f"   ✅ Найдена МАКС ширина: {max_row['_width_num']} м, координата: {max_coord}")
        
    except Exception as e:
        print(f"   ❌ Ошибка при парсинге Excel ширины: {e}")
        import traceback
        traceback.print_exc()
        return

    # Поиск фото по координатам (как было раньше)
    photos_dir = None
    for folder in Path(object_folder_path).iterdir():
        if folder.is_dir() and ("ширины" in folder.name.lower() or "Ширина" in folder.name):
            photos_dir = folder
            break
    
    if not photos_dir:
        print("   ⚠️ Папка с фотографиями ширины не найдена.")
        return

    print(f"   📁 Найдена папка с фото: {photos_dir.name}")

    def find_photo_by_coord(coord, folder):
        if not coord:
            return None
        clean_coord = coord.replace(',', '').replace(' ', '').replace('+', '')
        for file in folder.iterdir():
            if file.suffix.lower() in ['.jpg', '.jpeg', '.png']:
                file_clean = file.stem.replace(',', '').replace(' ', '').replace('+', '')
                if clean_coord in file_clean or file_clean in clean_coord:
                    return file
        return None

    min_photo = find_photo_by_coord(min_coord, photos_dir)
    max_photo = find_photo_by_coord(max_coord, photos_dir)

    def insert_photo_with_height(placeholder, photo_path, height_cm):
        if placeholder['type'] == 'paragraph':
            para = placeholder['element']
            para.clear()
            run = para.add_run()
            run.add_picture(str(photo_path), height=Cm(height_cm))
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            cell = placeholder['element']
            cell.text = ""
            para = cell.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(str(photo_path), height=Cm(height_cm))

    if min_photo:
        insert_photo_with_height(found_min_ph, min_photo, 11.0)
        print(f"   ✅ Вставлено фото МИН ширины (высота 11 см, коорд. {min_coord})")
    else:
        print(f"   ⚠️ Фото МИН ширины не найдено")

    if max_photo:
        insert_photo_with_height(found_max_ph, max_photo, 11.0)
        print(f"   ✅ Вставлено фото МАКС ширины (высота 11 см, коорд. {max_coord})")
    else:
        print(f"   ⚠️ Фото МАКС ширины не найдено")

    print("   ✅ Обработка фото ширины завершена!")


# ====================================================================
# ОБЩИЕ ВИДЫ (7 фото)
# ====================================================================
def get_total_length_from_longitudinal_excel(object_folder_path):
    for file in Path(object_folder_path).glob("*.xlsx"):
        if "Продольная ровность" in file.name:
            try:
                df = pd.read_excel(file, header=None)
                header_row_idx = None
                for i, row in df.iterrows():
                    if "Конец участка" in str(row.values):
                        header_row_idx = i
                        break
                if header_row_idx is None:
                    return None
                last_row = None
                for i in range(len(df) - 1, header_row_idx, -1):
                    val = df.iloc[i, 1]
                    if pd.notna(val) and str(val).strip() != "":
                        last_row = i
                        break
                if last_row is None:
                    return None
                end_piket_str = str(df.iloc[last_row, 1]).strip()
                if "+" in end_piket_str:
                    parts = end_piket_str.split("+")
                    km = int(parts[0])
                    m = int(parts[1])
                    return km * 1000 + m
                else:
                    return None
            except Exception as e:
                print(f"   ⚠️ Ошибка при чтении длины дороги: {e}")
                return None
    return None

def replace_text_preserve_style(paragraph, old_text, new_text, font_size=12):
    """
    Заменяет текст в параграфе, сохраняя стили и поля SEQ
    """
    if old_text not in paragraph.text:
        return False
    
    # Получаем все прогоны
    runs = paragraph.runs
    if not runs:
        # Если нет прогонов, создаем новый
        paragraph.text = paragraph.text.replace(old_text, new_text)
        # Устанавливаем шрифт для нового текста
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.font.name = 'Times New Roman'
        return True
    
    # Собираем весь текст параграфа
    full_text = ''.join(run.text for run in runs)
    
    # Если old_text не найден в полном тексте, пробуем заменить напрямую
    if old_text not in full_text:
        paragraph.text = paragraph.text.replace(old_text, new_text)
        # Устанавливаем шрифт для нового текста
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.font.name = 'Times New Roman'
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
            # Устанавливаем шрифт 12 для измененного прогона
            run.font.size = Pt(font_size)
            run.font.name = 'Times New Roman'
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
    
    # Обновляем размер шрифта для всех прогонов
    for run in runs:
        if run.font.size is None:
            run.font.size = Pt(font_size)
        if run.font.name is None:
            run.font.name = 'Times New Roman'
    
    return True

def clean_unused_placeholders(doc):
    """
    Удаляет незаполненные плейсхолдеры из документа
    """
    placeholders_to_remove = [
        '{{ФОТО_ВИД_ДОРОГА_1}}',
        '{{ФОТО_ВИД_ДОРОГА_2}}',
        '{{ФОТО_ВИД_ДОРОГА_3}}',
        '{{ФОТО_ВИД_ДОРОГА_4}}',
        '{{ФОТО_ВИД_ДОРОГА_5}}',
        '{{ФОТО_ВИД_ДОРОГА_6}}',
        '{{ФОТО_ВИД_ДОРОГА_7}}',
        '{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_1}}',
        '{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_2}}',
        '{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_3}}',
        '{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_4}}',
        '{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_5}}',
        '{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_6}}',
        '{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_7}}',
        # Добавляем также плейсхолдеры для других типов фото
        '{{ФОТО_КЕРН_1}}',
        '{{ФОТО_КЕРН_2}}',
        '{{ФОТО_КЕРН_3}}',
        '{{ФОТО_КЕРН_4}}',
        '{{ФОТО_КЕРН_5}}',
        '{{ФОТО_КЕРН_6}}',
        '{{ФОТО_КЕРН_7}}',
        '{{ФОТО_КЕРН_8}}',
        '{{ФОТО_КЕРН_9}}',
        '{{ФОТО_КЕРН_10}}',
        '{{ФОТО_КЕРН_11}}',
        '{{ФОТО_КЕРН_12}}',
        '{{ФОТО_ПРОЧНОСТЬ_1}}',
        '{{ФОТО_ПРОЧНОСТЬ_2}}',
        '{{ФОТО_ПРОЧНОСТЬ_3}}',
        '{{ФОТО_ПРОЧНОСТЬ_4}}',
        '{{ФОТО_ПРОЧНОСТЬ_5}}',
        '{{ФОТО_ПРОЧНОСТЬ_6}}',
        '{{ФОТО_ПРОЧНОСТЬ_7}}',
        '{{ФОТО_ПРОЧНОСТЬ_8}}',
        '{{ФОТО_ПРОЧНОСТЬ_9}}',
        '{{ФОТО_ПРОЧНОСТЬ_10}}',
        '{{ФОТО_ШИРИНА_1}}',
        '{{ФОТО_ШИРИНА_2}}',
        '{{ФОТО_ШИРИНА_3}}',
        '{{ФОТО_ШИРИНА_4}}',
        '{{ФОТО_ШИРИНА_5}}',
        '{{ФОТО_ШИРИНА_6}}',
        '{{ФОТО_ШИРИНА_7}}',
        '{{ФОТО_ШИРИНА_8}}',
        '{{ФОТО_ШИРИНА_9}}',
        '{{ФОТО_ШИРИНА_10}}',
    ]
    
    removed_count = 0
    
    # Обрабатываем параграфы - собираем пустые для удаления
    paragraphs_to_remove = []
    for paragraph in doc.paragraphs:
        for placeholder in placeholders_to_remove:
            if placeholder in paragraph.text:
                # Удаляем плейсхолдер из текста
                paragraph.text = paragraph.text.replace(placeholder, '')
                # Если параграф стал пустым, помечаем на удаление
                if not paragraph.text.strip():
                    paragraphs_to_remove.append(paragraph)
                removed_count += 1
                break
    
    # Удаляем пустые параграфы
    for paragraph in paragraphs_to_remove:
        parent = paragraph._element.getparent()
        if parent is not None:
            parent.remove(paragraph._element)
    
    # Обрабатываем таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for placeholder in placeholders_to_remove:
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, '')
                        removed_count += 1
                        break
    
    if removed_count > 0:
        print(f"   🧹 Очищено {removed_count} незаполненных плейсхолдеров")
    
    return removed_count

def distribute_photos_and_pikets(doc, object_folder_path):
    print("\n🛣️ [ФОТО-ПРОЦЕССОР] Обработка общих видов дороги...")
    
    obj_name = object_folder_path.name
    photos_folder = None
    for folder in Path(object_folder_path).iterdir():
        if folder.is_dir() and f"Фотоматериалы {obj_name}" in folder.name:
            photos_folder = folder
            break
    if not photos_folder:
        print(f"   ⚠️ Папка 'Фотоматериалы {obj_name}' не найдена.")
        return

    all_images = [f for f in photos_folder.iterdir() if f.suffix.lower() in ['.jpg', '.jpeg', '.png']]
    all_images.sort(key=lambda x: x.name)
    
    total_length_m = get_total_length_from_longitudinal_excel(object_folder_path)
    if total_length_m is None:
        total_length_m = 3600
        print("   ⚠️ Не удалось получить длину дороги, использую 3600 м.")
    print(f"   📏 Общая длина дороги: {total_length_m} м")

    # ==========================================================
    # ГРАДАЦИЯ: определяем количество фото и шаг округления
    # ==========================================================
    if total_length_m <= 500:
        num_photos = 3
        round_step = 10  # округление до 10 метров
    elif total_length_m <= 1000:
        num_photos = 5
        round_step = 10
    else:
        num_photos = 7
        round_step = 100  # для длинных дорог округление до 100 м

    # Проверяем, хватает ли фото
    if len(all_images) < num_photos:
        print(f"   ⚠️ В папке найдено всего {len(all_images)} фото, а нужно минимум {num_photos}.")
        return

    # Вычисляем равномерные индексы
    n = len(all_images)
    indices = [int(round(i * (n - 1) / (num_photos - 1))) for i in range(num_photos)]
    selected_images = [all_images[i] for i in indices]
    
    print(f"   📁 Найдено фото: {len(all_images)}")
    print(f"   📸 Выбрано фото: {num_photos}, индексы: {indices}")

    # Вычисляем пикеты с учётом округления
    pikets_m = [int(round(i * total_length_m / (num_photos - 1) / round_step)) * round_step for i in range(num_photos)]
    pikets_str = []
    for p in pikets_m:
        km = p // 1000
        m = p % 1000
        pikets_str.append(f"ПК{km}+{m:03d}")
    
    # Первый фиксируем строго
    pikets_str[0] = "ПК0+000"
    print(f"   📍 Пикеты: {pikets_str}")

    # Размер шрифта для подписей
    CAPTION_FONT_SIZE = 12

    # Вставляем фото и пикеты
    for i in range(num_photos):
        ph_photo = f"{{{{ФОТО_ВИД_ДОРОГА_{i+1}}}}}"
        ph_piket = f"{{{{ФОТО_ВИД_ДОРОГА_ПИКЕТ_{i+1}}}}}"
        
        # Вставка ФОТО
        found_photo = find_placeholder_in_document(doc, ph_photo)
        if found_photo:
            if found_photo['type'] == 'paragraph':
                para = found_photo['element']
                # Проверяем, есть ли в параграфе другие элементы
                full_text = para.text
                if ph_photo in full_text:
                    parts = full_text.split(ph_photo)
                    prefix = parts[0]
                    suffix = parts[1] if len(parts) > 1 else ""
                    
                    # Очищаем параграф
                    para.clear()
                    
                    # Добавляем префикс если есть (шрифт 12)
                    if prefix:
                        run = para.add_run(prefix)
                        run.font.size = Pt(CAPTION_FONT_SIZE)
                        run.font.name = 'Times New Roman'
                    
                    # Добавляем изображение
                    run = para.add_run()
                    run.add_picture(str(selected_images[i]), width=Cm(14.0))
                    
                    # Добавляем суффикс если есть (шрифт 12)
                    if suffix:
                        run = para.add_run(suffix)
                        run.font.size = Pt(CAPTION_FONT_SIZE)
                        run.font.name = 'Times New Roman'
                    
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                # Вставка в таблицу
                cell = found_photo['element']
                # Проверяем, есть ли в ячейке другие элементы
                full_text = cell.text
                if ph_photo in full_text:
                    parts = full_text.split(ph_photo)
                    prefix = parts[0]
                    suffix = parts[1] if len(parts) > 1 else ""
                    
                    cell.text = ""
                    if prefix:
                        para = cell.add_paragraph(prefix)
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in para.runs:
                            run.font.size = Pt(CAPTION_FONT_SIZE)
                            run.font.name = 'Times New Roman'
                    
                    para = cell.add_paragraph()
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = para.add_run()
                    run.add_picture(str(selected_images[i]), width=Cm(14.0))
                    
                    if suffix:
                        para = cell.add_paragraph(suffix)
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in para.runs:
                            run.font.size = Pt(CAPTION_FONT_SIZE)
                            run.font.name = 'Times New Roman'
            
            print(f"   ✅ Вставлено фото {i+1}: {selected_images[i].name}")
            
        # Вставка ПИКЕТА с сохранением стилей (шрифт 12)
        found_piket = find_placeholder_in_document(doc, ph_piket)
        if found_piket:
            if found_piket['type'] == 'paragraph':
                para = found_piket['element']
                # Используем функцию для замены с сохранением стилей
                if replace_text_preserve_style(para, ph_piket, pikets_str[i], CAPTION_FONT_SIZE):
                    print(f"   ✅ Вставлен пикет {i+1}: {pikets_str[i]}")
                else:
                    # Если не удалось заменить с сохранением стилей, делаем простую замену
                    para.text = para.text.replace(ph_piket, pikets_str[i])
                    # Устанавливаем шрифт 12
                    for run in para.runs:
                        run.font.size = Pt(CAPTION_FONT_SIZE)
                        run.font.name = 'Times New Roman'
                    print(f"   ✅ Вставлен пикет {i+1}: {pikets_str[i]} (простая замена)")
            else:
                # В таблице
                cell = found_piket['element']
                cell.text = cell.text.replace(ph_piket, pikets_str[i])
                # Устанавливаем шрифт 12
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(CAPTION_FONT_SIZE)
                        run.font.name = 'Times New Roman'
                print(f"   ✅ Вставлен пикет {i+1}: {pikets_str[i]}")

    print(f"   ✅ Обработка общих видов дороги завершена! (фото: {num_photos})")
    
    # После вставки всех фото и пикетов - очищаем незаполненные плейсхолдеры
    clean_unused_placeholders(doc)