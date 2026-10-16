"""
analyze_numbering.py
Скрипт для анализа нумерованных списков в документе Word.
Выводит информацию о всех абзацах с нумерацией.
"""

from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml

def analyze_document_numbering(doc_path):
    """
    Анализирует нумерацию в документе Word и выводит подробную информацию.
    """
    print("=" * 80)
    print(f"🔍 АНАЛИЗ НУМЕРОВАННЫХ СПИСКОВ В ДОКУМЕНТЕ")
    print("=" * 80)
    print(f"📄 Файл: {doc_path}")
    print("-" * 80)
    
    doc = Document(doc_path)
    
    # Получаем информацию о нумерации из XML
    numbering_part = doc.part.numbering_part
    if numbering_part is None:
        print("⚠️ В документе нет информации о нумерации.")
        return
    
    numbering_xml = numbering_part.element
    print("\n📋 ИНФОРМАЦИЯ О НУМЕРАЦИИ ИЗ XML:\n")
    
    # Находим все элементы <w:num> (определения нумерации)
    num_elements = numbering_xml.findall('.//w:num', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
    
    print(f"🔢 Найдено определений нумерации: {len(num_elements)}")
    
    for num_elem in num_elements:
        num_id = num_elem.get(qn('w:numId'))
        print(f"\n  📌 Нумерация ID: {num_id}")
        
        # Находим связанный абстрактный номер
        abstract_num_id_elem = num_elem.find('.//w:abstractNumId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if abstract_num_id_elem is not None:
            abstract_num_id = abstract_num_id_elem.get(qn('w:val'))
            print(f"     Абстрактный ID: {abstract_num_id}")
        
        # Проверяем, есть ли имя у нумерации
        name_elem = num_elem.find('.//w:name', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if name_elem is not None:
            name = name_elem.get(qn('w:val'))
            print(f"     Имя: {name}")
    
    # Находим все элементы <w:abstractNum> (абстрактные определения)
    abstract_num_elements = numbering_xml.findall('.//w:abstractNum', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
    
    print(f"\n📊 Абстрактных определений: {len(abstract_num_elements)}")
    
    for abstract_num in abstract_num_elements:
        abstract_num_id = abstract_num.get(qn('w:abstractNumId'))
        print(f"\n  📌 Абстрактный ID: {abstract_num_id}")
        
        # Находим уровни
        lvl_elements = abstract_num.findall('.//w:lvl', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        print(f"     Уровней: {len(lvl_elements)}")
        
        for lvl in lvl_elements:
            ilvl = lvl.get(qn('w:ilvl'))
            print(f"        Уровень {ilvl}:")
            
            # Формат нумерации
            num_fmt = lvl.find('.//w:numFmt', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if num_fmt is not None:
                fmt_val = num_fmt.get(qn('w:val'))
                print(f"          Формат: {fmt_val}")
            
            # Текст уровня (что отображается)
            lvl_text = lvl.find('.//w:lvlText', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if lvl_text is not None:
                text_val = lvl_text.get(qn('w:val'))
                print(f"          Текст: {text_val}")
            
            # Отступы
            p_pr = lvl.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if p_pr is not None:
                indent = p_pr.find('.//w:ind', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if indent is not None:
                    left = indent.get(qn('w:left'))
                    hanging = indent.get(qn('w:hanging'))
                    if left:
                        print(f"          Отступ слева: {left}")
                    if hanging:
                        print(f"          Отступ с висячей строкой: {hanging}")
    
    # Анализируем абзацы с нумерацией в документе
    print("\n" + "=" * 80)
    print("📝 АБЗАЦЫ С НУМЕРАЦИЕЙ В ДОКУМЕНТЕ")
    print("=" * 80)
    
    numbered_paragraphs = []
    
    for i, paragraph in enumerate(doc.paragraphs):
        # Проверяем, есть ли у параграфа нумерация
        p = paragraph._element
        p_pr = p.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        
        if p_pr is not None:
            num_pr = p_pr.find('.//w:numPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if num_pr is not None:
                # Получаем ID нумерации
                num_id_elem = num_pr.find('.//w:ilvl', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                ilvl = num_id_elem.get(qn('w:val')) if num_id_elem is not None else '?'
                
                num_id_elem = num_pr.find('.//w:numId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                num_id = num_id_elem.get(qn('w:val')) if num_id_elem is not None else '?'
                
                text = paragraph.text.strip()
                if text:
                    numbered_paragraphs.append({
                        'index': i,
                        'num_id': num_id,
                        'ilvl': ilvl,
                        'text': text[:100]  # Обрезаем длинный текст
                    })
                    
                    print(f"\n  📍 Абзац {i}:")
                    print(f"     ID нумерации: {num_id}")
                    print(f"     Уровень: {ilvl}")
                    print(f"     Текст: {text[:100]}{'...' if len(text) > 100 else ''}")
                    
                    # Показываем стиль параграфа
                    style = paragraph.style.name if paragraph.style else 'Нет стиля'
                    print(f"     Стиль: {style}")
    
    if not numbered_paragraphs:
        print("\n⚠️ В документе не найдено абзацев с нумерацией.")
    else:
        print(f"\n✅ Найдено абзацев с нумерацией: {len(numbered_paragraphs)}")
        
        # Группируем по ID нумерации
        num_ids = {}
        for p in numbered_paragraphs:
            num_id = p['num_id']
            if num_id not in num_ids:
                num_ids[num_id] = []
            num_ids[num_id].append(p)
        
        print("\n📊 ГРУППИРОВКА ПО ID НУМЕРАЦИИ:")
        for num_id, paras in num_ids.items():
            print(f"\n  ID {num_id}: {len(paras)} абзацев")
            for p in paras[:3]:  # Показываем первые 3
                print(f"    - [{p['ilvl']}] {p['text'][:60]}...")
            if len(paras) > 3:
                print(f"    ... и еще {len(paras) - 3} абзацев")
    
    return numbered_paragraphs

def find_placeholder_references(doc_path):
    """
    Находит все плейсхолдеры в документе, связанные с нумерацией.
    """
    print("\n" + "=" * 80)
    print("🔍 ПОИСК ПЛЕЙСХОЛДЕРОВ СВЯЗАННЫХ С НУМЕРАЦИЕЙ")
    print("=" * 80)
    
    doc = Document(doc_path)
    
    placeholder_patterns = [
        '{{НОМЕР_ПП}}',
        '{{НОМЕР_УЧСТОК}}',
        '{{НОМЕР_ШУРФ_КЕРН}}',
        '{{ПИКЕТ_ШИРИНА_MIN}}',
        '{{ПИКЕТ_ШИРИНА_MAX}}',
        '{{ПИКЕТ_ПРОЧНОСТЬ}}',
    ]
    
    found = []
    
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text
        for placeholder in placeholder_patterns:
            if placeholder in text:
                found.append({
                    'index': i,
                    'placeholder': placeholder,
                    'text': text.strip()[:100]
                })
                print(f"\n  📍 Абзац {i}:")
                print(f"     Плейсхолдер: {placeholder}")
                print(f"     Текст: {text.strip()[:100]}{'...' if len(text.strip()) > 100 else ''}")
                
                # Проверяем, есть ли у этого абзаца нумерация
                p = paragraph._element
                p_pr = p.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if p_pr is not None:
                    num_pr = p_pr.find('.//w:numPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    if num_pr is not None:
                        num_id_elem = num_pr.find('.//w:numId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                        num_id = num_id_elem.get(qn('w:val')) if num_id_elem is not None else '?'
                        ilvl_elem = num_pr.find('.//w:ilvl', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                        ilvl = ilvl_elem.get(qn('w:val')) if ilvl_elem is not None else '?'
                        print(f"     ID нумерации: {num_id}, уровень: {ilvl}")
    
    if not found:
        print("\n⚠️ Плейсхолдеры не найдены.")
    
    return found

def main():
    # Путь к файлу для анализа
    doc_path = r"D:\ProgrammEngine\road_inspection_automation\template\template_6.docx"
    
    if not Path(doc_path).exists():
        print(f"❌ Файл не найден: {doc_path}")
        return
    
    # Анализируем нумерацию
    numbered_paragraphs = analyze_document_numbering(doc_path)
    
    # Ищем плейсхолдеры
    placeholders = find_placeholder_references(doc_path)
    
    print("\n" + "=" * 80)
    print("📊 ИТОГОВЫЙ ОТЧЕТ")
    print("=" * 80)
    print(f"✅ Анализ завершен!")
    print(f"📁 Файл: {doc_path}")
    print(f"📝 Найдено абзацев с нумерацией: {len(numbered_paragraphs)}")
    print(f"🔍 Найдено плейсхолдеров: {len(placeholders)}")

if __name__ == "__main__":
    main()