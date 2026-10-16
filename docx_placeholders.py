"""Безопасная замена текстовых плейсхолдеров в DOCX без пересоздания абзацев."""

from __future__ import annotations


def replace_in_paragraph(paragraph, placeholder: str, replacement: str) -> int:
    """Заменяет все вхождения, в том числе разделённые между несколькими runs.

    Стиль вставленного текста наследуется от первого run, содержащего начало
    плейсхолдера. Незатронутые runs и XML-поля Word не пересоздаются.
    """
    if not placeholder or placeholder not in paragraph.text:
        return 0

    replacements = 0
    while placeholder in ''.join(run.text for run in paragraph.runs):
        runs = list(paragraph.runs)
        full_text = ''.join(run.text for run in runs)
        start = full_text.find(placeholder)
        end = start + len(placeholder)
        cursor = 0
        start_index = end_index = None

        for index, run in enumerate(runs):
            next_cursor = cursor + len(run.text)
            if start_index is None and start < next_cursor:
                start_index = index
            if end <= next_cursor:
                end_index = index
                break
            cursor = next_cursor

        if start_index is None or end_index is None:
            return replacements

        start_offset = start - sum(len(run.text) for run in runs[:start_index])
        end_offset = end - sum(len(run.text) for run in runs[:end_index])
        if start_index == end_index:
            run = runs[start_index]
            run.text = run.text[:start_offset] + replacement + run.text[end_offset:]
        else:
            runs[start_index].text = runs[start_index].text[:start_offset] + replacement
            for index in range(start_index + 1, end_index):
                runs[index].text = ''
            runs[end_index].text = runs[end_index].text[end_offset:]
        replacements += 1
    return replacements


def iter_container_paragraphs(container):
    """Возвращает абзацы контейнера и всех вложенных таблиц."""
    for paragraph in container.paragraphs:
        yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from iter_container_paragraphs(cell)


def iter_document_paragraphs(document):
    """Возвращает абзацы тела документа, таблиц и колонтитулов без дублей."""
    seen = set()
    containers = [document]
    for section in document.sections:
        containers.extend((section.header, section.footer, section.first_page_header,
                           section.first_page_footer, section.even_page_header,
                           section.even_page_footer))
    for container in containers:
        for paragraph in iter_container_paragraphs(container):
            element_id = id(paragraph._element)
            if element_id not in seen:
                seen.add(element_id)
                yield paragraph


def replace_in_document(document, replacements: dict[str, object]) -> int:
    """Заменяет все текстовые плейсхолдеры в документе, таблицах и колонтитулах."""
    count = 0
    for paragraph in iter_document_paragraphs(document):
        for placeholder, replacement in replacements.items():
            count += replace_in_paragraph(paragraph, placeholder, str(replacement))
    return count