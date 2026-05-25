"""
Склейщик документов (merger.py).
Склеивает подготовленные .docx в итоговый документ БЕЗ ПУСТЫХ СТРАНИЦ В НАЧАЛЕ,
С ИЗОБРАЖЕНИЯМИ И С ЖЕСТКИМ РАЗДЕЛЕНИЕМ ПО СТРАНИЦАМ (каждый файл на своем листе).
"""

import copy
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _remove_last_empty_paragraphs(doc: Document) -> None:
    """Удаляет пустые параграфы в конце документа, создающие лишние страницы."""
    while doc.paragraphs:
        last_para = doc.paragraphs[-1]
        if not last_para.text.strip() and not last_para._element.findall('.//' + qn('w:drawing')):
            last_para._element.getparent().remove(last_para._element)
        else:
            break


def _is_paragraph_empty(para_element) -> bool:
    """Вспомогательная функция для проверки: пустой ли XML-элемент параграфа."""
    text_elements = para_element.findall('.//' + qn('w:t'))
    drawing_elements = para_element.findall('.//' + qn('w:drawing'))

    has_content = False
    for t in text_elements:
        if t.text and t.text.strip():
            has_content = True
            break
    if drawing_elements:
        has_content = True

    return not has_content


def _copy_images_and_fix_rels(source_doc: Document, target_doc: Document, element_copy) -> None:
    """Переносит файлы картинок в целевой документ и обновляет ID связей (rId)."""
    blip_elements = element_copy.findall('.//' + qn('a:blip'))
    if not blip_elements:
        return

    source_part = source_doc.part
    target_part = target_doc.part

    for blip in blip_elements:
        old_rid = blip.get(qn('r:embed'))
        if not old_rid:
            continue

        try:
            source_rel = source_part.rels[old_rid]
            image_part = source_rel.target_part

            image_bytes = image_part.blob
            content_type = image_part.content_type

            # Регистрируем картинку в новом документе
            new_rid, _ = target_part.get_or_add_image_relationship_by_data(image_bytes, content_type)
            blip.set(qn('r:embed'), new_rid)
        except Exception as e:
            print(f"[Ворнинг] Не удалось скопировать связь изображения {old_rid}: {e}")


def merge_to_docx_and_pdf(
    paths: List[Path],
    out_docx: Path,
    out_pdf: Path,
    page_break_between_parts: bool = True,
    insert_titles: bool = True,
    pdf_render_dpi: int = 150,
) -> Tuple[List[str], List[str]]:
    """
    Склеивает список .docx файлов в один итоговый документ.
    Каждый исходный файл занимает целое количество листов и не сливается со следующим.
    """
    warnings = []
    errors = []

    if not paths:
        errors.append("Нет файлов для склейки")
        return warnings, errors

    merged_doc = Document()

    # Настройка стиля по умолчанию
    style = merged_doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # РЕШЕНИЕ ПРОБЛЕМЫ 1: Удаляем абсолютно все пустые параграфы в самом начале,
    # чтобы итоговый файл не начинался с белого листа
    for p in list(merged_doc.paragraphs):
        p._element.getparent().remove(p._element)

    for i, file_path in enumerate(paths):
        if not file_path.exists():
            errors.append(f"Файл не найден: {file_path.name}")
            continue

        try:
            # РЕШЕНИЕ ПРОБЛЕМЫ 2: Чтобы файлы не шли тупо друг за другом и не налезали,
            # мы принудительно создаем новый чистый раздел (Section) со следующей страницы
            if i > 0:
                merged_doc.add_section(WD_SECTION.NEW_PAGE)

            # Вставляем заголовок с именем файла, если требуется
            if insert_titles:
                title_para = merged_doc.add_paragraph()
                title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                title_run = title_para.add_run(file_path.stem)
                title_run.bold = True
                title_run.font.size = Pt(14)

                spacer = merged_doc.add_paragraph()
                spacer.paragraph_format.space_after = Pt(6)
            elif page_break_between_parts and i > 0 and not insert_titles:
                # Если заголовков нет, но разрыв просили — он уже создался через add_section выше
                pass

            source_doc = Document(str(file_path))
            _remove_last_empty_paragraphs(source_doc)

            # Перенос элементов тела документа (включая таблицы и картинки)
            for element in source_doc.element.body:
                # Пропускаем старые свойства разделов исходника, чтобы они не ломали общую разметку
                if element.tag.endswith('}sectPr'):
                    continue

                if element.tag.endswith('}p') and _is_paragraph_empty(element):
                    continue

                # Глубокое копирование XML-элемента
                element_copy = copy.deepcopy(element)

                # РЕШЕНИЕ ПРОБЛЕМЫ 3: Восстановление картинок.
                # Переносим бинарные файлы картинок в архив нового docx и чиним XML rId связи
                _copy_images_and_fix_rels(source_doc, merged_doc, element_copy)

                # Добавляем элемент в документ
                merged_doc.element.body.append(element_copy)

        except Exception as e:
            errors.append(f"Ошибка обработки {file_path.name}: {e}")

    # Финальная чистка хвостов
    _remove_last_empty_paragraphs(merged_doc)

    # 1. Сохраняем DOCX
    try:
        merged_doc.save(str(out_docx))
    except Exception as e:
        errors.append(f"Ошибка сохранения DOCX: {e}")
        return warnings, errors

    # 2. Сохраняем честный PDF
    # Так как мы идеально перенесли структуру со всеми картинками внутрь результирующего DOCX,
    # автономная библиотека docx2pdf соберет идентичный PDF файл со всеми изображениями
    try:
        from docx2pdf import convert
        # Конвертируем сохраненный docx в pdf
        convert(str(out_docx), str(out_pdf))
    except Exception as pdf_err:
        warnings.append(f"Не удалось сгенерировать PDF: {pdf_err}. Проверьте наличие прав доступа.")

    return warnings, errors
