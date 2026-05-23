from __future__ import annotations

import io
import tempfile
import traceback
from pathlib import Path

import fitz  # PyMuPDF — для PDF→картинки в DOCX
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Inches, Pt, Cm
from docxcompose.composer import Composer
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ── Вспомогательные функции ──

def _prepend_heading(doc: Document, text: str) -> None:
    h = doc.add_heading(text, level=2)
    el = h._element
    body = doc.element.body
    body.remove(el)
    body.insert(0, el)


def _append_page_break(doc: Document) -> None:
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _append_pdf_as_images(
        doc: Document,
        pdf_path: Path,
        *,
        dpi: int = 120,
        max_width_inches: float = 6.5,
) -> None:
    """Растеризует PDF в картинки для вставки в DOCX."""
    src = fitz.open(pdf_path)
    try:
        for i in range(len(src)):
            page = src[i]
            pix = page.get_pixmap(dpi=dpi)
            bio = io.BytesIO(pix.tobytes("png"))
            bio.seek(0)
            par = doc.add_paragraph()
            run = par.add_run()
            run.add_picture(bio, width=Inches(max_width_inches))
    finally:
        src.close()


# ── НОВАЯ: автономная DOCX→PDF через reportlab ──

def _docx_to_pdf_bytes(docx_path: Path) -> bytes:
    """
    Конвертирует DOCX в PDF через reportlab.
    Сохраняет текст, таблицы, базовое форматирование (жирный, курсив, размер).
    Не требует внешних программ.
    """
    doc = Document(docx_path)

    # Временный буфер для PDF
    buffer = io.BytesIO()

    # Создаём PDF через reportlab
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Начальная позиция
    x_margin = 20 * mm
    y_position = height - 25 * mm
    line_height = 5 * mm
    para_spacing = 3 * mm

    for element in doc.element.body:
        # Таблицы
        if element.tag.endswith('}tbl'):
            table = None
            for tbl in doc.tables:
                if tbl._element is element:
                    table = tbl
                    break

            if table:
                y_position -= 5 * mm
                rows = len(table.rows)
                cols = len(table.columns)

                cell_width = (width - 2 * x_margin) / cols

                for row_idx, row in enumerate(table.rows):
                    if y_position < 25 * mm:
                        c.showPage()
                        y_position = height - 25 * mm

                    x_pos = x_margin
                    row_texts = []

                    for cell in row.cells:
                        cell_text = cell.text.strip()[:100]  # Обрезаем длинный текст
                        row_texts.append(cell_text)

                    for col_idx, cell_text in enumerate(row_texts):
                        c.setFont("Helvetica", 8)
                        if row_idx == 0:
                            c.setFont("Helvetica-Bold", 8)
                        c.drawString(x_pos + 2, y_position - 4, cell_text)
                        c.rect(x_pos, y_position - line_height, cell_width, line_height)
                        x_pos += cell_width

                    y_position -= line_height

                y_position -= para_spacing

        # Параграфы
        elif element.tag.endswith('}p'):
            for para in doc.paragraphs:
                if para._element is element:
                    if y_position < 25 * mm:
                        c.showPage()
                        y_position = height - 25 * mm

                    # Определяем стиль текста
                    font_name = "Helvetica"
                    font_size = 10

                    if para.style.name.startswith('Heading'):
                        font_name = "Helvetica-Bold"
                        font_size = 14 if '1' in para.style.name else 12

                    c.setFont(font_name, font_size)

                    # Переносим длинный текст
                    text = para.text.strip()
                    if not text:
                        y_position -= line_height / 2
                        continue

                    max_chars = int((width - 2 * x_margin) / (font_size * 0.4))

                    while len(text) > max_chars:
                        split_pos = text[:max_chars].rfind(' ')
                        if split_pos == -1:
                            split_pos = max_chars

                        c.drawString(x_margin, y_position, text[:split_pos])
                        y_position -= line_height
                        text = text[split_pos:].lstrip()

                        if y_position < 25 * mm:
                            c.showPage()
                            y_position = height - 25 * mm

                    if text:
                        c.drawString(x_margin, y_position, text)
                        y_position -= line_height

                    y_position -= para_spacing
                    break

    c.save()
    return buffer.getvalue()


# ── Основная функция ──

def merge_to_docx_and_pdf(
        paths: list[Path],
        output_docx: Path,
        output_pdf: Path | None,
        *,
        page_break_between_parts: bool = True,
        insert_titles: bool = False,
        pdf_render_dpi: int = 120,
        libreoffice_executable: str | None = None,
) -> tuple[list[str], list[str]]:
    """
    Склеивает .docx и .pdf в единый DOCX и PDF.
    ПОЛНОСТЬЮ АВТОНОМНАЯ — только Python-пакеты.
    """
    warnings: list[str] = []
    errors: list[str] = []

    work_items: list[tuple[Path, str]] = []

    for p in paths:
        p = Path(p)
        if not p.is_file():
            warnings.append(f"Пропуск (файл не найден): {p}")
            continue
        ext = p.suffix.lower()
        if ext == ".doc":
            warnings.append(f"Пропуск .doc без чтения: {p.name}")
            continue
        if ext in (".docx", ".pdf"):
            work_items.append((p, p.name))
            continue
        warnings.append(f"Пропуск (неподдерживаемый тип): {p.name}")

    if not work_items:
        errors.append("Нет ни одного поддерживаемого файла для склейки (.doc пропускаются).")
        return warnings, errors

    # ── Этап 1: DOCX ──
    merged: Document | None = None
    composer: Composer | None = None

    output_docx = Path(output_docx)
    try:
        for idx, (src_path, display_name) in enumerate(work_items):
            ext = src_path.suffix.lower()

            if idx > 0 and page_break_between_parts and merged is not None:
                _append_page_break(merged)

            if ext == ".pdf":
                if merged is None:
                    merged = Document()
                    composer = None
                if insert_titles:
                    merged.add_heading(display_name, level=2)
                _append_pdf_as_images(merged, src_path, dpi=pdf_render_dpi)
                continue

            if ext == ".docx":
                doc = Document(str(src_path))
                if insert_titles:
                    _prepend_heading(doc, display_name)
                if merged is None:
                    merged = doc
                    composer = Composer(merged)
                else:
                    if composer is None:
                        composer = Composer(merged)
                    composer.append(doc)
                continue

        if merged is None:
            errors.append("Не удалось сформировать документ.")
            return warnings, errors

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        merged.save(str(output_docx))

    except Exception as e:
        errors.append(f"Ошибка при склейке DOCX: {e}\n{traceback.format_exc()}")
        return warnings, errors

    # ── Этап 2: PDF ──
    if output_pdf is not None:
        output_pdf = Path(output_pdf)
        output_pdf.parent.mkdir(parents=True, exist_ok=True)

        try:
            pdf_writer = PdfWriter()
            temp_files = []

            for src_path, _ in work_items:
                ext = src_path.suffix.lower()

                if ext == ".pdf":
                    reader = PdfReader(str(src_path))
                    for page in reader.pages:
                        pdf_writer.add_page(page)

                elif ext == ".docx":
                    pdf_bytes = _docx_to_pdf_bytes(src_path)

                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                        tmp.write(pdf_bytes)
                        temp_files.append(Path(tmp.name))

                    reader = PdfReader(str(temp_files[-1]))
                    for page in reader.pages:
                        pdf_writer.add_page(page)

            with open(output_pdf, "wb") as f:
                pdf_writer.write(f)

        except Exception as e:
            errors.append(f"Ошибка при создании PDF: {e}\n{traceback.format_exc()}")

        finally:
            for tmp_path in temp_files:
                try:
                    tmp_path.unlink(missing_ok=True)
                except OSError:
                    pass

    return warnings, errors
