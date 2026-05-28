from __future__ import annotations

import io
import os
import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from docxcompose.composer import Composer

from geo_documents.gigachat import (
    DEFAULT_GIGACHAT_URL,
    DEFAULT_MAX_TOKENS,
    EXPLANATORY_NOTE_TITLE,
    build_explanatory_note_prompt,
    call_simple,
)
from geo_documents.libreoffice import doc_to_docx, docx_to_pdf, find_soffice


@dataclass
class MergeOptions:
    page_break_between_files: bool = True
    insert_file_titles: bool = False
    pdf_dpi: int = 150
    output_dir: Path | None = None
    soffice_path: str | None = None
    export_pdf: bool = True
    append_folder_media: bool = True
    generate_explanatory_note: bool = True
    gigachat_url: str = DEFAULT_GIGACHAT_URL
    gigachat_max_tokens: int = DEFAULT_MAX_TOKENS
    add_page_numbers: bool = True


@dataclass
class MergeResult:
    docx_path: Path
    pdf_path: Path | None = None
    explanatory_note_added: bool = False
    warnings: list[str] = field(default_factory=list)


def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    """Устанавливает поля ячейки таблицы в твипах (1/20 пункта)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")

    for name, value in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)

    existing_tcMar = tcPr.find(qn("w:tcMar"))
    if existing_tcMar is not None:
        tcPr.remove(existing_tcMar)

    tcPr.append(tcMar)


def scale_landscape_contents_to_portrait(docx_stream):
    """
    Находит альбомные секции, переводит их в книжную ориентацию
    и пропорционально уменьшает размер шрифта, таблиц и картинок,
    чтобы содержимое гарантированно вписалось в книжный лист.
    """
    try:
        doc = Document(docx_stream)

        for section in doc.sections:
            if section.orientation == WD_ORIENT.LANDSCAPE:
                new_width, new_height = section.page_height, section.page_width
                section.orientation = WD_ORIENT.PORTRAIT
                section.page_width = new_width
                section.page_height = new_height

                usable_width = section.page_width - section.left_margin - section.right_margin
                section_element = section._sectPr

                for p_element in section_element.xpath(".//w:p"):
                    try:
                        paragraph = doc.paragraphs[doc._paragraphs.index(p_element)]
                        if paragraph.style and paragraph.style.font and paragraph.style.font.size:
                            paragraph.style.font.size = Pt(paragraph.style.font.size.pt * 0.8)
                        for run in paragraph.runs:
                            if run.font and run.font.size:
                                run.font.size = Pt(run.font.size.pt * 0.8)
                            elif run.font:
                                run.font.size = Pt(9.5)
                    except Exception:
                        pass

                for tbl_element in section_element.xpath(".//w:tbl"):
                    try:
                        table = doc.tables[doc._tables.index(tbl_element)]
                        scale_factor = 0.7
                        table.width = usable_width

                        tblGrid = table._tbl.find(qn("w:tblGrid"))
                        if tblGrid is not None:
                            for gridCol in tblGrid.findall(qn("w:gridCol")):
                                w_attr = gridCol.get(qn("w:w"))
                                if w_attr is not None:
                                    new_grid_w = int(int(w_attr) * scale_factor)
                                    gridCol.set(qn("w:w"), str(new_grid_w))

                        for row in table.rows:
                            trPr = row._tr.get_or_add_trPr()
                            trHeight = trPr.find(qn("w:trHeight"))
                            if trHeight is not None:
                                trPr.remove(trHeight)

                            for cell in row.cells:
                                tcPr = cell._tc.get_or_add_tcPr()
                                tcW = tcPr.find(qn("w:tcW"))
                                if tcW is not None:
                                    w_attr = tcW.get(qn("w:w"))
                                    if w_attr is not None and w_attr.isdigit():
                                        new_cell_w = int(int(w_attr) * scale_factor)
                                        tcW.set(qn("w:w"), str(new_cell_w))
                                        tcW.set(qn("w:type"), "dxa")

                                try:
                                    set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
                                except Exception:
                                    pass

                                for p in cell.paragraphs:
                                    if p.paragraph_format:
                                        p.paragraph_format.space_after = Pt(0)
                                        p.paragraph_format.space_before = Pt(0)
                                        p.paragraph_format.line_spacing = 1.0
                                    for r in p.runs:
                                        if r.font and r.font.size:
                                            r.font.size = Pt(max(7.5, r.font.size.pt * 0.8))
                                        elif r.font:
                                            r.font.size = Pt(8.5)

                        table.autofit = True
                    except Exception:
                        pass

                for shape in doc.inline_shapes:
                    try:
                        w = getattr(shape, "width", None)
                        h = getattr(shape, "height", None)
                        if w is not None and h is not None and isinstance(w, (int, float)):
                            if w > usable_width:
                                ratio = usable_width / w
                                shape.width = int(usable_width)
                                shape.height = int(h * ratio)
                    except Exception:
                        pass

        out_stream = io.BytesIO()
        doc.save(out_stream)
        out_stream.seek(0)
        return out_stream
    except Exception as e:
        print(f"Ошибка при масштабировании содержимого: {e}")
        return docx_stream


def convert_doc_to_docx_via_powershell(doc_input_path: str | Path) -> io.BytesIO | None:
    """Конвертация .doc в .docx через Microsoft Word (COM), только Windows."""
    abs_input = os.path.abspath(str(doc_input_path))
    temp_dir = tempfile.gettempdir()
    base_name = os.path.splitext(os.path.basename(abs_input))[0]
    expected_docx = os.path.join(temp_dir, f"__conv_{base_name}.docx")

    if os.path.exists(expected_docx):
        try:
            os.remove(expected_docx)
        except OSError:
            pass

    ps_script = f"""
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try {{
        $doc = $word.Documents.Open('{abs_input}')
        $doc.SaveAs2('{expected_docx}', 16)
        $doc.Close()
    }} finally {{
        $word.Quit()
    }}
    """
    try:
        subprocess.run(
            ["powershell", "-Command", ps_script],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=30,
        )
        if not os.path.exists(expected_docx):
            return None
        with open(expected_docx, "rb") as f:
            mem_stream = io.BytesIO(f.read())
        mem_stream.seek(0)
        return mem_stream
    except Exception:
        return None
    finally:
        if os.path.exists(expected_docx):
            try:
                os.remove(expected_docx)
            except OSError:
                pass


def convert_doc_to_docx(doc_path: str | Path, soffice_path: str | None = None) -> io.BytesIO:
    """Конвертация .doc → .docx: сначала Word, затем LibreOffice."""
    stream = convert_doc_to_docx_via_powershell(doc_path)
    if stream is not None:
        return stream

    soffice = find_soffice(preferred=soffice_path)
    if soffice:
        docx_path = doc_to_docx(soffice, doc_path)
        try:
            with open(docx_path, "rb") as f:
                mem_stream = io.BytesIO(f.read())
            mem_stream.seek(0)
            return mem_stream
        finally:
            try:
                docx_path.unlink(missing_ok=True)
                docx_path.parent.rmdir()
            except OSError:
                pass

    raise RuntimeError(
        f"Не удалось конвертировать {Path(doc_path).name}: "
        "нужен Microsoft Word или LibreOffice (soffice.exe)."
    )


def pdf_to_docx_stream(pdf_path: str | Path, dpi: int = 150) -> io.BytesIO:
    """Растеризация PDF в DOCX через PyMuPDF."""
    doc = Document()
    pdf = fitz.open(str(pdf_path))
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    try:
        for page_num in range(len(pdf)):
            if page_num > 0:
                doc.add_page_break()
            page = pdf[page_num]
            pix = page.get_pixmap(matrix=matrix)
            img_buffer = io.BytesIO(pix.tobytes("png"))
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(img_buffer, width=Inches(6.0))
    finally:
        pdf.close()

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream


def prepare_docx_stream(path: Path, options: MergeOptions) -> io.BytesIO:
    """Подготавливает фрагмент DOCX из .doc, .docx или .pdf."""
    ext = path.suffix.lower()

    if ext == ".doc":
        docx_stream = convert_doc_to_docx(path, options.soffice_path)
    elif ext == ".docx":
        with open(path, "rb") as f:
            docx_stream = io.BytesIO(f.read())
        docx_stream.seek(0)
    elif ext == ".pdf":
        return pdf_to_docx_stream(path, dpi=options.pdf_dpi)
    else:
        raise ValueError(f"Неподдерживаемый формат: {path.name}")

    return scale_landscape_contents_to_portrait(docx_stream)


def _add_file_title(doc: Document, filename: str) -> None:
    heading = doc.add_paragraph()
    run = heading.add_run(filename)
    run.bold = True
    run.font.size = Pt(12)


def _extract_text_from_document(doc: Document) -> str:
    parts: list[str] = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            parts.append(text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text.strip()
                    if text:
                        parts.append(text)
    return "\n".join(parts)


def extract_text_from_path(path: Path, options: MergeOptions) -> str:
    """Извлекает текст из исходного файла для передачи в GigaChat."""
    ext = path.suffix.lower()

    if ext == ".pdf":
        with fitz.open(str(path)) as pdf:
            pages = [page.get_text().strip() for page in pdf]
        return "\n".join(part for part in pages if part)

    if ext == ".docx":
        return _extract_text_from_document(Document(str(path)))

    if ext == ".doc":
        docx_stream = convert_doc_to_docx(path, options.soffice_path)
        return _extract_text_from_document(Document(docx_stream))

    return ""


def collect_report_text(paths: list[Path], options: MergeOptions) -> tuple[str, list[str]]:
    """Собирает текст файлов с нумерованными источниками для ссылок в записке."""
    chunks: list[str] = []
    source_labels: list[str] = []

    for index, path in enumerate(paths, start=1):
        source_id = f"Источник {index}"
        label = f'{source_id}: «{path.name}»'
        source_labels.append(label)
        text = extract_text_from_path(path, options).strip()
        if text:
            chunks.append(f"--- {label} ---\n{text}")

    if not chunks:
        return "", source_labels

    registry = (
        "Реестр исходных материалов (единственный допустимый источник фактов):\n"
        + "\n".join(f"- {label}" for label in source_labels)
    )
    return registry + "\n\n" + "\n\n".join(chunks), source_labels


def _append_page_field(paragraph) -> None:
    """Вставляет поле PAGE (номер страницы) в абзац."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    placeholder = OxmlElement("w:t")
    placeholder.text = "1"

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(placeholder)
    run._r.append(fld_end)


def _clear_footer(footer) -> None:
    for paragraph in list(footer.paragraphs):
        element = paragraph._element
        element.getparent().remove(element)


def add_page_numbers_to_document(doc: Document) -> None:
    """Добавляет сквозную нумерацию страниц по центру нижнего колонтитула."""
    for index, section in enumerate(doc.sections):
        footer = section.footer
        footer.is_linked_to_previous = index > 0
        if index > 0:
            continue

        _clear_footer(footer)
        paragraph = footer.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _append_page_field(paragraph)


_CITATION_RE = re.compile(r"\[ист\.:[^\]]+\]", re.IGNORECASE)


def _add_text_with_source_citations(paragraph, text: str) -> None:
    """Вставляет текст; ссылки [ист.: …] выделяются курсивом."""
    pos = 0
    for match in _CITATION_RE.finditer(text):
        if match.start() > pos:
            paragraph.add_run(text[pos : match.start()])
        cite_run = paragraph.add_run(match.group())
        cite_run.italic = True
        pos = match.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def insert_explanatory_note(doc: Document, note_text: str) -> None:
    """Вставляет пояснительную записку перед приложениями с медиа."""
    doc.add_page_break()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(EXPLANATORY_NOTE_TITLE)
    title_run.bold = True
    title_run.font.size = Pt(14)

    doc.add_paragraph()

    section_re = re.compile(r"^(\d+\.\d+(?:\.\d+)?)\s+(.+)$")
    list_re = re.compile(r"^(\d+\)|[а-гa-d]\)|[а-гa-d]\.)\s*", re.IGNORECASE)

    for line in note_text.splitlines():
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph()
            continue

        if stripped.startswith("#"):
            stripped = stripped.lstrip("#").strip()

        if stripped == EXPLANATORY_NOTE_TITLE or stripped.startswith("1. Пояснительная записка"):
            continue

        paragraph = doc.add_paragraph()
        section_match = section_re.match(stripped)
        if section_match:
            run = paragraph.add_run(stripped)
            run.bold = True
            run.font.size = Pt(12)
        elif list_re.match(stripped):
            paragraph.paragraph_format.left_indent = Pt(18)
            _add_text_with_source_citations(paragraph, stripped)
        else:
            _add_text_with_source_citations(paragraph, stripped)


def generate_explanatory_note(paths: list[Path], options: MergeOptions) -> str:
    report_text, source_labels = collect_report_text(paths, options)
    if not report_text.strip():
        raise RuntimeError(
            "Не удалось извлечь текст из документов для пояснительной записки."
        )

    prompt = build_explanatory_note_prompt(report_text, source_labels)
    response = call_simple(
        prompt,
        url=options.gigachat_url,
        max_tokens=options.gigachat_max_tokens,
    )
    return response.content


def try_convert_dwg_to_png(dwg_path):
    """
    Конвертирует DWG в PNG через доступные методы.
    Приоритет: ODA FileConverter → AutoCAD COM → FreeCAD → заглушка.
    Возвращает BytesIO с PNG или None.
    """
    import subprocess

    temp_dir = tempfile.gettempdir()
    base_name = os.path.splitext(os.path.basename(dwg_path))[0]
    png_output = os.path.join(temp_dir, f"__dwg_{base_name}.png")

    odf_paths = [
        r"C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe",
        r"C:\Program Files (x86)\ODA\ODAFileConverter\ODAFileConverter.exe",
    ]

    for odf in odf_paths:
        if os.path.exists(odf):
            try:
                subprocess.run(
                    [odf, dwg_path, png_output, "PNG", "200", "1"],
                    timeout=60,
                    capture_output=True,
                )
                if os.path.exists(png_output) and os.path.getsize(png_output) > 0:
                    png_buffer = io.BytesIO()
                    with open(png_output, "rb") as f:
                        png_buffer.write(f.read())
                    png_buffer.seek(0)
                    try:
                        os.remove(png_output)
                    except OSError:
                        pass
                    return png_buffer
            except Exception:
                continue

    try:
        import pythoncom

        pythoncom.CoInitialize()
        try:
            from pyautocad import Autocad

            acad = Autocad(create_if_not_exists=True)
            if acad:
                doc = acad.Application.Documents.Open(dwg_path)
                doc.Export(png_output, "PNG")
                doc.Close(False)

                if os.path.exists(png_output) and os.path.getsize(png_output) > 0:
                    png_buffer = io.BytesIO()
                    with open(png_output, "rb") as f:
                        png_buffer.write(f.read())
                    png_buffer.seek(0)
                    try:
                        os.remove(png_output)
                    except OSError:
                        pass
                    return png_buffer
        except Exception:
            pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    except ImportError:
        pass

    try:
        import FreeCAD

        doc = FreeCAD.open(dwg_path)
        FreeCAD.Gui.SendMsgToActiveView("ViewFit")
        FreeCAD.Gui.activeDocument().activeView().saveImage(png_output, 800, 600, "White")
        FreeCAD.closeDocument(doc.Name)

        if os.path.exists(png_output) and os.path.getsize(png_output) > 0:
            png_buffer = io.BytesIO()
            with open(png_output, "rb") as f:
                png_buffer.write(f.read())
            png_buffer.seek(0)
            try:
                os.remove(png_output)
            except OSError:
                pass
            return png_buffer
    except ImportError:
        pass

    return None


def insert_images_from_folder(doc, folder_path, max_width_inches=6.0):
    """
    Вставляет все изображения (JPG, PNG, BMP, GIF, TIFF) и DWG/DXF чертежи
    из папки в конец документа.
    """
    from PIL import Image, ImageDraw, ImageFont

    image_extensions = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif"}
    dwg_extensions = {".dwg", ".dxf"}

    image_files = []
    dwg_files = []

    if not os.path.exists(folder_path):
        return

    for f in sorted(os.listdir(folder_path)):
        if f.startswith("~$"):
            continue
        ext = os.path.splitext(f)[1].lower()
        full_path = os.path.join(folder_path, f)

        if ext in image_extensions:
            image_files.append(full_path)
        elif ext in dwg_extensions:
            dwg_files.append(full_path)

    for img_path in image_files:
        try:
            doc.add_page_break()

            caption = doc.add_paragraph()
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = caption.add_run(f"Приложение: {os.path.basename(img_path)}")
            run.bold = True
            run.font.size = Pt(11)

            doc.add_paragraph()

            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(img_path, width=Inches(max_width_inches))

        except Exception as e:
            doc.add_paragraph(f"[Ошибка вставки {os.path.basename(img_path)}: {e}]")

    for dwg_path in dwg_files:
        try:
            doc.add_page_break()

            caption = doc.add_paragraph()
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = caption.add_run(f"Чертёж: {os.path.basename(dwg_path)}")
            run.bold = True
            run.font.size = Pt(11)

            doc.add_paragraph()

            png_buffer = try_convert_dwg_to_png(dwg_path)

            if png_buffer:
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(png_buffer, width=Inches(max_width_inches))
            else:
                file_name = os.path.basename(dwg_path)
                file_size_kb = os.path.getsize(dwg_path) / 1024

                img_width = 800
                img_height = 600
                img = Image.new("RGB", (img_width, img_height), color=(245, 245, 245))
                draw = ImageDraw.Draw(img)

                draw.rectangle(
                    [10, 10, img_width - 11, img_height - 11],
                    outline=(200, 200, 200),
                    width=3,
                )

                try:
                    font_title = ImageFont.truetype("arial.ttf", 28)
                    font_info = ImageFont.truetype("arial.ttf", 20)
                except OSError:
                    font_title = ImageFont.load_default()
                    font_info = ImageFont.load_default()

                lines = [
                    "Чертёж AutoCAD",
                    "",
                    f"Файл: {file_name}",
                    f"Размер: {file_size_kb:.1f} КБ",
                    "",
                    "Для отображения чертежа установите:",
                    "• ODA FileConverter",
                    "• или AutoCAD",
                    "• или FreeCAD",
                ]

                y = 180
                for i, line in enumerate(lines):
                    if i == 0:
                        draw.text(
                            (img_width / 2, y),
                            line,
                            fill=(50, 50, 50),
                            font=font_title,
                            anchor="mt",
                        )
                    else:
                        draw.text(
                            (img_width / 2, y),
                            line,
                            fill=(100, 100, 100),
                            font=font_info,
                            anchor="mt",
                        )
                    y += 40

                img_buffer = io.BytesIO()
                img.save(img_buffer, format="PNG")
                img_buffer.seek(0)

                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(img_buffer, width=Inches(max_width_inches))

        except Exception as e:
            doc.add_paragraph(f"[Ошибка обработки {os.path.basename(dwg_path)}: {e}]")


def merge_to_docx_and_pdf(
    ordered_paths: list[str | Path],
    output_basename: str = "merged_report",
    options: MergeOptions | None = None,
) -> MergeResult:
    """Склеивает документы в DOCX и при наличии LibreOffice экспортирует PDF."""
    if not ordered_paths:
        raise ValueError("Список файлов для склейки пуст.")

    options = options or MergeOptions()
    warnings: list[str] = []

    paths = [Path(p).resolve() for p in ordered_paths]
    for path in paths:
        if not path.is_file():
            raise FileNotFoundError(f"Файл не найден: {path}")

    input_dir = paths[0].parent
    output_dir = Path(options.output_dir or input_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    output_docx = output_dir / f"{output_basename}.docx"

    base_doc: Document | None = None
    composer: Composer | None = None

    for index, path in enumerate(paths):
        if index > 0 and options.page_break_between_files and base_doc is not None:
            base_doc.add_page_break()

        fragment_stream = prepare_docx_stream(path, options)
        fragment = Document(fragment_stream)

        if index == 0:
            if options.insert_file_titles:
                base_doc = Document()
                _add_file_title(base_doc, path.name)
                composer = Composer(base_doc)
                composer.append(fragment)
            else:
                base_doc = fragment
                composer = Composer(base_doc)
        else:
            if options.insert_file_titles:
                assert base_doc is not None
                _add_file_title(base_doc, path.name)
            assert composer is not None
            composer.append(fragment)

    assert base_doc is not None and composer is not None

    explanatory_note_added = False
    if options.generate_explanatory_note:
        try:
            note_text = generate_explanatory_note(paths, options)
            insert_explanatory_note(base_doc, note_text)
            explanatory_note_added = True
        except Exception as exc:
            warnings.append(f"Пояснительная записка не добавлена: {exc}")

    if options.append_folder_media:
        insert_images_from_folder(base_doc, str(input_dir))

    if options.add_page_numbers:
        add_page_numbers_to_document(base_doc)

    composer.save(output_docx)

    pdf_path: Path | None = None
    if options.export_pdf:
        soffice = find_soffice(preferred=options.soffice_path)
        if soffice:
            try:
                pdf_path = docx_to_pdf(soffice, output_docx, outdir=output_dir)
            except Exception as exc:
                warnings.append(f"Не удалось создать PDF: {exc}")
        else:
            warnings.append(
                "LibreOffice не найден — создан только DOCX. "
                "Укажите путь к soffice.exe или установите LibreOffice."
            )

    return MergeResult(
        docx_path=output_docx,
        pdf_path=pdf_path,
        explanatory_note_added=explanatory_note_added,
        warnings=warnings,
    )


def main(
    input_dir=None,
    file_order=None,
    output_basename="merged_report",
    ordered_paths: list[str | Path] | None = None,
    options: MergeOptions | None = None,
) -> MergeResult:
    """CLI-совместимая обёртка над merge_to_docx_and_pdf."""
    if ordered_paths is None:
        if input_dir is None or file_order is None:
            raise ValueError("Укажите ordered_paths или пару input_dir + file_order.")
        input_dir = Path(input_dir)
        ordered_paths = []
        for name in file_order:
            matched = None
            for ext in (".docx", ".doc", ".pdf"):
                candidate = input_dir / f"{name}{ext}"
                if candidate.is_file():
                    matched = candidate
                    break
            if matched is None:
                raise FileNotFoundError(f"Не найден файл для: {name}")
            ordered_paths.append(matched)

    opts = options or MergeOptions()
    if opts.output_dir is None and input_dir is not None:
        opts.output_dir = Path(input_dir)

    return merge_to_docx_and_pdf(ordered_paths, output_basename, opts)


if __name__ == "__main__":
    custom_order = [
        "1.3 Содержание",
        "Б. ТЗ на геологию",
    ]
    source_folder = r"D:\Pycharm\Projects\GEO_DOCUMENTS\КРАСНАЯ ГОРКА"
    result = main(input_dir=source_folder, file_order=custom_order)
    print(f"DOCX: {result.docx_path}")
    if result.pdf_path:
        print(f"PDF: {result.pdf_path}")
    for warning in result.warnings:
        print(f"Предупреждение: {warning}")
