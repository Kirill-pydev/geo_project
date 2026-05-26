import os
import io
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxcompose.composer import Composer


def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    """Устанавливает поля ячейки таблицы в твипах (1/20 пункта)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')

    for name, value in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{name}')
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)

    existing_tcMar = tcPr.find(qn('w:tcMar'))
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

                # Сжатие шрифтов
                for p_element in section_element.xpath('.//w:p'):
                    try:
                        paragraph = doc.paragraphs[doc._paragraphs.index(p_element)]
                        if paragraph.style and paragraph.style.font and paragraph.style.font.size:
                            paragraph.style.font.size = Pt(paragraph.style.font.size.pt * 0.8)
                        for run in paragraph.runs:
                            if run.font and run.font.size:
                                run.font.size = Pt(run.font.size.pt * 0.8)
                            elif run.font:
                                run.font.size = Pt(9.5)
                    except:
                        pass

                # Сжатие таблиц
                for tbl_element in section_element.xpath('.//w:tbl'):
                    try:
                        table = doc.tables[doc._tables.index(tbl_element)]
                        scale_factor = 0.7
                        table.width = usable_width

                        tblGrid = table._tbl.find(qn('w:tblGrid'))
                        if tblGrid is not None:
                            for gridCol in tblGrid.findall(qn('w:gridCol')):
                                w_attr = gridCol.get(qn('w:w'))
                                if w_attr is not None:
                                    new_grid_w = int(int(w_attr) * scale_factor)
                                    gridCol.set(qn('w:w'), str(new_grid_w))

                        for row in table.rows:
                            trPr = row._tr.get_or_add_trPr()
                            trHeight = trPr.find(qn('w:trHeight'))
                            if trHeight is not None:
                                trPr.remove(trHeight)

                            for cell in row.cells:
                                tcPr = cell._tc.get_or_add_tcPr()
                                tcW = tcPr.find(qn('w:tcW'))
                                if tcW is not None:
                                    w_attr = tcW.get(qn('w:w'))
                                    if w_attr is not None and w_attr.isdigit():
                                        new_cell_w = int(int(w_attr) * scale_factor)
                                        tcW.set(qn('w:w'), str(new_cell_w))
                                        tcW.set(qn('w:type'), 'dxa')

                                try:
                                    set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
                                except:
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
                    except:
                        pass

                # Сжатие изображений
                for shape in doc.inline_shapes:
                    try:
                        w = getattr(shape, 'width', None)
                        h = getattr(shape, 'height', None)
                        if w is not None and h is not None and isinstance(w, (int, float)):
                            if w > usable_width:
                                ratio = usable_width / w
                                shape.width = int(usable_width)
                                shape.height = int(h * ratio)
                    except:
                        pass

        out_stream = io.BytesIO()
        doc.save(out_stream)
        out_stream.seek(0)
        return out_stream
    except Exception as e:
        print(f"Ошибка при масштабировании содержимого: {e}")
        return docx_stream


def convert_doc_to_docx_via_powershell(doc_input_path):
    """Качественная базовая конвертация .doc в .docx без изменения геометрии."""
    import tempfile
    import subprocess

    abs_input = os.path.abspath(doc_input_path)
    temp_dir = tempfile.gettempdir()
    base_name = os.path.splitext(os.path.basename(doc_input_path))[0]
    expected_docx = os.path.join(temp_dir, f"__conv_{base_name}.docx")

    if os.path.exists(expected_docx):
        try:
            os.remove(expected_docx)
        except:
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
            timeout=30
        )
        if not os.path.exists(expected_docx):
            return None
        with open(expected_docx, "rb") as f:
            mem_stream = io.BytesIO(f.read())
        return mem_stream
    except:
        return None
    finally:
        if os.path.exists(expected_docx):
            try:
                os.remove(expected_docx)
            except:
                pass


def try_convert_dwg_to_png(dwg_path):
    """
    Конвертирует DWG в PNG через доступные методы.
    Приоритет: ODA FileConverter → AutoCAD COM → FreeCAD → заглушка.
    Возвращает BytesIO с PNG или None.
    """
    import subprocess
    import tempfile

    temp_dir = tempfile.gettempdir()
    base_name = os.path.splitext(os.path.basename(dwg_path))[0]
    png_output = os.path.join(temp_dir, f"__dwg_{base_name}.png")

    # ── МЕТОД 1: ODA FileConverter ──
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
                    capture_output=True
                )
                if os.path.exists(png_output) and os.path.getsize(png_output) > 0:
                    png_buffer = io.BytesIO()
                    with open(png_output, "rb") as f:
                        png_buffer.write(f.read())
                    png_buffer.seek(0)
                    try:
                        os.remove(png_output)
                    except:
                        pass
                    return png_buffer
            except:
                continue

    # ── МЕТОД 2: AutoCAD через COM ──
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
                    except:
                        pass
                    pythoncom.CoUninitialize()
                    return png_buffer
        except:
            pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    except ImportError:
        pass

    # ── МЕТОД 3: FreeCAD ──
    try:
        import FreeCAD
        import importDXF

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
            except:
                pass
            return png_buffer
    except ImportError:
        pass

    return None


def insert_images_from_folder(doc, folder_path, max_width_inches=6.0):
    """
    Вставляет все изображения (JPG, PNG, BMP, GIF, TIFF) и DWG/DXF чертежи
    из папки в конец документа.
    DWG/DXF конвертируются в PNG через доступные методы.
    """
    from PIL import Image, ImageDraw, ImageFont

    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'}
    dwg_extensions = {'.dwg', '.dxf'}

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

    # Вставляем обычные изображения
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

    # Вставляем DWG/DXF как картинки
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
                # Вставляем сконвертированное изображение
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(png_buffer, width=Inches(max_width_inches))
            else:
                # Создаём информационную картинку-заглушку
                file_name = os.path.basename(dwg_path)
                file_size_kb = os.path.getsize(dwg_path) / 1024

                img_width = 800
                img_height = 600
                img = Image.new('RGB', (img_width, img_height), color=(245, 245, 245))
                draw = ImageDraw.Draw(img)

                draw.rectangle([10, 10, img_width - 11, img_height - 11], outline=(200, 200, 200), width=3)

                try:
                    font_title = ImageFont.truetype("arial.ttf", 28)
                    font_info = ImageFont.truetype("arial.ttf", 20)
                except:
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
                        draw.text((img_width / 2, y), line, fill=(50, 50, 50), font=font_title, anchor="mt")
                    else:
                        draw.text((img_width / 2, y), line, fill=(100, 100, 100), font=font_info, anchor="mt")
                    y += 40

                img_buffer = io.BytesIO()
                img.save(img_buffer, format='PNG')
                img_buffer.seek(0)

                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(img_buffer, width=Inches(max_width_inches))

        except Exception as e:
            doc.add_paragraph(f"[Ошибка обработки {os.path.basename(dwg_path)}: {e}]")


def main(input_dir=None, file_order=None, output_basename="merged_report"):
    if input_dir is None or file_order is None:
        print("Ошибка: Укажите входную папку и порядок файлов.")
        return

    input_dir = os.path.abspath(input_dir)
    downloads_path = Path.home() / "Downloads"
    output_dir = downloads_path / "result"
    output_file = output_dir / f"{output_basename}.docx"

    processed_mem_files = {}

    print("Шаг 1: Чтение файлов и сжатие альбомного контента...")
    for filename in os.listdir(input_dir):
        if filename.startswith("~$"):
            continue

        full_path = os.path.join(input_dir, filename)
        name_without_ext, ext = os.path.splitext(filename)
        ext = ext.lower()

        if ext not in ['.doc', '.docx']:
            continue

        if ext == '.doc':
            docx_stream = convert_doc_to_docx_via_powershell(full_path)
        else:
            with open(full_path, "rb") as f:
                docx_stream = io.BytesIO(f.read())

        if docx_stream:
            scaled_stream = scale_landscape_contents_to_portrait(docx_stream)
            processed_mem_files[name_without_ext] = scaled_stream

    missing_files = [name for name in file_order if name not in processed_mem_files]
    if missing_files:
        print(f"Ошибка: Не найдены файлы: {missing_files}")
        return

    print("Шаг 2: Сборка итогового документа из памяти...")
    output_dir.mkdir(parents=True, exist_ok=True)

    first_file_name = file_order[0]
    base_doc = Document(processed_mem_files[first_file_name])
    composer = Composer(base_doc)

    for file_name in file_order[1:]:
        base_doc.add_page_break()
        next_doc = Document(processed_mem_files[file_name])
        composer.append(next_doc)

    print("Шаг 3: Вставка изображений и чертежей из папки...")
    insert_images_from_folder(base_doc, input_dir)

    composer.save(output_file)
    print(f"\n[УСПЕХ] Документ успешно собран по адресу: {output_file}")


if __name__ == "__main__":
    custom_order = [
        "1.3 Содержание",
        "Б. ТЗ на геологию"
    ]
    source_folder = r"D:\Pycharm\Projects\GEO_DOCUMENTS\КРАСНАЯ ГОРКА"
    main(input_dir=source_folder, file_order=custom_order)