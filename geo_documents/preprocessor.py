import os
import shutil
import tempfile
import win32com.client
from pdf2docx import Converter


# --- Функция 1: Безопасный переворот из альбомной в книжную ---
def safe_convert_landscape_to_portrait(docx_input, docx_output):
    abs_input = os.path.abspath(docx_input)
    abs_output = os.path.abspath(docx_output)

    temp_dir = tempfile.gettempdir()
    temp_docx = os.path.join(temp_dir, "temp_proc.docx")
    temp_pdf = os.path.join(temp_dir, "temp_proc.pdf")

    # Копируем оригинал во временный файл для безопасности
    shutil.copy2(abs_input, temp_docx)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(temp_docx)
        for section in doc.Sections:
            section.PageSetup.Orientation = 0  # 0 = wdOrientPortrait (Книжная)
            for table in section.Range.Tables:
                table.AutoFitBehavior(2)  # 2 = wdAutoFitWindow (Подгонка таблиц)
            for shape in section.Range.InlineShapes:
                max_width = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin
                if shape.Width > max_width:
                    shape.LockAspectRatio = True
                    shape.Width = max_width
        doc.SaveAs(temp_pdf, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(SaveChanges=0)
    except Exception as e:
        print(f"Ошибка Word при перевороте {os.path.basename(docx_input)}: {e}")
        return False
    finally:
        word.Quit()

    try:
        cv = Converter(temp_pdf)
        cv.convert(abs_output, start=0, end=None)
        cv.close()
        return True
    except Exception as e:
        print(f"Ошибка pdf2docx для {os.path.basename(docx_input)}: {e}")
        return False
    finally:
        for temp_file in [temp_docx, temp_pdf]:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass


# --- Функция 2: Конвертация старого DOC в DOCX ---
def convert_doc_to_docx(doc_input, docx_output):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(doc_input), ReadOnly=True)
        doc.SaveAs2(os.path.abspath(docx_output), FileFormat=16)  # 16 = wdFormatXMLDocument (docx)
        doc.Close(SaveChanges=0)
        return True
    except Exception as e:
        print(f"Ошибка конвертации {os.path.basename(doc_input)} в DOCX: {e}")
        return False
    finally:
        word.Quit()


# --- ГЛАВНЫЙ АЛГОРИТМ ПРОЦЕССА ---
def main(input_dir=None, output_dir=None):
    """
    Препроцессор документов.
    Обрабатывает файлы ИЗ input_dir и сохраняет СРАЗУ в output_dir (data/).
    НИКАКОГО ПРОМЕЖУТОЧНОГО КОПИРОВАНИЯ ВСЕХ ФАЙЛОВ!
    """
    # Пути по умолчанию
    if input_dir is None:
        input_dir = "./source_files"
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

    input_dir = os.path.abspath(input_dir)
    output_dir = os.path.abspath(output_dir)

    print(f"[PREPROCESSOR] Input:  {input_dir}")
    print(f"[PREPROCESSOR] Output: {output_dir}")

    # Создаём выходную папку
    os.makedirs(output_dir, exist_ok=True)

    # Проверяем входную папку
    if not os.path.exists(input_dir):
        print(f"[PREPROCESSOR] ❌ Папка не найдена: {input_dir}")
        return False

    # Обработка файлов — СРАЗУ в output_dir
    for filename in os.listdir(input_dir):
        if filename.startswith("~$"):
            continue

        full_path = os.path.join(input_dir, filename)
        name_without_ext, ext = os.path.splitext(filename)
        ext = ext.lower()

        if ext not in ['.doc', '.docx']:
            continue

        print(f"[PREPROCESSOR] Обработка: {filename}")

        # Целевой путь в папке output (всегда .docx)
        target_docx = os.path.join(output_dir, f"{name_without_ext}.docx")

        if ext == '.doc':
            # Старый .doc → конвертируем через Word во временный, потом переворачиваем
            temp_docx = os.path.join(tempfile.gettempdir(), f"{name_without_ext}_converted.docx")
            if convert_doc_to_docx(full_path, temp_docx):
                success = safe_convert_landscape_to_portrait(temp_docx, target_docx)
                # Удаляем временный
                if os.path.exists(temp_docx):
                    os.remove(temp_docx)
            else:
                print(f"[PREPROCESSOR] ❌ Не удалось конвертировать DOC: {filename}")
                continue
        else:
            # Уже .docx → сразу переворачиваем
            success = safe_convert_landscape_to_portrait(full_path, target_docx)

        if success:
            print(f"[PREPROCESSOR] ✅ Сохранено: {target_docx}")
        else:
            print(f"[PREPROCESSOR] ❌ Ошибка: {filename}")

    print(f"[PREPROCESSOR] ✅ Готово! Файлы в: {output_dir}")
    return True


if __name__ == "__main__":
    main()