from __future__ import annotations

import os
from pathlib import Path

from PyQt6.QtCore import QSettings, Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from geo_documents.file_sorter import sort_key_from_filename, sorted_paths
from geo_documents.libreoffice import find_soffice
from geo_documents.merger import MergeOptions, merge_to_docx_and_pdf


class _MergeThread(QThread):
    """Вызывает merge_to_docx_and_pdf в отдельном потоке."""

    done = pyqtSignal(bool, str)
    crashed = pyqtSignal(str)

    def __init__(
        self,
        ordered_paths: list[str],
        output_basename: str,
        options: MergeOptions,
    ) -> None:
        super().__init__()
        self._ordered_paths = ordered_paths
        self._output_basename = output_basename
        self._options = options

    def run(self) -> None:
        import traceback

        try:
            result = merge_to_docx_and_pdf(
                ordered_paths=self._ordered_paths,
                output_basename=self._output_basename,
                options=self._options,
            )
            lines = ["Склейка успешно завершена!", "", f"DOCX:\n{result.docx_path}"]
            if result.explanatory_note_added:
                lines.append("\nПояснительная записка добавлена в документ.")
            if result.pdf_path:
                lines.extend(["", f"PDF:\n{result.pdf_path}"])
            for warning in result.warnings:
                lines.extend(["", f"⚠ {warning}"])
            self.done.emit(True, "\n".join(lines))
        except Exception as e:
            self.done.emit(False, str(e))
            self.crashed.emit(traceback.format_exc())


def _human_sort_key(name: str) -> str:
    return repr(sort_key_from_filename(name))


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Склейка отчётов (PDF / DOCX / DOC)")
        self.resize(880, 620)

        self._folder = Path.home()
        self._paths: list[Path] = []
        self._settings = QSettings("GEO_DOCUMENTS", "merge_app")
        self._merge_thread: _MergeThread | None = None

        root = QVBoxLayout(self)

        row1 = QHBoxLayout()
        self.ed_folder = QLineEdit(str(self._folder))
        btn_browse = QPushButton("Папка…")
        btn_browse.clicked.connect(self._pick_folder)
        btn_scan = QPushButton("Обновить список")
        btn_scan.clicked.connect(self._scan_folder)
        row1.addWidget(QLabel("Папка с файлами:"))
        row1.addWidget(self.ed_folder, stretch=1)
        row1.addWidget(btn_browse)
        row1.addWidget(btn_scan)
        root.addLayout(row1)

        self.list_w = QListWidget()
        self.list_w.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.list_w.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.list_w.setDefaultDropAction(Qt.DropAction.MoveAction)
        root.addWidget(self.list_w, stretch=1)

        row_btns = QHBoxLayout()
        self.btn_sort = QPushButton("Автосортировка")
        self.btn_sort.clicked.connect(self._autosort)
        self.btn_up = QPushButton("Вверх")
        self.btn_up.clicked.connect(lambda: self._move_selected(-1))
        self.btn_down = QPushButton("Вниз")
        self.btn_down.clicked.connect(lambda: self._move_selected(1))
        self.btn_remove = QPushButton("Убрать из списка")
        self.btn_remove.clicked.connect(self._remove_selected)
        self.btn_add = QPushButton("Добавить файлы…")
        self.btn_add.clicked.connect(self._add_files)
        row_btns.addWidget(self.btn_sort)
        row_btns.addWidget(self.btn_up)
        row_btns.addWidget(self.btn_down)
        row_btns.addWidget(self.btn_remove)
        row_btns.addWidget(self.btn_add)
        row_btns.addStretch(1)
        root.addLayout(row_btns)

        opts = QGroupBox("Параметры склейки")
        fl = QFormLayout(opts)
        self.cb_page_break = QCheckBox("Разрыв страницы между файлами")
        self.cb_page_break.setChecked(True)
        self.cb_titles = QCheckBox("Вставлять заголовок с именем файла")
        self.cb_titles.setChecked(False)
        self.sp_dpi = QSpinBox()
        self.sp_dpi.setRange(72, 300)
        self.sp_dpi.setValue(150)
        self.sp_dpi.setSuffix(" dpi")
        self.cb_export_pdf = QCheckBox("Экспортировать PDF через LibreOffice")
        self.cb_export_pdf.setChecked(True)
        self.cb_explanatory_note = QCheckBox("Пояснительная записка (GigaChat)")
        self.cb_explanatory_note.setChecked(True)
        self.cb_page_numbers = QCheckBox("Нумерация страниц (внизу по центру)")
        self.cb_page_numbers.setChecked(True)
        fl.addRow(self.cb_page_break)
        fl.addRow(self.cb_titles)
        fl.addRow("Качество PDF (DPI):", self.sp_dpi)
        fl.addRow(self.cb_export_pdf)
        fl.addRow(self.cb_explanatory_note)
        fl.addRow(self.cb_page_numbers)
        root.addWidget(opts)

        row_soffice = QHBoxLayout()
        row_soffice.addWidget(QLabel("LibreOffice (soffice.exe):"))
        self.ed_soffice = QLineEdit()
        saved_soffice = self._settings.value("soffice_path", "", str)
        detected = find_soffice(preferred=saved_soffice or None)
        self.ed_soffice.setText(saved_soffice or (detected or ""))
        self.ed_soffice.setPlaceholderText("Автопоиск или путь к soffice.exe")
        btn_soffice = QPushButton("Обзор…")
        btn_soffice.clicked.connect(self._pick_soffice)
        row_soffice.addWidget(self.ed_soffice, stretch=1)
        row_soffice.addWidget(btn_soffice)
        root.addLayout(row_soffice)

        row_out = QHBoxLayout()
        row_out.addWidget(QLabel("Имя без расширения:"))
        self.ed_basename = QLineEdit("merged_report")
        row_out.addWidget(self.ed_basename, stretch=1)
        root.addLayout(row_out)

        self.btn_merge = QPushButton("Склеить в DOCX и PDF")
        self.btn_merge.clicked.connect(self._merge)
        root.addWidget(self.btn_merge)

        hint = QLabel(
            "Поддерживаются файлы: .pdf, .doc, .docx\n"
            "• PDF — растеризация страниц (PyMuPDF), качество задаётся DPI\n"
            "• DOC — конвертация через Word или LibreOffice\n"
            "• DOCX — сохранение форматирования и автопереворот альбомных страниц\n"
            "Пояснительная записка (выжимка) формируется через GigaChat и вставляется "
            "перед изображениями и чертежами\n"
            "Изображения и чертежи (.jpg, .png, .dwg, .dxf) из папки добавляются в конец\n"
            "Результат сохраняется в выбранной папке: имя.docx и имя.pdf"
        )
        hint.setWordWrap(True)
        root.addWidget(hint)

    def _pick_folder(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "Папка с документами", str(self._folder))
        if d:
            self._folder = Path(d)
            self.ed_folder.setText(str(self._folder))
            self._scan_folder()

    def _pick_soffice(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "LibreOffice soffice.exe",
            self.ed_soffice.text() or "C:\\Program Files",
            "Исполняемые файлы (soffice.exe);;Все файлы (*.*)",
        )
        if path:
            self.ed_soffice.setText(path)

    def _scan_folder(self) -> None:
        self._folder = Path(self.ed_folder.text().strip() or ".")
        self.ed_folder.setText(str(self._folder))
        if not self._folder.is_dir():
            QMessageBox.warning(self, "Папка", "Укажите существующую папку.")
            return
        found: list[Path] = []
        for name in os.listdir(self._folder):
            low = name.lower()
            if low.endswith((".pdf", ".docx", ".doc")):
                if low.startswith("~$"):
                    continue
                found.append(self._folder / name)
        self._paths = sorted_paths(found)
        self._fill_list()

    def _fill_list(self) -> None:
        self.list_w.clear()
        for p in self._paths:
            item = QListWidgetItem(f"{p.name}   [{_human_sort_key(p.name)}]")
            item.setData(Qt.ItemDataRole.UserRole, str(p.resolve()))
            self.list_w.addItem(item)

    def _paths_from_list(self) -> list[Path]:
        out: list[Path] = []
        for i in range(self.list_w.count()):
            it = self.list_w.item(i)
            data = it.data(Qt.ItemDataRole.UserRole)
            if data:
                out.append(Path(str(data)))
        return out

    def _autosort(self) -> None:
        self._paths = sorted_paths(self._paths_from_list())
        self._fill_list()

    def _move_selected(self, delta: int) -> None:
        row = self.list_w.currentRow()
        if row < 0:
            return
        new_row = row + delta
        if new_row < 0 or new_row >= self.list_w.count():
            return
        item = self.list_w.takeItem(row)
        self.list_w.insertItem(new_row, item)
        self.list_w.setCurrentRow(new_row)

    def _remove_selected(self) -> None:
        for item in self.list_w.selectedItems():
            self.list_w.takeItem(self.list_w.row(item))

    def _add_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Добавить файлы",
            str(self._folder),
            "Документы (*.pdf *.doc *.docx);;Все файлы (*.*)",
        )
        if not files:
            return
        existing = {str(Path(x).resolve()) for x in self._paths_from_list()}
        for f in files:
            fp = Path(f).resolve()
            key = str(fp)
            if key in existing:
                continue
            existing.add(key)
            it = QListWidgetItem(f"{fp.name}   [{_human_sort_key(fp.name)}]")
            it.setData(Qt.ItemDataRole.UserRole, key)
            self.list_w.addItem(it)

    def _build_merge_options(self) -> MergeOptions:
        soffice_text = self.ed_soffice.text().strip()
        self._settings.setValue("soffice_path", soffice_text)
        return MergeOptions(
            page_break_between_files=self.cb_page_break.isChecked(),
            insert_file_titles=self.cb_titles.isChecked(),
            pdf_dpi=self.sp_dpi.value(),
            output_dir=self._folder.resolve(),
            soffice_path=soffice_text or None,
            export_pdf=self.cb_export_pdf.isChecked(),
            generate_explanatory_note=self.cb_explanatory_note.isChecked(),
            add_page_numbers=self.cb_page_numbers.isChecked(),
        )

    def _merge(self) -> None:
        if self._merge_thread and self._merge_thread.isRunning():
            QMessageBox.warning(self, "Занято", "Идёт склейка, подождите...")
            return

        paths = self._paths_from_list()
        if not paths:
            QMessageBox.information(self, "Склейка", "Список файлов пуст.")
            return

        basename = self.ed_basename.text().strip() or "merged_report"
        options = self._build_merge_options()

        self.btn_merge.setEnabled(False)
        self.btn_merge.setText("Склейка...")

        self._merge_thread = _MergeThread(
            ordered_paths=[str(p) for p in paths],
            output_basename=basename,
            options=options,
        )
        self._merge_thread.done.connect(self._on_merge_done)
        self._merge_thread.crashed.connect(self._on_merge_crashed)
        self._merge_thread.finished.connect(self._merge_thread.deleteLater)
        self._merge_thread.start()

    def _on_merge_done(self, success: bool, message: str) -> None:
        if success:
            QMessageBox.information(self, "Готово", message)
        else:
            QMessageBox.critical(self, "Ошибка", message)

        self.btn_merge.setEnabled(True)
        self.btn_merge.setText("Склеить в DOCX и PDF")

    def _on_merge_crashed(self, tb: str) -> None:
        QMessageBox.critical(self, "Сбой при склейке", tb)
        self.btn_merge.setEnabled(True)
        self.btn_merge.setText("Склеить в DOCX и PDF")
