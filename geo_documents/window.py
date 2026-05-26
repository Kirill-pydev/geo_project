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


class _MergeThread(QThread):
    """Вызывает merger.main() в отдельном потоке."""

    done = pyqtSignal(bool, str)
    crashed = pyqtSignal(str)

    def __init__(self, input_dir: str, file_order: list[str], output_basename: str) -> None:
        super().__init__()
        self._input_dir = input_dir
        self._file_order = file_order
        self._output_basename = output_basename

    def run(self) -> None:
        import traceback
        try:
            from geo_documents.merger import main as merger_main
            merger_main(
                input_dir=self._input_dir,
                file_order=self._file_order,
                output_basename=self._output_basename
            )
            self.done.emit(True, "Склейка успешно завершена!")
        except Exception as e:
            self.done.emit(False, str(e))
            self.crashed.emit(traceback.format_exc())


def _human_sort_key(name: str) -> str:
    return repr(sort_key_from_filename(name))


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Склейка отчётов (PDF / DOCX / DOC)")
        self.resize(880, 560)

        self._folder = Path.home()
        self._paths: list[Path] = []
        self._settings = QSettings("GEO_DOCUMENTS", "merge_app")
        self._merge_thread: _MergeThread | None = None

        self._downloads = Path.home() / "Downloads"
        self._result_dir = self._downloads / "result"

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
        fl.addRow(self.cb_page_break)
        fl.addRow(self.cb_titles)
        fl.addRow("Качество (DPI):", self.sp_dpi)
        root.addWidget(opts)

        row_out = QHBoxLayout()
        row_out.addWidget(QLabel("Имя без расширения:"))
        self.ed_basename = QLineEdit("merged_report")
        row_out.addWidget(self.ed_basename, stretch=1)
        root.addLayout(row_out)

        self.btn_merge = QPushButton("Склеить в DOCX и PDF")
        self.btn_merge.clicked.connect(self._merge)
        root.addWidget(self.btn_merge)

        hint = QLabel(
            "Поддерживаются файлы: .doc, .docx\n"
            ".doc — конвертация через PowerShell (Word)\n"
            ".docx — полное сохранение форматирования и автопереворот в книжную ориентацию\n"
            "Изображения (.jpg, .png, .bmp, .tiff) и чертежи (.dwg, .dxf) вставляются в конец документа\n"
            "Приложение полностью автономное, без LibreOffice."
        )
        hint.setWordWrap(True)
        root.addWidget(hint)

    def _pick_folder(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "Папка с документами", str(self._folder))
        if d:
            self._folder = Path(d)
            self.ed_folder.setText(str(self._folder))
            self._scan_folder()

    def _scan_folder(self) -> None:
        self._folder = Path(self.ed_folder.text().strip() or ".")
        self.ed_folder.setText(str(self._folder))
        if not self._folder.is_dir():
            QMessageBox.warning(self, "Папка", "Укажите существующую папку.")
            return
        found: list[Path] = []
        for name in os.listdir(self._folder):
            low = name.lower()
            if low.endswith((".docx", ".doc")):
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

    def _get_file_order(self) -> list[str]:
        """Имена файлов БЕЗ расширений в GUI-порядке."""
        return [p.stem for p in self._paths_from_list()]

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
            "Документы (*.doc *.docx);;Все файлы (*.*)",
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

    def _merge(self) -> None:
        if self._merge_thread and self._merge_thread.isRunning():
            QMessageBox.warning(self, "Занято", "Идёт склейка, подождите...")
            return

        paths = self._paths_from_list()
        if not paths:
            QMessageBox.information(self, "Склейка", "Список файлов пуст.")
            return

        file_order = self._get_file_order()
        input_dir = str(self._folder.resolve())
        basename = self.ed_basename.text().strip() or "merged_report"

        self.btn_merge.setEnabled(False)
        self.btn_merge.setText("Склейка...")

        self._merge_thread = _MergeThread(
            input_dir=input_dir,
            file_order=file_order,
            output_basename=basename,
        )
        self._merge_thread.done.connect(self._on_merge_done)
        self._merge_thread.crashed.connect(self._on_merge_crashed)
        self._merge_thread.finished.connect(self._merge_thread.deleteLater)
        self._merge_thread.start()

    def _on_merge_done(self, success: bool, message: str) -> None:
        basename = self.ed_basename.text().strip() or "merged_report"
        result_file = self._result_dir / f"{basename}.docx"

        if success:
            QMessageBox.information(
                self, "Готово",
                f"{message}\n\nФайл сохранён:\n{result_file}"
            )
        else:
            QMessageBox.critical(self, "Ошибка", message)

        self.btn_merge.setEnabled(True)
        self.btn_merge.setText("Склеить в DOCX и PDF")

    def _on_merge_crashed(self, tb: str) -> None:
        QMessageBox.critical(self, "Сбой при склейке", tb)
        self.btn_merge.setEnabled(True)
        self.btn_merge.setText("Склеить в DOCX и PDF")