from __future__ import annotations

import os
import shutil
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


class _PreprocessThread(QThread):
    """Вызывает preprocessor.main() в отдельном потоке."""

    done = pyqtSignal(bool, str)  # success, message
    progress = pyqtSignal(str)
    crashed = pyqtSignal(str)

    def __init__(self, input_dir: str, output_dir: str) -> None:
        super().__init__()
        self._input_dir = input_dir
        self._output_dir = output_dir

    def run(self) -> None:
        import traceback
        try:
            from geo_documents.preprocessor import main as preprocess_main
            preprocess_main(input_dir=self._input_dir, output_dir=self._output_dir)
            self.done.emit(True, "Подготовка завершена")
        except Exception as e:
            self.done.emit(False, str(e))
            self.crashed.emit(traceback.format_exc())


class _MergeThread(QThread):
    """Вызывает merger.py в отдельном потоке."""

    done = pyqtSignal(bool, list, list, str, str)
    crashed = pyqtSignal(str)

    def __init__(
            self,
            *,
            data_folder: Path,
            result_folder: Path,
            basename: str,
            page_break: bool,
            insert_titles: bool,
            dpi: int,
    ) -> None:
        super().__init__()
        self._data_folder = data_folder
        self._result_folder = result_folder
        self._basename = basename
        self._page_break = page_break
        self._insert_titles = insert_titles
        self._dpi = dpi

    def run(self) -> None:
        import traceback
        try:
            from geo_documents.merger import merge_to_docx_and_pdf
            from geo_documents.file_sorter import sorted_paths  # ВОТ ОНО, БЛЯДЬ!

            self._result_folder.mkdir(parents=True, exist_ok=True)

            # Используем sorted_paths вместо sorted()
            paths = sorted_paths([
                f for f in self._data_folder.glob("*.docx")
                if not f.name.startswith("~$")
            ])

            out_docx = self._result_folder / f"{self._basename}.docx"
            out_pdf = self._result_folder / f"{self._basename}.pdf"

            warnings, errors = merge_to_docx_and_pdf(
                paths,
                out_docx,
                out_pdf,
                page_break_between_parts=self._page_break,
                insert_titles=self._insert_titles,
                pdf_render_dpi=self._dpi,
            )

            success = out_docx.exists() and len(errors) == 0

            # Удаляем папку data после успешной склейки
            if self._data_folder.exists():
                shutil.rmtree(self._data_folder, ignore_errors=True)

            self.done.emit(success, warnings, errors, str(out_docx), str(out_pdf))
        except Exception:
            self.crashed.emit(traceback.format_exc())


def _human_sort_key(name: str) -> str:
    return repr(sort_key_from_filename(name))


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Склейка отчётов (PDF / DOCX / DOC)")
        self.resize(880, 560)

        self._folder = Path.home() / "Downloads"  # По умолчанию Загрузки
        self._paths: list[Path] = []
        self._settings = QSettings("GEO_DOCUMENTS", "merge_app")
        self._preprocess_thread: _PreprocessThread | None = None
        self._merge_thread: _MergeThread | None = None

        # Стандартные пути
        self._downloads = Path.home() / "Downloads"
        self._data_dir = self._downloads / "data"
        self._result_dir = self._downloads / "result"

        root = QVBoxLayout(self)

        row1 = QHBoxLayout()
        self.ed_folder = QLineEdit(str(self._folder))
        btn_browse = QPushButton("Папка…")
        btn_browse.clicked.connect(self._pick_folder)
        btn_scan = QPushButton("Обновить список")
        btn_scan.clicked.connect(self._scan_folder)
        row1.addWidget(QLabel("Папка с исходными файлами:"))
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

        self.status_label = QLabel("Готов к работе")
        self.status_label.setStyleSheet("color: #555; font-weight: bold;")
        root.addWidget(self.status_label)

        hint = QLabel(
            "1. Выберите папку с исходными .doc/.docx файлами\n"
            "2. Нажмите «Склеить»\n"
            "3. Подготовленные файлы → Загрузки/data/\n"
            "4. Результат склейки → Загрузки/result/\n"
            "5. Папка data/ удаляется автоматически"
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
            "Документы (*.pdf *.docx *.doc);;Все файлы (*.*)",
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
        if self._preprocess_thread and self._preprocess_thread.isRunning():
            QMessageBox.warning(self, "Занято", "Идёт подготовка файлов...")
            return
        if self._merge_thread and self._merge_thread.isRunning():
            QMessageBox.warning(self, "Занято", "Идёт склейка...")
            return

        paths = self._paths_from_list()
        if not paths:
            QMessageBox.information(self, "Склейка", "Список файлов пуст.")
            return

        # Очищаем старую data/ если есть
        if self._data_dir.exists():
            shutil.rmtree(self._data_dir, ignore_errors=True)

        # Берём исходную папку НАПРЯМУЮ — preprocessor сам разберётся
        input_dir = str(self._folder.resolve())

        self.btn_merge.setEnabled(False)
        self.status_label.setText("⚙ Шаг 1/2: Подготовка файлов...")

        # Шаг 1: preprocessor.py (читает из input_dir, сохраняет в data_dir)
        self._preprocess_thread = _PreprocessThread(
            input_dir=input_dir,
            output_dir=str(self._data_dir)
        )
        self._preprocess_thread.done.connect(self._on_preprocess_done)
        self._preprocess_thread.crashed.connect(self._on_preprocess_crashed)
        self._preprocess_thread.finished.connect(self._preprocess_thread.deleteLater)
        self._preprocess_thread.start()

    def _on_preprocess_done(self, success: bool, message: str) -> None:
        if not success:
            self.status_label.setText("❌ Ошибка подготовки")
            QMessageBox.critical(self, "Ошибка", f"Препроцессор:\n{message}")
            self.btn_merge.setEnabled(True)
            return

        self.status_label.setText("✅ Шаг 1/2 готов. Шаг 2/2: Склейка...")

        # Шаг 2: merger.py (читает из data_dir, сохраняет в result_dir)
        self._merge_thread = _MergeThread(
            data_folder=self._data_dir,
            result_folder=self._result_dir,
            basename=self.ed_basename.text().strip() or "merged_report",
            page_break=self.cb_page_break.isChecked(),
            insert_titles=self.cb_titles.isChecked(),
            dpi=self.sp_dpi.value(),
        )
        self._merge_thread.done.connect(self._on_merge_done)
        self._merge_thread.crashed.connect(self._on_merge_crashed)
        self._merge_thread.finished.connect(self._merge_thread.deleteLater)
        self._merge_thread.start()

    def _on_preprocess_crashed(self, tb: str) -> None:
        QMessageBox.critical(self, "Сбой", f"Ошибка препроцессора:\n\n{tb}")
        self.status_label.setText("❌ Сбой")
        self.btn_merge.setEnabled(True)

    def _on_merge_done(self, success: bool, warnings: list, errors: list, docx_path: str, pdf_path: str) -> None:
        msg_lines = ["✅ Готово!"]

        if Path(docx_path).exists():
            msg_lines.append(f"📄 DOCX: {docx_path}")
        if Path(pdf_path).exists():
            msg_lines.append(f"📄 PDF: {pdf_path}")

        if warnings:
            msg_lines.append("\n⚠ Предупреждения:\n- " + "\n- ".join(warnings))
        if errors:
            msg_lines.append("\n❌ Ошибки:\n- " + "\n- ".join(errors))

        QMessageBox.information(self, "Результат", "\n".join(msg_lines))
        self.status_label.setText(f"✅ Готово! Результат в {self._result_dir}")
        self.btn_merge.setEnabled(True)

    def _on_merge_crashed(self, tb: str) -> None:
        QMessageBox.critical(self, "Сбой склейки", tb)
        self.status_label.setText("❌ Сбой")
        self.btn_merge.setEnabled(True)
