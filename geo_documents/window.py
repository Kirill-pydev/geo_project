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
from geo_documents.merger import merge_to_docx_and_pdf  # ← ИСПРАВЛЕНО


class _MergeThread(QThread):
    """Склейка в отдельном потоке, чтобы окно не уходило в «Не отвечает» и не закрывалось ОС."""

    done = pyqtSignal(list, list)
    crashed = pyqtSignal(str)

    def __init__(
            self,
            *,
            paths: list[Path],
            out_docx: Path,
            out_pdf: Path,
            page_break: bool,
            insert_titles: bool,
            dpi: int,
            lo_path: str | None,
    ) -> None:
        super().__init__()
        self._paths = paths
        self._out_docx = out_docx
        self._out_pdf = out_pdf
        self._page_break = page_break
        self._insert_titles = insert_titles
        self._dpi = dpi
        self._lo_path = lo_path

    def run(self) -> None:
        import traceback

        try:
            warnings, errors = merge_to_docx_and_pdf(
                self._paths,
                self._out_docx,
                self._out_pdf,
                page_break_between_parts=self._page_break,
                insert_titles=self._insert_titles,
                pdf_render_dpi=self._dpi,
                libreoffice_executable=self._lo_path,
            )
            self.done.emit(warnings, errors)
        except Exception:
            self.crashed.emit(traceback.format_exc())


def _human_sort_key(name: str) -> str:
    return repr(sort_key_from_filename(name))


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Склейка отчётов (PDF / DOCX)")
        self.resize(880, 560)

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
        fl.addRow(self.cb_page_break)
        fl.addRow(self.cb_titles)
        fl.addRow("Растр PDF в DOCX:", self.sp_dpi)
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
            "PDF создаётся автономно. "
            "Альбомные страницы автоматически поворачиваются в книжные.\n"
            "Файлы .doc не читаются и пропускаются. "
            "Для качественного PDF положите wkhtmltopdf.exe в папку bin/"
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
            if low.endswith((".pdf", ".docx")):
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
            "Документы (*.pdf *.docx);;Все файлы (*.*)",
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
        if self._merge_thread is not None and self._merge_thread.isRunning():
            return
        paths = self._paths_from_list()
        if not paths:
            QMessageBox.warning(self, "Склейка", "Список файлов пуст.")
            return
        base = self.ed_basename.text().strip() or "merged_report"
        out_dir = self._folder if self._folder.is_dir() else Path.cwd()
        out_docx = out_dir / f"{base}.docx"
        out_pdf = out_dir / f"{base}.pdf"

        th = _MergeThread(
            paths=paths,
            out_docx=out_docx,
            out_pdf=out_pdf,
            page_break=self.cb_page_break.isChecked(),
            insert_titles=self.cb_titles.isChecked(),
            dpi=int(self.sp_dpi.value()),
            lo_path=None,
        )
        self._merge_thread = th
        th.done.connect(self._on_merge_done)
        th.crashed.connect(self._on_merge_crashed)
        th.finished.connect(th.deleteLater)
        self.btn_merge.setEnabled(False)
        th.start()

    def _merge_ui_unlock(self) -> None:
        self.btn_merge.setEnabled(True)
        self._merge_thread = None

    def _on_merge_done(self, warnings: list[str], errors: list[str]) -> None:
        msg_lines: list[str] = []
        base = self.ed_basename.text().strip() or "merged_report"
        out_dir = self._folder if self._folder.is_dir() else Path.cwd()
        out_docx = out_dir / f"{base}.docx"
        out_pdf = out_dir / f"{base}.pdf"
        if out_docx.is_file() or out_pdf.is_file():
            msg_lines.append("Сохранено:")
        if out_docx.is_file():
            msg_lines.append(f"DOCX: {out_docx}")
        if out_pdf.is_file():
            msg_lines.append(f"PDF: {out_pdf}")
        if warnings:
            msg_lines.append("\nПредупреждения:\n- " + "\n- ".join(warnings))
        if errors:
            msg_lines.append("\nОшибки:\n- " + "\n- ".join(errors))
            QMessageBox.warning(self, "Склейка завершена с ошибками", "\n".join(msg_lines))
        else:
            QMessageBox.information(self, "Готово", "\n".join(msg_lines) or "Готово.")
        self._merge_ui_unlock()

    def _on_merge_crashed(self, tb: str) -> None:
        QMessageBox.critical(
            self,
            "Сбой при склейке",
            "Произошла внутренняя ошибка (подробности ниже).\n\n" + tb,
        )
        self._merge_ui_unlock()
