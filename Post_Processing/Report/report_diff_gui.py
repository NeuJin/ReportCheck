#!/usr/bin/env python3
"""GUI for PPT report visual comparison.

Python version: 3.9+
Requires PySide2, Pillow, pywin32, and python-pptx from the bundled environment.
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path
from types import SimpleNamespace
from typing import List, Optional

from PySide2.QtCore import QObject, QSize, Qt, QThread, Signal
from PySide2.QtGui import QPixmap
from PySide2.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import pixel_report_diff as engine


class CompareWorker(QObject):
    finished = Signal(object, object, str)
    failed = Signal(str)

    def __init__(self, args: SimpleNamespace):
        super().__init__()
        self.args = args

    def run(self) -> None:
        try:
            engine.validate_inputs(self.args)
            output_dir = engine.resolve_output_dir(self.args.output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)

            pixel_results = []
            object_diffs = []
            if self.args.mode in ("pixel", "both"):
                pixel_results = engine.compare_pixels(self.args, output_dir)
            if self.args.mode in ("object", "both"):
                object_diffs = engine.compare_objects(self.args, output_dir)

            engine.write_summary(output_dir, pixel_results, object_diffs)
            self.finished.emit(pixel_results, object_diffs, str(output_dir))
        except Exception:
            self.failed.emit(traceback.format_exc())


class ImagePreview(QScrollArea):
    def __init__(self):
        super().__init__()
        self.label = QLabel("No slide selected")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setMinimumSize(QSize(640, 360))
        self.label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setWidget(self.label)
        self.setWidgetResizable(True)
        self._pixmap: Optional[QPixmap] = None

    def set_image(self, image_path: str) -> None:
        pixmap = QPixmap(image_path)
        if pixmap.isNull():
            self._pixmap = None
            self.label.setText("Cannot load image")
            return
        self._pixmap = pixmap
        self._fit_pixmap()

    def resizeEvent(self, event):  # type: ignore[override]
        super().resizeEvent(event)
        self._fit_pixmap()

    def _fit_pixmap(self) -> None:
        if self._pixmap is None:
            return
        size = self.viewport().size()
        scaled = self._pixmap.scaled(size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.label.setPixmap(scaled)


class ReportDiffWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ReportCheck - PPT Slide Diff")
        self.resize(1200, 760)

        self.pixel_results = []
        self.object_diffs = []
        self.output_dir = ""
        self.worker_thread: Optional[QThread] = None

        self.expected_edit = QLineEdit()
        self.actual_edit = QLineEdit()
        self.output_edit = QLineEdit(engine.DEFAULT_OUTPUT_DIR)
        self.output_edit.setToolTip("Relative folder inside this downloaded package only")

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["both", "pixel", "object"])
        self.threshold_spin = QSpinBox()
        self.threshold_spin.setRange(0, 255)
        self.threshold_spin.setValue(3)
        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(50, 300)
        self.dpi_spin.setValue(150)
        self.allowed_spin = QSpinBox()
        self.allowed_spin.setRange(0, 100)
        self.allowed_spin.setValue(0)

        self.only_diff_check = QCheckBox("Only different slides")
        self.only_diff_check.setChecked(True)
        self.only_diff_check.stateChanged.connect(self.populate_slide_list)

        self.run_button = QPushButton("Run compare")
        self.run_button.clicked.connect(self.run_compare)

        self.slide_list = QListWidget()
        self.slide_list.currentItemChanged.connect(self.slide_selected)

        self.preview = ImagePreview()
        self.detail_box = QTextEdit()
        self.detail_box.setReadOnly(True)
        self.detail_box.setMinimumHeight(140)

        self._build_layout()
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

    def _build_layout(self) -> None:
        expected_button = QPushButton("Browse")
        actual_button = QPushButton("Browse")
        expected_button.clicked.connect(lambda: self.pick_file(self.expected_edit))
        actual_button.clicked.connect(lambda: self.pick_file(self.actual_edit))

        form = QFormLayout()
        form.addRow("Expected report", self._path_row(self.expected_edit, expected_button))
        form.addRow("Actual report", self._path_row(self.actual_edit, actual_button))
        form.addRow("Output folder", self.output_edit)
        form.addRow("Mode", self.mode_combo)
        form.addRow("Pixel threshold", self.threshold_spin)
        form.addRow("DPI", self.dpi_spin)
        form.addRow("Allowed diff %", self.allowed_spin)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.addLayout(form)
        left_layout.addWidget(self.only_diff_check)
        left_layout.addWidget(self.run_button)
        left_layout.addWidget(QLabel("Slides"))
        left_layout.addWidget(self.slide_list, 1)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.addWidget(self.preview, 4)
        right_layout.addWidget(QLabel("Details"))
        right_layout.addWidget(self.detail_box, 1)

        splitter = QSplitter()
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        self.setCentralWidget(splitter)

    def _path_row(self, edit: QLineEdit, button: QPushButton) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(edit, 1)
        layout.addWidget(button)
        return row

    def pick_file(self, edit: QLineEdit) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select report",
            "",
            "Reports (*.ppt *.pptx *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp);;All files (*.*)",
        )
        if path:
            edit.setText(path)

    def build_args(self) -> SimpleNamespace:
        return SimpleNamespace(
            expected=self.expected_edit.text().strip(),
            actual=self.actual_edit.text().strip(),
            output_dir=self.output_edit.text().strip() or engine.DEFAULT_OUTPUT_DIR,
            mode=self.mode_combo.currentText(),
            threshold=self.threshold_spin.value(),
            allowed_percent=float(self.allowed_spin.value()),
            highlight_color=(255, 0, 0),
            alpha=51,
            dpi=self.dpi_spin.value(),
        )

    def run_compare(self) -> None:
        args = self.build_args()
        if not args.expected or not args.actual:
            QMessageBox.warning(self, "Missing input", "Select expected and actual reports first.")
            return

        self.run_button.setEnabled(False)
        self.slide_list.clear()
        self.detail_box.clear()
        self.preview.label.setText("Running compare...")
        self.preview._pixmap = None
        self.statusBar().showMessage("Comparing reports. PowerPoint may take a moment...")

        self.worker_thread = QThread(self)
        self.worker = CompareWorker(args)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.compare_finished)
        self.worker.failed.connect(self.compare_failed)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def compare_finished(self, pixel_results, object_diffs, output_dir: str) -> None:
        self.pixel_results = list(pixel_results)
        self.object_diffs = list(object_diffs)
        self.output_dir = output_dir
        self.run_button.setEnabled(True)
        self.populate_slide_list()
        diff_pages = len([result for result in self.pixel_results if not result.passed])
        self.statusBar().showMessage(
            "Done. Different slides: {}. Object differences: {}. Output: {}".format(
                diff_pages,
                len(self.object_diffs),
                output_dir,
            )
        )

    def compare_failed(self, message: str) -> None:
        self.run_button.setEnabled(True)
        self.statusBar().showMessage("Compare failed")
        self.detail_box.setPlainText(message)
        QMessageBox.critical(self, "Compare failed", message.splitlines()[-1] if message else "Unknown error")

    def populate_slide_list(self) -> None:
        self.slide_list.clear()
        only_diff = self.only_diff_check.isChecked()
        for result in self.pixel_results:
            if only_diff and result.passed:
                continue
            item = QListWidgetItem(
                "Slide {:03d} | {} | {:.6f}% | {} regions".format(
                    result.page,
                    "PASS" if result.passed else "DIFF",
                    result.difference_percent,
                    len(result.regions),
                )
            )
            item.setData(Qt.UserRole, result)
            self.slide_list.addItem(item)

        if self.slide_list.count() > 0:
            self.slide_list.setCurrentRow(0)
        else:
            self.preview.label.setText("No different slides")
            self.preview._pixmap = None
            self.detail_box.setPlainText(self.object_diff_text())

    def slide_selected(self, current: Optional[QListWidgetItem], previous: Optional[QListWidgetItem]) -> None:
        if current is None:
            return
        result = current.data(Qt.UserRole)
        self.preview.set_image(result.output_overlay)
        details = [
            "Slide: {}".format(result.page),
            "Status: {}".format("PASS" if result.passed else "DIFFERENT"),
            "Different pixels: {} / {}".format(result.different_pixels, result.compared_pixels),
            "Difference percent: {:.6f}%".format(result.difference_percent),
            "Max channel delta: {}".format(result.max_channel_delta),
            "Bounding box: {}".format(result.bbox),
            "Highlight regions: {}".format(len(result.regions)),
            "Overlay: {}".format(result.output_overlay),
            "Mask: {}".format(result.output_mask),
            "",
            self.object_diff_text(result.page),
        ]
        self.detail_box.setPlainText("\n".join(details))

    def object_diff_text(self, slide: Optional[int] = None) -> str:
        diffs = self.object_diffs
        if slide is not None:
            diffs = [diff for diff in diffs if diff.slide == slide]
        if not diffs:
            return "Object differences: 0"

        lines = ["Object differences: {}".format(len(diffs))]
        for diff in diffs[:80]:
            lines.append(
                "Slide {} | Object {} | {} | expected={!r} | actual={!r}".format(
                    diff.slide,
                    diff.object_index,
                    diff.field,
                    diff.expected,
                    diff.actual,
                )
            )
        if len(diffs) > 80:
            lines.append("... more differences are available in object_summary.json")
        return "\n".join(lines)


def main() -> int:
    app = QApplication(sys.argv)
    window = ReportDiffWindow()
    window.show()
    return app.exec_()


if __name__ == "__main__":
    raise SystemExit(main())
