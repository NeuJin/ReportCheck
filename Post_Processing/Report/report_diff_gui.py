#!/usr/bin/env python3
"""GUI for PPT report visual comparison."""

from __future__ import annotations

import sys
import traceback
from types import SimpleNamespace
from typing import Optional

from PySide2.QtCore import QObject, QSize, Qt, QThread, Signal, QEvent
from PySide2.QtGui import QPixmap
from PySide2.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTextEdit,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

import pixel_report_diff as engine


HELP = {
    "mode": "both = visual pixel diff + PPTX object diff. Recommended.",
    "threshold": "Ignores tiny per-channel RGB differences. 3 recommended; 5-8 for render noise; 0 strict.",
    "dpi": "Slide screenshot resolution. 150 balanced; 200 better for small text/charts but slower.",
    "allowed": "Allowed different-pixel percent before a slide is DIFF. 0 strict; 0.01-0.05 tolerates tiny render noise.",
    "align": "Match slides by visual similarity before diff. Use when report page counts or order differ.",
    "matching": "Minimum similarity percentage for auto matching. 82% is practical; raise if wrong slides match, lower if related slides do not match.",
    "boxes": "Show or hide red highlight rectangles. Very large regions use lighter fill. Hover a visible rectangle to see its bbox/size.",
    "unmatched": "Hide pairs where one side has no matching slide, such as extra_actual or missing_actual.",
    "output": "Relative output folder inside this downloaded package only.",
}


class CompareWorker(QObject):
    finished = Signal(object, object, str)
    failed = Signal(str)
    progress = Signal(int, str)

    def __init__(self, args: SimpleNamespace):
        super().__init__()
        self.args = args

    def run(self) -> None:
        try:
            self.progress.emit(5, "Validating input files...")
            engine.validate_inputs(self.args)
            output_dir = engine.resolve_output_dir(self.args.output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)

            pixel_results = []
            object_diffs = []
            if self.args.mode in ("pixel", "both"):
                self.progress.emit(20, "Rendering slides and comparing pixels...")
                pixel_results = engine.compare_pixels(self.args, output_dir)
                self.progress.emit(70, "Pixel comparison complete.")
            if self.args.mode in ("object", "both"):
                self.progress.emit(75, "Comparing PPTX objects...")
                object_diffs = engine.compare_objects(self.args, output_dir)
                self.progress.emit(90, "Object comparison complete.")

            self.progress.emit(95, "Writing summary files...")
            engine.write_summary(output_dir, pixel_results, object_diffs)
            self.progress.emit(100, "Done.")
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
        self.label.setMouseTracking(True)
        self.label.installEventFilter(self)
        self.setMouseTracking(True)
        self.setWidget(self.label)
        self.setWidgetResizable(True)
        self._pixmap: Optional[QPixmap] = None
        self._scaled: Optional[QPixmap] = None
        self._result = None
        self._show_boxes = True

    def set_result(self, result, show_boxes: bool) -> None:
        self._result = result
        self._show_boxes = show_boxes

    def set_image(self, image_path: str) -> None:
        pixmap = QPixmap(image_path)
        if pixmap.isNull():
            self._pixmap = None
            self._scaled = None
            self.label.setText("Cannot load image")
            return
        self._pixmap = pixmap
        self._fit_pixmap()

    def resizeEvent(self, event):  # type: ignore[override]
        super().resizeEvent(event)
        self._fit_pixmap()

    def eventFilter(self, obj, event):  # type: ignore[override]
        if obj is self.label and event.type() == QEvent.MouseMove:
            self._show_region_tooltip(event)
        return super().eventFilter(obj, event)

    def _fit_pixmap(self) -> None:
        if self._pixmap is None:
            return
        size = self.viewport().size()
        self._scaled = self._pixmap.scaled(size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.label.setPixmap(self._scaled)

    def _show_region_tooltip(self, event) -> None:
        if self._pixmap is None or self._scaled is None or self._result is None:
            return
        if not self._show_boxes:
            return
        if not self._result.regions:
            return

        label_size = self.label.size()
        scaled_size = self._scaled.size()
        offset_x = max(0, (label_size.width() - scaled_size.width()) // 2)
        offset_y = max(0, (label_size.height() - scaled_size.height()) // 2)
        px = event.pos().x() - offset_x
        py = event.pos().y() - offset_y
        if px < 0 or py < 0 or px >= scaled_size.width() or py >= scaled_size.height():
            return

        original_x = int(px * self._pixmap.width() / max(1, scaled_size.width()))
        original_y = int(py * self._pixmap.height() / max(1, scaled_size.height()))
        indexed_regions = []
        for idx, region in enumerate(self._result.regions, start=1):
            x1, y1, x2, y2 = region
            area = max(1, (x2 - x1) * (y2 - y1))
            indexed_regions.append((area, idx, region))

        for _area, idx, region in sorted(indexed_regions):
            x1, y1, x2, y2 = region
            if x1 <= original_x <= x2 and y1 <= original_y <= y2:
                width = x2 - x1
                height = y2 - y1
                text = (
                    "Highlight region #{idx}\n"
                    "bbox: ({x1}, {y1}) - ({x2}, {y2})\n"
                    "size: {width} x {height}px\n"
                    "Expected page: {expected}\n"
                    "Actual page: {actual}\n"
                    "Diff: {diff:.6f}%"
                ).format(
                    idx=idx,
                    x1=x1,
                    y1=y1,
                    x2=x2,
                    y2=y2,
                    width=width,
                    height=height,
                    expected=self._result.expected_page if self._result.expected_page is not None else "-",
                    actual=self._result.actual_page if self._result.actual_page is not None else "-",
                    diff=self._result.difference_percent,
                )
                QToolTip.showText(event.globalPos(), text, self.label)
                return


class ReportDiffWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ReportCheck - PPT Slide Diff")
        self.resize(1240, 800)

        self.pixel_results = []
        self.object_diffs = []
        self.output_dir = ""
        self.worker_thread: Optional[QThread] = None

        self.expected_edit = QLineEdit()
        self.actual_edit = QLineEdit()
        self.output_edit = QLineEdit(engine.DEFAULT_OUTPUT_DIR)
        self.output_edit.setToolTip(HELP["output"])

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["both", "pixel", "object"])
        self.mode_combo.setToolTip(HELP["mode"])

        self.threshold_spin = QSpinBox()
        self.threshold_spin.setRange(0, 255)
        self.threshold_spin.setValue(3)
        self.threshold_spin.setToolTip(HELP["threshold"])

        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(50, 300)
        self.dpi_spin.setValue(150)
        self.dpi_spin.setToolTip(HELP["dpi"])

        self.allowed_spin = QDoubleSpinBox()
        self.allowed_spin.setRange(0.0, 100.0)
        self.allowed_spin.setDecimals(4)
        self.allowed_spin.setSingleStep(0.01)
        self.allowed_spin.setValue(0.0)
        self.allowed_spin.setSuffix(" %")
        self.allowed_spin.setToolTip(HELP["allowed"])

        self.align_check = QCheckBox("Auto align slides")
        self.align_check.setChecked(True)
        self.align_check.setToolTip(HELP["align"])

        self.matching_spin = QDoubleSpinBox()
        self.matching_spin.setRange(0.0, 100.0)
        self.matching_spin.setDecimals(1)
        self.matching_spin.setSingleStep(1.0)
        self.matching_spin.setValue(82.0)
        self.matching_spin.setSuffix(" %")
        self.matching_spin.setToolTip(HELP["matching"])

        self.show_boxes_check = QCheckBox("Show highlight boxes")
        self.show_boxes_check.setChecked(True)
        self.show_boxes_check.setToolTip(HELP["boxes"])
        self.show_boxes_check.stateChanged.connect(lambda _state: self.update_preview_status())
        self.show_boxes_check.stateChanged.connect(self.refresh_current_preview)

        self.view_overlay_button = QPushButton("Overlay")
        self.toggle_original_button = QPushButton("Toggle original")
        self.original_status_label = QLabel("View: Overlay")
        self.view_overlay_button.setToolTip("Show actual report with red highlight boxes")
        self.toggle_original_button.setToolTip("Switch between original expected/old and actual/new slide images")
        self.view_overlay_button.clicked.connect(lambda: self.set_preview_mode("overlay"))
        self.toggle_original_button.clicked.connect(self.toggle_original_preview)
        self.preview_mode = "overlay"
        self.original_mode = "expected"

        self.hide_unmatched_check = QCheckBox("Hide unmatched slides")
        self.hide_unmatched_check.setChecked(True)
        self.hide_unmatched_check.setToolTip(HELP["unmatched"])
        self.hide_unmatched_check.stateChanged.connect(self.populate_slide_list)

        self.only_diff_check = QCheckBox("Only different slides")
        self.only_diff_check.setChecked(True)
        self.only_diff_check.stateChanged.connect(self.populate_slide_list)

        self.run_button = QPushButton("Run compare")
        self.run_button.clicked.connect(self.run_compare)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Ready")

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
        form.addRow("Output folder", self._control_with_help(self.output_edit, HELP["output"]))
        form.addRow("Mode", self._control_with_help(self.mode_combo, HELP["mode"]))
        form.addRow("Pixel threshold", self._control_with_help(self.threshold_spin, HELP["threshold"]))
        form.addRow("DPI", self._control_with_help(self.dpi_spin, HELP["dpi"]))
        form.addRow("Allowed diff", self._control_with_help(self.allowed_spin, HELP["allowed"]))
        form.addRow("Auto align", self._control_with_help(self.align_check, HELP["align"]))
        form.addRow("Min Matching", self._control_with_help(self.matching_spin, HELP["matching"]))
        form.addRow("Highlight", self._control_with_help(self.show_boxes_check, HELP["boxes"]))
        form.addRow("Unmatched", self._control_with_help(self.hide_unmatched_check, HELP["unmatched"]))

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.addLayout(form)
        left_layout.addWidget(self.only_diff_check)
        left_layout.addWidget(self.run_button)
        left_layout.addWidget(self.progress_bar)
        left_layout.addWidget(QLabel("Slides"))
        left_layout.addWidget(self.slide_list, 1)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        preview_toolbar = QHBoxLayout()
        preview_toolbar.addWidget(QLabel("View"))
        preview_toolbar.addWidget(self.view_overlay_button)
        preview_toolbar.addWidget(self.toggle_original_button)
        preview_toolbar.addWidget(self.original_status_label)
        preview_toolbar.addStretch(1)
        right_layout.addLayout(preview_toolbar)
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

    def _control_with_help(self, control: QWidget, help_text: str) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(control, 1)
        help_label = QLabel("?")
        help_label.setAlignment(Qt.AlignCenter)
        help_label.setFixedSize(18, 18)
        help_label.setToolTip(help_text)
        help_label.setStyleSheet("border: 1px solid #888; border-radius: 9px; color: #333; background: #f2f2f2;")
        layout.addWidget(help_label)
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
            align_slides=self.align_check.isChecked(),
            min_match_score=float(self.matching_spin.value()) / 100.0,
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
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0% - Starting")
        self.statusBar().showMessage("Comparing reports. PowerPoint may take a moment...")

        self.worker_thread = QThread(self)
        self.worker = CompareWorker(args)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.compare_finished)
        self.worker.failed.connect(self.compare_failed)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def update_progress(self, percent: int, message: str) -> None:
        self.progress_bar.setValue(percent)
        self.progress_bar.setFormat("{}% - {}".format(percent, message))
        self.statusBar().showMessage(message)

    def compare_finished(self, pixel_results, object_diffs, output_dir: str) -> None:
        self.pixel_results = list(pixel_results)
        self.object_diffs = list(object_diffs)
        self.output_dir = output_dir
        self.run_button.setEnabled(True)
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("100% - Done")
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
        self.progress_bar.setFormat("Failed")
        self.statusBar().showMessage("Compare failed")
        self.detail_box.setPlainText(message)
        QMessageBox.critical(self, "Compare failed", message.splitlines()[-1] if message else "Unknown error")

    def populate_slide_list(self) -> None:
        self.slide_list.clear()
        only_diff = self.only_diff_check.isChecked()
        hide_unmatched = self.hide_unmatched_check.isChecked()
        for result in self.pixel_results:
            if hide_unmatched and result.match_status in ("extra_actual", "missing_actual"):
                continue
            if only_diff and result.passed:
                continue
            item = QListWidgetItem(
                "Pair {:03d} | E:{} -> A:{} | {} | diff {:.6f}% | {} regions".format(
                    result.page,
                    result.expected_page if result.expected_page is not None else "-",
                    result.actual_page if result.actual_page is not None else "-",
                    result.match_status,
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

    def set_preview_mode(self, mode: str) -> None:
        self.preview_mode = mode
        self.update_preview_status()
        self.refresh_current_preview()

    def toggle_original_preview(self) -> None:
        if self.preview_mode not in ("expected", "actual"):
            self.preview_mode = self.original_mode
        elif self.preview_mode == "expected":
            self.preview_mode = "actual"
            self.original_mode = "actual"
        else:
            self.preview_mode = "expected"
            self.original_mode = "expected"
        self.update_preview_status()
        self.refresh_current_preview()

    def update_preview_status(self) -> None:
        if self.preview_mode == "expected":
            self.original_status_label.setText("Original: Expected")
        elif self.preview_mode == "actual":
            self.original_status_label.setText("Original: Actual")
        elif self.show_boxes_check.isChecked():
            self.original_status_label.setText("View: Overlay")
        else:
            self.original_status_label.setText("View: Actual clean")

    def refresh_current_preview(self) -> None:
        current = self.slide_list.currentItem()
        if current is not None:
            self.slide_selected(current, None)

    def image_path_for_mode(self, result) -> str:
        if self.preview_mode == "expected":
            return result.output_expected
        if self.preview_mode == "actual":
            return result.output_actual
        if self.show_boxes_check.isChecked():
            return result.output_overlay
        return result.output_actual

    def slide_selected(self, current: Optional[QListWidgetItem], previous: Optional[QListWidgetItem]) -> None:
        if current is None:
            return
        result = current.data(Qt.UserRole)
        self.preview.set_result(result, self.show_boxes_check.isChecked() and self.preview_mode == "overlay")
        self.preview.set_image(self.image_path_for_mode(result))
        if result.match_status in ("extra_actual", "missing_actual"):
            details = [
                "Pair: {}".format(result.page),
                "Expected page: {}".format(result.expected_page if result.expected_page is not None else "-"),
                "Actual page: {}".format(result.actual_page if result.actual_page is not None else "-"),
                "Match status: {}".format(result.match_status),
                "Unmatched slide. Detailed diff is hidden; compare only matched slide pairs.",
            ]
            self.detail_box.setPlainText("\n".join(details))
            return
        details = [
            "Pair: {}".format(result.page),
            "Expected page: {}".format(result.expected_page if result.expected_page is not None else "-"),
            "Actual page: {}".format(result.actual_page if result.actual_page is not None else "-"),
            "Status: {}".format("PASS" if result.passed else "DIFFERENT"),
            "Different pixels: {} / {}".format(result.different_pixels, result.compared_pixels),
            "Difference percent: {:.6f}%".format(result.difference_percent),
            "Max channel delta: {}".format(result.max_channel_delta),
            "Bounding box: {}".format(result.bbox),
            "Highlight regions: {}".format(len(result.regions)),
            "Expected original: {}".format(result.output_expected),
            "Actual original: {}".format(result.output_actual),
            "Overlay: {}".format(result.output_overlay),
            "Mask: {}".format(result.output_mask),
            "",
            self.object_diff_text(result.expected_page),
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
