#!/usr/bin/env python3
"""GUI for PPT report visual comparison."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import traceback
from collections import OrderedDict
from dataclasses import asdict, fields
from pathlib import Path
from types import SimpleNamespace
from typing import Dict, Optional

from PySide2.QtCore import QObject, QSize, Qt, QThread, Signal, QEvent
from PySide2.QtGui import QBrush, QColor, QIcon, QPixmap
from PySide2.QtWidgets import (
    QAction,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTabWidget,
    QTextEdit,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

import pixel_report_diff as engine


# Preset profiles — one-click matching strictness for common scenarios.
# (label, min_match_percent, gap_penalty, tooltip)
MATCHING_PROFILES = [
    ("Strict 95%",    95.0, -0.05,
     "Near-identical reports only. Use when comparing minor edits to the same deck."),
    ("Balanced 82%",  82.0, -0.12,
     "Default — accepts small layout/colour changes; rejects unrelated slides."),
    ("Cross-version 65%", 65.0, -0.18,
     "Different test runs of the same template (different data, same boilerplate)."),
    ("Permissive 50%", 50.0, -0.25,
     "Cross-project matching. Many low-confidence matches; review carefully."),
]

# Status icons make the slide list scannable at a glance.
MATCH_STATUS_ICON = {
    "matched": "✓",
    "low_confidence_match": "≈",
    "missing_actual": "⊘",
    "extra_actual": "+",
    "same_index": "·",
    "manual_match": "M",
}


HELP = {
    "mode": "both = visual pixel diff + PPTX object diff. Recommended.",
    "threshold": "Ignores tiny per-channel RGB differences. 3 recommended; 5-8 for render noise; 0 strict.",
    "dpi": "Slide screenshot resolution. 150 balanced; 200 better for small text/charts but slower.",
    "allowed": "Allowed different-pixel percent before a slide is DIFF. 0 strict; 0.01-0.05 tolerates tiny render noise.",
    "align": "Match slides by visual similarity before diff. Use when report page counts or order differ.",
    "matching": "Minimum similarity percentage for auto matching. 82% is practical; raise if wrong slides match, lower if related slides do not match.",
    "gap": "Penalty for leaving a slide unmatched during auto align. More negative = matcher pairs slides only when clearly similar; less negative = more eager to pair. -0.12 is the default.",
    "overrides": ('JSON file forcing specific slide pairings, e.g. {"1": 2, "3": null} '
                  'pairs expected slide 1 with actual slide 2 and forces expected slide 3 '
                  'to "missing". Use "Export current as template" after a run to get a '
                  'starting file that you can edit by hand.'),
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
            matches = None
            if self.args.mode in ("pixel", "both"):
                self.progress.emit(20, "Rendering slides and comparing pixels...")
                pixel_results, matches = engine.compare_pixels(self.args, output_dir)
                self.progress.emit(70, "Pixel comparison complete.")
            if self.args.mode in ("object", "both"):
                self.progress.emit(75, "Comparing PPTX objects...")
                object_diffs = engine.compare_objects(self.args, output_dir, matches=matches)
                self.progress.emit(90, "Object comparison complete.")

            self.progress.emit(95, "Writing summary files...")
            engine.write_summary(output_dir, pixel_results, object_diffs)
            self.progress.emit(100, "Done.")
            self.finished.emit(pixel_results, object_diffs, str(output_dir))
        except Exception:
            self.failed.emit(traceback.format_exc())


_PIXMAP_CACHE: "OrderedDict[str, QPixmap]" = OrderedDict()
_PIXMAP_CACHE_LIMIT = 60

_THUMB_CACHE: "OrderedDict[tuple, QPixmap]" = OrderedDict()
_THUMB_CACHE_LIMIT = 200


def _load_pixmap_cached(image_path: str) -> Optional[QPixmap]:
    """LRU-cached QPixmap.

    Slide navigation re-opens the same PNGs repeatedly; loading a 1-2 MB image
    from disk + decoding takes long enough to feel sluggish. Caching the
    decoded QPixmap makes every re-visit instant.
    """
    cached = _PIXMAP_CACHE.get(image_path)
    if cached is not None:
        _PIXMAP_CACHE.move_to_end(image_path)
        return cached
    pixmap = QPixmap(image_path)
    if pixmap.isNull():
        return None
    _PIXMAP_CACHE[image_path] = pixmap
    if len(_PIXMAP_CACHE) > _PIXMAP_CACHE_LIMIT:
        _PIXMAP_CACHE.popitem(last=False)
    return pixmap


def _load_thumbnail_cached(image_path: str, size: QSize) -> Optional[QPixmap]:
    """LRU-cached scaled thumbnail keyed by path + target size."""
    key = (image_path, size.width(), size.height())
    cached = _THUMB_CACHE.get(key)
    if cached is not None:
        _THUMB_CACHE.move_to_end(key)
        return cached
    full = _load_pixmap_cached(image_path)
    if full is None:
        return None
    thumb = full.scaled(size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    _THUMB_CACHE[key] = thumb
    if len(_THUMB_CACHE) > _THUMB_CACHE_LIMIT:
        _THUMB_CACHE.popitem(last=False)
    return thumb


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
        self._last_image_path: Optional[str] = None

    def set_result(self, result, show_boxes: bool) -> None:
        self._result = result
        self._show_boxes = show_boxes

    def set_image(self, image_path: str) -> None:
        # Skip work if we're already showing this exact image.
        if image_path == self._last_image_path and self._pixmap is not None:
            return
        pixmap = _load_pixmap_cached(image_path)
        if pixmap is None:
            self._pixmap = None
            self._scaled = None
            self._last_image_path = None
            self.label.setText("Cannot load image")
            return
        self._pixmap = pixmap
        self._last_image_path = image_path
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
        self.resize(1480, 860)

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

        self.gap_penalty_spin = QDoubleSpinBox()
        self.gap_penalty_spin.setRange(-1.0, 0.0)
        self.gap_penalty_spin.setDecimals(2)
        self.gap_penalty_spin.setSingleStep(0.01)
        self.gap_penalty_spin.setValue(engine.DEFAULT_GAP_PENALTY)
        self.gap_penalty_spin.setToolTip(HELP["gap"])

        self.overrides_edit = QLineEdit()
        self.overrides_edit.setPlaceholderText("optional — JSON file with manual slide pairings")
        self.overrides_edit.setToolTip(HELP["overrides"])
        self.overrides_browse_button = QPushButton("Browse")
        self.overrides_browse_button.clicked.connect(self.pick_overrides_file)
        self.export_overrides_button = QPushButton("Export current as template")
        self.export_overrides_button.setEnabled(False)
        self.export_overrides_button.setToolTip(
            "Save the alignment from the latest run as a JSON template you can "
            "edit by hand to override specific pairings."
        )
        self.export_overrides_button.clicked.connect(self.export_current_overrides)

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

        self.open_output_button = QPushButton("Open output folder")
        self.open_output_button.setEnabled(False)
        self.open_output_button.setToolTip(
            "Reveal the folder containing overlays, masks, and similarity_matrix.json"
        )
        self.open_output_button.clicked.connect(self.open_output_folder)

        self.profile_buttons = []
        for label, min_pct, gap, tip in MATCHING_PROFILES:
            button = QPushButton(label)
            button.setToolTip(tip)
            button.clicked.connect(
                lambda _checked=False, m=min_pct, g=gap: self._apply_profile(m, g)
            )
            self.profile_buttons.append(button)

        self.summary_label = QLabel("Run a comparison to see results.")
        self.summary_label.setWordWrap(True)
        self.summary_label.setStyleSheet(
            "QLabel { padding: 6px; background: #f5f5f5; border: 1px solid #ddd; border-radius: 3px; }"
        )

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Ready")

        # Pairs list (centre-left): comparison pairs with status icons + thumbnails.
        # Expected / Actual lists (right panel): full per-file slide strips, shown
        # side-by-side so the user can pick one from each and pair them visually.
        thumb_icon_size = QSize(120, 68)

        self.slide_list = QListWidget()
        self.slide_list.setIconSize(thumb_icon_size)
        self.slide_list.setSpacing(2)
        self.slide_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.slide_list.customContextMenuRequested.connect(self._show_pair_context_menu)
        self.slide_list.currentItemChanged.connect(self.slide_selected)

        self.expected_list = QListWidget()
        self.expected_list.setIconSize(thumb_icon_size)
        self.expected_list.setSpacing(2)
        self.expected_list.currentItemChanged.connect(self._expected_selected)

        self.actual_list = QListWidget()
        self.actual_list.setIconSize(thumb_icon_size)
        self.actual_list.setSpacing(2)
        self.actual_list.currentItemChanged.connect(self._actual_selected)

        # "Pair selected" button used between the two side-by-side lists.
        self.pair_selected_button = QPushButton("🔗  Pair selected")
        self.pair_selected_button.setToolTip(
            "Select one slide on the Expected side and one on the Actual side, "
            "then click this to create a manual pair. The new pair appears in "
            "the Pairs list immediately — no re-Run needed."
        )
        self.pair_selected_button.clicked.connect(self._pair_selected_slides)

        # Session save/load — lets the user resume a comparison later without
        # re-running anything.
        self.save_session_button = QPushButton("Save session…")
        self.save_session_button.setToolTip(
            "Save settings + manual overrides + current results to a JSON file "
            "you can open later without re-running compare."
        )
        self.save_session_button.clicked.connect(self.save_session)
        self.load_session_button = QPushButton("Load session…")
        self.load_session_button.setToolTip(
            "Restore a saved session (settings, overrides, and last results)."
        )
        self.load_session_button.clicked.connect(self.load_session)

        # In-memory manual overrides: 0-based expected idx → 0-based actual idx
        # (or None for forced-missing). Persisted across runs within the session.
        self.manual_overrides: Dict[int, Optional[int]] = {}

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

        # ── Files ────────────────────────────────────────────────────────────
        files_group = QGroupBox("Files")
        files_form = QFormLayout(files_group)
        files_form.addRow("Expected report", self._path_row(self.expected_edit, expected_button))
        files_form.addRow("Actual report", self._path_row(self.actual_edit, actual_button))
        files_form.addRow("Output folder", self._control_with_help(self.output_edit, HELP["output"]))

        # ── Comparison knobs ─────────────────────────────────────────────────
        comparison_group = QGroupBox("Comparison")
        comp_form = QFormLayout(comparison_group)
        comp_form.addRow("Mode", self._control_with_help(self.mode_combo, HELP["mode"]))
        comp_form.addRow("Pixel threshold", self._control_with_help(self.threshold_spin, HELP["threshold"]))
        comp_form.addRow("DPI", self._control_with_help(self.dpi_spin, HELP["dpi"]))
        comp_form.addRow("Allowed diff", self._control_with_help(self.allowed_spin, HELP["allowed"]))

        # ── Slide matching (with one-click presets) ──────────────────────────
        matching_group = QGroupBox("Slide matching")
        matching_layout = QVBoxLayout(matching_group)
        matching_layout.addWidget(self.align_check)

        profile_label = QLabel("Preset:")
        profile_label.setStyleSheet("QLabel { color: #555; font-size: 11px; }")
        matching_layout.addWidget(profile_label)
        profile_row = QHBoxLayout()
        profile_row.setSpacing(4)
        for button in self.profile_buttons:
            profile_row.addWidget(button)
        matching_layout.addLayout(profile_row)

        matching_form = QFormLayout()
        matching_form.addRow("Min matching", self._control_with_help(self.matching_spin, HELP["matching"]))
        matching_form.addRow("Gap penalty", self._control_with_help(self.gap_penalty_spin, HELP["gap"]))
        matching_form.addRow(
            "Manual overrides",
            self._path_row(self.overrides_edit, self.overrides_browse_button),
        )
        matching_layout.addLayout(matching_form)
        matching_layout.addWidget(self.export_overrides_button)

        # ── Display filters ──────────────────────────────────────────────────
        display_group = QGroupBox("Display filters")
        display_layout = QVBoxLayout(display_group)
        display_layout.addWidget(self.show_boxes_check)
        display_layout.addWidget(self.hide_unmatched_check)
        display_layout.addWidget(self.only_diff_check)

        # ── Action row + progress ────────────────────────────────────────────
        action_row = QHBoxLayout()
        action_row.addWidget(self.run_button)
        action_row.addWidget(self.open_output_button)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setSpacing(8)
        left_layout.addWidget(files_group)
        left_layout.addWidget(comparison_group)
        left_layout.addWidget(matching_group)
        left_layout.addWidget(display_group)
        left_layout.addLayout(action_row)
        left_layout.addWidget(self.progress_bar)
        left_layout.addWidget(self.summary_label)
        left_layout.addWidget(QLabel("Comparison pairs (right-click for manual override)"))
        left_layout.addWidget(self.slide_list, 1)

        # Session toolbar — pinned at the very top of the left panel.
        session_row = QHBoxLayout()
        session_row.addWidget(self.save_session_button)
        session_row.addWidget(self.load_session_button)
        session_row.addStretch(1)
        left_layout.insertLayout(0, session_row)

        # ── Right panel ──────────────────────────────────────────────────────
        # Top: Expected | [Pair] | Actual lists side by side.
        # Below: preview toolbar, preview, details.

        # ── Far-right strip: two narrow slide-list columns + Pair button below ──
        # Mimics PowerPoint's slide navigator. Kept narrow so the centre
        # comparison area takes the bulk of the window.
        expected_box = QWidget()
        ev = QVBoxLayout(expected_box)
        ev.setContentsMargins(0, 0, 0, 0)
        ev.setSpacing(2)
        ev.addWidget(QLabel("Expected"))
        ev.addWidget(self.expected_list, 1)

        actual_box = QWidget()
        av = QVBoxLayout(actual_box)
        av.setContentsMargins(0, 0, 0, 0)
        av.setSpacing(2)
        av.addWidget(QLabel("Actual"))
        av.addWidget(self.actual_list, 1)

        lists_splitter = QSplitter(Qt.Horizontal)
        lists_splitter.addWidget(expected_box)
        lists_splitter.addWidget(actual_box)
        lists_splitter.setStretchFactor(0, 1)
        lists_splitter.setStretchFactor(1, 1)

        far_right = QWidget()
        fr_layout = QVBoxLayout(far_right)
        fr_layout.setContentsMargins(4, 4, 4, 4)
        fr_layout.setSpacing(4)
        fr_layout.addWidget(lists_splitter, 1)
        fr_layout.addWidget(self.pair_selected_button)

        # ── Centre: main comparison window (preview + details) ──
        centre = QWidget()
        centre_layout = QVBoxLayout(centre)
        preview_toolbar = QHBoxLayout()
        preview_toolbar.addWidget(QLabel("View"))
        preview_toolbar.addWidget(self.view_overlay_button)
        preview_toolbar.addWidget(self.toggle_original_button)
        preview_toolbar.addWidget(self.original_status_label)
        preview_toolbar.addStretch(1)
        centre_layout.addLayout(preview_toolbar)
        centre_layout.addWidget(self.preview, 4)
        centre_layout.addWidget(QLabel("Details"))
        centre_layout.addWidget(self.detail_box, 1)

        # Three-column main splitter: settings/pairs | comparison | slide strip
        splitter = QSplitter()
        splitter.addWidget(left)
        splitter.addWidget(centre)
        splitter.addWidget(far_right)
        splitter.setStretchFactor(0, 0)  # settings/pairs panel stays its natural width
        splitter.setStretchFactor(1, 1)  # comparison area absorbs extra space
        splitter.setStretchFactor(2, 0)  # slide-list strip stays narrow
        splitter.setSizes([340, 820, 240])
        self.setCentralWidget(splitter)

    def _path_row(self, edit: QLineEdit, button: QPushButton) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(edit, 1)
        layout.addWidget(button)
        return row

    def _apply_profile(self, min_match_pct: float, gap_penalty: float) -> None:
        """One-click profile: snap the matching knobs to a known-good combination."""
        self.matching_spin.setValue(min_match_pct)
        self.gap_penalty_spin.setValue(gap_penalty)
        self.statusBar().showMessage(
            "Matching profile: min={:.0f}%, gap={:.2f}".format(min_match_pct, gap_penalty),
            4000,
        )

    def _populate_original_file_lists(self) -> None:
        """Fill the Expected / Actual tabs with one row per slide from each file.

        Lets the user browse a report end-to-end (PowerPoint-strip style),
        independent of which slides happened to pair up.
        """
        self.expected_list.blockSignals(True)
        self.actual_list.blockSignals(True)
        self.expected_list.clear()
        self.actual_list.clear()

        # Pull unique (slide_number → image_path) pairs from the per-pair results.
        expected_slides: Dict[int, str] = {}
        actual_slides: Dict[int, str] = {}
        for result in self.pixel_results:
            if result.expected_page is not None and result.output_expected:
                expected_slides.setdefault(result.expected_page, result.output_expected)
            if result.actual_page is not None and result.output_actual:
                actual_slides.setdefault(result.actual_page, result.output_actual)

        icon_size = self.expected_list.iconSize()
        for slide_no, image_path in sorted(expected_slides.items()):
            item = QListWidgetItem("Slide {:03d}".format(slide_no))
            thumb = _load_thumbnail_cached(image_path, icon_size)
            if thumb is not None:
                item.setIcon(QIcon(thumb))
            item.setData(Qt.UserRole, (slide_no, image_path))
            self.expected_list.addItem(item)

        for slide_no, image_path in sorted(actual_slides.items()):
            item = QListWidgetItem("Slide {:03d}".format(slide_no))
            thumb = _load_thumbnail_cached(image_path, icon_size)
            if thumb is not None:
                item.setIcon(QIcon(thumb))
            item.setData(Qt.UserRole, (slide_no, image_path))
            self.actual_list.addItem(item)

        self.expected_list.blockSignals(False)
        self.actual_list.blockSignals(False)

    def _expected_selected(self, current, _previous) -> None:
        self._show_standalone_slide(current, side="Expected")

    def _actual_selected(self, current, _previous) -> None:
        self._show_standalone_slide(current, side="Actual")

    def _show_standalone_slide(self, item: Optional[QListWidgetItem], side: str) -> None:
        if item is None:
            return
        data = item.data(Qt.UserRole)
        if data is None:
            return
        slide_no, image_path = data
        # Drop overlay-region metadata so tooltip code stays quiet on plain views.
        self.preview.set_result(None, False)
        self.preview.set_image(image_path)
        self.detail_box.setPlainText(
            "{} report — slide {}\n\nFile: {}".format(side, slide_no, image_path)
        )

    # ── Manual overrides via right-click ──────────────────────────────────

    def _show_pair_context_menu(self, position) -> None:
        item = self.slide_list.itemAt(position)
        if item is None:
            return
        result = item.data(Qt.UserRole)
        if result is None:
            return

        menu = QMenu(self)
        if result.expected_page is not None:
            menu.addAction(
                "Force pair with actual slide…",
                lambda: self._force_pair_with_actual(result),
            )
            menu.addAction(
                "Force missing (no actual pair)",
                lambda: self._force_missing(result),
            )
            menu.addAction(
                "Reset this pair to auto match",
                lambda: self._reset_override(result),
            )
        else:
            menu.addAction(
                "Pair this actual slide with expected slide…",
                lambda: self._force_pair_with_expected(result),
            )
        menu.addSeparator()
        menu.addAction("Delete this pair", lambda: self._delete_current_pair(result))
        menu.addSeparator()
        menu.addAction("Clear all manual overrides", self._clear_all_overrides)
        menu.exec_(self.slide_list.viewport().mapToGlobal(position))

    def _available_actual_slides(self) -> "list[int]":
        seen = set()
        for result in self.pixel_results:
            if result.actual_page is not None:
                seen.add(result.actual_page)
        return sorted(seen)

    def _available_expected_slides(self) -> "list[int]":
        seen = set()
        for result in self.pixel_results:
            if result.expected_page is not None:
                seen.add(result.expected_page)
        return sorted(seen)

    def _force_pair_with_actual(self, result) -> None:
        actuals = self._available_actual_slides()
        if not actuals:
            return
        current = result.actual_page if result.actual_page is not None else actuals[0]
        choice, ok = QInputDialog.getInt(
            self,
            "Force pair",
            "Pair expected slide {} with actual slide:".format(result.expected_page),
            current,
            min(actuals),
            max(actuals),
            1,
        )
        if not ok:
            return
        self.manual_overrides[result.expected_page - 1] = choice - 1
        self._mark_overrides_pending()

    def _force_pair_with_expected(self, result) -> None:
        expecteds = self._available_expected_slides()
        if not expecteds or result.actual_page is None:
            return
        choice, ok = QInputDialog.getInt(
            self,
            "Pair with expected",
            "Pair expected slide ? with actual slide {}:".format(result.actual_page),
            expecteds[0],
            min(expecteds),
            max(expecteds),
            1,
        )
        if not ok:
            return
        self.manual_overrides[choice - 1] = result.actual_page - 1
        self._mark_overrides_pending()

    def _force_missing(self, result) -> None:
        self.manual_overrides[result.expected_page - 1] = None
        self._mark_overrides_pending()

    def _reset_override(self, result) -> None:
        key = result.expected_page - 1
        if key in self.manual_overrides:
            del self.manual_overrides[key]
            self._mark_overrides_pending()

    def _clear_all_overrides(self) -> None:
        if not self.manual_overrides:
            return
        self.manual_overrides.clear()
        self.overrides_edit.clear()
        self.statusBar().showMessage("All manual overrides cleared.", 4000)

    def _delete_current_pair(self, result) -> None:
        """Drop a pair from the list. If it was matched, the deletion is recorded
        as a 'force missing' override so the next Run preserves the user's
        choice. The freed slides remain visible in the Expected/Actual columns
        on the right so they can be re-paired with one click."""
        try:
            self.pixel_results.remove(result)
        except ValueError:
            return

        if result.expected_page is not None and result.actual_page is not None:
            # Was an actual pairing — sticky-delete it so a re-Run doesn't
            # auto-recreate it from the same scores.
            self.manual_overrides[result.expected_page - 1] = None
            self._mark_overrides_pending()

        self._renumber_pairs()
        self.populate_slide_list()
        self._update_summary()
        self.statusBar().showMessage(
            "Pair deleted. Both slides remain in the side columns for re-pairing.",
            5000,
        )

    # ── Visual pair-by-click between Expected / Actual lists ──────────────

    def _pair_selected_slides(self) -> None:
        """Create a manual pair from the currently-selected Expected + Actual slides.

        The new pair is materialised *immediately* via engine.compare_page using
        the per-slide PNGs already in the output folder — so the user sees the
        diff right away without a full re-Run. The pairing is also recorded in
        self.manual_overrides so the next Run keeps it.
        """
        e_item = self.expected_list.currentItem()
        a_item = self.actual_list.currentItem()
        if e_item is None or a_item is None:
            QMessageBox.information(
                self,
                "Select slides",
                "Highlight one slide on the Expected side and one on the Actual "
                "side, then click Pair.",
            )
            return
        if not self.output_dir:
            QMessageBox.warning(
                self,
                "No output folder",
                "Run a comparison once first so the tool has somewhere to write the diff images.",
            )
            return

        e_slide, e_path = e_item.data(Qt.UserRole)
        a_slide, a_path = a_item.data(Qt.UserRole)

        for result in self.pixel_results:
            if result.expected_page == e_slide and result.actual_page == a_slide:
                self.statusBar().showMessage(
                    "Expected slide {} is already paired with actual slide {}.".format(
                        e_slide, a_slide
                    ),
                    4000,
                )
                return

        # Drop any conflicting pair, remembering the slides that lost their partner
        # so we can re-add them as missing/extra entries.
        orphan_actual: "list[tuple]" = []
        orphan_expected: "list[tuple]" = []
        kept = []
        for result in self.pixel_results:
            if result.expected_page == e_slide:
                if result.actual_page is not None and result.actual_page != a_slide:
                    orphan_actual.append((result.actual_page, result.output_actual))
                continue
            if result.actual_page == a_slide:
                if result.expected_page is not None and result.expected_page != e_slide:
                    orphan_expected.append((result.expected_page, result.output_expected))
                continue
            kept.append(result)

        try:
            new_pair = self._build_manual_pair(e_slide, e_path, a_slide, a_path,
                                                len(kept) + 1)
            new_results = kept + [new_pair]
            next_no = len(new_results) + 1
            for slide_no, image_path in orphan_expected:
                new_results.append(self._build_unmatched_pair(
                    slide_no, image_path, side="expected", pair_number=next_no))
                next_no += 1
            for slide_no, image_path in orphan_actual:
                new_results.append(self._build_unmatched_pair(
                    slide_no, image_path, side="actual", pair_number=next_no))
                next_no += 1
        except Exception as exc:
            QMessageBox.critical(self, "Pair generation failed", str(exc))
            return

        self.pixel_results = new_results
        self._renumber_pairs()
        self.manual_overrides[e_slide - 1] = a_slide - 1

        self.populate_slide_list()
        self._update_summary()
        self._mark_overrides_pending()
        self.statusBar().showMessage(
            "Paired expected slide {} ↔ actual slide {}.".format(e_slide, a_slide),
            5000,
        )

    def _build_manual_pair(
        self,
        expected_slide: int,
        expected_path: str,
        actual_slide: int,
        actual_path: str,
        pair_number: int,
    ):
        from PIL import Image
        expected_img = Image.open(expected_path).convert("RGB")
        actual_img = Image.open(actual_path).convert("RGB")
        return engine.compare_page(
            expected=expected_img,
            actual=actual_img,
            page_number=pair_number,
            expected_page=expected_slide,
            actual_page=actual_slide,
            match_score=1.0,
            match_status="manual_match",
            output_dir=Path(self.output_dir),
            threshold=self.threshold_spin.value(),
            allowed_percent=float(self.allowed_spin.value()),
            highlight_color=(255, 0, 0),
            alpha=51,
        )

    def _build_unmatched_pair(
        self,
        slide_number: int,
        image_path: str,
        side: str,  # 'expected' or 'actual'
        pair_number: int,
    ):
        from PIL import Image
        blank = Image.new("RGB", (1, 1), (255, 255, 255))
        try:
            slide_img = Image.open(image_path).convert("RGB")
        except Exception:
            slide_img = blank
        if side == "expected":
            expected_img, actual_img = slide_img, blank
            e_page, a_page, status = slide_number, None, "missing_actual"
        else:
            expected_img, actual_img = blank, slide_img
            e_page, a_page, status = None, slide_number, "extra_actual"
        return engine.compare_page(
            expected=expected_img,
            actual=actual_img,
            page_number=pair_number,
            expected_page=e_page,
            actual_page=a_page,
            match_score=None,
            match_status=status,
            output_dir=Path(self.output_dir),
            threshold=self.threshold_spin.value(),
            allowed_percent=float(self.allowed_spin.value()),
            highlight_color=(255, 0, 0),
            alpha=51,
        )

    def _renumber_pairs(self) -> None:
        """Re-sort + re-number pairs so the list reads naturally after edits."""
        def sort_key(result):
            e = result.expected_page if result.expected_page is not None else 10 ** 9
            a = result.actual_page if result.actual_page is not None else 10 ** 9
            return (e, a)
        self.pixel_results.sort(key=sort_key)
        for index, result in enumerate(self.pixel_results, start=1):
            result.page = index

    # ── Session save / load ───────────────────────────────────────────────

    def save_session(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "Save session", "session.json", "JSON files (*.json);;All files (*.*)"
        )
        if not path:
            return
        try:
            session = {
                "version": 1,
                "expected_path": self.expected_edit.text(),
                "actual_path": self.actual_edit.text(),
                "output_dir": self.output_dir,
                "settings": {
                    "output_dir_field": self.output_edit.text(),
                    "mode": self.mode_combo.currentText(),
                    "threshold": self.threshold_spin.value(),
                    "dpi": self.dpi_spin.value(),
                    "allowed_percent": self.allowed_spin.value(),
                    "align_slides": self.align_check.isChecked(),
                    "min_match_score": self.matching_spin.value(),
                    "gap_penalty": self.gap_penalty_spin.value(),
                    "overrides_path": self.overrides_edit.text(),
                    "show_boxes": self.show_boxes_check.isChecked(),
                    "hide_unmatched": self.hide_unmatched_check.isChecked(),
                    "only_diff": self.only_diff_check.isChecked(),
                },
                "manual_overrides": {
                    str(key + 1): (None if value is None else value + 1)
                    for key, value in self.manual_overrides.items()
                },
                "results": {
                    "pixel_results": [asdict(result) for result in self.pixel_results],
                    "object_diffs": [asdict(diff) for diff in self.object_diffs],
                },
            }
            Path(path).write_text(
                json.dumps(session, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            self.statusBar().showMessage("Session saved: {}".format(path), 6000)
        except Exception as exc:
            QMessageBox.critical(self, "Save session failed", str(exc))

    def load_session(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Load session", "", "JSON files (*.json);;All files (*.*)"
        )
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            self.expected_edit.setText(data.get("expected_path", ""))
            self.actual_edit.setText(data.get("actual_path", ""))
            self.output_dir = data.get("output_dir", "")

            settings_data = data.get("settings", {})
            if "output_dir_field" in settings_data:
                self.output_edit.setText(settings_data["output_dir_field"])
            if "mode" in settings_data:
                idx = self.mode_combo.findText(settings_data["mode"])
                if idx >= 0:
                    self.mode_combo.setCurrentIndex(idx)
            for key, widget in (
                ("threshold", self.threshold_spin),
                ("dpi", self.dpi_spin),
            ):
                if key in settings_data:
                    widget.setValue(int(settings_data[key]))
            for key, widget in (
                ("allowed_percent", self.allowed_spin),
                ("min_match_score", self.matching_spin),
                ("gap_penalty", self.gap_penalty_spin),
            ):
                if key in settings_data:
                    widget.setValue(float(settings_data[key]))
            for key, widget in (
                ("align_slides", self.align_check),
                ("show_boxes", self.show_boxes_check),
                ("hide_unmatched", self.hide_unmatched_check),
                ("only_diff", self.only_diff_check),
            ):
                if key in settings_data:
                    widget.setChecked(bool(settings_data[key]))
            if "overrides_path" in settings_data:
                self.overrides_edit.setText(settings_data["overrides_path"])

            self.manual_overrides = {
                int(key) - 1: (None if value is None else int(value) - 1)
                for key, value in data.get("manual_overrides", {}).items()
            }

            results_data = data.get("results", {})
            page_diff_fields = {f.name for f in fields(engine.PageDiff)}
            object_diff_fields = {f.name for f in fields(engine.ObjectDiff)}

            self.pixel_results = []
            for record in results_data.get("pixel_results", []):
                kwargs = {k: v for k, v in record.items() if k in page_diff_fields}
                if kwargs.get("bbox") is not None:
                    kwargs["bbox"] = tuple(kwargs["bbox"])
                if "regions" in kwargs:
                    kwargs["regions"] = [tuple(reg) for reg in kwargs["regions"]]
                self.pixel_results.append(engine.PageDiff(**kwargs))

            self.object_diffs = []
            for record in results_data.get("object_diffs", []):
                kwargs = {k: v for k, v in record.items() if k in object_diff_fields}
                self.object_diffs.append(engine.ObjectDiff(**kwargs))

            self.populate_slide_list()
            self._populate_original_file_lists()
            self._update_summary()
            self.open_output_button.setEnabled(bool(self.output_dir))
            self.export_overrides_button.setEnabled(bool(self.pixel_results))
            self.statusBar().showMessage("Session loaded: {}".format(path), 6000)
        except Exception as exc:
            QMessageBox.critical(self, "Load session failed", str(exc))

    def _mark_overrides_pending(self) -> None:
        """Persist current in-memory overrides to a temp JSON and surface a hint.

        The next Run uses this file via the existing --match-overrides plumbing,
        so GUI and CLI take exactly the same path through the engine.
        """
        if not self.manual_overrides:
            self.overrides_edit.clear()
            return
        overrides_json = {
            str(key + 1): (None if value is None else value + 1)
            for key, value in self.manual_overrides.items()
        }
        temp_path = Path(tempfile.gettempdir()) / "report_diff_overrides_pending.json"
        try:
            temp_path.write_text(
                json.dumps(overrides_json, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception as exc:
            QMessageBox.warning(self, "Cannot stage overrides", str(exc))
            return
        self.overrides_edit.setText(str(temp_path))
        self.statusBar().showMessage(
            "{} manual override(s) staged — click Run compare to apply.".format(
                len(self.manual_overrides)
            ),
            6000,
        )

    def pick_overrides_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select match-overrides JSON",
            self.overrides_edit.text().strip() or "",
            "JSON files (*.json);;All files (*.*)",
        )
        if path:
            self.overrides_edit.setText(path)

    def export_current_overrides(self) -> None:
        if not self.pixel_results:
            QMessageBox.information(
                self,
                "Nothing to export",
                "Run a comparison first — the current alignment is exported as a "
                "starting template.",
            )
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save overrides template",
            "match_overrides.json",
            "JSON files (*.json);;All files (*.*)",
        )
        if not path:
            return
        try:
            # Reconstruct SlideMatch list from PageDiff results so the engine can serialize it.
            matches = [
                engine.SlideMatch(
                    expected_index=(result.expected_page - 1) if result.expected_page is not None else None,
                    actual_index=(result.actual_page - 1) if result.actual_page is not None else None,
                    score=result.match_score,
                    status=result.match_status,
                )
                for result in self.pixel_results
            ]
            engine.export_overrides_template(matches, __import__("pathlib").Path(path))
            self.overrides_edit.setText(path)
            self.statusBar().showMessage("Overrides template saved: {}".format(path), 6000)
        except Exception as exc:
            QMessageBox.critical(self, "Export failed", str(exc))

    def open_output_folder(self) -> None:
        """Reveal the comparison's output folder in the OS file manager."""
        if not self.output_dir:
            return
        path = self.output_dir
        try:
            if sys.platform == "win32":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception as exc:
            QMessageBox.warning(self, "Cannot open folder", str(exc))

    def _update_summary(self) -> None:
        """Refresh the at-a-glance counters above the slide list."""
        if not self.pixel_results and not self.object_diffs:
            self.summary_label.setText("Run a comparison to see results.")
            return

        counts = {"matched": 0, "low_confidence_match": 0, "missing_actual": 0,
                  "extra_actual": 0, "same_index": 0, "manual_match": 0}
        diff_pairs = 0
        for result in self.pixel_results:
            counts[result.match_status] = counts.get(result.match_status, 0) + 1
            if not result.passed:
                diff_pairs += 1

        total = len(self.pixel_results)
        good = counts["matched"] + counts["same_index"]
        parts = ["{} pairs".format(total) if total else "0 pairs"]
        if good:
            parts.append("{} ✓".format(good))
        if counts["manual_match"]:
            parts.append("{} M manual".format(counts["manual_match"]))
        if counts["low_confidence_match"]:
            parts.append("{} ≈ low-conf".format(counts["low_confidence_match"]))
        if counts["missing_actual"]:
            parts.append("{} ⊘ missing".format(counts["missing_actual"]))
        if counts["extra_actual"]:
            parts.append("{} + extra".format(counts["extra_actual"]))
        if diff_pairs:
            parts.append("{} with pixel diffs".format(diff_pairs))
        if self.object_diffs:
            parts.append("{} object diffs".format(len(self.object_diffs)))
        self.summary_label.setText("  ·  ".join(parts))

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
            gap_penalty=float(self.gap_penalty_spin.value()),
            match_overrides=self.overrides_edit.text().strip() or None,
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
        self.preview._last_image_path = None
        # Stale pixmaps from a previous run would point at PNGs about to be
        # overwritten; drop them so the next selection re-reads fresh data.
        _PIXMAP_CACHE.clear()
        _THUMB_CACHE.clear()
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
        self.open_output_button.setEnabled(bool(output_dir))
        self.export_overrides_button.setEnabled(bool(self.pixel_results))
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("100% - Done")
        self.populate_slide_list()
        self._populate_original_file_lists()
        self._update_summary()
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
        # Suppress per-item currentItemChanged churn while we repopulate; we
        # restore the selection (and let the handler fire) once at the end.
        self.slide_list.blockSignals(True)
        self.slide_list.clear()
        only_diff = self.only_diff_check.isChecked()
        hide_unmatched = self.hide_unmatched_check.isChecked()
        for result in self.pixel_results:
            if hide_unmatched and result.match_status in ("extra_actual", "missing_actual"):
                continue
            if only_diff and result.passed:
                continue
            score_str = "{:.0f}%".format(result.match_score * 100) if result.match_score is not None else "-"
            icon = MATCH_STATUS_ICON.get(result.match_status, "?")
            item = QListWidgetItem(
                "{}  Pair {:03d}  E:{} → A:{}  ·  sim {}  ·  diff {:.4f}%  ·  {} regions".format(
                    icon,
                    result.page,
                    result.expected_page if result.expected_page is not None else "-",
                    result.actual_page if result.actual_page is not None else "-",
                    score_str,
                    result.difference_percent,
                    len(result.regions),
                )
            )
            # Thumbnail: overlay if we have one, else whichever side exists.
            thumb_source = result.output_overlay or result.output_actual or result.output_expected
            if thumb_source:
                thumb = _load_thumbnail_cached(thumb_source, self.slide_list.iconSize())
                if thumb is not None:
                    item.setIcon(QIcon(thumb))
            if result.match_status == "low_confidence_match":
                item.setForeground(QBrush(QColor(180, 100, 0)))  # amber — matched but low sim score
            elif result.match_status == "manual_match":
                item.setForeground(QBrush(QColor(40, 90, 180)))  # blue — user-asserted pairing
            elif result.match_status in ("extra_actual", "missing_actual"):
                item.setForeground(QBrush(QColor(140, 140, 140)))  # gray — unmatched
            item.setData(Qt.UserRole, result)
            self.slide_list.addItem(item)

        self.slide_list.blockSignals(False)
        if self.slide_list.count() > 0:
            self.slide_list.setCurrentRow(0)  # fires slide_selected once
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
        score_str = (
            "{:.2f}% (low confidence — same layout/concept?)".format(result.match_score * 100)
            if result.match_status == "low_confidence_match" and result.match_score is not None
            else "{:.2f}%".format(result.match_score * 100) if result.match_score is not None
            else "-"
        )
        details = [
            "Pair: {}".format(result.page),
            "Expected page: {}".format(result.expected_page if result.expected_page is not None else "-"),
            "Actual page: {}".format(result.actual_page if result.actual_page is not None else "-"),
            "Match status: {}".format(result.match_status),
            "Similarity score: {}".format(score_str),
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
            # The pixel-side slide list keys on the expected slide number; filter
            # object diffs the same way so the panel shows the right rows.
            diffs = [diff for diff in diffs if diff.expected_slide == slide]
        if not diffs:
            return "Object differences: 0"

        lines = ["Object differences: {}".format(len(diffs))]
        for diff in diffs[:80]:
            lines.append(
                "Pair {} | Exp:{} / Act:{} | Object {} | {} | expected={!r} | actual={!r}".format(
                    diff.pair_index,
                    diff.expected_slide if diff.expected_slide is not None else "-",
                    diff.actual_slide if diff.actual_slide is not None else "-",
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
