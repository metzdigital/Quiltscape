#!/usr/bin/env python3
"""
Standalone preview app for Quilt Motion Preview & Export.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

_EXTENSION_DIR = Path(__file__).resolve().parent
_SIDECAR_LIBS = _EXTENSION_DIR / "quilt_motion_exporter_libs"
if _SIDECAR_LIBS.exists():
    sys.path.insert(0, str(_SIDECAR_LIBS))

try:
    from PySide6 import QtCore, QtGui, QtWidgets  # type: ignore
    import PySide6  # type: ignore
except Exception as exc:
    raise SystemExit(f"PySide6 is required to run the preview app: {exc}")

qt_plugin_path = Path(PySide6.__file__).resolve().parent / "Qt" / "plugins"
if qt_plugin_path.exists():
    os.environ.setdefault("QT_PLUGIN_PATH", str(qt_plugin_path))

import quilt_motion_core as qmc

Point = Tuple[float, float]


class PreviewCanvas(QtWidgets.QWidget):
    def __init__(self, controller: "PreviewController") -> None:
        super().__init__()
        self.controller = controller
        self.setMinimumSize(400, 300)

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # noqa: N802
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        self.controller.draw(painter, self.width(), self.height())


class PreviewController(QtCore.QObject):
    BASE_SPEED_MM_PER_SEC = 35.0

    def __init__(self, model: qmc.MotionPathModel, exporters: Dict[str, qmc.ExportProfile]) -> None:
        super().__init__()
        self.model = model
        self.exporters = exporters

        self.progress_mm = 0.0
        self.speed_multiplier = 1.0
        self.playing = True
        self.last_tick: Optional[float] = None
        self._viewport: Optional[Tuple[float, float, float]] = None
        self._static_pixmap: Optional[QtGui.QPixmap] = None
        self._static_size: Tuple[int, int] = (0, 0)
        self._static_dirty = True
        self._updating_progress_slider = False

        # Pantograph defaults
        self.repeat_count = 2
        self.row_count = 2
        base_height_px = max(model.bounds[3] - model.bounds[1], 1e-3)
        self.base_row_distance_mm = base_height_px * model.px_to_mm
        self.row_distance_mm = self.base_row_distance_mm
        self.stagger = False
        self.stagger_percent = 50.0
        self._y_mismatch = False
        self.flip_horizontal = False
        self.flip_vertical = False
        self.mirror_alternate_rows = False
        self.mirror_alternate_rows_vertical = False
        self.export_entire_layout = False

    def set_status_callbacks(self, *, on_status, on_progress, on_y_warning) -> None:
        self.on_status = on_status
        self.on_progress = on_progress
        self.on_y_warning = on_y_warning
        self._refresh_y_warning()

    def tick(self) -> None:
        if not self.playing or not self.model.edges:
            return
        now = time.monotonic()
        if self.last_tick is None:
            self.last_tick = now
            return

        delta = now - self.last_tick
        self.last_tick = now
        advance = delta * self.BASE_SPEED_MM_PER_SEC * self.speed_multiplier
        self.progress_mm += advance
        if self.progress_mm >= self.model.total_length_mm:
            self.progress_mm = self.model.total_length_mm
            self.playing = False

    def draw(self, painter: QtGui.QPainter, width: int, height: int) -> None:
        if width <= 0 or height <= 0:
            return

        self._ensure_static_pixmap(width, height)
        if self._static_pixmap is not None:
            painter.drawPixmap(0, 0, self._static_pixmap)

        if not self.model.edges or self._viewport is None:
            self.on_status("Select at least one path to preview.")
            return

        scale, offset_x, offset_y = self._viewport
        self._draw_progress(painter, scale, offset_x, offset_y)

        if self._y_mismatch and self.model.start_point is not None:
            self._draw_warning_ring(painter, self.model.start_point, scale, offset_x, offset_y)
        if self._y_mismatch and self.model.end_point is not None:
            self._draw_warning_ring(painter, self.model.end_point, scale, offset_x, offset_y)

        point, needle_down = self.model.point_at(self.progress_mm)
        px = self._to_canvas(point, scale, offset_x, offset_y)
        radius = max(3.0, 4.0 * scale * 0.2)
        painter.setBrush(QtGui.QBrush(QtGui.QColor("#e53935" if needle_down else "#f06292")))
        painter.setPen(QtCore.Qt.NoPen)
        painter.drawEllipse(QtCore.QPointF(px[0], px[1]), radius, radius)

        stitched = min(self.progress_mm, self.model.total_length_mm)
        percent = (stitched / self.model.total_length_mm * 100.0) if self.model.total_length_mm else 0.0
        if not self._updating_progress_slider:
            self._updating_progress_slider = True
            self.on_progress(percent)
            self._updating_progress_slider = False

        self.on_status(
            f"Path length: {self.model.total_length_mm:.1f} mm   Previewed: {stitched:.1f} mm ({percent:.1f}%)"
        )

    def _ensure_static_pixmap(self, width: int, height: int) -> None:
        needs_new = self._static_pixmap is None or self._static_dirty or self._static_size != (width, height)
        if not needs_new:
            return

        pixmap = QtGui.QPixmap(width, height)
        pixmap.fill(QtGui.QColor("#f8f8f8"))
        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)

        if self.model.edges:
            self._viewport = self._compute_viewport(width, height)
            scale, offset_x, offset_y = self._viewport

            min_x, min_y, max_x, max_y = self._layout_bounds()
            tl = self._to_canvas((min_x, min_y), scale, offset_x, offset_y)
            br = self._to_canvas((max_x, max_y), scale, offset_x, offset_y)
            clip_rect = QtCore.QRectF(tl[0], tl[1], br[0] - tl[0], br[1] - tl[1])
            painter.save()
            painter.setClipRect(clip_rect)
            self._draw_full_pattern(painter, scale, offset_x, offset_y)
            painter.restore()
        else:
            self._viewport = (1.0, 0.0, 0.0)

        painter.end()
        self._static_pixmap = pixmap
        self._static_size = (width, height)
        self._static_dirty = False

    def _pantograph_offsets(self) -> List[Tuple[int, float, float]]:
        width_px = max(self.model.bounds[2] - self.model.bounds[0], 1e-3)
        height_px = max(self.model.bounds[3] - self.model.bounds[1], 1e-3)
        self._pattern_width_px = width_px
        self._pattern_height_px = height_px

        return qmc._compute_pantograph_offsets(
            self.model.bounds,
            repeat_count=self.repeat_count,
            row_count=self.row_count,
            row_distance_mm=self.row_distance_mm,
            px_to_mm=self.model.px_to_mm,
            stagger=self.stagger,
            stagger_percent=self.stagger_percent,
            start_point=self.model.start_point,
            end_point=self.model.end_point,
        )

    def _layout_bounds(self) -> Tuple[float, float, float, float]:
        return qmc._compute_layout_bounds(
            self.model.bounds,
            repeat_count=self.repeat_count,
            row_count=self.row_count,
            row_distance_mm=self.row_distance_mm,
            px_to_mm=self.model.px_to_mm,
            start_point=self.model.start_point,
            end_point=self.model.end_point,
        )

    def _build_export_model(self) -> qmc.MotionPathModel:
        layout_bounds = self._layout_bounds()
        raw_offsets = self._pantograph_offsets()
        offsets_by_row: Dict[int, List[Tuple[float, float]]] = {}
        for row_idx, dx, dy in raw_offsets:
            offsets_by_row.setdefault(row_idx, []).append((dx, dy))
        offsets: List[Tuple[int, float, float]] = []
        for row_idx in sorted(offsets_by_row.keys()):
            entries = sorted(offsets_by_row[row_idx], key=lambda v: v[0])
            if row_idx % 2 == 1:
                entries.reverse()
            for dx, dy in entries:
                offsets.append((row_idx, dx, dy))

        stitched_segments: List[qmc.MotionSegment] = []
        last_end: Optional[Point] = None

        def _close_enough(a: Point, b: Point, tol: float = 1e-6) -> bool:
            return math.isclose(a[0], b[0], abs_tol=tol) and math.isclose(a[1], b[1], abs_tol=tol)

        def _clip_segment(p0: Point, p1: Point) -> Optional[Tuple[Point, Point]]:
            min_x, min_y, max_x, max_y = layout_bounds
            dx = p1[0] - p0[0]
            dy = p1[1] - p0[1]
            u1, u2 = 0.0, 1.0

            def _update(p: float, q: float) -> bool:
                nonlocal u1, u2
                if math.isclose(p, 0.0, abs_tol=1e-12):
                    return q >= 0.0
                t = q / p
                if p < 0:
                    if t > u2:
                        return False
                    if t > u1:
                        u1 = t
                else:
                    if t < u1:
                        return False
                    if t < u2:
                        u2 = t
                return True

            if not (
                _update(-dx, p0[0] - min_x)
                and _update(dx, max_x - p0[0])
                and _update(-dy, p0[1] - min_y)
                and _update(dy, max_y - p0[1])
            ):
                return None
            clipped_start = (p0[0] + u1 * dx, p0[1] + u1 * dy)
            clipped_end = (p0[0] + u2 * dx, p0[1] + u2 * dy)
            return clipped_start, clipped_end

        def _clip_polyline(points: List[Point]) -> List[Point]:
            clipped: List[Point] = []
            for idx in range(1, len(points)):
                result = _clip_segment(points[idx - 1], points[idx])
                if result is None:
                    continue
                a, b = result
                if not clipped:
                    clipped.append(a)
                else:
                    if not _close_enough(clipped[-1], a):
                        clipped.append(a)
                clipped.append(b)
            return clipped

        for row_idx, dx, dy in offsets:
            mirror_row_h = (row_idx % 2 == 1) or (self.mirror_alternate_rows and (row_idx % 2 == 1))
            mirror_row_v = self.mirror_alternate_rows_vertical and (row_idx % 2 == 1)
            for seg in self.model.segments:
                pts: List[Point] = []
                for pt in seg.points:
                    tx, ty = self._transform_point(pt, mirror_row_h, mirror_row_v)
                    pts.append((tx + dx, ty + dy))
                if len(pts) < 2:
                    continue
                pts = _clip_polyline(pts)
                if len(pts) < 2:
                    continue
                if last_end is not None and not _close_enough(last_end, pts[0]):
                    stitched_segments.append(qmc.MotionSegment(points=[last_end, pts[0]], needle_down=True))
                stitched_segments.append(qmc.MotionSegment(points=pts, needle_down=seg.needle_down))
                last_end = pts[-1]

        if not stitched_segments:
            return self.model

        return qmc.MotionPathModel(
            stitched_segments,
            px_to_mm=self.model.px_to_mm,
            doc_height_px=self.model.doc_height_px,
        )

    def _transform_point(self, point: Point, mirror_row_h: bool, mirror_row_v: bool) -> Point:
        min_x, min_y, max_x, max_y = self.model.bounds
        cx = (min_x + max_x) / 2.0
        cy = (min_y + max_y) / 2.0
        x, y = point
        if self.flip_horizontal or mirror_row_h:
            x = 2 * cx - x
        if self.flip_vertical or mirror_row_v:
            y = 2 * cy - y
        return (x, y)

    def _compute_viewport(self, width: int, height: int) -> Tuple[float, float, float]:
        min_x, min_y, max_x, max_y = self._layout_bounds()
        margin = 0.05
        span_x = max(max_x - min_x, 1e-3)
        span_y = max(max_y - min_y, 1e-3)
        scale_x = width * (1.0 - margin) / span_x
        scale_y = height * (1.0 - margin) / span_y
        scale = min(scale_x, scale_y)
        offset_x = (width - span_x * scale) / 2.0 - min_x * scale
        offset_y = (height - span_y * scale) / 2.0 - min_y * scale
        return scale, offset_x, offset_y

    def _to_canvas(self, point: Point, scale: float, offset_x: float, offset_y: float) -> Point:
        return (point[0] * scale + offset_x, point[1] * scale + offset_y)

    def _stroke_width(self) -> float:
        span = max(
            self.model.bounds[2] - self.model.bounds[0],
            self.model.bounds[3] - self.model.bounds[1],
        )
        if span <= 0:
            return 1.0
        width = (span / 300.0) ** 0.7
        return max(min(width, 10.0), 0.08)

    def _draw_full_pattern(self, painter: QtGui.QPainter, scale: float, offset_x: float, offset_y: float) -> None:
        line_width = max(1.0, self._stroke_width() * scale)
        pen = QtGui.QPen()
        pen.setWidthF(line_width)
        offsets = self._pantograph_offsets()

        for row_idx, dx, dy in offsets:
            mirror_row_h = self.mirror_alternate_rows and (row_idx % 2 == 1)
            mirror_row_v = self.mirror_alternate_rows_vertical and (row_idx % 2 == 1)
            for seg in self.model.segments:
                color = QtGui.QColor("#2b6cb0" if seg.needle_down else "#d14343")
                pen.setColor(color)
                painter.setPen(pen)
                pts = [self._transform_point(pt, mirror_row_h, mirror_row_v) for pt in seg.points]
                pts = [(p[0] + dx, p[1] + dy) for p in pts]
                if len(pts) < 2:
                    continue
                path = QtGui.QPainterPath()
                first = self._to_canvas(pts[0], scale, offset_x, offset_y)
                path.moveTo(first[0], first[1])
                for pt in pts[1:]:
                    cx, cy = self._to_canvas(pt, scale, offset_x, offset_y)
                    path.lineTo(cx, cy)
                painter.drawPath(path)

    def _rgba_to_qcolor(self, rgba: Tuple[float, float, float, float]) -> QtGui.QColor:
        r, g, b, a = rgba
        color = QtGui.QColor(int(r * 255), int(g * 255), int(b * 255))
        color.setAlphaF(a)
        return color

    def _draw_progress(self, painter: QtGui.QPainter, scale: float, offset_x: float, offset_y: float) -> None:
        line_width = max(1.0, self._stroke_width() * 1.8 * scale)
        remaining = self.progress_mm
        transform = lambda p: self._transform_point(p, False, False)

        cell_size = max(self._stroke_width() * 0.01, 1e-6)
        coverage: Dict[Tuple[int, int], int] = {}

        def _sample_cells(p0: Point, p1: Point) -> List[Tuple[int, int]]:
            length = math.dist(p0, p1)
            if length <= 1e-12:
                gx = int(round(p0[0] / cell_size))
                gy = int(round(p0[1] / cell_size))
                return [(gx, gy)]
            steps = max(1, int(math.ceil(length / cell_size)))
            coords: List[Tuple[int, int]] = []
            for s in range(steps + 1):
                t = s / steps
                x = p0[0] + (p1[0] - p0[0]) * t
                y = p0[1] + (p1[1] - p0[1]) * t
                gx = int(round(x / cell_size))
                gy = int(round(y / cell_size))
                if not coords or coords[-1] != (gx, gy):
                    coords.append((gx, gy))
            return coords

        def _subdivide_and_draw(edge: qmc.MotionEdge) -> bool:
            if remaining <= edge.start_length_mm:
                return True
            length_remaining = min(remaining - edge.start_length_mm, edge.length_mm)
            if length_remaining <= 0:
                return True

            chunks = max(1, min(12, int(math.ceil(edge.length_mm / max(edge.length_mm / 5.0, 1e-6)))))
            chunk_len = edge.length_mm / chunks if chunks else edge.length_mm

            edge_cells: List[Tuple[int, int]] = []

            for i in range(chunks):
                seg_start_mm = i * chunk_len
                seg_draw_mm = min(chunk_len, max(0.0, length_remaining - seg_start_mm))
                if seg_draw_mm <= 0:
                    continue
                t0 = (seg_start_mm) / edge.length_mm if edge.length_mm else 0.0
                t1 = (seg_start_mm + seg_draw_mm) / edge.length_mm if edge.length_mm else 0.0
                t0 = max(0.0, min(1.0, t0))
                t1 = max(0.0, min(1.0, t1))
                seg_start = (
                    edge.start_px[0] + (edge.end_px[0] - edge.start_px[0]) * t0,
                    edge.start_px[1] + (edge.end_px[1] - edge.start_px[1]) * t0,
                )
                seg_end = (
                    edge.start_px[0] + (edge.end_px[0] - edge.start_px[0]) * t1,
                    edge.start_px[1] + (edge.end_px[1] - edge.start_px[1]) * t1,
                )

                draw_start = transform(seg_start)
                draw_end = transform(seg_end)

                cells = _sample_cells(draw_start, draw_end)
                if cells:
                    overlap_hits = [coverage.get(c, 0) for c in cells]
                    covered = [v for v in overlap_hits if v > 0]
                    threshold = max(1, int(len(cells) * 0.05))
                    passes_completed = max(covered) if len(covered) >= threshold else 0
                else:
                    passes_completed = 0

                color = self._rgba_to_qcolor(qmc._color_for_pass(passes_completed, edge.needle_down))
                pen = QtGui.QPen(color)
                pen.setWidthF(line_width)
                painter.setPen(pen)

                start_px = self._to_canvas(draw_start, scale, offset_x, offset_y)
                end_px = self._to_canvas(draw_end, scale, offset_x, offset_y)
                painter.drawLine(QtCore.QPointF(*start_px), QtCore.QPointF(*end_px))

                edge_cells.extend(cells)

                if seg_draw_mm + seg_start_mm + 1e-9 < length_remaining:
                    continue
                if length_remaining + 1e-9 < edge.length_mm:
                    for cell in edge_cells:
                        coverage[cell] = coverage.get(cell, 0) + 1
                    return False

            for cell in edge_cells:
                coverage[cell] = coverage.get(cell, 0) + 1

            return True

        for edge in self.model.edges:
            if not _subdivide_and_draw(edge):
                break

    def _draw_warning_ring(self, painter: QtGui.QPainter, point: Point, scale: float, offset_x: float, offset_y: float) -> None:
        px = self._to_canvas(point, scale, offset_x, offset_y)
        radius = max(6.0, 9.0 * scale * 0.2)
        pen = QtGui.QPen(QtGui.QColor("#b8860b"))
        pen.setWidthF(max(1.0, 2.0 * scale * 0.5))
        painter.setPen(pen)
        painter.setBrush(QtCore.Qt.NoBrush)
        painter.drawEllipse(QtCore.QPointF(px[0], px[1]), radius, radius)

    def _refresh_y_warning(self) -> None:
        start = self.model.start_point
        end = self.model.end_point
        delta_mm = 0.0
        mismatch = False
        if start is not None and end is not None:
            delta_mm = abs(start[1] - end[1]) * self.model.px_to_mm
            mismatch = delta_mm > 0.1 + 1e-9
        self._y_mismatch = mismatch
        if mismatch:
            message = (
                "WARNING: Start node and end node have different Y-axis positions "
                f"(dY = {delta_mm:.3f} mm > 0.1mm)"
            )
            self.on_y_warning(message)
        else:
            self.on_y_warning("")


class PreviewWindow(QtWidgets.QMainWindow):
    def __init__(self, controller: PreviewController) -> None:
        super().__init__()
        self.controller = controller
        self.setWindowTitle("Quilt Motion Preview")
        self.resize(1100, 600)

        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        layout = QtWidgets.QHBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        left_column = QtWidgets.QVBoxLayout()
        layout.addLayout(left_column, stretch=1)

        self.canvas = PreviewCanvas(controller)
        left_column.addWidget(self.canvas, stretch=1)

        self.warning_label = QtWidgets.QLabel("")
        self.warning_label.setStyleSheet("color: #b8860b;")
        left_column.addWidget(self.warning_label, stretch=0)

        sidebar = QtWidgets.QVBoxLayout()
        layout.addLayout(sidebar, stretch=0)

        controls = QtWidgets.QHBoxLayout()
        sidebar.addLayout(controls)

        self.play_button = QtWidgets.QPushButton("Pause")
        self.play_button.clicked.connect(self._toggle_play)
        controls.addWidget(self.play_button)

        restart_button = QtWidgets.QPushButton("Restart")
        restart_button.clicked.connect(self._restart)
        controls.addWidget(restart_button)

        self.optimize_button = QtWidgets.QPushButton("Optimize path")
        self.optimize_button.clicked.connect(self._optimize_path)
        controls.addWidget(self.optimize_button)

        sidebar.addSpacing(6)
        sidebar.addWidget(QtWidgets.QLabel("Preview speed"))
        self.speed_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.speed_slider.setRange(10, 500)
        self.speed_slider.setValue(100)
        self.speed_slider.valueChanged.connect(self._on_speed_changed)
        sidebar.addWidget(self.speed_slider)
        self.speed_value_label = QtWidgets.QLabel("1.00×")
        sidebar.addWidget(self.speed_value_label)

        sidebar.addSpacing(6)
        sidebar.addWidget(QtWidgets.QLabel("Preview progress"))
        self.progress_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.progress_slider.setRange(0, 1000)
        self.progress_slider.valueChanged.connect(self._on_progress_changed)
        sidebar.addWidget(self.progress_slider)

        self.preview_status = QtWidgets.QLabel("")
        sidebar.addWidget(self.preview_status)

        sidebar.addSpacing(8)
        sidebar.addWidget(self._separator())

        sidebar.addWidget(QtWidgets.QLabel("Pantograph layout"))
        pantograph = QtWidgets.QGridLayout()
        sidebar.addLayout(pantograph)

        self.repeat_spin = QtWidgets.QSpinBox()
        self.repeat_spin.setRange(1, 20)
        self.repeat_spin.setValue(self.controller.repeat_count)
        self.repeat_spin.valueChanged.connect(self._on_repeat_changed)
        pantograph.addWidget(QtWidgets.QLabel("Repeats"), 0, 0)
        pantograph.addWidget(self.repeat_spin, 0, 1)

        self.rows_spin = QtWidgets.QSpinBox()
        self.rows_spin.setRange(1, 20)
        self.rows_spin.setValue(self.controller.row_count)
        self.rows_spin.valueChanged.connect(self._on_rows_changed)
        pantograph.addWidget(QtWidgets.QLabel("Rows"), 1, 0)
        pantograph.addWidget(self.rows_spin, 1, 1)

        self.row_distance_spin = QtWidgets.QDoubleSpinBox()
        self.row_distance_spin.setRange(-5000.0, 5000.0)
        self.row_distance_spin.setValue(self.controller.row_distance_mm)
        self.row_distance_spin.setSingleStep(1.0)
        self.row_distance_spin.valueChanged.connect(self._on_row_distance_changed)
        pantograph.addWidget(QtWidgets.QLabel("Row distance (mm)"), 2, 0)
        pantograph.addWidget(self.row_distance_spin, 2, 1)

        self.stagger_toggle = QtWidgets.QCheckBox("Stagger alternate rows")
        self.stagger_toggle.setChecked(self.controller.stagger)
        self.stagger_toggle.toggled.connect(self._on_stagger_toggled)
        pantograph.addWidget(self.stagger_toggle, 3, 0, 1, 2)

        self.stagger_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.stagger_slider.setRange(0, 100)
        self.stagger_slider.setValue(int(self.controller.stagger_percent))
        self.stagger_slider.valueChanged.connect(self._on_stagger_percent_changed)
        pantograph.addWidget(QtWidgets.QLabel("Stagger %"), 4, 0)
        pantograph.addWidget(self.stagger_slider, 4, 1)

        self.mirror_toggle = QtWidgets.QCheckBox("Mirror every other row horizontally")
        self.mirror_toggle.setChecked(self.controller.mirror_alternate_rows)
        self.mirror_toggle.toggled.connect(self._on_mirror_rows_toggled)
        pantograph.addWidget(self.mirror_toggle, 5, 0, 1, 2)

        self.mirror_v_toggle = QtWidgets.QCheckBox("Mirror every other row vertically")
        self.mirror_v_toggle.setChecked(self.controller.mirror_alternate_rows_vertical)
        self.mirror_v_toggle.toggled.connect(self._on_mirror_rows_v_toggled)
        pantograph.addWidget(self.mirror_v_toggle, 6, 0, 1, 2)

        self.flip_h_toggle = QtWidgets.QCheckBox("Flip horizontally")
        self.flip_h_toggle.setChecked(self.controller.flip_horizontal)
        self.flip_h_toggle.toggled.connect(self._on_flip_h_toggled)
        pantograph.addWidget(self.flip_h_toggle, 7, 0, 1, 2)

        self.flip_v_toggle = QtWidgets.QCheckBox("Flip vertically")
        self.flip_v_toggle.setChecked(self.controller.flip_vertical)
        self.flip_v_toggle.toggled.connect(self._on_flip_v_toggled)
        pantograph.addWidget(self.flip_v_toggle, 8, 0, 1, 2)

        sidebar.addSpacing(6)
        sidebar.addWidget(QtWidgets.QLabel("Export format"))
        export_row = QtWidgets.QHBoxLayout()
        sidebar.addLayout(export_row)

        self.format_combo = QtWidgets.QComboBox()
        for key, profile in self.controller.exporters.items():
            self.format_combo.addItem(f"{key} – {profile.title}", key)
        export_row.addWidget(self.format_combo)

        self.export_layout_toggle = QtWidgets.QCheckBox("Export entire layout")
        self.export_layout_toggle.setChecked(self.controller.export_entire_layout)
        self.export_layout_toggle.toggled.connect(self._on_export_layout_toggled)
        export_row.addWidget(self.export_layout_toggle)

        self.export_button = QtWidgets.QPushButton("Export…")
        self.export_button.clicked.connect(self._export)
        sidebar.addWidget(self.export_button)

        self.export_status = QtWidgets.QLabel("")
        sidebar.addWidget(self.export_status)
        sidebar.addStretch(1)

        controller.set_status_callbacks(
            on_status=self.preview_status.setText,
            on_progress=self._set_progress_slider_value,
            on_y_warning=self.warning_label.setText,
        )

        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(16)

    def _separator(self) -> QtWidgets.QFrame:
        sep = QtWidgets.QFrame()
        sep.setFrameShape(QtWidgets.QFrame.HLine)
        sep.setFrameShadow(QtWidgets.QFrame.Sunken)
        return sep

    def _tick(self) -> None:
        self.controller.tick()
        if self.controller.playing:
            self.play_button.setText("Pause")
        self.canvas.update()

    def _toggle_play(self) -> None:
        if (
            not self.controller.playing
            and self.controller.model.total_length_mm > 0
            and math.isclose(self.controller.progress_mm, self.controller.model.total_length_mm, abs_tol=1e-6)
        ):
            self.controller.progress_mm = 0.0
        self.controller.playing = not self.controller.playing
        self.controller.last_tick = time.monotonic()
        self.play_button.setText("Pause" if self.controller.playing else "Play")
        self.canvas.update()

    def _restart(self) -> None:
        self.controller.progress_mm = 0.0
        self.controller.playing = False
        self.controller.last_tick = time.monotonic()
        self.play_button.setText("Play")
        self.preview_status.setText("Preview reset. Press Play to start.")
        self.canvas.update()

    def _optimize_path(self) -> None:
        if not self.controller.model.segments:
            self.preview_status.setText("No path to optimize.")
            return
        try:
            optimized_segments = qmc.optimize_motion_segments(
                self.controller.model.segments,
                start_point=self.controller.model.start_point,
                end_point=self.controller.model.end_point,
            )
        except Exception as exc:
            self.preview_status.setText(f"Optimization failed: {exc}")
            return

        if optimized_segments is self.controller.model.segments or optimized_segments == self.controller.model.segments:
            self.preview_status.setText("Path is already optimised.")
            return

        self.controller.model = qmc.MotionPathModel(
            optimized_segments,
            px_to_mm=self.controller.model.px_to_mm,
            doc_height_px=self.controller.model.doc_height_px,
        )
        base_height_px = max(self.controller.model.bounds[3] - self.controller.model.bounds[1], 1e-3)
        self.controller.base_row_distance_mm = base_height_px * self.controller.model.px_to_mm
        self.controller.row_distance_mm = self.controller.base_row_distance_mm
        self.row_distance_spin.setValue(self.controller.row_distance_mm)
        self.controller._static_dirty = True
        self.controller._refresh_y_warning()

        self.controller.progress_mm = 0.0
        self.controller.playing = False
        self.play_button.setText("Play")
        self.controller.last_tick = time.monotonic()
        self.preview_status.setText("Path optimised to reduce overlaps.")
        self.canvas.update()

    def _on_speed_changed(self, value: int) -> None:
        self.controller.speed_multiplier = max(0.1, value / 100.0)
        self.speed_value_label.setText(f"{self.controller.speed_multiplier:.2f}×")

    def _on_progress_changed(self, value: int) -> None:
        if self.controller._updating_progress_slider or not self.controller.model.total_length_mm:
            return
        percent = value / 10.0
        self.controller.progress_mm = percent / 100.0 * self.controller.model.total_length_mm
        self.controller.last_tick = time.monotonic()
        self.canvas.update()

    def _set_progress_slider_value(self, percent: float) -> None:
        self.controller._updating_progress_slider = True
        self.progress_slider.setValue(int(max(0.0, min(100.0, percent)) * 10.0))
        self.controller._updating_progress_slider = False

    def _on_repeat_changed(self, value: int) -> None:
        self.controller.repeat_count = max(1, value)
        self.controller._static_dirty = True
        self.canvas.update()

    def _on_rows_changed(self, value: int) -> None:
        self.controller.row_count = max(1, value)
        self.controller._static_dirty = True
        self.canvas.update()

    def _on_row_distance_changed(self, value: float) -> None:
        self.controller.row_distance_mm = float(value)
        self.controller._static_dirty = True
        self.canvas.update()

    def _on_stagger_toggled(self, checked: bool) -> None:
        self.controller.stagger = checked
        self.controller._static_dirty = True
        self.canvas.update()

    def _on_stagger_percent_changed(self, value: int) -> None:
        self.controller.stagger_percent = float(value)
        if self.controller.stagger:
            self.controller._static_dirty = True
            self.canvas.update()

    def _on_mirror_rows_toggled(self, checked: bool) -> None:
        self.controller.mirror_alternate_rows = checked
        self.controller._static_dirty = True
        self.canvas.update()

    def _on_mirror_rows_v_toggled(self, checked: bool) -> None:
        self.controller.mirror_alternate_rows_vertical = checked
        self.controller._static_dirty = True
        self.canvas.update()

    def _on_flip_h_toggled(self, checked: bool) -> None:
        self.controller.flip_horizontal = checked
        self.controller._static_dirty = True
        self.canvas.update()

    def _on_flip_v_toggled(self, checked: bool) -> None:
        self.controller.flip_vertical = checked
        self.controller._static_dirty = True
        self.canvas.update()

    def _on_export_layout_toggled(self, checked: bool) -> None:
        self.controller.export_entire_layout = checked

    def _build_export_model(self) -> qmc.MotionPathModel:
        return self.controller._build_export_model()  # type: ignore[attr-defined]

    def _export(self) -> None:
        index = self.format_combo.currentIndex()
        if index < 0:
            return
        key = self.format_combo.itemData(index)
        profile = self.controller.exporters[key]

        filename, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export Motion Path",
            str(Path.home() / f"quilt_path.{profile.extension}"),
            f"{profile.title} (*.{profile.extension})",
        )
        if not filename:
            return

        out_path = Path(filename)
        if out_path.suffix.lower() != f".{profile.extension.lower()}":
            out_path = out_path.with_suffix(f".{profile.extension.lower()}")

        try:
            out_path.parent.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            self.export_status.setText(f"No permission to create folder: {out_path.parent}")
            return
        except FileExistsError:
            pass

        try:
            export_model = self.controller._build_export_model() if self.controller.export_entire_layout else self.controller.model
            profile.writer(export_model, out_path)
        except PermissionError:
            self.export_status.setText(f"Cannot write to {out_path}. Pick a folder inside your home directory.")
            return
        except Exception as exc:
            self.export_status.setText(f"Export failed: {exc}")
            return

        self.export_status.setText(f"Exported to {out_path}")
        self.controller._static_dirty = True
        self.canvas.update()


def load_payload(path: Path) -> qmc.MotionPathModel:
    payload = json.loads(path.read_text())
    segments: List[qmc.MotionSegment] = []
    for seg in payload.get("segments", []):
        points = [(float(x), float(y)) for x, y in seg.get("points", [])]
        segments.append(qmc.MotionSegment(points=points, needle_down=bool(seg.get("needle_down", True))))
    px_to_mm = float(payload.get("px_to_mm", 1.0))
    doc_height_px = payload.get("doc_height_px")
    if doc_height_px is not None:
        doc_height_px = float(doc_height_px)
    return qmc.MotionPathModel(segments, px_to_mm=px_to_mm, doc_height_px=doc_height_px)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Quilt Motion Preview App")
    parser.add_argument("--input", type=Path, required=True, help="Path to preview JSON data.")
    parser.add_argument("--delete-input", action="store_true", help="Remove the input file after loading.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    model = load_payload(args.input)
    if args.delete_input:
        try:
            args.input.unlink()
        except Exception:
            pass

    app = QtWidgets.QApplication(sys.argv)
    controller = PreviewController(model, qmc.EXPORT_PROFILES)
    window = PreviewWindow(controller)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
