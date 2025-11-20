#!/usr/bin/env python3
"""
Inkscape Quilt Motion Preview & Export extension.

This script gathers the selected path elements from the active document,
preserves their draw order, animates them in a GTK preview window, and
exports the resulting stitch path to several long‑arm quilting formats.
"""

from __future__ import annotations

import math
import time
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import inkex
from inkex import bezier, units
from inkex.elements import PathElement
from inkex.localization import inkex_gettext as _
from inkex.paths import CubicSuperPath
from PIL import Image, ImageDraw

warnings.filterwarnings(
    "ignore",
    message="DynamicImporter\\.exec_module\\(\\) not found; falling back to load_module\\(\\)",
    category=ImportWarning,
)

GTK_AVAILABLE = False
GTK_LOAD_ERROR: Optional[str] = None

try:
    import gi  # type: ignore

    gi.require_version("Gtk", "3.0")
    gi.require_foreign("cairo")

    gi.require_version("Pango", "1.0")
    from gi.repository import GLib, Gtk, Pango, cairo as gi_cairo  # type: ignore
    import cairo  # type: ignore

    GTK_AVAILABLE = True
except Exception as exc:  # pragma: no cover - Gtk is only available inside Inkscape
    GTK_AVAILABLE = False
    GTK_LOAD_ERROR = str(exc)
    GLib = None  # type: ignore
    Gtk = None  # type: ignore


Point = Tuple[float, float]


@dataclass
class MotionSegment:
    """Represents a contiguous collection of points following the same stitch state."""

    points: List[Point]
    needle_down: bool = True


@dataclass
class MotionEdge:
    """Single linear edge in the flattened stitch sequence."""

    start_px: Point
    end_px: Point
    needle_down: bool
    start_length_mm: float
    length_mm: float

    @property
    def end_length_mm(self) -> float:
        return self.start_length_mm + self.length_mm


class MotionPathModel:
    """Holds a stitched path flattened to monotonic, ordered geometry."""

    def __init__(self, segments: List[MotionSegment], px_to_mm: float, doc_height_px: Optional[float] = None) -> None:
        stitched = [seg for seg in segments if len(seg.points) >= 2]
        self.segments = stitched
        self.px_to_mm = px_to_mm
        self.doc_height_px = doc_height_px
        self.doc_height_mm = (doc_height_px * px_to_mm) if doc_height_px else None
        self.edges: List[MotionEdge] = []
        self.total_length_mm = 0.0
        self.start_point: Optional[Point] = None
        self.end_point: Optional[Point] = None

        for idx, seg in enumerate(stitched):
            pts = seg.points
            for i in range(1, len(pts)):
                start = pts[i - 1]
                end = pts[i]
                delta = math.dist(start, end)
                length_mm = delta * px_to_mm
                edge = MotionEdge(
                    start_px=start,
                    end_px=end,
                    needle_down=seg.needle_down,
                    start_length_mm=self.total_length_mm,
                    length_mm=length_mm,
                )
                self.edges.append(edge)
                self.total_length_mm += length_mm

            if idx == 0:
                self.start_point = pts[0]
            self.end_point = pts[-1]

        xs = [pt[0] for seg in stitched for pt in seg.points]
        ys = [pt[1] for seg in stitched for pt in seg.points]
        if xs and ys:
            self.bounds = (min(xs), min(ys), max(xs), max(ys))
        else:
            self.bounds = (0.0, 0.0, 1.0, 1.0)

    def iter_segments_mm(self) -> List[Tuple[bool, List[Point]]]:
        """Return the motion path as absolute millimetre coordinates."""
        factor = self.px_to_mm
        converted: List[Tuple[bool, List[Point]]] = []
        for seg in self.segments:
            pts = [(x * factor, y * factor) for x, y in seg.points]
            converted.append((seg.needle_down, pts))
        return converted

    def point_at(self, length_mm: float) -> Tuple[Point, bool]:
        """Return the cartesian point and stitch state at a given cumulative length."""
        if not self.edges:
            return (0.0, 0.0), True
        clamped = max(0.0, min(length_mm, self.total_length_mm))
        for edge in self.edges:
            if clamped <= edge.end_length_mm or math.isclose(
                clamped, edge.end_length_mm, abs_tol=1e-6
            ):
                if edge.length_mm == 0:
                    return edge.end_px, edge.needle_down
                ratio = (clamped - edge.start_length_mm) / edge.length_mm
                x = edge.start_px[0] + (edge.end_px[0] - edge.start_px[0]) * ratio
                y = edge.start_px[1] + (edge.end_px[1] - edge.start_px[1]) * ratio
                return (x, y), edge.needle_down
        last_edge = self.edges[-1]
        return last_edge.end_px, last_edge.needle_down


def _flatten_path_element(
    element: PathElement, tolerance: float = 0.5
) -> List[MotionSegment]:
    """Convert an Inkscape path element into flattened motion segments."""
    transform = element.composed_transform()
    path = element.path.transform(transform)
    csp: CubicSuperPath = path.to_superpath()
    bezier.cspsubdiv(csp, tolerance)

    segments: List[MotionSegment] = []
    previous_end: Optional[Point] = None
    for subpath in csp:
        points = [tuple(node[1]) for node in subpath]
        if len(points) < 2:
            continue
        # Remove duplicate closing node
        if math.isclose(points[0][0], points[-1][0], abs_tol=1e-9) and math.isclose(
            points[0][1], points[-1][1], abs_tol=1e-9
        ):
            points = points[:-1]
        if not points:
            continue
        if previous_end and (
            not math.isclose(previous_end[0], points[0][0], abs_tol=1e-6)
            or not math.isclose(previous_end[1], points[0][1], abs_tol=1e-6)
        ):
            segments.append(MotionSegment(points=[previous_end, points[0]], needle_down=False))
        segments.append(MotionSegment(points=list(points), needle_down=True))
        previous_end = points[-1]
    return segments


class ExportProfile:
    """Holds metadata about every supported export format."""

    def __init__(
        self,
        title: str,
        extension: str,
        description: str,
        writer: Callable[[MotionPathModel, Path], None],
    ) -> None:
        self.title = title
        self.extension = extension
        self.description = description
        self.writer = writer


if GTK_AVAILABLE:
    class QuiltPreviewWindow(Gtk.Window):  # type: ignore[misc]
        """Interactive GTK window that previews and exports the motion path."""

        BASE_SPEED_MM_PER_SEC = 35.0

        def __init__(
            self,
            model: MotionPathModel,
            exporters: Dict[str, ExportProfile],
        ) -> None:
            super().__init__(title=_("Quilt Motion Preview"))
            self.set_default_size(1100, 600)
            self.model = model
            self.exporters = exporters

            self.progress_mm = 0.0
            self.speed_multiplier = 1.0
            self.playing = True
            self.last_tick: Optional[float] = None
            self._timeout_id: Optional[int] = None
            self._static_surface: Optional[cairo.ImageSurface] = None
            self._static_surface_size: Tuple[int, int] = (0, 0)
            self._surface_dirty = True
            self._viewport: Optional[Tuple[float, float, float]] = None
            self._prime_source: Optional[int] = None
            self._updating_progress_slider = False

            # Pantograph defaults
            self.repeat_count = 2
            self.row_count = 2
            base_height_px = max(model.bounds[3] - model.bounds[1], 1e-3)
            self.base_row_distance_mm = base_height_px * model.px_to_mm
            self.row_distance_mm = self.base_row_distance_mm
            self.stagger = False
            self.stagger_percent = 50.0

            self._build_ui()
            self.connect("destroy", self._on_destroy)
            self._timeout_id = GLib.timeout_add(16, self._tick)
            self._prime_source = GLib.idle_add(self._prime_surface)

        def _build_ui(self) -> None:
            root = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=16)
            root.set_margin_top(12)
            root.set_margin_bottom(12)
            root.set_margin_start(12)
            root.set_margin_end(12)
            self.add(root)

            self.drawing_area = Gtk.DrawingArea()
            self.drawing_area.set_hexpand(True)
            self.drawing_area.set_vexpand(True)
            self.drawing_area.connect("draw", self._on_draw)
            self.drawing_area.connect("size-allocate", self._on_size_allocate)
            root.pack_start(self.drawing_area, True, True, 0)

            sidebar = Gtk.Box(
                orientation=Gtk.Orientation.VERTICAL, spacing=8
            )
            sidebar.set_size_request(280, -1)
            root.pack_start(sidebar, False, False, 0)

            controls = Gtk.Box(spacing=6)
            sidebar.pack_start(controls, False, False, 0)

            self.play_button = Gtk.Button(label=_("Pause"))
            self.play_button.connect("clicked", self._toggle_play)
            controls.pack_start(self.play_button, True, True, 0)

            restart_button = Gtk.Button(label=_("Restart"))
            restart_button.connect("clicked", self._restart)
            controls.pack_start(restart_button, True, True, 0)

            speed_label = Gtk.Label(label=_("Preview speed"))
            speed_label.set_xalign(0.0)
            sidebar.pack_start(speed_label, False, False, 0)

            adjustment = Gtk.Adjustment(
                value=1.0,
                lower=0.1,
                upper=5.0,
                step_increment=0.1,
                page_increment=0.5,
                page_size=0.0,
            )
            self.speed_slider = Gtk.Scale(
                orientation=Gtk.Orientation.HORIZONTAL, adjustment=adjustment
            )
            self.speed_slider.connect("value-changed", self._on_speed_changed)
            self.speed_slider.set_value(1.0)
            sidebar.pack_start(self.speed_slider, False, False, 0)

            self.speed_value_label = Gtk.Label(label=_("1.00×"))
            self.speed_value_label.set_xalign(0.0)
            sidebar.pack_start(self.speed_value_label, False, False, 0)

            progress_label = Gtk.Label(label=_("Preview progress"))
            progress_label.set_xalign(0.0)
            sidebar.pack_start(progress_label, False, False, 0)

            progress_adjustment = Gtk.Adjustment(
                value=0.0,
                lower=0.0,
                upper=100.0,
                step_increment=0.1,
                page_increment=1.0,
                page_size=0.0,
            )
            self.progress_slider = Gtk.Scale(
                orientation=Gtk.Orientation.HORIZONTAL, adjustment=progress_adjustment
            )
            self.progress_slider.set_digits(1)
            self.progress_slider.connect("value-changed", self._on_progress_slider_changed)
            sidebar.pack_start(self.progress_slider, False, False, 0)

            self.preview_status = Gtk.Label()
            self.preview_status.set_xalign(0.0)
            self.preview_status.set_width_chars(50)
            self.preview_status.set_max_width_chars(50)
            self.preview_status.set_line_wrap(False)
            self.preview_status.set_ellipsize(Pango.EllipsizeMode.END)
            sidebar.pack_start(self.preview_status, False, False, 0)

            sidebar.pack_start(Gtk.Separator(orientation=Gtk.Orientation.HORIZONTAL), False, False, 4)

            pantograph_label = Gtk.Label(label=_("Pantograph layout"))
            pantograph_label.set_xalign(0.0)
            sidebar.pack_start(pantograph_label, False, False, 0)

            pantograph_grid = Gtk.Grid(column_spacing=6, row_spacing=4)
            sidebar.pack_start(pantograph_grid, False, False, 0)

            repeat_spin = Gtk.SpinButton(
                adjustment=Gtk.Adjustment(
                    value=2,
                    lower=1,
                    upper=20,
                    step_increment=1,
                    page_increment=2,
                    page_size=0,
                ),
                numeric=True,
            )
            repeat_spin.set_value(self.repeat_count)
            repeat_spin.connect("value-changed", self._on_repeat_changed)
            pantograph_grid.attach(Gtk.Label(label=_("Repeats"), xalign=0), 0, 0, 1, 1)
            pantograph_grid.attach(repeat_spin, 1, 0, 1, 1)
            self.repeat_spin = repeat_spin

            rows_spin = Gtk.SpinButton(
                adjustment=Gtk.Adjustment(
                    value=2,
                    lower=1,
                    upper=20,
                    step_increment=1,
                    page_increment=2,
                    page_size=0,
                ),
                numeric=True,
            )
            rows_spin.set_value(self.row_count)
            rows_spin.connect("value-changed", self._on_rows_changed)
            pantograph_grid.attach(Gtk.Label(label=_("Rows"), xalign=0), 0, 1, 1, 1)
            pantograph_grid.attach(rows_spin, 1, 1, 1, 1)
            self.rows_spin = rows_spin

            distance_adjustment = Gtk.Adjustment(
                value=self.row_distance_mm,
                lower=1.0,
                upper=5000.0,
                step_increment=1.0,
                page_increment=10.0,
                page_size=0.0,
            )
            distance_spin = Gtk.SpinButton(adjustment=distance_adjustment, digits=1)
            distance_spin.connect("value-changed", self._on_row_distance_changed)
            pantograph_grid.attach(Gtk.Label(label=_("Row distance (mm)"), xalign=0), 0, 2, 1, 1)
            pantograph_grid.attach(distance_spin, 1, 2, 1, 1)
            self.row_distance_spin = distance_spin

            stagger_toggle = Gtk.CheckButton(label=_("Stagger alternate rows"))
            stagger_toggle.set_active(self.stagger)
            stagger_toggle.connect("toggled", self._on_stagger_toggled)
            pantograph_grid.attach(stagger_toggle, 0, 3, 2, 1)
            self.stagger_toggle = stagger_toggle

            stagger_adjustment = Gtk.Adjustment(
                value=self.stagger_percent,
                lower=0.0,
                upper=100.0,
                step_increment=1.0,
                page_increment=10.0,
                page_size=0.0,
            )
            stagger_scale = Gtk.Scale(orientation=Gtk.Orientation.HORIZONTAL, adjustment=stagger_adjustment)
            stagger_scale.set_digits(0)
            stagger_scale.connect("value-changed", self._on_stagger_percent_changed)
            pantograph_grid.attach(Gtk.Label(label=_("Stagger %"), xalign=0), 0, 4, 1, 1)
            pantograph_grid.attach(stagger_scale, 1, 4, 1, 1)
            self.stagger_scale = stagger_scale

            export_label = Gtk.Label(label=_("Export format"))
            export_label.set_xalign(0.0)
            sidebar.pack_start(export_label, False, False, 0)

            self.format_combo = Gtk.ComboBoxText()
            for key, profile in self.exporters.items():
                self.format_combo.append_text(f"{key} – {profile.title}")
            self.format_combo.set_active(0)
            sidebar.pack_start(self.format_combo, False, False, 0)

            export_button = Gtk.Button(label=_("Export…"))
            export_button.connect("clicked", self._export)
            sidebar.pack_start(export_button, False, False, 4)

            self.export_status = Gtk.Label()
            self.export_status.set_xalign(0.0)
            self.export_status.set_width_chars(50)
            self.export_status.set_max_width_chars(50)
            self.export_status.set_line_wrap(False)
            self.export_status.set_ellipsize(Pango.EllipsizeMode.END)
            sidebar.pack_start(self.export_status, False, False, 0)

            self.show_all()

        def _toggle_play(self, *_args) -> None:
            if (
                not self.playing
                and self.model.total_length_mm > 0
                and math.isclose(self.progress_mm, self.model.total_length_mm, abs_tol=1e-6)
            ):
                self.progress_mm = 0.0
                self.drawing_area.queue_draw()
            self.playing = not self.playing
            self.play_button.set_label(_("Pause") if self.playing else _("Play"))
            self.last_tick = time.monotonic()

        def _restart(self, *_args) -> None:
            self.progress_mm = 0.0
            self.playing = False
            self.play_button.set_label(_("Play"))
            self.last_tick = time.monotonic()
            self.drawing_area.queue_draw()
            self._set_progress_slider_value(0.0)
            self.preview_status.set_text(_("Preview reset. Press Play to start."))
            self._surface_dirty = True

        def _on_speed_changed(self, slider: Gtk.Scale) -> None:
            self.speed_multiplier = slider.get_value()
            self.speed_value_label.set_label(f"{self.speed_multiplier:.2f}×")
        def _on_repeat_changed(self, spin: Gtk.SpinButton) -> None:
            self.repeat_count = max(1, int(spin.get_value()))
            self._invalidate_surface()

        def _on_rows_changed(self, spin: Gtk.SpinButton) -> None:
            self.row_count = max(1, int(spin.get_value()))
            self._invalidate_surface()

        def _on_row_distance_changed(self, spin: Gtk.SpinButton) -> None:
            self.row_distance_mm = max(1.0, spin.get_value())
            self._invalidate_surface()

        def _on_stagger_toggled(self, button: Gtk.CheckButton) -> None:
            self.stagger = button.get_active()
            self._invalidate_surface()

        def _on_stagger_percent_changed(self, scale: Gtk.Scale) -> None:
            self.stagger_percent = scale.get_value()
            if self.stagger:
                self._invalidate_surface()

        def _export(self, *_args) -> None:
            active = self.format_combo.get_active()
            if active < 0:
                return
            format_key = list(self.exporters.keys())[active]
            profile = self.exporters[format_key]

            dialog = Gtk.FileChooserDialog(
                title=_("Export Motion Path"),
                transient_for=self,
                modal=True,
                action=Gtk.FileChooserAction.SAVE,
            )
            dialog.add_buttons(
                Gtk.STOCK_CANCEL,
                Gtk.ResponseType.CANCEL,
                Gtk.STOCK_SAVE,
                Gtk.ResponseType.OK,
            )
            try:
                dialog.set_current_folder(str(Path.home()))
            except Exception:
                pass
            dialog.set_current_name(f"quilt_path.{profile.extension}")
            response = dialog.run()
            filename: Optional[str] = None
            if response == Gtk.ResponseType.OK:
                filename = dialog.get_filename()
            dialog.destroy()

            if not filename:
                return

            out_path = Path(filename)
            if out_path.suffix.lower() != f".{profile.extension.lower()}":
                out_path = out_path.with_suffix(f".{profile.extension.lower()}")

            try:
                out_path.parent.mkdir(parents=True, exist_ok=True)
            except PermissionError:
                self.export_status.set_text(
                    _("No permission to create folder: ") + str(out_path.parent)
                )
                return
            except FileExistsError:
                pass

            try:
                profile.writer(self.model, out_path)
            except PermissionError:
                self.export_status.set_text(
                    _("Cannot write to {path}. Pick a folder inside your home directory.").format(
                        path=str(out_path)
                    )
                )
                return
            except Exception as exc:  # pragma: no cover - GUI feedback
                self.export_status.set_text(_("Export failed: ") + str(exc))
                return

            self.export_status.set_text(_("Exported to ") + str(out_path))
            self._surface_dirty = True
        def _invalidate_surface(self) -> None:
            self._surface_dirty = True
            self.drawing_area.queue_draw()

        def _on_destroy(self, *_args) -> None:
            if self._timeout_id is not None:
                GLib.source_remove(self._timeout_id)
                self._timeout_id = None
            if self._prime_source is not None:
                GLib.source_remove(self._prime_source)
                self._prime_source = None
            Gtk.main_quit()

        def _tick(self) -> bool:
            if not self.playing or not self.model.edges:
                return True
            now = time.monotonic()
            if self.last_tick is None:
                self.last_tick = now
                return True

            delta = now - self.last_tick
            self.last_tick = now

            advance = delta * self.BASE_SPEED_MM_PER_SEC * self.speed_multiplier
            self.progress_mm += advance
            if self.progress_mm >= self.model.total_length_mm:
                self.progress_mm = self.model.total_length_mm
                self.playing = False
                self.play_button.set_label(_("Play"))
            self.drawing_area.queue_draw()
            return True

        def _on_draw(self, widget: Gtk.DrawingArea, cr) -> None:
            width = widget.get_allocated_width()
            height = widget.get_allocated_height()
            self._ensure_static_surface(width, height)
            if self._static_surface is None:
                return

            cr.set_source_surface(self._static_surface, 0, 0)
            cr.paint()

            if not self.model.edges or self._viewport is None:
                self.preview_status.set_text(_("Select at least one path to preview."))
                return

            scale, offset_x, offset_y = self._viewport
            cr.translate(offset_x, offset_y)
            cr.scale(scale, scale)

            self._draw_progress(cr)

            point, needle_down = self.model.point_at(self.progress_mm)
            cr.save()
            cr.translate(point[0], point[1])
            cr.set_source_rgba(0.9, 0.2, 0.2, 1.0 if needle_down else 0.7)
            radius = 4.0 / scale
            cr.arc(0, 0, radius, 0, math.tau)
            cr.fill()
            cr.restore()

            stitched = min(self.progress_mm, self.model.total_length_mm)
            percent = (stitched / self.model.total_length_mm * 100.0) if self.model.total_length_mm else 0.0
            self._set_progress_slider_value(percent)
            self.preview_status.set_text(
                _("Path length: {length:.1f} mm   Previewed: {progress:.1f} mm ({pct:.1f}%)").format(
                    length=self.model.total_length_mm, progress=stitched, pct=percent
                )
            )

        def _ensure_static_surface(self, width: int, height: int) -> None:
            needs_new = (
                self._static_surface is None
                or self._surface_dirty
                or self._static_surface_size != (width, height)
            )
            if not needs_new:
                return

            surface = cairo.ImageSurface(cairo.FORMAT_ARGB32, width, height)
            ctx = cairo.Context(surface)
            ctx.set_source_rgb(0.98, 0.98, 0.98)
            ctx.paint()

            if self.model.edges:
                viewport = self._compute_viewport(width, height)
                self._viewport = viewport
                ctx.translate(viewport[1], viewport[2])
                ctx.scale(viewport[0], viewport[0])
                self._draw_full_pattern(ctx)
            else:
                self._viewport = (1.0, 0.0, 0.0)

            self._static_surface = surface
            self._static_surface_size = (width, height)
            self._surface_dirty = False

        def _set_progress_slider_value(self, percent: float) -> None:
            if not hasattr(self, "progress_slider"):
                return
            self._updating_progress_slider = True
            self.progress_slider.set_value(max(0.0, min(100.0, percent)))
            self._updating_progress_slider = False

        def _on_progress_slider_changed(self, slider: Gtk.Scale) -> None:
            if self._updating_progress_slider or not self.model.total_length_mm:
                return
            percent = slider.get_value()
            self.progress_mm = (
                percent / 100.0 * self.model.total_length_mm
            )
            self.last_tick = time.monotonic()
            self.drawing_area.queue_draw()

        def _pantograph_offsets(self) -> List[Tuple[float, float]]:
            width_px = max(self.model.bounds[2] - self.model.bounds[0], 1e-3)
            height_px = max(self.model.bounds[3] - self.model.bounds[1], 1e-3)
            self._pattern_width_px = width_px
            self._pattern_height_px = height_px
            row_spacing_px = max(self.row_distance_mm / self.model.px_to_mm, 1.0)
            stagger_px = width_px * (self.stagger_percent / 100.0) if self.stagger else 0.0

            if self.model.start_point is not None and self.model.end_point is not None:
                delta_x = self.model.end_point[0] - self.model.start_point[0]
                delta_y = self.model.end_point[1] - self.model.start_point[1]
            else:
                delta_x = width_px
                delta_y = 0.0

            offsets: List[Tuple[float, float]] = []
            for row in range(self.row_count):
                base_dx = stagger_px if (self.stagger and row % 2 == 1) else 0.0
                row_dy = row * row_spacing_px
                for repeat in range(self.repeat_count):
                    dx = base_dx + repeat * delta_x
                    dy = row_dy + repeat * delta_y
                    offsets.append((dx, dy))
            return offsets

        def _pantograph_bounds(self) -> Tuple[float, float, float, float]:
            min_x, min_y, max_x, max_y = self.model.bounds
            offsets = self._pantograph_offsets()
            total_min_x = float("inf")
            total_min_y = float("inf")
            total_max_x = float("-inf")
            total_max_y = float("-inf")
            for dx, dy in offsets:
                total_min_x = min(total_min_x, min_x + dx)
                total_min_y = min(total_min_y, min_y + dy)
                total_max_x = max(total_max_x, max_x + dx)
                total_max_y = max(total_max_y, max_y + dy)
            if not offsets:
                total_min_x, total_min_y, total_max_x, total_max_y = min_x, min_y, max_x, max_y
            return (total_min_x, total_min_y, total_max_x, total_max_y)

        def _on_size_allocate(self, widget, allocation) -> None:
            if self._static_surface_size != (allocation.width, allocation.height):
                self._surface_dirty = True

        def _prime_surface(self) -> bool:
            if not self.drawing_area.get_realized():
                return True
            width = self.drawing_area.get_allocated_width()
            height = self.drawing_area.get_allocated_height()
            if width <= 0 or height <= 0:
                return True
            self._ensure_static_surface(width, height)
            self.drawing_area.queue_draw()
            self._prime_source = None
            return False

        def _compute_viewport(self, width: int, height: int) -> Tuple[float, float, float]:
            min_x, min_y, max_x, max_y = self._pantograph_bounds()
            margin = 0.05
            span_x = max(max_x - min_x, 1e-3)
            span_y = max(max_y - min_y, 1e-3)
            scale_x = width * (1.0 - margin) / span_x
            scale_y = height * (1.0 - margin) / span_y
            scale = min(scale_x, scale_y)
            offset_x = (width - span_x * scale) / 2.0 - min_x * scale
            offset_y = (height - span_y * scale) / 2.0 - min_y * scale
            return scale, offset_x, offset_y

        def _draw_full_pattern(self, cr) -> None:
            cr.save()
            cr.set_line_width(1.0)
            offsets = self._pantograph_offsets()

            for dx, dy in offsets:
                cr.save()
                cr.translate(dx, dy)
                for seg in self.model.segments:
                    color = (0.2, 0.4, 0.7, 0.7) if seg.needle_down else (0.85, 0.2, 0.2, 0.8)
                    cr.set_source_rgba(*color)
                    cr.new_path()
                    pts = seg.points
                    cr.move_to(*pts[0])
                    for pt in pts[1:]:
                        cr.line_to(*pt)
                    cr.stroke()
                cr.restore()
            cr.restore()

        def _draw_progress(self, cr) -> None:
            cr.save()
            cr.set_line_width(2.0)
            remaining = self.progress_mm

            for edge in self.model.edges:
                if remaining <= edge.start_length_mm:
                    continue
                start = edge.start_px
                end = edge.end_px
                length_remaining = min(remaining - edge.start_length_mm, edge.length_mm)
                if length_remaining <= 0:
                    continue

                if length_remaining < edge.length_mm:
                    ratio = length_remaining / edge.length_mm if edge.length_mm else 0.0
                    end = (
                        edge.start_px[0] + (edge.end_px[0] - edge.start_px[0]) * ratio,
                        edge.start_px[1] + (edge.end_px[1] - edge.start_px[1]) * ratio,
                    )

                if edge.needle_down:
                    cr.set_source_rgba(0.1, 0.55, 0.85, 0.9)
                else:
                    cr.set_source_rgba(0.9, 0.25, 0.25, 0.9)
                cr.move_to(*start)
                cr.line_to(*end)
                cr.stroke()

                if length_remaining < edge.length_mm:
                    break

            cr.restore()
else:
    class QuiltPreviewWindow:  # pragma: no cover - placeholder when GTK missing
        pass




# Export writers -------------------------------------------------------------
def _write_dxf(model: MotionPathModel, outfile: Path) -> None:
    def section(lines: List[str]) -> List[str]:
        chunk: List[str] = []
        for i in range(0, len(lines), 2):
            chunk.extend(lines[i : i + 2])
        return chunk

    entities: List[str] = []
    for idx, (needle_down, pts) in enumerate(model.iter_segments_mm()):
        if len(pts) < 2:
            continue
        layer = "STITCH" if needle_down else "TRAVEL"
        entities.extend(
            [
                "0",
                "LWPOLYLINE",
                "8",
                layer,
                "90",
                str(len(pts)),
                "70",
                "0",
            ]
        )
        for x, y in pts:
            x_c, y_c = _cartesian_coords(model, x, y)
            entities.extend(["10", f"{x_c:.4f}", "20", f"{y_c:.4f}"])
    content = [
        "0",
        "SECTION",
        "2",
        "HEADER",
        "0",
        "ENDSEC",
        "0",
        "SECTION",
        "2",
        "ENTITIES",
        *entities,
        "0",
        "ENDSEC",
        "0",
        "EOF",
    ]
    outfile.write_text("\n".join(content), encoding="ascii")


def _write_gif(model: MotionPathModel, outfile: Path) -> None:
    frame_count = 60 if model.total_length_mm > 0 else 1
    width = 700
    height = 700
    min_x, min_y, max_x, max_y = model.bounds
    span_x = max(max_x - min_x, 1e-3)
    span_y = max(max_y - min_y, 1e-3)
    margin = 0.1
    scale = min(
        (width * (1.0 - margin)) / span_x,
        (height * (1.0 - margin)) / span_y,
    )
    offset_x = (width - span_x * scale) / 2 - min_x * scale
    offset_y = (height - span_y * scale) / 2 - min_y * scale

    def transform(point: Point) -> Tuple[int, int]:
        return (
            int(round(point[0] * scale + offset_x)),
            int(round(point[1] * scale + offset_y)),
        )

    base_paths: List[List[Tuple[int, int]]] = []
    stroke_colors: List[str] = []
    for seg in model.segments:
        pts = [transform(pt) for pt in seg.points]
        base_paths.append(pts)
        stroke_colors.append("#4a7bc6" if seg.needle_down else "#c4c4c4")

    frames: List[Image.Image] = []
    for idx in range(frame_count):
        progress = (
            (idx / (frame_count - 1)) * model.total_length_mm
            if frame_count > 1
            else model.total_length_mm
        )
        img = Image.new("RGB", (width, height), "#fefefe")
        draw = ImageDraw.Draw(img)

        for path, color in zip(base_paths, stroke_colors):
            if len(path) >= 2:
                draw.line(path, fill=color, width=2)

        remaining = progress
        for edge in model.edges:
            if remaining <= edge.start_length_mm:
                continue
            length_remaining = min(remaining - edge.start_length_mm, edge.length_mm)
            if length_remaining <= 0:
                continue
            start_pt = transform(edge.start_px)
            if length_remaining < edge.length_mm:
                ratio = length_remaining / edge.length_mm if edge.length_mm else 0.0
                interp = (
                    edge.start_px[0] + (edge.end_px[0] - edge.start_px[0]) * ratio,
                    edge.start_px[1] + (edge.end_px[1] - edge.start_px[1]) * ratio,
                )
                end_pt = transform(interp)
            else:
                end_pt = transform(edge.end_px)

            draw.line([start_pt, end_pt], fill="#00bcd4", width=3)

            if length_remaining < edge.length_mm:
                break

        if model.total_length_mm > 0:
            point, needle_down = model.point_at(progress)
            px = transform(point)
            r = 4
            color = "#e53935" if needle_down else "#f06292"
            draw.ellipse([px[0] - r, px[1] - r, px[0] + r, px[1] + r], fill=color)

        frames.append(img)

    Path(outfile).parent.mkdir(parents=True, exist_ok=True)
    frames[0].save(
        outfile,
        save_all=True,
        append_images=frames[1:],
        duration=60,
        loop=0,
        disposal=2,
    )
def _cartesian_coords(model: MotionPathModel, x_mm: float, y_mm: float) -> Tuple[float, float]:
    if model.doc_height_mm is not None:
        return x_mm, model.doc_height_mm - y_mm
    return x_mm, -y_mm


EXPORT_PROFILES: Dict[str, ExportProfile] = {
    "DXF": ExportProfile(
        title="AutoCAD DXF (polyline)",
        extension="dxf",
        description="Light‑weight polyline DXF",
        writer=_write_dxf,
    ),
    "GIF": ExportProfile(
        title="Animated GIF",
        extension="gif",
        description="Preview animation exported as GIF",
        writer=lambda model, dest: _write_gif(model, dest),
    ),
}


class QuiltMotionExportExtension(inkex.EffectExtension):
    """Entry point for Inkscape."""

    def effect(self) -> None:  # pragma: no cover - Inkscape runtime
        if not GTK_AVAILABLE:
            detail = (
                _(
                    "PyGObject (Gtk 3) is required for the preview UI. "
                    "Install the packages python3-gi, python3-gi-cairo, python3-cairo, and gir1.2-gtk-3.0."
                )
            )
            if GTK_LOAD_ERROR:
                detail = f"{detail}\n{GTK_LOAD_ERROR}"
            raise inkex.AbortExtension(detail)

        selection = self.svg.selection.filter(PathElement)
        if not selection:
            raise inkex.AbortExtension(_("Please select at least one path."))

        tolerance = 0.4
        ordered_segments: List[MotionSegment] = []
        for elem in selection.values():
            ordered_segments.extend(_flatten_path_element(elem, tolerance=tolerance))

        if not ordered_segments:
            raise inkex.AbortExtension(_("No drawable segments were found."))

        try:
            px_per_mm = self.svg.unittouu("1mm")
        except Exception:
            px_per_mm = None
        if px_per_mm and px_per_mm != 0:
            px_to_mm = 1.0 / px_per_mm
        else:
            px_to_mm = units.convert_unit("1px", "mm")

        doc_height_attr = self.svg.get("height")
        doc_height_px: Optional[float] = None
        try:
            if doc_height_attr:
                doc_height_px = self.svg.unittouu(doc_height_attr)
            elif getattr(self.svg, "viewbox_height", None):
                doc_height_px = float(self.svg.viewbox_height)
        except Exception:
            doc_height_px = None

        model = MotionPathModel(ordered_segments, px_to_mm=px_to_mm, doc_height_px=doc_height_px)
        window = QuiltPreviewWindow(model, EXPORT_PROFILES)
        window.present()
        Gtk.main()


if __name__ == "__main__":  # pragma: no cover
    QuiltMotionExportExtension().run()
