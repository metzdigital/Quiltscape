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
import heapq
import warnings
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import inkex
from inkex import bezier, units
from inkex.elements import PathElement
from inkex.localization import inkex_gettext as _
from inkex.paths import CubicSuperPath

_EXTENSION_DIR = Path(__file__).resolve().parent
_SIDECAR_LIBS = _EXTENSION_DIR / "quilt_motion_exporter_libs"
if _SIDECAR_LIBS.exists():
    sys.path.insert(0, str(_SIDECAR_LIBS))

try:
    from PIL import Image, ImageDraw

    PIL_AVAILABLE = True
    PIL_LOAD_ERROR: Optional[str] = None
except Exception as exc:  # pragma: no cover - optional dependency
    Image = None  # type: ignore
    ImageDraw = None  # type: ignore
    PIL_AVAILABLE = False
    PIL_LOAD_ERROR = str(exc)

warnings.filterwarnings(
    "ignore",
    message="DynamicImporter\\.exec_module\\(\\) not found; falling back to load_module\\(\\)",
    category=ImportWarning,
)

try:
    import tkinter as tk
    from tkinter import filedialog, ttk

    TK_AVAILABLE = True
    TK_LOAD_ERROR: Optional[str] = None
except Exception as exc:  # pragma: no cover - Tk is standard but may be missing
    TK_AVAILABLE = False
    TK_LOAD_ERROR = str(exc)


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

        self._refine_edges()

    def _refine_edges(self) -> None:
        if not self.edges:
            return
        span_x = self.bounds[2] - self.bounds[0]
        span_y = self.bounds[3] - self.bounds[1]
        span_px = max(span_x, span_y)
        if span_px <= 0:
            return
        target_px = max(span_px / 400.0, 0.5)
        target_mm = target_px * self.px_to_mm
        if target_mm <= 0:
            return

        new_edges: List[MotionEdge] = []
        total = 0.0
        for edge in self.edges:
            segments = max(1, int(math.ceil(edge.length_mm / target_mm)))
            for i in range(segments):
                t0 = i / segments
                t1 = (i + 1) / segments
                start = (
                    edge.start_px[0] + (edge.end_px[0] - edge.start_px[0]) * t0,
                    edge.start_px[1] + (edge.end_px[1] - edge.start_px[1]) * t0,
                )
                end = (
                    edge.start_px[0] + (edge.end_px[0] - edge.start_px[0]) * t1,
                    edge.start_px[1] + (edge.end_px[1] - edge.start_px[1]) * t1,
                )
                seg_len = edge.length_mm * (t1 - t0)
                new_edges.append(
                    MotionEdge(start_px=start, end_px=end, needle_down=edge.needle_down, start_length_mm=total, length_mm=seg_len)
                )
                total += seg_len

        self.edges = new_edges
        self.total_length_mm = total

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


def _compute_pantograph_offsets(
    bounds: Tuple[float, float, float, float],
    repeat_count: int,
    row_count: int,
    row_distance_mm: float,
    px_to_mm: float,
    stagger: bool,
    stagger_percent: float,
    start_point: Optional[Point],
    end_point: Optional[Point],
) -> List[Tuple[int, float, float]]:
    """Return offsets (row_idx, dx, dy) for pantograph repeats."""
    min_x, min_y, max_x, max_y = bounds
    width_px = max(max_x - min_x, 1e-3)
    height_px = max(max_y - min_y, 1e-3)
    row_spacing_px = height_px + max(row_distance_mm / px_to_mm, 1.0)
    stagger_px = width_px * (stagger_percent / 100.0) if stagger else 0.0

    if start_point is not None and end_point is not None:
        delta_x = end_point[0] - start_point[0]
        delta_y = end_point[1] - start_point[1]
    else:
        delta_x = width_px
        delta_y = 0.0

    offsets: List[Tuple[int, float, float]] = []
    target_min_x = min_x
    target_max_x = max_x + (repeat_count - 1) * delta_x
    for row in range(row_count):
        base_dx = stagger_px if (stagger and row % 2 == 1) else 0.0
        row_dy = row * row_spacing_px
        row_dx: List[float] = [base_dx + repeat * delta_x for repeat in range(repeat_count)]
        row_dx.sort()
        if row_dx:
            while min_x + row_dx[0] > target_min_x + 1e-6:
                row_dx.insert(0, row_dx[0] - delta_x)
            while max_x + row_dx[-1] < target_max_x - 1e-6:
                row_dx.append(row_dx[-1] + delta_x)
        for dx in row_dx:
            dy = row_dy + (dx - base_dx) / delta_x * delta_y if delta_x else row_dy
            offsets.append((row, dx, dy))
    return offsets


def _compute_layout_bounds(
    bounds: Tuple[float, float, float, float],
    repeat_count: int,
    row_count: int,
    row_distance_mm: float,
    px_to_mm: float,
    start_point: Optional[Point],
    end_point: Optional[Point],
) -> Tuple[float, float, float, float]:
    """Rectangular layout bounds ignoring stagger offsets."""
    min_x, min_y, max_x, max_y = bounds
    width = max_x - min_x
    height = max_y - min_y

    offsets: List[Tuple[float, float]] = []
    row_spacing_px = height + max(row_distance_mm / px_to_mm, 1.0)

    if start_point is not None and end_point is not None:
        delta_x = end_point[0] - start_point[0]
        delta_y = end_point[1] - start_point[1]
    else:
        delta_x = width
        delta_y = 0.0

    for row in range(row_count):
        row_dy = row * row_spacing_px
        for repeat in range(repeat_count):
            dx = repeat * delta_x
            dy = row_dy + repeat * delta_y
            offsets.append((dx, dy))

    if not offsets:
        return (min_x, min_y, max_x, max_y)

    total_min_x = min(min_x + dx for dx, _dy in offsets)
    total_min_y = min(min_y + dy for _dx, dy in offsets)
    total_max_x = max(min_x + dx + width for dx, _dy in offsets)
    total_max_y = max(min_y + dy + height for _dx, dy in offsets)
    return (total_min_x, total_min_y, total_max_x, total_max_y)


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


def optimize_motion_segments(
    segments: List[MotionSegment],
    start_point: Optional[Point] = None,
    end_point: Optional[Point] = None,
    tolerance: float = 1e-6,
) -> List[MotionSegment]:
    """Reorder a continuous stitched path to reduce geometric overlaps.

    This preserves:

    - The stitched design (the set of needle‑down edges), and
    - The logical start and end locations of the motion path.

    The optimisation only touches needle-down segments and falls back to
    the original segments if:

    - The stitched graph is disconnected, or
    - The optimised path would change the stitched geometry, or
    - The overlap metric is not improved.
    """

    def quantize_point(point: Point) -> Tuple[float, float]:
        return (round(point[0], 6), round(point[1], 6))

    def stitched_edge_counts(segment_list: List[MotionSegment]) -> Dict[Tuple[Tuple[float, float], Tuple[float, float]], int]:
        counts: Dict[Tuple[Tuple[float, float], Tuple[float, float]], int] = {}
        for segment in segment_list:
            if not segment.needle_down or len(segment.points) < 2:
                continue
            for index in range(len(segment.points) - 1):
                start = segment.points[index]
                end = segment.points[index + 1]
                if math.isclose(start[0], end[0], abs_tol=1e-9) and math.isclose(
                    start[1], end[1], abs_tol=1e-9
                ):
                    continue
                start_key = quantize_point(start)
                end_key = quantize_point(end)
                if start_key == end_key:
                    continue
                canonical = (
                    start_key if start_key <= end_key else end_key,
                    end_key if start_key <= end_key else start_key,
                )
                counts[canonical] = counts.get(canonical, 0) + 1
        return counts

    def overlap_length(
        segment_list: List[MotionSegment],
        baseline_counts: Dict[Tuple[Tuple[float, float], Tuple[float, float]], int],
    ) -> float:
        """Return total extra length beyond baseline multiplicity."""
        seen: Dict[Tuple[Tuple[float, float], Tuple[float, float]], int] = {}
        overlap_total = 0.0
        for segment in segment_list:
            if not segment.needle_down or len(segment.points) < 2:
                continue
            for index in range(len(segment.points) - 1):
                start = segment.points[index]
                end = segment.points[index + 1]
                if math.isclose(start[0], end[0], abs_tol=1e-9) and math.isclose(
                    start[1], end[1], abs_tol=1e-9
                ):
                    continue
                start_key = quantize_point(start)
                end_key = quantize_point(end)
                if start_key == end_key:
                    continue
                canonical = (
                    start_key if start_key <= end_key else end_key,
                    end_key if start_key <= end_key else start_key,
                )
                length = math.dist(start, end)
                current_count = seen.get(canonical, 0) + 1
                baseline = baseline_counts.get(canonical, 0)
                if current_count > baseline:
                    overlap_total += length
                seen[canonical] = current_count
        return overlap_total

    def segment_intersections(a: Point, b: Point, c: Point, d: Point) -> List[Point]:
        """Return intersection points between two closed segments."""
        def _on_segment(p: Point, q: Point, r: Point) -> bool:
            return min(p[0], r[0]) - 1e-9 <= q[0] <= max(p[0], r[0]) + 1e-9 and min(p[1], r[1]) - 1e-9 <= q[1] <= max(p[1], r[1]) + 1e-9

        def _orient(p: Point, q: Point, r: Point) -> float:
            return (q[0] - p[0]) * (r[1] - p[1]) - (q[1] - p[1]) * (r[0] - p[0])

        o1 = _orient(a, b, c)
        o2 = _orient(a, b, d)
        o3 = _orient(c, d, a)
        o4 = _orient(c, d, b)

        points: List[Point] = []

        # General intersection
        if (o1 == 0 and _on_segment(a, c, b)) or (o2 == 0 and _on_segment(a, d, b)) or (o3 == 0 and _on_segment(c, a, d)) or (o4 == 0 and _on_segment(c, b, d)):
            if o1 == 0 and _on_segment(a, c, b):
                points.append(c)
            if o2 == 0 and _on_segment(a, d, b):
                points.append(d)
            if o3 == 0 and _on_segment(c, a, d):
                points.append(a)
            if o4 == 0 and _on_segment(c, b, d):
                points.append(b)

        denom = (a[0] - b[0]) * (c[1] - d[1]) - (a[1] - b[1]) * (c[0] - d[0])
        if abs(denom) > 1e-12:
            t = ((a[0] - c[0]) * (c[1] - d[1]) - (a[1] - c[1]) * (c[0] - d[0])) / denom
            u = ((a[0] - c[0]) * (a[1] - b[1]) - (a[1] - c[1]) * (a[0] - b[0])) / denom
            if -1e-9 <= t <= 1 + 1e-9 and -1e-9 <= u <= 1 + 1e-9:
                px = a[0] + t * (b[0] - a[0])
                py = a[1] + t * (b[1] - a[1])
                points.append((px, py))

        if not points:
            return []
        # Deduplicate near-coincident points
        unique: List[Point] = []
        for p in points:
            if not any(math.isclose(p[0], q[0], abs_tol=1e-9) and math.isclose(p[1], q[1], abs_tol=1e-9) for q in unique):
                unique.append(p)
        return unique

    stitched_segments = [
        segment
        for segment in segments
        if segment.needle_down and len(segment.points) >= 2
    ]
    if not stitched_segments:
        return segments

    # Flatten to raw edges and split at every intersection so the optimiser
    # can choose alternate routes through intersection nodes.
    raw_edges: List[Tuple[Point, Point]] = []
    for segment in stitched_segments:
        for idx in range(len(segment.points) - 1):
            a = segment.points[idx]
            b = segment.points[idx + 1]
            if math.isclose(a[0], b[0], abs_tol=1e-9) and math.isclose(a[1], b[1], abs_tol=1e-9):
                continue
            raw_edges.append((a, b))

    if not raw_edges:
        return segments

    split_points: List[List[Point]] = [[edge[0], edge[1]] for edge in raw_edges]
    for i in range(len(raw_edges)):
        for j in range(i + 1, len(raw_edges)):
            p1, p2 = raw_edges[i]
            p3, p4 = raw_edges[j]
            pts = segment_intersections(p1, p2, p3, p4)
            if not pts:
                continue
            split_points[i].extend(pts)
            split_points[j].extend(pts)

    split_edges: List[Tuple[Point, Point]] = []
    for idx, edge in enumerate(raw_edges):
        a, b = edge
        pts = split_points[idx]
        # Sort points along the edge using projection parameter.
        dx = b[0] - a[0]
        dy = b[1] - a[1]
        def param(pt: Point) -> float:
            length_sq = dx * dx + dy * dy
            if length_sq <= 0:
                return 0.0
            return ((pt[0] - a[0]) * dx + (pt[1] - a[1]) * dy) / length_sq
        pts_sorted = sorted(pts, key=param)
        # Create sub-edges.
        for i in range(1, len(pts_sorted)):
            p_start = pts_sorted[i - 1]
            p_end = pts_sorted[i]
            if math.isclose(p_start[0], p_end[0], abs_tol=1e-9) and math.isclose(
                p_start[1], p_end[1], abs_tol=1e-9
            ):
                continue
            split_edges.append((p_start, p_end))

    if not split_edges:
        return segments

    vertex_index_by_point: Dict[Tuple[float, float], int] = {}
    vertices: List[Point] = []

    def vertex_for(point: Point) -> int:
        key = quantize_point(point)
        existing_index = vertex_index_by_point.get(key)
        if existing_index is not None:
            return existing_index
        new_index = len(vertices)
        vertex_index_by_point[key] = new_index
        vertices.append(point)
        return new_index

    base_edges: List[Tuple[int, int, float]] = []
    used_vertices: set = set()
    edge_map: Dict[Tuple[int, int], float] = {}
    for a, b in split_edges:
        va = vertex_for(a)
        vb = vertex_for(b)
        if va == vb:
            continue
        length = math.dist(vertices[va], vertices[vb])
        key = (va, vb) if va <= vb else (vb, va)
        # Keep only the shortest instance of a geometric edge; duplicated
        # overlaps in the original drawing are treated as optional.
        prev = edge_map.get(key)
        if prev is None or length < prev - 1e-9:
            edge_map[key] = length
        used_vertices.add(va)
        used_vertices.add(vb)

    for (va, vb), length in edge_map.items():
        base_edges.append((va, vb, length))

    if not base_edges:
        return segments

    # Baseline counts: each unique stitched sub-edge must appear at least once
    # (duplicate overlaps in the original drawing are treated as optional).
    baseline_edge_counts: Dict[Tuple[Tuple[float, float], Tuple[float, float]], int] = {}
    for vertex_index_a, vertex_index_b, _length in base_edges:
        pa = quantize_point(vertices[vertex_index_a])
        pb = quantize_point(vertices[vertex_index_b])
        key = (pa, pb) if pa <= pb else (pb, pa)
        if key not in baseline_edge_counts:
            baseline_edge_counts[key] = 1

    vertex_count = len(vertices)
    adjacency: List[List[Tuple[int, float, int]]] = [[] for _ in range(vertex_count)]
    for edge_index, (vertex_index_a, vertex_index_b, length) in enumerate(base_edges):
        adjacency[vertex_index_a].append((vertex_index_b, length, edge_index))
        adjacency[vertex_index_b].append((vertex_index_a, length, edge_index))

    # Ensure the stitched graph is a single connected component.
    starting_vertex = next(iter(used_vertices))
    visited_vertices: set = set()
    stack: List[int] = [starting_vertex]
    while stack:
        current_vertex = stack.pop()
        if current_vertex in visited_vertices:
            continue
        visited_vertices.add(current_vertex)
        for neighbour_vertex, _length, _edge_index in adjacency[current_vertex]:
            if neighbour_vertex not in visited_vertices:
                stack.append(neighbour_vertex)
    if visited_vertices != used_vertices:
        return segments

    degrees: List[int] = [0] * vertex_count
    for vertex_index_a, vertex_index_b, _length in base_edges:
        degrees[vertex_index_a] += 1
        degrees[vertex_index_b] += 1
    odd_vertices: List[int] = [
        index for index, degree in enumerate(degrees) if degree % 2 == 1
    ]

    def shortest_paths(source_vertex: int) -> Tuple[List[float], List[Optional[int]]]:
        distances: List[float] = [float("inf")] * vertex_count
        previous_vertices: List[Optional[int]] = [None] * vertex_count
        distances[source_vertex] = 0.0
        heap: List[Tuple[float, int]] = [(0.0, source_vertex)]
        while heap:
            distance, vertex_index_current = heapq.heappop(heap)
            if distance > distances[vertex_index_current] + 1e-12:
                continue
            for neighbour_vertex, edge_length, _edge_index in adjacency[vertex_index_current]:
                new_distance = distance + edge_length
                if new_distance + 1e-12 < distances[neighbour_vertex]:
                    distances[neighbour_vertex] = new_distance
                    previous_vertices[neighbour_vertex] = vertex_index_current
                    heapq.heappush(heap, (new_distance, neighbour_vertex))
        return distances, previous_vertices

    # Track how many times each base edge is duplicated when fixing parity.
    duplicate_paths: List[List[int]] = []

    # Determine desired end-point parities so that the final walk keeps the
    # same logical start and end locations as the original motion path.
    if start_point is not None:
        start_key = quantize_point(start_point)
        start_vertex_index = vertex_index_by_point.get(start_key, starting_vertex)
    else:
        start_vertex_index = starting_vertex

    if end_point is not None:
        end_key = quantize_point(end_point)
        end_vertex_index = vertex_index_by_point.get(end_key, start_vertex_index)
    else:
        end_vertex_index = start_vertex_index

    target_odd: set = set()
    if start_vertex_index != end_vertex_index:
        target_odd = {start_vertex_index, end_vertex_index}

    required_parity_vertices = set(odd_vertices) ^ target_odd

    required_count = len(required_parity_vertices)
    if required_count % 2 != 0:
        # Should not happen in a valid undirected graph, but guard anyway.
        return segments

    def shortest_path_with_trace(source_vertex_index: int, target_vertex_index: int) -> Tuple[float, List[int]]:
        distances: List[float] = [float("inf")] * vertex_count
        previous_vertices: List[Optional[int]] = [None] * vertex_count
        previous_edge: List[Optional[int]] = [None] * vertex_count
        distances[source_vertex_index] = 0.0
        heap: List[Tuple[float, int]] = [(0.0, source_vertex_index)]

        while heap:
            distance, vertex_index_current = heapq.heappop(heap)
            if distance > distances[vertex_index_current] + 1e-12:
                continue
            if vertex_index_current == target_vertex_index:
                break
            for neighbour_vertex, edge_length, edge_index in adjacency[vertex_index_current]:
                new_distance = distance + edge_length
                if new_distance + 1e-12 < distances[neighbour_vertex]:
                    distances[neighbour_vertex] = new_distance
                    previous_vertices[neighbour_vertex] = vertex_index_current
                    previous_edge[neighbour_vertex] = edge_index
                    heapq.heappush(heap, (new_distance, neighbour_vertex))

        if not math.isfinite(distances[target_vertex_index]):
            return float("inf"), []

        path_edges: List[int] = []
        current_vertex = target_vertex_index
        while current_vertex != source_vertex_index:
            prev = previous_vertices[current_vertex]
            edge_index = previous_edge[current_vertex]
            if prev is None or edge_index is None:
                return float("inf"), []
            path_edges.append(edge_index)
            current_vertex = prev
        path_edges.reverse()
        return distances[target_vertex_index], path_edges

    if 0 < required_count <= 16:
        # Exact minimum-weight perfect matching via DP (Held-Karp style).
        required_vertices_sorted = sorted(required_parity_vertices)
        m = len(required_vertices_sorted)
        # Precompute pairwise shortest paths.
        pair_costs: Dict[Tuple[int, int], Tuple[float, List[int]]] = {}
        for i in range(m):
            for j in range(i + 1, m):
                u = required_vertices_sorted[i]
                v = required_vertices_sorted[j]
                dist, path_edges = shortest_path_with_trace(u, v)
                pair_costs[(i, j)] = (dist, path_edges)

        full_mask = (1 << m) - 1
        dp: List[float] = [float("inf")] * (1 << m)
        choice: List[Optional[Tuple[int, int]]] = [None] * (1 << m)
        dp[0] = 0.0

        for mask in range(1 << m):
            if dp[mask] == float("inf"):
                continue
            # Find first unmatched vertex.
            try:
                first = next(idx for idx in range(m) if not (mask & (1 << idx)))
            except StopIteration:
                continue
            for second in range(first + 1, m):
                if mask & (1 << second):
                    continue
                pair = (first, second)
                cost, _path_edges = pair_costs.get(pair, (float("inf"), []))
                if not math.isfinite(cost) or not _path_edges:
                    continue
                next_mask = mask | (1 << first) | (1 << second)
                new_cost = dp[mask] + cost
                if new_cost < dp[next_mask] - 1e-12:
                    dp[next_mask] = new_cost
                    choice[next_mask] = pair

        if dp[full_mask] == float("inf"):
            return segments

        # Reconstruct chosen pairs.
        mask = full_mask
        chosen_pairs: List[Tuple[int, int]] = []
        while mask:
            pair = choice[mask]
            if pair is None:
                break
            chosen_pairs.append(pair)
            mask &= ~(1 << pair[0])
            mask &= ~(1 << pair[1])

        duplicate_paths: List[List[int]] = []
        for first, second in chosen_pairs:
            u = required_vertices_sorted[first]
            v = required_vertices_sorted[second]
            cost, path_edges = pair_costs[(first, second)]
            if not path_edges or not math.isfinite(cost):
                return segments
            duplicate_paths.append(path_edges)

    elif 16 < required_count <= 24:
        # Greedy pairing for larger sets to keep runtime bounded.
        remaining_vertices = set(required_parity_vertices)
        while remaining_vertices:
            candidates = sorted(remaining_vertices)
            best_pair: Optional[Tuple[int, int]] = None
            best_cost = float("inf")
            best_path_edges: List[int] = []

            for i in range(len(candidates)):
                u = candidates[i]
                for j in range(i + 1, len(candidates)):
                    v = candidates[j]
                    cost, path_edges = shortest_path_with_trace(u, v)
                    if not path_edges or not math.isfinite(cost):
                        continue
                    if cost < best_cost:
                        best_cost = cost
                        best_pair = (u, v)
                        best_path_edges = path_edges

            if best_pair is None or not best_path_edges:
                return segments

            duplicate_paths.append(best_path_edges)
            remaining_vertices.remove(best_pair[0])
            remaining_vertices.remove(best_pair[1])
    elif required_count > 24:
        return segments

    extended_edges: List[Tuple[int, int, float]] = list(base_edges)
    for path_edges in duplicate_paths:
        for edge_index in path_edges:
            base_vertex_a, base_vertex_b, base_length = base_edges[edge_index]
            extended_edges.append((base_vertex_a, base_vertex_b, base_length))

    adjacency_multigraph: List[List[Tuple[int, int]]] = [
        [] for _ in range(vertex_count)
    ]
    for edge_index, (vertex_index_a, vertex_index_b, _length) in enumerate(
        extended_edges
    ):
        adjacency_multigraph[vertex_index_a].append((vertex_index_b, edge_index))
        adjacency_multigraph[vertex_index_b].append((vertex_index_a, edge_index))

    used_edges: List[bool] = [False] * len(extended_edges)

    stack_vertices: List[int] = [start_vertex_index]
    euler_vertices: List[int] = []

    while stack_vertices:
        vertex_index_current = stack_vertices[-1]
        adjacency_list = adjacency_multigraph[vertex_index_current]
        while adjacency_list and used_edges[adjacency_list[-1][1]]:
            adjacency_list.pop()
        if not adjacency_list:
            euler_vertices.append(stack_vertices.pop())
        else:
            neighbour_vertex, edge_index = adjacency_list.pop()
            if used_edges[edge_index]:
                continue
            used_edges[edge_index] = True
            stack_vertices.append(neighbour_vertex)

    euler_vertices.reverse()
    if not euler_vertices:
        return segments

    optimized_points: List[Point] = []
    for vertex_index_current in euler_vertices:
        point = vertices[vertex_index_current]
        if not optimized_points:
            optimized_points.append(point)
            continue
        last_point = optimized_points[-1]
        if math.isclose(point[0], last_point[0], abs_tol=tolerance) and math.isclose(
            point[1], last_point[1], abs_tol=tolerance
        ):
            continue
        optimized_points.append(point)

    if len(optimized_points) < 2:
        return segments

    optimized_segment = MotionSegment(points=optimized_points, needle_down=True)
    optimized_segments: List[MotionSegment] = []
    inserted_stitched_segment = False
    for segment in segments:
        if segment.needle_down and len(segment.points) >= 2:
            if not inserted_stitched_segment:
                optimized_segments.append(optimized_segment)
                inserted_stitched_segment = True
            continue
        optimized_segments.append(segment)

    optimized_edge_counts = stitched_edge_counts(optimized_segments)
    for edge_key, baseline_count in baseline_edge_counts.items():
        if optimized_edge_counts.get(edge_key, 0) < baseline_count:
            # Missing required geometry from the original design.
            return segments

    original_overlap = overlap_length(segments, baseline_edge_counts)
    optimized_overlap = overlap_length(optimized_segments, baseline_edge_counts)

    if optimized_overlap + tolerance >= original_overlap:
        return segments

    return optimized_segments


if TK_AVAILABLE:
    class QuiltPreviewWindow:
        """Interactive Tk window that previews and exports the motion path."""

        BASE_SPEED_MM_PER_SEC = 35.0

        def __init__(
            self,
            model: MotionPathModel,
            exporters: Dict[str, ExportProfile],
        ) -> None:
            self.model = model
            self.exporters = exporters

            self.progress_mm = 0.0
            self.speed_multiplier = 1.0
            self.playing = True
            self.last_tick: Optional[float] = None
            self._tick_after_id: Optional[str] = None
            self._viewport: Optional[Tuple[float, float, float]] = None
            self._updating_progress_slider = False
            self._redraw_pending = False

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

            self.root = tk.Tk()
            self.root.title(_("Quilt Motion Preview"))
            self.root.geometry("1100x600")
            self.root.protocol("WM_DELETE_WINDOW", self._on_destroy)

            self._build_ui()
            self._refresh_y_warning()
            self._schedule_tick()
            self._schedule_redraw()

        def present(self) -> None:
            self.root.mainloop()

        def _build_ui(self) -> None:
            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(0, weight=1)

            main = ttk.Frame(self.root, padding=12)
            main.grid(row=0, column=0, sticky="nsew")
            main.columnconfigure(0, weight=1)
            main.rowconfigure(0, weight=1)

            left_column = ttk.Frame(main)
            left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
            left_column.columnconfigure(0, weight=1)
            left_column.rowconfigure(0, weight=1)

            self.canvas = tk.Canvas(left_column, background="#f8f8f8", highlightthickness=1)
            self.canvas.grid(row=0, column=0, sticky="nsew")
            self.canvas.bind("<Configure>", self._on_canvas_resize)

            self.y_warning_label = ttk.Label(left_column, foreground="#b8860b")
            self.y_warning_label.grid(row=1, column=0, sticky="w", pady=(6, 0))

            sidebar = ttk.Frame(main, width=280)
            sidebar.grid(row=0, column=1, sticky="ns")
            sidebar.columnconfigure(0, weight=1)

            controls = ttk.Frame(sidebar)
            controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
            controls.columnconfigure((0, 1, 2), weight=1)

            self.play_button = ttk.Button(controls, text=_("Pause"), command=self._toggle_play)
            self.play_button.grid(row=0, column=0, sticky="ew")

            restart_button = ttk.Button(controls, text=_("Restart"), command=self._restart)
            restart_button.grid(row=0, column=1, sticky="ew", padx=4)

            self.optimize_button = ttk.Button(controls, text=_("Optimize path"), command=self._optimize_path)
            self.optimize_button.grid(row=0, column=2, sticky="ew")

            ttk.Label(sidebar, text=_("Preview speed")).grid(row=1, column=0, sticky="w")
            self.speed_var = tk.DoubleVar(value=1.0)
            self.speed_slider = ttk.Scale(sidebar, from_=0.1, to=5.0, orient="horizontal", variable=self.speed_var, command=self._on_speed_changed)
            self.speed_slider.grid(row=2, column=0, sticky="ew")
            self.speed_value_label = ttk.Label(sidebar, text=_("1.00×"))
            self.speed_value_label.grid(row=3, column=0, sticky="w", pady=(0, 6))

            ttk.Label(sidebar, text=_("Preview progress")).grid(row=4, column=0, sticky="w")
            self.progress_var = tk.DoubleVar(value=0.0)
            self.progress_slider = ttk.Scale(sidebar, from_=0.0, to=100.0, orient="horizontal", variable=self.progress_var, command=self._on_progress_slider_changed)
            self.progress_slider.grid(row=5, column=0, sticky="ew")

            self.preview_status = ttk.Label(sidebar, text="", width=50)
            self.preview_status.grid(row=6, column=0, sticky="w", pady=(0, 8))

            ttk.Separator(sidebar, orient="horizontal").grid(row=7, column=0, sticky="ew", pady=6)

            ttk.Label(sidebar, text=_("Pantograph layout")).grid(row=8, column=0, sticky="w")
            pantograph = ttk.Frame(sidebar)
            pantograph.grid(row=9, column=0, sticky="ew")
            pantograph.columnconfigure(1, weight=1)

            self.repeat_var = tk.IntVar(value=self.repeat_count)
            ttk.Label(pantograph, text=_("Repeats")).grid(row=0, column=0, sticky="w")
            repeat_spin = ttk.Spinbox(pantograph, from_=1, to=20, textvariable=self.repeat_var, width=5, command=self._on_repeat_changed)
            repeat_spin.grid(row=0, column=1, sticky="w")

            self.rows_var = tk.IntVar(value=self.row_count)
            ttk.Label(pantograph, text=_("Rows")).grid(row=1, column=0, sticky="w")
            rows_spin = ttk.Spinbox(pantograph, from_=1, to=20, textvariable=self.rows_var, width=5, command=self._on_rows_changed)
            rows_spin.grid(row=1, column=1, sticky="w")

            self.row_distance_var = tk.DoubleVar(value=self.row_distance_mm)
            ttk.Label(pantograph, text=_("Row distance (mm)")).grid(row=2, column=0, sticky="w")
            row_distance_spin = ttk.Spinbox(pantograph, from_=1.0, to=5000.0, increment=1.0, textvariable=self.row_distance_var, width=7, command=self._on_row_distance_changed)
            row_distance_spin.grid(row=2, column=1, sticky="w")
            self.row_distance_spin = row_distance_spin

            self.stagger_var = tk.BooleanVar(value=self.stagger)
            stagger_toggle = ttk.Checkbutton(pantograph, text=_("Stagger alternate rows"), variable=self.stagger_var, command=self._on_stagger_toggled)
            stagger_toggle.grid(row=3, column=0, columnspan=2, sticky="w")

            self.stagger_percent_var = tk.DoubleVar(value=self.stagger_percent)
            ttk.Label(pantograph, text=_("Stagger %")).grid(row=4, column=0, sticky="w")
            self.stagger_scale = ttk.Scale(pantograph, from_=0.0, to=100.0, orient="horizontal", variable=self.stagger_percent_var, command=self._on_stagger_percent_changed)
            self.stagger_scale.grid(row=4, column=1, sticky="ew")

            self.mirror_rows_var = tk.BooleanVar(value=self.mirror_alternate_rows)
            mirror_toggle = ttk.Checkbutton(pantograph, text=_("Mirror every other row horizontally"), variable=self.mirror_rows_var, command=self._on_mirror_rows_toggled)
            mirror_toggle.grid(row=5, column=0, columnspan=2, sticky="w")

            self.mirror_rows_v_var = tk.BooleanVar(value=self.mirror_alternate_rows_vertical)
            mirror_v_toggle = ttk.Checkbutton(pantograph, text=_("Mirror every other row vertically"), variable=self.mirror_rows_v_var, command=self._on_mirror_rows_v_toggled)
            mirror_v_toggle.grid(row=6, column=0, columnspan=2, sticky="w")

            self.flip_h_var = tk.BooleanVar(value=self.flip_horizontal)
            flip_h_toggle = ttk.Checkbutton(pantograph, text=_("Flip horizontally"), variable=self.flip_h_var, command=self._on_flip_h_toggled)
            flip_h_toggle.grid(row=7, column=0, columnspan=2, sticky="w")

            self.flip_v_var = tk.BooleanVar(value=self.flip_vertical)
            flip_v_toggle = ttk.Checkbutton(pantograph, text=_("Flip vertically"), variable=self.flip_v_var, command=self._on_flip_v_toggled)
            flip_v_toggle.grid(row=8, column=0, columnspan=2, sticky="w")

            ttk.Label(sidebar, text=_("Export format")).grid(row=10, column=0, sticky="w", pady=(8, 0))
            export_row = ttk.Frame(sidebar)
            export_row.grid(row=11, column=0, sticky="ew")
            export_row.columnconfigure(0, weight=1)

            self.format_combo = ttk.Combobox(export_row, state="readonly")
            self.format_combo["values"] = [f"{key} – {profile.title}" for key, profile in self.exporters.items()]
            if self.format_combo["values"]:
                self.format_combo.current(0)
            self.format_combo.grid(row=0, column=0, sticky="ew")

            self.export_layout_var = tk.BooleanVar(value=self.export_entire_layout)
            export_layout_toggle = ttk.Checkbutton(export_row, text=_("Export entire layout"), variable=self.export_layout_var, command=self._on_export_layout_toggled)
            export_layout_toggle.grid(row=0, column=1, sticky="w", padx=(6, 0))

            export_button = ttk.Button(sidebar, text=_("Export…"), command=self._export)
            export_button.grid(row=12, column=0, sticky="w", pady=(8, 0))

            self.export_status = ttk.Label(sidebar, text="", width=50)
            self.export_status.grid(row=13, column=0, sticky="w", pady=(2, 0))

        def _toggle_play(self, *_args) -> None:
            if (
                not self.playing
                and self.model.total_length_mm > 0
                and math.isclose(self.progress_mm, self.model.total_length_mm, abs_tol=1e-6)
            ):
                self.progress_mm = 0.0
                self._schedule_redraw()
            self.playing = not self.playing
            self.play_button.config(text=_("Pause") if self.playing else _("Play"))
            self.last_tick = time.monotonic()

        def _restart(self, *_args) -> None:
            self.progress_mm = 0.0
            self.playing = False
            self.play_button.config(text=_("Play"))
            self.last_tick = time.monotonic()
            self._set_progress_slider_value(0.0)
            self.preview_status.config(text=_("Preview reset. Press Play to start."))
            if hasattr(self, "_edge_progress"):
                self._edge_progress = {}
            self._schedule_redraw()

        def _optimize_path(self, *_args) -> None:
            if not self.model.segments:
                self.preview_status.config(text=_("No path to optimize."))
                return
            try:
                optimized_segments = optimize_motion_segments(
                    self.model.segments,
                    start_point=self.model.start_point,
                    end_point=self.model.end_point,
                )
            except Exception as exc:  # pragma: no cover - GUI feedback
                self.preview_status.config(text=_("Optimization failed: ") + str(exc))
                return

            if optimized_segments is self.model.segments or optimized_segments == self.model.segments:
                self.preview_status.config(text=_("Path is already optimised."))
                return

            self.model = MotionPathModel(
                optimized_segments,
                px_to_mm=self.model.px_to_mm,
                doc_height_px=self.model.doc_height_px,
            )
            base_height_px = max(self.model.bounds[3] - self.model.bounds[1], 1e-3)
            self.base_row_distance_mm = base_height_px * self.model.px_to_mm
            self.row_distance_mm = self.base_row_distance_mm
            self.row_distance_var.set(self.row_distance_mm)
            self._refresh_y_warning()

            self.progress_mm = 0.0
            self.playing = False
            self.play_button.config(text=_("Play"))
            self.last_tick = time.monotonic()
            self._schedule_redraw()
            self.preview_status.config(text=_("Path optimised to reduce overlaps."))

        def _on_speed_changed(self, *_args) -> None:
            self.speed_multiplier = float(self.speed_var.get())
            self.speed_value_label.config(text=f"{self.speed_multiplier:.2f}×")

        def _on_repeat_changed(self) -> None:
            self.repeat_count = max(1, int(self.repeat_var.get()))
            self._schedule_redraw()

        def _on_rows_changed(self) -> None:
            self.row_count = max(1, int(self.rows_var.get()))
            self._schedule_redraw()

        def _on_row_distance_changed(self) -> None:
            self.row_distance_mm = max(1.0, float(self.row_distance_var.get()))
            self._schedule_redraw()

        def _on_stagger_toggled(self) -> None:
            self.stagger = bool(self.stagger_var.get())
            self._schedule_redraw()

        def _on_stagger_percent_changed(self, *_args) -> None:
            self.stagger_percent = float(self.stagger_percent_var.get())
            if self.stagger:
                self._schedule_redraw()

        def _on_export_layout_toggled(self) -> None:
            self.export_entire_layout = bool(self.export_layout_var.get())

        def _on_mirror_rows_toggled(self) -> None:
            self.mirror_alternate_rows = bool(self.mirror_rows_var.get())
            self._schedule_redraw()

        def _on_mirror_rows_v_toggled(self) -> None:
            self.mirror_alternate_rows_vertical = bool(self.mirror_rows_v_var.get())
            self._schedule_redraw()

        def _on_flip_h_toggled(self) -> None:
            self.flip_horizontal = bool(self.flip_h_var.get())
            self._schedule_redraw()

        def _on_flip_v_toggled(self) -> None:
            self.flip_vertical = bool(self.flip_v_var.get())
            self._schedule_redraw()

        def _export(self, *_args) -> None:
            active = self.format_combo.current()
            if active < 0:
                return
            format_key = list(self.exporters.keys())[active]
            profile = self.exporters[format_key]

            filename = filedialog.asksaveasfilename(
                title=_("Export Motion Path"),
                initialdir=str(Path.home()),
                initialfile=f"quilt_path.{profile.extension}",
                defaultextension=f".{profile.extension}",
                filetypes=[(profile.title, f"*.{profile.extension}")],
            )
            if not filename:
                return

            out_path = Path(filename)
            if out_path.suffix.lower() != f".{profile.extension.lower()}":
                out_path = out_path.with_suffix(f".{profile.extension.lower()}")

            try:
                out_path.parent.mkdir(parents=True, exist_ok=True)
            except PermissionError:
                self.export_status.config(text=_("No permission to create folder: ") + str(out_path.parent))
                return
            except FileExistsError:
                pass

            try:
                export_model = self._build_export_model() if self.export_entire_layout else self.model
                profile.writer(export_model, out_path)
            except PermissionError:
                self.export_status.config(
                    text=_("Cannot write to {path}. Pick a folder inside your home directory.").format(
                        path=str(out_path)
                    )
                )
                return
            except Exception as exc:  # pragma: no cover - GUI feedback
                self.export_status.config(text=_("Export failed: ") + str(exc))
                return

            self.export_status.config(text=_("Exported to ") + str(out_path))
            self._schedule_redraw()

        def _on_destroy(self) -> None:
            if self._tick_after_id is not None:
                self.root.after_cancel(self._tick_after_id)
                self._tick_after_id = None
            self.root.destroy()

        def _schedule_tick(self) -> None:
            self._tick_after_id = self.root.after(16, self._tick)

        def _tick(self) -> None:
            if not self.playing or not self.model.edges:
                self._schedule_tick()
                return
            now = time.monotonic()
            if self.last_tick is None:
                self.last_tick = now
                self._schedule_tick()
                return

            delta = now - self.last_tick
            self.last_tick = now

            advance = delta * self.BASE_SPEED_MM_PER_SEC * self.speed_multiplier
            self.progress_mm += advance
            if self.progress_mm >= self.model.total_length_mm:
                self.progress_mm = self.model.total_length_mm
                self.playing = False
                self.play_button.config(text=_("Play"))
            self._schedule_redraw()
            self._schedule_tick()

        def _on_canvas_resize(self, *_args) -> None:
            self._schedule_redraw()

        def _schedule_redraw(self) -> None:
            if self._redraw_pending:
                return
            self._redraw_pending = True
            self.root.after_idle(self._redraw)

        def _redraw(self) -> None:
            self._redraw_pending = False
            if not hasattr(self, "canvas"):
                return
            width = self.canvas.winfo_width()
            height = self.canvas.winfo_height()
            if width <= 1 or height <= 1:
                return

            self.canvas.delete("all")
            self.canvas.create_rectangle(0, 0, width, height, fill="#f8f8f8", outline="")

            if not self.model.edges:
                self.preview_status.config(text=_("Select at least one path to preview."))
                return

            self._viewport = self._compute_viewport(width, height)
            scale, offset_x, offset_y = self._viewport

            self._draw_full_pattern(scale, offset_x, offset_y)
            self._draw_progress(scale, offset_x, offset_y)

            if self._y_mismatch and self.model.start_point is not None:
                self._draw_warning_ring(self.model.start_point, scale, offset_x, offset_y)
            if self._y_mismatch and self.model.end_point is not None:
                self._draw_warning_ring(self.model.end_point, scale, offset_x, offset_y)

            point, needle_down = self.model.point_at(self.progress_mm)
            px = self._to_canvas(point, scale, offset_x, offset_y)
            r = max(2.0, 4.0 * scale * 0.2)
            color = "#e53935" if needle_down else "#f06292"
            self.canvas.create_oval(px[0] - r, px[1] - r, px[0] + r, px[1] + r, fill=color, outline="")

            stitched = min(self.progress_mm, self.model.total_length_mm)
            percent = (stitched / self.model.total_length_mm * 100.0) if self.model.total_length_mm else 0.0
            self._set_progress_slider_value(percent)
            self.preview_status.config(
                text=_("Path length: {length:.1f} mm   Previewed: {progress:.1f} mm ({pct:.1f}%)").format(
                    length=self.model.total_length_mm, progress=stitched, pct=percent
                )
            )

        def _set_progress_slider_value(self, percent: float) -> None:
            self._updating_progress_slider = True
            self.progress_var.set(max(0.0, min(100.0, percent)))
            self._updating_progress_slider = False

        def _on_progress_slider_changed(self, _value) -> None:
            if self._updating_progress_slider or not self.model.total_length_mm:
                return
            percent = float(self.progress_var.get())
            self.progress_mm = percent / 100.0 * self.model.total_length_mm
            self.last_tick = time.monotonic()
            self._schedule_redraw()

        def _pantograph_offsets(self) -> List[Tuple[int, float, float]]:
            width_px = max(self.model.bounds[2] - self.model.bounds[0], 1e-3)
            height_px = max(self.model.bounds[3] - self.model.bounds[1], 1e-3)
            self._pattern_width_px = width_px
            self._pattern_height_px = height_px

            return _compute_pantograph_offsets(
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

        def _pantograph_bounds(self) -> Tuple[float, float, float, float]:
            min_x, min_y, max_x, max_y = self.model.bounds
            offsets = self._pantograph_offsets()
            total_min_x = float("inf")
            total_min_y = float("inf")
            total_max_x = float("-inf")
            total_max_y = float("-inf")
            for _row_idx, dx, dy in offsets:
                total_min_x = min(total_min_x, min_x + dx)
                total_min_y = min(total_min_y, min_y + dy)
                total_max_x = max(total_max_x, max_x + dx)
                total_max_y = max(total_max_y, max_y + dy)
            if not offsets:
                total_min_x, total_min_y, total_max_x, total_max_y = min_x, min_y, max_x, max_y
            return (total_min_x, total_min_y, total_max_x, total_max_y)

        def _layout_bounds(self) -> Tuple[float, float, float, float]:
            return _compute_layout_bounds(
                self.model.bounds,
                repeat_count=self.repeat_count,
                row_count=self.row_count,
                row_distance_mm=self.row_distance_mm,
                px_to_mm=self.model.px_to_mm,
                start_point=self.model.start_point,
                end_point=self.model.end_point,
            )

        def _build_export_model(self) -> MotionPathModel:
            """Return a MotionPathModel representing the full layout."""
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

            stitched_segments: List[MotionSegment] = []
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
                        stitched_segments.append(MotionSegment(points=[last_end, pts[0]], needle_down=True))
                    stitched_segments.append(MotionSegment(points=pts, needle_down=seg.needle_down))
                    last_end = pts[-1]

            if not stitched_segments:
                return self.model

            return MotionPathModel(
                stitched_segments,
                px_to_mm=self.model.px_to_mm,
                doc_height_px=self.model.doc_height_px,
            )

        def _stroke_width(self) -> float:
            span = max(
                self.model.bounds[2] - self.model.bounds[0],
                self.model.bounds[3] - self.model.bounds[1],
            )
            if span <= 0:
                return 1.0
            width = (span / 300.0) ** 0.7
            return max(min(width, 10.0), 0.08)

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

        def _draw_full_pattern(self, scale: float, offset_x: float, offset_y: float) -> None:
            line_width = max(1.0, self._stroke_width() * scale)
            offsets = self._pantograph_offsets()

            for row_idx, dx, dy in offsets:
                mirror_row_h = self.mirror_alternate_rows and (row_idx % 2 == 1)
                mirror_row_v = self.mirror_alternate_rows_vertical and (row_idx % 2 == 1)
                for seg in self.model.segments:
                    color = "#2b6cb0" if seg.needle_down else "#d14343"
                    pts = [self._transform_point(pt, mirror_row_h, mirror_row_v) for pt in seg.points]
                    pts = [(p[0] + dx, p[1] + dy) for p in pts]
                    if len(pts) < 2:
                        continue
                    coords = []
                    for pt in pts:
                        cx, cy = self._to_canvas(pt, scale, offset_x, offset_y)
                        coords.extend([cx, cy])
                    self.canvas.create_line(*coords, fill=color, width=line_width)

        def _rgba_to_hex(self, rgba: Tuple[float, float, float, float]) -> str:
            r, g, b, _a = rgba
            return f"#{int(r * 255):02x}{int(g * 255):02x}{int(b * 255):02x}"

        def _draw_progress(self, scale: float, offset_x: float, offset_y: float) -> None:
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

            def _subdivide_and_draw(edge: MotionEdge) -> bool:
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
                        if len(covered) >= threshold:
                            passes_completed = max(covered)
                        else:
                            passes_completed = 0
                    else:
                        passes_completed = 0

                    color = self._rgba_to_hex(_color_for_pass(passes_completed, edge.needle_down))
                    start_px = self._to_canvas(draw_start, scale, offset_x, offset_y)
                    end_px = self._to_canvas(draw_end, scale, offset_x, offset_y)
                    self.canvas.create_line(*start_px, *end_px, fill=color, width=line_width)

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

        def _draw_warning_ring(self, point: Point, scale: float, offset_x: float, offset_y: float) -> None:
            px = self._to_canvas(point, scale, offset_x, offset_y)
            r = max(6.0, 9.0 * scale * 0.2)
            self.canvas.create_oval(
                px[0] - r,
                px[1] - r,
                px[0] + r,
                px[1] + r,
                outline="#b8860b",
                width=max(1.0, 2.0 * scale * 0.5),
            )

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
                message = _(
                    "WARNING: Start node and end node have different Y-axis positions (dY = {delta:.3f} mm > 0.1mm)"
                ).format(delta=delta_mm)
                self.y_warning_label.config(text=message)
            else:
                self.y_warning_label.config(text="")
else:
    class QuiltPreviewWindow:  # pragma: no cover - placeholder when Tk missing
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


def _write_qct_dxf(model: MotionPathModel, outfile: Path) -> None:
    def format_number(value: float) -> str:
        text = f"{value:.4f}".rstrip("0").rstrip(".")
        if text == "-0":
            text = "0"
        return text

    def add_code(lines: List[str], code: str) -> None:
        lines.append(f" {code} ")

    def add_text(lines: List[str], value: str) -> None:
        lines.append(value)

    def add_number(lines: List[str], value: float) -> None:
        lines.append(f"{format_number(value)} ")

    lines: List[str] = []
    add_code(lines, "0")
    add_text(lines, "SECTION")
    add_code(lines, "2")
    add_text(lines, "ENTITIES")

    for needle_down, pts in model.iter_segments_mm():
        if not needle_down or len(pts) < 2:
            continue
        for start, end in zip(pts, pts[1:]):
            x1, y1 = _cartesian_coords(model, start[0], start[1])
            x2, y2 = _cartesian_coords(model, end[0], end[1])
            add_code(lines, "0")
            add_text(lines, "LINE")
            add_code(lines, "8")
            add_text(lines, "Layer")
            add_code(lines, "10")
            add_number(lines, x1)
            add_code(lines, "20")
            add_number(lines, y1)
            add_code(lines, "11")
            add_number(lines, x2)
            add_code(lines, "21")
            add_number(lines, y2)

    add_code(lines, "0")
    add_text(lines, "ENDSEC")
    add_code(lines, "0")
    add_text(lines, "EOF")

    payload = "\r\n".join(lines) + "\r\n"
    outfile.write_bytes(payload.encode("ascii"))


def _write_gif(model: MotionPathModel, outfile: Path) -> None:
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow is required for GIF export.")
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
        stroke_colors.append("#008080" if seg.needle_down else "#f4a1a1")

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

            draw.line([start_pt, end_pt], fill="#003c83", width=3)

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


def _color_for_pass(passes_completed: int, needle_down: bool) -> Tuple[float, float, float, float]:
    """Return RGBA for a given pass count and stitch state."""
    if not needle_down:
        return (0.9, 0.15, 0.15, 0.9)  # red for needle-up jumps
    if passes_completed <= 0:
        return (0.1, 0.55, 0.85, 0.9)  # blue first pass
    if passes_completed == 1:
        return (0.95, 0.85, 0.25, 0.9)  # yellow second pass
    return (0.98, 0.55, 0.1, 0.9)  # orange 3+ passes


EXPORT_PROFILES: Dict[str, ExportProfile] = {
    "DXF": ExportProfile(
        title="AutoCAD DXF (polyline)",
        extension="dxf",
        description="Light‑weight polyline DXF",
        writer=_write_dxf,
    ),
    "QCT": ExportProfile(
        title="QCT DXF (lines)",
        extension="dxf",
        description="QCT-compatible line DXF",
        writer=_write_qct_dxf,
    ),
}
if PIL_AVAILABLE:
    EXPORT_PROFILES["GIF"] = ExportProfile(
        title="Animated GIF",
        extension="gif",
        description="Preview animation exported as GIF",
        writer=lambda model, dest: _write_gif(model, dest),
    )


class QuiltMotionExportExtension(inkex.EffectExtension):
    """Entry point for Inkscape."""

    def effect(self) -> None:  # pragma: no cover - Inkscape runtime
        if not TK_AVAILABLE:
            detail = _("Tkinter is required for the preview UI.")
            if TK_LOAD_ERROR:
                detail = f"{detail}\n{TK_LOAD_ERROR}"
            raise inkex.AbortExtension(detail)

        selection = self.svg.selection.filter(PathElement)
        if not selection:
            raise inkex.AbortExtension(_("Please select at least one path."))

        tolerance = 0.4
        try:
            bbox = selection.bounding_box()
        except Exception:
            bbox = None
        if bbox is not None:
            major = max(bbox.width, bbox.height)
            if major and major > 0:
                tolerance = max(min(major / 500.0, 0.4), 0.02)
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


if __name__ == "__main__":  # pragma: no cover
    QuiltMotionExportExtension().run()
