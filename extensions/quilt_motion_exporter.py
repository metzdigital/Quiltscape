#!/usr/bin/env python3
"""
Inkscape Quilt Motion Preview & Export extension.

This script gathers the selected path elements from the active document,
preserves their draw order, launches a standalone preview application, and
exports the resulting stitch path to several long‑arm quilting formats.
"""

from __future__ import annotations

import json
import math
import time
import heapq
import os
import shutil
import subprocess
import tempfile
import threading
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
    start_x = start_point[0] if start_point is not None else min_x
    end_x = end_point[0] if end_point is not None else max_x
    for row in range(row_count):
        base_dx = stagger_px if (stagger and row % 2 == 1) else 0.0
        row_dy = row * row_spacing_px
        row_dx: List[float] = [base_dx + repeat * delta_x for repeat in range(repeat_count)]
        row_dx.sort()
        if row_dx:
            while start_x + row_dx[0] > target_min_x + 1e-6:
                row_dx.insert(0, row_dx[0] - delta_x)
            while end_x + row_dx[-1] < target_max_x - 1e-6:
                row_dx.append(row_dx[-1] + delta_x)
        for dx in row_dx:
            offsets.append((row, dx, row_dy))
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
            dy = row_dy
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
        if (
            len(points) < 3
            and math.isclose(points[0][0], points[-1][0], abs_tol=1e-9)
            and math.isclose(points[0][1], points[-1][1], abs_tol=1e-9)
        ):
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
    "DXF": ExportProfile(
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


def _preview_payload(
    segments: List[MotionSegment],
    px_to_mm: float,
    doc_height_px: Optional[float],
) -> Dict[str, object]:
    return {
        "px_to_mm": px_to_mm,
        "doc_height_px": doc_height_px,
        "segments": [
            {"needle_down": seg.needle_down, "points": [list(pt) for pt in seg.points]}
            for seg in segments
        ],
    }


def _find_preview_python() -> str:
    config_path = _EXTENSION_DIR / "quilt_motion_preview_python.txt"
    if config_path.exists():
        configured = config_path.read_text(encoding="utf-8").strip()
        if configured:
            return configured
    override = os.environ.get("QUILT_PREVIEW_PYTHON")
    if override:
        return override
    return shutil.which("python3") or shutil.which("python") or sys.executable


def _launch_preview_app(payload: Dict[str, object]) -> None:
    preview_script = _EXTENSION_DIR / "quilt_motion_preview_app.py"
    if not preview_script.exists():
        raise FileNotFoundError(f"Preview app not found: {preview_script}")

    with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as handle:
        json.dump(payload, handle)
        temp_path = handle.name

    python_exe = _find_preview_python()
    process = subprocess.Popen(
        [python_exe, str(preview_script), "--input", temp_path, "--delete-input"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
    )
    threading.Thread(target=process.wait, daemon=True).start()


class QuiltMotionExportExtension(inkex.EffectExtension):
    """Entry point for Inkscape."""

    def effect(self) -> None:  # pragma: no cover - Inkscape runtime
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
        payload = _preview_payload(model.segments, px_to_mm, doc_height_px)
        try:
            _launch_preview_app(payload)
        except Exception as exc:
            raise inkex.AbortExtension(_("Failed to launch preview app: ") + str(exc))


if __name__ == "__main__":  # pragma: no cover
    QuiltMotionExportExtension().run()
