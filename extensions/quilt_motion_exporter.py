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
            self._y_mismatch = False
            self.flip_horizontal = False
            self.flip_vertical = False
            self.mirror_alternate_rows = False
            self.mirror_alternate_rows_vertical = False

            self._build_ui()
            self._refresh_y_warning()
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

            left_column = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
            left_column.set_hexpand(True)
            left_column.set_vexpand(True)
            root.pack_start(left_column, True, True, 0)

            self.drawing_area = Gtk.DrawingArea()
            self.drawing_area.set_hexpand(True)
            self.drawing_area.set_vexpand(True)
            self.drawing_area.connect("draw", self._on_draw)
            self.drawing_area.connect("size-allocate", self._on_size_allocate)
            left_column.pack_start(self.drawing_area, True, True, 0)

            self.y_warning_label = Gtk.Label()
            self.y_warning_label.set_xalign(0.0)
            self.y_warning_label.set_use_markup(True)
            left_column.pack_start(self.y_warning_label, False, False, 0)

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

            optimize_button = Gtk.Button(label=_("Optimize path"))
            optimize_button.connect("clicked", self._optimize_path)
            controls.pack_start(optimize_button, True, True, 0)
            self.optimize_button = optimize_button

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

            mirror_toggle = Gtk.CheckButton(label=_("Mirror every other row horizontally"))
            mirror_toggle.set_active(self.mirror_alternate_rows)
            mirror_toggle.connect("toggled", self._on_mirror_rows_toggled)
            pantograph_grid.attach(mirror_toggle, 0, 5, 2, 1)
            self.mirror_toggle = mirror_toggle

            mirror_v_toggle = Gtk.CheckButton(label=_("Mirror every other row vertically"))
            mirror_v_toggle.set_active(self.mirror_alternate_rows_vertical)
            mirror_v_toggle.connect("toggled", self._on_mirror_rows_v_toggled)
            pantograph_grid.attach(mirror_v_toggle, 0, 6, 2, 1)
            self.mirror_v_toggle = mirror_v_toggle

            flip_h_toggle = Gtk.CheckButton(label=_("Flip horizontally"))
            flip_h_toggle.set_active(self.flip_horizontal)
            flip_h_toggle.connect("toggled", self._on_flip_h_toggled)
            pantograph_grid.attach(flip_h_toggle, 0, 7, 2, 1)
            self.flip_h_toggle = flip_h_toggle

            flip_v_toggle = Gtk.CheckButton(label=_("Flip vertically"))
            flip_v_toggle.set_active(self.flip_vertical)
            flip_v_toggle.connect("toggled", self._on_flip_v_toggled)
            pantograph_grid.attach(flip_v_toggle, 0, 8, 2, 1)
            self.flip_v_toggle = flip_v_toggle

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
            if hasattr(self, "_edge_progress"):
                self._edge_progress = {}

        def _optimize_path(self, *_args) -> None:
            if not self.model.segments:
                self.preview_status.set_text(_("No path to optimize."))
                return
            try:
                optimized_segments = optimize_motion_segments(
                    self.model.segments,
                    start_point=self.model.start_point,
                    end_point=self.model.end_point,
                )
            except Exception as exc:  # pragma: no cover - GUI feedback
                self.preview_status.set_text(_("Optimization failed: ") + str(exc))
                return

            if optimized_segments is self.model.segments:
                self.preview_status.set_text(_("Path is already optimised."))
                return

            if optimized_segments == self.model.segments:
                self.preview_status.set_text(_("Path is already optimised."))
                return

            self.model = MotionPathModel(
                optimized_segments,
                px_to_mm=self.model.px_to_mm,
                doc_height_px=self.model.doc_height_px,
            )
            base_height_px = max(self.model.bounds[3] - self.model.bounds[1], 1e-3)
            self.base_row_distance_mm = base_height_px * self.model.px_to_mm
            self.row_distance_mm = self.base_row_distance_mm
            if hasattr(self, "row_distance_spin"):
                adjustment = self.row_distance_spin.get_adjustment()
                adjustment.set_value(self.row_distance_mm)
            self._refresh_y_warning()

            self.progress_mm = 0.0
            self.playing = False
            self.play_button.set_label(_("Play"))
            self.last_tick = time.monotonic()
            self._surface_dirty = True
            self.drawing_area.queue_draw()
            self.preview_status.set_text(_("Path optimised to reduce overlaps."))

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

        def _on_mirror_rows_toggled(self, button: Gtk.CheckButton) -> None:
            self.mirror_alternate_rows = button.get_active()
            self._invalidate_surface()

        def _on_mirror_rows_v_toggled(self, button: Gtk.CheckButton) -> None:
            self.mirror_alternate_rows_vertical = button.get_active()
            self._invalidate_surface()

        def _on_flip_h_toggled(self, button: Gtk.CheckButton) -> None:
            self.flip_horizontal = button.get_active()
            self._invalidate_surface()

        def _on_flip_v_toggled(self, button: Gtk.CheckButton) -> None:
            self.flip_vertical = button.get_active()
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

            transform = lambda p: self._transform_point(p, False, False)

            self._draw_progress(cr)

            if self._y_mismatch and self.model.start_point is not None:
                cr.save()
                start_pt = transform(self.model.start_point)
                cr.translate(start_pt[0], start_pt[1])
                cr.set_source_rgba(0.8, 0.62, 0.06, 1.0)
                cr.set_line_width((self._stroke_width() * 10.0) / scale)
                radius = 9.0 / scale
                cr.arc(0, 0, radius, 0, math.tau)
                cr.stroke()
                cr.restore()

            if self._y_mismatch and self.model.end_point is not None:
                cr.save()
                end_pt = transform(self.model.end_point)
                cr.translate(end_pt[0], end_pt[1])
                cr.set_source_rgba(0.8, 0.62, 0.06, 1.0)
                cr.set_line_width((self._stroke_width() * 10.0) / scale)
                radius = 9.0 / scale
                cr.arc(0, 0, radius, 0, math.tau)
                cr.stroke()
                cr.restore()

            point, needle_down = self.model.point_at(self.progress_mm)
            cr.save()
            point = transform(point)
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

        def _pantograph_offsets(self) -> List[Tuple[int, float, float]]:
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

            offsets: List[Tuple[int, float, float]] = []
            for row in range(self.row_count):
                base_dx = stagger_px if (self.stagger and row % 2 == 1) else 0.0
                row_dy = row * row_spacing_px
                for repeat in range(self.repeat_count):
                    dx = base_dx + repeat * delta_x
                    dy = row_dy + repeat * delta_y
                    offsets.append((row, dx, dy))
            return offsets

        def _transform_point(self, point: Point, mirror_row_h: bool, mirror_row_v: bool) -> Point:
            """Apply flip/mirror transforms around the pattern centre."""
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

        def _stroke_width(self) -> float:
            span = max(
                self.model.bounds[2] - self.model.bounds[0],
                self.model.bounds[3] - self.model.bounds[1],
            )
            if span <= 0:
                return 1.0
            width = (span / 300.0) ** 0.7  # slightly heavier scaling for large layouts
            return max(min(width, 10.0), 0.08)

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
            cr.set_line_width(self._stroke_width())
            offsets = self._pantograph_offsets()

            for row_idx, dx, dy in offsets:
                cr.save()
                cr.translate(dx, dy)
                for seg in self.model.segments:
                    color = (0.2, 0.4, 0.7, 0.7) if seg.needle_down else (0.85, 0.2, 0.2, 0.8)
                    cr.set_source_rgba(*color)
                    cr.new_path()
                    mirror_row_h = self.mirror_alternate_rows and (row_idx % 2 == 1)
                    mirror_row_v = self.mirror_alternate_rows_vertical and (row_idx % 2 == 1)
                    pts = [self._transform_point(pt, mirror_row_h, mirror_row_v) for pt in seg.points]
                    cr.move_to(*pts[0])
                    for pt in pts[1:]:
                        cr.line_to(*pt)
                    cr.stroke()
                cr.restore()
            cr.restore()

        def _draw_progress(self, cr) -> None:
            cr.save()
            cr.set_line_width(self._stroke_width() * 1.8)
            remaining = self.progress_mm
            transform = lambda p: self._transform_point(p, False, False)

            def _edge_key(a: Point, b: Point) -> Tuple[Tuple[float, float], Tuple[float, float]]:
                p0 = (round(a[0], 6), round(a[1], 6))
                p1 = (round(b[0], 6), round(b[1], 6))
                return (p0, p1) if p0 <= p1 else (p1, p0)

            # Spatial coverage grid for previously completed edges (or drawn portions).
            # Coverage is applied only after an edge finishes to avoid self-triggered retrace coloring.
            cell_size = self._stroke_width() * 0.01
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

                # Subdivide the edge into up to 12 chunks (more responsive color for longer edges).
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
                    # Overlap detection: require a modest fraction of samples already covered to count as a retrace.
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
                    cr.set_source_rgba(*_color_for_pass(passes_completed, edge.needle_down))

                    cr.move_to(*draw_start)
                    cr.line_to(*draw_end)
                    cr.stroke()

                    edge_cells.extend(cells)

                    # If we haven't finished this edge, stop drawing further edges.
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

            cr.restore()

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
                message = _("WARNING: Start node and end node have different Y-axis positions (dY = {delta:.3f} mm > 0.1mm)").format(
                    delta=delta_mm
                )
                self.y_warning_label.set_markup(f'<span foreground="#b8860b">{message}</span>')
            else:
                self.y_warning_label.set_markup("")
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
        Gtk.main()


if __name__ == "__main__":  # pragma: no cover
    QuiltMotionExportExtension().run()
