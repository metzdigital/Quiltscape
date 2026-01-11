#!/usr/bin/env python3
"""
Core motion model + exporters for Quilt Motion Preview & Export.
"""

from __future__ import annotations

import math
import heapq
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    from PIL import Image, ImageDraw

    PIL_AVAILABLE = True
    PIL_LOAD_ERROR: Optional[str] = None
except Exception as exc:  # pragma: no cover - optional dependency
    Image = None  # type: ignore
    ImageDraw = None  # type: ignore
    PIL_AVAILABLE = False
    PIL_LOAD_ERROR = str(exc)


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
    row_spacing_px = height_px + (row_distance_mm / px_to_mm)
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
    row_spacing_px = height + (row_distance_mm / px_to_mm)

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


def optimize_motion_segments(
    segments: List[MotionSegment],
    start_point: Optional[Point] = None,
    end_point: Optional[Point] = None,
    tolerance: float = 1e-6,
) -> List[MotionSegment]:
    """Reorder a stitched path using junctions and shortest coverage."""

    def quantize_point(point: Point) -> Tuple[float, float]:
        return (round(point[0], 6), round(point[1], 6))

    def segment_intersections(a: Point, b: Point, c: Point, d: Point) -> List[Point]:
        def _on_segment(p: Point, q: Point, r: Point) -> bool:
            return min(p[0], r[0]) - 1e-9 <= q[0] <= max(p[0], r[0]) + 1e-9 and min(p[1], r[1]) - 1e-9 <= q[1] <= max(p[1], r[1]) + 1e-9

        def _orient(p: Point, q: Point, r: Point) -> float:
            return (q[0] - p[0]) * (r[1] - p[1]) - (q[1] - p[1]) * (r[0] - p[0])

        o1 = _orient(a, b, c)
        o2 = _orient(a, b, d)
        o3 = _orient(c, d, a)
        o4 = _orient(c, d, b)

        intersections: List[Point] = []

        if math.isclose(o1, 0.0, abs_tol=1e-9) and _on_segment(a, c, b):
            intersections.append(c)
        if math.isclose(o2, 0.0, abs_tol=1e-9) and _on_segment(a, d, b):
            intersections.append(d)
        if math.isclose(o3, 0.0, abs_tol=1e-9) and _on_segment(c, a, d):
            intersections.append(a)
        if math.isclose(o4, 0.0, abs_tol=1e-9) and _on_segment(c, b, d):
            intersections.append(b)

        if (o1 > 0 and o2 < 0 or o1 < 0 and o2 > 0) and (o3 > 0 and o4 < 0 or o3 < 0 and o4 > 0):
            denom = (a[0] - b[0]) * (c[1] - d[1]) - (a[1] - b[1]) * (c[0] - d[0])
            if not math.isclose(denom, 0.0, abs_tol=1e-12):
                px = (
                    (a[0] * b[1] - a[1] * b[0]) * (c[0] - d[0])
                    - (a[0] - b[0]) * (c[0] * d[1] - c[1] * d[0])
                ) / denom
                py = (
                    (a[0] * b[1] - a[1] * b[0]) * (c[1] - d[1])
                    - (a[1] - b[1]) * (c[0] * d[1] - c[1] * d[0])
                ) / denom
                intersections.append((px, py))

        return intersections

    stitched_segments = [seg for seg in segments if seg.needle_down and len(seg.points) >= 2]
    if not stitched_segments:
        return segments

    edges: List[Tuple[Point, Point]] = []
    for seg in stitched_segments:
        edges.extend([(a, b) for a, b in zip(seg.points, seg.points[1:])])
    if not edges:
        return segments

    # Split edges at intersections.
    split_edges: List[Tuple[Point, Point]] = []
    for idx, (a, b) in enumerate(edges):
        if math.isclose(a[0], b[0], abs_tol=1e-9) and math.isclose(a[1], b[1], abs_tol=1e-9):
            continue
        points = [a, b]
        for jdx, (c, d) in enumerate(edges):
            if idx == jdx:
                continue
            for ipt in segment_intersections(a, b, c, d):
                points.append(ipt)
        unique: List[Point] = []
        for pt in points:
            if not any(math.isclose(pt[0], q[0], abs_tol=1e-6) and math.isclose(pt[1], q[1], abs_tol=1e-6) for q in unique):
                unique.append(pt)
        if abs(a[0] - b[0]) >= abs(a[1] - b[1]):
            def tval(p: Point) -> float:
                return (p[0] - a[0]) / (b[0] - a[0]) if not math.isclose(a[0], b[0], abs_tol=1e-9) else 0.0
        else:
            def tval(p: Point) -> float:
                return (p[1] - a[1]) / (b[1] - a[1]) if not math.isclose(a[1], b[1], abs_tol=1e-9) else 0.0
        unique.sort(key=tval)
        for p0, p1 in zip(unique, unique[1:]):
            if math.isclose(p0[0], p1[0], abs_tol=1e-9) and math.isclose(p0[1], p1[1], abs_tol=1e-9):
                continue
            split_edges.append((p0, p1))

    def edge_key(a: Point, b: Point) -> Tuple[Tuple[float, float], Tuple[float, float]]:
        ak = quantize_point(a)
        bk = quantize_point(b)
        return (ak, bk) if ak <= bk else (bk, ak)

    unique_edges: Dict[Tuple[Tuple[float, float], Tuple[float, float]], Tuple[Point, Point]] = {}
    for a, b in split_edges:
        unique_edges[edge_key(a, b)] = (a, b)
    if not unique_edges:
        return segments

    # Build graph.
    nodes: Dict[Tuple[float, float], int] = {}
    node_points: List[Point] = []

    def node_id(pt: Point) -> int:
        key = quantize_point(pt)
        if key not in nodes:
            nodes[key] = len(node_points)
            node_points.append(pt)
        return nodes[key]

    edge_list: List[Tuple[int, int, float]] = []
    adjacency: Dict[int, List[Tuple[int, int, float]]] = {}
    for a, b in unique_edges.values():
        u = node_id(a)
        v = node_id(b)
        weight = math.dist(a, b)
        edge_idx = len(edge_list)
        edge_list.append((u, v, weight))
        adjacency.setdefault(u, []).append((v, edge_idx, weight))
        adjacency.setdefault(v, []).append((u, edge_idx, weight))

    if start_point is None:
        start_point = stitched_segments[0].points[0]
    if end_point is None:
        end_point = stitched_segments[-1].points[-1]

    start_id = node_id(start_point)
    end_id = node_id(end_point)

    degrees = {node: len(nei) for node, nei in adjacency.items()}
    odd_nodes = [node for node, deg in degrees.items() if deg % 2 == 1]
    if len(odd_nodes) < 2:
        odd_nodes = []

    # Ensure the stitched graph is connected.
    visited = set()
    stack = [start_id]
    while stack:
        node = stack.pop()
        if node in visited:
            continue
        visited.add(node)
        stack.extend([n for n, _e, _w in adjacency.get(node, [])])
    if len(visited) != len(adjacency):
        return segments

    # Dijkstra for shortest paths between odd nodes.
    def shortest_path(src: int, dst: int) -> Tuple[float, List[int]]:
        dist: Dict[int, float] = {src: 0.0}
        prev: Dict[int, Tuple[int, int]] = {}
        heap: List[Tuple[float, int]] = [(0.0, src)]
        while heap:
            d, node = heapq.heappop(heap)
            if node == dst:
                break
            if d > dist.get(node, float("inf")) + 1e-12:
                continue
            for nxt, edge_idx, weight in adjacency.get(node, []):
                nd = d + weight
                if nd < dist.get(nxt, float("inf")) - 1e-12:
                    dist[nxt] = nd
                    prev[nxt] = (node, edge_idx)
                    heapq.heappush(heap, (nd, nxt))
        if dst not in dist:
            return float("inf"), []
        # reconstruct edge path
        edge_path: List[int] = []
        node = dst
        while node != src:
            node, edge_idx = prev[node]
            edge_path.append(edge_idx)
        edge_path.reverse()
        return dist[dst], edge_path

    odd_ids = odd_nodes[:]
    best_endpoints = (start_id, end_id)
    best_pairs: List[Tuple[int, int, List[int]]] = []

    def greedy_pairing(nodes_list: List[int]) -> Tuple[float, List[Tuple[int, int, List[int]]]]:
        remaining = nodes_list[:]
        pairs: List[Tuple[int, int, List[int]]] = []
        total = 0.0
        while remaining:
            a = remaining.pop(0)
            best = None
            best_cost = float("inf")
            best_path: List[int] = []
            for b in remaining:
                cost, path = shortest_path(a, b)
                if cost < best_cost:
                    best_cost = cost
                    best = b
                    best_path = path
            if best is None:
                break
            remaining.remove(best)
            total += best_cost
            pairs.append((a, best, best_path))
        return total, pairs

    pairing_paths: List[Tuple[int, int, List[int]]] = []
    if start_id != end_id:
        start_is_odd = start_id in odd_ids
        end_is_odd = end_id in odd_ids
        if start_is_odd and end_is_odd:
            pair_nodes = [n for n in odd_ids if n not in (start_id, end_id)]
        elif start_is_odd and not end_is_odd:
            candidates = [n for n in odd_ids if n != start_id]
            if not candidates:
                return segments
            best_cost = float("inf")
            best_choice: Optional[Tuple[int, List[int]]] = None
            for candidate in candidates:
                cost, path = shortest_path(end_id, candidate)
                if cost < best_cost:
                    best_cost = cost
                    best_choice = (candidate, path)
            if best_choice is None or not best_choice[1]:
                return segments
            pairing_paths.append((end_id, best_choice[0], best_choice[1]))
            pair_nodes = [n for n in odd_ids if n not in (start_id, best_choice[0])]
        elif not start_is_odd and end_is_odd:
            candidates = [n for n in odd_ids if n != end_id]
            if not candidates:
                return segments
            best_cost = float("inf")
            best_choice = None
            for candidate in candidates:
                cost, path = shortest_path(start_id, candidate)
                if cost < best_cost:
                    best_cost = cost
                    best_choice = (candidate, path)
            if best_choice is None or not best_choice[1]:
                return segments
            pairing_paths.append((start_id, best_choice[0], best_choice[1]))
            pair_nodes = [n for n in odd_ids if n not in (end_id, best_choice[0])]
        else:
            cost, path = shortest_path(start_id, end_id)
            if cost == float("inf") or not path:
                return segments
            pairing_paths.append((start_id, end_id, path))
            pair_nodes = odd_ids[:]
    else:
        pair_nodes = odd_ids[:]

    if pair_nodes:
        _cost, pairs = greedy_pairing(pair_nodes)
        best_pairs = pairing_paths + pairs
    else:
        best_pairs = pairing_paths

    # Build multigraph by duplicating edges along matched paths.
    multigraph: Dict[int, List[Tuple[int, int]]] = {}
    edge_instances: List[Tuple[int, int]] = []

    def add_edge(u: int, v: int) -> None:
        edge_id = len(edge_instances)
        edge_instances.append((u, v))
        multigraph.setdefault(u, []).append((v, edge_id))
        multigraph.setdefault(v, []).append((u, edge_id))

    for u, v, _w in edge_list:
        add_edge(u, v)

    for _a, _b, edge_path in best_pairs:
        for edge_idx in edge_path:
            u, v, _w = edge_list[edge_idx]
            add_edge(u, v)

    # Hierholzer's algorithm for Euler trail.
    start_node, end_node = best_endpoints
    stack = [start_node]
    trail: List[int] = []
    multigraph_copy: Dict[int, List[Tuple[int, int]]] = {k: v[:] for k, v in multigraph.items()}

    while stack:
        v = stack[-1]
        if multigraph_copy.get(v):
            nxt, edge_id = multigraph_copy[v].pop()
            neighbor_list = multigraph_copy.get(nxt, [])
            for idx, (_nn, eid) in enumerate(neighbor_list):
                if eid == edge_id:
                    neighbor_list.pop(idx)
                    break
            stack.append(nxt)
        else:
            trail.append(stack.pop())
    trail.reverse()

    if not trail or trail[0] != start_node or trail[-1] != end_node:
        return segments

    optimized_points = [node_points[node] for node in trail]
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

    def stitched_length(segment_list: List[MotionSegment]) -> float:
        total = 0.0
        for segment in segment_list:
            if not segment.needle_down or len(segment.points) < 2:
                continue
            for start, end in zip(segment.points, segment.points[1:]):
                total += math.dist(start, end)
        return total

    if stitched_length(optimized_segments) + tolerance < stitched_length(segments):
        return optimized_segments
    return optimized_segments


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
        return (0.9, 0.15, 0.15, 0.9)
    if passes_completed <= 0:
        return (0.1, 0.55, 0.85, 0.9)
    if passes_completed == 1:
        return (0.95, 0.85, 0.25, 0.9)
    return (0.98, 0.55, 0.1, 0.9)


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
