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
            dy = row_dy + repeat * delta_y
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
    """Reorder a continuous stitched path to reduce geometric overlaps."""

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

    def split_path(points: List[Point]) -> List[List[Point]]:
        if len(points) < 2:
            return [points]

        split_pts: List[Point] = []
        for i in range(len(points) - 1):
            a = points[i]
            b = points[i + 1]
            split_pts.append(a)
            for j in range(i + 1, len(points) - 1):
                c = points[j]
                d = points[j + 1]
                intersections = segment_intersections(a, b, c, d)
                for ipt in intersections:
                    if (
                        math.isclose(ipt[0], a[0], abs_tol=1e-6)
                        and math.isclose(ipt[1], a[1], abs_tol=1e-6)
                    ) or (
                        math.isclose(ipt[0], b[0], abs_tol=1e-6)
                        and math.isclose(ipt[1], b[1], abs_tol=1e-6)
                    ):
                        continue
                    split_pts.append(ipt)
        split_pts.append(points[-1])

        refined = sorted(split_pts, key=lambda p: (p[0], p[1]))
        deduped: List[Point] = []
        for pt in refined:
            if not deduped or not (math.isclose(pt[0], deduped[-1][0], abs_tol=1e-6) and math.isclose(pt[1], deduped[-1][1], abs_tol=1e-6)):
                deduped.append(pt)

        segments_out: List[List[Point]] = []
        for i in range(len(deduped) - 1):
            segments_out.append([deduped[i], deduped[i + 1]])
        return segments_out

    stitched_segments = [seg for seg in segments if seg.needle_down and len(seg.points) >= 2]
    if not stitched_segments:
        return segments

    stitched_points: List[Point] = []
    for seg in stitched_segments:
        stitched_points.extend(seg.points)
    if not stitched_points:
        return segments

    baseline_edge_counts = stitched_edge_counts(segments)
    if not baseline_edge_counts:
        return segments

    graph: Dict[Point, List[Point]] = {}
    for seg in stitched_segments:
        for i in range(len(seg.points) - 1):
            a = seg.points[i]
            b = seg.points[i + 1]
            if math.isclose(a[0], b[0], abs_tol=1e-9) and math.isclose(a[1], b[1], abs_tol=1e-9):
                continue
            graph.setdefault(a, []).append(b)
            graph.setdefault(b, []).append(a)

    if start_point is None:
        start_point = stitched_segments[0].points[0]
    if end_point is None:
        end_point = stitched_segments[-1].points[-1]

    if start_point not in graph or end_point not in graph:
        return segments

    degrees = {node: len(neighbors) for node, neighbors in graph.items()}
    odd_nodes = [node for node, deg in degrees.items() if deg % 2 == 1]
    if odd_nodes and (start_point not in odd_nodes or end_point not in odd_nodes):
        return segments

    # Ensure the stitched graph is connected.
    visited = set()
    stack = [start_point]
    while stack:
        node = stack.pop()
        if node in visited:
            continue
        visited.add(node)
        stack.extend(graph.get(node, []))
    if len(visited) != len(graph):
        return segments

    def edge_key(a: Point, b: Point) -> Tuple[Point, Point]:
        return (a, b) if a <= b else (b, a)

    edge_weights: Dict[Tuple[Point, Point], float] = {}
    for node, neighbors in graph.items():
        for neighbor in neighbors:
            key = edge_key(node, neighbor)
            if key not in edge_weights:
                edge_weights[key] = math.dist(node, neighbor)

    # Add extra edges to make all degrees even, except start/end.
    odds = [node for node, deg in degrees.items() if deg % 2 == 1]
    if odds:
        # Pair odd nodes with minimum distance matching.
        heap: List[Tuple[float, Tuple[Point, Point]]] = []
        for i, node in enumerate(odds):
            for other in odds[i + 1 :]:
                dist = math.dist(node, other)
                heapq.heappush(heap, (dist, (node, other)))

        pairs: List[Tuple[Point, Point]] = []
        used = set()
        while heap:
            _dist, (a, b) = heapq.heappop(heap)
            if a in used or b in used:
                continue
            used.add(a)
            used.add(b)
            pairs.append((a, b))
        for a, b in pairs:
            graph[a].append(b)
            graph[b].append(a)
            degrees[a] += 1
            degrees[b] += 1
            edge_weights[edge_key(a, b)] = math.dist(a, b)

    # Hierholzer's algorithm for Eulerian trail.
    graph_copy: Dict[Point, List[Point]] = {node: neighbors[:] for node, neighbors in graph.items()}
    path: List[Point] = []
    stack = [start_point]
    while stack:
        v = stack[-1]
        if graph_copy[v]:
            u = graph_copy[v].pop()
            graph_copy[u].remove(v)
            stack.append(u)
        else:
            path.append(stack.pop())
    path.reverse()

    if not path or path[0] != start_point or path[-1] != end_point:
        return segments

    optimized_points: List[Point] = []
    for point in path:
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
            return segments

    original_overlap = overlap_length(segments, baseline_edge_counts)
    optimized_overlap = overlap_length(optimized_segments, baseline_edge_counts)

    if optimized_overlap + tolerance >= original_overlap:
        return segments

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
