import unittest
from pathlib import Path

import sys

ROOT = Path(__file__).resolve().parents[1]
EXT_DIR = ROOT / "extensions"
if str(EXT_DIR) not in sys.path:
    sys.path.insert(0, str(EXT_DIR))

try:
    import inkex  # type: ignore
    from inkex import bezier
    from inkex.paths import CubicSuperPath
except Exception:
    inkex = None

import quilt_motion_core as qmc  # noqa: E402


def _stitched_length(segments):
    total = 0.0
    for seg in segments:
        if not seg.needle_down or len(seg.points) < 2:
            continue
        for a, b in zip(seg.points, seg.points[1:]):
            total += ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2) ** 0.5
    return total


def _flatten_path_element(element, tolerance=0.5):
    transform = element.composed_transform()
    path = element.path.transform(transform)
    csp: CubicSuperPath = path.to_superpath()
    bezier.cspsubdiv(csp, tolerance)

    segments = []
    previous_end = None
    for subpath in csp:
        points = [tuple(node[1]) for node in subpath]
        if len(points) < 2:
            continue
        if (
            len(points) < 3
            and abs(points[0][0] - points[-1][0]) <= 1e-9
            and abs(points[0][1] - points[-1][1]) <= 1e-9
        ):
            continue
        if previous_end and (
            abs(previous_end[0] - points[0][0]) > 1e-6
            or abs(previous_end[1] - points[0][1]) > 1e-6
        ):
            segments.append(qmc.MotionSegment(points=[previous_end, points[0]], needle_down=False))
        segments.append(qmc.MotionSegment(points=list(points), needle_down=True))
        previous_end = points[-1]
    return segments


class OptimizeSvgTests(unittest.TestCase):
    def test_optimize_reduces_length_for_optimize_test_svg(self) -> None:
        if inkex is None:
            self.skipTest("inkex not available")
        svg_path = ROOT / "optimize_test.svg"
        if not svg_path.exists():
            self.skipTest("optimize_test.svg not found")
        with svg_path.open("rb") as handle:
            doc = inkex.load_svg(handle)
        root = doc.getroot()
        paths = root.xpath(".//svg:path", namespaces={"svg": "http://www.w3.org/2000/svg"})
        self.assertTrue(paths, "No paths found in optimize_test.svg")
        segments = []
        for path in paths:
            segments.extend(_flatten_path_element(path, tolerance=0.4))
        self.assertTrue(segments, "No segments generated from optimize_test.svg")

        start_point = segments[0].points[0]
        end_point = segments[-1].points[-1]
        original_len = _stitched_length(segments)
        optimized = qmc.optimize_motion_segments(segments, start_point=start_point, end_point=end_point)
        optimized_len = _stitched_length(optimized)

        self.assertLess(
            optimized_len,
            original_len,
            f"Expected optimization to reduce length (orig={original_len}, opt={optimized_len})",
        )
        self.assertAlmostEqual(
            optimized[0].points[0][0], start_point[0], places=6
        )
        self.assertAlmostEqual(
            optimized[0].points[0][1], start_point[1], places=6
        )
        self.assertAlmostEqual(
            optimized[0].points[-1][0], end_point[0], places=6
        )
        self.assertAlmostEqual(
            optimized[0].points[-1][1], end_point[1], places=6
        )
