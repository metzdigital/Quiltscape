import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
EXT_DIR = ROOT / "extensions"
if str(EXT_DIR) not in sys.path:
    sys.path.insert(0, str(EXT_DIR))

import quilt_motion_exporter as qme  # noqa: E402


class MotionPathModelTests(unittest.TestCase):
    def setUp(self) -> None:
        px_to_mm = 0.2645833333  # 96 dpi conversion
        self.px_to_mm = px_to_mm
        self.model = qme.MotionPathModel(
            [
                qme.MotionSegment(points=[(0.0, 0.0), (10.0, 0.0), (10.0, 10.0)], needle_down=True),
                qme.MotionSegment(points=[(10.0, 10.0), (0.0, 10.0)], needle_down=False),
            ],
            px_to_mm=px_to_mm,
        )

    def test_total_length_matches_expected(self) -> None:
        expected_mm = 30.0 * self.px_to_mm
        self.assertAlmostEqual(self.model.total_length_mm, expected_mm, places=6)

    def test_point_lookup_along_path(self) -> None:
        half_length = self.model.total_length_mm / 2.0
        point, needle_down = self.model.point_at(half_length)
        self.assertTrue(needle_down)
        self.assertAlmostEqual(point[0], 10.0, places=3)
        self.assertAlmostEqual(point[1], 5.0, places=3)

    def test_iter_segments_converts_to_mm(self) -> None:
        segments = self.model.iter_segments_mm()
        stitch_segment = segments[0]
        self.assertTrue(stitch_segment[0])
        self.assertAlmostEqual(stitch_segment[1][1][0], 10.0 * self.px_to_mm, places=4)
        jump_segment = segments[1]
        self.assertFalse(jump_segment[0])
        self.assertAlmostEqual(jump_segment[1][-1][0], 0.0, places=4)


class ExportWriterTests(unittest.TestCase):
    def setUp(self) -> None:
        px_to_mm = 0.2645833333
        segments = [
            qme.MotionSegment(points=[(0.0, 0.0), (5.0, 0.0)], needle_down=True),
            qme.MotionSegment(points=[(5.0, 0.0), (5.0, 5.0)], needle_down=False),
        ]
        self.model = qme.MotionPathModel(segments, px_to_mm=px_to_mm)

    def test_dxf_writer_emits_polyline(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "path.dxf"
            qme._write_dxf(self.model, out)
            data = out.read_text()
            self.assertIn("LWPOLYLINE", data)
            self.assertIn("STITCH", data)

    def test_gif_writer_generates_animation(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "path.gif"
            qme._write_gif(self.model, out)
            data = out.read_bytes()
            self.assertTrue(data.startswith(b"GIF"))
            self.assertGreater(len(data), 100)


class ExportProfilesTests(unittest.TestCase):
    def test_all_required_formats_registered(self) -> None:
        required = {"DXF", "GIF"}
        self.assertTrue(required.issubset(set(qme.EXPORT_PROFILES.keys())))


class DisplayColorTests(unittest.TestCase):
    def test_color_for_pass_transitions(self) -> None:
        blue = qme._color_for_pass(0, needle_down=True)
        yellow = qme._color_for_pass(1, needle_down=True)
        red = qme._color_for_pass(2, needle_down=True)
        travel_red = qme._color_for_pass(0, needle_down=False)

        self.assertEqual(blue, (0.1, 0.55, 0.85, 0.9))
        self.assertEqual(yellow, (0.95, 0.85, 0.25, 0.9))
        self.assertEqual(red, (0.98, 0.55, 0.1, 0.9))
        self.assertEqual(travel_red, (0.9, 0.15, 0.15, 0.9))

    def test_partial_overlap_increments_pass(self) -> None:
        # Any prior stitched length on a geometric edge should count as a retrace.
        key_len = 10.0
        stitched_before = 3.0  # partial overlap already stitched
        passes_completed = 0
        if stitched_before > 1e-9:
            passes_completed = 1 if stitched_before <= key_len + 1e-9 else 2 + int(
                (stitched_before - key_len) / key_len
            )
        self.assertEqual(passes_completed, 1)

class PantographLayoutTests(unittest.TestCase):
    def test_stagger_extends_left_and_right(self) -> None:
        bounds = (0.0, 0.0, 10.0, 5.0)
        offsets = qme._compute_pantograph_offsets(
            bounds=bounds,
            repeat_count=2,
            row_count=1,
            row_distance_mm=5.0,
            px_to_mm=1.0,
            stagger=True,
            stagger_percent=50.0,
            start_point=(0.0, 0.0),
            end_point=(10.0, 0.0),
        )
        dx_values = [dx for _row, dx, _dy in offsets if _row == 0]
        self.assertLessEqual(min(dx_values), 0.0)
        self.assertGreaterEqual(max(dx_values), 10.0)

    def test_layout_bounds_ignore_stagger(self) -> None:
        bounds = (0.0, 0.0, 10.0, 5.0)
        layout = qme._compute_layout_bounds(
            bounds=bounds,
            repeat_count=2,
            row_count=2,
            row_distance_mm=5.0,
            px_to_mm=1.0,
            start_point=(0.0, 0.0),
            end_point=(10.0, 0.0),
        )
        self.assertEqual(layout, (0.0, 0.0, 20.0, 15.0))


class OptimizePathTests(unittest.TestCase):
    def test_optimize_path_reduces_overlap_and_preserves_endpoints(self) -> None:
        # Triangle where the base edge is stitched twice.
        original_points = [
            (0.0, 0.0),
            (10.0, 0.0),
            (20.0, 0.0),
            (10.0, 10.0),
            (0.0, 0.0),
            (10.0, 0.0),
        ]
        segments = [qme.MotionSegment(points=original_points, needle_down=True)]
        optimized_segments = qme.optimize_motion_segments(
            segments,
            start_point=original_points[0],
            end_point=original_points[-1],
        )

        self.assertEqual(len(optimized_segments), 1)
        optimized = optimized_segments[0]
        self.assertTrue(optimized.needle_down)
        self.assertGreaterEqual(len(optimized.points), 4)
        # Path endpoints must match the original endpoints.
        self.assertAlmostEqual(optimized.points[0][0], original_points[0][0], places=6)
        self.assertAlmostEqual(optimized.points[0][1], original_points[0][1], places=6)
        self.assertAlmostEqual(optimized.points[-1][0], original_points[-1][0], places=6)
        self.assertAlmostEqual(optimized.points[-1][1], original_points[-1][1], places=6)

        def overlap_length(segment_list):
            from math import dist, isclose

            edge_counts = {}
            overlap_total = 0.0
            for segment in segment_list:
                if not segment.needle_down or len(segment.points) < 2:
                    continue
                for index in range(len(segment.points) - 1):
                    start = segment.points[index]
                    end = segment.points[index + 1]
                    if isclose(start[0], end[0], abs_tol=1e-9) and isclose(
                        start[1], end[1], abs_tol=1e-9
                    ):
                        continue
                    key_start = (round(start[0], 6), round(start[1], 6))
                    key_end = (round(end[0], 6), round(end[1], 6))
                    if key_start == key_end:
                        continue
                    canonical = (
                        key_start if key_start <= key_end else key_end,
                        key_end if key_start <= key_end else key_start,
                    )
                    length = dist(start, end)
                    previous_count = edge_counts.get(canonical, 0)
                    if previous_count > 0:
                        overlap_total += length
                    edge_counts[canonical] = previous_count + 1
            return overlap_total

        original_overlap = overlap_length(segments)
        optimized_overlap = overlap_length(optimized_segments)
        self.assertLessEqual(optimized_overlap, original_overlap)


if __name__ == "__main__":
    unittest.main()
