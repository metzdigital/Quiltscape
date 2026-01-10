import tempfile
import unittest
from pathlib import Path

import sys

ROOT = Path(__file__).resolve().parents[1]
EXT_DIR = ROOT / "extensions"
if str(EXT_DIR) not in sys.path:
    sys.path.insert(0, str(EXT_DIR))

import quilt_motion_core as qmc  # noqa: E402


class MotionPathModelEdgeTests(unittest.TestCase):
    def test_edges_are_refined_for_long_segments(self) -> None:
        seg = qmc.MotionSegment(points=[(0.0, 0.0), (400.0, 0.0)], needle_down=True)
        model = qmc.MotionPathModel([seg], px_to_mm=1.0)
        self.assertGreater(len(model.edges), 1)
        self.assertAlmostEqual(model.total_length_mm, 400.0, places=6)

    def test_point_at_clamps_to_endpoints(self) -> None:
        seg = qmc.MotionSegment(points=[(0.0, 0.0), (10.0, 0.0)], needle_down=True)
        model = qmc.MotionPathModel([seg], px_to_mm=1.0)
        start, _ = model.point_at(-10.0)
        end, _ = model.point_at(999.0)
        self.assertEqual(start, (0.0, 0.0))
        self.assertEqual(end, (10.0, 0.0))


class PantographOffsetTests(unittest.TestCase):
    def test_offsets_respect_delta_y(self) -> None:
        bounds = (0.0, 0.0, 10.0, 5.0)
        offsets = qmc._compute_pantograph_offsets(
            bounds=bounds,
            repeat_count=3,
            row_count=1,
            row_distance_mm=5.0,
            px_to_mm=1.0,
            stagger=False,
            stagger_percent=0.0,
            start_point=(0.0, 0.0),
            end_point=(10.0, 2.0),
        )
        dy_values = [dy for _row, _dx, dy in offsets]
        self.assertGreater(max(dy_values), min(dy_values))

    def test_layout_bounds_include_row_spacing(self) -> None:
        bounds = (0.0, 0.0, 10.0, 5.0)
        layout = qmc._compute_layout_bounds(
            bounds=bounds,
            repeat_count=1,
            row_count=3,
            row_distance_mm=5.0,
            px_to_mm=1.0,
            start_point=(0.0, 0.0),
            end_point=(10.0, 0.0),
        )
        self.assertEqual(layout, (0.0, 0.0, 10.0, 25.0))


class ExportWriterEdgeTests(unittest.TestCase):
    def setUp(self) -> None:
        segments = [
            qmc.MotionSegment(points=[(0.0, 0.0), (5.0, 0.0)], needle_down=True),
            qmc.MotionSegment(points=[(5.0, 0.0), (5.0, 5.0)], needle_down=False),
        ]
        self.model = qmc.MotionPathModel(segments, px_to_mm=1.0)

    def test_qct_writer_has_expected_header(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "path.dxf"
            qmc._write_qct_dxf(self.model, out)
            data = out.read_text()
            self.assertIn("SECTION", data)
            self.assertTrue(data.strip().endswith("EOF"))

    def test_cartesian_coords_flip_y(self) -> None:
        model = qmc.MotionPathModel(
            [qmc.MotionSegment(points=[(0.0, 0.0), (0.0, 10.0)], needle_down=True)],
            px_to_mm=1.0,
            doc_height_px=100.0,
        )
        x, y = qmc._cartesian_coords(model, 0.0, 10.0)
        self.assertEqual(x, 0.0)
        self.assertEqual(y, 90.0)


class OptimizePathEdgeTests(unittest.TestCase):
    def test_optimize_preserves_travel_segments(self) -> None:
        stitched = qmc.MotionSegment(points=[(0.0, 0.0), (10.0, 0.0)], needle_down=True)
        travel = qmc.MotionSegment(points=[(10.0, 0.0), (15.0, 0.0)], needle_down=False)
        segments = [stitched, travel]
        optimized = qmc.optimize_motion_segments(segments, start_point=(0.0, 0.0), end_point=(10.0, 0.0))
        self.assertEqual(len(optimized), 2)
        self.assertFalse(optimized[1].needle_down)
