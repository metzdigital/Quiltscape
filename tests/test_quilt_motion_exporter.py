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

    def test_text_writer_contains_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "path.txt"
            qme._write_txt(self.model, out)
            lines = out.read_text().splitlines()
            self.assertTrue(lines)
            first = lines[0].split()
            self.assertEqual(len(first), 2)
            self.assertFalse("," in lines[0])

    def test_dxf_writer_emits_polyline(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "path.dxf"
            qme._write_dxf(self.model, out)
            data = out.read_text()
            self.assertIn("LWPOLYLINE", data)
            self.assertIn("STITCH", data)

    def test_generic_pointset_writer_includes_sections(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "path.bqm"
            qme._write_generic_pointset(self.model, out, "BQM")
            data = out.read_text().splitlines()
            self.assertEqual(data[1], "VERSION 1")
            self.assertIn("BEGIN STITCH", data)
            self.assertIn("BEGIN JUMP", data)

    def test_hqf_writer_generates_header(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "path.hqf"
            qme._write_hqf(self.model, out)
            lines = out.read_text().splitlines()
            self.assertTrue(lines[0].startswith("#"))
            self.assertEqual(lines[1], "VERSION 1")
            self.assertIn("BEGIN STITCH", lines)


class ExportProfilesTests(unittest.TestCase):
    def test_all_required_formats_registered(self) -> None:
        required = {"BQM", "DXF", "HQF", "IQP", "PAT", "PLT", "QCC", "QLI", "SSD", "TXT"}
        self.assertTrue(required.issubset(set(qme.EXPORT_PROFILES.keys())))


if __name__ == "__main__":
    unittest.main()
