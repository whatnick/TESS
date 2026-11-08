import tempfile
import unittest
from pathlib import Path

from scripts.prepare_elecrow import (
    fitted_positions, odb_drills, original_drills, read_bom, read_positions, write_drills, write_outline,
)


class PackagingTests(unittest.TestCase):
    def test_default_bom_filters_nonparts_and_unlisted_headers(self):
        bom = read_bom()
        fitted, excluded = fitted_positions(bom, read_positions())
        self.assertEqual({r["Designator"] for r in fitted}, {r for g in bom for r in g["refs"]})
        self.assertTrue({"P2", "P3", "P6", "P7"} <= {r["Designator"] for r in excluded})

    def test_missing_fitted_placement_is_an_error(self):
        with self.assertRaisesRegex(ValueError, "Missing placements"):
            fitted_positions(read_bom(), [])

    def test_unknown_extra_placement_is_an_error(self):
        with self.assertRaisesRegex(ValueError, "Unclassified"):
            fitted_positions(read_bom(), read_positions() + [{"Designator": "U999"}])

    def test_odb_drills_are_explicit_metric_without_modal_coordinates(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "features"
            source.write_text("UNITS=INCH\n$0 r25.5906\nP 0.5 1.0 0 P 0 0\n")
            target = Path(directory) / "NPTH.drl"
            write_drills(target, odb_drills(source), "NON_PLATED")
            self.assertIn("X12.700000Y25.400000", target.read_text())
            self.assertIn(";TYPE=NON_PLATED", target.read_text())

    def test_analytical_outline_preserves_arcs_and_closes(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "outline.gbr"
            result = write_outline(target)
            self.assertEqual(result["arcs"], 4)
            self.assertLessEqual(result["max_endpoint_join_normalisation_mm"], 0.0001)
            self.assertIn("G03", target.read_text())

    def test_altium_trailing_zero_omission_and_modal_coordinates(self):
        points = original_drills(
            "M48\n;FILE_FORMAT=2:5\nINCH,LZ\nT01F00S00C0.03150\n"
            "%\nT01\nX009Y0025\nY0035\nM30\n"
        )
        self.assertAlmostEqual(points[0][1], 22.86)
        self.assertAlmostEqual(points[0][2], 6.35)
        self.assertAlmostEqual(points[1][1], 22.86)
        self.assertAlmostEqual(points[1][2], 8.89)


if __name__ == "__main__":
    unittest.main()
