import unittest

from deck_engine.models.manifest import GeometryIn, Layout, Slot
from deck_engine.qa_checker import (
    check_overflow,
    compute_headroom_in,
    estimate_text_height_in,
    estimate_wrapped_line_count,
)


def slot(name, left_in, top_in, width_in, height_in, type_="text", required=False):
    return Slot(
        shape_name=name, type=type_, role="body", required=required,
        geometry_in=GeometryIn(left_in=left_in, top_in=top_in, width_in=width_in, height_in=height_in),
    )


class TestEstimateWrappedLineCount(unittest.TestCase):
    def test_empty_text_is_one_line(self):
        self.assertEqual(estimate_wrapped_line_count("", 20), 1)

    def test_text_shorter_than_chars_per_line_is_one_line(self):
        self.assertEqual(estimate_wrapped_line_count("short", 20), 1)

    def test_text_wraps_to_multiple_lines(self):
        self.assertEqual(estimate_wrapped_line_count("x" * 36, 36), 1)
        self.assertEqual(estimate_wrapped_line_count("x" * 37, 36), 2)
        self.assertEqual(estimate_wrapped_line_count("x" * 72, 36), 2)
        self.assertEqual(estimate_wrapped_line_count("x" * 73, 36), 3)


class TestEstimateTextHeightIn(unittest.TestCase):
    """font_size_pt=8, box_width_in=2.2 is chosen so the module's own
    constants (0.5 avg-char-width factor, 0.2in inset) produce a round
    chars_per_line=36 -- lets these pin an exact expected height rather than
    just checking directional trends."""

    def test_single_line_text_height(self):
        height = estimate_text_height_in("x" * 36, font_size_pt=8, box_width_in=2.2)
        self.assertAlmostEqual(height, (8 * 1.2) / 72, places=6)

    def test_wrapped_text_height_scales_with_line_count(self):
        one_line = estimate_text_height_in("x" * 36, font_size_pt=8, box_width_in=2.2)
        two_lines = estimate_text_height_in("x" * 37, font_size_pt=8, box_width_in=2.2)
        self.assertAlmostEqual(two_lines, one_line * 2, places=6)

    def test_bullet_list_sums_each_items_lines(self):
        items = ["x" * 36, "y" * 36, "z" * 36]  # 1 line each -> 3 lines total
        height = estimate_text_height_in(items, font_size_pt=8, box_width_in=2.2)
        line_height = (8 * 1.2) / 72
        self.assertAlmostEqual(height, 3 * line_height, places=6)


class TestComputeHeadroomIn(unittest.TestCase):
    def test_headroom_is_gap_to_nearest_shape_below(self):
        title = slot("TITLE", left_in=1.0, top_in=1.0, width_in=5.0, height_in=0.5)  # bottom = 1.5
        body = slot("BODY", left_in=1.0, top_in=2.0, width_in=5.0, height_in=1.0)  # top = 2.0
        layout = Layout(layout_id="x", slide_index=0, title="x", use_when="x", slots=[title, body])

        headroom = compute_headroom_in(title, layout, slide_height_in=7.5)
        self.assertAlmostEqual(headroom, 0.5, places=6)  # 2.0 - 1.5

    def test_headroom_falls_back_to_slide_bottom_when_nothing_below(self):
        body = slot("BODY", left_in=1.0, top_in=2.0, width_in=5.0, height_in=1.0)  # bottom = 3.0
        layout = Layout(layout_id="x", slide_index=0, title="x", use_when="x", slots=[body])

        headroom = compute_headroom_in(body, layout, slide_height_in=7.5, bottom_margin_in=0.3)
        self.assertAlmostEqual(headroom, 7.5 - 0.3 - 3.0, places=6)

    def test_non_overlapping_horizontal_neighbor_does_not_block_headroom(self):
        # two_column-style: body (left column) and image (right column) sit at
        # the same vertical band but never overlap horizontally -- the image
        # must not be treated as "below" the body.
        body = slot("BODY", left_in=0.9, top_in=2.35, width_in=5.7, height_in=4.1)
        image = slot("IMAGE", left_in=7.05, top_in=2.15, width_in=5.38, height_in=4.35, type_="image")
        layout = Layout(layout_id="two_column", slide_index=0, title="x", use_when="x", slots=[body, image])

        headroom = compute_headroom_in(body, layout, slide_height_in=7.5, bottom_margin_in=0.3)
        body_bottom = 2.35 + 4.1
        self.assertAlmostEqual(headroom, 7.5 - 0.3 - body_bottom, places=6)


class TestCheckOverflow(unittest.TestCase):
    def test_short_text_in_a_roomy_box_does_not_overflow(self):
        body = slot("BODY", left_in=1.0, top_in=1.0, width_in=10.0, height_in=3.0)
        layout = Layout(layout_id="x", slide_index=0, title="x", use_when="x", slots=[body])
        self.assertIsNone(check_overflow("Short line.", font_size_pt=18, slot=body, layout=layout, slide_height_in=7.5))

    def test_long_bullet_list_in_a_tight_box_overflows(self):
        title = slot("TITLE", left_in=1.0, top_in=1.0, width_in=5.0, height_in=0.3)
        body = slot("BODY", left_in=1.0, top_in=1.3, width_in=2.0, height_in=0.3)  # small box, tight headroom
        layout = Layout(layout_id="x", slide_index=0, title="x", use_when="x", slots=[title, body])
        items = ["a fairly long bullet point that will not fit in a tiny box"] * 6
        reason = check_overflow(items, font_size_pt=24, slot=body, layout=layout, slide_height_in=7.5)
        self.assertIsNotNone(reason)
        self.assertIn("overflow", reason)

    def test_missing_geometry_yields_no_verdict(self):
        no_geom = Slot(shape_name="X", type="text", role="body", geometry_in=None)
        layout = Layout(layout_id="x", slide_index=0, title="x", use_when="x", slots=[no_geom])
        self.assertIsNone(check_overflow("anything", font_size_pt=18, slot=no_geom, layout=layout, slide_height_in=7.5))

    def test_missing_font_size_yields_no_verdict(self):
        body = slot("BODY", left_in=1.0, top_in=1.0, width_in=1.0, height_in=0.1)
        layout = Layout(layout_id="x", slide_index=0, title="x", use_when="x", slots=[body])
        self.assertIsNone(check_overflow("x" * 500, font_size_pt=None, slot=body, layout=layout, slide_height_in=7.5))


if __name__ == "__main__":
    unittest.main()
