import unittest

from deck_engine.chart_builder import (
    build_chart_spec,
    classify_data_shape,
    nearest_supported_type,
    resolve_chart_type,
    suggest_chart_type,
)
from deck_engine.models.chart import SUPPORTED_CHART_TYPES, ChartData


def data(categories, series, **kw):
    return ChartData(categories=categories, series=series, **kw)


class TestDataShapeClassification(unittest.TestCase):
    def test_multiple_series_is_comparison(self):
        d = data(["North", "South", "East"], {"2023": [1, 2, 3], "2024": [4, 5, 6]})
        self.assertEqual(classify_data_shape(d), "multi_series_comparison")

    def test_single_series_over_quarters_is_time_series(self):
        d = data(["Q1", "Q2", "Q3", "Q4"], {"Revenue": [2.1, 2.5, 2.9, 3.4]})
        self.assertEqual(classify_data_shape(d), "time_series")

    def test_single_series_over_years_is_time_series(self):
        d = data(["2021", "2022", "2023", "2024"], {"Users": [10, 20, 35, 60]})
        self.assertEqual(classify_data_shape(d), "time_series")

    def test_single_series_over_months_is_time_series(self):
        d = data(["Jan", "Feb", "Mar"], {"Signups": [5, 8, 13]})
        self.assertEqual(classify_data_shape(d), "time_series")

    def test_single_series_marked_share_is_part_of_whole(self):
        d = data(["Cloud", "On-prem", "Hybrid"], {"Share": [50, 30, 20]}, kind="share")
        self.assertEqual(classify_data_shape(d), "part_of_whole")

    def test_single_series_plain_categories_is_categorical(self):
        d = data(["Alpha", "Beta", "Gamma"], {"Score": [7, 4, 9]})
        self.assertEqual(classify_data_shape(d), "categorical")


class TestSuggestion(unittest.TestCase):
    def test_time_series_suggests_line(self):
        self.assertEqual(suggest_chart_type(data(["Q1", "Q2"], {"r": [1, 2]})).chart_type, "line")

    def test_comparison_suggests_column(self):
        self.assertEqual(
            suggest_chart_type(data(["A", "B"], {"x": [1, 2], "y": [3, 4]})).chart_type, "column"
        )

    def test_share_suggests_pie(self):
        self.assertEqual(
            suggest_chart_type(data(["A", "B"], {"s": [60, 40]}, kind="share")).chart_type, "pie"
        )

    def test_every_suggestion_and_alternative_is_supported(self):
        cases = [
            data(["Q1", "Q2"], {"r": [1, 2]}),
            data(["A", "B"], {"x": [1, 2], "y": [3, 4]}),
            data(["A", "B"], {"s": [60, 40]}, kind="share"),
            data(["A", "B", "C"], {"v": [1, 2, 3]}),
        ]
        for d in cases:
            s = suggest_chart_type(d)
            self.assertIn(s.chart_type, SUPPORTED_CHART_TYPES)
            for alt in s.alternatives:
                self.assertIn(alt, SUPPORTED_CHART_TYPES)
            self.assertTrue(s.reason)  # a one-line reason is always present (5.6)


class TestResolveAndOverride(unittest.TestCase):
    def test_no_request_uses_suggestion(self):
        d = data(["Q1", "Q2"], {"r": [1, 2]})
        chart_type, note = resolve_chart_type(d, None)
        self.assertEqual(chart_type, "line")
        self.assertIsNone(note)

    def test_supported_request_is_honored_without_note(self):
        d = data(["Q1", "Q2"], {"r": [1, 2]})
        chart_type, note = resolve_chart_type(d, "bar")
        self.assertEqual(chart_type, "bar")
        self.assertIsNone(note)

    def test_supported_request_is_case_insensitive(self):
        chart_type, note = resolve_chart_type(data(["A", "B"], {"x": [1, 2]}), "PIE")
        self.assertEqual(chart_type, "pie")

    def test_unsupported_request_maps_to_nearest_with_note(self):
        d = data(["A", "B"], {"x": [1, 2]})
        chart_type, note = resolve_chart_type(d, "waterfall")
        self.assertEqual(chart_type, "column")  # D7 exotic -> nearest supported
        self.assertIsNotNone(note)
        self.assertIn("waterfall", note)

    def test_unknown_exotic_falls_back_to_default(self):
        supported, note = nearest_supported_type("some_made_up_chart")
        self.assertIn(supported, SUPPORTED_CHART_TYPES)
        self.assertIsNotNone(note)

    def test_build_chart_spec_carries_note_on_unsupported(self):
        spec = build_chart_spec(data(["A", "B"], {"x": [1, 2]}), "funnel")
        self.assertIn(spec.chart_type, SUPPORTED_CHART_TYPES)
        self.assertIsNotNone(spec.note)

    def test_per_chart_type_key_is_an_override(self):
        # a `type` on the data itself is the same override path the CLI uses
        d = data(["Q1", "Q2"], {"r": [1, 2]}, type="area")
        spec = build_chart_spec(d, d.type)
        self.assertEqual(spec.chart_type, "area")


class TestChartDataValidation(unittest.TestCase):
    def test_empty_categories_rejected(self):
        with self.assertRaises(Exception):
            ChartData(categories=[], series={"x": []})

    def test_empty_series_rejected(self):
        with self.assertRaises(Exception):
            ChartData(categories=["A"], series={})

    def test_series_length_must_match_categories(self):
        with self.assertRaises(Exception):
            ChartData(categories=["A", "B", "C"], series={"x": [1, 2]})

    def test_missing_data_is_never_fabricated(self):
        # 5.6: missing data -> caller must supply it; the model refuses to
        # construct a chart from nothing rather than inventing values.
        with self.assertRaises(Exception):
            ChartData(categories=["A", "B"], series={"x": []})


if __name__ == "__main__":
    unittest.main()
