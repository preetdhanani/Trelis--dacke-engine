import unittest
from unittest import mock

from deck_engine.layout_selector import CONTENT_LAYOUT_IDS, FALLBACK_LAYOUT_ID, describe_layouts, select_layout
from deck_engine.models.slide_intent import SlideIntent
from deck_engine.template_registry import load_tenant_assets


def make_intent(**overrides):
    defaults = dict(
        purpose="Make a point",
        suggested_layout="text_bullets",
        headline="A Headline",
        content_outline="Some outline",
        needs_exhibit="none",
    )
    defaults.update(overrides)
    return SlideIntent(**defaults)


class TestLayoutSelector(unittest.TestCase):
    def setUp(self):
        self.manifest = load_tenant_assets("default").manifest

    def test_describe_layouts_covers_every_content_layout(self):
        desc = describe_layouts(self.manifest)
        for layout_id in CONTENT_LAYOUT_IDS:
            self.assertIn(layout_id, desc)

    def test_valid_suggested_layout_is_used_directly_no_tiebreak(self):
        intent = make_intent(suggested_layout="quote_callout")
        with mock.patch("deck_engine.layout_selector.call_structured") as mock_call:
            layout, flag = select_layout(intent, self.manifest, "anthropic", "claude-sonnet-5")
        mock_call.assert_not_called()  # deterministic path never calls the LLM (D5)
        self.assertEqual(layout.layout_id, "quote_callout")
        self.assertIsNone(flag)

    def test_title_and_closing_are_never_selectable_even_if_suggested(self):
        # title/closing_contact are fixed bookends, not part of the Planner's
        # choices -- an invalid suggestion naming one of them must still fall
        # through to tiebreak/fallback, never select a bookend layout.
        intent = make_intent(suggested_layout="title")
        with mock.patch("deck_engine.layout_selector.call_structured", return_value={"layout_id": "text_bullets"}):
            layout, flag = select_layout(intent, self.manifest, "anthropic", "claude-sonnet-5")
        self.assertEqual(layout.layout_id, "text_bullets")
        self.assertIsNotNone(flag)

    def test_invalid_suggestion_triggers_tiebreak_and_uses_its_result(self):
        intent = make_intent(suggested_layout="not_a_real_layout")
        with mock.patch(
            "deck_engine.layout_selector.call_structured", return_value={"layout_id": "exhibit_data"}
        ) as mock_call:
            layout, flag = select_layout(intent, self.manifest, "anthropic", "claude-sonnet-5")
        mock_call.assert_called_once()
        self.assertEqual(layout.layout_id, "exhibit_data")
        self.assertIsNotNone(flag)
        self.assertIn("not_a_real_layout", flag)

    def test_invalid_suggestion_and_failed_tiebreak_falls_back_safely(self):
        intent = make_intent(suggested_layout="not_a_real_layout")
        with mock.patch("deck_engine.layout_selector.call_structured", side_effect=Exception("network down")):
            layout, flag = select_layout(intent, self.manifest, "anthropic", "claude-sonnet-5")
        self.assertEqual(layout.layout_id, FALLBACK_LAYOUT_ID)
        self.assertIsNotNone(flag)

    def test_tiebreak_returning_invalid_layout_falls_back_safely(self):
        intent = make_intent(suggested_layout=None)
        with mock.patch("deck_engine.layout_selector.call_structured", return_value={"layout_id": "also_not_real"}):
            layout, flag = select_layout(intent, self.manifest, "anthropic", "claude-sonnet-5")
        self.assertEqual(layout.layout_id, FALLBACK_LAYOUT_ID)
        self.assertIsNotNone(flag)


if __name__ == "__main__":
    unittest.main()
