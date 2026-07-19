import unittest
from unittest import mock

from deck_engine.deck_planner import _coerce_list, plan_deck
from deck_engine.template_registry import load_tenant_assets


class TestCoerceList(unittest.TestCase):
    """Same double-encoding defense as content_generator._coerce_dict, applied
    to the Planner's array-typed 'slide_intents' field."""

    def test_passes_through_real_list(self):
        result, error = _coerce_list([{"a": 1}])
        self.assertEqual(result, [{"a": 1}])
        self.assertIsNone(error)

    def test_recovers_double_encoded_json_string(self):
        raw = '[{"purpose": "x"}]'
        result, error = _coerce_list(raw)
        self.assertEqual(result, [{"purpose": "x"}])
        self.assertIsNone(error)

    def test_reports_error_on_garbage_string(self):
        result, error = _coerce_list("not json at all [")
        self.assertEqual(result, [])
        self.assertIsNotNone(error)

    def test_reports_error_on_wrong_type(self):
        result, error = _coerce_list({"not": "a list"})
        self.assertEqual(result, [])
        self.assertIsNotNone(error)


VALID_INTENT = {
    "purpose": "Explain why margins are declining",
    "suggested_layout": "text_bullets",
    "headline": "Why Margins Are Declining",
    "content_outline": "Competition, input costs, discounting",
    "needs_exhibit": "none",
}


class TestPlanDeck(unittest.TestCase):
    def setUp(self):
        self.manifest = load_tenant_assets("default").manifest

    def test_successful_plan_first_attempt(self):
        with mock.patch(
            "deck_engine.deck_planner.call_structured",
            return_value={"slide_intents": [VALID_INTENT, VALID_INTENT]},
        ):
            intents, error = plan_deck("some brief", self.manifest, "anthropic", "claude-sonnet-5")
        self.assertIsNone(error)
        self.assertEqual(len(intents), 2)
        self.assertEqual(intents[0].headline, "Why Margins Are Declining")

    def test_double_encoded_response_is_recovered(self):
        import json

        with mock.patch(
            "deck_engine.deck_planner.call_structured",
            return_value={"slide_intents": json.dumps([VALID_INTENT])},
        ):
            intents, error = plan_deck("some brief", self.manifest, "anthropic", "claude-sonnet-5")
        self.assertIsNone(error)
        self.assertEqual(len(intents), 1)

    def test_retry_recovers_from_initial_garbage(self):
        responses = [
            {"slide_intents": "not valid json ["},
            {"slide_intents": [VALID_INTENT]},
        ]
        with mock.patch("deck_engine.deck_planner.call_structured", side_effect=responses):
            intents, error = plan_deck("some brief", self.manifest, "anthropic", "claude-sonnet-5")
        self.assertIsNone(error)
        self.assertEqual(len(intents), 1)

    def test_persistent_failure_is_flagged_not_silently_empty(self):
        with mock.patch(
            "deck_engine.deck_planner.call_structured",
            return_value={"slide_intents": "still not valid json ["},
        ):
            intents, error = plan_deck("some brief", self.manifest, "anthropic", "claude-sonnet-5")
        self.assertEqual(intents, [])
        self.assertIsNotNone(error)

    def test_malformed_individual_intent_is_dropped_not_fatal(self):
        bad_intent = {"purpose": "missing other required fields"}
        with mock.patch(
            "deck_engine.deck_planner.call_structured",
            return_value={"slide_intents": [bad_intent, VALID_INTENT]},
        ):
            intents, error = plan_deck("some brief", self.manifest, "anthropic", "claude-sonnet-5")
        self.assertIsNone(error)
        self.assertEqual(len(intents), 1)  # the bad one was dropped, the good one kept


if __name__ == "__main__":
    unittest.main()
