import unittest
from unittest import mock

from deck_engine.models.manifest import Layout
from deck_engine.models.revision import RevisionState
from deck_engine.models.slide_spec import SlideSpec
from deck_engine.revision_engine import (
    MAX_ATTEMPTS,
    MAX_LOG_ENTRIES,
    apply_field_edit,
    classify_feedback,
    revise,
    summarize_entry,
)
from deck_engine.template_registry import load_tenant_assets


class TestClassifyFeedback(unittest.TestCase):
    def test_returns_field_edit_when_classified_as_such(self):
        with mock.patch("deck_engine.revision_engine.call_structured", return_value={"kind": "field_edit"}):
            self.assertEqual(classify_feedback("change 12% to 15%", "anthropic", "m"), "field_edit")

    def test_returns_regen_when_classified_as_such(self):
        with mock.patch("deck_engine.revision_engine.call_structured", return_value={"kind": "regen"}):
            self.assertEqual(classify_feedback("re-approach this slide", "anthropic", "m"), "regen")

    def test_falls_back_to_regen_on_unexpected_value(self):
        with mock.patch("deck_engine.revision_engine.call_structured", return_value={"kind": "something_else"}):
            self.assertEqual(classify_feedback("x", "anthropic", "m"), "regen")

    def test_falls_back_to_regen_on_call_failure(self):
        with mock.patch("deck_engine.revision_engine.call_structured", side_effect=RuntimeError("down")):
            self.assertEqual(classify_feedback("x", "anthropic", "m"), "regen")


class TestApplyFieldEdit(unittest.TestCase):
    def setUp(self):
        self.layout = load_tenant_assets("default").manifest.layout_by_id("text_bullets")
        self.spec = SlideSpec(
            layout_id="text_bullets",
            slots={"IM_BULLETS_TITLE": "Old Title", "IM_BULLETS_KICKER": "Old Kicker", "IM_BULLETS_BODY": ["a", "b"]},
        )

    def test_updates_only_the_targeted_slot(self):
        with mock.patch(
            "deck_engine.revision_engine.call_structured",
            return_value={"shape_name": "IM_BULLETS_TITLE", "new_value": "New Title"},
        ):
            updated, error = apply_field_edit(self.spec, self.layout, "change the title", "anthropic", "m")
        self.assertIsNone(error)
        self.assertEqual(updated.slots["IM_BULLETS_TITLE"], "New Title")
        self.assertEqual(updated.slots["IM_BULLETS_KICKER"], "Old Kicker")  # untouched
        self.assertEqual(updated.slots["IM_BULLETS_BODY"], ["a", "b"])  # untouched

    def test_rejects_unknown_shape_name(self):
        with mock.patch(
            "deck_engine.revision_engine.call_structured",
            return_value={"shape_name": "NOT_A_REAL_SLOT", "new_value": "x"},
        ):
            updated, error = apply_field_edit(self.spec, self.layout, "change something", "anthropic", "m")
        self.assertIsNone(updated)
        self.assertIsNotNone(error)

    def test_rejects_value_that_fails_validation(self):
        with mock.patch(
            "deck_engine.revision_engine.call_structured",
            return_value={"shape_name": "IM_BULLETS_TITLE", "new_value": "x" * 200},  # exceeds max_chars=60
        ):
            updated, error = apply_field_edit(self.spec, self.layout, "make it very long", "anthropic", "m")
        self.assertIsNone(updated)
        self.assertIsNotNone(error)

    def test_bullets_only_slot_is_never_offered_as_a_field_edit_target(self):
        # IM_BULLETS_BODY is a "bullets" slot, deliberately excluded from field-edit's schema (v1 narrowing)
        with mock.patch(
            "deck_engine.revision_engine.call_structured",
            return_value={"shape_name": "IM_BULLETS_BODY", "new_value": "not a list"},
        ):
            updated, error = apply_field_edit(self.spec, self.layout, "change a bullet", "anthropic", "m")
        self.assertIsNone(updated)
        self.assertIsNotNone(error)


class TestSummarizeEntry(unittest.TestCase):
    def test_summary_line_includes_kind_attempt_and_feedback(self):
        from deck_engine.models.revision import FeedbackLogEntry

        entry = FeedbackLogEntry(attempt=1, kind="regen", user_feedback="make it punchier")
        line = summarize_entry(entry)
        self.assertIn("regen", line)
        self.assertIn("#1", line)
        self.assertIn("make it punchier", line)


class TestRevise(unittest.TestCase):
    def setUp(self):
        self.layout = load_tenant_assets("default").manifest.layout_by_id("text_bullets")
        self.spec = SlideSpec(
            layout_id="text_bullets",
            slots={"IM_BULLETS_TITLE": "Old Title", "IM_BULLETS_KICKER": "Old Kicker", "IM_BULLETS_BODY": ["a", "b"]},
        )

    def _state(self):
        return RevisionState(label="slide 1", layout=self.layout, role="content slide", content_brief="brief", current_spec=self.spec)

    def test_field_edit_path_updates_only_target_slot_and_logs_it(self):
        with mock.patch("deck_engine.revision_engine.call_structured") as m:
            m.side_effect = [
                {"kind": "field_edit"},
                {"shape_name": "IM_BULLETS_TITLE", "new_value": "New Title"},
            ]
            state = revise(self._state(), "change the title", "anthropic", "m")
        self.assertEqual(state.current_spec.slots["IM_BULLETS_TITLE"], "New Title")
        self.assertEqual(state.current_spec.slots["IM_BULLETS_KICKER"], "Old Kicker")
        self.assertEqual(len(state.feedback_log), 1)
        self.assertEqual(state.feedback_log[0].kind, "field_edit")
        self.assertFalse(state.escalated)
        self.assertIsNone(state.last_error)

    def test_regen_path_calls_generate_single_slide_and_replaces_spec(self):
        new_spec = SlideSpec(layout_id="text_bullets", slots={"IM_BULLETS_TITLE": "Regenerated"})
        with mock.patch("deck_engine.revision_engine.call_structured", return_value={"kind": "regen"}), mock.patch(
            "deck_engine.revision_engine.generate_single_slide", return_value=(new_spec, None)
        ) as gen:
            state = revise(self._state(), "re-approach this slide", "anthropic", "m")
        gen.assert_called_once()
        self.assertEqual(state.current_spec, new_spec)
        self.assertEqual(state.feedback_log[0].kind, "regen")

    def test_field_edit_failure_falls_back_to_regen_automatically(self):
        new_spec = SlideSpec(layout_id="text_bullets", slots={"IM_BULLETS_TITLE": "Fallback Regen"})
        with mock.patch("deck_engine.revision_engine.call_structured") as m, mock.patch(
            "deck_engine.revision_engine.generate_single_slide", return_value=(new_spec, None)
        ) as gen:
            m.side_effect = [
                {"kind": "field_edit"},
                {"shape_name": "NOT_REAL", "new_value": "x"},  # apply_field_edit fails
            ]
            state = revise(self._state(), "change something ambiguous", "anthropic", "m")
        gen.assert_called_once()  # fell back to regen instead of failing outright
        self.assertEqual(state.current_spec, new_spec)
        self.assertEqual(state.feedback_log[0].kind, "regen")

    def test_escalates_after_max_attempts_of_genuine_failure(self):
        with mock.patch("deck_engine.revision_engine.call_structured", return_value={"kind": "regen"}), mock.patch(
            "deck_engine.revision_engine.generate_single_slide", return_value=(None, {"errors": ["nope"]})
        ) as gen:
            state = self._state()
            for _ in range(MAX_ATTEMPTS):
                state = revise(state, "impossible ask", "anthropic", "m")
            self.assertTrue(state.escalated)
            self.assertEqual(gen.call_count, MAX_ATTEMPTS)

            # a further call must NOT spend another LLM call
            state = revise(state, "one more try", "anthropic", "m")
            self.assertEqual(gen.call_count, MAX_ATTEMPTS)
            self.assertTrue(state.escalated)

    def test_successful_iteration_never_escalates_even_past_max_attempts(self):
        # repeated SUCCESSFUL revisions must not trip the failure-based escalation cap
        new_spec = SlideSpec(layout_id="text_bullets", slots={"IM_BULLETS_TITLE": "V"})
        with mock.patch("deck_engine.revision_engine.call_structured", return_value={"kind": "regen"}), mock.patch(
            "deck_engine.revision_engine.generate_single_slide", return_value=(new_spec, None)
        ) as gen:
            state = self._state()
            for i in range(MAX_ATTEMPTS + 2):
                state = revise(state, f"tweak {i}", "anthropic", "m")
            self.assertFalse(state.escalated)
            self.assertEqual(gen.call_count, MAX_ATTEMPTS + 2)

    def test_feedback_log_is_capped_and_oldest_entry_is_summarized(self):
        new_spec = SlideSpec(layout_id="text_bullets", slots={"IM_BULLETS_TITLE": "V"})
        with mock.patch("deck_engine.revision_engine.call_structured", return_value={"kind": "regen"}), mock.patch(
            "deck_engine.revision_engine.generate_single_slide", return_value=(new_spec, None)
        ):
            state = self._state()
            for i in range(MAX_LOG_ENTRIES + 2):
                state = revise(state, f"feedback {i}", "anthropic", "m")
            self.assertEqual(len(state.feedback_log), MAX_LOG_ENTRIES)
            self.assertEqual(state.feedback_log[0].kind, "summary")


if __name__ == "__main__":
    unittest.main()
