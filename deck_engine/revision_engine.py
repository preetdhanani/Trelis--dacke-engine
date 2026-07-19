"""5.10 Revision Engine (M6). Applies user feedback to a single slide without
disturbing any other slide (FR-6), routing to a surgical field edit or a
full-slide regeneration by classifying the feedback's intent (FR-7).

D10/D11 compliance by construction, not by convention: regeneration calls
`content_generator.generate_single_slide` -- the *exact* function used for
initial generation, not a second code path -- so a regenerated slide is
structurally guaranteed to inherit the seed's brand/theme the same way any
other slide does (D8). The caller is responsible for re-rendering the
returned SlideSpec through `renderer.replace_rendered_slide`, which itself
re-duplicates the pristine seed slide fresh (D13) rather than touching the
already-rendered output slide.

Deliberate v1 narrowing: field-edit only targets **text**-type slots. A
"small ask" like "change 12% to 15%" is a targeted text substitution; a
bullets-list ask ("add a point about X", "reorder these") is a structural
decision about which item / how many, not a value fill-in, so it's routed to
a full regen instead of a fabricated bullets-splitting scheme. If the
classifier says field_edit but no text slot matches, `revise()` falls back to
regen automatically for that attempt (never a wasted/failed turn).

Escalation (5.10): after MAX_ATTEMPTS *failed* revise() calls on one slide
(calls that produced no usable spec -- not merely calls in general, so a user
iterating successfully never gets blocked), further calls are refused (no
LLM call spent) and the slide is marked escalated -- an ambiguous or
out-of-scope ask, not an infinite retry. The feedback log is capped at
MAX_LOG_ENTRIES; the oldest entry beyond that is folded into a single summary
line rather than dropped (5.10: "summarize older attempts").
"""
import json

from .content_generator import generate_single_slide
from .llm_providers import call_structured
from .models.revision import FeedbackLogEntry, RevisionState
from .models.slide_spec import SlideSpec, validate_against_layout

MAX_ATTEMPTS = 3
MAX_LOG_ENTRIES = 4

_CLASSIFY_SCHEMA = {
    "type": "object",
    "properties": {
        "kind": {
            "type": "string",
            "enum": ["field_edit", "regen"],
            "description": (
                "field_edit: a small, targeted change to one existing value/word/number "
                "(e.g. 'change 12% to 15%', 'fix the typo in the title'). "
                "regen: a structural ask to re-approach, restructure, or substantially rewrite the slide "
                "(e.g. 're-approach this slide', 'make this more concise', 'add a point about X')."
            ),
        }
    },
    "required": ["kind"],
    "additionalProperties": False,
}

_CLASSIFY_SYSTEM_PROMPT = (
    "You are the feedback classifier for a presentation deck engine's Revision Engine (SDD 5.10). "
    "Classify the user's feedback about one slide as exactly one of the schema's two kinds."
)


def classify_feedback(feedback_text, provider, model):
    """Returns "field_edit" or "regen". Falls back to "regen" on any
    classification failure -- regen always goes through the exact same
    validated fill path as initial generation (D10), so an ambiguous
    classification is safer to default toward it than to risk a field edit
    landing on the wrong slot."""
    try:
        raw = call_structured(provider, model, _CLASSIFY_SYSTEM_PROMPT, feedback_text, _CLASSIFY_SCHEMA, "classify_feedback")
        kind = raw.get("kind") if isinstance(raw, dict) else None
        return kind if kind in ("field_edit", "regen") else "regen"
    except Exception:
        return "regen"


def _field_edit_schema(layout):
    text_slot_names = [s.shape_name for s in layout.non_image_slots() if s.type == "text"]
    return {
        "type": "object",
        "properties": {
            "shape_name": {"type": "string", "enum": text_slot_names},
            "new_value": {"type": "string"},
        },
        "required": ["shape_name", "new_value"],
        "additionalProperties": False,
    }, text_slot_names


_FIELD_EDIT_SYSTEM_TEMPLATE = """You are the field-edit agent for a presentation deck engine's Revision Engine (SDD 5.10).
Given the slide's current text-slot values and a small, targeted feedback ask, identify
which ONE slot the feedback is about and produce that slot's corrected value only.
Respect max_chars if given -- it is a hard limit, not a suggestion.

Current values:
{current_values}
"""


def apply_field_edit(spec: SlideSpec, layout, feedback_text, provider, model):
    """Returns (updated_spec, error). On success, `updated_spec` is byte-
    identical to `spec` except for the one targeted slot's value (FR-6: "only
    the flagged slide changes"; here, narrower still -- only the one field).
    `error` is set (and updated_spec is None) if no text slot exists to
    target, the model names an unknown slot, or the new value fails
    business-rule validation -- the caller should fall back to a full regen
    in that case rather than treat it as fatal."""
    schema, text_slot_names = _field_edit_schema(layout)
    if not text_slot_names:
        return None, "layout has no text slots to field-edit"

    current_values = "\n".join(f"- {name}: {spec.slots.get(name)!r}" for name in text_slot_names)
    system_prompt = _FIELD_EDIT_SYSTEM_TEMPLATE.format(current_values=current_values)

    try:
        raw = call_structured(provider, model, system_prompt, feedback_text, schema, "apply_field_edit")
    except Exception as e:
        return None, f"field-edit call failed: {e}"

    shape_name = raw.get("shape_name") if isinstance(raw, dict) else None
    new_value = raw.get("new_value") if isinstance(raw, dict) else None
    if shape_name not in text_slot_names or not isinstance(new_value, str):
        return None, f"field-edit response did not name a valid text slot: {raw!r}"

    updated_slots = dict(spec.slots)
    updated_slots[shape_name] = new_value
    updated_spec = SlideSpec(layout_id=spec.layout_id, content_kind=spec.content_kind, slots=updated_slots, exhibit=spec.exhibit)
    errors = validate_against_layout(updated_spec, layout)
    if errors:
        return None, f"field edit produced an invalid slide: {errors}"
    return updated_spec, None


def summarize_entry(entry: FeedbackLogEntry) -> str:
    """Folds one old log entry into a single compact line (5.10: "summarize
    older attempts into a single line")."""
    return f"[{entry.kind} #{entry.attempt}] {entry.user_feedback[:80]}"


def _append_log_entry(state: RevisionState, entry: FeedbackLogEntry):
    state.feedback_log.append(entry)
    if len(state.feedback_log) > MAX_LOG_ENTRIES:
        oldest = state.feedback_log[0]
        summary = FeedbackLogEntry(attempt=oldest.attempt, kind="summary", user_feedback=summarize_entry(oldest))
        state.feedback_log = [summary] + state.feedback_log[2:]


def regenerate_slide(content_brief, layout, feedback_text, feedback_log, deck_context, provider, model, role):
    """Full-slide regeneration (D10): reuses `generate_single_slide` --
    the exact same generation/fill path as initial content, not a second code
    path -- with an augmented brief folding in the prior feedback log (5.10:
    accumulating, capped context) and the new feedback."""
    log_lines = "\n".join(f"- {e.user_feedback if e.kind == 'summary' else e.user_feedback}" for e in feedback_log)
    augmented_brief = content_brief
    if log_lines:
        augmented_brief += f"\n\nPrior revision feedback on this slide (already addressed once each, still unresolved):\n{log_lines}"
    augmented_brief += f"\n\nNew feedback to address now: {feedback_text}"
    if deck_context:
        augmented_brief += f"\n\nDeck context (for consistency with the rest of the deck): {deck_context}"
    return generate_single_slide(augmented_brief, layout, provider, model, role=role)


def revise(state: RevisionState, feedback_text, provider, model, deck_context=""):
    """The main entry point (SDD 5.10's `revise(slide_id, feedback_text)`):
    classify the feedback, apply a field edit or a full regen, log the
    attempt (capped/summarized), and escalate after MAX_ATTEMPTS. Mutates and
    returns `state`; once `state.escalated` is True, further calls are a
    no-op (never an infinite retry loop, D9's "stop, don't guess" discipline
    applied to revision)."""
    if state.escalated:
        return state

    kind = classify_feedback(feedback_text, provider, model)
    new_spec, error = None, None

    if kind == "field_edit":
        new_spec, error = apply_field_edit(state.current_spec, state.layout, feedback_text, provider, model)
        if new_spec is None:
            kind = "regen"  # fall back rather than waste the attempt

    if kind == "regen":
        new_spec, error = regenerate_slide(
            state.content_brief, state.layout, feedback_text, state.feedback_log, deck_context, provider, model, state.role
        )

    state.attempts += 1
    snapshot = new_spec.slots if new_spec is not None else None
    _append_log_entry(state, FeedbackLogEntry(attempt=state.attempts, kind=kind, user_feedback=feedback_text, spec_snapshot=snapshot))

    if new_spec is not None:
        state.current_spec = new_spec
        state.last_error = None
    else:
        state.failed_attempts += 1
        state.last_error = str(error) if error is not None else "revision produced no usable slide"
        if state.failed_attempts >= MAX_ATTEMPTS:
            state.escalated = True
    return state
