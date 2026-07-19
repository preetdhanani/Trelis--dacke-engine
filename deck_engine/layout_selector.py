"""5.4 Layout Selector. Maps each SlideIntent to a concrete manifest layout --
deterministic first: validate the Planner's `suggested_layout` hint against
the real manifest (D5 -- never trust an LLM-produced string blindly, even one
constrained to an enum, without checking it against ground truth). Only when
that hint is missing/invalid does an optional cheap LLM tiebreak run, reusing
whichever provider/model the rest of the pipeline is already using. If even
that doesn't produce a valid layout_id, fall back to the safe general layout
(text_bullets) and flag it (SDD 5.4's stated failure mode) -- never raise,
never leave a slide without a layout.

Owns CONTENT_LAYOUT_IDS -- the layouts a SlideIntent may ever be mapped to
(everything except the two fixed bookends, title/closing_contact -- see
content_generator.py's module docstring for why those stay fixed).
"""
from .llm_providers import call_structured

CONTENT_LAYOUT_IDS = ["section_divider", "text_bullets", "image_only", "two_column", "exhibit_data", "quote_callout"]

FALLBACK_LAYOUT_ID = "text_bullets"


def describe_layouts(manifest, layout_ids=CONTENT_LAYOUT_IDS):
    lines = []
    for layout_id in layout_ids:
        layout = manifest.layout_by_id(layout_id)
        lines.append(f"- {layout_id}: {layout.use_when}")
    return "\n".join(lines)


def _tiebreak(intent, manifest, provider, model):
    """A cheap, single structured-output call to pick a layout_id when the
    Planner's hint didn't survive validation. Returns a layout_id string or
    None on any failure (network, malformed response, etc.) -- the caller
    always has select_layout's deterministic fallback behind this."""
    schema = {
        "type": "object",
        "properties": {"layout_id": {"type": "string", "enum": CONTENT_LAYOUT_IDS}},
        "required": ["layout_id"],
        "additionalProperties": False,
    }
    system_prompt = (
        "You are the Layout Selector for a presentation deck engine (SDD component 5.4). "
        "Given a slide's purpose, headline, and content outline, pick exactly one best-fitting "
        "layout_id from the list below.\n\n" + describe_layouts(manifest)
    )
    user_prompt = (
        f"purpose: {intent.purpose}\nheadline: {intent.headline}\n"
        f"content_outline: {intent.content_outline}\nneeds_exhibit: {intent.needs_exhibit}"
    )
    try:
        raw = call_structured(provider, model, system_prompt, user_prompt, schema, "emit_layout_choice")
        layout_id = raw.get("layout_id")
        return layout_id if isinstance(layout_id, str) else None
    except Exception:
        return None


def select_layout(intent, manifest, provider, model):
    """Returns (layout, flag_reason: str | None). flag_reason is None when the
    Planner's own suggested_layout hint validated cleanly; otherwise it
    explains what went wrong and how it was resolved, for the caller to log
    (not necessarily a hard needs_review -- a resolved tiebreak is a
    successful selection, just one worth being visible about)."""
    valid_ids = {l.layout_id for l in manifest.layouts if l.layout_id in CONTENT_LAYOUT_IDS}

    if intent.suggested_layout in valid_ids:
        return manifest.layout_by_id(intent.suggested_layout), None

    tiebreak_id = _tiebreak(intent, manifest, provider, model)
    if tiebreak_id in valid_ids:
        return (
            manifest.layout_by_id(tiebreak_id),
            f"suggested_layout {intent.suggested_layout!r} was not a valid content layout; tiebreak selected {tiebreak_id!r}",
        )

    return (
        manifest.layout_by_id(FALLBACK_LAYOUT_ID),
        f"suggested_layout {intent.suggested_layout!r} was invalid and tiebreak did not resolve it; defaulted to {FALLBACK_LAYOUT_ID!r}",
    )
