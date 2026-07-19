"""5.3 Deck Planner (reasoning). Turns a brief into an ordered list of
SlideIntents -- the deck skeleton (how many slides, each one's job and
headline) -- via one structured-output call (D6). Does NOT write final copy
(5.5's job) and does NOT make the final layout choice (5.4's job): each
intent's `suggested_layout` is only a hint (D5), always validated downstream.

Only plans the *content* slides -- title and closing_contact are fixed,
code-guaranteed bookends (see content_generator.py's module docstring), so
they're excluded from the layouts the Planner is even offered here. That
keeps "every deck opens/closes correctly" a structural guarantee rather than
something an LLM could get wrong.
"""
import json

from .layout_selector import CONTENT_LAYOUT_IDS, describe_layouts
from .llm_providers import call_structured
from .models.slide_intent import SlideIntent

SYSTEM_PROMPT_TEMPLATE = """You are the Deck Planner for a presentation deck engine (SDD component 5.3).
Given a plain-language brief, break it into an ordered list of {min_slides}-{max_slides}
slide intents -- the deck's skeleton, not final copy. Each intent is:
- purpose: one sentence on what this slide is for
- suggested_layout: your best guess at which layout fits (a HINT only -- it will
  be validated against the real template, not trusted blindly), one of: {layout_ids}
- headline: a working headline for the slide
- content_outline: a short outline of what this slide should say (the next stage
  writes the actual copy from this, so be concrete about the point/argument/data,
  not vague)
- needs_exhibit: "chart" if this slide's point is best made with a chart, "diagram"
  if it's a process/flow, otherwise "none"

Available layouts and when to use them:
{layout_desc}

Group related points onto one slide; don't split one idea across several, and
don't cram unrelated ideas onto one. If the brief is vague, still make a
reasonable editorial judgment call rather than refusing.
"""


def _schema(min_slides, max_slides):
    intent_schema = {
        "type": "object",
        "properties": {
            "purpose": {"type": "string"},
            "suggested_layout": {"type": "string", "enum": CONTENT_LAYOUT_IDS},
            "headline": {"type": "string"},
            "content_outline": {"type": "string"},
            "needs_exhibit": {"type": "string", "enum": ["none", "chart", "diagram"]},
        },
        "required": ["purpose", "suggested_layout", "headline", "content_outline", "needs_exhibit"],
        "additionalProperties": False,
    }
    return {
        "type": "object",
        "properties": {
            "slide_intents": {
                "type": "array",
                "minItems": min_slides,
                "maxItems": max_slides,
                "items": intent_schema,
            }
        },
        "required": ["slide_intents"],
    }


def _coerce_list(value):
    """Same double-encoding defense as content_generator._coerce_list -- models
    occasionally return an array-typed field as a JSON-encoded string instead
    of the real structure. Returns (list, error)."""
    if isinstance(value, list):
        return value, None
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except json.JSONDecodeError as e:
            return [], f"'slide_intents' was a string but not valid JSON: {e}"
        if isinstance(parsed, list):
            return parsed, None
        return [], f"'slide_intents' decoded to a {type(parsed).__name__}, expected a list"
    return [], f"'slide_intents' was type {type(value).__name__}, expected a list"


def plan_deck(brief: str, manifest, provider: str, model: str, min_slides: int = 2, max_slides: int = 8):
    """Returns (intents: list[SlideIntent], error: str | None). On failure
    (even after one bounded retry, D6), error is a diagnostic string and
    intents is []  -- callers must flag this, never silently produce a
    zero-slide deck and call it done (D9)."""
    system_prompt = SYSTEM_PROMPT_TEMPLATE.format(
        min_slides=min_slides, max_slides=max_slides,
        layout_ids=", ".join(CONTENT_LAYOUT_IDS), layout_desc=describe_layouts(manifest),
    )
    schema = _schema(min_slides, max_slides)

    def attempt(prompt):
        raw = call_structured(provider, model, system_prompt, prompt, schema, "emit_plan")
        raw_intents, coerce_error = _coerce_list(raw.get("slide_intents", []))
        if coerce_error:
            return None, coerce_error
        intents = []
        for d in raw_intents:
            if not isinstance(d, dict):
                continue
            try:
                intents.append(SlideIntent(**d))
            except Exception:
                continue  # a malformed individual intent just gets dropped, not fatal
        if not intents:
            return None, "no usable slide_intents in the response"
        return intents, None

    intents, error = attempt(brief)
    if intents:
        return intents, None

    retry_prompt = f"{brief}\n\nYour previous response was invalid: {error}\nReturn a fresh, correctly-structured plan."
    intents, error2 = attempt(retry_prompt)
    if intents:
        return intents, None
    return [], f"{error}; retry also failed: {error2}"
