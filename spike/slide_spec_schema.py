"""S2 -- minimal SlideSpec schema (SDD SS6.2) + manifest-driven business-rule
validation, split in two layers per D6:

  1. ENVELOPE_SCHEMA: the coarse JSON schema forced on the model at the API
     level (enum'd layout_id, generic slots dict). This is the "~99.8% schema
     valid" layer -- gets the gross shape right, can't express per-layout
     max_chars/max_items/allowed_values since those vary by which layout_id
     the model picked.
  2. validate_against_manifest(): the Pydantic/business-rule layer -- once we
     know which layout_id came back, check its *specific* slot set, required-
     ness, max_chars, max_items, max_chars_per_item, allowed_values.
"""
import json
from pathlib import Path

MANIFEST_PATH = Path(__file__).parent.parent / "Templates" / "MasterDeck_manifest.json"
MANIFEST = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
LAYOUTS_BY_ID = {l["layout_id"]: l for l in MANIFEST["layouts"]}
LAYOUT_IDS = list(LAYOUTS_BY_ID.keys())

ENVELOPE_SCHEMA = {
    "type": "object",
    "properties": {
        "layout_id": {
            "type": "string",
            "enum": LAYOUT_IDS,
            "description": "Which seed-slide layout best fits this brief.",
        },
        "content_kind": {
            "type": "string",
            "enum": ["text", "chart", "diagram"],
        },
        "slots": {
            "type": "object",
            "description": (
                "Keys are the shape_name values of the chosen layout's non-image slots. "
                "Text slots are a single string; bullets slots are an array of strings. "
                "Never include image-typed slots here."
            ),
            "additionalProperties": True,
        },
    },
    "required": ["layout_id", "content_kind", "slots"],
}


def slot_defs_for(layout_id):
    layout = LAYOUTS_BY_ID[layout_id]
    return {s["shape_name"]: s for s in layout["slots"] if s["type"] != "image"}


def validate_against_manifest(layout_id, slots):
    """Returns list[str] of business-rule violations (empty == valid)."""
    errors = []
    if layout_id not in LAYOUTS_BY_ID:
        return [f"unknown layout_id '{layout_id}'"]
    if not isinstance(slots, dict):
        return [f"'slots' must be an object, got {type(slots).__name__}"]

    slot_defs = slot_defs_for(layout_id)

    for key in slots:
        if key not in slot_defs:
            errors.append(f"unknown slot '{key}' is not in layout '{layout_id}'")

    for name, slot in slot_defs.items():
        val = slots.get(name)
        if slot.get("required") and (val is None or val == "" or val == []):
            errors.append(f"missing required slot '{name}'")
            continue
        if val is None:
            continue

        if slot["type"] == "text":
            if not isinstance(val, str):
                errors.append(f"slot '{name}' should be a string, got {type(val).__name__}")
                continue
            max_chars = slot.get("max_chars")
            if max_chars and len(val) > max_chars:
                errors.append(f"slot '{name}' exceeds max_chars={max_chars} (got {len(val)})")
            allowed = slot.get("allowed_values")
            if allowed and val not in allowed:
                errors.append(f"slot '{name}' value {val!r} not in allowed_values {allowed}")

        elif slot["type"] == "bullets":
            if not isinstance(val, list) or not all(isinstance(v, str) for v in val):
                errors.append(f"slot '{name}' should be a list of strings, got {val!r}")
                continue
            max_items = slot.get("max_items")
            if max_items and len(val) > max_items:
                errors.append(f"slot '{name}' has {len(val)} items, max_items={max_items}")
            max_cpi = slot.get("max_chars_per_item")
            if max_cpi:
                for i, item in enumerate(val):
                    if len(item) > max_cpi:
                        errors.append(f"slot '{name}' item[{i}] exceeds max_chars_per_item={max_cpi} (got {len(item)})")

    return errors
