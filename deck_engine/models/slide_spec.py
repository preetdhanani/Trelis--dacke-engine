"""SlideSpec (SDD SS6.2) -- the contract between reasoning and rendering. Keys
in `slots` are the manifest's exact shape_name values, so the Renderer can
address shapes directly with no name-translation layer.
"""
from typing import Dict, List, Optional, Union

from pydantic import BaseModel, Field

from .manifest import Layout

SlotValue = Union[str, List[str], None]


class SlideSpec(BaseModel):
    layout_id: str
    content_kind: str = "text"  # "text" | "chart" | "diagram" -- only "text" is exercised in M1
    slots: Dict[str, SlotValue] = Field(default_factory=dict)
    exhibit: Optional[dict] = None


def validate_against_layout(spec: SlideSpec, layout: Layout) -> List[str]:
    """D6's business-rule validation layer: once layout_id is known, check that
    *specific* layout's required slots, max_chars, max_items,
    max_chars_per_item, and allowed_values. Returns violation strings; empty
    means valid.
    """
    if spec.layout_id != layout.layout_id:
        return [f"spec.layout_id={spec.layout_id!r} does not match layout {layout.layout_id!r}"]

    errors: List[str] = []
    slot_defs = {s.shape_name: s for s in layout.non_image_slots()}

    for key in spec.slots:
        if key not in slot_defs:
            errors.append(f"unknown slot '{key}' is not in layout '{layout.layout_id}'")

    for name, slot in slot_defs.items():
        val = spec.slots.get(name)
        if slot.required and (val is None or val == "" or val == []):
            errors.append(f"missing required slot '{name}'")
            continue
        if val is None:
            continue

        if slot.type == "text":
            if not isinstance(val, str):
                errors.append(f"slot '{name}' should be a string, got {type(val).__name__}")
                continue
            if slot.max_chars and len(val) > slot.max_chars:
                errors.append(f"slot '{name}' exceeds max_chars={slot.max_chars} (got {len(val)})")
            if slot.allowed_values and val not in slot.allowed_values:
                errors.append(f"slot '{name}' value {val!r} not in allowed_values {slot.allowed_values}")

        elif slot.type == "bullets":
            if not isinstance(val, list) or not all(isinstance(v, str) for v in val):
                errors.append(f"slot '{name}' should be a list of strings, got {val!r}")
                continue
            if slot.max_items and len(val) > slot.max_items:
                errors.append(f"slot '{name}' has {len(val)} items, max_items={slot.max_items}")
            if slot.max_chars_per_item:
                for i, item in enumerate(val):
                    if len(item) > slot.max_chars_per_item:
                        errors.append(
                            f"slot '{name}' item[{i}] exceeds max_chars_per_item={slot.max_chars_per_item} (got {len(item)})"
                        )

    return errors
