"""Pydantic models for the Template Manifest (SDD SS6.3) -- the machine-readable
contract describing each seed slide's shapes. Field names mirror the real
MasterDeck_manifest.json exactly, since slots are addressed by the
literal shape_name (D3), not by a translated role name.
"""
from typing import Any, Dict, List, Optional

from pydantic import BaseModel


class GeometryIn(BaseModel):
    left_in: float
    top_in: float
    width_in: float
    height_in: float


class Slot(BaseModel):
    shape_name: str
    type: str  # "text" | "bullets" | "image"
    role: str
    required: bool = False
    max_chars: Optional[int] = None
    max_items: Optional[int] = None
    max_chars_per_item: Optional[int] = None
    allowed_values: Optional[List[str]] = None
    geometry_in: Optional[GeometryIn] = None
    fit: Optional[str] = None
    note: Optional[str] = None


class Layout(BaseModel):
    layout_id: str
    slide_index: int
    title: str
    use_when: str
    slots: List[Slot]

    def slot(self, shape_name: str) -> Optional[Slot]:
        return next((s for s in self.slots if s.shape_name == shape_name), None)

    def non_image_slots(self) -> List[Slot]:
        return [s for s in self.slots if s.type != "image"]

    def image_slots(self) -> List[Slot]:
        return [s for s in self.slots if s.type == "image"]


class Manifest(BaseModel):
    template_name: str
    template_version: str
    pipeline: str
    fill_protocol: List[str]
    brand: Dict[str, Any]
    slide_dimensions_in: Dict[str, float]
    shape_name_convention: str
    layouts: List[Layout]

    def layout_by_id(self, layout_id: str) -> Layout:
        for layout in self.layouts:
            if layout.layout_id == layout_id:
                return layout
        raise KeyError(f"unknown layout_id '{layout_id}'; known: {[l.layout_id for l in self.layouts]}")
