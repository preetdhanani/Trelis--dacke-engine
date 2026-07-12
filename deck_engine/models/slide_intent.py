"""SlideIntent -- the Deck Planner's (5.3) output: the deck skeleton, not
final copy. One entry per intended slide, for everything *between* the fixed
title/closing bookends (see content_generator.py's module docstring for why
those two stay fixed rather than planner-decided).
"""
from typing import Literal, Optional

from pydantic import BaseModel


class SlideIntent(BaseModel):
    purpose: str  # what this slide is for, e.g. "Argue why now is the right time to enter DTC"
    suggested_layout: Optional[str] = None  # a HINT only -- Layout Selector (5.4) validates, never trusts blindly (D5)
    headline: str  # working headline/title for the slide
    content_outline: str  # what this slide should say -- fed to the Content Generator (5.5)
    needs_exhibit: Literal["none", "chart", "diagram"] = "none"
