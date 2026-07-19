"""Revision models (SDD 5.10, M6). The contract between the Revision Engine's
feedback-log bookkeeping and the caller.
"""
from typing import Dict, List, Literal, Optional

from pydantic import BaseModel

from .manifest import Layout
from .slide_spec import SlideSpec


class FeedbackLogEntry(BaseModel):
    """One row of a slide's accumulating feedback log (SDD 6.1). `kind` is
    "summary" for an older entry folded down to a single line once the log
    exceeds MAX_LOG_ENTRIES (5.10: "cap at last 3-5 attempts; summarize older
    attempts into a single line") -- never silently dropped."""
    attempt: int
    kind: Literal["field_edit", "regen", "summary"]
    user_feedback: str
    spec_snapshot: Optional[Dict[str, object]] = None


class RevisionState(BaseModel):
    """Everything the Revision Engine needs to revise one slide again: its
    layout, the content/role context a regen would reuse, its current spec,
    and the capped feedback log. Not persisted (SDD 5.12's state store is a
    separate, unbuilt milestone) -- lives for the duration of one run.
    """
    label: str  # human-readable, e.g. "title slide" or "slide 3 (text_bullets)"
    layout: Layout
    role: str  # passed to generate_single_slide's `role` on regen
    content_brief: str  # the base content (headline + outline) a regen starts from
    current_spec: SlideSpec
    feedback_log: List[FeedbackLogEntry] = []
    attempts: int = 0  # total revise() calls ever made (used for log entry numbering)
    failed_attempts: int = 0  # attempts that produced no usable spec -- this is what escalates (5.10)
    escalated: bool = False
    last_error: Optional[str] = None  # set when the most recent attempt failed to produce a usable spec
