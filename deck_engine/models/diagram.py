"""Diagram models (SDD 5.7, M4). v1 supports exactly one diagram type --
linear chevron-flow -- so there is no data-shape -> type decision to make
(unlike chart.py's 6-type matrix). This is a deliberate, documented narrowing
of the original SDD 5.7 concept (LLM node/edge extraction from prose;
linear-or-branch layout; image-fallback for diagrams too complex to lay out):
v1 takes caller-supplied ordered step labels directly (never fabricated, D9,
mirroring --chart-data), is linear-only (no branching), and has no complexity
axis left to overwhelm a layout algorithm -- "too complex" degenerates to a
bounded step-count validation error here, not a layout failure needing an
image-fallback path.

SUPPORTED_DIAGRAM_TYPES documents today's only value, in the same spot
SUPPORTED_CHART_TYPES lives in chart.py, so a v2 second diagram type has an
obvious, natural place to grow a real classify/suggest/resolve pipeline plus
a DiagramSpec wrapper analogous to ChartSpec. Neither exists yet -- a wrapper
holding one constant type and an always-None note would be ceremony with
nothing to resolve between (documented deferral, not a silent omission,
matching how chart.py defers scatter/bubble/stock).
"""
from typing import List

from pydantic import BaseModel, field_validator

SUPPORTED_DIAGRAM_TYPES = {"chevron_flow"}  # v1's only value; a v2 second type joins this set

MIN_STEPS = 2
MAX_STEPS = 6
# Conservative bound so a label stays legible inside a chevron even at the
# tightest case (MAX_STEPS chevrons sharing IM_IMAGEONLY_HERO's 11.53in
# width): after inter-chevron gaps each chevron is ~1.8in wide, and the
# chevron's arrow point/notch tapers eat further into usable interior width.
# At a 12pt font, wrapped over 2-3 lines, ~24 chars is a safe ceiling.
MAX_CHARS_PER_STEP = 24


class DiagramData(BaseModel):
    """User-supplied, ordered step labels for a linear chevron-flow diagram.
    Never fabricated by the engine (D9, mirroring ChartData): it comes from
    the caller (CLI --diagram-data), in display order, one chevron per step.
    """
    steps: List[str]

    @field_validator("steps")
    @classmethod
    def _step_count_in_range(cls, v):
        if not (MIN_STEPS <= len(v) <= MAX_STEPS):
            raise ValueError(
                f"a linear chevron-flow diagram needs between {MIN_STEPS} and {MAX_STEPS} "
                f"steps (got {len(v)})"
            )
        return v

    @field_validator("steps")
    @classmethod
    def _steps_are_nonblank_and_within_length(cls, v):
        for i, step in enumerate(v):
            if not step.strip():
                raise ValueError(f"step[{i}] is empty or blank")
            if len(step) > MAX_CHARS_PER_STEP:
                raise ValueError(
                    f"step[{i}] {step!r} exceeds max_chars_per_step={MAX_CHARS_PER_STEP} "
                    f"(got {len(step)}); shorten it so it fits legibly inside a chevron"
                )
        return v
