"""5.7 Diagram Builder (M4). From a caller-supplied ordered list of step
labels, build a validated DiagramData for a linear chevron-flow diagram.

Deterministic, no LLM (a deliberate narrowing of the SDD's original node/edge-
extraction idea): the caller supplies the steps directly, mirroring how chart
*data* is caller-supplied via --chart-data, never fabricated (5.6/5.7, D9).

No classify_*/suggest_*/nearest_supported_* pipeline exists here, unlike
chart_builder.py: v1 has exactly one diagram type, so there is no data-shape
-> type decision to make. This file's job today is real validation (bounds on
step count / label length, in models/diagram.py) plus this thin spec-building
function -- kept as a separate module (not inlined into the CLI) purely as a
structural extension point, so a v2 second diagram type has an obvious,
natural place to grow real decision-making logic, mirroring chart_builder.py.
"""
from typing import List

from .models.diagram import DiagramData


def build_diagram_data(steps: List[str]) -> DiagramData:
    """Validate caller-supplied step labels into a DiagramData. A thin
    pass-through today (the real logic lives in DiagramData's validators);
    the CLI calls through here rather than the raw model so a v2 real
    decision step can be inserted without CLI changes, mirroring
    build_chart_spec."""
    return DiagramData(steps=steps)
