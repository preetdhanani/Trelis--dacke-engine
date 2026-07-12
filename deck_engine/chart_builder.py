"""5.6 Chart Builder. From user-supplied structured data, deterministically
**suggest** a chart type with a one-line reason, let the caller confirm or
override, and resolve to a supported-only `ChartSpec` the Renderer turns into
a native (editable) `python-pptx` chart.

Deterministic by design (no LLM): the data's *shape* dictates the suggestion
via a fixed table (5.6). The reasoning step already happened upstream (the
Planner decided this slide needs a chart); picking a chart type from a data
shape is a lookup, not a judgement call, so an LLM here would only add cost and
nondeterminism.

D7 is enforced through `models/chart.SUPPORTED_CHART_TYPES`: every type this
module can ever suggest or resolve to is in that matrix. An override naming an
unsupported/exotic type is remapped to the nearest supported one with an
explanation (never silently honored into a broken chart, never hard-rejected).
Missing data is the caller's problem to supply, not this module's to invent
(5.6: "missing data -> prompt user, do not fabricate") -- ChartData won't even
validate without categories and at least one series.
"""
from typing import Optional, Tuple

from .models.chart import (
    NEAREST_SUPPORTED,
    SUPPORTED_CHART_TYPES,
    ChartData,
    ChartSpec,
    ChartSuggestion,
)

DEFAULT_TYPE = "column"  # the safe general categorical default

# Case-insensitive substrings that mark a category axis as time-like -> a
# single series over these reads best as a trend line.
_TIME_TOKENS = ("q1", "q2", "q3", "q4", "quarter", "month", "year", "week", "day", "fy", "h1", "h2")
_MONTHS = (
    "jan", "feb", "mar", "apr", "may", "jun",
    "jul", "aug", "sep", "oct", "nov", "dec",
)


def _looks_like_time_axis(categories) -> bool:
    """True if the category labels look like an ordered time axis (quarters,
    months, years, dates). Deterministic and conservative: it only needs to be
    right often enough to make a *suggestion*; the user still confirms."""
    hits = 0
    for c in categories:
        cl = str(c).strip().lower()
        if any(tok in cl for tok in _TIME_TOKENS) or cl[:3] in _MONTHS:
            hits += 1
            continue
        # a bare 4-digit year like "2024", or a YYYY-prefixed date
        digits = cl.replace("-", "").replace("/", "")
        if len(cl) >= 4 and cl[:4].isdigit() and 1900 <= int(cl[:4]) <= 2100:
            hits += 1
    # majority of labels time-like -> treat the whole axis as time
    return hits >= max(1, (len(categories) + 1) // 2)


def classify_data_shape(data: ChartData) -> str:
    """Return the data's shape category, per the 5.6 shape->type table:
    'multi_series_comparison' | 'time_series' | 'part_of_whole' | 'categorical'.
    """
    if len(data.series) > 1:
        return "multi_series_comparison"
    if _looks_like_time_axis(data.categories):
        return "time_series"
    if data.kind == "share":
        return "part_of_whole"
    return "categorical"


# shape -> (suggested supported type, one-line reason)
_SHAPE_TABLE = {
    "multi_series_comparison": ("column", "multiple series compared across the same categories read clearly as clustered columns"),
    "time_series": ("line", "a single series over an ordered time axis reads best as a trend line"),
    "part_of_whole": ("pie", "a single series of parts that sum to a whole reads best as a pie"),
    "categorical": ("column", "a single categorical series compares discrete values, which columns show directly"),
}

# per-shape override menu offered alongside the suggestion (supported-only)
_ALTERNATIVES = {
    "multi_series_comparison": ["bar", "line", "area"],
    "time_series": ["area", "column", "bar"],
    "part_of_whole": ["doughnut", "bar", "column"],
    "categorical": ["bar", "line", "pie"],
}


def suggest_chart_type(data: ChartData) -> ChartSuggestion:
    """Deterministic data-shape -> supported-type suggestion with a reason and
    a menu of supported alternatives to override with (5.6)."""
    shape = classify_data_shape(data)
    chart_type, reason = _SHAPE_TABLE[shape]
    alternatives = [t for t in _ALTERNATIVES[shape] if t in SUPPORTED_CHART_TYPES]
    return ChartSuggestion(chart_type=chart_type, reason=reason, alternatives=alternatives)


def nearest_supported_type(requested: str) -> Tuple[str, str]:
    """Map an unsupported/exotic requested type to the nearest supported one,
    with an explanation (D7 failure mode). Returns (supported_type, note)."""
    key = requested.strip().lower().replace(" ", "_").replace("-", "_")
    nearest = NEAREST_SUPPORTED.get(key, DEFAULT_TYPE)
    return nearest, (
        f"requested chart type {requested!r} is not supported in v1 (D7); "
        f"using the nearest supported type {nearest!r} instead"
    )


def resolve_chart_type(data: ChartData, requested_type: Optional[str] = None) -> Tuple[str, Optional[str]]:
    """Turn an optional requested/override type into a concrete supported type.

    - no request            -> the deterministic suggestion (note=None)
    - request is supported   -> honored as-is (note=None)
    - request is unsupported -> nearest supported + an explanatory note (D7)

    Returns (chart_type, note).
    """
    if requested_type is None:
        return suggest_chart_type(data).chart_type, None
    key = requested_type.strip().lower()
    if key in SUPPORTED_CHART_TYPES:
        return key, None
    return nearest_supported_type(requested_type)


def build_chart_spec(data: ChartData, requested_type: Optional[str] = None) -> ChartSpec:
    """Produce the Renderer's placement instruction: a confirmed supported
    chart type + the data. `requested_type` (from a --chart-type flag, a
    per-chart `type` in the data, or an interactive confirmation) overrides the
    suggestion; an unsupported request is remapped with a note, never dropped."""
    chart_type, note = resolve_chart_type(data, requested_type)
    return ChartSpec(chart_type=chart_type, data=data, note=note)
