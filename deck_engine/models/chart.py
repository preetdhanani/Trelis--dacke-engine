"""Chart models (SDD 5.6, D7). The contract between user-supplied structured
data, the Chart Builder's type suggestion, and the Renderer's native-chart
insertion.

D7's support matrix is enforced here in one place: `SUPPORTED_CHART_TYPES`
maps our canonical type names to `python-pptx`'s `XL_CHART_TYPE` members. Only
categorical-data chart types are in v1 -- column/bar/line/pie/doughnut/area --
because they all build correctly from `CategoryChartData` (categories + one or
more series of values). D7 also *names* scatter/bubble/stock "if needed", but
those need an XY/OHLC data model rather than categories+series, so they are
deliberately out of M3's scope (documented deferral, not a silent omission) --
offering them here would produce broken charts from categorical data.
"""
from typing import Dict, List, Literal, Optional

from pydantic import BaseModel, field_validator, model_validator

# Canonical name -> XL_CHART_TYPE member name. The renderer resolves the member
# off pptx.enum.chart.XL_CHART_TYPE by this name, so the enum import stays out
# of the model layer (which the CLI and tests import without needing pptx).
SUPPORTED_CHART_TYPES: Dict[str, str] = {
    "column": "COLUMN_CLUSTERED",
    "bar": "BAR_CLUSTERED",
    "line": "LINE_MARKERS",
    "pie": "PIE",
    "doughnut": "DOUGHNUT",
    "area": "AREA",
}

# Common exotic / unsupported requests -> the nearest supported type (D7's
# "requested unsupported type -> suggest nearest supported + explain" failure
# mode). Anything not in this map falls back to the general default, "column".
NEAREST_SUPPORTED: Dict[str, str] = {
    "waterfall": "column",
    "funnel": "bar",
    "mekko": "column",
    "marimekko": "column",
    "gantt": "bar",
    "treemap": "pie",
    "sunburst": "doughnut",
    "stacked_column": "column",
    "stacked_bar": "bar",
    "histogram": "column",
    "scatter": "line",
    "bubble": "line",
    "stock": "line",
    "radar": "line",
    "spider": "line",
}


class ChartData(BaseModel):
    """User-supplied structured data for one chart. Never fabricated by the
    engine (D9 / 5.6 "missing data -> prompt user, do not fabricate"): it comes
    from the caller (CLI --chart-data), mirroring CategoryChartData's shape.
    """
    categories: List[str]
    series: Dict[str, List[float]]  # series name -> one value per category
    title: Optional[str] = None
    # An explicit part-of-whole hint. The shape classifier cannot reliably infer
    # "these values are shares of a whole" from numbers alone, so the caller can
    # say so; when set, a single-series chart is suggested as pie rather than
    # column. Purely advisory -- the user still confirms/overrides.
    kind: Optional[Literal["share"]] = None
    # Per-chart type override (skips the suggestion). Validated against the
    # support matrix downstream, not here, so an unsupported value can be
    # remapped-with-explanation rather than hard-rejected at parse time.
    type: Optional[str] = None

    @field_validator("categories")
    @classmethod
    def _at_least_one_category(cls, v):
        if not v:
            raise ValueError("chart data needs at least one category")
        return v

    @field_validator("series")
    @classmethod
    def _at_least_one_series(cls, v):
        if not v:
            raise ValueError("chart data needs at least one series")
        return v

    @model_validator(mode="after")
    def _series_lengths_match_categories(self):
        n = len(self.categories)
        for name, values in self.series.items():
            if len(values) != n:
                raise ValueError(
                    f"series {name!r} has {len(values)} values but there are {n} categories"
                )
        return self


class ChartSuggestion(BaseModel):
    """The Chart Builder's proposal: a supported type, a one-line reason, and
    the other supported types offered as overrides (5.6: "suggest with a
    one-line reason; user confirms/overrides")."""
    chart_type: str
    reason: str
    alternatives: List[str]


class ChartSpec(BaseModel):
    """The resolved placement instruction the Renderer consumes: a confirmed
    supported chart type + the data to build it from. `note` carries any
    caveat worth surfacing (e.g. an unsupported request was remapped)."""
    chart_type: str
    data: ChartData
    note: Optional[str] = None
