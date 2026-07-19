"""5.9 QA / Overflow Checker (M5). Detects text/bullets slots whose rendered
content is likely to overflow their placeholder box, flagging the slide
needs_review rather than shipping a silently broken layout (D9).

Real-template discovery (updates the original SDD 5.9 assumption): every
text/bullets shape in MasterDeck.pptx uses MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT --
PowerPoint grows the box to fit its text rather than clipping it. So the
practical failure mode isn't invisible clipped text; it's a grown box
visually colliding with whatever sits below/around it, or bleeding off the
slide. Detection therefore compares estimated text height against a *safe
growth budget*: the slot's own designed height plus the real headroom to the
nearest element below it (or to the slide's bottom edge if nothing is below),
not just the slot's own static height.

No new dependency and no LLM/vision call: python-pptx cannot run PowerPoint's
text-layout engine, so height is estimated with a standard proportional-font
line-wrap heuristic (chars-per-line from box width and font size), not
measured pixel-exact. Per D9/5.9: false positives are acceptable (route to
review); false negatives are the risk to minimize.
"""
from math import ceil

AVG_CHAR_WIDTH_FACTOR = 0.5  # avg glyph width ~= 0.5 * font size, typical for Cambria/Calibri-class fonts
LINE_SPACING_FACTOR = 1.2  # typical PowerPoint single-line-spacing multiplier
TEXT_INSET_IN = 0.2  # python-pptx/PowerPoint default ~0.1in left+right text-frame inset
DEFAULT_BOTTOM_MARGIN_IN = 0.3  # breathing room kept from the slide's bottom edge when nothing sits below


def estimate_wrapped_line_count(text, chars_per_line):
    """Estimate wrapped line count for one paragraph. A reasoned estimate,
    not a rendered measurement -- word-wrap doesn't break exactly on this
    count (no font-layout engine available without a new dependency)."""
    if not text:
        return 1
    return max(1, ceil(len(text) / chars_per_line))


def estimate_text_height_in(paragraphs, font_size_pt, box_width_in):
    """Estimate the rendered height (inches) of one or more paragraphs (a
    single string for a text slot, or a list of bullet items -- one
    paragraph per item) inside a box of the given width, at the given font
    size."""
    if isinstance(paragraphs, str):
        paragraphs = [paragraphs]
    usable_width_in = max(0.1, box_width_in - TEXT_INSET_IN)
    box_width_pt = usable_width_in * 72
    chars_per_line = max(1, box_width_pt / (font_size_pt * AVG_CHAR_WIDTH_FACTOR))
    total_lines = sum(estimate_wrapped_line_count(p, chars_per_line) for p in paragraphs)
    line_height_in = (font_size_pt * LINE_SPACING_FACTOR) / 72
    return total_lines * line_height_in


def compute_headroom_in(slot, layout, slide_height_in, bottom_margin_in=DEFAULT_BOTTOM_MARGIN_IN):
    """How much a slot's box may grow downward before it collides with the
    nearest shape below it in the same horizontal band, or bleeds past the
    slide's bottom edge if nothing is below it. Pure geometry over the
    manifest's own slot geometry_in -- no pptx object needed, so this is
    testable with plain manifest data."""
    geom = slot.geometry_in
    this_left, this_right = geom.left_in, geom.left_in + geom.width_in
    this_bottom = geom.top_in + geom.height_in

    nearest_below = None
    for other in layout.slots:
        if other is slot or other.geometry_in is None:
            continue
        og = other.geometry_in
        overlaps_horizontally = og.left_in < this_right and (og.left_in + og.width_in) > this_left
        if overlaps_horizontally and og.top_in >= this_bottom:
            if nearest_below is None or og.top_in < nearest_below:
                nearest_below = og.top_in

    ceiling_in = nearest_below if nearest_below is not None else (slide_height_in - bottom_margin_in)
    return max(0.0, ceiling_in - this_bottom)


def check_overflow(paragraphs, font_size_pt, slot, layout, slide_height_in):
    """Returns an overflow reason string if the estimated text height exceeds
    the slot's safe growth budget (its own height + headroom to the next
    element/slide edge); None if it looks safe or there isn't enough
    geometry/font information to judge (no data to check -> no verdict,
    never a guess)."""
    if slot.geometry_in is None or font_size_pt is None:
        return None
    estimated_in = estimate_text_height_in(paragraphs, font_size_pt, slot.geometry_in.width_in)
    headroom_in = compute_headroom_in(slot, layout, slide_height_in)
    budget_in = slot.geometry_in.height_in + headroom_in
    if estimated_in > budget_in:
        return (
            f"estimated text height {estimated_in:.2f}in exceeds its safe growth budget "
            f"{budget_in:.2f}in (box {slot.geometry_in.height_in:.2f}in + {headroom_in:.2f}in headroom) "
            "-- may visually overflow/collide with neighboring content"
        )
    return None
