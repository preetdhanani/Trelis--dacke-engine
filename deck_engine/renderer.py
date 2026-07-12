"""5.8 Renderer -- the fill engine (deterministic). Implements the seed-slide-
duplication mechanic (D3) with duplicate-before-fill ordering (D13): every
output slide is produced by duplicating a *pristine* seed slide from the
source template and filling the duplicate's named shapes -- never via
`add_slide(layout)` + `placeholders`, and never by re-cloning an
already-filled slide (confirmed necessary by the spike's negative control).
"""
import copy

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches

from .models.chart import SUPPORTED_CHART_TYPES

IMAGE_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def duplicate_slide(output_prs, seed_slide):
    """Duplicate `seed_slide` (from the pristine template) into `output_prs` as
    a new slide, with the seed's full content already in place. Filling still
    happens afterward via fill_slot().

    Deliberately does NOT call output_prs.slides.add_slide(): that convenience
    method calls `slide.shapes.clone_layout_placeholders(...)` internally,
    which touches `.shapes` -- a lazyproperty, cached on first access -- before
    the seed's content is swapped in below. Once poisoned, `.shapes` keeps
    returning that stale, empty collection for the rest of that Slide object's
    life, even after the cSld swap installs a real, populated shape tree.
    Going through the low-level part API instead avoids ever touching
    `.shapes` until after the swap is complete.
    """
    slide_layout = seed_slide.slide_layout
    rId, dest = output_prs.part.add_slide(slide_layout)
    output_prs.slides._sldIdLst.add_sldId(rId)

    src_cSld = seed_slide._element.find(qn("p:cSld"))
    dest_cSld = dest._element.find(qn("p:cSld"))
    new_cSld = copy.deepcopy(src_cSld)
    dest._element.replace(dest_cSld, new_cSld)

    # Remap every image relationship the clone carries (D13). A naive XML
    # clone copies r:embed ids verbatim; those ids are dangling in the new
    # slide part's own .rels unless explicitly re-related here -- this is the
    # exact failure the spike's negative control reproduced and confirmed.
    for blip in new_cSld.findall(".//" + qn("a:blip")):
        r_embed = blip.get(qn("r:embed"))
        if not r_embed:
            continue
        image_part = seed_slide.part.related_part(r_embed)
        new_rId = dest.part.relate_to(image_part, IMAGE_RT)
        blip.set(qn("r:embed"), new_rId)

    return dest


def find_shape(slide, name):
    for shp in slide.shapes:
        if shp.name == name:
            return shp
    return None


def _set_text_preserving_format(text_frame, value):
    paragraphs = text_frame.paragraphs
    p0 = paragraphs[0]
    if not p0.runs:
        p0.text = value
    else:
        p0.runs[0].text = value
        for extra_run in p0.runs[1:]:
            extra_run._r.getparent().remove(extra_run._r)
    for extra_p in paragraphs[1:]:
        extra_p._p.getparent().remove(extra_p._p)


def _set_bullets_preserving_format(text_frame, items):
    template_p = text_frame.paragraphs[0]._p
    parent = template_p.getparent()

    new_ps = []
    for item in items:
        p_clone = copy.deepcopy(template_p)
        r_elems = p_clone.findall(qn("a:r"))
        if r_elems:
            r_elems[0].find(qn("a:t")).text = item
            for extra in r_elems[1:]:
                p_clone.remove(extra)
        new_ps.append(p_clone)

    for p in list(text_frame.paragraphs):
        p._p.getparent().remove(p._p)
    for p_el in new_ps:
        parent.append(p_el)


def _remove_shape(shape):
    shape._element.getparent().remove(shape._element)


def _add_native_chart(slide, slot, chart_spec):
    """Insert a native (editable) python-pptx chart at the slot's geometry per
    D7/5.8-point-5 -- the *preferred* way to fill an exhibit slot, since it
    stays a real editable chart rather than a flat picture. The chart type name
    is resolved through SUPPORTED_CHART_TYPES so nothing outside the D7 matrix
    can ever reach add_chart().
    """
    xl_member = SUPPORTED_CHART_TYPES.get(chart_spec.chart_type)
    if xl_member is None:  # defensive: build_chart_spec should already guarantee this
        raise RuntimeError(f"unsupported chart type {chart_spec.chart_type!r} reached the renderer (D7)")
    xl_type = getattr(XL_CHART_TYPE, xl_member)

    data = chart_spec.data
    chart_data = CategoryChartData()
    chart_data.categories = data.categories
    for series_name, values in data.series.items():
        chart_data.add_series(series_name, values)

    geom = slot.geometry_in
    graphic_frame = slide.shapes.add_chart(
        xl_type,
        Inches(geom.left_in), Inches(geom.top_in),
        Inches(geom.width_in), Inches(geom.height_in),
        chart_data,
    )
    # keep the exhibit addressable by its manifest name (D3), same as the
    # picture path renames its inserted shape.
    graphic_frame.name = slot.shape_name

    chart = graphic_frame.chart
    if data.title:
        chart.has_title = True
        chart.chart_title.text_frame.text = data.title
    else:
        chart.has_title = False
    # a legend only earns its space when there's more than one series
    chart.has_legend = len(data.series) > 1
    return graphic_frame


def fill_slot(slide, slot, value, image_source_path=None, chart_spec=None):
    """Returns True if the slot was actually filled, False if it was left as
    the seed's own placeholder content. Image slots with no source are
    skipped rather than raising: a slide that can't get a real exhibit must be
    flagged (needs_review), never silently broken or crashed on (D9) -- the
    caller (render_slide) surfaces which required slots were skipped so the CLI
    can flag them.

    An image-typed exhibit slot can be filled two ways (5.8 point 5, D7): a
    native `chart_spec` (preferred -- stays editable) takes precedence, else a
    rendered `image_source_path` picture; with neither, the grey placeholder
    rect is left in place.
    """
    shape = find_shape(slide, slot.shape_name)
    if shape is None:
        raise RuntimeError(f"named shape '{slot.shape_name}' not found on duplicated slide")

    if slot.type == "text":
        if value is None:
            return False
        text = value if not slot.max_chars or len(value) <= slot.max_chars else value[: slot.max_chars - 1].rstrip() + "…"
        _set_text_preserving_format(shape.text_frame, text)
        return True

    elif slot.type == "bullets":
        if not value:
            return False
        items = list(value)
        if slot.max_items:
            items = items[: slot.max_items]
        if slot.max_chars_per_item:
            items = [
                i if len(i) <= slot.max_chars_per_item else i[: slot.max_chars_per_item - 1].rstrip() + "…"
                for i in items
            ]
        _set_bullets_preserving_format(shape.text_frame, items)
        return True

    elif slot.type == "image":
        if chart_spec is not None:
            _add_native_chart(slide, slot, chart_spec)
            _remove_shape(shape)  # remove the grey placeholder rect, per fill_protocol
            return True
        if image_source_path is None:
            return False  # grey placeholder rect stays -- caller flags needs_review
        geom = slot.geometry_in
        picture = slide.shapes.add_picture(
            str(image_source_path),
            Inches(geom.left_in), Inches(geom.top_in),
            width=Inches(geom.width_in), height=Inches(geom.height_in),
        )
        # add_picture() names the new shape generically ("Picture N"); rename it
        # back to the slot's shape_name so manifest addressing (D3) still
        # resolves this slot on a future lookup (spike S1 finding).
        picture.name = slot.shape_name
        _remove_shape(shape)  # remove the grey placeholder rect, per fill_protocol
        return True

    else:
        raise RuntimeError(f"unknown slot type: {slot.type}")


def render_slide(output_prs, seed_slide, layout, spec, image_paths=None, chart_specs=None):
    """Duplicate `seed_slide` into `output_prs` and fill it per `layout`'s
    slots from `spec.slots` (D3: duplicate first; D13: fill only the
    duplicate, never the seed).

    Image-typed exhibit slots are filled from `chart_specs[shape_name]` (a
    native chart, preferred) or `image_paths[shape_name]` (a picture); a slot
    with neither is left as the grey placeholder.

    Returns (new_slide, skipped_required_slots): the second element lists the
    shape_name of every *required* slot that was left unfilled (only possible
    for image slots with no chart or picture source -- see fill_slot) so the
    caller can flag the slide needs_review instead of shipping it silently
    incomplete.
    """
    image_paths = image_paths or {}
    chart_specs = chart_specs or {}
    new_slide = duplicate_slide(output_prs, seed_slide)
    skipped_required = []
    for slot in layout.slots:
        if slot.type == "image":
            filled = fill_slot(
                new_slide, slot, None,
                image_source_path=image_paths.get(slot.shape_name),
                chart_spec=chart_specs.get(slot.shape_name),
            )
        else:
            filled = fill_slot(new_slide, slot, spec.slots.get(slot.shape_name))
        if not filled and slot.required:
            skipped_required.append(slot.shape_name)
    return new_slide, skipped_required


def strip_seed_slides(prs, count):
    """Remove the first `count` slides -- the pristine seed/index slides that
    came along when the output presentation was opened from a copy of the
    template -- leaving only the newly-rendered slides. Must be called *after*
    all render_slide() calls; removing by index while more slides are still
    being appended would shift indices under you.
    """
    id_lst = prs.slides._sldIdLst
    sld_ids = list(id_lst)
    for i in range(count - 1, -1, -1):
        id_lst.remove(sld_ids[i])
