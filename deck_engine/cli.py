"""M2/M3 -- brief -> deck across the full 8-layout set, fully unattended.

Pipeline: Deck Planner (5.3) turns the brief into an ordered list of
SlideIntents -> Layout Selector (5.4) maps each intent to a concrete manifest
layout (validating the Planner's suggested_layout hint, never trusting it
blindly, D5) -> Content Generator (5.5) fills that layout's slots -> Chart
Builder (5.6) turns user-supplied data into a native chart for exhibit slots
-> Renderer (5.8) duplicates the right seed slide and fills it (D3/D13). Every
deck is bookended by a fixed `title` slide and a fixed `closing_contact` slide
-- these are NOT planner-decided (see content_generator.py's module docstring
for why); the Planner only plans what goes in between.

Exhibit slots (M3): a slide the Planner marked needs_exhibit=chart gets a
native, editable python-pptx chart if the caller supplied its data via
--chart-data (mapped in planner order to chart-needing slides). Chart data is
never fabricated (5.6/D9) -- a chart-needing slide with no data supplied
renders with its grey placeholder rect left in place and is flagged
needs_review, never silently shipped broken.

Diagram slots (M4): a slide the Planner marked needs_exhibit=diagram gets a
native, editable chevron-flow diagram (a grouped row of chevron autoshapes,
5.7) if the caller supplied its ordered step labels via --diagram-data
(mapped in planner order to diagram-needing slides, exactly like chart data).
Diagram steps are never fabricated (D9) -- same placeholder + needs_review
treatment when no data is supplied.

Usage (from the repo root, so the `deck_engine` package is importable):
    python -m deck_engine.cli --brief "..." --out output/deck.pptx
    python -m deck_engine.cli --brief-file brief.txt --out output/deck.pptx

Provider defaults to Anthropic (claude-sonnet-5, the SDD's reference model --
requires ANTHROPIC_API_KEY). Swap with --provider ollama (local, no key --
requires an Ollama server running with the model pulled) or --provider gemini
(requires GOOGLE_API_KEY or GEMINI_API_KEY). --model overrides the
provider's default model name.

    python -m deck_engine.cli --provider ollama --model gemma4:latest --brief "..." --out output/deck.pptx
    python -m deck_engine.cli --provider gemini --brief "..." --out output/deck.pptx

The closing slide's contact block is real personal information (name, role,
email, phone) -- the LLM never invents it. Pass it with --contact, e.g.:
    python -m deck_engine.cli --brief "..." --contact "Jane Doe|Senior Partner|jane@firm.com * 555-2020" --out output/deck.pptx
--contact accepts literal "\\n" or "|" as line separators (both become real
newlines in the slide). If omitted, the closing slide is still rendered but
keeps the template's own placeholder contact text -- a clear warning is
printed rather than silently shipping a real-looking but fake contact block.
"""
import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

from .chart_builder import build_chart_spec, suggest_chart_type
from .config import DEFAULT_TENANT_ID, DEFAULT_PROVIDER
from .content_generator import generate_single_slide
from .deck_planner import plan_deck
from .diagram_builder import build_diagram_data
from .layout_selector import select_layout
from .llm_providers import DEFAULT_MODEL, PROVIDERS, check_provider_ready, ProviderError
from .models.chart import ChartData
from .models.revision import RevisionState
from .models.slide_spec import SlideSpec
from .renderer import render_slide, replace_rendered_slide, strip_seed_slides
from .revision_engine import revise
from .template_registry import load_tenant_assets, ManifestValidationError

TITLE_LAYOUT_ID = "title"
CLOSING_LAYOUT_ID = "closing_contact"


def build_arg_parser():
    p = argparse.ArgumentParser(description="M2/M3: brief -> .pptx deck across the full 8-layout set, unattended.")
    p.add_argument("--brief", help="Brief text, inline.")
    p.add_argument("--brief-file", type=Path, help="Path to a text file containing the brief.")
    p.add_argument("--contact", help='Closing slide contact block, e.g. "Jane Doe|Senior Partner|jane@firm.com * 555-2020".')
    p.add_argument("--tenant", default=DEFAULT_TENANT_ID)
    p.add_argument("--provider", choices=PROVIDERS, default=DEFAULT_PROVIDER)
    p.add_argument("--model", default=None, help="Defaults to the chosen provider's reference model if omitted.")
    p.add_argument("--min-slides", type=int, default=2, help="Minimum content slides for the Deck Planner (excludes title/closing).")
    p.add_argument("--max-slides", type=int, default=8, help="Maximum content slides for the Deck Planner (excludes title/closing).")
    p.add_argument(
        "--chart-data",
        type=Path,
        help="JSON file: a list of chart-data blocks {categories, series, title?, kind?, type?}, mapped in planner "
        "order to slides the Planner marked needs_exhibit=chart. The engine never fabricates chart data (5.6/D9).",
    )
    p.add_argument(
        "--chart-type",
        default=None,
        help="Global chart-type override (column|bar|line|pie|doughnut|area) for chart blocks lacking their own "
        "'type'. Unsupported/exotic types are remapped to the nearest supported one with a note (D7).",
    )
    p.add_argument(
        "--confirm-charts",
        action="store_true",
        help="Interactively confirm/override each chart's suggested type. Default: auto-accept the deterministic "
        "suggestion (unattended), printing what was chosen and why.",
    )
    p.add_argument(
        "--diagram-data",
        type=Path,
        help="JSON file: a list of diagram-data blocks {steps: [str, ...]} (2-6 linear steps each), mapped in "
        "planner order to slides the Planner marked needs_exhibit=diagram. Rendered as a native, editable "
        "chevron-flow diagram (5.7/M4). Diagram steps are never fabricated (D9).",
    )
    p.add_argument(
        "--interactive-revise",
        action="store_true",
        help="After the deck is generated, offer an interactive loop (5.10, M6) to flag any single "
        "text slide with free-text feedback and see it revised in place -- a small ask ('change 12%% "
        "to 15%%') edits just that value; a structural ask ('re-approach this slide') triggers a full "
        "regeneration. Capped at 3 failed attempts per slide, then it's flagged for a human look. "
        "The closing slide is excluded (its contact info is never LLM-generated/revised).",
    )
    p.add_argument("--out", type=Path, required=True)
    return p


def _load_chart_data(path):
    """Read the --chart-data file into a list of raw chart-block dicts. Accepts
    either a top-level JSON list or a single block object."""
    raw = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(raw, dict):
        raw = [raw]
    if not isinstance(raw, list):
        raise ValueError("--chart-data must be a JSON list of chart blocks (or a single block object)")
    return raw


def _load_diagram_data(path):
    """Read the --diagram-data file into a list of raw diagram-block dicts.
    Accepts either a top-level JSON list or a single block object. Kept as a
    twin of _load_chart_data (not a shared helper) so the working chart path
    stays untouched."""
    raw = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(raw, dict):
        raw = [raw]
    if not isinstance(raw, list):
        raise ValueError("--diagram-data must be a JSON list of diagram blocks (or a single block object)")
    return raw


def _confirm_chart_type(suggestion, title, input_fn=input):
    """Present the Chart Builder's suggestion + alternatives and let the user
    accept (Enter) or type an override (5.6: "user confirms/overrides").
    Returns the override string, or None to accept the suggestion. `input_fn`
    is injectable so this is testable without a live TTY."""
    label = f" for '{title}'" if title else ""
    print(f"  chart{label}: suggested '{suggestion.chart_type}' -- {suggestion.reason}")
    print(f"    supported alternatives: {', '.join(suggestion.alternatives)}")
    resp = input_fn(f"    Enter to accept '{suggestion.chart_type}', or type a type to override: ").strip()
    return resp or None


def _normalize_contact(raw: str) -> str:
    text = raw.replace("\\n", "\n")
    if "\n" not in text and "|" in text:
        text = "\n".join(part.strip() for part in text.split("|"))
    return text


def _print_overflow_flags(slide_label, layout_id, overflow_warnings):
    """QA/Overflow Checker (5.9, M5): print a needs_review flag for every
    text/bullets slot whose estimated height exceeds its safe growth budget.
    Mirrors the existing skipped-required-slot flagging below -- never a
    silent broken slide (D9)."""
    for shape_name, reason in overflow_warnings:
        print(
            f"  FLAGGED (needs_review) {slide_label} ({layout_id}): slot '{shape_name}' {reason}",
            file=sys.stderr,
        )


def _print_render_flags(label, layout_id, skipped, overflow, needs_exhibit=None):
    """Shared FLAGGED-printer for both the initial render loop and the
    revision loop (5.10, M6) -- same messages either way."""
    for shape_name in skipped:
        if needs_exhibit == "chart":
            reason = "needs a chart but no --chart-data block was provided for it"
        elif needs_exhibit == "diagram":
            reason = "needs a diagram but no --diagram-data block was provided for it"
        else:
            reason = "has no exhibit source available"
        print(
            f"  FLAGGED (needs_review) {label} ({layout_id}): required slot '{shape_name}' {reason} "
            "-- placeholder left in place",
            file=sys.stderr,
        )
    _print_overflow_flags(label, layout_id, overflow)


def _open_in_default_app(path):
    """Best-effort cross-platform "open this file" -- lets the user actually
    look at the deck mid-revision (5.10/M6) instead of judging it from a text
    label alone. Never fatal: if it fails, the caller still has the saved
    path to open by hand."""
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # noqa: only defined on Windows
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
        return True
    except OSError:
        return False


def _format_slide_summary(spec, indent="     "):
    """Render a slide's current slot values as readable terminal lines --
    the closest thing to "seeing the slide" without opening PowerPoint (5.10/
    M6): feedback like "change 12% to 15%" is hard to give blind."""
    lines = []
    for shape_name, value in spec.slots.items():
        if value is None or value == "" or value == []:
            continue
        text = ", ".join(value) if isinstance(value, list) else str(value)
        if len(text) > 90:
            text = text[:87] + "..."
        lines.append(f"{indent}{shape_name}: {text!r}")
    return "\n".join(lines) if lines else f"{indent}(no text content in this exhibit/placeholder slide)"


def _run_interactive_revision(output_prs, revisable, provider, model, deck_context, out_path):
    """5.10/M6: let the user flag any single revisable slide with free-text
    feedback and see it revised in place, looping until they're done. Every
    other slide is untouched (FR-6); regeneration always re-duplicates the
    pristine seed fresh (D13) via replace_rendered_slide.

    Charts/diagrams don't have a meaningful text summary, so 'p'/'preview'
    saves the deck-so-far to `out_path` and opens it in the system's default
    app -- the only way to actually *see* an exhibit slide before deciding
    what feedback to give.
    """
    print("\n--- Interactive revision (M6) ---")
    print("Each slide's current text is shown below its number. Flag one to revise it,")
    print("type 'p' to save + open the deck so far in PowerPoint, or Enter to finish.\n")
    while True:
        for i, r in enumerate(revisable, 1):
            flag = " [ESCALATED -- needs a human look]" if r["state"].escalated else ""
            print(f"  {i}. {r['label']}{flag}")
            print(_format_slide_summary(r["state"].current_spec))
        choice = input("\nSlide number to revise, 'p' to preview, or Enter to finish: ").strip()
        if not choice:
            break
        if choice.lower() in ("p", "preview"):
            out_path.parent.mkdir(parents=True, exist_ok=True)
            output_prs.save(str(out_path))
            print(f"  saved current state to {out_path} -- opening it now...")
            if not _open_in_default_app(out_path):
                print(f"  couldn't auto-open it -- open {out_path} manually to take a look.")
            continue
        try:
            entry = revisable[int(choice) - 1]
        except (ValueError, IndexError):
            print("  not a valid slide number.")
            continue
        if entry["state"].escalated:
            print("  this slide already hit the 3-failed-attempt cap -- pick another slide.")
            continue
        feedback = input(f"  Feedback for {entry['label']}: ").strip()
        if not feedback:
            continue

        entry["state"] = revise(entry["state"], feedback, provider, model, deck_context=deck_context)
        if entry["state"].last_error is not None:
            print(f"  revision did not produce a usable slide ({entry['state'].last_error}); slide left unchanged.")
            if entry["state"].escalated:
                print("  ESCALATED: 3 failed attempts on this slide -- stopping automatic revision here.")
            continue

        _, skipped, overflow = replace_rendered_slide(
            output_prs, entry["position"], entry["seed"], entry["layout"], entry["state"].current_spec,
            **entry["render_kwargs"],
        )
        _print_render_flags(entry["label"], entry["layout"].layout_id, skipped, overflow, entry.get("needs_exhibit"))
        print(f"  revised and re-rendered {entry['label']}.")


def _content_brief(intent) -> str:
    """The per-slide prompt fed to the Content Generator -- the intent's own
    outline, not the full original brief, since the Planner already did the
    decomposition."""
    return f"{intent.headline}\n\n{intent.content_outline}"


def main(argv=None):
    args = build_arg_parser().parse_args(argv)

    if not args.brief and not args.brief_file:
        print("error: provide --brief or --brief-file", file=sys.stderr)
        return 2
    brief = args.brief or args.brief_file.read_text(encoding="utf-8")
    model = args.model or DEFAULT_MODEL[args.provider]

    try:
        check_provider_ready(args.provider)
    except ProviderError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    try:
        assets = load_tenant_assets(args.tenant)
    except ManifestValidationError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    manifest = assets.manifest
    title_layout = manifest.layout_by_id(TITLE_LAYOUT_ID)
    closing_layout = manifest.layout_by_id(CLOSING_LAYOUT_ID)

    try:
        print(f"[1/4] generating title slide (provider={args.provider} model={model})...")
        title_spec, title_failure = generate_single_slide(brief, title_layout, args.provider, model, role="opening/title slide")
        if title_failure:
            print(f"  FLAGGED (needs_review) title slide: {title_failure['errors']}", file=sys.stderr)
            # Title is a mandatory structural bookend (D3) -- every deck opens
            # with one, regardless of brief content. Even if both generation
            # attempts failed business-rule validation (e.g. a one-character
            # max_chars overflow), render the best-effort content instead of
            # silently shipping a deck with no opening slide at all: a
            # slightly-over-length title is a cosmetic risk, a missing title
            # slide is a structural one.
            title_spec = SlideSpec(layout_id=TITLE_LAYOUT_ID, content_kind="text", slots=title_failure["slots"])

        print("[2/4] planning deck structure...")
        intents, plan_error = plan_deck(brief, manifest, args.provider, model, args.min_slides, args.max_slides)
        if plan_error:
            print(f"error: deck planning failed: {plan_error}", file=sys.stderr)
            return 1
        print(f"  {len(intents)} content slide(s) planned")

        print("[3/4] selecting layouts + generating content per slide...")
        content_slides = []  # list of (layout, spec, intent)
        content_failed = []
        for i, intent in enumerate(intents, 1):
            layout, select_flag = select_layout(intent, manifest, args.provider, model)
            if select_flag:
                print(f"  slide {i}: {select_flag}", file=sys.stderr)
            spec, gen_failure = generate_single_slide(
                _content_brief(intent), layout, args.provider, model, role=f"content slide ({intent.purpose})"
            )
            if gen_failure:
                print(f"  FLAGGED (needs_review) slide {i} ({layout.layout_id}): {gen_failure['errors']}", file=sys.stderr)
                content_failed.append(gen_failure)
                continue
            content_slides.append((layout, spec, intent))
    except ProviderError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    if not content_slides:
        print("error: no valid content slides were generated -- nothing to render", file=sys.stderr)
        return 1

    # Chart Builder (5.6, M3): resolve user-supplied chart data into native
    # chart specs, keyed by content-slide index -> {shape_name: ChartSpec}.
    # Chart data is mapped in planner order to slides the Planner marked
    # needs_exhibit=chart whose selected layout actually has an image slot.
    # The engine never invents chart data (5.6/D9): a chart-needing slide with
    # no data provided just keeps its placeholder and gets flagged at render.
    chart_specs_per_slide = {}
    chart_blocks = []
    if args.chart_data:
        try:
            chart_blocks = _load_chart_data(args.chart_data)
        except (OSError, ValueError, json.JSONDecodeError) as e:
            print(f"error: could not read --chart-data: {e}", file=sys.stderr)
            return 1

    chart_cursor = 0
    for idx, (layout, spec, intent) in enumerate(content_slides):
        image_slots = layout.image_slots()
        if intent.needs_exhibit != "chart" or not image_slots:
            continue
        if chart_cursor >= len(chart_blocks):
            continue  # planner wanted a chart but no data supplied -- placeholder stays, flagged at render
        block = chart_blocks[chart_cursor]
        chart_cursor += 1
        try:
            data = ChartData(**block)
        except Exception as e:
            print(f"  FLAGGED slide {idx + 1} ({layout.layout_id}): chart data invalid, skipping chart: {e}", file=sys.stderr)
            continue
        requested = data.type or args.chart_type
        if args.confirm_charts:
            requested = _confirm_chart_type(suggest_chart_type(data), data.title) or requested
        chart_spec = build_chart_spec(data, requested)
        origin = "override" if requested else "suggested"
        print(f"  chart (slide {idx + 1}, {layout.layout_id}): type '{chart_spec.chart_type}' [{origin}]")
        if chart_spec.note:
            print(f"    note: {chart_spec.note}", file=sys.stderr)
        chart_specs_per_slide.setdefault(idx, {})[image_slots[0].shape_name] = chart_spec

    if chart_cursor < len(chart_blocks):
        print(
            f"  WARNING: {len(chart_blocks) - chart_cursor} chart-data block(s) were unused "
            "(more blocks than chart-needing slides)",
            file=sys.stderr,
        )

    # Diagram Builder (5.7, M4): resolve user-supplied step labels into native
    # chevron-flow diagram data, keyed by content-slide index ->
    # {shape_name: DiagramData}. Mapped in planner order to slides the Planner
    # marked needs_exhibit=diagram whose layout has an image slot -- the exact
    # twin of the chart wiring above. Steps are never invented (D9): a
    # diagram-needing slide with no data keeps its placeholder, flagged at render.
    diagram_specs_per_slide = {}
    diagram_blocks = []
    if args.diagram_data:
        try:
            diagram_blocks = _load_diagram_data(args.diagram_data)
        except (OSError, ValueError, json.JSONDecodeError) as e:
            print(f"error: could not read --diagram-data: {e}", file=sys.stderr)
            return 1

    diagram_cursor = 0
    for idx, (layout, spec, intent) in enumerate(content_slides):
        image_slots = layout.image_slots()
        if intent.needs_exhibit != "diagram" or not image_slots:
            continue
        if diagram_cursor >= len(diagram_blocks):
            continue  # planner wanted a diagram but no data supplied -- placeholder stays, flagged at render
        block = diagram_blocks[diagram_cursor]
        diagram_cursor += 1
        try:
            diagram_data = build_diagram_data(**block)
        except Exception as e:
            print(f"  FLAGGED slide {idx + 1} ({layout.layout_id}): diagram data invalid, skipping diagram: {e}", file=sys.stderr)
            continue
        print(f"  diagram (slide {idx + 1}, {layout.layout_id}): chevron-flow, {len(diagram_data.steps)} step(s)")
        diagram_specs_per_slide.setdefault(idx, {})[image_slots[0].shape_name] = diagram_data

    if diagram_cursor < len(diagram_blocks):
        print(
            f"  WARNING: {len(diagram_blocks) - diagram_cursor} diagram-data block(s) were unused "
            "(more blocks than diagram-needing slides)",
            file=sys.stderr,
        )

    # Closing slide: never LLM-generated (real contact info must never be
    # invented). Only fill it if the caller actually gave us something real;
    # otherwise leave the slot unset so the renderer keeps the seed's own
    # placeholder text rather than fabricate a person's details.
    #
    # The contact slot's shape_name is resolved from the manifest by its
    # `role` (a tenant-agnostic contract every closing_contact layout shares),
    # never hardcoded to the reference tenant's own shape name (M7 caught this
    # as a real bug: a second tenant's differently-named slot silently never
    # got filled, because the previous code hardcoded "IM_CLOSING_CONTACT").
    contact_slot = next((s for s in closing_layout.non_image_slots() if s.role == "contact_block"), None)
    if contact_slot is None:
        print(f"error: closing_contact layout has no slot with role='contact_block'", file=sys.stderr)
        return 1
    if args.contact:
        closing_spec = SlideSpec(
            layout_id=CLOSING_LAYOUT_ID,
            content_kind="text",
            slots={contact_slot.shape_name: _normalize_contact(args.contact)},
        )
    else:
        closing_spec = SlideSpec(layout_id=CLOSING_LAYOUT_ID, content_kind="text", slots={})
        print("  WARNING: no --contact given -- closing slide will keep the template's placeholder contact text", file=sys.stderr)

    n_title = 1 if title_spec else 0
    print(
        f"  title: {n_title}/1, content: {len(content_slides)} valid "
        f"({len(content_failed)} failed), closing: 1 (placeholder-or-real)"
    )

    total_slides = n_title + len(content_slides) + 1
    print(f"[4/4] rendering {total_slides} slide(s) via seed-slide duplication...")
    output_prs = assets.open_template()
    n_seed_slides = len(output_prs.slides)
    slide_height_in = manifest.slide_dimensions_in["height"]

    # revisable (5.10, M6): one entry per LLM-generated slide, in output_prs
    # append order, so each entry's position is known up front. The closing
    # slide is deliberately excluded -- its contact info is caller-supplied,
    # never LLM-generated, and must never become something revise() regenerates.
    revisable = []

    if title_spec:
        title_seed = output_prs.slides[title_layout.slide_index]
        _, skipped, title_overflow = render_slide(
            output_prs, title_seed, title_layout, title_spec, slide_height_in=slide_height_in
        )
        _print_render_flags("title slide", title_layout.layout_id, skipped, title_overflow)
        revisable.append({
            "label": "title slide", "position": n_seed_slides + len(revisable), "seed": title_seed, "layout": title_layout,
            "needs_exhibit": None, "render_kwargs": {"slide_height_in": slide_height_in},
            "state": RevisionState(label="title slide", layout=title_layout, role="opening/title slide",
                                   content_brief=brief, current_spec=title_spec),
        })

    for idx, (layout, spec, intent) in enumerate(content_slides):
        seed = output_prs.slides[layout.slide_index]  # always the pristine seed at its fixed index (D13)
        render_kwargs = {
            "chart_specs": chart_specs_per_slide.get(idx),
            "diagram_specs": diagram_specs_per_slide.get(idx),
            "brand_colors": manifest.brand.get("colors", {}),
            "slide_height_in": slide_height_in,
        }
        _, skipped, overflow = render_slide(output_prs, seed, layout, spec, **render_kwargs)
        label = f"slide {idx + 1}"
        _print_render_flags(label, layout.layout_id, skipped, overflow, intent.needs_exhibit)
        revisable.append({
            "label": f"{label} ({layout.layout_id})", "position": n_seed_slides + len(revisable), "seed": seed, "layout": layout,
            "needs_exhibit": intent.needs_exhibit, "render_kwargs": render_kwargs,
            "state": RevisionState(label=label, layout=layout, role=f"content slide ({intent.purpose})",
                                   content_brief=_content_brief(intent), current_spec=spec),
        })

    closing_seed = output_prs.slides[closing_layout.slide_index]
    _, closing_skipped, closing_overflow = render_slide(
        output_prs, closing_seed, closing_layout, closing_spec, slide_height_in=slide_height_in
    )
    _print_render_flags("closing slide", closing_layout.layout_id, closing_skipped, closing_overflow)

    # Strip the pristine seed slides *before* the revision loop (not just at
    # the end): captured `seed` Slide object references stay valid and
    # reusable for further duplication after this (verified) -- doing it here
    # means the deck is already in its final, save-able shape, so the
    # revision loop can preview/save it at any point, not just once at exit.
    strip_seed_slides(output_prs, n_seed_slides)
    for entry in revisable:
        entry["position"] -= n_seed_slides  # every slide's index shifted down once the seeds were removed

    if args.interactive_revise:
        _run_interactive_revision(output_prs, revisable, args.provider, model, deck_context=brief, out_path=args.out)

    print(f"saving to {args.out}...")
    args.out.parent.mkdir(parents=True, exist_ok=True)
    output_prs.save(str(args.out))
    print(f"done: {args.out} ({len(output_prs.slides)} slides)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
