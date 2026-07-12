"""Interactive manual-test console for the Deck Engine.

A thin, friendly wrapper around the real `deck_engine.cli` -- it walks you
through provider/model/brief/chart choices, then builds the exact same argv the
production CLI takes and calls `cli.main()`. That means you are exercising the
actual pipeline (Planner -> Layout Selector -> Content Generator -> Chart
Builder -> Renderer), never a reimplementation of it.

Run it from the repo root:
    python -m deck_engine.manual_test

What it does for you that a raw `python -m deck_engine.cli ...` line doesn't:
  - prompts (hidden) for an API key if the chosen provider needs one and it
    isn't already in the environment -- set for THIS PROCESS ONLY, never written
    to disk or echoed;
  - offers ready-made example briefs + matching chart data (no hand-crafting
    long briefs or JSON), plus a "type your own" and "load from file" path;
  - lets you drive the interactive chart confirm/override loop (--confirm-charts);
  - offers to open the finished .pptx when it's done.

This module is for humans at a terminal; the unattended entry point is still
`deck_engine.cli`.
"""
import getpass
import os
import sys
import tempfile
from pathlib import Path

from . import cli
from .llm_providers import DEFAULT_MODEL, PROVIDERS

# --- ready-made example scenarios ---------------------------------------------
# Each is (label, brief, chart_data_or_None, suggested_min, suggested_max).
# chart_data is a JSON string written to a temp file and passed via --chart-data.

_BRIEF_QUICK = (
    "Introduce our company -- an AI studio that turns messy internal processes into "
    "reliable automated systems. Cover what we do, who we help, and why now. Keep it tight."
)

_BRIEF_CHARTS = (
    "Show that our pilot with a retail client worked. Include a chart of monthly active "
    "users over the first six months (Jan through Jun). Then include a chart breaking down "
    "where the time savings came from (Support, Ops, Finance) as a share of the whole. "
    "Close with the client's quote: 'We stopped firefighting and started planning.'"
)
_CHARTS_CHARTS = """[
  {"categories": ["Jan","Feb","Mar","Apr","May","Jun"],
   "series": {"Monthly Active Users": [120, 260, 430, 690, 980, 1340]},
   "title": "Pilot Adoption: Monthly Active Users"},
  {"categories": ["Support","Ops","Finance"],
   "series": {"Share of Time Saved (%)": [45, 35, 20]},
   "kind": "share",
   "title": "Where the Time Savings Came From"}
]"""

_BRIEF_INTENSE = (
    "Prepare an executive board deck for our firm's proposed AI-driven operational "
    "transformation of Meridian Freight, a mid-market North American logistics carrier. "
    "Build the narrative in order, treating each 'include a chart' as a slide needing an "
    "exhibit chart:\n"
    "1. Section divider: 'The Case for Transformation'.\n"
    "2. Market context: include a chart of total addressable market by year (FY21-FY25).\n"
    "3. Meridian vs field: include a chart comparing Meridian's on-time delivery against the "
    "industry average across 2022, 2023, 2024.\n"
    "4. Cost base: include a chart of operating cost by category (Fuel, Labor, Maintenance, "
    "Admin, Other) as a share of the whole.\n"
    "5. Upside: include a chart of projected revenue by year (Year 1 to Year 5).\n"
    "6. Efficiency: include a chart of efficiency gain by function (Warehousing, Routing, "
    "Dispatch, Billing).\n"
    "7. Revenue mix: include a chart of post-transformation revenue mix (Subscription, "
    "Services, Support) as a share of the whole.\n"
    "8. Implementation roadmap as a process-flow diagram across four phases (Assess, Pilot, "
    "Scale, Optimize).\n"
    "9. Proof point: land the COO quote 'In eight weeks the pilot did what our last vendor "
    "promised in a year.'\n"
    "10. Close with next steps."
)
_CHARTS_INTENSE = """[
  {"categories": ["FY21","FY22","FY23","FY24","FY25"],
   "series": {"TAM ($B)": [18.4, 22.1, 27.9, 34.2, 41.6]},
   "title": "North American Freight-Tech TAM ($B)"},
  {"categories": ["2022","2023","2024"],
   "series": {"Meridian": [88.2, 89.1, 90.4], "Industry Avg": [93.5, 94.2, 95.1]},
   "title": "On-Time Delivery Rate (%): Meridian vs. Industry"},
  {"categories": ["Fuel","Labor","Maintenance","Admin","Other"],
   "series": {"Share of Operating Cost (%)": [34, 41, 12, 8, 5]},
   "kind": "share", "title": "Operating Cost Base by Category"},
  {"categories": ["Year 1","Year 2","Year 3","Year 4","Year 5"],
   "series": {"Projected Revenue ($M)": [412, 468, 551, 642, 758]},
   "title": "Projected Revenue Post-Transformation ($M)", "type": "area"},
  {"categories": ["Warehousing","Routing","Dispatch","Billing"],
   "series": {"Efficiency Gain (%)": [22, 38, 29, 47]},
   "title": "Expected Efficiency Gain by Function", "type": "bar"},
  {"categories": ["Subscription","Services","Support"],
   "series": {"Revenue Mix (%)": [58, 27, 15]},
   "kind": "share", "title": "Post-Transformation Revenue Mix", "type": "doughnut"}
]"""

_EXAMPLES = {
    "1": ("Quick smoke (text only, ~2-3 slides, no charts)", _BRIEF_QUICK, None, 2, 3),
    "2": ("Charts demo (line + pie + quote, ~3-4 slides)", _BRIEF_CHARTS, _CHARTS_CHARTS, 3, 5),
    "3": ("Intense (Meridian board deck, all 6 chart types + diagram degrade)", _BRIEF_INTENSE, _CHARTS_INTENSE, 8, 12),
}

_KEY_ENV = {"anthropic": "ANTHROPIC_API_KEY", "gemini": "GEMINI_API_KEY"}


# --- small input helpers ------------------------------------------------------

def _ask(prompt, default=None):
    suffix = f" [{default}]" if default is not None else ""
    resp = input(f"{prompt}{suffix}: ").strip()
    return resp or (default if default is not None else "")


def _ask_yes_no(prompt, default=False):
    d = "Y/n" if default else "y/N"
    resp = input(f"{prompt} [{d}]: ").strip().lower()
    if not resp:
        return default
    return resp in ("y", "yes")


def _choose(prompt, options, default_key):
    """options: dict[key -> label]. Returns the chosen key."""
    print(prompt)
    for key, label in options.items():
        marker = " (default)" if key == default_key else ""
        print(f"  [{key}] {label}{marker}")
    resp = input(f"choice [{default_key}]: ").strip() or default_key
    while resp not in options:
        resp = input(f"pick one of {list(options)} [{default_key}]: ").strip() or default_key
    return resp


def _ensure_api_key(provider):
    """If the provider needs a key and none is in the environment, prompt for
    it (hidden) and set it for THIS PROCESS ONLY. Returns True if we're good to
    go, False if the user declined to supply a needed key."""
    env_var = _KEY_ENV.get(provider)
    if not env_var:  # ollama needs no key
        return True
    # gemini accepts either GOOGLE_API_KEY or GEMINI_API_KEY
    already = os.environ.get(env_var) or (os.environ.get("GOOGLE_API_KEY") if provider == "gemini" else None)
    if already:
        print(f"  ({env_var} already present in this shell -- using it)")
        return True
    print(f"  {provider} needs an API key. It will be used for this process only -- not saved, not echoed.")
    key = getpass.getpass(f"  paste {env_var} (input hidden, Enter to cancel): ").strip()
    if not key:
        print("  no key given -- cannot use this provider.")
        return False
    os.environ[env_var] = key
    return True


def _collect_brief_and_charts():
    """Returns (brief_text, brief_file, chart_data_path, min_slides, max_slides).
    Exactly one of brief_text/brief_file is set."""
    source = _choose(
        "\nBrief source?",
        {"1": "Built-in example", "2": "Type / paste your own", "3": "Load from a file"},
        "1",
    )
    if source == "1":
        key = _choose(
            "\nWhich example?",
            {k: v[0] for k, v in _EXAMPLES.items()},
            "2",
        )
        _label, brief, chart_json, dmin, dmax = _EXAMPLES[key]
        chart_path = _write_temp_chart_data(chart_json) if chart_json else None
        return brief, None, chart_path, dmin, dmax

    if source == "2":
        print("\nEnter/paste your brief. Finish with a line containing only '.' (a single dot):")
        lines = []
        while True:
            line = input()
            if line.strip() == ".":
                break
            lines.append(line)
        brief = "\n".join(lines).strip()
        chart_path = _maybe_chart_file()
        return brief, None, chart_path, 2, 8

    # source == "3"
    path = _ask("\nPath to brief text file")
    chart_path = _maybe_chart_file()
    return None, path, chart_path, 2, 8


def _write_temp_chart_data(json_str):
    fd, path = tempfile.mkstemp(prefix="deck_charts_", suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(json_str)
    return path


def _maybe_chart_file():
    if not _ask_yes_no("Provide a --chart-data JSON file for chart slides?", default=False):
        return None
    return _ask("  path to chart-data JSON") or None


def _run_once():
    print("\n" + "=" * 68)
    print(" Deck Engine -- Manual Test Console")
    print("=" * 68)

    provider = {"1": "gemini", "2": "anthropic", "3": "ollama"}[
        _choose(
            "\nProvider?",
            {"1": "gemini", "2": "anthropic", "3": "ollama (local, no key)"},
            "1",
        )
    ]
    if not _ensure_api_key(provider):
        return

    model = _ask(f"Model (Enter for default)", DEFAULT_MODEL[provider])

    brief, brief_file, chart_path, dmin, dmax = _collect_brief_and_charts()
    if brief is not None and not brief:
        print("empty brief -- aborting this run.")
        return

    min_slides = _ask("Min content slides", str(dmin))
    max_slides = _ask("Max content slides", str(dmax))

    contact = _ask(
        "Contact line for the closing slide (Enter to skip; use '|' between lines)",
        "Prit Dhanani|AI Engineer|prit.dhanani@example.com",
    )

    chart_type = ""
    confirm_charts = False
    if chart_path:
        chart_type = _ask("Global --chart-type override (Enter for auto-suggest)", "")
        confirm_charts = _ask_yes_no("Confirm/override each chart's type interactively?", default=False)

    out = _ask("Output .pptx path", "output/manual_test.pptx")

    # --- build the exact argv the real CLI takes, then call it -------------
    argv = ["--provider", provider, "--model", model,
            "--min-slides", str(min_slides), "--max-slides", str(max_slides),
            "--out", out]
    if brief is not None:
        argv += ["--brief", brief]
    else:
        argv += ["--brief-file", brief_file]
    if contact:
        argv += ["--contact", contact]
    if chart_path:
        argv += ["--chart-data", chart_path]
    if chart_type:
        argv += ["--chart-type", chart_type]
    if confirm_charts:
        argv += ["--confirm-charts"]

    print("\n--- running deck_engine.cli " + " ".join(
        # don't echo the (already-hidden) key; argv has none anyway
        (a if " " not in a else f'"{a[:40]}..."' if len(a) > 40 else f'"{a}"') for a in argv
    ) + "\n")

    rc = cli.main(argv)

    if chart_path and chart_path.startswith(tempfile.gettempdir()):
        try:
            os.remove(chart_path)
        except OSError:
            pass

    if rc == 0:
        out_path = Path(out)
        print(f"\n[OK] wrote {out_path.resolve()}")
        if sys.platform == "win32" and out_path.exists() and _ask_yes_no("Open it now?", default=True):
            os.startfile(str(out_path.resolve()))  # noqa: only defined on Windows
    else:
        print(f"\n[FAILED] exit code {rc} -- see the error above.")


def main():
    try:
        while True:
            _run_once()
            if not _ask_yes_no("\nRun another?", default=False):
                break
    except (KeyboardInterrupt, EOFError):
        print("\nbye.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
