# Deck Engine

Turn a plain-language brief into a **brand-consistent, editable PowerPoint deck**,
unattended. You describe what the deck is about; the engine plans the slides,
writes the copy, picks the right layouts, builds native charts, and renders a
`.pptx` that opens clean in PowerPoint — on your template, in your fonts and
colors, every time.

It works by **duplicating hand-built "seed slides"** from a template and filling
their named shapes, so the model never touches design — it only supplies text and
data inside tight, template-defined limits. That division of labor is what keeps
output good (see [Why the output stays good](#why-the-output-stays-good)).

> **Status:** M0–M4 built and verified end-to-end (Anthropic, Gemini, and local
> Ollama). QA/overflow, the revision engine, and multi-tenant
> onboarding are on the roadmap — see [Status & roadmap](#status--roadmap).
> The full design rationale lives in
> [DOCS/02_system_design_document.md](DOCS/02_system_design_document.md).

---

## Requirements

- **Python 3.13** (tested on 3.13.14)
- Three dependencies only — `pip install -r requirements.txt`:
  - `python-pptx` (rendering + native charts)
  - `anthropic` (only needed if you use `--provider anthropic`)
  - `pydantic` (validation)
- Ollama and Gemini are reached over plain stdlib HTTP (`urllib`), so they add **no
  extra dependency**.

An API key is needed only for hosted providers (Anthropic, Gemini). Local Ollama
needs none.

---

## Quickstart (interactive)

The easiest way to try it — an interactive console that walks you through every
choice and offers ready-made example briefs:

```
python -m deck_engine.manual_test
```

It prompts for provider + key, a brief (built-in example / your own / a file),
optional chart data, slide count, and where to save — then runs the real pipeline
and offers to open the result. Nothing is written to disk except the deck; API
keys you paste are held for that one process only.

---

## Unattended CLI

For scripted / repeatable runs:

```
python -m deck_engine.cli --brief "..." --out output/deck.pptx
```

### Flags

| Flag | Required | Default | Purpose |
|---|---|---|---|
| `--brief` | one of these | — | Brief text, inline |
| `--brief-file` | one of these | — | Path to a text file with the brief |
| `--out` | ✅ | — | Output `.pptx` path |
| `--contact` | | — | Closing-slide contact block, `\|`-separated (see below) |
| `--provider` | | `anthropic` | `anthropic` \| `gemini` \| `ollama` |
| `--model` | | provider default | Override the model |
| `--min-slides` | | `2` | Min **content** slides (excludes title/closing) |
| `--max-slides` | | `8` | Max **content** slides |
| `--chart-data` | | — | JSON file of chart data (see [Chart data](#chart-data)) |
| `--chart-type` | | — | Global chart-type override for blocks lacking their own `type` |
| `--confirm-charts` | | off | Prompt to confirm/override each chart's type |
| `--diagram-data` | | — | JSON file of diagram steps (see [Diagram data](#diagram-data)) |
| `--tenant` | | `default` | Which template + manifest to use |

### Examples

```bash
# Simple text deck via the default provider (Anthropic)
python -m deck_engine.cli \
  --brief "Introduce our AI studio: what we do, who we help, why now." \
  --contact "Prit Dhanani|AI Engineer|prit.dhanani@example.com" \
  --out output/intro.pptx

# Chart-heavy board deck via Gemini
python -m deck_engine.cli --provider gemini \
  --brief-file output/M3_intense_brief.txt \
  --chart-data output/M3_intense_charts.json \
  --min-slides 8 --max-slides 12 \
  --contact "Prit Dhanani|AI Engineer|prit.dhanani@example.com" \
  --out output/board_deck.pptx

# Deck with a process-flow diagram (steps you supply; see "Diagram data")
python -m deck_engine.cli \
  --brief "Explain our client onboarding process end to end." \
  --diagram-data output/onboarding_steps.json \
  --out output/onboarding.pptx
```

---

## Providers & API keys

| Provider | Key (env var) | Default model | Notes |
|---|---|---|---|
| `anthropic` | `ANTHROPIC_API_KEY` | `claude-sonnet-5` | Reference/validated provider |
| `gemini` | `GEMINI_API_KEY` or `GOOGLE_API_KEY` | `gemini-3.1-flash-lite` | Reliable for large multi-chart decks |
| `ollama` | *(none)* | `gemma4:26b-a4b-it-q4_K_M` | Local. Set `OLLAMA_HOST` (default `http://localhost:11434`) and `OLLAMA_TIMEOUT` (default `240`s) |

Set the key in your shell before running, e.g. on Windows PowerShell:
`$env:GEMINI_API_KEY = "..."`. The CLI fails fast with a clear message if the
chosen provider's key is missing — it never starts generating and then dies
halfway.

---

## What you must supply (and what the engine won't invent)

The engine writes narrative and picks layouts, but it **never fabricates specific
facts** ("flag, don't guess"). Two things must come from you:

- **Contact details (`--contact`).** Real names/emails/phones. Format uses `|`
  between lines: `"Jane Doe|Senior Partner|jane@firm.com"`. If you omit it, the
  closing slide keeps the template's placeholder and prints a warning — it will
  not make up a person.
- **Chart data (`--chart-data`).** If a slide needs a chart and you didn't supply
  its numbers, that slide renders a grey placeholder flagged `needs_review` — it
  will not invent figures. Charts are only as real as the data you provide.
- **Diagram steps (`--diagram-data`).** Same rule for process-flow diagrams: if a
  slide needs one and you didn't supply its steps, the placeholder stays and the
  slide is flagged — the engine never invents the steps of your process.

Everything else — how many slides, which layouts, the copy, the chart *type* — the
engine decides.

---

## Chart data

A JSON **list of blocks**, mapped in planner order to the slides the planner marked
as needing a chart. Each block:

```json
[
  {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": {"Revenue ($M)": [2.1, 2.5, 2.9, 3.4]},
    "title": "Quarterly Revenue",
    "kind": "share",
    "type": "line"
  }
]
```

| Field | Required | Meaning |
|---|---|---|
| `categories` | ✅ | X-axis labels |
| `series` | ✅ | `name -> [one value per category]`; multiple series = grouped comparison |
| `title` | | Chart title |
| `kind` | | `"share"` hints a part-of-whole chart (→ suggests pie) |
| `type` | | Force this chart's type (see below); overrides the suggestion |

**Supported chart types:** `column`, `bar`, `line`, `pie`, `doughnut`, `area`.
The type is **suggested automatically from the data shape** (multiple series →
column, single time series → line, a `share` → pie, else → column). You can
override per-block (`type`) or globally (`--chart-type`). An unsupported/exotic
request (waterfall, funnel, …) is remapped to the nearest supported type with a
printed note — never a broken chart. (Scatter/bubble/stock are deferred — they
need a different data model.)

---

## Diagram data

A JSON **list of blocks**, mapped in planner order to the slides the planner marked
as needing a diagram (exactly like chart data). Each block:

```json
[
  { "steps": ["Intake", "Review", "Approve", "Deliver"] }
]
```

Rendered as a **native, editable chevron-flow**: one arrow-shaped chevron per step,
in your template's brand color, grouped into a single named shape (each chevron
stays individually editable — it's real shapes, not a picture). Constraints:
**2–6 steps**, each label **≤ 24 characters** (so it stays legible inside a
chevron); an invalid block is flagged and skipped, never a crash. v1 is
linear-only — branching flows are deferred until something smarter (you, or a
future agent) can decide where a flow splits.

---

## Why the output stays good

Quality comes from rails, not from one clever prompt:

1. **The template does the design.** Every slide is a clone of a hand-built seed
   slide; fonts/colors/positioning are inherited. The model can't make it ugly.
2. **Per-slot limits** (`max_chars`, `max_items`, allowed values) from the
   template manifest are injected into the prompt **and** enforced as a JSON
   schema **and** re-checked after generation.
3. **"Omit, don't invent."** The generator is told to leave a slot empty rather
   than fabricate detail the brief doesn't support.
4. **One bounded retry** feeds exact validation errors back before flagging a
   slide `needs_review`.
5. **Deterministic choices** for layout selection and chart type — the LLM only
   hints; the engine validates against the real template.

---

## Output & verification

The deck is written to `--out`. To sanity-check a run beyond opening it:

- Re-open programmatically: `python -c "from pptx import Presentation;
  p=Presentation('output/deck.pptx'); print(len(p.slides))"`
- `output/validate_m1.ps1` opens a deck in real PowerPoint via COM and exports
  per-slide PNGs (repoint `$outPath`/`$shotDir` at the top). A clean open with no
  repair prompt means the XML — including native charts — is valid.

---

## Project layout

```
deck_engine/
  cli.py               unattended entry point (brief -> .pptx)
  manual_test.py       interactive test console (wraps cli)
  config.py            tenant registry (template + manifest paths)
  template_registry.py loads + validates a tenant's template against its manifest
  deck_planner.py      5.3 brief -> ordered slide intents
  layout_selector.py   5.4 intent -> concrete template layout (deterministic)
  content_generator.py 5.5 fills one slide's slots (structured output + retry)
  chart_builder.py     5.6 data shape -> chart type -> native chart spec
  diagram_builder.py   5.7 ordered steps -> chevron-flow diagram data
  renderer.py          5.8 seed-slide duplication + fill + native charts/diagrams
  llm_providers.py     provider dispatch (Anthropic / Ollama / Gemini)
  models/              Pydantic contracts: manifest, slide_intent, slide_spec, chart
Templates/             the .pptx template + its JSON manifest
tests/                 stdlib unittest suite
DOCS/                  design docs (spike plan, SDD, PRD)
```

---

## Testing

```
python -m unittest discover -s tests
```

80 tests, no external services required (all LLM/network calls are mocked).

---

## Status & roadmap

| Milestone | State |
|---|---|
| M0 Foundations (config, template+manifest loader) | ✅ done |
| M1 Core loop (brief → text deck) | ✅ done |
| M2 Deck Planner + Layout Selector (8 layouts) | ✅ done |
| M3 Chart Builder (native, editable, D7 types) | ✅ done |
| M4 Diagram Builder (chevron process-flow, native, editable) | ✅ done |
| M5 QA / overflow checker | ⏳ next |
| M6 Revision engine (field edit / regen) | ⏳ planned |
| M7 Second tenant (prove tenant-agnosticism) | ⏳ planned |

### Known limitations (today)

- **Diagrams are linear-only (M4 scope).** A process-flow diagram is a straight
  2–6-step chevron chain from steps you supply via `--diagram-data`; branching
  flows and prose-to-diagram extraction are deferred (they need a judgement call
  the deterministic v1 deliberately doesn't make). A diagram-needing slide with
  no data still renders a flagged placeholder — gracefully, never a crash.
- **No tenant voice guide / cross-slide context yet.** Each content slide is
  written from its own outline; house-tone and neighboring-headline context
  (designed in SDD §5.5) aren't wired in. This is the main lever left for
  sharper copy — confirmed in dogfood runs against local Ollama models, where
  content read as generic/formulaic without house-tone grounding. Likely v2
  territory once an agent can carry richer context than a single-slide prompt.
- **Diagram chevrons are a single flat brand color (M4 v1 styling).** All
  chevrons in a flow currently render in the same solid `blue_primary` with no
  shading/gradient progression — functional and on-brand, but visually flat
  compared to a hand-designed flow diagram. A styling pass (alternating shades,
  a gradient across steps) is a known follow-up, not yet scheduled against a
  milestone.
- **Local Ollama on large schemas.** Big multi-chart briefs can truncate JSON on
  a CPU-only local model; Gemini and Anthropic are the proven paths for those.
  Reasoning-capable local models (qwen3.x, gemma4, ...) additionally need
  `think: false` forced on every call (fixed in `llm_providers.py`) since they'll
  otherwise burn the whole token budget on a hidden reasoning trace before
  writing any real content.

For the full "why" behind every decision, read
[DOCS/02_system_design_document.md](DOCS/02_system_design_document.md).
