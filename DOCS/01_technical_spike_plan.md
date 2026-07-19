# Technical Feasibility Spike Plan
### Presentation Automation Engine ("Deck Engine")

| | |
|---|---|
| **Document** | 1 of 3 — Technical Spike Plan |
| **Version** | 1.0 |
| **Status** | Draft for execution |
| **Owner** | Prit (AI Engineer) |
| **Related docs** | `02_system_design_document.md`, `03_product_requirements_document.md` |
| **Precedes** | System Design Document (findings from this spike feed the SDD) |

---

## 1. Purpose

Before writing architecture or production code, this spike **de-risks the technical bets the entire system stands on**. Each bet, if false, changes the architecture — so we prove or disprove them cheaply first, on throwaway code, against quantitative pass/fail criteria.

**Time-box:** 2–4 focused days.
**Output:** a short findings memo (Section 7) that either greenlights the build as designed, or names the specific architectural change a failed bet forces.
**This is not production code.** Inputs are hardcoded, no UI, no persistence, no cleanup. The code is disposable; the *findings* are the deliverable.

> **Why a spike at all:** the most common failure mode for deck-automation projects is discovering three sprints in that (a) the LLM can't reliably emit the structured spec the renderer needs, or (b) the rendering library can't do what was assumed. A 2–4 day spike on those two bets prevents unwinding design decisions built on a false premise.

---

## 2. The bets we are de-risking

The system rests on five load-bearing assumptions. Everything downstream (data model, agent design, revision loop) assumes these hold.

| # | Assumption | If false… |
|---|---|---|
| A1 | A `.pptx` built to spec exposes named placeholders that `python-pptx` can reliably target, and filling them preserves brand fidelity. | The template→render contract collapses; the whole approach is non-viable without a different rendering library. |
| A2 | Claude can emit a valid, complete **slide-spec** (structured JSON) from a plain brief, reliably enough to run unattended. | Need validation + retry scaffolding, a narrower schema, or a human-in-loop at generation time. |
| A3 | Over-length generated content (text overflow) can be **reliably detected** before export. | Silent broken slides reach the user; need a heavier pre-generation constraint model. |
| A4 | Native `python-pptx` charts inherit the template theme and remain editable. | Charts must become rendered images (not editable) or a shape-composition subsystem. |
| A5 | A useful **process/flow diagram** can be generated as *editable native shapes*, not a flat image. | Diagrams in v1 fall back to rendered images (informational but non-editable) or are deferred. |

---

## 3. Spike experiments

Each experiment is independent and can be run in isolation. Run S1 and S2 first — they are the two bets most likely to kill the approach.

### S1 — Template fill fidelity  *(de-risks A1)* — ✅ **Structural check run; confirmed**

**Question:** Can we open a real branded `.pptx`, add slides from its existing layouts, address every placeholder by a stable name, fill them, and get a clean, on-brand `.pptx` back?

**Status update:** ran a structural cross-check against the real `MasterDeck.pptx` + `MasterDeck_manifest.json`. Result: **all 8 manifest layouts' required shape names are present, by exact match, on the correct `slide_index`.** File has 9 real slides (slide 8 is a human-readable index, unreferenced by the manifest — correctly ignored). This confirms the *addressing* half of S1. **Still to run:** the actual fill-and-open round-trip (steps 2–4 below) — duplicate each seed slide, write real content into its named shapes, save, and open in PowerPoint to visually confirm brand fidelity and no repair prompt.

**Method (remaining):**
1. ~~Take a `.pptx` built to spec~~ — **done**, real file in hand.
2. ~~Enumerate layouts/shapes~~ — **done** (see status above).
3. For each of the 8 seed slides: duplicate it (XML-clone, per D13's duplicate-first ordering), fill its named shapes with sample content, and — for slots typed `image` — insert a real picture and remove the grey placeholder rect.
4. Save as a new file; open in PowerPoint (or LibreOffice) and inspect.

**Key finding already confirmed:** this template does **not** use distinct PowerPoint Slide Layouts per slide type — all 9 slides sit on the generic `Blank` layout, with per-type structure carried entirely in hand-named shapes. The rendering mechanic is therefore **seed-slide duplication + shape-name addressing** (see SDD D3), not `add_slide(slide_layout)` + `placeholders`. The SDD has been updated accordingly.

**Pass criteria:**
- 100% of the template's placeholders are addressable by a **stable** identifier (name or idx that doesn't shift between runs).
- Output opens with **no repair prompt** and theme (colors, fonts, logo, footer) intact.
- Text written into a placeholder inherits the template's font/size/color (no manual restyle needed).

**If it fails:** evaluate an alternate library (e.g., a higher-level engine) or move rendering to a document-conversion pipeline. This is the highest-severity failure — surface it loudly.

---

### S2 — Structured slide-spec reliability  *(de-risks A2)*

**Question:** Using structured outputs, how reliably does Claude produce a schema-valid, *usable* slide-spec from a varied brief?

**Method:**
1. Define a minimal `SlideSpec` JSON schema (layout choice + placeholder values, with `enum` for layout and `maxLength` per text field).
2. Call the API with **tool-use forcing** (`tool_choice={"type":"tool","name":...}`) or native structured outputs (JSON-schema mode) — both are GA on current models.
3. Run **N = 100** varied briefs (short, long, messy, multi-point, comparison, data-bearing).
4. Validate every output with **Pydantic**. Log schema-valid rate. Human-rate a 20-sample for "usable first draft (yes/no)".

**Reference model choice:** reasoning/generation → `claude-sonnet-5` or `claude-opus-4-8`; cheap classification (layout hinting, chart-type suggestion) → `claude-haiku-4-5`.

**Pass criteria:**
- **≥ 99%** schema-valid on first attempt (industry-reported tool-use validity is ~99.8%).
- **≥ 90%** of the 20-sample rated "usable first draft."
- Bounded retry (feed validation error back once) closes any remaining gap to ~100%.

**If it fails:** narrow/flatten the schema (deep nesting >3–4 levels raises failure rates), tighten field descriptions (they are part of the prompt), and add the validate→retry loop as a hard requirement in the SDD.

---

### S3 — Overflow detection  *(de-risks A3)*

**Question:** When content is too long for its placeholder, can we reliably detect it before export?

**Method:**
1. Deliberately feed over-length content into placeholders.
2. Test two detectors: (a) **prevention** — `maxLength` per field in the schema keyed to the manifest's char-limits; (b) **post-render measurement** — estimate rendered text height vs. placeholder height (font size × wrapped line count vs. box height in EMU).
3. Measure detection rate across ~30 deliberately-overflowing cases.

**Pass criteria:**
- A detection method flags **≥ 95%** of overflow cases.
- (Auto-*fixing* is out of scope for the spike — we only need reliable *detection* so the SDD can route overflow to a `needs_review` state.)

**If it fails:** rely more heavily on schema char-caps as prevention and accept a manual review gate.

---

### S4 — Native chart rendering into theme  *(de-risks A4)*

**Question:** Do native `python-pptx` charts adopt the template's theme accent colors and stay editable?

**Method:**
1. On a template slide, add `COLUMN_CLUSTERED`, `LINE`, and `PIE` charts via `add_chart` with `CategoryChartData`.
2. Open in PowerPoint: confirm charts are **editable objects** (right-click → Edit Data works), colors come from theme Accent 1–6, values are correct.

**Supported-type matrix to confirm** (from `XL_CHART_TYPE`): native support exists for **column/bar (clustered/stacked), line, pie, doughnut, area, scatter, radar, bubble, stock**. **No native support** for **waterfall, funnel, marimekko/mekko, treemap, sunburst, bullet, Gantt, heatmap** — exactly the exotic consulting charts. Confirm this so the SDD scopes the suggestion engine to supported types only.

**Pass criteria:**
- The three standard charts render editable, theme-colored, data-correct.

**If it fails:** charts become rendered images (matplotlib → PNG → picture) — informational but not editable — or a deferred shape-composition component.

---

### S5 — Editable diagram generation  *(de-risks A5)*

**Question:** Can we generate a **process/flow diagram as editable native shapes** (boxes + connectors) from an LLM-produced node/edge spec, laid out by a deterministic algorithm?

**Method:**
1. LLM emits a structured node/edge spec (Mermaid-like: nodes with labels, directed edges) from a plain-text process description.
2. A deterministic layout function positions nodes (linear top-down or left-right; single-branch fork) and draws `MSO_SHAPE` rectangles + connectors via `python-pptx`.
3. Test one **linear 4–6 step** flow and one **single-branch** flow. Open in PowerPoint.

**Grounding:** Mermaid has **no native `.pptx` export** (open request since 2022); rendering Mermaid to PNG/SVG yields a *flat image* (not editable). Editable output therefore requires native-shape composition + our own layout logic. S5 proves whether a *constrained* version of this is tractable for v1.

**Pass criteria:**
- Both diagrams render **legibly** (no overlapping shapes/connectors), are **editable shape-by-shape**, and match theme colors.

**If it fails / is too costly for v1:** diagrams fall back to **rendered image + explicit `non-editable` flag** for v1, and editable-native-shape diagrams move to v2. (Charts still stay native per S4 — diagrams are the harder case.)

---

### S6 — End-to-end thin thread  *(optional, integrative)*

**Question:** Can we chain S1+S2+S4 into: brief → 3-slide deck (title + one content + one chart), fully unattended?

**Pass criteria:** produces an openable, on-brand 3-slide `.pptx` with zero manual steps. This is the smallest proof the pipeline concept holds end-to-end; it becomes the seed of the M1 vertical slice in the SDD.

---

## 4. Explicitly out of scope for the spike

No UI, no auth, no multi-tenant isolation, no persistence/database, no revision loop, no feedback log, no exotic charts, no arbitrary-template ingestion, no manifest-authoring tool, no confidentiality/compliance controls, no cost optimization. Hardcode all inputs. Delete the repo after the findings memo is written.

---

## 5. Environment & inputs

**Runtime:** Python 3.11+, `python-pptx`, `anthropic` SDK, `pydantic`. Optional for S3/S5: a vision-model call; `mermaid-cli`/`graphviz` **only** if testing the image-fallback path.

**Inputs required before starting:**
1. **One `.pptx` template built to spec** — ~6 named layouts with clearly named placeholders. *This is the single most important input; S1's validity depends on it being a real designed template, not a default theme.* If not ready, generate a stand-in to unblock, but flag results as provisional until re-run against the real file.
2. **API access** to current models (`claude-sonnet-5` / `claude-opus-4-8` / `claude-haiku-4-5`).
3. A handful of **representative briefs** (real anonymized engagement inputs are ideal for S2's quality sample).

---

## 6. Deliverables

1. A **throwaway repo** with one script per experiment (`s1_template_fill.py`, `s2_spec_reliability.py`, …).
2. A **1–2 page findings memo** mapping each experiment to **Pass / Adjust / Fail**, with the evidence (screenshots, validity %, sample ratings) and a one-line recommendation each.
3. A **recommendation** feeding the SDD: confirm the architecture as designed, or specify the exact change each failed bet forces.

---

## 7. Decision gate

| Experiment | Pass → | Fail → (architectural consequence) |
|---|---|---|
| **S1** template fill | Proceed; `.pptx`-as-template confirmed | Re-evaluate rendering library / conversion pipeline. **Blocking.** |
| **S2** spec reliability | Structured-output engine confirmed | Add mandatory validate→retry; flatten schema; possibly human-in-loop at generation. |
| **S3** overflow detection | Route overflow → `needs_review` state | Lean on schema char-caps + manual gate. |
| **S4** charts | Native charts in v1 | Charts become rendered images (non-editable) or deferred. |
| **S5** diagrams | Editable native-shape diagrams in v1 | Diagrams become image+flag in v1; editable diagrams → v2. |
| **S6** e2e (optional) | Seed the M1 vertical slice | Isolate which stage broke; re-run that experiment. |

**Green light rule:** S1 and S2 **must** pass to proceed with this architecture. S3–S5 failures are *survivable* — they narrow v1 scope rather than kill it.

---

## 8. Assumptions carried into this spike

- The template can and will be authored to spec (named placeholders) — the spike does **not** attempt to ingest arbitrary user templates.
- Deck sizes are modest (≤ ~40 slides), so `python-pptx` memory behavior (which degrades past ~500–1000 slides) is not a concern at this stage.
- "Presentations" means client-facing consulting/PE decks where **native editable `.pptx`** is a hard requirement (rules out flat-image-only output).
