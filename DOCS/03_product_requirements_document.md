# Product Requirements Document
### Presentation Automation Engine ("Deck Engine")

| | |
|---|---|
| **Document** | 3 of 3 — Product Requirements Document (PRD) |
| **Version** | 1.0 |
| **Status** | Draft for sign-off |
| **Owner** | Prit (AI Engineer) |
| **Related docs** | `01_technical_spike_plan.md`, `02_system_design_document.md` |
| **Purpose** | The authoritative statement of *what* v1 delivers and how "done" is judged. The SDD covers *how*. |

---

## 1. Problem statement

Consulting and PE-facing deck production splits into two very different kinds of work. The **reasoning** — turning messy findings into a defensible story — is expensive and non-delegable. The **production** — brand-compliant layout, formatting, and exhibits — is low-leverage but consumes the majority of the hours: industry write-ups estimate the production layer (formatting, alignment, brand compliance, QA) at roughly **60–70% of deck-building time**, and note that generic AI tools help with content generation but leave that production layer largely manual, because it is a *production* problem (deterministic, at the file's object-model level), not a *generation* problem.

The market has effectively validated where the value sits: everyone can bolt an LLM onto an outline; the defensible, paid-for capability is **deterministic, brand-perfect rendering into an editable native `.pptx`**. That is precisely the seam this product is built around.

---

## 2. Vision & product framing

A **tenant-agnostic engine** that renders any tenant's branded template. Branding lives entirely in a swappable **template + manifest** — never in the code. A firm adopts the product by having a built-to-spec template + manifest prepared, then generates decks from briefs. The reusable, sellable asset is the **engine**; per-firm **template onboarding** is a productized setup step.

**Explicitly not** a "type a prompt, get a deck" flat-output generator, and **not** (in v1) an "upload any existing deck and we'll figure it out" tool.

---

## 3. Goals & success metrics

| Goal | Metric | v1 target |
|---|---|---|
| Cut deck production time | Time from brief → usable first draft | ≥ 50% reduction vs. manual, on a real dogfood deck |
| Usable first drafts | % slides rated "usable first draft" (human sample) | ≥ 90% |
| Brand fidelity | Off-brand slides in output (wrong color/font/layout) | **0** (structurally guaranteed by template inheritance) |
| Structured reliability | SlideSpec schema-valid rate (post-retry) | ~100% |
| No silent failures | Broken/overflowing slides shipped un-flagged | **0** |
| Regen effectiveness | Slides resolved within ≤ 3 revision attempts | ≥ 90% |

---

## 4. Users & personas

| Persona | Who | Needs |
|---|---|---|
| **Consultant (end user)** | Producer of the deck | Brief in → editable on-brand `.pptx` out; fast per-slide fixes; native file they can finesse before a client meeting. |
| **Template admin / onboarder** | Prit (or a templating step) | A clear, one-time way to register a tenant's `.pptx` + author its manifest and voice guide. |
| **Firm / tenant (buyer)** | The consultancy adopting it | Their brand, not a generic look; isolation of their content; output indistinguishable from hand-built decks. |

---

## 5. Scope

### In scope (v1)
- Brief → editable native `.pptx` on a tenant's built-to-spec template.
- Deck planning (narrative + ordered slide intents) and **deterministic layout selection** from the manifest.
- **Text** generation in the tenant's voice, constrained to layout placeholders.
- **Charts** with data-shape-based **type suggestion + user confirmation**, rendered as **native** charts (supported subset only).
- **One diagram type** — process/flow — as editable native shapes *(with image fallback if S5 is deferred)*.
- **Per-slide revision:** field-level edit + full-slide regeneration; accumulating (capped) feedback log; escalation after 3 attempts.
- **Multi-tenant-ready** architecture: per-tenant template + manifest + voice guide; isolated state.
- **Needs-review** flagging for overflow/failed slides.
- Zero-data-retention API calls as default.

### Out of scope (v1) → see Roadmap §9
- Arbitrary-template ingestion / auto-manifest generation.
- Self-serve template-builder wizard.
- Exotic consulting charts (waterfall, mekko, funnel, Gantt, treemap).
- Decorative/filler images (by design — content is only informational).
- Confidentiality/compliance controls beyond ZDR default.
- Real-time collaboration, animations, embedded video, diff-view UI.

---

## 6. Functional requirements

Each requirement has acceptance criteria. "The system" = the Deck Engine.

**FR-1 — Generate deck from brief**
Given a brief and a tenant, the system produces an ordered, multi-slide editable `.pptx`.
*Accept:* output opens in PowerPoint with no repair prompt; slide order matches the planned narrative; theme intact.

**FR-2 — Automatic layout selection**
Each slide's layout is chosen from the tenant's manifest based on the slide's purpose.
*Accept:* every slide uses a valid manifest layout; no free-floating/unstructured slides; ambiguous cases default safely and are flagged.

**FR-3 — In-voice text generation**
Text is generated in the tenant's configured voice and fits the layout's placeholders.
*Accept:* SlideSpec is schema-valid; no field exceeds its manifest `max_chars`; a human sample rates ≥ 90% "usable first draft."

**FR-4 — Charts with suggestion + confirmation**
For data-bearing slides, the system suggests a chart type from data shape, lets the user confirm/override, and renders a **native** chart.
*Accept:* suggestion matches the data shape (categorical→bar, part-of-whole→pie, time-series→line); only supported types are offered; chart is an editable object with theme colors and correct data; unsupported requests get a nearest-supported suggestion + explanation; missing data is requested, never fabricated.

**FR-5 — Process/flow diagram**
The system builds a process/flow diagram from plain text (user does **not** write node/edge syntax).
*Accept:* a 4–6 step linear flow and a single-branch flow render legibly, theme-colored, and **editable shape-by-shape**; if editable generation is deferred (S5), output is a rendered image explicitly flagged `non-editable`.
*Status (M4, SDD v1.6):* **linear flows: met** — 2–6 step chevron flows render brand-colored and editable shape-by-shape (native grouped autoshapes; no image fallback was needed, so the S5 deferral clause never triggered). Two parts deliberately narrowed, not silently dropped: **input is ordered step labels** (`--diagram-data`), not free prose — prose-to-steps extraction is deferred to v2 (an LLM/agent job feeding the same contract); **single-branch flows are deferred** to v2 — a branch point needs a judgement call the deterministic v1 diagram path intentionally doesn't make.

**FR-6 — Per-slide flag & feedback**
The user can flag any single slide and attach free-text feedback without affecting other slides.
*Accept:* only the flagged slide changes; all others remain byte-stable/locked.

**FR-7 — Field-edit vs. regeneration routing**
Feedback is classified: small asks → surgical field edit; structural asks → full-slide regeneration.
*Accept:* "change 12% to 15%" edits only that value and changes nothing else; "re-approach this slide" regenerates it; the regen prompt includes prior rejected attempts and deck context.

**FR-8 — Structural style consistency on regeneration**
A regenerated slide is always consistent with the deck's brand (colors/fonts/layout).
*Accept:* regeneration reuses the same layout-fill path as generation; no regenerated slide is ever off-brand (guaranteed by template inheritance, not by prompting).

**FR-9 — Editable native export**
Output is a single native `.pptx` a consultant can finesse.
*Accept:* text is editable in placeholders; charts editable via Edit Data; diagrams (when native) editable per shape; no flattened/"AI-generated" artifacts.

**FR-10 — Needs-review flagging**
Overflowing or failed slides are surfaced, never silently shipped.
*Accept:* overflow is detected (≥ 95% of cases) and marked `needs_review`; slides failing after retry are marked `failed`/`needs_review`; nothing broken ships un-flagged.

**FR-11 — Template + manifest onboarding**
A tenant is onboarded by registering a built-to-spec `.pptx` and its manifest (+ voice guide).
*Accept:* the manifest is validated against the real `.pptx` at load (mismatch = fatal, pre-render); end users never author or see the manifest.

**FR-12 — Multi-tenant isolation**
Each tenant's templates, content, and outputs are isolated.
*Accept:* a second onboarded tenant renders on its own brand with no leakage of content or templates between tenants.

---

## 7. Non-functional requirements

- **Output format:** native, fully-editable `.pptx` — hard requirement (rules out flat-image-only output).
- **Brand fidelity:** zero off-brand slides — enforced structurally via template inheritance.
- **Reliability:** SlideSpec ~100% schema-valid after bounded retry; no silent broken slides.
- **Content honesty:** no decorative filler; every element is informational (text, chart, diagram) or a user-supplied image.
- **Latency:** target a full small deck (~10–15 slides) generated in a few minutes unattended (tune post-spike).
- **Tenancy:** isolation enforced at the data layer.
- **Data handling:** ZDR API calls by default (full confidentiality controls deferred to v2).
- **Scale assumption:** decks ≤ ~40 slides (beyond which `python-pptx` memory degrades — out of v1 scope).

---

## 8. Definition of Done (v1)

v1 is done when **all** of the following hold:

- [ ] Brief → editable `.pptx` on a built-to-spec template, unattended (FR-1).
- [ ] Full ~6-layout set with deterministic manifest-driven selection (FR-2).
- [ ] In-voice text within placeholder limits, ≥ 90% usable first draft (FR-3).
- [ ] Native charts with suggestion + confirmation, supported subset only (FR-4).
- [ ] One process/flow diagram type — editable native shapes, or image+flag if S5 deferred (FR-5).
- [ ] Per-slide flag + field-edit/regen routing + capped feedback log + 3-attempt escalation (FR-6/7).
- [ ] Regeneration structurally on-brand via shared fill path (FR-8).
- [ ] Editable native `.pptx` export (FR-9).
- [ ] Overflow/failure detection → `needs_review`, zero silent bad slides (FR-10).
- [ ] Template + manifest onboarding with load-time validation (FR-11).
- [ ] Second tenant onboarded, isolation proven (FR-12).
- [ ] Dogfooded on a real deck.

Anything beyond this list is **out of v1** and belongs to the roadmap.

---

## 9. Roadmap (post-v1)

**v2 — Productization & depth**
- Guided **template-builder wizard** (upload logo/colors → pick from N layout shapes → generate a conforming `.pptx` + draft manifest). Still not "accept any file."
- **Exotic charts** as editable **shape-composition** components (waterfall, mekko, funnel, Gantt, RAG scorecards, roadmaps).
- **Editable complex diagrams** (org charts, swimlanes) if deferred from v1.
- **Confidentiality/compliance** workstream (retention, storage, access controls) — the prerequisite before non-dogfood client data flows through the system at scale.
- **Diff-view** on regeneration (old vs. new before replace).

**v3 — Arbitrary-template ingestion**
- Parse an existing `.pptx` and **auto-draft** its manifest for human review/correction (the hardest problem — deliberately last, funded by v1/v2 revenue).
- think-cell-class exhibit fidelity.

---

## 10. Assumptions, constraints, dependencies

**Assumptions**
- "Presentations" = client-facing consulting/PE decks requiring editable `.pptx`.
- Every tenant template is built to spec (named placeholders) — no arbitrary ingestion in v1.
- Decks ≤ ~40 slides.
- ZDR API default is acceptable as the sole confidentiality measure in v1.

**Constraints**
- `python-pptx`: no template/`.potx` concept (use `.pptx`); cannot add layouts at runtime; cannot write into table placeholders (use `add_table`); no native exotic chart types; embedded-image loss on slide-copy.
- Editable diagrams require native-shape composition (Mermaid→image is non-editable).

**Dependencies**
- A real built-to-spec `.pptx` template + manifest + voice guide per tenant (blocking input — parallels §12 of the SDD).
- API access to current models (`claude-sonnet-5` / `claude-opus-4-8` / `claude-haiku-4-5`) with structured outputs.
- Spike findings (`01_technical_spike_plan.md`) resolving the spike-gated decisions.

---

## 11. Open decisions (owner: Prit)

1. Confirm the **exact v1 layout list** (names + purposes).
2. Confirm the **v1 chart subset**; confirm exotic charts are v2.
3. Decide whether the **diagram image-fallback** is acceptable for v1 or diagrams must be editable-or-deferred.
4. Choose **storage/hosting** (local MVP vs. cloud from day one).
5. Define where the **tenant voice guide** lives and who authors it.
6. Confirm **who authors the manifest** per new tenant in practice.

---

*This PRD is deliberately scoped to a buildable v1 that proves the reasoning-and-rendering engine on a controlled template. Multi-tenant arbitrary-upload is the productization phase — a real, harder project that begins once the core engine is validated, not before.*
