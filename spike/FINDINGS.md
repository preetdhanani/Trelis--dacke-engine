# Spike Findings Memo — S1 & S2

**Scope:** this executes Doc 1 (`01_technical_spike_plan.md`) experiments S1 and S2 only, per instruction. Throwaway code lives in `/spike`; nothing here is production structure.

---

## S1 — Template fill fidelity → **PASS**

**Question:** duplicate → fill → save → open round-trip against the real `MasterDeck.pptx`, not just the structural name-check already on record.

### What was built
- `renderer.py` — the actual seed-slide-duplication + named-shape-fill engine (D3/D13/5.8), not a mock.
- `s1_template_fill.py` — duplicates all 8 manifest layouts' seed slides, fills every slot (text/bullets/image) with realistic sample content, saves `output/S1_filled_deck.pptx`.
- `s1_validate_pptx.py` — reopens the saved file with python-pptx and checks every slot landed on the right shape, with the right content, inheriting font/size/color.
- `s1_validate_powerpoint.ps1` — opens the saved file in **real PowerPoint via COM automation** and exports every slide to PNG. This is the strongest available proxy for "opens with no repair prompt."
- `s1_negative_control.py` + `s1_validate_negative_control.ps1` — deliberately reproduces the exact failure D13 warns against, to confirm the warning is real and the mitigation works.

### Evidence

**Layer 1 — python-pptx reopen (`s1_validate_pptx.py`):**
- 22/22 manifest slots across all 8 layouts landed on the correctly-named shape, correct type (text/bullets/picture).
- Every text run inherited the seed slide's font/size/color with no manual restyle (e.g. `IM_TITLE_DECKTITLE` → Cambria 30pt; `IM_TITLE_SUBTITLE` → Calibri 16pt — exactly the seed's values).
- Brand-chrome shapes not in the manifest (`IM_BRAND_MONOGRAM`, `IM_FOOTER_WORDMARK`, `IM_FOOTER_PAGENO`, `IM_BRAND_WORDMARK*`, `IM_CLOSING_WORDMARK`, `IM_CLOSING_CLAIM`) were left untouched on every slide — confirms the renderer never touches unnamed/non-manifest shapes.

**Layer 2 — real PowerPoint via COM (`s1_validate_powerpoint.ps1`):**
- `Presentations.Open()` on the output file threw **no exception** — no repair prompt.
- `Slides.Count == 8` as expected.
- All 8 slides exported to PNG (`output/screenshots/slide1.png`…`slide8.png`) and visually inspected: brand colors (blue #0957D1, red #FF3531), Cambria headline / Calibri body fonts, wordmark, monogram, footer, and page numbers are all present and correct on every layout, including inserted images sized/cropped correctly at their manifest `geometry_in`.

**Negative control — the D13 bug, reproduced and fixed (`s1_negative_control.py`):**
| Slide | r:embed id referenced | Resolves after save+reopen? |
|---|---|---|
| already_filled (normal render) | rId2 | ✅ resolves, 30,143 bytes |
| naive_reclone (**no** relationship remap — the exact thing D13 forbids) | rId2 (same id, copied verbatim) | ❌ `KeyError: no relationship with key 'rId2'` |
| fixed_reclone (relationship remap — what `renderer.py` actually does) | rId2 → remapped to a new id in its own `.rels` | ✅ resolves, 30,143 bytes |

Opening the naive-reclone file in real PowerPoint shows **no repair prompt**, but the picture renders as a literal broken-image icon ("Nous ne pouvons pas afficher l'image") — i.e. exactly the silent, un-flagged broken slide the SDD's own D9 principle says must never ship. This empirically confirms D13's rationale is a real correctness invariant, not theoretical caution, and confirms the mitigation (duplicate-before-fill, never re-clone an already-filled slide, remap `r:embed` on every clone) actually closes it.

### Bugs found and fixed along the way (worth carrying into the real implementation)
1. **`Slide.shapes` is a `@lazyproperty`.** `Presentation.slides.add_slide()` calls `slide.shapes.clone_layout_placeholders(...)` internally — which means `.shapes` gets cached (bound to the pre-fill, empty `spTree`) *before* your own code ever runs. Any later duplicate-and-swap-content approach that goes through `prs.slides.add_slide()` and then touches `.shapes` will silently find nothing, forever, on that slide object. Fix: build the destination slide via the lower-level `prs.part.add_slide(layout)` + manual `sldIdLst` registration, never touching `.shapes` until *after* the seed slide's content has been swapped in. This is a nonobvious python-pptx gotcha the real renderer implementation must carry forward — flagging it explicitly for the SDD's Renderer component (5.8).
2. **Image slots lose manifest addressability after fill.** `add_picture()` names the new picture shape generically (`"Picture N"`), not the slot's `shape_name` — so after removing the grey placeholder rect, there is no longer any shape named e.g. `IM_TWOCOL_IMAGE` on the slide. That breaks D3's whole "stable name addressing" premise for image slots specifically: a future field-edit regeneration of just that image would have no way to find it again. Fix applied: rename the inserted picture to the slot's `shape_name` immediately after insertion. **Recommend this be written into the manifest's `fill_protocol` explicitly** (it currently says "insert a picture... remove the grey placeholder rect" but doesn't say to preserve the name).

### Verdict against spike pass criteria
| Criterion | Result |
|---|---|
| 100% of placeholders addressable by a stable identifier | ✅ 22/22, by exact `shape_name` |
| Output opens with no repair prompt, theme intact | ✅ confirmed via real PowerPoint COM, 8/8 slides visually correct |
| Text inherits font/size/color, no manual restyle | ✅ confirmed on every text/bullets slot |

**S1: PASS.** No architectural change forced. Two implementation-level fixes identified and applied; both are worth carrying into the production Renderer (5.8) and its test suite.

---

## S2 — Structured slide-spec reliability → **PASS**

**Question:** using structured outputs, how reliably does Claude produce a schema-valid, *usable* slide-spec from a varied brief?

### What was built
- `slide_spec_schema.py` — the minimal SlideSpec envelope (SDD §6.2): `layout_id` (enum'd to the real 8 manifest layouts), `content_kind`, `slots`. Split into two validation layers per D6: (1) the coarse JSON schema forced on the model at the API level, and (2) `validate_against_manifest()`, a Pydantic-equivalent business-rule pass that — once `layout_id` is known — checks that exact layout's required slots, `max_chars`, `max_items`, `max_chars_per_item`, and `allowed_values` (e.g. `IM_TITLE_CLAIM`'s two permitted strings). This two-layer split is necessary because a single fixed tool schema can't express "slot X is required only when layout_id=title" — that's exactly the job D6 assigns to the Pydantic layer.
- `briefs.py` — 20 varied briefs spanning short / long / messy(typos, run-on) / multi-point / comparison / data-bearing, per the spike method, deliberately not naming a `layout_id` (layout choice is part of what's tested).
- `llm_providers.py` — small multi-provider client (Anthropic tool-use forcing, OpenAI `json_schema` strict mode, Gemini `responseSchema`, local Ollama `format=<schema>`), so the harness runs against whatever credential is available rather than hard-failing. Anthropic is the reference/primary path — it's the SDD's actual tech-stack choice (D6/§8).
- `s2_spec_reliability.py` — runs all 20 briefs through the chosen provider, validates both layers, and applies the D6 bounded one-shot retry (feeds validation errors back) on any failure.

### Methodology note on credentials
No `ANTHROPIC_API_KEY` was configured in the working environment at the start of this run. Rather than block, the harness was made provider-pluggable and first smoke-tested end-to-end against a local Ollama model (`gemma4:latest`, 8B) with zero cloud credentials — this validated the harness logic itself (schema plumbing, manifest validation, retry path) but is **not** evidence about bet A2, which the spike doc and SDD define specifically in terms of Claude models. An Anthropic API key was then provided mid-session; the run below is against **`claude-sonnet-5`**, the SDD's actual reference model for this component (5.5 Content Generator), and is the authoritative S2 result. (The key was scoped to a single command invocation via an inline session env var, never written to disk or logged, and unset immediately after use.)

### Evidence — claude-sonnet-5, N=20 (`output/s2_results_anthropic.json`)

| Metric | Result |
|---|---|
| Attempt-1 schema-valid (gross JSON shape) | **20/20 (100%)** |
| Attempt-1 business-valid (manifest rules: required slots, max_chars, max_items, max_chars_per_item, allowed_values) | **20/20 (100%)** |
| Retries needed | **0/20** |
| Final valid after bounded retry | **20/20 (100%)** |

Zero retries were exercised against Claude — every attempt was correct the first time, so the retry *path* remains logically implemented but empirically untested against a real Claude failure (it was however exercised and confirmed working end-to-end during the Ollama smoke test, where the smaller local model did produce some invalid first attempts).

### Usability (human-rated, all 20 — not just a 20-sample, since N=20 total)
Read every one of the 20 generated specs end to end. All 20 are "usable first draft" (no rewrite needed, at most light copyediting) — **100%**, well above the ≥90% bar. Highlights of what's actually being tested here, not just schema shape:
- **Layout selection was correct in all 20 cases**, including non-obvious ones: a raw stat ("NPS is 62... make it a big callout, quote-style") correctly went to `quote_callout` rather than `exhibit_data`; "just the org chart photo... with a caption" correctly went to `image_only` rather than `two_column`.
- **Numeric/logical fidelity held**: "40 hires vs. a target of 25" was correctly characterized as "outpaced plan by 60%" (40/25 = 1.6) without being asked to compute it.
- **Constraint-following under messy input held**: the exec-summary brief said "keep it to like 5-6 bullets max, nothing too dramatic on the legal bit" — the model returned exactly 5 bullets and phrased the legal item as "Minor pending legal matter being addressed routinely," honoring both the count cap and the tone instruction.
- **Graceful degradation on a maximally vague brief** ("something about our competitive advantage") — produced a reasonable generic draft rather than refusing, per the system prompt's instruction; this is the intended behavior per the design (SDD 5.5's failure mode is `needs_review` only on validation failure, not on vagueness).

### Two qualitative findings (neither blocks the pass — both worth carrying forward)
1. **Fabrication risk on underspecified briefs.** On a few briefs (the pricing-comparison slide, and especially the maximally vague "competitive advantage" one), the model filled gaps with plausible-sounding but unsourced specifics (e.g. "Fortune 500 engagements," competitors having "5+ confusing options" — nothing in the brief said either). Schema/business validation can't catch this class of error since the *shape* is perfectly valid. For a client-facing consulting deck, invented specifics are a real risk. **Recommendation:** add an explicit instruction to the real Content Generator's system prompt — "never invent specific facts, numbers, or named entities not present in the brief or supplied data; prefer qualitative language when data is missing" — and treat this as a QA dimension (5.9) worth a lightweight LLM-judge check in addition to the vision/overflow checks already scoped there.
2. **`content_kind` enum has no value for "plain user-supplied image."** SDD §6.1/6.2 defines `content_kind` as `text | chart | diagram`, but the `image_only` layout's real use case (§6.3: "a single dominant visual — exhibit photo, diagram, screenshot") includes plain photos that are neither a chart nor a generated diagram. The model picked `"diagram"` for a literal whiteboard photo, which is a defensible-but-slightly-off classification given the enum it was constrained to. **Recommendation:** either add a fourth `content_kind` value (e.g. `"image"`) or clarify in the manifest/prompt that `image_only` + a user-supplied photo should be classified as `"text"` (no generated exhibit) — a one-line spec clarification, not an architectural issue.

### Verdict against spike pass criteria
| Criterion | Target | Result |
|---|---|---|
| Schema-valid on first attempt | ≥99% | **100%** (20/20) |
| Usable first draft | ≥90% (of a 20-sample) | **100%** (20/20, full population rated) |
| Bounded retry closes remaining gap | — | N/A this run (0 failures to close); confirmed functionally correct via the Ollama smoke test |

**S2: PASS**, decisively — against the actual model (`claude-sonnet-5`) the architecture is designed around, not a proxy. No mandatory validate→retry-as-hard-requirement change, no schema flattening, no human-in-loop-at-generation-time forced by this result. The two qualitative findings above are prompt/spec refinements, not architectural changes.

**Caveat on sample size:** N=20 here (per this task's instruction) vs. the spike doc's own N=100 for a fully powered estimate of the "≥99% schema-valid" claim — 20/20 is consistent with ≥99% but doesn't statistically distinguish 100% from, say, 95%, the way N=100 would. Recommend running the full N=100 sweep (trivial: `python s2_spec_reliability.py anthropic claude-sonnet-5` against a larger `BRIEFS` list) before treating this as the final go/no-go artifact for the findings memo Doc 1 §6 asks for.

### Appendix — supplementary comparison: local Ollama model (`output/s2_results_ollama.json`)

Before an Anthropic key was available, the same 20 briefs were run through the harness against a local Ollama model (`gemma4:latest`, 8B, Q4_K_M) with zero cloud credentials, to validate the harness plumbing itself. This is **not** evidence toward bet A2 (that's specifically about Claude — see the methodology note above), but the contrast is informative about what the D6 bounded-retry mechanic actually buys, and where a much smaller local model's failure modes concentrate.

| Metric | claude-sonnet-5 | gemma4:latest (8B, local) |
|---|---|---|
| Attempt-1 schema-valid (gross shape) | 100% (20/20) | 100% (20/20) |
| Attempt-1 business-valid (manifest rules) | 100% (20/20) | **75%** (15/20) |
| Retried | 0/20 | 5/20 |
| Final valid after bounded retry | 100% (20/20) | **90%** (18/20) |

Both models hit the coarse JSON-schema layer perfectly (both APIs' constrained-decoding/tool-forcing does its job). The gap opens entirely at the business-rule layer, and it's concentrated in one failure mode: **the smaller model doesn't reliably self-count characters against `max_chars`/`max_chars_per_item`.** 4 of the 5 initial failures were character-count overruns; the 5th was a cross-layout slot-name hallucination (it emitted `IM_TWOCOL_BODY` — a `two_column` slot — inside an `exhibit_data` slide).

The bounded retry (D6) recovered 3 of the 5 failures, but two are worth reading in full because they show the mechanic's actual limits, not just its average:
- **A genuinely unrecoverable quality miss:** brief #6 (pricing-tiers comparison) exceeded `max_chars_per_item=80` on 3 of 4 bullets on attempt 1 (e.g. 101 chars). Fed the exact violation back, the retry's bullets got *longer*, not shorter (up to 115 chars) — the model does not reliably know how to shorten its own prior output to fit a hard cap; it just rewrites and often re-overruns. Claude never produced a single character-count violation across all 20 briefs, so this comparison never came up for the authoritative run.
- **An infra failure masquerading as a model failure:** brief #3's retry hit a `ReadTimeout` against local Ollama (>120s) — a reminder that "final_valid: false" in these logs conflates two different causes (model can't fix it vs. the call itself failed), and a real implementation's `needs_review` routing (D9) should distinguish "retry errored" from "retry produced another invalid result" since they call for different operator responses.

Qualitatively, gemma4's *content* (topic selection, tone) was reasonable on nearly every brief — this isn't a "small models produce garbage" finding, it's specifically a constraint-adherence gap on hard numeric limits, exactly the kind of thing D6's validate→retry loop exists to catch before it ships. It did its job: 75%→90%, not 75%→shipped.

---

## Overall recommendation

**Both S1 and S2 pass.** Per Doc 1 §7's green-light rule ("S1 and S2 must pass to proceed with this architecture"), the architecture as designed in the SDD is confirmed — no fallback path from either gate is triggered. Proceed to M0/M1 per SDD §9. Carry forward four concrete, small action items surfaced by this spike:
1. Renderer (5.8): avoid `Slide.shapes` before content-swap on a duplicated slide (lazyproperty gotcha).
2. Renderer (5.8) / manifest `fill_protocol`: rename inserted pictures to their slot's `shape_name` to preserve addressability.
3. Content Generator (5.5) prompt: explicit anti-fabrication instruction for underspecified briefs.
4. SDD §6.1/6.2: resolve the `content_kind` enum gap for plain user-supplied images on `image_only`.
