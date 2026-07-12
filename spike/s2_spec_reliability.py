"""S2 -- structured-output reliability test, per spike doc method:
  1. minimal SlideSpec schema (slide_spec_schema.py) -- done.
  2. call the model with structured outputs forced (tool-use / json-schema mode).
  3. run across ~20 varied briefs (briefs.py).
  4. validate every output with Pydantic-equivalent business rules; log schema-valid
     rate; bounded (one-shot) retry feeding the validation error back.

CAVEAT logged loudly, not buried: the spike doc's pass criteria (S2, section 7's
green-light rule) and the A2 bet they gate are specifically about *Claude*
models (claude-sonnet-5 / claude-opus-4-8 / claude-haiku-4-5 -- see spike doc
S2 + SDD S8). No ANTHROPIC_API_KEY (or any other cloud key) was available in
this environment; the only usable provider was a local Ollama model
(gemma4:latest, 8B, Q4_K_M quantized). This run is evidence about the harness
and about small local models, but it does NOT close the A2 bet -- that requires
re-running this exact script against claude-sonnet-5 or claude-haiku-4-5 once a
key is available (just: $env:ANTHROPIC_API_KEY = "..." then rerun).
"""
import json
import sys
from pathlib import Path

from briefs import BRIEFS
from llm_providers import call_structured, pick_default_provider, available_providers
from slide_spec_schema import ENVELOPE_SCHEMA, LAYOUTS_BY_ID, validate_against_manifest

OUT_DIR = Path(__file__).parent / "output"

MODEL_BY_PROVIDER = {
    "anthropic": "claude-sonnet-5",
    "openai": "gpt-4o-mini",
    "gemini": "gemini-2.0-flash",
    "ollama": "gemma4:latest",
}


def build_system_prompt():
    lines = [
        "You are the Content Generator for a presentation deck engine (SDD component 5.5).",
        "Given a plain-language brief, choose the single best-fitting layout_id from the "
        "enum and a content_kind, then fill the 'slots' object using ONLY the shape_name "
        "keys valid for that layout_id -- never invent a key, never include an image-typed "
        "slot. Text slots are one string; bullets slots are an array of strings. Respect "
        "every max_chars / max_items / max_chars_per_item / allowed_values constraint below "
        "-- these are hard limits, not suggestions. If the brief is vague, still make a "
        "reasonable editorial choice rather than refusing.",
        "",
        "Layouts and their fillable (non-image) slots:",
    ]
    for layout_id, layout in LAYOUTS_BY_ID.items():
        slot_descs = []
        for s in layout["slots"]:
            if s["type"] == "image":
                continue
            bits = [s["type"]]
            if s.get("required"):
                bits.append("required")
            if s.get("max_chars"):
                bits.append(f"max_chars={s['max_chars']}")
            if s.get("max_items"):
                bits.append(f"max_items={s['max_items']}")
            if s.get("max_chars_per_item"):
                bits.append(f"max_chars_per_item={s['max_chars_per_item']}")
            if s.get("allowed_values"):
                bits.append(f"allowed_values={s['allowed_values']}")
            slot_descs.append(f"{s['shape_name']} ({', '.join(bits)})")
        lines.append(f"- {layout_id} [{layout['use_when']}]: " + "; ".join(slot_descs))
    return "\n".join(lines)


SYSTEM_PROMPT = build_system_prompt()


def run_one(provider, model, brief):
    record = {"brief": brief, "provider": provider, "model": model}

    try:
        parsed, raw = call_structured(provider, model, SYSTEM_PROMPT, brief, ENVELOPE_SCHEMA)
        record["attempt1_raw"] = raw
    except Exception as e:
        record["attempt1_error"] = f"{type(e).__name__}: {e}"
        record["schema_valid"] = False
        record["business_valid"] = False
        record["retried"] = False
        record["final_valid"] = False
        return record

    envelope_errors = []
    layout_id = parsed.get("layout_id")
    if layout_id not in LAYOUTS_BY_ID:
        envelope_errors.append(f"layout_id {layout_id!r} not a valid enum value")
    if parsed.get("content_kind") not in ("text", "chart", "diagram"):
        envelope_errors.append(f"content_kind {parsed.get('content_kind')!r} not valid")
    if "slots" not in parsed or not isinstance(parsed.get("slots"), dict):
        envelope_errors.append("missing/invalid 'slots' object")

    record["attempt1_parsed"] = parsed
    record["schema_valid"] = len(envelope_errors) == 0
    business_errors = envelope_errors if envelope_errors else validate_against_manifest(layout_id, parsed.get("slots", {}))
    record["attempt1_errors"] = business_errors
    record["business_valid"] = len(business_errors) == 0

    if not business_errors:
        record["retried"] = False
        record["final_valid"] = True
        return record

    # bounded retry: feed the validation error back once (D6)
    retry_prompt = (
        f"{brief}\n\n"
        f"Your previous attempt was invalid for these reasons:\n"
        + "\n".join(f"- {e}" for e in business_errors)
        + "\nFix it and return a corrected, fully valid slide spec."
    )
    try:
        parsed2, raw2 = call_structured(provider, model, SYSTEM_PROMPT, retry_prompt, ENVELOPE_SCHEMA)
        record["attempt2_raw"] = raw2
        record["attempt2_parsed"] = parsed2
        layout_id2 = parsed2.get("layout_id")
        errors2 = validate_against_manifest(layout_id2, parsed2.get("slots", {})) if layout_id2 in LAYOUTS_BY_ID else ["invalid layout_id on retry"]
        record["attempt2_errors"] = errors2
        record["retried"] = True
        record["final_valid"] = len(errors2) == 0
    except Exception as e:
        record["attempt2_error"] = f"{type(e).__name__}: {e}"
        record["retried"] = True
        record["final_valid"] = False

    return record


def main():
    avail = available_providers()
    provider = sys.argv[1] if len(sys.argv) > 1 else pick_default_provider()
    model = sys.argv[2] if len(sys.argv) > 2 else MODEL_BY_PROVIDER[provider]
    out_path = OUT_DIR / f"s2_results_{provider}.json"
    print(f"providers available: {avail}")
    print(f"using provider={provider} model={model}")
    if provider != "anthropic":
        print(
            "*** CAVEAT: spike doc S2's pass criteria are specifically about Claude models. "
            "This run does not close bet A2 -- see FINDINGS.md. ***"
        )

    results = []
    for i, brief in enumerate(BRIEFS, 1):
        print(f"[{i}/{len(BRIEFS)}] {brief[:60]}...")
        results.append(run_one(provider, model, brief))

    out_path.write_text(json.dumps(results, indent=2), encoding="utf-8")

    n = len(results)
    n_schema_valid_1 = sum(r.get("schema_valid") for r in results)
    n_business_valid_1 = sum(r.get("business_valid") for r in results)
    n_final_valid = sum(r.get("final_valid") for r in results)
    n_retried = sum(r.get("retried") for r in results)

    print(f"\n=== S2 summary ({provider}/{model}, N={n}) ===")
    print(f"attempt-1 schema-valid (gross shape):   {n_schema_valid_1}/{n} ({100*n_schema_valid_1/n:.0f}%)")
    print(f"attempt-1 business-valid (manifest rules): {n_business_valid_1}/{n} ({100*n_business_valid_1/n:.0f}%)")
    print(f"retried:                                 {n_retried}/{n}")
    print(f"final valid after bounded retry:         {n_final_valid}/{n} ({100*n_final_valid/n:.0f}%)")
    print(f"\nfull results written to {out_path}")


if __name__ == "__main__":
    main()
