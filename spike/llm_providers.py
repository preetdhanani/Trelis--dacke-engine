"""S2 -- minimal multi-provider structured-output client (throwaway spike code).

The SDD's actual tech-stack choice (D6/S8) is the `anthropic` SDK with tool-use
forcing / JSON-schema mode -- that's the primary, reference path here. OpenAI,
Gemini, and a local Ollama model are wired in as swappable alternatives so this
harness runs against whatever credential/runtime is actually available in a
given environment, rather than hard-failing with no key configured.

IMPORTANT CAVEAT (see FINDINGS.md): S2's pass criteria and the underlying
architectural bet (spike doc A2) are specifically about *Claude* models. A run
against a different provider is informative about the harness, not a verdict
on A2, unless the provider used *is* Claude.
"""
import json
import os

import requests

OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")


def available_providers():
    avail = {}
    avail["anthropic"] = bool(os.environ.get("ANTHROPIC_API_KEY"))
    avail["openai"] = bool(os.environ.get("OPENAI_API_KEY"))
    avail["gemini"] = bool(os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY"))
    try:
        requests.get(f"{OLLAMA_HOST}/api/tags", timeout=2)
        avail["ollama"] = True
    except Exception:
        avail["ollama"] = False
    return avail


def pick_default_provider():
    """Preference order matches the SDD's actual model choice (Claude first);
    falls back to whatever is actually usable in this environment."""
    avail = available_providers()
    for name in ("anthropic", "openai", "gemini", "ollama"):
        if avail.get(name):
            return name
    raise RuntimeError(f"no LLM provider available: {avail}")


def call_structured(provider, model, system_prompt, user_prompt, json_schema, schema_name="slide_spec"):
    """Returns (parsed_dict, raw_text). Raises on transport/API error."""
    if provider == "anthropic":
        return _call_anthropic(model, system_prompt, user_prompt, json_schema, schema_name)
    if provider == "openai":
        return _call_openai(model, system_prompt, user_prompt, json_schema, schema_name)
    if provider == "gemini":
        return _call_gemini(model, system_prompt, user_prompt, json_schema)
    if provider == "ollama":
        return _call_ollama(model, system_prompt, user_prompt, json_schema)
    raise ValueError(f"unknown provider '{provider}'")


def _call_anthropic(model, system_prompt, user_prompt, json_schema, schema_name):
    import anthropic

    client = anthropic.Anthropic()
    tool = {"name": schema_name, "description": "Emit the slide spec.", "input_schema": json_schema}
    resp = client.messages.create(
        model=model,
        max_tokens=1024,
        system=system_prompt,
        tools=[tool],
        tool_choice={"type": "tool", "name": schema_name},
        messages=[{"role": "user", "content": user_prompt}],
    )
    for block in resp.content:
        if block.type == "tool_use":
            return block.input, json.dumps(block.input)
    raise RuntimeError("no tool_use block in Anthropic response")


def _call_openai(model, system_prompt, user_prompt, json_schema, schema_name):
    api_key = os.environ["OPENAI_API_KEY"]
    strict_schema = dict(json_schema)
    strict_schema["additionalProperties"] = False
    resp = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}"},
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "response_format": {
                "type": "json_schema",
                "json_schema": {"name": schema_name, "schema": strict_schema, "strict": True},
            },
        },
        timeout=60,
    )
    resp.raise_for_status()
    content = resp.json()["choices"][0]["message"]["content"]
    return json.loads(content), content


def _call_gemini(model, system_prompt, user_prompt, json_schema):
    api_key = os.environ.get("GOOGLE_API_KEY") or os.environ["GEMINI_API_KEY"]
    resp = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}",
        json={
            "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
            "systemInstruction": {"parts": [{"text": system_prompt}]},
            "generationConfig": {
                "responseMimeType": "application/json",
                "responseSchema": json_schema,
            },
        },
        timeout=60,
    )
    resp.raise_for_status()
    content = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
    return json.loads(content), content


def _call_ollama(model, system_prompt, user_prompt, json_schema):
    resp = requests.post(
        f"{OLLAMA_HOST}/api/chat",
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "format": json_schema,
            "stream": False,
            "options": {"temperature": 0.4},
        },
        timeout=120,
    )
    resp.raise_for_status()
    content = resp.json()["message"]["content"]
    return json.loads(content), content
