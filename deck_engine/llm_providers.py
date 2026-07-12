"""Multi-provider structured-output client for the Content Generator (5.5).

D6 ("structured outputs + Pydantic + bounded retry") is provider-agnostic --
only the wire mechanism for forcing/constraining JSON output differs between
vendors. `claude-sonnet-5` (Anthropic) remains the SDD's reference model
(D6/SS8); Ollama (local, no credential needed) and Gemini are supported as
swappable alternatives -- e.g. for offline/local dev, or when only a
non-Anthropic credential is available.

Uses stdlib `urllib` rather than `requests` so this stays within the
project's three declared dependencies (python-pptx, anthropic, pydantic); the
`anthropic` SDK import is local to its own call function so ollama/gemini-only
usage never needs it importable.
"""
import http.client
import json
import os
import time
import urllib.error
import urllib.request

DEFAULT_MODEL = {
    "anthropic": "claude-sonnet-5",
    # Whatever's actually pulled locally -- override with --model. gemma4:latest
    # is what was available/tested on the dev machine this was built against.
    "ollama": "gemma4:26b-a4b-it-q4_K_M",
    # A pinned version (e.g. "gemini-2.0-flash") goes stale and gets retired --
    # confirmed the hard way: it 404'd as "no longer available" well before
    # this default needed touching for any other reason. Use the rolling
    # "-latest" alias instead so this default doesn't silently break again;
    # override with --model for a specific pinned version if reproducibility
    # across runs matters more than always getting the current model.
    "gemini": "gemini-3.1-flash-lite",
}

PROVIDERS = tuple(DEFAULT_MODEL)

OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
# Larger local models (e.g. a 26B quantized model) on long/complex briefs can
# comfortably exceed 3 minutes on CPU -- override via env var rather than
# hardcoding a bigger constant every time a bigger model gets pulled.
OLLAMA_TIMEOUT = int(os.environ.get("OLLAMA_TIMEOUT", "240"))


class ProviderError(Exception):
    """A structured-output call to an LLM provider failed, was unreachable, or
    returned something that couldn't be parsed as the requested JSON object."""


def check_provider_ready(provider):
    """Cheap fail-fast credential/reachability check, meant to be called once
    before any real generation work -- mirrors template_registry's "fail fast
    at load, never mid-render" philosophy applied to the LLM side.
    """
    if provider == "anthropic":
        if not os.environ.get("ANTHROPIC_API_KEY"):
            raise ProviderError("ANTHROPIC_API_KEY is not set")
    elif provider == "gemini":
        if not (os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")):
            raise ProviderError("GOOGLE_API_KEY or GEMINI_API_KEY is not set")
    elif provider == "ollama":
        try:
            urllib.request.urlopen(f"{OLLAMA_HOST}/api/tags", timeout=3)
        except Exception as e:
            raise ProviderError(f"could not reach local Ollama at {OLLAMA_HOST}: {e}") from e
    else:
        raise ValueError(f"unknown provider '{provider}' (expected one of {PROVIDERS})")


def call_structured(provider, model, system_prompt, user_prompt, schema, tool_name="emit"):
    """Returns the parsed JSON object the model produced, matching `schema` as
    closely as that provider's structured-output mechanism guarantees (none
    of them guarantee it perfectly -- callers still validate, per D6)."""
    if provider == "anthropic":
        return _call_anthropic(model, system_prompt, user_prompt, schema, tool_name)
    if provider == "ollama":
        return _call_ollama(model, system_prompt, user_prompt, schema)
    if provider == "gemini":
        return _call_gemini(model, system_prompt, user_prompt, schema)
    raise ValueError(f"unknown provider '{provider}' (expected one of {PROVIDERS})")


def _call_anthropic(model, system_prompt, user_prompt, schema, tool_name):
    import anthropic  # local import: only needed when provider=anthropic

    client = anthropic.Anthropic()
    tool = {"name": tool_name, "description": "Emit the requested slide data.", "input_schema": schema}
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=2048,
            system=system_prompt,
            tools=[tool],
            tool_choice={"type": "tool", "name": tool_name},
            messages=[{"role": "user", "content": user_prompt}],
        )
    except anthropic.APIError as e:
        raise ProviderError(f"anthropic API error: {e}") from e
    for block in resp.content:
        if block.type == "tool_use":
            return block.input
    raise ProviderError(f"anthropic: no tool_use block in response for tool '{tool_name}'")


# HTTP status codes worth retrying: rate limiting + transient server-side
# failures. A 4xx that isn't 429 is our fault (bad request/auth) -- don't retry.
_RETRYABLE_STATUS = {429, 500, 502, 503, 504}


def _http_post_json(url, payload, headers=None, timeout=120, retries=2):
    """POST `payload` as JSON and parse the JSON response. Every failure --
    HTTP error, unreachable host, read timeout, or a raw socket-level error
    like a connection reset -- is wrapped in a clean, flaggable ProviderError
    rather than escaping as a traceback (D9). Transient failures (429/5xx and
    network-level errors) are retried with a short linear backoff, since a
    reset or rate-limit blip usually clears on the next attempt; a permanent
    error (e.g. HTTP 400/401/404) is raised immediately without wasting retries.
    """
    data = json.dumps(payload).encode("utf-8")
    last_err = None
    for attempt in range(retries + 1):
        req = urllib.request.Request(
            url, data=data, method="POST", headers={"Content-Type": "application/json", **(headers or {})}
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            last_err = ProviderError(f"HTTP {e.code} from {url}: {body}")
            if e.code in _RETRYABLE_STATUS and attempt < retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            raise last_err from e
        except (urllib.error.URLError, TimeoutError, OSError, http.client.HTTPException) as e:
            # All the non-HTTP-status failure modes, wrapped uniformly:
            #  - urllib.error.URLError: host unreachable / DNS / connection refused
            #  - TimeoutError: read timeout on an already-open connection (e.g.
            #    slow local Ollama inference) -- NOT a URLError subclass
            #  - OSError: raw socket errors like ConnectionResetError
            #    ("forcibly closed by the remote host", WinError 10054) and
            #    BrokenPipeError -- seen against Gemini mid-run
            #  - http.client.HTTPException: IncompleteRead / BadStatusLine
            # Any of these escaping raw would be an ugly traceback instead of a
            # clean ProviderError; all are plausibly transient, so retry.
            last_err = ProviderError(f"network error talking to {url}: {e.__class__.__name__}: {e}")
            if attempt < retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            raise last_err from e
    raise last_err  # defensive: loop always returns or raises above


def _call_ollama(model, system_prompt, user_prompt, schema):
    result = _http_post_json(
        f"{OLLAMA_HOST}/api/chat",
        {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "format": schema,  # Ollama's structured-output mode: constrain to this JSON schema
            "stream": False,
            # num_predict/num_ctx defaults are small enough that a multi-slide
            # M2 response (several slide_intents, each with a few string
            # fields) can truncate mid-JSON -- observed directly as an
            # "Unterminated string" json.JSONDecodeError. Set both generously.
            "options": {"temperature": 0.4, "num_predict": 4096, "num_ctx": 8192},
        },
        timeout=OLLAMA_TIMEOUT,
    )
    content = result.get("message", {}).get("content")
    if not content:
        raise ProviderError(f"ollama: no message content in response: {result}")
    try:
        return json.loads(content)
    except json.JSONDecodeError as e:
        raise ProviderError(f"ollama: response content was not valid JSON: {e}") from e


def _sanitize_schema_for_gemini(schema):
    """Gemini's responseSchema is an OpenAPI-3.0-style subset of JSON Schema --
    it rejects/ignores keywords like `additionalProperties`. Strip anything
    outside that subset recursively rather than risk the whole call failing.
    """
    if isinstance(schema, dict):
        return {k: _sanitize_schema_for_gemini(v) for k, v in schema.items() if k != "additionalProperties"}
    if isinstance(schema, list):
        return [_sanitize_schema_for_gemini(v) for v in schema]
    return schema


def _call_gemini(model, system_prompt, user_prompt, schema):
    api_key = os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ProviderError("GOOGLE_API_KEY or GEMINI_API_KEY must be set to use provider=gemini")

    result = _http_post_json(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}",
        {
            "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
            "systemInstruction": {"parts": [{"text": system_prompt}]},
            "generationConfig": {
                "responseMimeType": "application/json",
                "responseSchema": _sanitize_schema_for_gemini(schema),
            },
        },
        timeout=120,
    )
    try:
        content = result["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError) as e:
        raise ProviderError(f"gemini: unexpected response shape: {result}") from e
    try:
        return json.loads(content)
    except json.JSONDecodeError as e:
        raise ProviderError(f"gemini: response content was not valid JSON: {e}") from e
