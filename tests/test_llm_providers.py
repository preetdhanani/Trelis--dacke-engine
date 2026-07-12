import unittest
from unittest import mock

from deck_engine.llm_providers import (
    DEFAULT_MODEL,
    PROVIDERS,
    ProviderError,
    _http_post_json,
    call_structured,
    check_provider_ready,
)


class TestProviderDispatch(unittest.TestCase):
    def test_all_providers_have_a_default_model(self):
        for provider in PROVIDERS:
            self.assertIn(provider, DEFAULT_MODEL)
            self.assertTrue(DEFAULT_MODEL[provider])

    def test_unknown_provider_raises_value_error(self):
        with self.assertRaises(ValueError):
            call_structured("not_a_real_provider", "some-model", "sys", "user", {})

    def test_check_provider_ready_rejects_unknown_provider(self):
        with self.assertRaises(ValueError):
            check_provider_ready("not_a_real_provider")


class TestProviderReadyChecks(unittest.TestCase):
    def test_anthropic_requires_api_key(self):
        with mock.patch.dict("os.environ", {}, clear=True):
            with self.assertRaises(ProviderError):
                check_provider_ready("anthropic")

    def test_gemini_requires_api_key(self):
        with mock.patch.dict("os.environ", {}, clear=True):
            with self.assertRaises(ProviderError):
                check_provider_ready("gemini")

    def test_gemini_accepts_either_key_env_var(self):
        with mock.patch.dict("os.environ", {"GEMINI_API_KEY": "fake"}, clear=True):
            check_provider_ready("gemini")  # must not raise
        with mock.patch.dict("os.environ", {"GOOGLE_API_KEY": "fake"}, clear=True):
            check_provider_ready("gemini")  # must not raise


class TestHttpPostJsonErrorHandling(unittest.TestCase):
    """Regression tests: raw exceptions out of urllib.request.urlopen must be
    wrapped as clean, flaggable ProviderErrors, never escape as tracebacks.
    Two distinct real-run failures drove these:
      - a TimeoutError (read timeout on an already-open connection, e.g. slow
        local Ollama inference) -- NOT a urllib.error.URLError subclass;
      - a ConnectionResetError ("forcibly closed by the remote host",
        WinError 10054) mid-run against Gemini -- an OSError, also not a
        URLError subclass.
    `retries=0` keeps these instant (no backoff sleeps).
    """

    def test_timeout_error_is_wrapped_as_provider_error(self):
        with mock.patch("deck_engine.llm_providers.urllib.request.urlopen", side_effect=TimeoutError("timed out")):
            with self.assertRaises(ProviderError):
                _http_post_json("http://localhost:11434/api/chat", {"model": "x"}, timeout=1, retries=0)

    def test_url_error_is_wrapped_as_provider_error(self):
        import urllib.error

        with mock.patch(
            "deck_engine.llm_providers.urllib.request.urlopen",
            side_effect=urllib.error.URLError("connection refused"),
        ):
            with self.assertRaises(ProviderError):
                _http_post_json("http://localhost:11434/api/chat", {"model": "x"}, timeout=1, retries=0)

    def test_connection_reset_is_wrapped_as_provider_error(self):
        with mock.patch(
            "deck_engine.llm_providers.urllib.request.urlopen",
            side_effect=ConnectionResetError(10054, "forcibly closed by the remote host"),
        ):
            with self.assertRaises(ProviderError):
                _http_post_json("https://example.com/x", {"model": "x"}, timeout=1, retries=0)

    def test_transient_network_error_is_retried_then_succeeds(self):
        class FakeResp:
            def __enter__(self):
                return self

            def __exit__(self, *a):
                return False

            def read(self):
                return b'{"ok": true}'

        # first attempt resets the connection, second attempt succeeds
        with mock.patch(
            "deck_engine.llm_providers.urllib.request.urlopen",
            side_effect=[ConnectionResetError("reset"), FakeResp()],
        ) as m, mock.patch("deck_engine.llm_providers.time.sleep") as sleep:
            result = _http_post_json("https://example.com/x", {"model": "x"}, retries=2)
        self.assertEqual(result, {"ok": True})
        self.assertEqual(m.call_count, 2)
        sleep.assert_called()  # backed off before the successful retry


if __name__ == "__main__":
    unittest.main()
