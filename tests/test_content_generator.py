import unittest

from deck_engine.content_generator import _coerce_dict


class TestCoerceDict(unittest.TestCase):
    """Regression test for a real failure seen against claude-sonnet-5: a
    structured-output object field can come back double-encoded as a JSON
    string instead of the real object, even under tool-use forcing with an
    explicit {"type": "object"} schema. _coerce_dict recovers the real
    structure instead of letting it blow up as a Pydantic dict_type error.
    """

    def test_coerce_dict_passes_through_real_dict(self):
        self.assertEqual(_coerce_dict({"a": 1}), {"a": 1})

    def test_coerce_dict_recovers_double_encoded_json_string(self):
        raw = '{"IM_BULLETS_TITLE": "Title", "IM_BULLETS_BODY": ["a"]}'
        self.assertEqual(_coerce_dict(raw), {"IM_BULLETS_TITLE": "Title", "IM_BULLETS_BODY": ["a"]})

    def test_coerce_dict_returns_empty_on_garbage(self):
        self.assertEqual(_coerce_dict("nope"), {})
        self.assertEqual(_coerce_dict([1, 2, 3]), {})


if __name__ == "__main__":
    unittest.main()
