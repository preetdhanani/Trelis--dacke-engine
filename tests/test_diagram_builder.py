import unittest

from deck_engine.diagram_builder import build_diagram_data
from deck_engine.models.diagram import MAX_CHARS_PER_STEP, MAX_STEPS, MIN_STEPS, DiagramData


def steps(n, label="Step"):
    """Terse builder for n distinct, in-bounds step labels."""
    return [f"{label} {i + 1}" for i in range(n)]


class TestDiagramDataValidation(unittest.TestCase):
    def test_rejects_too_few_steps(self):
        for n in range(MIN_STEPS):
            with self.assertRaises(Exception):
                DiagramData(steps=steps(n))

    def test_rejects_too_many_steps(self):
        with self.assertRaises(Exception):
            DiagramData(steps=steps(MAX_STEPS + 1))

    def test_accepts_boundary_step_counts(self):
        self.assertEqual(len(DiagramData(steps=steps(MIN_STEPS)).steps), MIN_STEPS)
        self.assertEqual(len(DiagramData(steps=steps(MAX_STEPS)).steps), MAX_STEPS)

    def test_rejects_blank_or_whitespace_step_label(self):
        with self.assertRaises(Exception):
            DiagramData(steps=["Intake", ""])
        with self.assertRaises(Exception):
            DiagramData(steps=["Intake", "   "])

    def test_rejects_step_label_over_max_chars(self):
        with self.assertRaises(Exception):
            DiagramData(steps=["Intake", "x" * (MAX_CHARS_PER_STEP + 1)])

    def test_accepts_step_label_at_max_chars_boundary(self):
        data = DiagramData(steps=["Intake", "x" * MAX_CHARS_PER_STEP])
        self.assertEqual(len(data.steps[1]), MAX_CHARS_PER_STEP)


class TestBuildDiagramData(unittest.TestCase):
    def test_build_diagram_data_preserves_step_order(self):
        labels = ["Intake", "Review", "Approve", "Deliver"]
        self.assertEqual(build_diagram_data(labels).steps, labels)

    def test_build_diagram_data_raises_on_invalid_input(self):
        with self.assertRaises(Exception):
            build_diagram_data(["only one step"])


if __name__ == "__main__":
    unittest.main()
