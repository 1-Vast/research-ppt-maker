"""Checks that compact journal-club storyboards prioritize original figures and decisive evidence.

Run: python -m unittest discover -s tests -v
"""
from __future__ import annotations
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from check_storyboard import audit


class StoryboardTests(unittest.TestCase):
    def test_original_overview_and_reasonable_density(self):
        data = {"papers": [{"name": "Paper A", "paper_type": "original-method",
                 "role": "primary", "available_original_overview": True,
                 "budget": 6, "slides": [
                    {"purpose": "cover"},
                    {"purpose": "context"},
                    {"purpose": "method-overview", "visual": "original-paper-figure", "source": "Fig. 1"},
                    {"purpose": "mechanism", "visual": "explainer"},
                    {"purpose": "mechanism", "visual": "explainer"},
                    {"purpose": "evidence", "experiment_id": "primary"}]}]}
        self.assertEqual(audit(data), [])

    def test_redraw_first_and_process_title_are_warned(self):
        data = {"papers": [{"name": "A", "paper_type": "original-method",
            "available_original_overview": True, "slides": [
            {"purpose": "cover", "title": "原始方法论文"},
            {"purpose": "method-overview", "visual": "redraw", "source": "Fig. 1"},
            {"purpose": "evidence", "experiment_id": "one"},
            {"purpose": "evidence", "experiment_id": "two"},
            {"purpose": "evidence", "experiment_id": "three"}]}]}
        codes = {x["code"] for x in audit(data)}
        self.assertIn("original_figure_not_first", codes)
        self.assertIn("production_meta_in_title", codes)
        self.assertIn("too_many_evidence_topics", codes)

    def test_review_two_cases_and_multi_paper_budget(self):
        data = {"papers": [
          {"name": "Main", "paper_type": "original-method", "role": "primary",
           "slides": [{"purpose": "other"}] * 9},
          {"name": "Review", "paper_type": "review", "role": "supporting",
           "slides": [{"purpose": "cover"}, {"purpose": "context"},
                      {"purpose": "evidence", "experiment_id": "A"},
                      {"purpose": "evidence", "experiment_id": "B"}]}
        ]}
        codes = {x["code"] for x in audit(data)}
        self.assertIn("too_many_review_cases", codes)
        self.assertIn("multi_paper_density", codes)

    def test_unavailable_original_figure_can_be_redrawn(self):
        data = {"papers": [{"name": "No original figure", "paper_type": "original-method",
             "available_original_overview": False,
             "slides": [{"purpose": "method-overview", "visual": "redraw"}]}]}
        self.assertFalse(audit(data))

    def test_a_long_deep_reading_is_not_invalid_by_itself(self):
        data = {"papers": [{"name": "Deep", "paper_type": "original-method",
           "available_original_overview": True, "budget": 20,
           "slides": [{"purpose": "cover"},
                      {"purpose": "method-overview", "visual": "original-figure", "source": "Fig. 1"},
                      {"purpose": "mechanism"}] * 4}]}
        self.assertFalse(any(w["code"] == "over_budget" for w in audit(data)))


if __name__ == "__main__":
    unittest.main()
