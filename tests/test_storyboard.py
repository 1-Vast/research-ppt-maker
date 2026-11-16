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
    def test_default_paper_only_rejects_project_slides(self):
        data = {'papers': [{'name': 'Paper', 'paper_type': 'review', 'role': 'supporting', 'budget': 1,
                            'slides': [{'purpose': 'context'}]}],
                'project_slides': [{'purpose': 'progress', 'source': 'user experiment'}]}
        self.assertIn('unrequested_project_slides', {w['code'] for w in audit(data)})

    def test_authorized_paper_plus_project(self):
        data = {'scope': 'paper-plus-project', 'project_authorized': True,
                'papers': [{'name': 'Paper', 'paper_type': 'review', 'role': 'supporting',
                            'budget': 1, 'slides': [{'purpose': 'context'}]}],
                'project_slides': [{'purpose': 'progress', 'source': 'user/experiments/run1.csv'}]}
        self.assertEqual(audit(data), [])

    def test_project_mode_without_explicit_authorization_or_source_warns(self):
        data = {'scope': 'paper-plus-project',
                'papers': [{'name': 'Paper', 'paper_type': 'review', 'role': 'supporting',
                            'budget': 1, 'slides': [{'purpose': 'context'}]}],
                'project_slides': [{'purpose': 'progress'}]}
        codes = {w['code'] for w in audit(data)}
        self.assertIn('missing_project_authorization', codes)
        self.assertIn('untraceable_project_evidence', codes)

    def test_project_only_mode(self):
        data = {'scope': 'project-only', 'project_authorized': True, 'papers': [],
                'project_slides': [{'source': 'user/experiment.json', 'purpose': 'progress'}]}
        self.assertEqual(audit(data), [])

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
             "available_original_overview": False, "role": "primary", "budget": 1,
             "slides": [{"purpose": "method-overview", "visual": "redraw", "source": "Methods"}]}]}
        self.assertFalse(audit(data))

    def test_a_long_deep_reading_is_not_invalid_by_itself(self):
        data = {"papers": [{"name": "Deep", "paper_type": "original-method",
           "available_original_overview": True, "budget": 20,
           "slides": [{"purpose": "cover"},
                      {"purpose": "method-overview", "visual": "original-figure", "source": "Fig. 1"},
                      {"purpose": "mechanism"}] * 4}]}
        self.assertFalse(any(w["code"] == "over_budget" for w in audit(data)))

    def test_explicit_deep_reading_keeps_original_figure_checks(self):
        paper = {"paper_type": "original-method", "role": "primary", "budget": 16,
                 "available_original_overview": True,
                 "slides": [{"purpose": "method-overview", "visual": "redraw", "source": "Fig. 1"}]
                           + [{"purpose": "evidence", "experiment_id": x} for x in ("A", "B", "C")]}
        codes = {w["code"] for w in audit({"mode": "deep-reading", "papers": [paper]})}
        self.assertNotIn("too_many_evidence_topics", codes)
        self.assertIn("original_figure_not_first", codes)
        paper["overview_reason"] = "Original labels unreadable even in the best available source"
        self.assertFalse(audit({"mode": "deep-reading", "papers": [paper]}))
        paper["slides"] = [{"purpose": "mechanism"}]
        self.assertIn("missing_method_overview", {w["code"] for w in audit({"mode": "deep-reading", "papers": [paper]})})

    def test_primary_review_and_equal_papers_do_not_inherit_pair_limit(self):
        paper = {"paper_type": "review", "role": "primary", "budget": 12,
                 "slides": [{"purpose": "evidence", "experiment_id": x} for x in ("A", "B", "C")]}
        self.assertFalse(audit({"papers": [paper]}))
        papers = [dict(paper, role="equal", slides=[{"purpose": "context"}] * 8) for _ in range(3)]
        self.assertFalse(audit({"papers": papers}))

    def test_density_exception_does_not_override_declared_budget(self):
        paper = {"paper_type": "original-method", "role": "primary", "budget": 2,
                 "available_original_overview": False, "density_reason": "User requested all three comparisons",
                 "slides": [{"purpose": "overview", "visual": "redraw"}]
                           + [{"purpose": "evidence", "experiment_id": x} for x in ("A", "B", "C")]}
        codes = {w["code"] for w in audit({"papers": [paper]})}
        self.assertNotIn("too_many_evidence_topics", codes)
        self.assertIn("over_budget", codes)

    def test_malformed_and_missing_plan_fields_do_not_pass_silently(self):
        for data in ([], {"papers": [None]}, {"papers": [{"slide_plan": []}]},
                     {"papers": [{"slides": [None, {}], "budget": True, "role": []}], "mode": []}):
            with self.subTest(data=data):
                self.assertTrue(audit(data))

    def test_original_overview_precedes_mechanism(self):
        paper = {"paper_type": "original-method", "role": "primary", "budget": 2,
                 "available_original_overview": True,
                 "slides": [{"purpose": "mechanism", "visual": "explainer"},
                            {"purpose": "overview", "visual": "original-figure", "source": "Fig. 1"}]}
        self.assertIn("mechanism_before_overview", {w["code"] for w in audit({"papers": [paper]})})

    def test_null_reason_does_not_hide_warnings_and_redraw_needs_source(self):
        paper = {"paper_type": "original-method", "role": "primary", "budget": 4,
                 "available_original_overview": True, "overview_reason": None,
                 "slides": [{"purpose": "overview", "visual": "redraw"}]
                           + [{"purpose": "evidence", "experiment_id": x} for x in ("A", "B", "C")]}
        codes = {w["code"] for w in audit({"density_reason": None, "papers": [paper]})}
        self.assertIn("original_figure_not_first", codes)
        self.assertIn("too_many_evidence_topics", codes)
        paper["available_original_overview"] = False
        self.assertIn("missing_overview_source", {w["code"] for w in audit({"papers": [paper]})})


if __name__ == "__main__":
    unittest.main()
