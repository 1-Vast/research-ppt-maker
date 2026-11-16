#!/usr/bin/env python3
"""Advisory quality checks for a multi-paper journal-club storyboard.

Input schema (a compact subset of the internal storyboard):
{"papers": [{"name": "...", "paper_type": "original-method",
 "role": "primary", "available_original_overview": true,
 "budget": 7, "slides": [
   {"purpose": "cover", "visual": "paper-first-page"},
   {"purpose": "context"},
   {"purpose": "method-overview", "visual": "original-paper-figure",
    "source": "Fig. 1d"},
   {"purpose": "mechanism", "visual": "explainer"},
   {"purpose": "evidence", "experiment_id": "masking"},
 ]}]}

Optional mode="deep-reading" disables default topic/count advice, not the declared
budget or original-overview checks. density_reason (top-level or per paper)
records an exception to default density advice; overview_reason explains a
necessary redraw-first fallback without removing overview/source checks.
Warnings are for human decisions, never mechanical failure or proof of quality.
Usage: python scripts/check_storyboard.py work/storyboard.json [--strict]
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

METHOD_TYPES = {"original-method", "method", "model"}
REVIEW_TYPES = {"review", "perspective", "conceptual"}
OVERVIEW_VISUAL = {"original-paper-figure", "original-figure", "paper-overview"}
EVIDENCE_PURPOSE = {"evidence", "experiment"}
OVERVIEW_PURPOSE = {"method-overview", "overview", "model-overview"}


def has_reason(value):
    return isinstance(value, str) and bool(value.strip())


def audit(data: dict) -> list[dict]:
    warnings: list[dict] = []
    if not isinstance(data, dict):
        return [{"code": "invalid_storyboard", "detail": "Expected a JSON object."}]
    papers = data.get("papers", [])
    if not isinstance(papers, list) or not papers:
        return [{"code": "missing_papers", "detail": "Expected a nonempty papers list."}]
    mode = data.get("mode", "journal-club")
    if mode not in ("journal-club", "deep-reading"):
        warnings.append({"code": "invalid_mode", "detail": "Use journal-club or deep-reading."})
    compact = mode == "journal-club" and not has_reason(data.get("density_reason"))
    valid_papers = []
    for idx, paper in enumerate(papers, 1):
        if not isinstance(paper, dict):
            warnings.append({"code": "invalid_paper", "detail": f"Paper {idx} must be an object."})
            continue
        name = paper.get("name") or f"paper {idx}"
        paper_type = str(paper.get("paper_type", "")).lower()
        if not paper_type:
            warnings.append({"paper": name, "code": "missing_paper_type", "detail": "Declare paper_type to select applicable checks."})
        if paper.get("role") not in ("primary", "supporting", "equal"):
            warnings.append({"paper": name, "code": "missing_or_invalid_role", "detail": "Declare role as primary, supporting or equal."})
        budget = paper.get("budget")
        if type(budget) is not int or budget <= 0:
            warnings.append({"paper": name, "code": "missing_or_invalid_budget", "detail": "Declare a positive integer budget, including covers."})
            budget = None
        slides = paper.get("slides", [])
        if not isinstance(slides, list) or not slides:
            warnings.append({"paper": name, "code": "missing_slides", "detail": "No slides list; convert slide_plan to the checker schema."})
            continue
        if budget and len(slides) > budget:
            warnings.append({"paper": name, "code": "over_budget",
                             "detail": f"{len(slides)} slides exceed proposed budget {budget}; revise the budget with reasons or remove low-value pages."})
        valid_slides = []
        for i, slide in enumerate(slides, 1):
            if not isinstance(slide, dict) or not isinstance(slide.get("purpose"), str) or not slide["purpose"].strip():
                warnings.append({"paper": name, "slide": i, "code": "invalid_slide", "detail": "Each slide needs an object with a purpose."})
                continue
            valid_slides.append((i, slide))
        valid_papers.append(paper)
        available = paper.get("available_original_overview")
        if paper_type in METHOD_TYPES:
            if not isinstance(available, bool):
                warnings.append({"paper": name, "code": "missing_overview_availability", "detail": "Declare available_original_overview as true or false after checking sources."})
            overviews = [(i, slide) for i, slide in valid_slides if slide["purpose"] in OVERVIEW_PURPOSE]
            if not overviews:
                warnings.append({"paper": name, "code": "missing_method_overview", "detail": "No method overview in the plan."})
            else:
                first_index, first = overviews[0]
                if any(i < first_index and slide["purpose"] == "mechanism" for i, slide in valid_slides):
                    warnings.append({"paper": name, "code": "mechanism_before_overview", "detail": "Introduce the overall model before detailed mechanism pages."})
                if available is True and first.get("visual") not in OVERVIEW_VISUAL and not has_reason(paper.get("overview_reason")):
                    warnings.append({"paper": name, "code": "original_figure_not_first", "slide": first_index,
                                     "detail": "Use the original architecture first, or record overview_reason for a necessary fallback."})
                if not first.get("source"):
                    warnings.append({"paper": name, "code": "missing_overview_source", "slide": first_index,
                                     "detail": "An overview, including a redraw, needs a traceable source."})
        experiments = []
        for i, slide in valid_slides:
            if slide["purpose"] in EVIDENCE_PURPOSE:
                experiments.append(str(slide.get("experiment_id") or f"slide-{i}"))
            caption = (str(slide.get("title", "")) + " " + str(slide.get("eyebrow", ""))).lower()
            if any(x in caption for x in ("据原文重绘", "原始方法论文", "本次讲解标注", "完整模型 · 据原文")):
                warnings.append({"paper": name, "code": "production_meta_in_title", "slide": i,
                                 "detail": "Move production/source notes to compact references; keep a meaningful scientific title."})
        distinct = len(set(experiments))
        default_density = compact and not has_reason(paper.get("density_reason"))
        if default_density and paper_type in METHOD_TYPES and distinct > 2:
            warnings.append({"paper": name, "code": "too_many_evidence_topics",
                             "detail": f"{distinct} experiment topics; consider 1–2 decisive studies or record density_reason for the larger scope."})
        if default_density and paper_type in REVIEW_TYPES and paper.get("role") == "supporting" and distinct > 1:
            warnings.append({"paper": name, "code": "too_many_review_cases",
                             "detail": "A supporting review usually needs one representative cited study; record density_reason if multiple cases are necessary."})
    # This default is specific to one primary method paper plus one supporting review.
    paired_default = (len(papers) == len(valid_papers) == 2 and
                      any(p.get("paper_type") in METHOD_TYPES and p.get("role") == "primary" for p in valid_papers) and
                      any(p.get("paper_type") in REVIEW_TYPES and p.get("role") == "supporting" for p in valid_papers))
    if compact and paired_default and not any(has_reason(p.get("density_reason")) for p in valid_papers):
        total = sum(len(p["slides"]) for p in valid_papers)
        if total > 11:
            warnings.append({"code": "multi_paper_density", "detail": f"{total} slides; consider the 9–11 slide starting budget or record density_reason for the larger scope."})
    return warnings


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("storyboard", type=Path)
    p.add_argument("--strict", action="store_true", help="exit 1 on warnings (optional)")
    args = p.parse_args()
    data = json.loads(args.storyboard.read_text(encoding="utf-8"))
    problems = audit(data)
    print(json.dumps({"warning_count": len(problems), "warnings": problems},
                     ensure_ascii=False, indent=2))
    return int(args.strict and bool(problems))


if __name__ == "__main__":
    raise SystemExit(main())
