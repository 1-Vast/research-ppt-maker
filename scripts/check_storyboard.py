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


def audit(data: dict) -> list[dict]:
    warnings: list[dict] = []
    papers = data.get("papers", [])
    if not isinstance(papers, list) or not papers:
        return [{"code": "missing_papers", "detail": "Expected a nonempty papers list."}]

    for idx, paper in enumerate(papers, 1):
        name = paper.get("name") or f"paper {idx}"
        slides = paper.get("slides", [])
        if not isinstance(slides, list) or not slides:
            warnings.append({"paper": name, "code": "missing_slides", "detail": "No slide plan."})
            continue
        budget = paper.get("budget")
        if isinstance(budget, int) and budget > 0 and len(slides) > budget:
            warnings.append({"paper": name, "code": "over_budget",
                             "detail": f"{len(slides)} slides exceed proposed budget {budget}; justify exceptions or remove low-value pages."})
        paper_type = str(paper.get("paper_type", "")).lower()
        if paper_type in METHOD_TYPES and paper.get("available_original_overview") is True:
            overviews = [(i, s) for i, s in enumerate(slides)
                         if s.get("purpose") in OVERVIEW_PURPOSE]
            if not overviews:
                warnings.append({"paper": name, "code": "missing_method_overview",
                                 "detail": "No overview in the plan despite an available original model figure."})
            else:
                first_index, first = overviews[0]
                if first.get("visual") not in OVERVIEW_VISUAL or not first.get("source"):
                    warnings.append({"paper": name, "code": "original_figure_not_first",
                                     "slide": first_index + 1,
                                     "detail": "First method overview should prioritize a sourced original architecture figure; put the redraw on a later mechanism slide."})
        experiments = []
        for i, slide in enumerate(slides, 1):
            if slide.get("purpose") in EVIDENCE_PURPOSE:
                experiments.append(slide.get("experiment_id") or f"slide-{i}")
            caption = (str(slide.get("title", "")) + " " + str(slide.get("eyebrow", ""))).lower()
            if any(x in caption for x in ("据原文重绘", "原始方法论文", "本次讲解标注", "完整模型 · 据原文")):
                warnings.append({"paper": name, "code": "production_meta_in_title",
                                 "slide": i, "detail": "Move production/source notes to compact references; keep a meaningful scientific title."})
        # Multiple pages covering the same experiment are legitimate if needed.
        distinct = len(set(experiments))
        if paper_type in METHOD_TYPES and distinct > 2:
            warnings.append({"paper": name, "code": "too_many_evidence_topics",
                             "detail": f"{distinct} experiment topics; journal clubs usually need 1–2 decisive studies, unless a larger scope is justified."})
        if paper_type in REVIEW_TYPES and distinct > 1:
            warnings.append({"paper": name, "code": "too_many_review_cases",
                             "detail": "A supporting review usually needs at most one representative cited study; do not imply the review performed it."})
    if len(papers) >= 2:
        total = sum(len(p.get("slides", [])) for p in papers)
        if total > 11 and any(p.get("role") == "supporting" for p in papers):
            warnings.append({"code": "multi_paper_density", "detail": f"{total} slides with a supporting paper; consider a 9–11 slide compact group meeting, but prioritize completeness and readability."})
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
