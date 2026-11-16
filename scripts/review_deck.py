#!/usr/bin/env python3
"""Inspect a PPTX for repeated explanations, stale labels and revision scope.

Only emits review candidates. It cannot judge scientific truth or inspect image
content. Run on final PPTX after rendering and human scientific review.

  python scripts/review_deck.py new.pptx --output report.json
  python scripts/review_deck.py new.pptx --baseline old.pptx --output report.json
"""
from __future__ import annotations

import argparse
from collections import defaultdict
from difflib import SequenceMatcher
import hashlib
import json
from pathlib import Path
import re

from pptx import Presentation


PAPER_FIGURE_REF = re.compile(r"(?:Fig(?:ure)?\.?\s*)(\d+[a-z]?)", re.I)
EVIDENCE_WORDS = re.compile(r"实验|证据|预测|泛化|验证|结果|性能|测试")
METRIC_WORDS = re.compile(r"PCC|Pearson|AUROC|AUPRC|C-index|准确率|相关系数|\bFID\b|\br\s*[=><]", re.I)
ARCH_WORDS = re.compile(r"架构图|架构。|模型架构|原图模型|模型结构图|architecture", re.I)


def slide_texts(slide):
    def texts(shapes):
        for shape in shapes:
            if shape.shape_type == 6:
                yield from texts(shape.shapes)
            elif shape.has_text_frame and shape.text.strip():
                yield shape.text.strip()
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if not cell.is_spanned and cell.text.strip():
                            yield cell.text.strip()
    return list(texts(slide.shapes))


def normalized_segment(text):
    return re.sub(r"[\s，。；：:！!？?、,.\-—“”\"'（）()\[\]]", "", text).lower()


def check_slides(prs):
    issues = []
    for i, slide in enumerate(prs.slides, 1):
        texts = slide_texts(slide)
        seen = defaultdict(set)
        for shape_id, text in enumerate(texts):
            for segment in re.split(r"(?<=[。；！？!?])|\n", text):
                key = normalized_segment(segment.strip())
                if len(key) >= 20:
                    seen[key].add(shape_id)
        for key, shape_ids in seen.items():
            if len(shape_ids) > 1:
                issues.append({"slide": i, "code": "duplicate_explanation",
                               "severity": "review", "text": key[:90],
                               "detail": f"Long sentence repeated in {len(shape_ids)} text shapes"})
        flat = "\n".join(texts)
        standalone = {t.strip().upper() for t in texts if len(t.strip()) <= 2}
        if "A/B/C" in flat and "标注" in flat and not {"A", "B", "C"}.issubset(standalone):
            issues.append({"slide": i, "code": "orphaned_annotation_reference",
                           "severity": "review",
                           "detail": "Text cites added A/B/C labels but these text labels are absent as separate shapes; inspect image and rendering."})
        if EVIDENCE_WORDS.search(flat) and METRIC_WORDS.search(flat) and ARCH_WORDS.search(flat):
            issues.append({"slide": i, "code": "mixed_architecture_and_metric_evidence",
                           "severity": "review",
                           "figure_refs": sorted(set(PAPER_FIGURE_REF.findall(flat))),
                           "detail": "Metric claims share a slide with architecture-figure wording; verify a matching results panel/table is displayed."})
    return issues



def check_project_scope(prs, scope="paper-only"):
    """Candidate only: detect project-progress headers in paper-only decks."""
    if scope != "paper-only":
        return []
    issues = []
    for i, slide in enumerate(prs.slides, 1):
        title = slide.shapes.title
        titles = [title.text.strip()] if title is not None else []
        titles.extend(shape.text.strip() for shape in slide.shapes
                      if shape.has_text_frame and shape.top < prs.slide_height / 4)
        if any(len(t) <= 32 and "\n" not in t and t.startswith(("项目进展", "项目实验", "项目研究进展"))
               for t in titles):
            issues.append({
                "slide": i,
                "code": "possible_unrequested_project_progress",
                "severity": "review",
                "detail": "Wording may refer to the user's project progress. Check whether it describes cited work or actual project content, and verify the user's request if it is project content."
            })
    return issues


def picture_signature(slide):
    return [(hashlib.sha256(shape.image.blob).hexdigest(),
             int(shape.left), int(shape.top), int(shape.width), int(shape.height),
             shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom,
             shape.rotation)
            for shape in slide.shapes if shape.shape_type == 13]


def compare_decks(old, new):
    old_prs, new_prs = Presentation(old), Presentation(new)
    matched = min(len(old_prs.slides), len(new_prs.slides))
    text_changed = 0
    picture_same = 0
    picture_total = 0
    identical_layout = 0
    page_details = []
    for j in range(matched):
        before, after = old_prs.slides[j], new_prs.slides[j]
        a, b = slide_texts(before), slide_texts(after)
        if a != b:
            text_changed += 1
        pi, pn = picture_signature(before), picture_signature(after)
        total = max(len(pi), len(pn))
        pic_match = sum(x == y for x, y in zip(pi, pn))
        picture_total += total
        picture_same += pic_match
        geometry = lambda slide: [(int(s.shape_type), int(s.left), int(s.top), int(s.width), int(s.height)) for s in slide.shapes]
        same_layout = geometry(before) == geometry(after)
        if same_layout:
            identical_layout += 1
        page_details.append({"slide": j + 1, "text_changed": a != b,
                             "text_similarity": round(SequenceMatcher(None, "\n".join(a), "\n".join(b)).ratio(), 3),
                             "pictures_unchanged": pic_match, "picture_slots": total,
                             "same_shape_geometry": same_layout})
    same_count = len(old_prs.slides) == len(new_prs.slides)
    unpaired_pictures = sum(len(picture_signature(prs.slides[j]))
                            for prs in (old_prs, new_prs)
                            for j in range(matched, len(prs.slides)))
    picture_total += unpaired_pictures
    group_count = sum(s.shape_type == 6 for prs in (old_prs, new_prs)
                      for slide in prs.slides for s in slide.shapes)
    same_canvas = (old_prs.slide_width == new_prs.slide_width and
                   old_prs.slide_height == new_prs.slide_height)
    pictures_untouched = (same_count and same_canvas and not group_count and
                          picture_total > 0 and picture_total == picture_same)
    return {"baseline_slide_count": len(old_prs.slides), "new_slide_count": len(new_prs.slides),
            "slides_with_text_changes": text_changed, "shared_slide_pairs": matched,
            "identical_shape_geometry_slides": identical_layout,
            "unchanged_pictures": picture_same, "picture_slots": picture_total,
            "unpaired_slide_pictures": unpaired_pictures,
            "all_picture_content_and_positions_unchanged": pictures_untouched,
            "structure_revision_review_needed": bool(same_count and pictures_untouched),
            "comparison_mode": "same_slide_index_and_picture_order",
            "same_slide_dimensions": same_canvas,
            "groups_not_compared": group_count,
            "caution": "Unchanged pictures are only a review clue for structural rewrite requests, never proof of failure. Comparison pairs slides by index and pictures by order; inserted/reordered slides are not semantically matched. Grouped picture geometry, charts, SmartArt, image content, theme changes and scientific correctness require manual review. An unchanged result is withheld when groups or slide dimensions differ.",
            "slides": page_details}


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("deck", type=Path)
    p.add_argument("--baseline", type=Path)
    p.add_argument("--scope", choices=("paper-only", "paper-plus-project", "project-only"), default="paper-only")
    p.add_argument("--output", type=Path)
    args = p.parse_args(argv)
    new = Presentation(args.deck)
    report = {"deck": str(args.deck), "slide_count": len(new.slides), "scope": args.scope,
              "limitations": "Candidates use editable text, including group text and table cells; image text, charts and SmartArt are not read. Missing editable labels may still exist inside images. No candidate is an automatic failure or proof of scientific accuracy.",
              "content_review_candidates": check_slides(new) + check_project_scope(new, args.scope)}
    if args.baseline:
        report["revision_comparison"] = compare_decks(args.baseline, args.deck)
    output = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output, encoding="utf-8")
        print(f"Review candidates: {len(report['content_review_candidates'])}; report: {args.output}")
    else:
        print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
