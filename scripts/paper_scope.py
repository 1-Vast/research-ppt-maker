"""Shared, deliberately conservative cross-paper scope hints."""
from __future__ import annotations

import re

CROSS_PAPER_PURPOSES = {
    "cross-paper-comparison", "cross-paper-synthesis", "paper-comparison",
    "multi-paper-comparison", "multi-paper-synthesis",
}
CROSS_PAPER_LABEL = re.compile(
    r"^(?:(?:双|两|多)(?:篇)?论文|跨论文|跨文献)\s*(?:[/:：·—-]\s*)?"
    r"(?:对照|对比|比较|综合|总结)|"
    r"^(?:cross[- ]paper|multi[- ]paper)\s+(?:comparison|synthesis|summary)",
    re.I,
)


def is_cross_paper_label(value):
    """Flag explicit cross-paper headings, not ordinary within-paper baselines."""
    return isinstance(value, str) and bool(CROSS_PAPER_LABEL.search(value.strip()))


def is_cross_paper_slide(slide):
    purpose = str(slide.get("purpose", "")).strip().lower().replace("_", "-")
    return (purpose in CROSS_PAPER_PURPOSES or
            any(is_cross_paper_label(slide.get(field)) for field in ("title", "eyebrow")))
