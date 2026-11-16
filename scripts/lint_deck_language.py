#!/usr/bin/env python3
"""Heuristic copy-edit review for research PPT JSON specs and PPTX files.

Warnings are prompts for human review, not proof of bad Chinese or bad science.
Usage: python scripts/lint_deck_language.py path/to/deck.json|deck.pptx [--strict]
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

REVIEW_TERMS = {
    "待检验主张": "换成听众能回答的具体问题，例如“隐藏哪个蛋白后，还能预测什么？”",
    "作用路径": "写清谁对什么做了哪一步运算，避免把内部审稿栏目投到屏幕上",
    "结论范围": "直接说明这项实验不能推出什么",
    "读出接口": "说明是怎样聚合向量、训练分类器或预测头",
    "赋予通道语义": "说明怎样加入蛋白身份以及为什么跨数据集可复用",
    "跨尺度赋能": "直接列出实际被聚合与预测的细胞、区域或组织",
    "知识写入表示": "指出具体训练目标如何约束哪些变量",
    "怎样检验": "改写为具体的对照、划分和指标，或删除重复栏目名",
    "观察与结论": "直接写观测数字及其可支持的解释",
    "测谁": "写明蛋白身份或测量对象；不要用口语提示代替专业名称",
    "另丢弃": "写明额外随机移除/遮挡哪些通道以及条件",
    "看这条路": "直接描述具体信息传递路径",
    "看“×”": "说明图例 × 对应的评估条件",
    "mask 查询": "明确目标蛋白的 Mask Token、身份条件与解码器输入；核对原文",
    "每位置 s(p)": "首次给出每个空间位置的汇总表示及 p 的含义",
}
GENERIC_TITLES = {"研究背景", "方法介绍", "方法创新", "实验结果", "核心贡献", "模型框架"}
TERM_VARIANTS = {
    "看“×”": re.compile(r"""看\s*["'“‘]?\s*×"""),
    "mask 查询": re.compile(r"\bmask\s*查询", re.IGNORECASE),
    "每位置 s(p)": re.compile(r"每位置\s*s\s*[（(]\s*p\s*[）)]", re.IGNORECASE),
}
FIELDS = {
    "title", "idea", "lede", "text", "note", "name", "did", "limit", "approach",
    "question", "before", "after", "why", "summary", "points", "boundary", "caption",
    "example", "example_label", "label", "prior_label", "columns", "rows", "categories",
}
META_FIELDS = {"source", "cite", "original_title", "authors", "affiliation", "first_page",
               "first_page_caption", "eyebrow", "path", "output"}
CONTAINERS = {"context", "prior", "gap", "lanes", "steps", "formula", "items",
              "figure", "table", "chart"}


def fields(value, path=""):
    if isinstance(value, dict):
        for key, child in value.items():
            if key in META_FIELDS:
                continue
            if key not in FIELDS and key not in CONTAINERS:
                continue
            child_path = f"{path}.{key}" if path else key
            if isinstance(child, str):
                if key in FIELDS:
                    yield child_path, child
            else:
                yield from fields(child, child_path)
    elif isinstance(value, list):
        for i, child in enumerate(value):
            yield from fields(child, f"{path}[{i}]")
    elif isinstance(value, str):
        yield path, value


def inspect(spec):
    results = []
    for i, slide in enumerate(spec.get("slides", []), 1):
        for path, value in fields(slide):
            # Small grey text and subscripts/superscripts remain visible in the
            # builder. Remove their markup, not the words inside it.
            s = re.sub(r"[~_^]\{([^}]+)\}", r"\1", value.replace("**", ""))
            for term, hint in REVIEW_TERMS.items():
                if matches_term(term, s):
                    results.append((i, path, f"审稿/抽象词“{term}”：{hint}"))
            if path == "title" and s.strip() in GENERIC_TITLES:
                results.append((i, path, "标题缺少本页的具体科研判断"))
            if path in {"idea", "lede"} and len(s) > 135:
                results.append((i, path, "核心讲解句较长：考虑拆成对象、操作、作用两句"))
            if re.search(r"(?:通过.{0,30}){2,}从而", s):
                results.append((i, path, "因果链过长：拆句并写出中间操作和输出"))
    return results


def matches_term(term, text):
    variant = TERM_VARIANTS.get(term)
    return bool(variant.search(text)) if variant else term in text



def pptx_fields(slide):
    """Read editable slide text, including groups and tables, without OCR."""
    def walk(shapes, prefix=""):
        for shape in shapes:
            here = f"{prefix}shape-{shape.shape_id}"
            if shape.shape_type == 6:
                yield from walk(shape.shapes, here + "/")
            elif shape.has_text_frame and shape.text.strip():
                yield here, shape.text.strip()
            if shape.has_table:
                for ri, row in enumerate(shape.table.rows):
                    for ci, cell in enumerate(row.cells):
                        if not cell.is_spanned and cell.text.strip():
                            yield f"{here}/table[{ri},{ci}]", cell.text.strip()
    yield from walk(slide.shapes)


def inspect_pptx(path):
    """Apply the same wording heuristics to the delivered presentation.

    Cannot inspect text baked into screenshots, SmartArt or chart labels.
    """
    from pptx import Presentation
    results = []
    for slide_no, slide in enumerate(Presentation(path).slides, 1):
        for field, value in pptx_fields(slide):
            cleaned = re.sub(r"[~_^]\{([^}]+)\}", r"\1", value.replace("**", ""))
            for term, hint in REVIEW_TERMS.items():
                if matches_term(term, cleaned):
                    results.append((slide_no, field, f"表达候选“{term}”：{hint}"))
            if len(cleaned) > 135 and ("通过" in cleaned and "从而" in cleaned):
                results.append((slide_no, field, "长句可能包含多重因果连接，需人工检查"))
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("spec", type=Path, help="JSON authoring spec or final PPTX")
    parser.add_argument("--strict", action="store_true", help="exit 1 if there are warnings")
    args = parser.parse_args()
    if args.spec.suffix.lower() == ".pptx":
        issues = inspect_pptx(args.spec)
    else:
        spec = json.loads(args.spec.read_text(encoding="utf-8"))
        issues = inspect(spec)
    for slide, path, why in issues:
        print(f"slide {slide} {path}: {why}")
    print(f"Language review: {len(issues)} potential issue(s); human review is mandatory.")
    return 1 if args.strict and issues else 0


if __name__ == "__main__":
    raise SystemExit(main())
