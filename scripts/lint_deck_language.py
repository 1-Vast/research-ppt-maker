#!/usr/bin/env python3
"""Heuristic copy-edit review for research PPT JSON specs.

Warnings are prompts for human review, not proof of bad Chinese or bad science.
Usage: python scripts/lint_deck_language.py path/to/deck.json [--strict]
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
}
GENERIC_TITLES = {"研究背景", "方法介绍", "方法创新", "实验结果", "核心贡献", "模型框架"}
FIELDS = {
    "title", "idea", "lede", "text", "note", "name", "did", "limit", "approach",
    "question", "before", "after", "why", "summary", "points", "boundary", "caption",
}
META_FIELDS = {"source", "cite", "original_title", "authors", "affiliation", "first_page",
               "first_page_caption", "eyebrow", "path", "output", "label"}


def fields(value, path=""):
    if isinstance(value, dict):
        for key, child in value.items():
            if key in META_FIELDS or key in {"table", "chart", "figure"}:
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


def inspect(spec):
    results = []
    for i, slide in enumerate(spec.get("slides", []), 1):
        for path, value in fields(slide):
            s = re.sub(r"\*\*|~\{[^}]+\}", "", value)
            for term, hint in REVIEW_TERMS.items():
                if term in s:
                    results.append((i, path, f"审稿/抽象词“{term}”：{hint}"))
            if path == "title" and s.strip() in GENERIC_TITLES:
                results.append((i, path, "标题缺少本页的具体科研判断"))
            if path in {"idea", "lede"} and len(s) > 135:
                results.append((i, path, "核心讲解句较长：考虑拆成对象、操作、作用两句"))
            if re.search(r"(?:通过.{0,30}){2,}从而", s):
                results.append((i, path, "因果链过长：拆句并写出中间操作和输出"))
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("spec", type=Path)
    parser.add_argument("--strict", action="store_true", help="exit 1 if there are warnings")
    args = parser.parse_args()
    spec = json.loads(args.spec.read_text(encoding="utf-8"))
    issues = inspect(spec)
    for slide, path, why in issues:
        print(f"slide {slide} {path}: {why}")
    print(f"Language review: {len(issues)} potential issue(s); human review is mandatory.")
    return 1 if args.strict and issues else 0


if __name__ == "__main__":
    raise SystemExit(main())
