#!/usr/bin/env python3
"""Check PPTX structure and flag geometry/font clues; this is not visual QA."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import zipfile

from pptx import Presentation


def walk(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == 6:
            yield from walk(shape.shapes)


def inspect(deck: Path) -> dict:
    with zipfile.ZipFile(deck) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f'Corrupt ZIP member: {bad}')
    presentation = Presentation(deck)
    slides = []
    for number, slide in enumerate(presentation.slides, 1):
        editable = dict(text=0, tables=0, charts=0)
        small, outside = [], []
        # Group children have local coordinates: only test top-level geometry.
        for shape in slide.shapes:
            if (shape.left < 0 or shape.top < 0 or
                    shape.left + shape.width > presentation.slide_width + 9144 or
                    shape.top + shape.height > presentation.slide_height + 9144):
                outside.append(dict(shape_id=shape.shape_id, name=shape.name))
        for shape in walk(slide.shapes):
            editable['tables'] += int(shape.has_table)
            editable['charts'] += int(shape.has_chart)
            frames = []
            if shape.has_text_frame:
                editable['text'] += int(bool(shape.text.strip()))
                frames.append(shape.text_frame)
            if shape.has_table:
                frames.extend(cell.text_frame for row in shape.table.rows for cell in row.cells)
            for frame in frames:
                for paragraph in frame.paragraphs:
                    for run in paragraph.runs:
                        size = run.font.size or paragraph.font.size
                        if size and size.pt < 16 and run.text.strip():
                            small.append(dict(shape_id=shape.shape_id, points=round(size.pt, 1), text=run.text[:90]))
        slides.append(dict(slide=number, editable_objects=editable,
                           out_of_bounds_candidates=outside, small_font_candidates=small))
    return dict(deck=str(deck), slide_count=len(slides), zip_integrity='ok',
                visual_review_required=True,
                limitations='Flags are review clues, not design failures. Inherited fonts, image text, overflow, overlap and scientific correctness require visual inspection.',
                slides=slides)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('deck', type=Path)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = inspect(args.deck.resolve())
        if not result['slide_count']:
            raise ValueError('No slides in presentation')
    except Exception as exc:
        sys.stderr.write(f'Cannot inspect deck: {exc}\n')
        return 1
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Opened {result["slide_count"]} slides; visual review still required: {args.output}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
