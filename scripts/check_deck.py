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


def inspect(deck: Path, min_pt: float = 12) -> dict:
    with zipfile.ZipFile(deck) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f'Corrupt ZIP member: {bad}')
    presentation = Presentation(deck)
    slides = []
    for number, slide in enumerate(presentation.slides, 1):
        editable = dict(text=0, tables=0, charts=0)
        small, outside, picture_text_overlaps = [], [], []
        # Candidate labels placed directly over source figures. This is a geometry
        # warning only; the renderer and reviewer must judge image content.
        pictures = [x for x in slide.shapes if x.shape_type == 13]
        for picture in pictures:
            for label in slide.shapes:
                if label.shape_id == picture.shape_id or not label.has_text_frame or not label.text.strip():
                    continue
                if label.width <= 0 or label.height <= 0:
                    continue
                ix = max(0, min(picture.left + picture.width, label.left + label.width)
                         - max(picture.left, label.left))
                iy = max(0, min(picture.top + picture.height, label.top + label.height)
                         - max(picture.top, label.top))
                if ix * iy / (label.width * label.height) >= 0.35:
                    picture_text_overlaps.append(dict(
                        picture_id=picture.shape_id, text_id=label.shape_id,
                        text=label.text.strip()[:60]))
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
                        if size and size.pt < min_pt and run.text.strip():
                            small.append(dict(shape_id=shape.shape_id, points=round(size.pt, 1), text=run.text[:90]))
        slides.append(dict(slide=number, editable_objects=editable,
                           overlap_groups_not_inspected=sum(x.shape_type == 6 for x in slide.shapes),
                           out_of_bounds_candidates=outside, small_font_candidates=small,
                           picture_text_overlap_candidates=picture_text_overlaps))
    return dict(deck=str(deck), slide_count=len(slides), zip_integrity='ok',
                visual_review_required=True,
                limitations='Flags are review clues, not design failures. Overlap checks use top-level, axis-aligned bounding boxes and flag intersections covering at least 35% of a text box. They do not inspect image content, rotation, transparency or stacking order; decorative images can cause false positives, and small intersections or grouped objects can escape detection. Inherited fonts, image text, overflow and scientific correctness require visual inspection.',
                slides=slides)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('deck', type=Path)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--min-pt', type=float, default=12, help='flag explicit font sizes below this')
    args = parser.parse_args(argv)
    try:
        result = inspect(args.deck.resolve(), args.min_pt)
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
