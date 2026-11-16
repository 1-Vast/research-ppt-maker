"""Regression tests: slide-copy QA and revision-scope candidates.

Run: python -m unittest discover -s tests -v
Requires the normal python-pptx / Pillow environment.
"""
from __future__ import annotations

import io
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from review_deck import compare_decks, check_slides


def make_deck(path, duplicate=False, labelled=False, text="sample"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    buf = io.BytesIO()
    Image.new("RGB", (80, 60), "white").save(buf, format="PNG")
    buf.seek(0)
    slide.shapes.add_picture(buf, Inches(1), Inches(1), width=Inches(2))
    sentence = "减少全连接代价，同时为跨蛋白重建与多尺度聚合提供接口。"
    for phrase in ([sentence, sentence] if duplicate else [text]):
        box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(5), Inches(0.6))
        box.text = phrase
    if labelled:
        box = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(5), Inches(0.6))
        box.text = "A/B/C 是本次讲解标注"
    prs.save(path)


class RevisionReviewTests(unittest.TestCase):
    def test_duplicate_explanation(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "duplicate.pptx"
            make_deck(path, duplicate=True)
            self.assertIn("duplicate_explanation", {x["code"] for x in check_slides(Presentation(path))})

    def test_clean_explanation(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "clean.pptx"
            make_deck(path)
            self.assertFalse(check_slides(Presentation(path)))

    def test_missing_abc_annotation_candidate(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "old-caption.pptx"
            make_deck(path, labelled=True)
            self.assertIn("orphaned_annotation_reference",
                          {x["code"] for x in check_slides(Presentation(path))})

    def test_same_images_with_rewritten_text_flag_revision(self):
        with tempfile.TemporaryDirectory() as tmp:
            old, new = Path(tmp) / "old.pptx", Path(tmp) / "new.pptx"
            make_deck(old, text="old")
            make_deck(new, text="new")
            report = compare_decks(old, new)
            self.assertEqual(report["slides_with_text_changes"], 1)
            self.assertTrue(report["all_picture_content_and_positions_unchanged"])
            self.assertTrue(report["structure_revision_review_needed"])

    def test_moved_image_not_misclassified_as_unchanged(self):
        with tempfile.TemporaryDirectory() as tmp:
            old, new = Path(tmp) / "old.pptx", Path(tmp) / "new.pptx"
            make_deck(old)
            make_deck(new)
            pptx = Presentation(new)
            pptx.slides[0].shapes[0].left += Inches(0.2)
            pptx.save(new)
            self.assertFalse(compare_decks(old, new)["all_picture_content_and_positions_unchanged"])


if __name__ == "__main__":
    unittest.main()
