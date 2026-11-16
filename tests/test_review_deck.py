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
from review_deck import compare_decks, check_slides, check_project_scope


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
    def test_project_content_is_reviewed_only_in_paper_scope(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        label = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        label.text = "项目进展 / 实验结果"
        self.assertIn("possible_unrequested_project_progress",
                      {x["code"] for x in check_project_scope(prs)})
        self.assertEqual(check_project_scope(prs, "paper-plus-project"), [])
        self.assertEqual(check_project_scope(prs, "project-only"), [])

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

    def test_crop_or_rotation_is_a_picture_change(self):
        for attribute, value in (("crop_left", 0.1), ("rotation", 15)):
            with self.subTest(attribute=attribute), tempfile.TemporaryDirectory() as tmp:
                old, new = Path(tmp) / "old.pptx", Path(tmp) / "new.pptx"
                make_deck(old)
                make_deck(new)
                prs = Presentation(new)
                setattr(prs.slides[0].shapes[0], attribute, value)
                prs.save(new)
                self.assertFalse(compare_decks(old, new)["all_picture_content_and_positions_unchanged"])

    def test_added_slide_pictures_are_counted(self):
        with tempfile.TemporaryDirectory() as tmp:
            old, new = Path(tmp) / "old.pptx", Path(tmp) / "new.pptx"
            make_deck(old)
            make_deck(new)
            prs = Presentation(new)
            picture = prs.slides[0].shapes[0]
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(io.BytesIO(picture.image.blob), Inches(1), Inches(1))
            prs.save(new)
            report = compare_decks(old, new)
            self.assertEqual(report["picture_slots"], 2)
            self.assertEqual(report["unpaired_slide_pictures"], 1)
            self.assertFalse(report["all_picture_content_and_positions_unchanged"])

    def test_group_coverage_withholds_unchanged_claim(self):
        with tempfile.TemporaryDirectory() as tmp:
            old, new = Path(tmp) / "old.pptx", Path(tmp) / "new.pptx"
            for path in (old, new):
                make_deck(path)
                prs = Presentation(path)
                slide = prs.slides[0]
                slide.shapes.add_group_shape(list(slide.shapes))
                prs.save(path)
            report = compare_decks(old, new)
            self.assertEqual(report["groups_not_compared"], 2)
            self.assertFalse(report["all_picture_content_and_positions_unchanged"])

    def test_group_and_table_copy_are_checked(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        sentence = "该实验的对照条件必须明确，才能解释模型在跨队列条件下的性能变化。"
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        box.text = sentence
        slide.shapes.add_group_shape([box])
        table = slide.shapes.add_table(1, 1, Inches(1), Inches(3), Inches(4), Inches(1)).table
        table.cell(0, 0).text = sentence
        self.assertIn("duplicate_explanation", {x["code"] for x in check_slides(prs)})


if __name__ == "__main__":
    unittest.main()
