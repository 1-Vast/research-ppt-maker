"""Regression checks for PPT language warnings and source-figure overlap detection.

Run: python -m unittest discover -s tests -v
Requires python-pptx and Pillow (the normal deck builder dependencies).
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
from lint_deck_language import inspect as lint_language
from check_deck import inspect as check_deck


class LanguageChecks(unittest.TestCase):
    def test_reviewer_jargon_is_flagged(self):
        spec = {"slides": [{"layout": "evidence", "title": "跨队列泛化",
                            "points": ["待检验主张：模型能否迁移？", "观察与结论：r=0.67"]}]}
        issues = lint_language(spec)
        assert any("待检验主张" in msg for _, _, msg in issues)
        assert any("观察与结论" in msg for _, _, msg in issues)

    def test_natural_explanation_does_not_trigger_jargon(self):
        spec = {"slides": [{"layout": "method",
                            "title": "同一空间位置的不同蛋白先交换信息",
                            "lede": "模型在每个位置比较不同蛋白，再沿图像位置交换特征。"}]}
        self.assertFalse(lint_language(spec))


class FigureOverlapChecks(unittest.TestCase):
    def _create_deck(self, filepath, overlap=True):
        deck = Presentation()
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        buffer = io.BytesIO()
        Image.new("RGB", (300, 200), "white").save(buffer, format="PNG")
        buffer.seek(0)
        slide.shapes.add_picture(buffer, Inches(1), Inches(1), width=Inches(4), height=Inches(3))
        top = 2 if overlap else 5
        text = slide.shapes.add_textbox(Inches(2), Inches(top), Inches(1.2), Inches(0.4))
        text.text = "A"
        deck.save(filepath)

    def test_overlapping_letter_detected(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "overlap.pptx"
            self._create_deck(path, overlap=True)
            report = check_deck(path)
            hits = report["slides"][0]["picture_text_overlap_candidates"]
            self.assertTrue(any(x["text"] == "A" for x in hits))

    def test_external_annotation_not_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "external.pptx"
            self._create_deck(path, overlap=False)
            report = check_deck(path)
            self.assertEqual(report["slides"][0]["picture_text_overlap_candidates"], [])


if __name__ == "__main__":
    unittest.main()
