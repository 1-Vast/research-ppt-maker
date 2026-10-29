#!/usr/bin/env python3
"""Build a paper-explanation deck in the ivory-and-clay house style from a JSON spec.

Layouts: cover, background, method, innovation, evidence (see
references/deck-spec.md). Text is measured with the installed fonts so blocks
stack instead of overlapping; text that still does not fit is shrunk to a floor
size and reported. The deck still needs full rendering and visual review.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_TICK_MARK
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

W, H = 13.333, 7.5            # 16:9 canvas in inches
MX = 0.65                     # side margin
BOTTOM = 6.72                 # lowest edge of body content; the footer sits below
FLOOR = 13                    # smallest body size before a slide is reported as overfull

C = dict(canvas='FAF9F5', panel='F0EEE6', white='FFFFFF', hair='E0DDD3',
         ink='141413', ink2='3D3C38', muted='7A786F', faint='BDB9AE',
         clay='D97757', clay_deep='B5583A', clay_tint='F6E7DE',
         slate='5E7A94', olive='6B7F4E')

FONT_DIRS = [Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts',
             Path(os.environ.get('LOCALAPPDATA', '~')) / 'Microsoft/Windows/Fonts',
             Path('/Library/Fonts'), Path.home() / 'Library/Fonts', Path('/System/Library/Fonts')]
# role -> candidates (PowerPoint face, regular file, bold file, variable instance, collection index)
FONT_STACKS = {
    'serif': [('Noto Serif SC SemiBold', 'NotoSerifSC-VF.ttf', None, 'SemiBold', 0),
              ('STZhongsong', 'STZHONGS.TTF', None, None, 0), ('SimSun', 'simsun.ttc', None, None, 0)],
    'serif_text': [('Noto Serif SC', 'NotoSerifSC-VF.ttf', None, 'Regular', 0),
                   ('STSong', 'STSONG.TTF', None, None, 0), ('SimSun', 'simsun.ttc', None, None, 0)],
    'sans': [('Microsoft YaHei', 'msyh.ttc', 'msyhbd.ttc', None, 0), ('DengXian', 'Deng.ttf', 'Dengb.ttf', None, 0),
             ('PingFang SC', 'PingFang.ttc', 'PingFang.ttc', None, 0)],
    'math': [('Cambria Math', 'cambria.ttc', None, None, 1)],
}

LABELS = {
    'zh': dict(background='研究背景', method='方法框架', innovation='方法创新', evidence='实验结果',
               q_background='研究背景', q_method='方法框架',
               q_innovation='方法创新', q_evidence='实验结果',
               context='研究场景', example='例', prior='已有工作', did='进展', limit='不足',
               gap='研究需求', question='研究问题', approach='本文思路', idea='核心想法', first_page='原文首页',
               innov='创新点', before='已有方法', after='本文改进', why='主要作用', boundary='证据边界'),
    'en': dict(background='Context', method='Method', innovation='Novelty', evidence='Evidence',
               q_background='Why this work', q_method='How it works',
               q_innovation="What's new", q_evidence='What the evidence shows',
               context='Setting', example='Example', prior='Prior work', did='Achieved', limit='Limit',
               gap='Open problem', question='This paper asks', approach='This paper proposes', idea='Core idea', first_page='First page',
               innov='Change', before='Before', after='This paper', why='Why it helps', boundary='Limits'),
}

BOLD = re.compile(r'\*\*(.+?)\*\*')
INLINE = re.compile(r'_\{(.+?)\}|\^\{(.+?)\}|~\{(.+?)\}')
SCRIPTS = {1: 'sub', 2: 'sup', 3: 'small'}
SCALE = {'': 1.0, 'sub': 0.66, 'sup': 0.66, 'small': 0.75}
TOKEN = re.compile(r"[A-Za-z0-9\u00C0-\u024F\u0370-\u03FF.,:;!?%+\-_/'\"()\[\]=<>|~&*#@$]+ ?| +|.", re.S)
NO_LINE_START = set('，。、；：！？）》」』】,.;:!?)]}%’”')


def rgb(name: str) -> RGBColor:
    return RGBColor.from_string(C.get(name, name))


def runs_of(text: str) -> list[tuple[str, bool, str]]:
    """Split markup into (text, bold, script) runs: **bold**, _{sub}, ^{sup}, ~{small grey}."""
    out = []
    for i, part in enumerate(BOLD.split(text)):  # odd parts were inside ** **
        pos = 0
        for m in INLINE.finditer(part):
            out.append((part[pos:m.start()], i % 2 == 1, ''))
            out.append((m.group(m.lastindex), i % 2 == 1, SCRIPTS[m.lastindex]))
            pos = m.end()
        out.append((part[pos:], i % 2 == 1, ''))
    return [run for run in out if run[0]]


def paras(content) -> list[str]:
    items = content if isinstance(content, list) else [content]
    return [line for item in items if item is not None for line in str(item).split('\n')]


def find_font_file(name: str | None) -> Path | None:
    for folder in FONT_DIRS if name else []:
        if (folder / name).is_file():
            return folder / name
    return None


def strip_style(shape) -> None:
    """Drop the theme style so shapes get no inherited outline or shadow."""
    style = shape._element.find(qn('p:style'))
    if style is not None:
        shape._element.remove(style)


def first_page_png(pdf: Path, page: int = 0, zoom: float = 2.0) -> bytes:
    try:
        import pymupdf
    except ImportError:
        import fitz as pymupdf
    with pymupdf.open(pdf) as doc:
        return doc[page].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False).tobytes('png')


class Fonts:
    """Pick the first installed font per role and measure text with it."""

    def __init__(self):
        self.face, self.chosen, self.cache = {}, {}, {}
        for role, stack in FONT_STACKS.items():
            pick = next((c for c in stack if find_font_file(c[1])), None)
            self.face[role], self.chosen[role] = (pick or stack[0])[0], pick

    def width(self, text: str, role: str, bold: bool, size: float) -> float:
        font = self._load(role, bold)
        if font is None:  # no font file: CJK ~1 em, Latin ~0.56 em
            return size * sum(1.0 if ord(ch) >= 0x2E80 else 0.56 for ch in text)
        return font.getlength(text) * size / 100

    def _load(self, role: str, bold: bool):
        key = (role, bold)
        if key not in self.cache:
            self.cache[key] = None
            pick = self.chosen.get(role)
            if pick:
                _, regular, bold_file, instance, index = pick
                try:
                    from PIL import ImageFont
                    font = ImageFont.truetype(str(find_font_file(bold and bold_file or regular)), 100, index=index)
                    if instance:
                        font.set_variation_by_name(instance)
                    self.cache[key] = font
                except Exception:
                    pass
        return self.cache[key]


class Deck:
    def __init__(self, spec: dict, base: Path):
        self.spec, self.base = spec, base
        self.L = LABELS[spec.get('lang', 'zh')]
        self.lang_tag = 'en-US' if spec.get('lang') == 'en' else 'zh-CN'  # enables CJK line-break rules
        self.fonts = Fonts()
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.warnings: list[str] = []
        self.page, self.cite, self.group = 0, spec.get('cite', ''), []

    # ---------- primitives ----------
    def warn(self, message: str) -> None:
        self.warnings.append(f'slide {self.page}: {message}')

    def path(self, value: str) -> Path:
        p = Path(value)
        return p if p.is_absolute() else self.base / p

    def lines(self, para: str, role: str, size: float, width: float) -> int:
        limit, count, cur = width * 72 * 0.97, 1, 0.0
        for text, bold, script in runs_of(para):
            s = size * SCALE[script]
            for tok in TOKEN.findall(text):
                w = self.fonts.width(tok, role, bold, s)
                if w > limit and len(tok) > 1:  # PowerPoint breaks overlong tokens by character
                    for ch in tok:
                        cw = self.fonts.width(ch, role, bold, s)
                        count, cur = (count + 1, cw) if cur + cw > limit and cur else (count, cur + cw)
                elif cur + w <= limit or cur == 0 or tok[0] in NO_LINE_START:
                    cur += w
                else:
                    count, cur = count + 1, self.fonts.width(tok.lstrip(), role, bold, s)
        return count

    def height(self, content, size, width, role='sans', line=1.38, gap=0.45) -> float:
        ps = paras(content)
        n = sum(self.lines(p, role, size, width) for p in ps)
        return (n * size * line + max(0, len(ps) - 1) * gap * size) / 72

    def text(self, slide, x, y, w, content, size, role='sans', color='ink2', line=1.38, gap=0.45,
             align='l', emph=None, bold=False, spacing=None, h=None, anchor='t', draw=True) -> float:
        """Add a frameless text box; returns its measured height in inches."""
        height = self.height(content, size, w, role, line, gap)
        if not draw:
            return height
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h or height))
        tf = box.text_frame
        tf.word_wrap, tf.auto_size = True, MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = dict(t=MSO_ANCHOR.TOP, m=MSO_ANCHOR.MIDDLE, b=MSO_ANCHOR.BOTTOM)[anchor]
        for i, ptext in enumerate(paras(content)):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = dict(l=PP_ALIGN.LEFT, c=PP_ALIGN.CENTER, r=PP_ALIGN.RIGHT)[align]
            p.line_spacing = Pt(size * line)
            p._p.get_or_add_pPr().set('hangingPunct', '1')
            if i:
                p.space_before = Pt(size * gap)
            for run_text, run_bold, script in runs_of(ptext):
                run = p.add_run()
                run.text = run_text
                tone = 'muted' if script == 'small' else emph if run_bold and emph else color
                self.style_run(run, role, size * (SCALE[script] if script == 'small' else 1), tone,
                               bold or run_bold, script, spacing)
        return height

    def style_run(self, run, role, size, color, bold=False, script='', spacing=None) -> None:
        font = run.font
        font.size, font.bold = Pt(size), bold
        font.color.rgb = rgb(color)
        rpr = run._r.get_or_add_rPr()
        rpr.set('lang', self.lang_tag)
        if script in ('sub', 'sup'):
            rpr.set('baseline', '-25000' if script == 'sub' else '30000')
        if spacing:
            rpr.set('spc', str(int(spacing * 100)))
        face = self.fonts.face[role]
        for tag, value in (('a:latin', face), ('a:ea', self.fonts.face['sans'] if role == 'math' else face),
                           ('a:cs', face)):
            el = rpr.find(qn(tag))
            (el if el is not None else etree.SubElement(rpr, qn(tag))).set('typeface', value)

    def fit(self, content, w, max_h, size, floor, label, keep_lines=False, **kw) -> tuple[float, float]:
        """Largest size (0.5 pt steps, not below floor) whose text fits max_h.

        keep_lines: shrink until every hand-broken line stays on one line (titles).
        """
        role, manual = kw.get('role', 'sans'), keep_lines and len(paras(content)) > 1
        while True:
            h = self.height(content, size, w, **kw)
            unbroken = not manual or all(self.lines(p, role, size, w) == 1 for p in paras(content))
            if (h <= max_h and unbroken) or size <= floor:
                break
            size -= 0.5
        if h > max_h + 0.01:
            self.warn(f'{label} needs {h:.2f} in but has {max_h:.2f} in at {size} pt')
        return size, h

    def largest(self, measures, size, avail, label, floor=FLOOR) -> float:
        """One body size for several blocks: the largest at which every measure(size) fits avail."""
        measures = measures if isinstance(measures, list) else [measures]
        while any(m(size) > avail + 0.005 for m in measures) and size > floor:
            size -= 0.5
        over = max(m(size) for m in measures) - avail
        if over > 0.01:
            self.warn(f'{label} overflows by {over:.2f} in even at {size} pt; cut text')
        return size

    def rect(self, slide, x, y, w, h, fill=None, line=None, lw=0.75, radius=0.0):
        kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        strip_style(shape)
        if radius:
            shape.adjustments[0] = min(0.5, radius / min(w, h))
        if fill:
            shape.fill.solid()
            shape.fill.fore_color.rgb = rgb(fill)
        else:
            shape.fill.background()
        if line:
            shape.line.color.rgb, shape.line.width = rgb(line), Pt(lw)
        else:
            shape.line.fill.background()
        return shape

    def line(self, slide, x1, y1, x2, y2, color='hair', lw=0.75, head=False):
        conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        strip_style(conn)
        conn.line.color.rgb, conn.line.width = rgb(color), Pt(lw)
        if head:
            etree.SubElement(conn.line._get_or_add_ln(), qn('a:tailEnd'), type='triangle', w='med', len='med')
        return conn

    def badge(self, slide, x, y, d, label, fill='clay'):
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
        strip_style(dot)
        dot.fill.solid()
        dot.fill.fore_color.rgb = rgb(fill)
        dot.line.fill.background()
        tf = dot.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(label)
        self.style_run(run, 'serif', d * 72 * 0.46, 'white')

    def picture(self, slide, data: bytes, x, y, w, h, valign='m', border=None, shadow=False):
        from PIL import Image
        with Image.open(io.BytesIO(data)) as im:
            iw, ih = im.size
        scale = min(w / iw, h / ih)
        pw, ph = iw * scale, ih * scale
        px, py = x + (w - pw) / 2, y + ((h - ph) / 2 if valign == 'm' else 0)
        pic = slide.shapes.add_picture(io.BytesIO(data), Inches(px), Inches(py), Inches(pw), Inches(ph))
        if border:
            pic.line.color.rgb, pic.line.width = rgb(border), Pt(0.75)
        if shadow:
            effects = etree.SubElement(pic._element.spPr, qn('a:effectLst'))
            shade = etree.SubElement(effects, qn('a:outerShdw'), blurRad='165100', dist='50800',
                                     dir='5400000', algn='t', rotWithShape='0')
            colour = etree.SubElement(shade, qn('a:srgbClr'), val=C['ink'])
            etree.SubElement(colour, qn('a:alpha'), val='16000')
        return px, py, pw, ph

    def image_bytes(self, value) -> bytes:
        p = self.path(value)
        return first_page_png(p) if p.suffix.lower() == '.pdf' else p.read_bytes()

    # ---------- page furniture ----------
    def new_slide(self):
        self.page += 1
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb('canvas')
        return slide

    def chrome(self, slide, sd: dict, layout: str) -> float:
        """Draw tracker, eyebrow, title and footer; return the top of the body area."""
        L = self.L
        if len(self.group) > 1:
            names = [f'**{L[k]}**' if k == layout else L[k] for k in self.group]
            self.text(slide, W - MX - 5.2, 0.44, 5.2, '　·　'.join(names), 10.5, color='faint',
                      emph='clay_deep', align='r', line=1.2)
        self.text(slide, MX, 0.44, 7.0, sd.get('eyebrow', L['q_' + layout]), 12, color='clay_deep',
                  spacing=1.2, line=1.2)
        size, h = self.fit(sd['title'], W - 2 * MX, 0.95, 28, 23, 'title', keep_lines=True,
                           role='serif', line=1.18, gap=0)
        self.text(slide, MX, 0.78, W - 2 * MX, sd['title'], size, role='serif', color='ink', line=1.18, gap=0)
        self.line(slide, MX, 7.0, W - MX, 7.0)
        source = sd.get('source', self.cite)
        if source:
            self.text(slide, MX, 7.09, W - 2 * MX - 1.2, source, 10, color='muted', line=1.2)
        self.text(slide, W - MX - 1, 7.07, 1, f'{self.page:02d}', 11, role='serif', color='muted',
                  align='r', line=1.2)
        return 0.78 + h + 0.36

    def label(self, slide, x, y, w, text, color='clay_deep', draw=True) -> float:
        return self.text(slide, x, y, w, text, 12, color=color, spacing=1.2, line=1.2, draw=draw) + 0.1

    # ---------- layouts ----------
    def cover(self, sd: dict) -> None:
        slide, L = self.new_slide(), self.L
        self.cite = sd.get('cite', self.cite)
        px = 8.15
        self.rect(slide, px, 0, W - px, H, 'panel')
        if sd.get('first_page'):
            x, y, w, h = self.picture(slide, self.image_bytes(sd['first_page']), px + 0.6, 0.6,
                                      W - px - 1.2, 6.0, border='hair', shadow=True)
            self.text(slide, px + 0.3, y + h + 0.16, W - px - 0.6, sd.get('first_page_caption', L['first_page']),
                      10, color='muted', align='c', line=1.2)
        else:
            self.warn('cover has no first_page; add the paper first-page screenshot or PDF')
        x, w, y = 0.8, px - 1.45, 1.0
        if sd.get('eyebrow'):
            y += self.text(slide, x, y, w, sd['eyebrow'], 12, color='clay_deep', spacing=1.2, line=1.2) + 0.3
        size, h = self.fit(sd['title'], w, 2.45, 36, 28, 'cover title', keep_lines=True, role='serif',
                           line=1.16, gap=0)
        y += self.text(slide, x, y, w, sd['title'], size, role='serif', color='ink', line=1.16, gap=0) + 0.24
        if sd.get('original_title'):
            y += self.text(slide, x, y, w, sd['original_title'], 16, role='serif_text', line=1.3) + 0.18
        meta = [t for t in (sd.get('authors'), sd.get('affiliation')) if t]
        if meta:
            y += self.text(slide, x, y, w, meta, 12, color='muted', line=1.4, gap=0.2)
        if sd.get('idea'):
            size, h = self.fit(sd['idea'], w, 1.9, 18, 15, 'cover idea', line=1.5)
            top = 6.75 - h
            label_top = top - 0.12 - 12 * 1.2 / 72
            if label_top - 0.2 < y + 0.25:
                self.warn('cover idea collides with the title block; shorten title or idea')
            self.line(slide, x, label_top - 0.2, x + 0.5, label_top - 0.2, 'clay', 2)
            self.label(slide, x, label_top, w, L['idea'])
            self.text(slide, x, top, w, sd['idea'], size, color='ink', line=1.5, emph='clay_deep')

    def background(self, sd: dict) -> None:
        slide, L = self.new_slide(), self.L
        top = self.chrome(slide, sd, 'background')
        avail = BOTTOM - top
        ax, aw, bx, bw = MX, 3.45, 4.5, 4.5
        cx = bx + bw + 0.4
        cw, inner = W - MX - cx, W - MX - cx - 0.6
        ctx, prior, gap = sd.get('context', {}), sd.get('prior', []), sd.get('gap', {})
        example = [f"**{ctx.get('example_label', L['example'])}**", ctx['example']] if ctx.get('example') else None

        def context(z, draw=False):
            y = top + self.label(slide, ax, top, aw, ctx.get('label', L['context']), draw=draw)
            y += self.text(slide, ax, y, aw, ctx.get('text', ''), z, draw=draw)
            if example:
                eh = self.text(slide, 0, 0, aw - 0.36, example, z - 1.5, line=1.4, gap=0.25, draw=False) + 0.36
                y = max(y + 0.25, BOTTOM - eh) if draw else y + 0.25  # pin the example to the bottom edge
                if draw:
                    self.rect(slide, ax, y, aw, eh, 'panel', radius=0.06)
                    self.text(slide, ax + 0.18, y + 0.18, aw - 0.36, example, z - 1.5, line=1.4, gap=0.25,
                              emph='ink')
                y += eh
            return y - top

        def prior_col(z, draw=False):
            y = top + self.label(slide, bx, top, bw, sd.get('prior_label', L['prior']), color='slate', draw=draw)
            for i, item in enumerate(prior):
                if i:
                    if draw:
                        self.line(slide, bx, y + 0.12, bx + bw, y + 0.12)
                    y += 0.24
                start = y
                head = f"**{item['name']}**" + (f"　~{{{item['cite']}}}" if item.get('cite') else '')
                y += self.text(slide, bx + 0.2, y, bw - 0.2, head, z + 0.5, color='ink', line=1.3,
                               draw=draw) + 0.06
                body = [f"**{L['did']}**　{item['did']}" if item.get('did') else None,
                        f"**{L['limit']}**　{item['limit']}" if item.get('limit') else None]
                y += self.text(slide, bx + 0.2, y, bw - 0.2, [b for b in body if b], z - 1, line=1.38,
                               gap=0.3, emph='ink', draw=draw)
                if draw:
                    self.rect(slide, bx, start + 0.04, 0.04, y - start - 0.04, 'slate')
            return y - top

        def gap_col(z, draw=False, y0=0.0):
            y = y0 + self.label(slide, cx + 0.3, y0, inner, gap.get('label', L['gap']), draw=draw)
            y += self.text(slide, cx + 0.3, y, inner, gap.get('text', ''), z + 2.5, role='serif_text',
                           color='ink', line=1.45, emph='clay_deep', draw=draw)
            response = gap.get('approach') or gap.get('question')
            if response:
                y += 0.28
                y += self.label(slide, cx + 0.3, y, inner, L['approach'] if gap.get('approach') else L['question'], draw=draw)
                y += self.text(slide, cx + 0.3, y, inner, response, z - 0.5, color='ink', line=1.42,
                               draw=draw)
            return y - y0 + 0.6

        z = self.largest([context, prior_col, gap_col], 15.5, avail, 'background columns')
        self.rect(slide, cx, top, cw, avail, 'clay_tint', radius=0.05)  # gap panel spans the body height
        self.rect(slide, cx, top, 0.06, avail, 'clay')
        mid = top + min(avail, max(context(z), prior_col(z))) / 2  # arrows sit beside the content
        for gx in (ax + aw + (bx - ax - aw) / 2, bx + bw + 0.2):
            self.line(slide, gx - 0.13, mid, gx + 0.13, mid, 'clay', 1.5, head=True)
        context(z, True)
        prior_col(z, True)
        gap_col(z, True, top + 0.3 + max(0, (avail - gap_col(z)) / 2))

    def _steps_row(self, slide, steps, x, y, w, z, draw=False, card_h=None) -> float:
        n, gapw, pad, d = len(steps), 0.42, 0.16, 0.3
        cw = (w - gapw * (n - 1)) / n
        heights = []
        for i, step in enumerate(steps):
            cx = x + i * (cw + gapw)
            name_w = cw - 2 * pad - d - 0.12
            nh = max(d, self.height(f"**{step['name']}**", z + 0.5, name_w, line=1.25))
            dh = self.height(step.get('text', ''), z, cw - 2 * pad, line=1.36)
            heights.append(pad + nh + 0.08 + dh + pad)
            if draw:
                key = step.get('key')
                self.rect(slide, cx, y, cw, card_h, 'clay_tint' if key else 'white',
                          line='clay' if key else 'hair', lw=1.25 if key else 0.75, radius=0.07)
                self.badge(slide, cx + pad, y + pad, d, step.get('badge', i + 1))
                self.text(slide, cx + pad + d + 0.12, y + pad + max(0, (d - nh) / 2), name_w,
                          f"**{step['name']}**", z + 0.5, color='ink', line=1.25, emph='ink')
                self.text(slide, cx + pad, y + pad + nh + 0.08, cw - 2 * pad, step.get('text', ''), z,
                          line=1.36, emph='ink')
                if i < n - 1:
                    self.line(slide, cx + cw + 0.08, y + pad + d / 2, cx + cw + gapw - 0.08, y + pad + d / 2,
                              'muted', 1.25, head=True)
        return max(heights)

    def _formula(self, slide, f: dict, x, y, w, draw=False) -> float:
        stacked = w < 7  # narrow panels put the note under the formula instead of beside it
        fw = w if stacked else w * 0.5
        nw = w - 0.5 if stacked else w - fw - 0.5
        fh = self.height(f['text'], 18, fw - 0.5, role='math', line=1.3)
        nh = self.height(f.get('note', ''), 13.5, nw, line=1.42) if f.get('note') else 0
        h = (fh + (nh + 0.14 if nh else 0) if stacked else max(fh, nh)) + 0.38
        if draw:
            self.rect(slide, x, y, w, h, 'panel', radius=0.06)
            fy = y + 0.19 if stacked else y + (h - fh) / 2
            self.text(slide, x + 0.25, fy, fw - 0.5, f['text'], 18, role='math', color='ink', line=1.3)
            if nh:
                nx, ny = (x + 0.25, fy + fh + 0.14) if stacked else (x + fw + 0.25, y + (h - nh) / 2)
                if not stacked:
                    self.line(slide, x + fw, y + 0.18, x + fw, y + h - 0.18, 'faint')
                self.text(slide, nx, ny, nw, f['note'], 13.5, line=1.42, emph='ink')
        return h

    def method(self, sd: dict) -> None:
        slide = self.new_slide()
        y = self.chrome(slide, sd, 'method')
        if sd.get('lede'):
            size, h = self.fit(sd['lede'], W - 2 * MX, 0.9, 17, 15, 'method lede', role='serif_text', line=1.42)
            y += self.text(slide, MX, y, W - 2 * MX, sd['lede'], size, role='serif_text', color='ink',
                           line=1.42, emph='clay_deep') + 0.28
        lanes = sd.get('lanes') or [dict(steps=sd.get('steps', []))]
        if sd.get('figure'):
            self._method_figure(slide, sd, lanes, y)
            return
        fh = self._formula(slide, sd['formula'], MX, 0, W - 2 * MX) if sd.get('formula') else 0
        area = BOTTOM - y - (fh + 0.28 if fh else 0)
        lw = 0.95 if any(lane.get('label') for lane in lanes) else 0
        rows_w = W - 2 * MX - lw

        def total(z):
            return sum(self._steps_row(slide, ln['steps'], 0, 0, rows_w, z) for ln in lanes) + \
                0.26 * (len(lanes) - 1)

        z = self.largest(total, 14.5, area, 'method steps')
        ry = y + max(0, (area - total(z)) / 2)
        for lane in lanes:
            rh = self._steps_row(slide, lane['steps'], 0, 0, rows_w, z)
            if lane.get('label'):
                self.text(slide, MX, ry, lw - 0.15, f"**{lane['label']}**", 13, color='muted', emph='ink2',
                          h=rh, anchor='m', line=1.3)
            self._steps_row(slide, lane['steps'], MX + lw, ry, rows_w, z, True, card_h=rh)
            ry += rh + 0.26
        if fh:
            self._formula(slide, sd['formula'], MX, BOTTOM - fh, W - 2 * MX, True)

    def _method_figure(self, slide, sd, lanes, y) -> None:
        fig = sd['figure']
        fx, fw = MX, 6.55
        sx = fx + fw + 0.45
        sw = W - MX - sx
        cap_h = self.text(slide, fx, 0, fw, fig.get('caption', ''), 10.5, color='muted', line=1.3,
                          draw=False) if fig.get('caption') else 0
        px, py, pw, ph = self.picture(slide, self.image_bytes(fig['path']), fx, y, fw,
                                      BOTTOM - y - cap_h - 0.12, valign='t', border='hair')
        if fig.get('caption'):
            self.text(slide, fx, py + ph + 0.12, fw, fig['caption'], 10.5, color='muted', line=1.3)
        steps = [s for lane in lanes for s in lane['steps']]
        fh = self._formula(slide, sd['formula'], sx, 0, sw) if sd.get('formula') else 0
        d = 0.3

        def column(z, draw=False):
            yy = y
            for i, step in enumerate(steps):
                if i:
                    if draw:
                        self.line(slide, sx, yy + 0.13, W - MX, yy + 0.13)
                    yy += 0.26
                if draw:
                    self.badge(slide, sx, yy, d, step.get('badge', i + 1), 'clay' if step.get('key') else 'ink2')
                nh = self.text(slide, sx + d + 0.15, yy + 0.02, sw - d - 0.15, f"**{step['name']}**", z + 0.5,
                               color='ink', emph='ink', line=1.25, draw=draw)
                yy += max(d, nh) + 0.06
                yy += self.text(slide, sx + d + 0.15, yy, sw - d - 0.15, step.get('text', ''), z, line=1.38,
                                emph='ink', draw=draw)
            return yy - y

        z = self.largest(column, 14.5, BOTTOM - y - (fh + 0.28 if fh else 0), 'method steps')
        column(z, True)
        if fh:
            self._formula(slide, sd['formula'], sx, BOTTOM - fh, sw, True)

    def innovation(self, sd: dict) -> None:
        slide, L = self.new_slide(), self.L
        top = self.chrome(slide, sd, 'innovation')
        c1, w1, c2, w2 = MX, 2.1, 3.0, 3.25  # the "this paper" column carries the most text
        c3 = c2 + w2 + 0.5
        w3 = W - MX - c3
        items = sd.get('items', [])
        summary_h = self.height(sd['summary'], 15.5, W - 2 * MX, role='serif_text', line=1.4) + 0.3 \
            if sd.get('summary') else 0
        head_h = 12 * 1.2 / 72 + 0.2

        def rows(z, draw=False):
            y = top + head_h
            for i, item in enumerate(items):
                body3 = [b for b in (item.get('after'), f"**{L['why']}**　{item['why']}" if item.get('why') else None)
                         if b]
                rh = 0.3 + max(0.4 + self.height(f"**{item['name']}**", z + 1.5, w1, line=1.3),
                                self.height(item.get('before', ''), z, w2, line=1.4),
                                self.height(body3, z, w3 - 0.5, line=1.4, gap=0.4))
                if draw:
                    self.text(slide, c1, y + 0.16, w1, f'{i + 1:02d}', 20, role='serif', color='clay', line=1.0)
                    self.text(slide, c1, y + 0.56, w1, f"**{item['name']}**", z + 1.5, color='ink', emph='ink',
                              line=1.3)
                    self.text(slide, c2, y + 0.16, w2, item.get('before', ''), z, line=1.4, emph='ink')
                    self.line(slide, c2 + w2 + 0.1, y + rh / 2, c3 - 0.1, y + rh / 2, 'clay', 1.5, head=True)
                    self.text(slide, c3 + 0.25, y + 0.16, w3 - 0.5, body3, z, color='ink', line=1.4, gap=0.4,
                              emph='clay_deep')
                    if i < len(items) - 1:
                        self.line(slide, c1, y + rh, c2 + w2, y + rh)
                        self.line(slide, c3 + 0.25, y + rh, W - MX - 0.25, y + rh, 'white', 1)
                y += rh
            return y - top

        z = self.largest(rows, 15, BOTTOM - top - summary_h, 'innovation rows')
        body_h = rows(z)
        self.rect(slide, c3, top - 0.08, w3, body_h + 0.16, 'clay_tint', radius=0.05)
        self.label(slide, c1, top, w1, L['innov'], color='muted')
        self.label(slide, c2, top, w2, L['before'], color='slate')
        self.label(slide, c3 + 0.25, top + 0.04, w3 - 0.5, L['after'])
        rows(z, True)
        if sd.get('summary'):
            y = top + body_h + 0.3
            self.line(slide, MX, y + 0.02, MX + 0.5, y + 0.02, 'clay', 2)
            self.text(slide, MX, y + 0.16, W - 2 * MX, sd['summary'], 15.5, role='serif_text', color='ink',
                      line=1.4, emph='clay_deep')

    def _table(self, slide, spec, x, y, w, row_h) -> float:
        cols, rows = spec['columns'], spec['rows']
        size = min(16, max(13, row_h * 72 / 2.3))  # taller rows get larger, easier-to-read figures
        hot = spec.get('highlight_row')
        widths = spec.get('widths') or [2.2] + [1] * (len(cols) - 1)
        frame = slide.shapes.add_table(len(rows) + 1, len(cols), Inches(x), Inches(y), Inches(w),
                                       Inches(row_h * (len(rows) + 1)))
        table = frame.table
        tbl_pr = table._tbl.tblPr
        tbl_pr.set('firstRow', '0')
        tbl_pr.set('bandRow', '0')
        style_id = tbl_pr.find(qn('a:tableStyleId'))
        if style_id is None:
            style_id = etree.SubElement(tbl_pr, qn('a:tableStyleId'))
        style_id.text = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'  # "No Style, No Grid"
        for j, cwid in enumerate(widths):
            table.columns[j].width = Inches(w * cwid / sum(widths))
        last = len(rows)
        for i in range(last + 1):
            table.rows[i].height = Inches(row_h)
            for j in range(len(cols)):
                cell = table.cell(i, j)
                cell.margin_left = cell.margin_right = Inches(0.1)
                cell.margin_top = cell.margin_bottom = Inches(0.02)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                lit = i > 0 and i - 1 == hot
                tc_pr = cell._tc.get_or_add_tcPr()
                rules = {'T': ('ink', 1.25) if i == 0 else None,  # booktabs: top, mid and bottom rules
                         'B': ('ink', 1.25) if i == last else ('ink2', 0.75) if i == 0 else ('hair', 0.5)}
                for pos, side in enumerate('LRTB'):
                    rule = rules.get(side)
                    ln = etree.Element(qn(f'a:ln{side}'), w=str(int((rule[1] if rule else 0) * 12700)))
                    if rule:
                        etree.SubElement(etree.SubElement(ln, qn('a:solidFill')), qn('a:srgbClr'), val=C[rule[0]])
                    else:
                        etree.SubElement(ln, qn('a:noFill'))
                    tc_pr.insert(pos, ln)
                if lit:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb('clay_tint')
                else:
                    cell.fill.background()
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
                run = p.add_run()
                run.text = str(cols[j] if i == 0 else rows[i - 1][j])
                if i == 0:
                    self.style_run(run, 'sans', size - 1.5, 'muted', bold=True)
                else:
                    self.style_run(run, 'sans', size, 'clay_deep' if lit else 'ink2', bold=lit)
        return row_h * (len(rows) + 1)

    def _chart(self, slide, spec, x, y, w, h) -> None:
        data = CategoryChartData()
        data.categories = spec['categories']
        data.add_series(spec.get('series', ''), spec['values'])
        chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h),
                                       data).chart
        chart.has_legend = chart.has_title = False
        plot = chart.plots[0]
        plot.gap_width, plot.vary_by_categories, plot.has_data_labels = 55, False, True
        labels = plot.data_labels
        labels.number_format, labels.number_format_is_linked = spec.get('number_format', 'General'), False
        labels.position = XL_LABEL_POSITION.OUTSIDE_END
        for i, cat in enumerate(spec['categories']):
            fill = plot.series[0].points[i].format.fill
            fill.solid()
            fill.fore_color.rgb = rgb('clay' if cat == spec.get('highlight') else 'faint')
        axis = chart.category_axis
        axis.reverse_order, axis.major_tick_mark = True, XL_TICK_MARK.NONE
        axis.format.line.color.rgb = rgb('hair')
        chart.value_axis.visible = False
        chart.value_axis.has_major_gridlines = False
        chart.font.size, chart.font.color.rgb = Pt(13), rgb('ink2')
        chart.font.name = self.fonts.face['sans']
        etree.SubElement(chart.font._rPr, qn('a:ea')).set('typeface', self.fonts.face['sans'])

    def evidence(self, sd: dict) -> None:
        slide, L = self.new_slide(), self.L
        top = self.chrome(slide, sd, 'evidence')
        vx, vw = MX, 7.0
        rx = vx + vw + 0.5
        rw = W - MX - rx
        cap = sd.get('caption', '')
        cap_h = self.text(slide, vx, 0, vw, cap, 10.5, color='muted', line=1.35, draw=False) + 0.14 if cap else 0
        room = BOTTOM - top - cap_h
        if sd.get('table'):
            n = len(sd['table']['rows']) + 1
            vh = self._table(slide, sd['table'], vx, top, vw, min(0.6, max(0.4, room / n)))
            cap_y = top + vh + 0.14
        elif sd.get('chart'):
            self._chart(slide, sd['chart'], vx, top, vw, room)
            cap_y = top + room + 0.14
        elif sd.get('figure'):
            _, py, _, ph = self.picture(slide, self.image_bytes(sd['figure']), vx, top, vw, room, valign='t',
                                        border='hair')
            cap_y = py + ph + 0.14
        else:
            cap_y = top
            self.warn('evidence slide has no table, chart or figure')
        if cap:
            self.text(slide, vx, cap_y, vw, cap, 10.5, color='muted', line=1.35)
        points, boundary = sd.get('points', []), sd.get('boundary')
        note = [f"**{L['boundary']}**", boundary] if boundary else None

        def column(z, draw=False):
            y = top
            for i, point in enumerate(points):
                if i:
                    if draw:
                        self.line(slide, rx, y + 0.14, W - MX, y + 0.14)
                    y += 0.28
                if draw:
                    self.text(slide, rx, y, 0.5, f'{i + 1:02d}', 17, role='serif', color='clay', line=1.1)
                y += self.text(slide, rx + 0.55, y, rw - 0.55, point, z, line=1.42, emph='ink', draw=draw)
            if note:
                bh = self.text(slide, 0, 0, rw - 0.4, note, z - 1.5, line=1.4, gap=0.25, draw=False) + 0.36
                if draw:
                    self.rect(slide, rx, BOTTOM - bh, rw, bh, 'panel', radius=0.06)
                    self.text(slide, rx + 0.2, BOTTOM - bh + 0.18, rw - 0.4, note, z - 1.5, line=1.4, gap=0.25,
                              emph='ink')
                y += 0.3 + bh
            return y - top

        column(self.largest(column, 15, BOTTOM - top, 'evidence column'), True)

    # ---------- build ----------
    def build(self, out: Path) -> None:
        slides = self.spec['slides']
        for i, sd in enumerate(slides):
            layout = sd.get('layout')
            if layout not in ('cover', 'background', 'method', 'innovation', 'evidence'):
                raise ValueError(f'slide {i + 1}: unknown layout {layout!r}')
            if layout == 'cover' or i == 0:
                end = next((j for j in range(i + 1, len(slides)) if slides[j].get('layout') == 'cover'), len(slides))
                self.group = [s['layout'] for s in slides[i:end] if s['layout'] != 'cover']
            getattr(self, layout)(sd)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(out)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('spec', type=Path, help='deck spec JSON (references/deck-spec.md)')
    parser.add_argument('-o', '--output', type=Path, help='output .pptx (default: spec "output")')
    args = parser.parse_args(argv)
    spec = json.loads(args.spec.read_text(encoding='utf-8'))
    base = args.spec.resolve().parent
    out = args.output or (base / spec['output'] if spec.get('output') else None)
    if out is None:
        parser.error('give -o or an "output" field in the spec')
    deck = Deck(spec, base)
    deck.build(out.resolve())
    fonts = ', '.join(f'{role}={face}' + ('' if deck.fonts.chosen[role] else ' (not installed)')
                      for role, face in deck.fonts.face.items())
    print(f'Saved {deck.page} slides: {out.resolve()}\nFonts: {fonts}')
    for message in deck.warnings:
        print(f'WARNING {message}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
