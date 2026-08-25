# -*- coding: utf-8 -*-
"""Teal study-kit slide helpers.

Design system = the lab's teal "v7" palette + Malgun Gothic (ppt_rules.md), ported into the
teaching-slide helper API for lecture decks (section/content/code/table/image/bullet).
Kept local on purpose: a lecture project's own ppt_common.py is a different (warm/slate) palette and is not edited.

Korean font floors (ppt_rules): title bar 30-32pt, table header 17pt, body 16pt, secondary 15pt,
never below 13pt.
"""
import sys, io
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn as _qn
from lxml import etree

# shared, palette-agnostic mechanics — single source of truth for the fiddly bits (notes,
# template load + slide clear, blank-layout lookup, overflow-safe images). Canonical home is the
# pptx-editing skill; project style stays below.
# 260806: an absolute path was hardcoded here, which breaks on any machine with a different
# drive/install location. Resolve the standard skill location instead (override: PPTX_KIT).
import os as _os


def _skill_scripts_dir():
    env = _os.environ.get("PPTX_KIT")
    if env:                                   # may point at pptx_kit.py or at its folder
        return _os.path.dirname(env) if env.lower().endswith(".py") else env
    base = _os.environ.get("CLAUDE_SKILLS_DIR") or _os.path.join(
        _os.path.expanduser("~"), ".claude", "skills")
    return _os.path.join(base, "pptx-editing", "scripts")


_scripts = _skill_scripts_dir()
if not _os.path.isfile(_os.path.join(_scripts, "pptx_kit.py")):
    raise ImportError(
        f"pptx_kit.py not found under {_scripts} — install the pptx-editing skill "
        "(repo: python sync_skill.py --install), or set PPTX_KIT to its path.")
sys.path.insert(0, _scripts)
from pptx_kit import (new_deck as _new_deck_core, blank_slide_layout as _blank,
                      rect, slide_number as _slide_number_core, speaker_note,
                      fit_picture, overflows)

# ── teal v7 palette (create_resub_pptx.py / ppt_rules.md) ──
TEAL      = RGBColor(0x2C, 0x7A, 0x7B)
TEAL_DARK = RGBColor(0x12, 0x34, 0x3B)
TEAL_CARD = RGBColor(0x1B, 0x4A, 0x50)
TEAL_ACC  = RGBColor(0x9F, 0xD0, 0xCB)
TEAL_SUB  = RGBColor(0x7F, 0xB3, 0xAE)
SAGE      = RGBColor(0xAF, 0xC4, 0xC2)
GREY      = RGBColor(0x6B, 0x7B, 0x85)
BODY      = RGBColor(0x23, 0x30, 0x3A)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
RED       = RGBColor(0xC7, 0x5D, 0x4F)
RED_BG    = RGBColor(0xFB, 0xEE, 0xEA)
PANEL     = RGBColor(0xF4, 0xF8, 0xF7)   # very light teal panel
CODEBG    = RGBColor(0xEE, 0xF3, 0xF2)
RULE_L    = RGBColor(0xD5, 0xDE, 0xDD)

# accent cycle, teal-harmonised
ACC = [TEAL, RED, RGBColor(0x4A, 0x8F, 0xA8), RGBColor(0xB0, 0x8A, 0x3E), GREY, TEAL_CARD]

FK = "Malgun Gothic"   # Korean body
FC = "Consolas"        # code / formulas

SW = None; SH = None
_TEMPLATE = None


def new_deck(template=None):
    """Load the teal v7 template (inherits theme/master), clear all slides, track page size.
    Mechanics live in pptx_kit.new_deck; here we just record SW/SH for the styled components."""
    global SW, SH
    prs = _new_deck_core(template)
    SW, SH = prs.slide_width, prs.slide_height
    return prs


# _blank (blank-layout lookup) and rect are imported from pptx_kit above.

# ── primitives ──
def tb(s, l, t, w, h, text, sz=16, bold=False, color=BODY, align=PP_ALIGN.LEFT, font=FK):
    box = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text
    p.font.size = Pt(sz); p.font.bold = bold; p.font.color.rgb = color
    p.font.name = font; p.alignment = align
    return box


def ap(tf, text, sz=16, bold=False, color=BODY, font=FK, align=PP_ALIGN.LEFT, spc=4):
    p = tf.add_paragraph(); p.text = text
    p.font.size = Pt(sz); p.font.bold = bold; p.font.color.rgb = color
    p.font.name = font; p.alignment = align
    if spc: p.space_before = Pt(spc)
    return p


def _bg(s, color=WHITE):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = color; r.line.fill.background()


def _slide_num(prs, s):
    return _slide_number_core(prs, s, color=GREY, font=FK)


# ── slide templates ──
def cover(prs, kicker, title, subtitle=""):
    s = prs.slides.add_slide(_blank(prs))
    _bg(s, TEAL_DARK)
    rect(s, 1.1, 2.05, 3.2, 0.10, TEAL)
    tb(s, 1.1, 1.45, 11.0, 0.6, kicker, sz=22, bold=True, color=TEAL_SUB)
    tb(s, 1.1, 2.35, 11.2, 1.9, title, sz=40, bold=True, color=WHITE)
    if subtitle:
        tb(s, 1.1, 4.75, 11.2, 1.1, subtitle, sz=18, color=SAGE)
    return s


def section_slide(prs, roman, title, subtitle=""):
    s = prs.slides.add_slide(_blank(prs))
    _bg(s, TEAL_DARK)
    tb(s, 1.1, 1.55, 4.0, 1.9, roman, sz=96, bold=True, color=TEAL)
    rect(s, 1.15, 3.62, 3.0, 0.09, TEAL_ACC)
    tb(s, 1.1, 3.85, 11.0, 1.0, title, sz=38, bold=True, color=WHITE)
    if subtitle:
        tb(s, 1.1, 4.85, 11.0, 0.8, subtitle, sz=17, color=SAGE)
    return s


def content_slide(prs, title, breadcrumb=""):
    """White slide with a small teal dot + breadcrumb + title (v7 content header)."""
    s = prs.slides.add_slide(_blank(prs))
    _bg(s, WHITE)
    d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.72), Inches(0.52), Inches(0.16), Inches(0.16))
    d.fill.solid(); d.fill.fore_color.rgb = TEAL; d.line.fill.background()
    if breadcrumb:
        tb(s, 0.98, 0.44, 8.0, 0.32, breadcrumb, sz=12, color=GREY)
    tb(s, 0.70, 0.78, 12.0, 0.62, title, sz=26, bold=True, color=TEAL_DARK)
    rect(s, 0.72, 1.44, 11.9, 0.03, RULE_L)
    _slide_num(prs, s)
    return s


def _hang(p, width):
    """Hanging indent of `width` on one paragraph (python-pptx exposes no API for it)."""
    pPr = p._pPr if p._pPr is not None else p._p.get_or_add_pPr()
    pPr.set("marL", str(int(width)))
    pPr.set("indent", str(int(-width)))


def bullet(s, l, t, w, h, items, sz=16, gap=7, ci=0):
    """items: [str] or [(str, accent_index)] or (True, 'header') for a sub-header row."""
    box = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for it in items:
        if isinstance(it, tuple) and it and it[0] is True:      # header row
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            r = p.add_run(); r.text = it[1]
            r.font.size = Pt(sz); r.font.bold = True; r.font.color.rgb = TEAL_DARK; r.font.name = FK
            if not first: p.space_before = Pt(gap + 4)
            first = False; continue
        if isinstance(it, tuple):
            text, c = it
        else:
            text, c = it, ci
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        r1 = p.add_run(); r1.text = "•  "
        r1.font.size = Pt(sz); r1.font.color.rgb = ACC[c % len(ACC)]; r1.font.name = FK
        r2 = p.add_run(); r2.text = text
        r2.font.size = Pt(sz); r2.font.color.rgb = BODY; r2.font.name = FK
        # hang the wrapped lines under the text, not back at the marker: the bullet is an
        # inline run, so without this a two-line item reads as two separate items
        _hang(p, Pt(sz) * 0.95)
        if not first: p.space_before = Pt(gap)
        first = False
    return box


def card(s, l, t, w, h, header, lines, header_color=TEAL, body_bg=PANEL, sz=15):
    """Rounded panel with a coloured header band + left accent bar."""
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = body_bg
    box.line.color.rgb = header_color; box.line.width = Pt(1.0)
    box.adjustments[0] = 0.06
    hb = rect(s, l, t, w, 0.42, header_color)
    hb.adjustments  # header band
    tfh = hb.text_frame; tfh.vertical_anchor = MSO_ANCHOR.MIDDLE
    tfh.margin_left = Inches(0.16)
    ph = tfh.paragraphs[0]; ph.text = header
    ph.font.size = Pt(sz + 1); ph.font.bold = True; ph.font.color.rgb = WHITE; ph.font.name = FK
    bx = s.shapes.add_textbox(Inches(l + 0.16), Inches(t + 0.52), Inches(w - 0.32), Inches(h - 0.64))
    tf = bx.text_frame; tf.word_wrap = True
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ln; p.font.size = Pt(sz); p.font.color.rgb = BODY; p.font.name = FK
        if i: p.space_before = Pt(4)
    return box


def formula(s, l, t, w, h, lines, sz=17, bg=CODEBG, accent=TEAL):
    """Monospace/maths block (unicode maths, not LaTeX) with a left accent bar."""
    rect(s, l, t, w, h, bg)
    rect(s, l, t, 0.07, h, accent)
    bx = s.shapes.add_textbox(Inches(l + 0.22), Inches(t + 0.10), Inches(w - 0.36), Inches(h - 0.20))
    tf = bx.text_frame; tf.word_wrap = True
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ln if ln else " "
        p.font.size = Pt(sz); p.font.name = FC; p.font.color.rgb = BODY
        p.line_spacing = 1.15
    return bx


def code_block(s, l, t, w, h, code, sz=13, accent=TEAL):
    return formula(s, l, t, w, h, code.split("\n"), sz=sz, accent=accent)


def table_block(s, headers, rows, col_x, col_w=None, y0=1.75, y1=6.80, sz=15, bold=None,
                right=12.6):
    """Draw table_slide's table INTO AN EXPLICIT y-range on a slide you already made.

    Use this whenever anything else goes on the slide below the table — a card, a second
    table, a diagram. `table_slide` sizes its rows from ALL the space left on the slide, so a
    shape placed under it silently ends up underneath the last rows; the collision is invisible
    to a coordinate check and only `audit_text_fit.py` or a render catches it (a deck build
    260819: four slides at once). Returns the y of the closing rule.

    `right` is the table's right edge. Set it when placing two tables SIDE BY SIDE — otherwise
    both draw their rules out to the slide margin and the left table's row lines run underneath
    the right one, which looks like stray lines through its rows (2026-08-20).
    """
    L = col_x[0]; W = right - L
    rect(s, L, y0, W, 0.035, TEAL_DARK)
    for j, h in enumerate(headers):
        tb(s, col_x[j], y0 + 0.10, (col_w[j] if col_w else 3.0), 0.42, h,
           sz=sz + 1, bold=True, color=TEAL_DARK)
    rect(s, L, y0 + 0.62, W, 0.02, TEAL_DARK)
    avail = max(0.6, y1 - (y0 + 0.74))
    rh = min(0.66, max(0.28, avail / max(1, len(rows))))
    y = y0 + 0.74
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            b = bool(bold[i][j]) if bold else False
            tb(s, col_x[j], y, (col_w[j] if col_w else 3.0), rh, str(cell), sz=sz, color=BODY, bold=b)
        if i < len(rows) - 1:
            rect(s, L, y + rh - 0.03, W, 0.012, RULE_L)
        y += rh
    rect(s, L, y, W, 0.035, TEAL_DARK)
    return y + 0.035


def table_slide(prs, title, headers, rows, col_x, col_w=None, sz=15, breadcrumb="", note=None, top=1.75, bold=None):
    """Whole slide = one table (+ optional grey note under it).

    ⚠ The rows EXPAND to fill everything left on the slide, so this cannot share a slide with
    anything below the table — use `table_block` for that. bold: optional matrix (list of list
    of bool) matching `rows`' shape — True bolds that cell.
    """
    s = content_slide(prs, title, breadcrumb)
    L = col_x[0]; W = 12.6 - L
    rect(s, L, top, W, 0.035, TEAL_DARK)
    for j, h in enumerate(headers):
        tb(s, col_x[j], top + 0.10, (col_w[j] if col_w else 3.0), 0.42, h,
           sz=sz + 2, bold=True, color=TEAL_DARK)
    rect(s, L, top + 0.62, W, 0.02, TEAL_DARK)
    # derive row height from the space actually left above the slide-number band (6.80in)
    avail = max(1.2, 6.80 - (top + 0.74) - (0.75 if note else 0.0))
    rh = min(0.66, max(0.30, avail / max(1, len(rows))))
    y = top + 0.74
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            b = bool(bold[i][j]) if bold else False
            tb(s, col_x[j], y, (col_w[j] if col_w else 3.0), rh, str(cell), sz=sz, color=BODY, bold=b)
        if i < len(rows) - 1:
            rect(s, L, y + rh - 0.03, W, 0.012, RULE_L)
        y += rh
    rect(s, L, y, W, 0.035, TEAL_DARK)
    if note:
        tb(s, L, y + 0.14, W, 0.7, note, sz=13, color=GREY)
    return s


def image_slide(prs, title, img_path, caption=None, breadcrumb="", max_w=11.2, max_h=4.6, top=1.72):
    from PIL import Image
    import os
    s = content_slide(prs, title, breadcrumb)
    if not os.path.exists(img_path):
        tb(s, 0.8, 3.0, 11.5, 0.6, f"[missing figure: {img_path}]", sz=15, color=RED)
        return s
    iw, ih = Image.open(img_path).size
    ar = iw / ih
    w = max_w; h = w / ar
    if h > max_h:
        h = max_h; w = h * ar
    left = (SW / 914400.0 - w) / 2.0
    s.shapes.add_picture(img_path, Inches(left), Inches(top), Inches(w), Inches(h))
    if caption:
        tb(s, 0.8, top + h + 0.10, 11.7, 0.75, caption, sz=13, color=GREY, align=PP_ALIGN.CENTER)
    return s


def two_col(prs, title, left_head, left_lines, right_head, right_lines,
            breadcrumb="", lead=None, sz=15):
    """Concept slide: optional one-line lead + two labelled cards."""
    s = content_slide(prs, title, breadcrumb)
    y = 1.68
    if lead:
        tb(s, 0.75, 1.58, 11.9, 0.5, lead, sz=17, bold=True, color=TEAL_DARK)
        y = 2.22
    # size the cards to their content. Korean glyphs are ~2x the width of Latin, so measure in
    # half-width units and add wrapped rows; pad generously so a slight underestimate never clips.
    def _units(ln):
        return sum(2 if ("가" <= ch <= "힣" or "ㄱ" <= ch <= "ㆎ") else 1 for ch in ln)
    def _rows(lines, wrap_units=50):
        return sum(1 + max(0, (_units(ln) - 1) // wrap_units) for ln in lines)
    n = max(_rows(left_lines), _rows(right_lines))
    h = min(6.75 - y, 0.82 + n * 0.34)
    card(s, 0.75, y, 5.85, h, left_head, left_lines, header_color=TEAL, sz=sz)
    card(s, 6.78, y, 5.85, h, right_head, right_lines, header_color=ACC[2], sz=sz)
    return s


def node(s, x, y, w, h, label, color=TEAL, txt=WHITE, sz=14):
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = color; box.line.fill.background()
    tf = box.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = label; p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(sz); p.font.bold = True; p.font.color.rgb = txt; p.font.name = FK
    return box


def arrow(s, x1, y1, x2, y2, color=GREY, width=2.0):
    cnx = s.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cnx.line.color.rgb = color; cnx.line.width = Pt(width)
    spPr = cnx._element.find(_qn('p:spPr'))
    ln = spPr.find(_qn('a:ln'))
    if ln is None:
        ln = etree.SubElement(spPr, _qn('a:ln'))
    etree.SubElement(ln, _qn('a:tailEnd'), {'type': 'arrow', 'w': 'lg', 'len': 'lg'})
    return cnx


def footnote(s, text, y=6.92):
    tb(s, 0.75, y, 11.9, 0.42, text, sz=13, color=GREY)


# speaker_note is imported from pptx_kit (shared mechanics) at the top of this file.
