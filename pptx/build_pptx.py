"""Build the SOLABIMA 2026 PowerPoint poster template.

Generates solabima2026_poster_template.pptx: an A0 portrait poster
(841 x 1189 mm) matching the design of the LaTeX template in ../latex
— a boxless header (title/subtitle/authors/affiliation in blue,
directly on the white page, centered in a 70%-wide column and
flanked by the SOLABIMA logo and a placeholder institution logo),
two equal content columns of white cards with an orange-tinted
border (only the top-left and bottom-right corners rounded) holding
blue bold-sans section titles over black serif body text, and an
optional funding/sponsor-logo band near the bottom. Edit this script
and re-run it to regenerate the file, or edit the .pptx directly in
PowerPoint.

Font sizes below are taken directly from the LaTeX template's own
\\renewcommand redefinitions of \\tiny..\\Huge (a real point size,
identical unit to PowerPoint's pt) and the explicit \\Huge/\\Large/
\\large/\\small/\\footnotesize commands used per element in
../latex/solabima2026_poster.tex, then rounded down a few percent as
a safety margin for Georgia/Arial's slightly wider metrics than Latin
Modern. An earlier version of this script guessed much smaller sizes
(e.g. 16-17pt body text vs. the ~31pt the LaTeX source actually uses)
which, on this page's real A0 physical dimensions, produced correct
card heights on paper but wildly undersized, sparse-looking text —
confirmed by rendering both versions to comparable-resolution PNGs
(pdftoppm for the LaTeX PDF, `qlmanage -t` for this file) and
measuring line pitch in pixels. Each block's height is now computed
from its actual (measured) text/image content via MeasureCanvas
below, rather than a hand-guessed constant, so it won't leave the
large dead gaps that guessing produced either.

Corner rounding relies on PowerPoint's built-in "Round Diagonal
Corner Rectangle" autoshape, which by default rounds the top-left and
bottom-right corners (matching the Office shape-gallery icon). Note:
macOS QuickLook's .pptx thumbnail renderer does not draw this preset
at all (confirmed with an isolated test file) and seems to ignore
custom line widths on shapes generally — so a QuickLook preview of
this file will show no card borders regardless of how the file is
generated. That appears to be a QuickLook-specific limitation (the
shape's XML — fill, line color, line width — is well-formed OOXML),
not a bug in the generated file, but it means QuickLook can't be used
to visually confirm the borders render; only real PowerPoint/Keynote/
LibreOffice can. This was written without one of those available on
the build machine.

Run:  python3 build_pptx.py
"""
import os
from PIL import Image, ImageFont
from pptx import Presentation
from pptx.util import Mm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
LOGOS_DIR = os.path.join(HERE, "..", "latex", "logos")
ASSETS_DIR = os.path.join(HERE, "assets")
OUT_PATH = os.path.join(HERE, "solabima2026_poster_template.pptx")

# ---------------------------------------------------------------
# Palette (matches latex/solabima2026_poster.tex, sampled from the
# real SOLABIMA 2026 logo)
# ---------------------------------------------------------------
BLUE = RGBColor(0x1C, 0x6E, 0xA8)
ORANGE = RGBColor(0xF1, 0x96, 0x51)
ORANGE_DARK = RGBColor(0xC2, 0x66, 0x1F)
BORDER_ORANGE = RGBColor(0xF7, 0xC5, 0x9F)  # solabimaOrange!55!white
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
BLUE_TINT_5 = RGBColor(0xF4, 0xF8, 0xFB)   # ~5% blue over white
BLUE_TINT_8 = RGBColor(0xED, 0xF3, 0xF8)   # ~8% blue over white

FONT_SERIF = "Georgia"   # body text (matches lmodern serif in LaTeX)
FONT_SANS = "Arial"      # headers + title (matches sans-serif in LaTeX)

# Real font files, used only to measure text so blocks can be sized
# to their actual content instead of a guessed constant.
_FONT_FILES = {
    (FONT_SERIF, False): "/System/Library/Fonts/Supplemental/Georgia.ttf",
    (FONT_SERIF, True): "/System/Library/Fonts/Supplemental/Georgia Bold.ttf",
    (FONT_SANS, False): "/System/Library/Fonts/Supplemental/Arial.ttf",
    (FONT_SANS, True): "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
}
_FONT_CACHE = {}

MM_PER_PT = 25.4 / 72.0

PAGE_W = Mm(841)
PAGE_H = Mm(1189)
MARGIN = Mm(15)
COL_GAP = Mm(25)
N_COLS = 2
COL_W = int((PAGE_W - 2 * MARGIN - (N_COLS - 1) * COL_GAP) / N_COLS)
COL_W_MM = COL_W / Mm(1)
GAP = Mm(19)  # vertical gap between stacked blocks in a column (measured from the LaTeX PDF)

BODY_PAD = Mm(14)
BODY_PAD_MM = 14.0
RADIUS_MM = 7.0  # absolute target corner radius, matches \SolabimaR in LaTeX
CARD_BORDER_W = Pt(4)

# Font sizes: LaTeX's own \tiny..\Huge point values (real pt, same
# unit as PowerPoint) rounded down ~5-10% as a safety margin against
# Georgia/Arial running a bit wider than Latin Modern at equal size.
BODY_SIZE = 29        # \normalsize (31pt) - paragraphs, bullets
CAPTION_SIZE = 24     # \small (26pt), bold - figure/table/equation captions
REF_SIZE = 20         # \footnotesize (22pt) - referencias
ACK_SIZE = 24         # \small (26pt) - agradecimientos
TABLE_SIZE = 21       # table cells
SECTION_TITLE_SIZE = 48   # tikzposter block-title default (~\LARGE, 54pt)
HEADER_TITLE_SIZE = 72    # \Huge (77pt)
HEADER_SUBTITLE_SIZE = 40  # \Large (45pt)
HEADER_AUTHORS_SIZE = 40   # \Large (45pt)
HEADER_INSTITUTE_SIZE = 32  # \large (37pt)


def col_x(i):
    return MARGIN + i * (COL_W + COL_GAP)


def _pil_font(font_name, bold, size_pt):
    key = (font_name, bold, round(size_pt))
    f = _FONT_CACHE.get(key)
    if f is None:
        f = ImageFont.truetype(_FONT_FILES[(font_name, bold)], max(1, round(size_pt)))
        _FONT_CACHE[key] = f
    return f


def _wrap_lines(text, font, max_width_pt):
    words = text.split()
    if not words:
        return [""]
    lines = []
    cur = words[0]
    for word in words[1:]:
        trial = cur + " " + word
        if font.getlength(trial) <= max_width_pt:
            cur = trial
        else:
            lines.append(cur)
            cur = word
    lines.append(cur)
    return lines


def text_height_mm(text, size_pt, width_mm, font=FONT_SERIF, bold=False, line_spacing=1.15):
    """Height (mm) a text box needs for `text` word-wrapped to width_mm."""
    f = _pil_font(font, bold, size_pt)
    max_w_pt = width_mm / MM_PER_PT
    n_lines = 0
    for para in text.split("\n"):
        n_lines += len(_wrap_lines(para, f, max_w_pt)) if para else 1
    return n_lines * size_pt * line_spacing * MM_PER_PT


def bullets_height_mm(items, size_pt, width_mm, font=FONT_SERIF, line_spacing=1.15,
                       indent_mm=8, space_after_mm=3):
    f = _pil_font(font, False, size_pt)
    max_w_pt = (width_mm - indent_mm) / MM_PER_PT
    line_pitch_mm = size_pt * line_spacing * MM_PER_PT
    total = 0.0
    for i, item in enumerate(items):
        total += len(_wrap_lines(item, f, max_w_pt)) * line_pitch_mm
        if i < len(items) - 1:
            total += space_after_mm
    return total


def picture_height_mm_for_width(path, width_mm):
    with Image.open(path) as im:
        px_w, px_h = im.size
    return width_mm * (px_h / px_w)


def radius_adj(w, h, target_mm=RADIUS_MM):
    """adj fraction (of the shape's shorter side) giving an absolute
    radius of target_mm, regardless of this shape's own aspect ratio."""
    short_side_mm = min(w, h) / Mm(1)
    return min(0.5, target_mm / short_side_mm)


def round_diag_rect(slide, x, y, w, h, fill=None, border_color=None,
                     border_w=CARD_BORDER_W):
    """Rectangle with only the top-left and bottom-right corners
    rounded (PowerPoint's "Round Diagonal Corner Rectangle")."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUND_2_DIAG_RECTANGLE, x, y, w, h)
    adj = radius_adj(w, h)
    shape.adjustments[0] = adj
    shape.adjustments[1] = adj
    _style_shape(shape, fill, border_color, border_w)
    return shape


def _style_shape(shape, fill, border_color, border_w):
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if border_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = border_color
        shape.line.width = border_w
    shape.shadow.inherit = False


def add_text(slide, x, y, w, h, text, size=18, color=BLACK, bold=False,
             italic=False, align=PP_ALIGN.LEFT, font=FONT_SERIF, anchor=None,
             line_spacing=1.15):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if anchor is not None:
        tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = font
        run.font.color.rgb = color
    return box


def add_bullets(slide, x, y, w, h, items, size=17, color=BLACK,
                 font=FONT_SERIF, space_after=Pt(8)):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.15
        p.space_after = space_after
        pPr = p._p.get_or_add_pPr()
        buFont = pPr.makeelement(qn("a:buFont"), {"typeface": font})
        buClr = pPr.makeelement(qn("a:buClr"), {})
        srgb = buClr.makeelement(qn("a:srgbClr"), {"val": "F19651"})
        buClr.append(srgb)
        buChar = pPr.makeelement(qn("a:buChar"), {"char": "•"})
        pPr.append(buClr)
        pPr.append(buFont)
        pPr.append(buChar)
        pPr.set("marL", "285750")
        pPr.set("indent", "-285750")
        run = p.add_run()
        run.text = item
        run.font.size = Pt(size)
        run.font.name = font
        run.font.color.rgb = color
    return box


def caption_text(slide, x, y, w, label, rest, size=CAPTION_SIZE):
    """'Label: rest' caption, bold, with only the label in orange —
    matches the LaTeX tikzfigure caption style."""
    h = Mm(text_height_mm(label + " " + rest, size, w / Mm(1), FONT_SERIF, True, 1.15))
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run()
    r1.text = label + " "
    r1.font.bold = True
    r1.font.size = Pt(size)
    r1.font.name = FONT_SERIF
    r1.font.color.rgb = ORANGE
    r2 = p.add_run()
    r2.text = rest
    r2.font.bold = True
    r2.font.size = Pt(size)
    r2.font.name = FONT_SERIF
    r2.font.color.rgb = BLACK
    return h / Mm(1)


def add_table(slide, x, y, w, rows, header_fill=BLUE_TINT_8, font_size=TABLE_SIZE,
              row_h_mm=15):
    n_rows = len(rows)
    n_cols = len(rows[0])
    row_h = Mm(row_h_mm)
    table_shape = slide.shapes.add_table(n_rows, n_cols, x, y, w, row_h * n_rows)
    table = table_shape.table
    table.first_row = False  # don't let the default table style re-color the header row
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Mm(3)
            cell.margin_right = Mm(3)
            cell.margin_top = Mm(1)
            cell.margin_bottom = Mm(1)
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            run = para.runs[0]
            run.font.size = Pt(font_size)
            run.font.name = FONT_SERIF
            run.font.bold = (r == 0)
            run.font.color.rgb = BLACK
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_fill if r == 0 else WHITE
    return row_h_mm * n_rows


class MeasureCanvas:
    """Accumulates the height (mm) a sequence of content elements
    would need at a given width, without drawing anything — used to
    size a card to its actual content before the card is drawn."""

    def __init__(self, width_mm):
        self.width_mm = width_mm
        self.h = 0.0

    def text(self, text, size=BODY_SIZE, **kw):
        self.h += text_height_mm(text, size, self.width_mm, kw.get("font", FONT_SERIF),
                                  kw.get("bold", False), kw.get("line_spacing", 1.15))

    def bullets(self, items, size=BODY_SIZE, **kw):
        self.h += bullets_height_mm(items, size, self.width_mm, kw.get("font", FONT_SERIF),
                                     kw.get("line_spacing", 1.15), kw.get("indent_mm", 8),
                                     kw.get("space_after_mm", 3))

    def image(self, path, width_mm, caption=None, caption_size=CAPTION_SIZE):
        self.h += picture_height_mm_for_width(path, width_mm)
        if caption:
            self.h += text_height_mm(caption[0] + " " + caption[1], caption_size,
                                      self.width_mm, FONT_SERIF, True, 1.15)

    def placeholder(self, height_mm, text_label, caption=None, caption_size=CAPTION_SIZE):
        self.h += height_mm
        if caption:
            self.h += text_height_mm(caption[0] + " " + caption[1], caption_size,
                                      self.width_mm, FONT_SERIF, True, 1.15)

    def table(self, rows, row_h_mm=15, caption=None, caption_size=CAPTION_SIZE):
        if caption:
            self.h += text_height_mm(caption[0] + " " + caption[1], caption_size,
                                      self.width_mm, FONT_SERIF, True, 1.15)
        self.h += row_h_mm * len(rows)

    def gap(self, mm):
        self.h += mm


class DrawCanvas:
    """Draws the same content elements as MeasureCanvas, stacked
    top-down from (x_mm, y_mm)."""

    def __init__(self, slide, x_mm, y_mm, width_mm):
        self.slide = slide
        self.x_mm = x_mm
        self.cy = y_mm
        self.width_mm = width_mm

    def text(self, text, size=BODY_SIZE, color=BLACK, bold=False, italic=False,
             align=PP_ALIGN.LEFT, font=FONT_SERIF, line_spacing=1.15):
        h = text_height_mm(text, size, self.width_mm, font, bold, line_spacing)
        add_text(self.slide, Mm(self.x_mm), Mm(self.cy), Mm(self.width_mm), Mm(h),
                 text, size=size, color=color, bold=bold, italic=italic, align=align,
                 font=font, line_spacing=line_spacing)
        self.cy += h

    def bullets(self, items, size=BODY_SIZE, color=BLACK, font=FONT_SERIF,
                line_spacing=1.15, indent_mm=8, space_after_mm=3):
        h = bullets_height_mm(items, size, self.width_mm, font, line_spacing,
                               indent_mm, space_after_mm)
        add_bullets(self.slide, Mm(self.x_mm), Mm(self.cy), Mm(self.width_mm), Mm(h),
                    items, size=size, color=color, font=font, space_after=Pt(space_after_mm / MM_PER_PT))
        self.cy += h

    def image(self, path, width_mm, caption=None, caption_size=CAPTION_SIZE, gap_before_caption=4):
        h = picture_height_mm_for_width(path, width_mm)
        x = self.x_mm + (self.width_mm - width_mm) / 2
        self.slide.shapes.add_picture(path, Mm(x), Mm(self.cy), width=Mm(width_mm), height=Mm(h))
        self.cy += h
        if caption:
            self.cy += gap_before_caption
            cap_h = caption_text(self.slide, Mm(self.x_mm), Mm(self.cy), Mm(self.width_mm),
                                  caption[0], caption[1], size=caption_size)
            self.cy += cap_h

    def placeholder(self, height_mm, text_label, caption=None, caption_size=CAPTION_SIZE,
                     gap_before_caption=4):
        box = round_diag_rect(self.slide, Mm(self.x_mm), Mm(self.cy), Mm(self.width_mm),
                               Mm(height_mm), fill=BLUE_TINT_5, border_color=BLUE,
                               border_w=Pt(1.5))
        ln = box.line._get_or_add_ln()
        dash = ln.makeelement(qn("a:prstDash"), {"val": "dash"})
        ln.append(dash)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = text_label
        run.font.size = Pt(BODY_SIZE - 4)
        run.font.color.rgb = BLUE
        run.font.name = FONT_SERIF
        self.cy += height_mm
        if caption:
            self.cy += gap_before_caption
            cap_h = caption_text(self.slide, Mm(self.x_mm), Mm(self.cy), Mm(self.width_mm),
                                  caption[0], caption[1], size=caption_size)
            self.cy += cap_h

    def table(self, rows, row_h_mm=15, caption=None, caption_size=CAPTION_SIZE):
        if caption:
            cap_h = caption_text(self.slide, Mm(self.x_mm), Mm(self.cy), Mm(self.width_mm),
                                  caption[0], caption[1], size=caption_size)
            self.cy += cap_h + 4
        add_table(self.slide, Mm(self.x_mm), Mm(self.cy), Mm(self.width_mm), rows,
                  font_size=TABLE_SIZE, row_h_mm=row_h_mm)
        self.cy += row_h_mm * len(rows)

    def gap(self, mm):
        self.cy += mm


def block(slide, x_mm, y_mm, w_mm, title, content_fn):
    """One section, sized to its actual content: a single white card
    (top-left/bottom-right corners rounded) with an orange-tinted
    border. No separate title-bar fill — only the text distinguishes
    the blue bold-sans title from the black serif body below it,
    matching the boxless SolabimaBox style in the LaTeX template.
    `content_fn(canvas)` is called once against a MeasureCanvas to
    size the card, then again against a DrawCanvas to fill it, so it
    must be a pure function of its canvas argument (no side effects
    of its own). Returns the y (mm) just below the card."""
    inner_w = w_mm - 2 * BODY_PAD_MM
    title_h = text_height_mm(title, SECTION_TITLE_SIZE, inner_w, FONT_SANS, True, 1.15)
    gap_after_title = 8.0

    measurer = MeasureCanvas(inner_w)
    content_fn(measurer)
    content_h = measurer.h

    height_mm = BODY_PAD_MM + title_h + gap_after_title + content_h + BODY_PAD_MM

    round_diag_rect(slide, Mm(x_mm), Mm(y_mm), Mm(w_mm), Mm(height_mm),
                     fill=WHITE, border_color=BORDER_ORANGE)
    inner_x = x_mm + BODY_PAD_MM
    add_text(slide, Mm(inner_x), Mm(y_mm + BODY_PAD_MM), Mm(inner_w), Mm(title_h), title,
             size=SECTION_TITLE_SIZE, bold=True, color=BLUE, font=FONT_SANS)

    drawer = DrawCanvas(slide, inner_x, y_mm + BODY_PAD_MM + title_h + gap_after_title, inner_w)
    content_fn(drawer)

    return y_mm + height_mm


def main():
    prs = Presentation()
    prs.slide_width = PAGE_W
    prs.slide_height = PAGE_H
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # ---------------- Header ----------------
    # Boxless: title/subtitle/authors/affiliation in blue/dark-gray
    # directly on the white page, centered in a 70%-wide column,
    # flanked by the SOLABIMA logo (left) and a placeholder
    # institution logo (right) — matches \TP@maketitle in the LaTeX
    # template. Font sizes are LaTeX's \Huge/\Large/\large; the block
    # of header text is bottom-anchored to a measured target (256mm,
    # from the top of the first card in the rendered LaTeX PDF) via
    # the gap before body_top, so header-font choices here don't
    # drift the column layout below.
    avail_w_mm = (PAGE_W - 2 * MARGIN) / Mm(1)
    header_w_mm = avail_w_mm * 0.70
    header_x_mm = (avail_w_mm - header_w_mm) / 2 + MARGIN / Mm(1)

    header = DrawCanvas(slide, header_x_mm, 20, header_w_mm)
    title_y0 = header.cy
    header.text("TÍTULO DEL TRABAJO EN MAYÚSCULAS", size=HEADER_TITLE_SIZE, bold=True,
                color=BLUE, align=PP_ALIGN.CENTER, font=FONT_SANS)
    header.gap(8)
    header.text("Subtítulo opcional", size=HEADER_SUBTITLE_SIZE, bold=True,
                color=ORANGE_DARK, align=PP_ALIGN.CENTER, font=FONT_SANS)
    header.gap(12)
    header.text("Nombre Apellido¹, Nombre Apellido², Nombre Apellido¹,*",
                size=HEADER_AUTHORS_SIZE, color=DARK_GRAY, align=PP_ALIGN.CENTER,
                font=FONT_SERIF)
    header.gap(10)
    header.text(
        "¹Facultad de Ciencias Imaginarias, Universidad de Algún Lugar, País\n"
        "²Instituto de Ejemplos Genéricos, Ciudad Ficticia\n"
        "*Autor de correspondencia: correo@ejemplo.py",
        size=HEADER_INSTITUTE_SIZE, color=DARK_GRAY, align=PP_ALIGN.CENTER,
        font=FONT_SERIF)
    header_bottom_mm = header.cy
    header_cy_mm = (title_y0 + header_bottom_mm) / 2

    # Flanking logos, vertically centered on the header block, each
    # centered within its own side margin (the strip between the
    # page margin and the 70%-wide title column). Heights match the
    # LaTeX template's explicit cm sizes exactly.
    side_zone_w_mm = header_x_mm - MARGIN / Mm(1)

    badge_path = os.path.join(LOGOS_DIR, "logo_solabima_badge.png")
    badge_h_mm = 130  # 13cm, matches LaTeX
    badge_w_mm = badge_h_mm * (Image.open(badge_path).width / Image.open(badge_path).height)
    badge_x_mm = MARGIN / Mm(1) + (side_zone_w_mm - badge_w_mm) / 2
    slide.shapes.add_picture(badge_path, Mm(badge_x_mm), Mm(header_cy_mm - badge_h_mm / 2),
                              height=Mm(badge_h_mm))

    right_logo_path = os.path.join(LOGOS_DIR, "logo_header_right.png")
    right_h_mm = 85  # 8.5cm, matches LaTeX
    right_w_mm = right_h_mm * (Image.open(right_logo_path).width / Image.open(right_logo_path).height)
    right_zone_x_mm = header_x_mm + header_w_mm
    right_x_mm = right_zone_x_mm + (side_zone_w_mm - right_w_mm) / 2
    slide.shapes.add_picture(right_logo_path, Mm(right_x_mm), Mm(header_cy_mm - right_h_mm / 2),
                              height=Mm(right_h_mm))

    # Target measured directly from the rendered LaTeX PDF (top of
    # the first column-1 card sits at ~256mm from the page top).
    body_top_mm = 256.0

    # Footer geometry (also measured from the LaTeX PDF: ~90mm tall),
    # computed up front so column heights can be checked against it.
    footer_h_mm = 90.0
    footer_y_mm = PAGE_H / Mm(1) - MARGIN / Mm(1) - footer_h_mm
    col_bottom_limit_mm = footer_y_mm - 15

    # ---------------- Column 1 ----------------
    x_mm = col_x(0) / Mm(1)
    y_mm = body_top_mm

    def motivacion_content(c):
        c.text(
            "Presente aquí el contexto y la motivación del trabajo: "
            "el problema abordado, su relevancia dentro de la biología "
            "matemática y los antecedentes más importantes.")
        c.gap(6)
        c.text(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed "
            "do eiusmod tempor incididunt ut labore et dolore magna "
            "aliqua. Ut enim ad minim veniam, quis nostrud exercitation "
            "ullamco laboris nisi ut aliquip ex ea commodo consequat.")
        c.gap(8)
        c.bullets([
            "Antecedente o dato relevante 1.",
            "Antecedente o dato relevante 2.",
            "Vacío de conocimiento que motiva este trabajo.",
            "Duis aute irure dolor in reprehenderit in voluptate velit esse "
            "cillum dolore eu fugiat nulla pariatur.",
        ])

    y_mm = block(slide, x_mm, y_mm, COL_W_MM, "Motivación / Introducción", motivacion_content)
    y_mm += GAP / Mm(1)

    def objetivos_content(c):
        c.bullets([
            "Objetivo general del trabajo.",
            "Objetivo específico 1.",
            "Objetivo específico 2.",
        ])

    y_mm = block(slide, x_mm, y_mm, COL_W_MM, "Objetivos", objetivos_content)
    y_mm += GAP / Mm(1)

    def metodologia_content(c):
        c.text(
            "Describa el enfoque, modelo matemático, datos y/o "
            "algoritmos utilizados.")
        c.gap(5)
        c.text(
            "Sed ut perspiciatis unde omnis iste natus error sit "
            "voluptatem accusantium doloremque laudantium. Totam rem "
            "aperiam, eaque ipsa quae ab illo inventore veritatis et "
            "quasi architecto beatae vitae dicta sunt explicabo.")
        c.gap(5)
        c.bullets([
            "Diseño del estudio / modelo propuesto.",
            "Fuentes de datos o parámetros.",
            "Herramientas y métodos de análisis.",
            "Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut "
            "odit aut fugit.",
        ])
        c.gap(8)
        c.image(os.path.join(ASSETS_DIR, "equation_sir.png"), 230,
                caption=("Ejemplo:", "modelo SIR de dinámica epidémica"))

    y_mm = block(slide, x_mm, y_mm, COL_W_MM, "Metodología", metodologia_content)
    col1_bottom_mm = y_mm

    # ---------------- Column 2 ----------------
    x_mm = col_x(1) / Mm(1)
    y_mm = body_top_mm

    def resultados_content(c):
        c.text(
            "Presente los resultados principales, apoyados en figuras "
            "y/o tablas.")
        c.gap(6)
        c.text(
            "Totam rem aperiam, eaque ipsa quae ab illo inventore "
            "veritatis et quasi architecto beatae vitae dicta sunt "
            "explicabo. Nemo enim ipsam voluptatem quia voluptas sit "
            "aspernatur aut odit aut fugit.")
        c.gap(8)
        c.placeholder(100, "Espacio para figura principal\n(gráfico, mapa, simulación, etc.)",
                      caption=("Figura:", "Descripción breve del resultado."))
        c.gap(10)
        c.table([
            ["Parámetro", "Valor", "IC 95%", "p"],
            ["Parámetro 1", "3.14", "(3.10, 3.18)", "0.001"],
            ["Parámetro 2", "2.72", "(2.68, 2.76)", "0.007"],
            ["Parámetro 3", "1.62", "(1.58, 1.66)", "0.042"],
        ], caption=("Tabla:", "Descripción breve de la tabla."))
        c.gap(8)
        c.bullets([
            "Sed ut perspiciatis unde omnis iste natus error sit "
            "voluptatem accusantium doloremque laudantium.",
            "Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut "
            "odit aut fugit, sed quia consequuntur magni dolores.",
            "Neque porro quisquam est, qui dolorem ipsum quia dolor sit "
            "amet, consectetur, adipisci velit.",
        ])

    y_mm = block(slide, x_mm, y_mm, COL_W_MM, "Resultados", resultados_content)
    y_mm += GAP / Mm(1)

    def conclusiones_content(c):
        c.bullets([
            "Conclusión principal 1.",
            "Conclusión principal 2.",
            "Proyecciones o trabajo futuro.",
        ])

    y_mm = block(slide, x_mm, y_mm, COL_W_MM, "Conclusiones", conclusiones_content)
    y_mm += GAP / Mm(1)

    def referencias_content(c):
        c.text(
            "[1] Apellido, N. (Año). Título del artículo. Revista, "
            "vol(núm), páginas.", size=REF_SIZE)
        c.gap(4)
        c.text(
            "[2] Apellido, N. & Apellido, N. (Año). Título del "
            "artículo. Revista, vol(núm), páginas.", size=REF_SIZE)
        c.gap(4)
        c.text(
            "[3] Apellido, N., Apellido, N. & Apellido, N. (Año). Título "
            "del artículo. Revista, vol(núm), páginas.", size=REF_SIZE)

    y_mm = block(slide, x_mm, y_mm, COL_W_MM, "Referencias", referencias_content)
    y_mm += GAP / Mm(1)

    def agradecimientos_content(c):
        c.text("Financiamiento, institución o colegas a reconocer.", size=ACK_SIZE)
        c.gap(4)
        c.text(
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco "
            "laboris nisi ut aliquip ex ea commodo consequat.", size=ACK_SIZE)

    y_mm = block(slide, x_mm, y_mm, COL_W_MM, "Agradecimientos", agradecimientos_content)
    col2_bottom_mm = y_mm

    assert col1_bottom_mm <= col_bottom_limit_mm, (
        f"Column 1 ({col1_bottom_mm:.0f}mm) overruns the footer limit "
        f"({col_bottom_limit_mm:.0f}mm) - shrink content or the footer.")
    assert col2_bottom_mm <= col_bottom_limit_mm, (
        f"Column 2 ({col2_bottom_mm:.0f}mm) overruns the footer limit "
        f"({col_bottom_limit_mm:.0f}mm) - shrink content or the footer.")

    # ---------------- Footer: optional funding/sponsor logos ----------------
    # Matches the LaTeX template's optional \ifSolabimaFundingLogos
    # band: an explanatory note plus generic placeholder logos, not
    # the fixed SOLABIMA/UNA/FPUNA institutional set (those only
    # appear in the header now).
    round_diag_rect(slide, MARGIN, Mm(footer_y_mm), PAGE_W - 2 * MARGIN, Mm(footer_h_mm),
                     fill=WHITE, border_color=BORDER_ORANGE)

    note_y_mm = footer_y_mm + BODY_PAD_MM
    add_text(slide, MARGIN + BODY_PAD, Mm(note_y_mm),
              PAGE_W - 2 * MARGIN - 2 * BODY_PAD, Mm(20),
              "Esta franja es opcional: puede usarla para agregar más logos, "
              "información adicional, un código QR, etc., o eliminarla por "
              "completo si no la necesita.",
              size=CAPTION_SIZE - 3, italic=True, color=DARK_GRAY, align=PP_ALIGN.CENTER,
              font=FONT_SERIF)

    logo_paths = [
        os.path.join(LOGOS_DIR, "logo_custom1.png"),
        os.path.join(LOGOS_DIR, "logo_custom2.png"),
    ]
    logo_h_mm = 55  # 5.5cm, matches LaTeX
    logo_gap_mm = 12
    widths_mm = [logo_h_mm * (Image.open(p).width / Image.open(p).height) for p in logo_paths]
    total_w_mm = sum(widths_mm) + logo_gap_mm * (len(logo_paths) - 1)
    start_x_mm = MARGIN / Mm(1) + (PAGE_W / Mm(1) - 2 * MARGIN / Mm(1) - total_w_mm) / 2
    logo_y_mm = footer_y_mm + footer_h_mm - BODY_PAD_MM - logo_h_mm
    lx_mm = start_x_mm
    for path, w_mm in zip(logo_paths, widths_mm):
        slide.shapes.add_picture(path, Mm(lx_mm), Mm(logo_y_mm), height=Mm(logo_h_mm))
        lx_mm += w_mm + logo_gap_mm

    prs.save(OUT_PATH)
    print(f"wrote {OUT_PATH}")
    print(f"column 1 bottom: {col1_bottom_mm:.0f}mm, column 2 bottom: {col2_bottom_mm:.0f}mm, "
          f"limit: {col_bottom_limit_mm:.0f}mm")


if __name__ == "__main__":
    main()
