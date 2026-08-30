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

Corner rounding relies on PowerPoint's built-in "Round Diagonal
Corner Rectangle" autoshape, which by default rounds the top-left and
bottom-right corners (matching the Office shape-gallery icon). This
was written without being able to render/verify the .pptx locally (no
PowerPoint/Keynote/LibreOffice on the build machine) — if a card's
rounded corners land on the wrong pair when you open the file, that
default-corner assumption is the first thing to check (see
round_diag_rect below).

Run:  python3 build_pptx.py
"""
import os
from PIL import Image
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

PAGE_W = Mm(841)
PAGE_H = Mm(1189)
MARGIN = Mm(15)
COL_GAP = Mm(25)
N_COLS = 2
COL_W = int((PAGE_W - 2 * MARGIN - (N_COLS - 1) * COL_GAP) / N_COLS)
GAP = Mm(20)  # vertical gap between stacked blocks in a column

TITLE_H = Mm(20)   # space reserved for a block's inset title line
BODY_PAD = Mm(14)
RADIUS_MM = 7.0  # absolute target corner radius, matches \SolabimaR in LaTeX


def col_x(i):
    return MARGIN + i * (COL_W + COL_GAP)


def radius_adj(w, h, target_mm=RADIUS_MM):
    """adj fraction (of the shape's shorter side) giving an absolute
    radius of target_mm, regardless of this shape's own aspect ratio."""
    short_side_mm = min(w, h) / Mm(1)
    return min(0.5, target_mm / short_side_mm)


def round_diag_rect(slide, x, y, w, h, fill=None, border_color=None,
                     border_w=Pt(2)):
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


def block(slide, x, y, w, height, title, body_items=None, body_text=None,
          title_size=28, body_size=17, body_builder=None):
    """One section: a single white card (top-left/bottom-right corners
    rounded) with an orange-tinted border. No separate title-bar fill
    — only the text distinguishes the blue bold-sans title from the
    black serif body below it, matching the boxless SolabimaBox style
    in the LaTeX template."""
    round_diag_rect(slide, x, y, w, height, fill=WHITE, border_color=BORDER_ORANGE)
    inner_x = x + BODY_PAD
    inner_w = w - 2 * BODY_PAD
    if title:
        add_text(slide, inner_x, y + BODY_PAD, inner_w, TITLE_H, title,
                  size=title_size, bold=True, color=BLUE, font=FONT_SANS)
        inner_y = y + BODY_PAD + TITLE_H + Mm(6)
    else:
        inner_y = y + BODY_PAD
    inner_h = (y + height) - inner_y - BODY_PAD
    if body_items:
        add_bullets(slide, inner_x, inner_y, inner_w, inner_h, body_items, size=body_size)
    elif body_text:
        add_text(slide, inner_x, inner_y, inner_w, inner_h, body_text, size=body_size)
    elif body_builder:
        body_builder(inner_x, inner_y, inner_w, inner_h)
    return y + height


def placeholder_box(slide, x, y, w, height, caption):
    """Dashed, pale-blue-tinted placeholder (matches the LaTeX
    dashed+tinted figure box)."""
    box = round_diag_rect(slide, x, y, w, height, fill=BLUE_TINT_5,
                           border_color=BLUE, border_w=Pt(1.25))
    ln = box.line._get_or_add_ln()
    dash = ln.makeelement(qn("a:prstDash"), {"val": "dash"})
    ln.append(dash)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = caption
    run.font.size = Pt(16)
    run.font.color.rgb = BLUE
    run.font.name = FONT_SERIF


def caption_text(slide, x, y, w, label, rest, size=15):
    """'Label: rest' caption, bold, with only the label in orange —
    matches the LaTeX tikzfigure caption style."""
    box = slide.shapes.add_textbox(x, y, w, Mm(14))
    tf = box.text_frame
    tf.word_wrap = True
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
    return box


def add_table(slide, x, y, w, rows, header_fill=BLUE_TINT_8, font_size=14):
    n_rows = len(rows)
    n_cols = len(rows[0])
    row_h = Mm(14)
    table_shape = slide.shapes.add_table(n_rows, n_cols, x, y, w, row_h * n_rows)
    table = table_shape.table
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
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_fill if r == 0 else WHITE
    return table_shape, row_h * n_rows


def picture_height_for_width(path, width):
    with Image.open(path) as im:
        px_w, px_h = im.size
    return int(width * (px_h / px_w))


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
    # template.
    avail_w = PAGE_W - 2 * MARGIN
    header_w = int(avail_w * 0.70)
    header_x = int(MARGIN + (avail_w - header_w) / 2)

    title_y = Mm(20)
    title_h = Mm(75)
    add_text(slide, header_x, title_y, header_w, title_h,
             "TÍTULO DEL TRABAJO EN MAYÚSCULAS", size=58, bold=True,
             color=BLUE, align=PP_ALIGN.CENTER, font=FONT_SANS,
             anchor=MSO_ANCHOR.TOP)

    subtitle_y = title_y + title_h + Mm(8)
    subtitle_h = Mm(26)
    add_text(slide, header_x, subtitle_y, header_w, subtitle_h,
             "Subtítulo opcional", size=30, bold=True, color=ORANGE_DARK,
             align=PP_ALIGN.CENTER, font=FONT_SANS)

    authors_y = subtitle_y + subtitle_h + Mm(14)
    authors_h = Mm(22)
    add_text(slide, header_x, authors_y, header_w, authors_h,
             "Nombre Apellido¹, Nombre Apellido², Nombre Apellido¹,*",
             size=24, color=DARK_GRAY, align=PP_ALIGN.CENTER, font=FONT_SERIF)

    institute_y = authors_y + authors_h + Mm(10)
    institute_h = Mm(45)
    add_text(slide, header_x, institute_y, header_w, institute_h,
             "¹Facultad de Ciencias Imaginarias, Universidad de Algún Lugar, País\n"
             "²Instituto de Ejemplos Genéricos, Ciudad Ficticia\n"
             "*Autor de correspondencia: correo@ejemplo.py",
             size=18, color=DARK_GRAY, align=PP_ALIGN.CENTER, font=FONT_SERIF)

    header_bottom = institute_y + institute_h
    header_cy = (title_y + header_bottom) // 2

    # Flanking logos, vertically centered on the header block, each
    # centered within its own side margin (the strip between the
    # page margin and the 70%-wide title column).
    side_zone_w = header_x - MARGIN

    badge_path = os.path.join(LOGOS_DIR, "logo_solabima_badge.png")
    badge_h = Mm(115)
    badge_w = int(badge_h * (Image.open(badge_path).size[0] / Image.open(badge_path).size[1]))
    badge_x = int(MARGIN + (side_zone_w - badge_w) / 2)
    slide.shapes.add_picture(badge_path, badge_x, int(header_cy - badge_h / 2),
                              height=badge_h)

    right_logo_path = os.path.join(LOGOS_DIR, "logo_header_right.png")
    right_h = Mm(75)
    right_w = int(right_h * (Image.open(right_logo_path).size[0] / Image.open(right_logo_path).size[1]))
    right_zone_x = header_x + header_w
    right_x = int(right_zone_x + (side_zone_w - right_w) / 2)
    slide.shapes.add_picture(right_logo_path, right_x, int(header_cy - right_h / 2),
                              height=right_h)

    body_top = header_bottom + Mm(25)

    # Footer geometry, computed up front so column heights can be
    # sized to end just above it (mirrors the fixed \TP@blocktop
    # anchor used for the optional funding-logos band in the LaTeX
    # version).
    footer_h = Mm(90)
    footer_y = PAGE_H - MARGIN - footer_h
    col_bottom_limit = footer_y - Mm(15)

    # ---------------- Column 1 ----------------
    x = col_x(0)
    y = body_top

    def motivacion_body(bx, by, bw, bh):
        add_text(slide, bx, by, bw, Mm(35),
                 "Presente aquí el contexto y la motivación del trabajo: "
                 "el problema abordado, su relevancia dentro de la biología "
                 "matemática y los antecedentes más importantes.", size=16)
        add_text(slide, bx, by + Mm(38), bw, Mm(45),
                 "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed "
                 "do eiusmod tempor incididunt ut labore et dolore magna "
                 "aliqua. Ut enim ad minim veniam, quis nostrud exercitation "
                 "ullamco laboris nisi ut aliquip ex ea commodo consequat.",
                 size=16)
        add_bullets(slide, bx, by + Mm(88), bw, bh - Mm(88), [
            "Antecedente o dato relevante 1.",
            "Antecedente o dato relevante 2.",
            "Vacío de conocimiento que motiva este trabajo.",
            "Duis aute irure dolor in reprehenderit in voluptate velit esse "
            "cillum dolore eu fugiat nulla pariatur.",
        ], size=16)

    y = block(slide, x, y, COL_W, Mm(215), "Motivación / Introducción",
              body_builder=motivacion_body)
    y += GAP

    y = block(slide, x, y, COL_W, Mm(115), "Objetivos", body_items=[
        "Objetivo general del trabajo.",
        "Objetivo específico 1.",
        "Objetivo específico 2.",
    ], body_size=16)
    y += GAP

    def metodologia_body(bx, by, bw, bh):
        add_text(slide, bx, by, bw, Mm(28),
                 "Describa el enfoque, modelo matemático, datos y/o "
                 "algoritmos utilizados.", size=16)
        add_text(slide, bx, by + Mm(30), bw, Mm(45),
                 "Sed ut perspiciatis unde omnis iste natus error sit "
                 "voluptatem accusantium doloremque laudantium. Totam rem "
                 "aperiam, eaque ipsa quae ab illo inventore veritatis et "
                 "quasi architecto beatae vitae dicta sunt explicabo.",
                 size=16)
        add_bullets(slide, bx, by + Mm(78), bw, Mm(60), [
            "Diseño del estudio / modelo propuesto.",
            "Fuentes de datos o parámetros.",
            "Herramientas y métodos de análisis.",
            "Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut "
            "odit aut fugit.",
        ], size=16)
        eq_top = by + Mm(140)
        eq_path = os.path.join(ASSETS_DIR, "equation_sir.png")
        eq_w = Mm(230)
        eq_h = picture_height_for_width(eq_path, eq_w)
        eq_x = int(bx + (bw - eq_w) / 2)
        slide.shapes.add_picture(eq_path, eq_x, eq_top, width=eq_w, height=eq_h)
        caption_text(slide, bx, eq_top + eq_h + Mm(4), bw,
                     "Ejemplo:", "modelo SIR de dinámica epidémica",
                     size=15)

    y = block(slide, x, y, COL_W, Mm(400), "Metodología",
              body_builder=metodologia_body)
    col1_bottom = y

    # ---------------- Column 2 ----------------
    x = col_x(1)
    y = body_top

    def resultados_body(bx, by, bw, bh):
        add_text(slide, bx, by, bw, Mm(30),
                 "Presente los resultados principales, apoyados en figuras "
                 "y/o tablas.", size=16)
        add_text(slide, bx, by + Mm(32), bw, Mm(45),
                 "Totam rem aperiam, eaque ipsa quae ab illo inventore "
                 "veritatis et quasi architecto beatae vitae dicta sunt "
                 "explicabo. Nemo enim ipsam voluptatem quia voluptas sit "
                 "aspernatur aut odit aut fugit.", size=16)
        fig_top = by + Mm(82)
        fig_h = Mm(150)
        placeholder_box(slide, bx, fig_top, bw, fig_h,
                         "Espacio para figura principal\n"
                         "(gráfico, mapa, simulación, etc.)")
        caption_text(slide, bx, fig_top + fig_h + Mm(4), bw,
                     "Figura:", "Descripción breve del resultado.")

        table_top = fig_top + fig_h + Mm(24)
        caption_text(slide, bx, table_top, bw,
                     "Tabla:", "Descripción breve de la tabla.")
        _, table_h = add_table(slide, bx, table_top + Mm(16), bw, [
            ["Parámetro", "Valor", "IC 95%", "p"],
            ["Parámetro 1", "3.14", "(3.10, 3.18)", "0.001"],
            ["Parámetro 2", "2.72", "(2.68, 2.76)", "0.007"],
            ["Parámetro 3", "1.62", "(1.58, 1.66)", "0.042"],
        ])

        bullets_top = table_top + Mm(16) + table_h + Mm(10)
        add_bullets(slide, bx, bullets_top, bw, bh - (bullets_top - by), [
            "Sed ut perspiciatis unde omnis iste natus error sit "
            "voluptatem accusantium doloremque laudantium.",
            "Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut "
            "odit aut fugit, sed quia consequuntur magni dolores.",
            "Neque porro quisquam est, qui dolorem ipsum quia dolor sit "
            "amet, consectetur, adipisci velit.",
        ], size=16)

    y = block(slide, x, y, COL_W, Mm(420), "Resultados", body_builder=resultados_body)
    y += GAP

    y = block(slide, x, y, COL_W, Mm(90), "Conclusiones", body_items=[
        "Conclusión principal 1.",
        "Conclusión principal 2.",
        "Proyecciones o trabajo futuro.",
    ], body_size=16)
    y += GAP

    y = block(slide, x, y, COL_W, Mm(120), "Referencias", body_text=(
        "[1] Apellido, N. (Año). Título del artículo. Revista, "
        "vol(núm), páginas.\n\n"
        "[2] Apellido, N. & Apellido, N. (Año). Título del "
        "artículo. Revista, vol(núm), páginas.\n\n"
        "[3] Apellido, N., Apellido, N. & Apellido, N. (Año). Título "
        "del artículo. Revista, vol(núm), páginas."
    ), body_size=13)
    y += GAP

    y = block(slide, x, y, COL_W, Mm(75), "Agradecimientos", body_text=(
        "Financiamiento, institución o colegas a reconocer.\n\n"
        "Ut enim ad minim veniam, quis nostrud exercitation ullamco "
        "laboris nisi ut aliquip ex ea commodo consequat."
    ), body_size=15)
    col2_bottom = y

    assert col1_bottom <= col_bottom_limit, (
        f"Column 1 ({col1_bottom/Mm(1):.0f}mm) overruns the footer "
        f"limit ({col_bottom_limit/Mm(1):.0f}mm) - shrink content or "
        f"the footer.")
    assert col2_bottom <= col_bottom_limit, (
        f"Column 2 ({col2_bottom/Mm(1):.0f}mm) overruns the footer "
        f"limit ({col_bottom_limit/Mm(1):.0f}mm) - shrink content or "
        f"the footer.")

    # ---------------- Footer: optional funding/sponsor logos ----------------
    # Matches the LaTeX template's optional \ifSolabimaFundingLogos
    # band: an explanatory note plus generic placeholder logos, not
    # the fixed SOLABIMA/UNA/FPUNA institutional set (those only
    # appear in the header now).
    round_diag_rect(slide, MARGIN, footer_y, PAGE_W - 2 * MARGIN, footer_h,
                     fill=WHITE, border_color=BORDER_ORANGE)

    note_y = footer_y + BODY_PAD
    add_text(slide, MARGIN + BODY_PAD, note_y, PAGE_W - 2 * MARGIN - 2 * BODY_PAD,
              Mm(16),
              "Esta franja es opcional: puede usarla para agregar más logos, "
              "información adicional, un código QR, etc., o eliminarla por "
              "completo si no la necesita.",
              size=15, italic=True, color=DARK_GRAY, align=PP_ALIGN.CENTER,
              font=FONT_SERIF)

    logo_paths = [
        os.path.join(LOGOS_DIR, "logo_custom1.png"),
        os.path.join(LOGOS_DIR, "logo_custom2.png"),
    ]
    logo_h = Mm(55)
    logo_gap = Mm(12)
    widths = []
    for path in logo_paths:
        with Image.open(path) as im:
            px_w, px_h = im.size
        widths.append(int(logo_h * (px_w / px_h)))
    total_w = sum(widths) + logo_gap * (len(logo_paths) - 1)
    start_x = int(MARGIN + (PAGE_W - 2 * MARGIN - total_w) / 2)
    logo_y = footer_y + footer_h - BODY_PAD - logo_h
    lx = start_x
    for path, w in zip(logo_paths, widths):
        slide.shapes.add_picture(path, lx, logo_y, height=logo_h)
        lx += w + logo_gap

    prs.save(OUT_PATH)
    print(f"wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
