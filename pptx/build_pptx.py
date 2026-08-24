"""Build the SOLABIMA 2026 PowerPoint poster template.

Generates solabima2026_poster_template.pptx: an A0 portrait poster
(841 x 1189 mm) matching the layout and color palette of the LaTeX
template in ../latex — two columns, boxed sections, logos in a
footer band. Edit this script and re-run it to regenerate the file,
or just edit the .pptx directly in PowerPoint / Keynote.

Run:  python3 build_pptx.py
"""
import os
import qrcode
from pptx import Presentation
from pptx.util import Mm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
LOGOS_DIR = os.path.join(HERE, "..", "latex", "logos")
OUT_PATH = os.path.join(HERE, "solabima2026_poster_template.pptx")
QR_PATH = os.path.join(HERE, "_qr_solabima2026.png")

# ---------------------------------------------------------------
# Palette (matches latex/solabima2026_poster.tex)
# ---------------------------------------------------------------
BLUE = RGBColor(0x00, 0x3B, 0x49)
TEAL = RGBColor(0x1B, 0x8A, 0x8F)
GOLD = RGBColor(0xE0, 0xA4, 0x58)
GREY = RGBColor(0xF4, 0xF4, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT = RGBColor(0x20, 0x20, 0x20)

PAGE_W = Mm(841)
PAGE_H = Mm(1189)
MARGIN = Mm(25)
COL_GAP = Mm(25)
N_COLS = 2
COL_W = (PAGE_W - 2 * MARGIN - (N_COLS - 1) * COL_GAP) / N_COLS

FONT = "Georgia"
FONT_SANS = "Arial"

TITLE_BAR_H = Mm(20)
BODY_PAD = Mm(10)


def col_x(i):
    return MARGIN + i * (COL_W + COL_GAP)


def make_qr(path: str, url: str) -> None:
    img = qrcode.make(url)
    img.save(path)


def add_rect(slide, x, y, w, h, fill=None, border_color=None, border_w=Pt(1.5)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
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
    return shape


def add_text(slide, x, y, w, h, text, size=18, color=DARK_TEXT, bold=False,
             italic=False, align=PP_ALIGN.LEFT, font=FONT_SANS, anchor=None,
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


def add_bullets(slide, x, y, w, h, items, size=17, color=DARK_TEXT,
                 font=FONT_SANS):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.15
        p.space_after = Pt(8)
        pPr = p._p.get_or_add_pPr()
        buFont = pPr.makeelement(qn("a:buFont"), {"typeface": font})
        buChar = pPr.makeelement(qn("a:buChar"), {"char": "▪"})
        pPr.append(buFont)
        pPr.append(buChar)
        pPr.set("marL", "228600")
        pPr.set("indent", "-228600")
        run = p.add_run()
        run.text = item
        run.font.size = Pt(size)
        run.font.name = font
        run.font.color.rgb = color
    return box


def block(slide, x, y, w, height, title, body_items=None, body_text=None,
          title_size=26, body_size=17, body_builder=None):
    """Draw one boxed section: blue title bar + white bordered body."""
    add_rect(slide, x, y, w, height, fill=WHITE, border_color=BLUE, border_w=Pt(2))
    if title:
        add_rect(slide, x, y, w, TITLE_BAR_H, fill=BLUE, border_color=BLUE, border_w=Pt(2))
        add_text(slide, x + BODY_PAD, y, w - 2 * BODY_PAD, TITLE_BAR_H, title,
                  size=title_size, bold=True, color=WHITE, font=FONT,
                  anchor=MSO_ANCHOR.MIDDLE)
        body_y = y + TITLE_BAR_H + BODY_PAD
        body_h = height - TITLE_BAR_H - 2 * BODY_PAD
    else:
        body_y = y + BODY_PAD
        body_h = height - 2 * BODY_PAD
    body_x = x + BODY_PAD
    body_w = w - 2 * BODY_PAD
    if body_items:
        add_bullets(slide, body_x, body_y, body_w, body_h, body_items, size=body_size)
    elif body_text:
        add_text(slide, body_x, body_y, body_w, body_h, body_text, size=body_size)
    elif body_builder:
        body_builder(body_x, body_y, body_w, body_h)
    return y + height


def placeholder_box(slide, x, y, w, height, caption):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, height)
    box.fill.background()
    box.line.color.rgb = RGBColor(0x99, 0x99, 0x99)
    box.line.width = Pt(1)
    line_elem = box.line._get_or_add_ln()
    dash = line_elem.makeelement(qn("a:prstDash"), {"val": "dash"})
    line_elem.append(dash)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = caption
    run.font.size = Pt(15)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x77, 0x77, 0x77)
    run.font.name = FONT_SANS


def main():
    make_qr(QR_PATH, "https://www.solabima2026.pol.una.py")

    prs = Presentation()
    prs.slide_width = PAGE_W
    prs.slide_height = PAGE_H
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # ---------------- Header (title only, logos moved to footer) ----------------
    header_h = Mm(150)
    add_rect(slide, 0, 0, PAGE_W, header_h, fill=BLUE)

    add_text(slide, MARGIN, Mm(20), PAGE_W - 2 * MARGIN, Mm(35),
             "TITULO DEL TRABAJO EN MAYUSCULAS", size=54, bold=True,
             color=WHITE, align=PP_ALIGN.CENTER, font=FONT)
    add_text(slide, MARGIN, Mm(55), PAGE_W - 2 * MARGIN, Mm(16),
             "Subtitulo opcional", size=26, italic=True, color=WHITE,
             align=PP_ALIGN.CENTER, font=FONT)
    add_text(slide, MARGIN, Mm(77), PAGE_W - 2 * MARGIN, Mm(16),
             "Nombre Apellido¹, Nombre Apellido², Nombre Apellido¹,*",
             size=24, color=WHITE, align=PP_ALIGN.CENTER, font=FONT)
    add_text(slide, MARGIN, Mm(100), PAGE_W - 2 * MARGIN, Mm(40),
             "¹Facultad Politecnica, Universidad Nacional de Asuncion (FPUNA), Paraguay"
             "    ²Institucion / Departamento\n"
             "*Autor de correspondencia: correo@ejemplo.py",
             size=18, color=WHITE, align=PP_ALIGN.CENTER, font=FONT_SANS)

    body_top = header_h + Mm(25)
    gap = Mm(18)

    # ---------------- Column 1 ----------------
    x = col_x(0)
    y = body_top
    y = block(slide, x, y, COL_W, Mm(190), "Motivacion / Introduccion",
              body_builder=lambda bx, by, bw, bh: (
                  add_text(slide, bx, by, bw, Mm(60),
                           "Presente aqui el contexto y la motivacion del trabajo: "
                           "el problema abordado, su relevancia dentro de la "
                           "biologia matematica y los antecedentes mas importantes.",
                           size=17),
                  add_bullets(slide, bx, by + Mm(65), bw, bh - Mm(65),
                              ["Antecedente o dato relevante 1.",
                               "Antecedente o dato relevante 2.",
                               "Vacio de conocimiento que motiva este trabajo."]),
              ))
    y += gap
    y = block(slide, x, y, COL_W, Mm(110), "Objetivos",
              body_items=["Objetivo general del trabajo.",
                          "Objetivo especifico 1.",
                          "Objetivo especifico 2."])
    y += gap
    y = block(slide, x, y, COL_W, Mm(260), "Metodologia",
              body_builder=lambda bx, by, bw, bh: (
                  add_text(slide, bx, by, bw, Mm(35),
                           "Describa el enfoque, modelo matematico, datos y/o "
                           "algoritmos utilizados.", size=17),
                  add_bullets(slide, bx, by + Mm(38), bw, Mm(70),
                              ["Diseno del estudio / modelo propuesto.",
                               "Fuentes de datos o parametros.",
                               "Herramientas y metodos de analisis."]),
                  placeholder_box(slide, bx, by + Mm(112), bw, bh - Mm(112),
                                   "Espacio para ecuacion, diagrama de flujo\n"
                                   "o esquema del modelo"),
              ))
    col1_bottom = y

    # ---------------- Column 2 ----------------
    x = col_x(1)
    y = body_top
    y = block(slide, x, y, COL_W, Mm(390), "Resultados",
              body_builder=lambda bx, by, bw, bh: (
                  add_text(slide, bx, by, bw, Mm(30),
                           "Presente los resultados principales, apoyados en "
                           "figuras y/o tablas.", size=17),
                  placeholder_box(slide, bx, by + Mm(35), bw, Mm(220),
                                   "Figura 1. Espacio para figura principal\n"
                                   "(grafico, mapa, simulacion, etc.)"),
                  placeholder_box(slide, bx, by + Mm(265), bw, bh - Mm(265),
                                   "Tabla 1. Espacio para tabla de resultados\n"
                                   "(parametro, valor, IC 95%, p)"),
              ))
    y += gap

    ref_h = Mm(200)

    def ref_body(bx, by, bw, bh):
        qr_size = Mm(60)
        qr_x = bx + bw - qr_size
        text_w = bw - qr_size - Mm(15)
        add_text(slide, bx, by, text_w, bh,
                 "[1] Apellido, N. (Ano). Titulo del articulo. Revista, "
                 "vol(num), paginas.\n\n"
                 "[2] Apellido, N. & Apellido, N. (Ano). Titulo del "
                 "articulo. Revista, vol(num), paginas.", size=15)
        slide.shapes.add_picture(QR_PATH, qr_x, by, width=qr_size, height=qr_size)
        add_text(slide, qr_x - Mm(10), by + qr_size + Mm(2), qr_size + Mm(20),
                 Mm(20), "Escanee para mas\ninformacion", size=12,
                 align=PP_ALIGN.CENTER, color=DARK_TEXT)

    y = block(slide, x, y, COL_W, ref_h, "Referencias", body_builder=ref_body)
    y += gap
    y = block(slide, x, y, COL_W, Mm(100), "Agradecimientos",
              body_text="Financiamiento, institucion o colegas a reconocer.")
    col2_bottom = y

    # ---------------- Footer: institutional logos ----------------
    footer_y = max(col1_bottom, col2_bottom) + Mm(25)
    footer_h = Mm(140)
    footer_w = PAGE_W - 2 * MARGIN
    add_rect(slide, MARGIN, footer_y, footer_w, footer_h, fill=WHITE,
             border_color=BLUE, border_w=Pt(2))
    add_text(slide, MARGIN, footer_y + Mm(12), footer_w, Mm(14),
             "INSTITUCIONES ORGANIZADORAS Y AUSPICIANTES", size=15, bold=True,
             color=BLUE, align=PP_ALIGN.CENTER, font=FONT_SANS)

    logo_files = ["logo_solabima.png", "logo_una.png", "logo_fpuna.png",
                  "logo_arasy.png", "logo_custom1.png", "logo_custom2.png"]
    n = len(logo_files)
    logo_h = Mm(45)
    logo_w = Mm(90)
    lgap = Mm(15)
    total_w = n * logo_w + (n - 1) * lgap
    start_x = MARGIN + (footer_w - total_w) / 2
    logo_y = footer_y + footer_h - logo_h - Mm(20)
    for i, fname in enumerate(logo_files):
        path = os.path.join(LOGOS_DIR, fname)
        lx = start_x + i * (logo_w + lgap)
        slide.shapes.add_picture(path, lx, logo_y, height=logo_h)

    prs.save(OUT_PATH)
    print(f"wrote {OUT_PATH}")
    if os.path.exists(QR_PATH):
        os.remove(QR_PATH)


if __name__ == "__main__":
    main()
