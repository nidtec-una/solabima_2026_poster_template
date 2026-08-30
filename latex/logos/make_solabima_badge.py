"""Generate a high-resolution square "event badge" version of the
SOLABIMA 2026 logo, matching the layout of the congress's official
Instagram graphic (../../latex/logos/solabima_ig.png: hand + wordmark
up top, congress name below, diagonal-cut ribbon with location/date
at the bottom) but as a plain rectangle on white instead of a circle
on a dark background, to match this template's design system.

Uses the real official flat logo (logo_solabima.png, already
downloaded from solabima2026.pol.una.py) as the centerpiece -- only
the surrounding badge (congress name, ribbon) is newly composed here,
not a redrawing of the hand/wordmark artwork itself.

Run: python3 make_solabima_badge.py
"""
from PIL import Image, ImageDraw, ImageFont

W, H = 2200, 2400

BLUE = (28, 110, 168)
ORANGE = (241, 150, 81)
WHITE = (255, 255, 255)

CONGRESS_LINES = [
    "XIV CONGRESO DE LA",
    "SOCIEDAD LATINOAMERICANA",
    "DE BIOLOGÍA MATEMÁTICA",
]
LOCATION_LINE = "San Lorenzo, Paraguay 2026"
DATE_LINE = "5 – 9 Octubre"


def font(path_options, size):
    for path in path_options:
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def text_w(draw, text, fnt):
    bbox = draw.textbbox((0, 0), text, font=fnt)
    return bbox[2] - bbox[0]


def ribbon_polygon(x0, y0, x1, y1, notch):
    """Rectangle with a concave V notch cut into the left edge,
    pointing right into the ribbon body (classic ribbon-tail shape),
    matching the reference graphic's banner."""
    return [
        (x0, y0),
        (x1, y0),
        (x1, y1),
        (x0, y1),
        (x0 + notch, (y0 + y1) / 2),
    ]


def main():
    img = Image.new("RGB", (W, H), WHITE)
    draw = ImageDraw.Draw(img)

    bold = ["/System/Library/Fonts/HelveticaNeue.ttc",
            "/System/Library/Fonts/Helvetica.ttc"]

    # Hand + wordmark logo, real asset, top area.
    logo = Image.open("logo_solabima.png").convert("RGBA")
    logo_w = int(W * 0.86)
    logo_h = int(logo_w * logo.height / logo.width)
    logo_resized = logo.resize((logo_w, logo_h), Image.LANCZOS)
    logo_x = (W - logo_w) // 2
    logo_y = 90
    img.paste(logo_resized, (logo_x, logo_y), logo_resized)

    # Congress name, three centered lines, in blue (white text only
    # made sense on the reference's dark background).
    name_font = font(bold, 78)
    y = logo_y + logo_h + 40
    for line in CONGRESS_LINES:
        w = text_w(draw, line, name_font)
        draw.text(((W - w) / 2, y), line, font=name_font, fill=BLUE)
        y += 96

    # Orange ribbon banner with the ribbon-tail notch, bottom area.
    location_font = font(bold, 68)
    date_font = font(bold, 62)
    banner_h = 230
    banner_top = y + 50
    margin_x = 90
    notch = 70
    draw.polygon(
        ribbon_polygon(margin_x, banner_top, W - margin_x,
                        banner_top + banner_h, notch),
        fill=ORANGE,
    )
    lines = [(LOCATION_LINE, location_font, 34), (DATE_LINE, date_font, 122)]
    for line, fnt, ty in lines:
        w = text_w(draw, line, fnt)
        draw.text(((W - w) / 2 + notch / 3, banner_top + ty), line, font=fnt,
                   fill=WHITE)

    img.save("logo_solabima_badge.png")
    print("wrote logo_solabima_badge.png", img.size)


if __name__ == "__main__":
    main()
