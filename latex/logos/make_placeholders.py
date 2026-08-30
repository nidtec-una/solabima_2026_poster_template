"""Generate simple placeholder logo images.

logo_solabima.png is the real official logo (downloaded from
solabima2026.pol.una.py) - do not overwrite it with this script. This
only (re)generates the generic placeholders: logo_header_right (the
logo slot next to the title, opposite SOLABIMA) and two "custom"
slots for funding-agency logos in the optional footer band. Run
`python3 make_placeholders.py` to regenerate them if needed.
"""
from PIL import Image, ImageDraw, ImageFont

LOGOS = [
    ("logo_header_right", "Logo de\ntu Uni"),
    ("logo_custom1", "LOGO"),
    ("logo_custom2", "LOGO"),
]

SIZE = (340, 240)
BG = (240, 240, 240)
BORDER = (150, 150, 150)
TEXT = (90, 90, 90)


def make_logo(label: str, path: str) -> None:
    img = Image.new("RGB", SIZE, BG)
    draw = ImageDraw.Draw(img)
    draw.rectangle([2, 2, SIZE[0] - 3, SIZE[1] - 3], outline=BORDER, width=4)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 44)
    except OSError:
        font = ImageFont.load_default()
    lines = label.split("\n")
    total_h = sum(draw.textbbox((0, 0), ln, font=font)[3] for ln in lines) + 10 * (len(lines) - 1)
    y = (SIZE[1] - total_h) / 2
    for ln in lines:
        bbox = draw.textbbox((0, 0), ln, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        draw.text(((SIZE[0] - w) / 2, y), ln, fill=TEXT, font=font)
        y += h + 10
    img.save(path)


if __name__ == "__main__":
    for name, label in LOGOS:
        make_logo(label, f"{name}.png")
        print(f"wrote {name}.png")
