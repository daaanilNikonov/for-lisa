#!/usr/bin/env python3
"""Build Forus-branded PDF competence card for Юлиана Юнусова (1С-ЭПД)."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "epd-competence-center"
PHOTO = OUT_DIR / "yuliana_cutout.png"
PHOTO_FALLBACK = OUT_DIR / "yuliana_photo.jpg"
PDF_OUT = OUT_DIR / "Yuliana_Yunusova_1C_EPD_competence_card.pdf"
PNG_OUT = OUT_DIR / "Yuliana_Yunusova_1C_EPD_competence_card.png"

# ГК Форус — тёмный шаблон
BLUE = (38, 166, 224)  # #26A6E0
NEAR_BLACK = (26, 26, 26)  # #1A1A1A
CARD = (63, 63, 63)  # #3F3F3F
SOFT = (191, 191, 191)  # #BFBFBF
WHITE = (255, 255, 255)
MUTED = (118, 118, 119)  # #767677

W, H = 1600, 900  # landscape card for chat / print preview

FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"


def font(path: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, size)


def circle_mask(size: int) -> Image.Image:
    m = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(m)
    d.ellipse((0, 0, size - 1, size - 1), fill=255)
    return m


def prepare_avatar(src: Path, size: int = 520) -> Image.Image:
    img = Image.open(src).convert("RGBA")
    # Prefer upper body / face for circular crop
    w, h = img.size
    side = min(w, h)
    left = (w - side) // 2
    top = max(0, int(h * 0.02))
    if top + side > h:
        top = h - side
    img = img.crop((left, top, left + side, top + side)).resize(
        (size, size), Image.Resampling.LANCZOS
    )
    mask = circle_mask(size)
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(img, (0, 0), mask)
    # Soft blue ring
    ring = Image.new("RGBA", (size + 16, size + 16), (0, 0, 0, 0))
    rd = ImageDraw.Draw(ring)
    rd.ellipse((0, 0, size + 15, size + 15), outline=BLUE + (230,), width=6)
    ring.paste(out, (8, 8), out)
    return ring


def draw_card() -> Image.Image:
    canvas_img = Image.new("RGB", (W, H), NEAR_BLACK)
    draw = ImageDraw.Draw(canvas_img)

    # Atmospheric gradient panel (left)
    for x in range(0, 520):
        t = x / 520
        r = int(NEAR_BLACK[0] + (CARD[0] - NEAR_BLACK[0]) * t * 0.35)
        g = int(NEAR_BLACK[1] + (CARD[1] - NEAR_BLACK[1]) * t * 0.35)
        b = int(NEAR_BLACK[2] + (BLUE[2] - NEAR_BLACK[2]) * t * 0.12)
        draw.line([(x, 0), (x, H)], fill=(r, g, b))

    # Top accent line
    draw.rectangle([0, 0, W, 8], fill=BLUE)

    # Bottom soft bar
    draw.rectangle([0, H - 56, W, H], fill=(20, 20, 20))
    draw.rectangle([0, H - 56, 280, H - 52], fill=BLUE)

    # Brand
    brand = font(FONT_BOLD, 28)
    draw.text((64, 40), "ГК ФОРУС", font=brand, fill=WHITE)
    sub = font(FONT_REG, 18)
    draw.text((64, 78), "Центр продаж", font=sub, fill=SOFT)

    # Avatar
    photo_path = PHOTO if PHOTO.exists() else PHOTO_FALLBACK
    avatar = prepare_avatar(photo_path, 500)
    # subtle glow behind avatar
    glow = Image.new("RGBA", (560, 560), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gd.ellipse((20, 20, 540, 540), fill=BLUE + (40,))
    glow = glow.filter(ImageFilter.GaussianBlur(28))
    canvas_rgba = canvas_img.convert("RGBA")
    canvas_rgba.alpha_composite(glow, (70, 170))
    canvas_rgba.alpha_composite(avatar, (98, 188))
    canvas_img = canvas_rgba.convert("RGB")
    draw = ImageDraw.Draw(canvas_img)

    # Text block
    tx = 680
    ty = 170
    label = font(FONT_BOLD, 20)
    draw.text((tx, ty), "ЦЕНТР КОМПЕТЕНЦИЙ", font=label, fill=BLUE)

    name_f = font(FONT_BOLD, 54)
    draw.text((tx, ty + 46), "Юлиана Юнусова", font=name_f, fill=WHITE)

    role_f = font(FONT_REG, 24)
    draw.text((tx, ty + 120), "Проект 1С-ЭПД · Центр продаж", font=role_f, fill=SOFT)

    # Divider
    draw.rectangle([tx, ty + 170, tx + 420, ty + 174], fill=BLUE)

    body = font(FONT_REG, 22)
    lines = [
        "Кратко: эксперт по продажам и сложным",
        "клиентским кейсам сервиса 1С-ЭПД.",
        "",
        "Можно обращаться по вопросам:",
        "• выбор решения (1С-ЭПД / Доки / Клиент ЭДО)",
        "• тарифы, пакеты титулов и акции",
        "• сложные сценарии по клиентам",
        "• возвраты и нестандартные ситуации",
        "• подпись, МЧД и линия поддержки",
    ]
    y = ty + 200
    for line in lines:
        color = WHITE if line.startswith("Можно") or line.startswith("Кратко") else SOFT
        if line.startswith("•"):
            color = WHITE
        draw.text((tx, y), line, font=body, fill=color)
        y += 34 if line else 18

    # Footer
    foot = font(FONT_REG, 16)
    draw.text(
        (64, H - 38),
        "Сервис 1С-ЭПД  ·  внутренняя карточка центра компетенций",
        font=foot,
        fill=MUTED,
    )
    draw.text((W - 64, H - 38), "forus.ru", font=foot, fill=BLUE, anchor="rm")

    return canvas_img


def export_pdf(img: Image.Image) -> None:
    # Physical size ~ 180×101 mm (landscape announcement card)
    width_pt = 510  # ~180 mm
    height_pt = 287  # ~101 mm
    c = canvas.Canvas(str(PDF_OUT), pagesize=(width_pt, height_pt))
    c.drawImage(ImageReader(img), 0, 0, width=width_pt, height=height_pt, mask="auto")
    c.showPage()
    c.save()


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    img = draw_card()
    img.save(PNG_OUT, "PNG", optimize=True)
    export_pdf(img)
    print(f"PNG: {PNG_OUT}")
    print(f"PDF: {PDF_OUT}")


if __name__ == "__main__":
    main()
