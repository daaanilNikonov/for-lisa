#!/usr/bin/env python3
"""Build a 3x3 A4 PDF of autumn labels from the 3x4 source sheet."""

from pathlib import Path

import cv2
import img2pdf
import numpy as np
from PIL import Image


SRC = Path(__file__).resolve().parents[1] / "assets" / "source_labels_12.png"
# Fallback to Cursor-uploaded asset path if local copy missing
FALLBACK_SRC = Path(
    "/home/ubuntu/.cursor/projects/workspace/assets/9a26e90e-5b9a-4872-ab11-0ee6bba48df2.png"
)
OUT_DIR = Path(__file__).resolve().parent
DPI = 300


def upscale(pil_img: Image.Image, tw: int, th: int) -> Image.Image:
    arr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    h, w = arr.shape[:2]
    while w * 2 < tw and h * 2 < th:
        arr = cv2.resize(arr, (w * 2, h * 2), interpolation=cv2.INTER_LANCZOS4)
        h, w = arr.shape[:2]
    arr = cv2.resize(arr, (tw, th), interpolation=cv2.INTER_LANCZOS4)
    blur = cv2.GaussianBlur(arr, (0, 0), 1.0)
    arr = cv2.addWeighted(arr, 1.35, blur, -0.35, 0)
    return Image.fromarray(cv2.cvtColor(arr, cv2.COLOR_BGR2RGB))


def trim_white(crop: Image.Image, pad: int = 2) -> Image.Image:
    arr = np.array(crop)
    mask = ~((arr[:, :, 0] > 248) & (arr[:, :, 1] > 248) & (arr[:, :, 2] > 248))
    ys, xs = np.where(mask)
    if not len(xs):
        return crop
    x0 = max(0, int(xs.min()) - pad)
    y0 = max(0, int(ys.min()) - pad)
    x1 = min(crop.width, int(xs.max()) + 1 + pad)
    y1 = min(crop.height, int(ys.max()) + 1 + pad)
    return crop.crop((x0, y0, x1, y1))


def main() -> None:
    src = SRC if SRC.exists() else FALLBACK_SRC
    im = Image.open(src).convert("RGB")

    cols = [(4, 286), (288, 566), (569, 849)]
    rows = [(10, 333), (335, 645), (647, 955), (957, 1260)]

    labels = []
    for y0, y1 in rows:
        for x0, x1 in cols:
            labels.append(trim_white(im.crop((x0, y0, x1, y1))))

    selected = labels[:9]  # top 3 rows → 9 larger labels

    page_w = int(round(210 / 25.4 * DPI))
    page_h = int(round(297 / 25.4 * DPI))
    avg_aspect = sum(l.width / l.height for l in selected) / len(selected)

    margin = int(round(6 / 25.4 * DPI))
    gap = int(round(3.5 / 25.4 * DPI))
    label_w = (page_w - 2 * margin - 2 * gap) // 3
    label_h = int(round(label_w / avg_aspect))
    max_h = (page_h - 2 * margin - 2 * gap) // 3
    if label_h > max_h:
        label_h = max_h
        label_w = int(round(label_h * avg_aspect))

    total_w = 3 * label_w + 2 * gap
    total_h = 3 * label_h + 2 * gap
    ox = (page_w - total_w) // 2
    oy = (page_h - total_h) // 2

    page = Image.new("RGB", (page_w, page_h), (255, 255, 255))
    for i, lab in enumerate(selected):
        r, c = divmod(i, 3)
        page.paste(
            upscale(lab, label_w, label_h),
            (ox + c * (label_w + gap), oy + r * (label_h + gap)),
        )

    png_path = OUT_DIR / "autumn_labels_9_a4_300dpi.png"
    pdf_path = OUT_DIR / "autumn_labels_9.pdf"
    page.save(png_path, dpi=(DPI, DPI), compress_level=6)

    a4 = (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
    with open(pdf_path, "wb") as f:
        f.write(img2pdf.convert(png_path.as_posix(), layout_fun=img2pdf.get_layout_fun(a4)))

    preview = page.copy()
    preview.thumbnail((1100, 1550), Image.Resampling.LANCZOS)
    preview.save(OUT_DIR / "autumn_labels_9_preview.png", optimize=True)
    print(f"Wrote {pdf_path}")


if __name__ == "__main__":
    main()
