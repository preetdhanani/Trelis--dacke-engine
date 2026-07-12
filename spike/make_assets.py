"""Throwaway: generate placeholder raster images for image-typed slots.
Not decorative filler in the product sense -- just test fixtures for S1.
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).parent / "assets"
OUT.mkdir(exist_ok=True)


def make_image(path, w, h, bg, label):
    img = Image.new("RGB", (w, h), bg)
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 36)
    except Exception:
        font = ImageFont.load_default()
    d.rectangle([8, 8, w - 8, h - 8], outline=(255, 255, 255), width=4)
    bbox = d.textbbox((0, 0), label, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text(((w - tw) / 2, (h - th) / 2), label, fill=(255, 255, 255), font=font)
    img.save(path)


# Sized roughly to their target slot geometry (aspect only matters a little; pptx will size/crop to the slot box)
make_image(OUT / "two_col_support.jpg", 1076, 870, (9, 87, 209), "SUPPORT IMAGE")
make_image(OUT / "image_only_hero.jpg", 2306, 1080, (17, 17, 17), "HERO IMAGE")
make_image(OUT / "exhibit_chart.jpg", 1680, 870, (255, 53, 49), "EXHIBIT VISUAL")
print("assets written to", OUT)
