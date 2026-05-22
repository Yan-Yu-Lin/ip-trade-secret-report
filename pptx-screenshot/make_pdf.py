#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11"
# dependencies = ["Pillow", "img2pdf"]
# ///

import img2pdf
from pathlib import Path

SLIDES_DIR = Path(__file__).parent / "slides"
OUTPUT = Path(__file__).parent / "presentation.pdf"

pngs = sorted(SLIDES_DIR.glob("slide-*.png"))
if not pngs:
    print("No slides found")
    exit(1)

with open(OUTPUT, "wb") as f:
    f.write(img2pdf.convert([str(p) for p in pngs], layout_fun=img2pdf.get_layout_fun((img2pdf.mm_to_pt(338.667), img2pdf.mm_to_pt(190.5)))))

print(f"Saved PDF with {len(pngs)} pages to {OUTPUT}")
