"""
measure_reference_design.py

Reverse-engineering tool, not a guessing tool: extracts exact, measured
design values from a reference Transfer News PDF instead of eyeballing a
rendered preview. Point it at the mission's real sample document to get hard
numbers to hardcode into generate_transfer_document.py.

Two measurement paths:
  - Vector/text (pdfplumber, already a project dependency): page size,
    fonts + sizes actually used, and the exact fill color of every drawn
    rectangle (table cells, badges, banners) straight from the PDF's color
    operators -- not sampled/estimated from pixels.
  - Raster (PyMuPDF, optional -- only needed for this diagnostic, not for
    the generator itself): extracts the literal embedded watermark image
    bytes plus its exact placement rectangle and, if the image carries a
    soft mask, its real alpha data. Install with `pip install pymupdf` if
    it isn't already available.

Usage:
    python measure_reference_design.py reference/sample_transfer_news.pdf
    python measure_reference_design.py reference/sample_transfer_news.pdf --pages 1,2,6
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

import pdfplumber

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


def _rgb_to_hex(color) -> str:
    if color is None:
        return "none"
    if isinstance(color, (int, float)):
        v = round(color * 255)
        return f"{v:02X}{v:02X}{v:02X}"
    if len(color) == 3:
        r, g, b = (round(c * 255) for c in color)
        return f"{r:02X}{g:02X}{b:02X}"
    if len(color) == 4:
        return f"cmyk{tuple(round(c, 3) for c in color)}"
    return str(color)


def measure_page_vector(page, label: str) -> None:
    print(f"\n=== {label}: {page.width:.1f} x {page.height:.1f} pt "
          f"({page.width / 72:.3f}in x {page.height / 72:.3f}in) ===")

    fonts = Counter((round(c["size"], 1), c["fontname"]) for c in page.chars)
    if fonts:
        print("Fonts/sizes in use (size pt, font): char count")
        for (size, font), count in fonts.most_common(12):
            print(f"  {size:>5}pt  {font:<30} {count}")

    fills = Counter(_rgb_to_hex(r.get("non_stroking_color")) for r in page.rects)
    if fills:
        print("Rectangle fill colors on this page: count")
        for color, count in fills.most_common(20):
            print(f"  #{color}: {count}")

    strokes = Counter(
        (_rgb_to_hex(r.get("stroking_color")), round(r.get("linewidth", 0), 2))
        for r in page.rects if r.get("stroking_color") is not None
    )
    if strokes:
        print("Border colors/widths on this page: count")
        for (color, width), count in strokes.most_common(10):
            print(f"  #{color} @ {width}pt: {count}")

    if page.images:
        print("Embedded images on this page (exact placement, from the PDF itself):")
        for img in page.images:
            w_in, h_in = img["width"] / 72, img["height"] / 72
            x0_in, top_in = img["x0"] / 72, img["top"] / 72
            cx_in, cy_in = x0_in + w_in / 2, top_in + h_in / 2
            print(
                f"  name={img.get('name')} bbox=({img['x0']:.1f},{img['top']:.1f})-"
                f"({img['x1']:.1f},{img['bottom']:.1f})pt "
                f"size={w_in:.3f}in x {h_in:.3f}in "
                f"center=({cx_in:.3f}in, {cy_in:.3f}in) srcsize={img.get('srcsize')}"
            )
    else:
        print("No embedded raster images detected on this page via pdfplumber.")


def measure_page_raster(doc, page_number_zero_based: int, out_dir: Path, label: str) -> None:
    if fitz is None:
        return
    page = doc[page_number_zero_based]
    images = page.get_images(full=True)
    if not images:
        return
    print(f"\n--- {label}: raster image extraction (PyMuPDF) ---")
    out_dir.mkdir(parents=True, exist_ok=True)
    for index, image_info in enumerate(images):
        xref = image_info[0]
        smask_xref = image_info[1]
        base = doc.extract_image(xref)
        ext = base["ext"]
        out_path = out_dir / f"{label.replace(' ', '_')}_image{index}_xref{xref}.{ext}"
        out_path.write_bytes(base["image"])
        rects = page.get_image_rects(xref)
        rect_desc = [
            f"({r.x0 / 72:.3f}in,{r.y0 / 72:.3f}in)-({r.x1 / 72:.3f}in,{r.y1 / 72:.3f}in)"
            for r in rects
        ]
        has_alpha = smask_xref != 0
        print(
            f"  xref {xref}: extracted -> {out_path} | src pixels={base.get('width')}x{base.get('height')} "
            f"| has real alpha/soft-mask={has_alpha} | placed at {rect_desc}"
        )
        if has_alpha:
            smask = doc.extract_image(smask_xref)
            mask_path = out_dir / f"{label.replace(' ', '_')}_image{index}_xref{xref}_alpha.{smask['ext']}"
            mask_path.write_bytes(smask["image"])
            print(f"    alpha channel extracted -> {mask_path} (this IS the real opacity data, not an estimate)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Measure exact design values from a reference Transfer News PDF.")
    parser.add_argument("pdf", help="Path to the reference PDF (the sample design spec).")
    parser.add_argument("--pages", default="1", help="Comma-separated 1-based page numbers to inspect (default: 1).")
    parser.add_argument("--extract-images", default="reference/_extracted_images",
                         help="Folder to write extracted raster images/alpha masks into (requires PyMuPDF).")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        sys.exit(f"Not found: {pdf_path}")

    page_numbers = [int(p) for p in args.pages.split(",") if p.strip()]

    with pdfplumber.open(str(pdf_path)) as pdf:
        for n in page_numbers:
            if 1 <= n <= len(pdf.pages):
                measure_page_vector(pdf.pages[n - 1], f"Page {n}")
            else:
                print(f"Page {n} out of range (document has {len(pdf.pages)} pages)")

    if fitz is None:
        print(
            "\nPyMuPDF not installed -- skipping raster/watermark image extraction. "
            "Run `pip install pymupdf` and re-run this script to pull the exact "
            "embedded watermark asset and its real alpha data instead of estimating opacity."
        )
        return

    doc = fitz.open(str(pdf_path))
    out_dir = Path(args.extract_images)
    for n in page_numbers:
        if 1 <= n <= doc.page_count:
            measure_page_raster(doc, n - 1, out_dir, f"page{n}")
    doc.close()


if __name__ == "__main__":
    main()
