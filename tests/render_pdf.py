"""Render the generated sample PDF to PNGs for visual inspection."""
import pathlib
import sys

import fitz  # PyMuPDF

pdf = sys.argv[1] if len(sys.argv) > 1 else "tests/output/quarterly-report.pdf"
out = pathlib.Path(sys.argv[2] if len(sys.argv) > 2 else "tests/output/render")
out.mkdir(parents=True, exist_ok=True)

doc = fitz.open(pdf)
print("pages:", len(doc))
for i, page in enumerate(doc):
    if i >= 7:
        break
    page.get_pixmap(dpi=90).save(str(out / f"p{i + 1}.png"))
    print("saved p%d" % (i + 1))
