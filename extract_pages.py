"""Extract pages from PDF to PNG images using PyMuPDF."""
import fitz  # PyMuPDF
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_PATH = os.path.join(SCRIPT_DIR, "input", "chapter-4.pdf")
OUT_DIR = os.path.join(SCRIPT_DIR, "input", "pages")
os.makedirs(OUT_DIR, exist_ok=True)

# Extract first 5 pages at 150 DPI
DPI = 150
ZOOM = DPI / 72.0  # PDF default is 72 DPI

doc = fitz.open(PDF_PATH)
for page_num in range(min(5, len(doc))):
    page = doc.load_page(page_num)
    mat = fitz.Matrix(ZOOM, ZOOM)
    pix = page.get_pixmap(matrix=mat)
    out_path = os.path.join(OUT_DIR, f"{page_num + 1}.png")
    pix.save(out_path)
    print(f"Saved {out_path} ({pix.width}x{pix.height})")

doc.close()
print("Done extracting pages.")
