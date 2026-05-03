"""
Crop math expressions and graphs from page images.

WORKFLOW:
1. Extract PDF pages to images (e.g., using pdf2image or similar)
2. Run this script to crop math formulas from the page images
3. Run build_pilot.js to assemble the Word document

CROP GUIDELINES:
- Crop ONLY the math formula, NOT Chinese text
- Include full multi-line formulas (don't cut off)
- Text before/after formulas should be transcribed as text in build_pilot.js

Each crop is defined as (page, x1, y1, x2, y2) in pixel coords.
Adjust coordinates based on your PDF extraction resolution.
"""
from PIL import Image
import os

# Use relative paths from project root
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(SCRIPT_DIR, "input", "pages")  # Extracted page images
OUT = os.path.join(SCRIPT_DIR, "crops")           # Output cropped images
os.makedirs(OUT, exist_ok=True)

# Crop definitions: name -> (page_file, (left, upper, right, lower))
# IMPORTANT: Coordinates should exclude Chinese text, include only formulas
# Adjust these coordinates based on your PDF extraction resolution

crops = {
    # --- Page 3: Indeterminate forms (∞/∞ and 0/0) — fraction bars at y=547 ---
    # Keep as OMML, but crop as fallback: x=457-554, y≈520-580
    "p3_indeterminate": ("3.png", (420, 515, 590, 585)),

    # L'Hopital theorem — conditions block (no fractions, text lines at y=1067-1141)
    "p3_lhopital_cond": ("3.png", (140, 1055, 910, 1155)),

    # L'Hopital theorem — conclusion formula (fraction bars at y=1180, text y=1145-1219)
    "p3_lhopital_conc": ("3.png", (140, 1135, 570, 1230)),

    # --- Page 4: Examples ---
    # Example 4-1: Problem formula (fraction bar at y=404, text y=384-434)
    "p4_ex1_problem": ("4.png", (140, 375, 520, 445)),

    # Example 4-1: Solution chain (fraction bars at y=513, text y=479-552)
    "p4_ex1_solution": ("4.png", (140, 465, 910, 565)),

    # Example 4-2: Problem formula (fraction bar at y=581, text y=558-610)
    "p4_ex2_problem": ("4.png", (140, 548, 520, 620)),

    # Example 4-2: Solution formula (fraction bars at y=679, text y=655-708)
    "p4_ex2_solution": ("4.png", (140, 645, 910, 720)),

    # Example 4-3: Problem formula (fraction bar at y=776, text y=767-807)
    "p4_ex3_problem": ("4.png", (140, 757, 520, 820)),

    # Example 4-3: Full solution — all 3 lines (fraction bars y=875,939; text y=852-968)
    "p4_ex3_solution": ("4.png", (140, 838, 910, 978)),

    # Example 4-4: Problem formula (fraction bars y=1123-1148, text y=1112-1179)
    "p4_ex4_problem": ("4.png", (140, 1103, 520, 1190)),

    # Example 4-4: Full solution — both lines (fraction bars y=1193-1287, text y=1181-1320)
    "p4_ex4_solution": ("4.png", (140, 1175, 910, 1325)),

    # --- Page 5: Famous limits ---
    # Famous limit 1: sin(x)/x (fraction bar at y=414, text y=394-443)
    "p5_limit1": ("5.png", (140, 383, 710, 455)),

    # Famous limit 2: (1+1/n)^n (fraction bar at y=476, text y=456-505)
    "p5_limit2": ("5.png", (140, 443, 710, 518)),

    # --- Cover pages ---
    "cover_toc": ("1.png", (50, 50, 874, 1266)),
    "cover_chapter": ("2.png", (50, 50, 874, 1266)),
}

for name, (fname, box) in crops.items():
    img = Image.open(f"{SRC}/{fname}")
    crop = img.crop(box)
    # Add small white padding
    padded = Image.new("RGB", (crop.width + 10, crop.height + 10), "white")
    padded.paste(crop, (5, 5))
    out_path = f"{OUT}/{name}.png"
    padded.save(out_path)
    print(f"Saved {out_path} ({padded.width}x{padded.height})")

print("Done.")
