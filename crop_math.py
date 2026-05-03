"""
Crop math expressions and graphs from page images.
Each crop is defined as (page, x1, y1, x2, y2) in pixel coords (924x1316).
"""
from PIL import Image
import os

SRC = "/home/claude/chapter4_extracted"
OUT = "/home/claude/math_crops"
os.makedirs(OUT, exist_ok=True)

# Page 3 crops (0-indexed coords from 924x1316 image)
# The ∞/∞ and 0/0 fraction display at center of page 3
# Page 3: "∞/∞ 型和 0/0 型" centered display formula
# Page 3: L'Hopital definition block (4 lines of limits)
# Page 4: multiple worked examples
# Page 5: two limit display formulas (sin x/x and (1-cos x)/x)

crops = {
    # (page_file, label, (left, upper, right, lower))
    # --- Page 3 ---
    "p3_indeterminate_forms": ("3.jpeg", (230, 385, 700, 450)),   # ∞/∞ and 0/0 display
    "p3_lhopital_def": ("3.jpeg", (60, 720, 870, 880)),           # the theorem block with limits
    
    # --- Page 4 ---
    "p4_ex1_formula": ("4.jpeg", (140, 265, 880, 335)),           # lim ln x / (x-1) worked
    "p4_ex2_formula": ("4.jpeg", (140, 430, 880, 480)),           # lim 2^x / 8x
    "p4_ex3_formula": ("4.jpeg", (60, 570, 880, 660)),            # lim 3^x / x^3 full working
    "p4_ex3_formula2": ("4.jpeg", (60, 660, 880, 720)),           # continuation
    "p4_ex4_formula": ("4.jpeg", (60, 855, 880, 945)),            # lim √x / (ln x)^2
    "p4_ex4_formula2": ("4.jpeg", (60, 945, 880, 1010)),          # continuation
    
    # --- Page 5 ---
    "p5_sinx_limits": ("5.jpeg", (180, 310, 740, 420)),           # sin x/x and (1-cosx)/x display
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
