# PDF to Word Transcription Project

This project transcribes a Chinese PDF textbook (Chapter 4 - Calculus) into a Word document while preserving the original formatting.

## Approach: Hybrid Math Handling

The project uses a **hybrid approach** for math formulas:

| Type | Method | Example |
|------|--------|---------|
| **Standard math** | Native OMML equations | Fractions, limits, exponents |
| **Simple inline** | Unicode symbols | x², →, ∞, ≤ |
| **Cover pages** | Image crops | Table of contents, title page |

**Why this approach?**
- **OMML equations** are editable, searchable, and scale perfectly in Word
- **Unicode** works well for simple notation within text
- **Image crops** are only used for complex layouts that OMML can't represent

## Project Structure

```
PDF-extractor/
├── input/
│   ├── chapter-4.pdf      # Source PDF file
│   └── pages/             # Extracted page images (PNG)
├── crops/                 # Cropped math formula images
├── output/
│   └── chapter4_pilot.docx  # Generated Word document
├── extract_pages.py       # Extract PDF pages to images
├── crop_math.py          # Crop math formulas from page images
├── build_pilot.js        # Build Word document with docx library
└── README.md
```

## Workflow

### 1. Extract PDF Pages
```bash
python extract_pages.py
```
This extracts the first 5 pages from `input/chapter-4.pdf` as PNG images at 150 DPI into `input/pages/`.

### 2. Crop Math Formulas
```bash
python crop_math.py
```
This crops math expressions from the page images based on predefined coordinates in `crop_math.py`. Each crop is saved to the `crops/` folder.

### 3. Build Word Document
```bash
node build_pilot.js
```
This assembles the Word document by:
- Adding transcribed Chinese text
- Inserting cropped math formula images
- Applying formatting (fonts, spacing, headers)

## Key Design Decisions

### Math Expressions
Math formulas are kept as images rather than attempting to convert to equations because:
1. Chinese textbook math notation can be complex
2. Preserves exact original formatting
3. Avoids transcription errors in mathematical notation

### Crop Guidelines
When adjusting `crop_math.py`:
- **Crop ONLY the math formula**, not surrounding Chinese text
- Include full multi-line formulas (don't cut off parts)
- Chinese text before/after formulas should be transcribed as text in `build_pilot.js`
- Add small white padding around crops for better appearance

### Text vs Image Decision
- **Chinese text**: Transcribe as text (searchable, editable)
- **Math formulas**: Crop as images (preserves exact notation)
- **Example labels** (例 4-1): Transcribe as text
- **Solution explanations** (解：...): Transcribe as text

## Crop Coordinate Format

In `crop_math.py`, crops are defined as:
```python
"crop_name": ("page.png", (left, upper, right, lower))
```
Coordinates are in pixels based on the extracted image resolution (150 DPI).

## Dependencies

### Python
- `pymupdf` (fitz) - PDF page extraction
- `pillow` - Image processing

### Node.js
- `docx` - Word document generation

## Common Issues & Solutions

### Issue: Math formula cut off
Adjust the crop coordinates in `crop_math.py` to include the full formula.

### Issue: Chinese text appears in cropped image
Move the crop boundaries to exclude text, then add the text separately in `build_pilot.js`.

### Issue: Duplicated content
Ensure text is not both transcribed AND included in a crop. Check `build_pilot.js` for text that should be removed if it's in an image.

## File Locations
- **Source PDF**: `input/chapter-4.pdf`
- **Output Word**: `output/chapter4_pilot.docx`

## Running the Full Pipeline
```bash
# 1. Set up Python environment
python -m venv .venv
source .venv/bin/activate
pip install pymupdf pillow

# 2. Install Node dependencies
npm install docx

# 3. Run pipeline
python extract_pages.py
python crop_math.py
node build_pilot.js
```

The final Word document will be at `output/chapter4_pilot.docx`.
