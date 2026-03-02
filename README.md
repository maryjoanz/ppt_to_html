Here’s your README converted into clean, structured Markdown while preserving all headings, lists, code blocks, and semantics from the original text.

---

# PPTX to Accessible HTML Converter

Converts PowerPoint (`.pptx`) files to accessible HTML, preserving SME‑authored alt text, heading structure, images, tables, and lists. Designed for screen reader users, the script strips out decorative images.

---

## Requirements

Python 3.7 or higher and one dependency:

```bash
pip install python-pptx
```

---

## Usage

**Convert a single file:**

```bash
python pptx_to_accessible_html.py presentation.pptx
```

Output will be saved as `presentation.html` in the same folder.

**Specify a custom output filename:**

```bash
python pptx_to_accessible_html.py presentation.pptx -o output.html
```

**Convert an entire folder of `.pptx` files:**

```bash
python pptx_to_accessible_html.py ./my_slides_folder/
```

Each file will produce a matching `.html` file in the same folder.

**Include speaker notes:**

```bash
python pptx_to_accessible_html.py presentation.pptx --include-notes
```

---

## What the Script Converts

- Slide title → `<h2>`
- Bold/large in‑slide text → `<h3>`
- Body text → `<p>`
- Bullet points → `<ul>` / `<li>`
- Tables → `<table>` with `<thead>` and `<th scope="col">`
- Images with alt text → `<img>` with SME‑authored `alt` attribute
- Decorative images → skipped entirely
- Speaker notes (optional) → `<aside>`

---

## Alt Text and Decorative Images

The script reads alt text exactly as authored in PowerPoint — no AI generation or modification.

**To add alt text in PowerPoint:**

Right‑click an image → **Edit Alt Text** → type a description.

**To mark an image as decorative:**

Right‑click an image → **Edit Alt Text** → check *Mark as decorative*.

The script will skip decorative images entirely and omit them from the HTML output.

Images with neither alt text nor a decorative flag receive a fallback description of:

> “Image on slide N”

These should be reviewed and updated in the source PowerPoint before final conversion.

---

## Image Sizing

Images are rendered at their actual PowerPoint dimensions, converted to pixels at 96 DPI.  
A `max-width: 100%` rule ensures they scale down gracefully on narrow screens without distortion.

---

## Accessibility Features

- Skip navigation link (“Skip to main content”) at the top of every page  
- Semantic HTML5 landmarks (`<main>`, `<section>`, `<aside>`)  
- `aria-label` on each slide section  
- Table headers use `scope="col"` for screen reader compatibility  
- Self‑contained output — images embedded as base64 data URIs, requiring no external assets  

---

## Auditing Images Before Converting

Use the companion script `inspect_pptx_images.py` to audit every image in a deck:

```bash
python inspect_pptx_images.py presentation.pptx
```

This prints each image’s name, alt text, and decorative status — useful for catching missing or incomplete alt text before generating the final HTML.

---

## Recommended Workflow

1. SMEs author alt text and mark decorative images in PowerPoint  
2. Run `inspect_pptx_images.py` to verify all images are correctly tagged  
3. Run `pptx_to_accessible_html.py` to generate the HTML  
4. Test with a screen reader  

