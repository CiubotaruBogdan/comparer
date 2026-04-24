# Comparer

Compare two DOCX or editable PDF documents side by side, with paragraph-level and word-level difference highlighting.

---

## Features

- **Drag & drop** or click to load documents (`.docx`, `.pdf`)
- **Paragraph-level diff** — identifies equal, modified, added, and deleted paragraphs
- **Word-level diff** — inside modified paragraphs, highlights the exact words that changed
- **Side-by-side viewer** with synchronized scrolling
- **Similarity score** — overall percentage match between the two documents
- **Statistics bar** — counts for identical / modified / added / deleted paragraphs
- **Swap** — re-run the comparison with documents reversed
- **Export** — save a self-contained HTML report of the comparison

## Highlight legend

| Color | Meaning |
|---|---|
| 🔴 Red | Paragraph present only in Document 1 (deleted) |
| 🟢 Green | Paragraph present only in Document 2 (added) |
| 🟡 Yellow | Paragraph exists in both but was modified |

---

## Requirements

```
PyQt5 >= 5.15
python-docx >= 0.8.11
PyMuPDF >= 1.23
```

Install with:

```bash
pip install -r requirements.txt
```

## Run

```bash
python main.py
```

## Build standalone `.exe`

```bash
build.bat
```

Output: `dist/Comparer.exe` — no Python or dependencies required on the target machine.

> **Note:** PDF comparison works only on text-based (editable) PDFs. Scanned image PDFs require OCR pre-processing.
