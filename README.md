# aidoc-export
Convert LLM answers written in Markdown into **DOCX** and **PDF** reports. One Python module, three dependencies (`markdown-it-py`, `python-docx`, `reportlab`).

## Install

```bash
pip install -e .           # library + CLI
pip install -e ".[api]"    # plus the FastAPI example
```

## Use

```python
from aidoc_export import ExportConfig, Source, export_docx, export_pdf

cfg = ExportConfig(title="Weekly summary", header=True,
                   font_path="fonts/NanumGothic.ttf",
                   sources=[Source("Market report 2025", url="https://example.com/r.pdf", page=42)])

export_pdf(markdown, "summary.pdf", cfg)
docx_bytes = export_docx(markdown, config=cfg)   # returns bytes when no output is given
```

```bash
aidoc-export answer.md -o answer.pdf --title "Weekly" --header --font fonts/NanumGothic.ttf
```

For an HTTP endpoint, see the FastAPI example in the Korean README (`POST /export` with `{markdown, format, title, sources}` returns the file).

## Notes

- PDFs need a **TrueType (.ttf)** font with the glyphs you use. Lookup order: `font_path` → `AIDOC_FONT` / `AIDOC_FONT_BOLD` env vars → common system locations. CFF-based `.otf` / `.ttc` fonts (e.g. Noto Sans CJK) are not supported by reportlab and raise `FontNotFoundError`.
- Supported: headings, bold/strike/code/links, `==highlight==`, nested and task lists, GFM tables with alignment (header row repeats), quotes, code blocks, rules, `<!-- pagebreak -->`, header text, `n / total` page numbers, a numbered sources section.
- Not supported: images, table cell merging, table of contents. Italic is not visible in PDFs unless the font has an italic face.
- Ordered lists in DOCX are written as literal numbers rather than Word auto-numbering.
