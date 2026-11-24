---
name: file-convert-skill
description: Convert HWPX, HWP, Word, PDF, and PowerPoint files through the supported pairs, and verify layout fidelity including fonts, spacing, and pagination by comparing rendered pages.
---

Use `python scripts/convert.py --list` to see the supported directions.
Run `python scripts/convert.py INPUT OUTPUT --report report.json`. Conversion verifies rendered pages by default and keeps the output only if the pages match. Read the report before claiming a match. A match is limited to the tested render settings and does not establish editability or structural equivalence.

For HWPX/HWP and DOCX conversion, use the native Hancom engine and compare against Microsoft Word. Check editable tables, images, equations, footnotes, endnotes, headers, footers, page numbering, and font substitution separately in the destination application when those matter. For editable HWPX text, images and tables, use the optional `--hwpx-structure` importer and check the documented omissions. Never rename a file to simulate conversion.

PDF to a document uses pdf2docx reconstruction and can change typography and layout. PDF to PowerPoint places a rendered image onto each slide and does not reconstruct editable text. Explain these limits before interpreting the results. Use `--allow-differences` only when the user accepts a result that differs visually. A page count mismatch still blocks saving. Check rendered pages before sharing any result with known differences. See [setup and verification](README.md).

For a fixed-layout PDF with selectable text, try `--pdf-objects` with DOCX output. It reconstructs positioned Word text boxes, shapes, and pictures. Check selected pages in desktop Word for visibility, clipping, and editability before sharing. This mode approximates typography and does not recreate semantic tables or reflowing paragraphs.
