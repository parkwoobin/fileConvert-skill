---
name: file-convert-skill
description: Convert HWPX, HWP, Word, PDF, and PowerPoint documents locally and verify their layout. Use when a user wants a document in another supported format or asks whether formatting and page layout survived a conversion.
---

Use the bundled converter on a Windows desktop with the Office apps required for the requested formats. Read [setup and format details](README.md) if the required apps or the correct conversion mode are unclear.

1. Resolve the skill folder from this `SKILL.md`. Run `python <skill-folder>/scripts/convert.py --list` to check the requested direction.
2. Use absolute input and output paths. Run `python <skill-folder>/scripts/convert.py INPUT OUTPUT --report REPORT.json`. Choose new paths for both the output and report.
3. Inspect the report before describing the result. The default saves output only when the rendered pages match. A render match measures appearance at the tested settings. Check requested editable structures, such as tables, formulas, footnotes and page fields, in the output independently.

When a user accepts visual differences, add `--allow-differences`. For an editable document whose pages may reflow, also add `--allow-page-count-change`. PDF → HWP/HWPX still requires the original page count.

For HWPX → DOCX, `--hwpx-structure` prioritizes editable paragraphs, pictures and tables and has format-specific omissions. For PDF → DOCX, use `--pdf-reflow` when semantic paragraphs and tables matter, or `--pdf-objects` when preserving the placement of editable text boxes and shapes matters. The converter chooses between those modes by default. For PDF → PPT/PPTX, explain that the resulting slides contain rendered page images. Scanned text needs OCR before a text-editable conversion.
