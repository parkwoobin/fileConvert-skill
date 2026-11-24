---
name: file-convert-skill
description: Convert HWPX, HWP, Word, PDF, and PowerPoint documents locally and verify their layout. Use when a user wants a document in another supported format or asks whether formatting and page layout survived a conversion.
---

Use the bundled converter on a Windows desktop with the Office apps required for the requested formats. Read [setup and format details](README.md) if the required apps or the correct conversion mode are unclear.

1. Resolve the skill folder from this `SKILL.md`. Run `python <skill-folder>/scripts/convert.py --list` to check the requested direction.
2. Use absolute input and output paths. Run `python <skill-folder>/scripts/convert.py INPUT OUTPUT --report REPORT.json`. Choose new paths for both the output and report.
3. Inspect the report before describing the result. HWPX → DOCX uses editable XML import by default and saves when the rendered page count matches. Use `--require-visual-match` if the user requires identical rendering. Other routes require a render match by default. Check requested editable structures, such as tables, formulas, footnotes and page fields, in the output independently.

When a user accepts visual differences, add `--allow-differences`. For an editable document whose pages may reflow, also add `--allow-page-count-change`. PDF → HWP/HWPX still requires the original page count.

For PDF → DOCX, use `--pdf-reflow` when semantic paragraphs and tables matter, or `--pdf-objects` when preserving the placement of editable text boxes and shapes matters. The converter chooses between those modes by default. For PDF → PPT/PPTX, explain that the resulting slides contain rendered page images. Scanned text needs OCR before a text-editable conversion.
