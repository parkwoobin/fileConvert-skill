"""FileConvert local CLI."""
import argparse
from contextlib import closing
from functools import partial
from io import BytesIO
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
from zipfile import ZipFile

HANCOM = {".hwp", ".hwpx"}
WORD = {".doc", ".docx"}
SLIDES = {".ppt", ".pptx"}
PAIRS = {
    (a, b)
    for left, right in [
        (".hwpx", ".docx"), (".hwp", ".docx"),
        *((".pdf", ext) for ext in sorted(HANCOM | WORD | SLIDES)),
    ]
    for a, b in [(left, right), (right, left)]
} | {(".doc", ".docx"), (".hwp", ".hwpx")}
DPI = 144
OLE = bytes.fromhex("d0cf11e0a1b11ae1")


def extension(path):
    return path.suffix.lower()


def validate(path):
    ext = extension(path)
    with path.open("rb") as source:
        signature = source.read(8)
    if ext in {".hwp", ".doc", ".ppt"}:
        valid = signature == OLE
    elif ext == ".pdf":
        valid = signature.startswith(b"%PDF-")
    else:
        with ZipFile(path) as package:
            expected = {
                ".hwpx": "Contents/content.hpf",
                ".docx": "word/document.xml",
                ".pptx": "ppt/presentation.xml",
            }[ext]
            valid = expected in package.namelist()
    if not valid:
        raise ValueError(f"Invalid or unsupported {ext} container: {path.name}")


def office(engine, source, target, fit_table_images=False):
    if sys.platform != "win32":
        raise RuntimeError(f"{engine} conversion requires a Windows desktop and the installed application")
    command = [
        sys.executable, str(Path(__file__).with_name("office.py")),
        engine, str(source), str(target),
        os.environ.get("FILECONVERT_HANCOM_PATH_MODULE", ""),
    ]
    if engine == "word" and fit_table_images:
        command.append("--fit-table-images")
    try:
        result = subprocess.run(command, capture_output=True, timeout=120, check=False)
    except subprocess.TimeoutExpired as error:
        raise RuntimeError(f"{engine} timed out. Check the application for a file access or licensing prompt.") from error
    if result.returncode:
        diagnostic = result.stderr.decode(errors="replace").strip().splitlines()
        raise RuntimeError(f"{engine}: {diagnostic[-1] if diagnostic else 'conversion failed'}")
    validate(target)


def render(source, target, fit_table_images=False):
    ext = extension(source)
    if ext == ".pdf":
        shutil.copyfile(source, target)
        return
    engine = "hancom" if ext in HANCOM else "word" if ext in WORD else "powerpoint"
    office(engine, source, target, fit_table_images)


def page_image(page):
    with closing(page.render(scale=DPI / 72, draw_annots=True)) as bitmap:
        return bitmap.to_pil().convert("RGB")


def pdf_to_slides(source, target):
    import pypdfium2 as pdfium
    from pptx import Presentation
    from pptx.util import Pt

    deck = Presentation()
    with closing(pdfium.PdfDocument(source)) as pdf:
        with closing(pdf[0]) as first:
            width, height = first.get_size()
        deck.slide_width = Pt(width)
        deck.slide_height = Pt(height)
        for page_index in range(len(pdf)):
            with closing(pdf[page_index]) as page:
                page_width, page_height = page.get_size()
                fit = min(width / page_width, height / page_height)
                image_width = page_width * fit
                image_height = page_height * fit
                stream = BytesIO()
                page_image(page).save(stream, format="PNG")
                stream.seek(0)
                slide = deck.slides.add_slide(deck.slide_layouts[6])
                slide.shapes.add_picture(stream, Pt((width-image_width)/2), Pt((height-image_height)/2),
                                         width=Pt(image_width), height=Pt(image_height))
    deck.save(target)


def pdf_to_word(source, target):
    from pdf2docx import Converter

    with closing(Converter(str(source))) as document:
        document.convert(str(target), parse_lattice_table=True, parse_stream_table=True)
    validate(target)


def upstream_to_word(source, target):
    script = {".hwp": "hwp2docx.py", ".hwpx": "hwpx_to_docx.py"}[extension(source)]
    exporter = Path(__file__).with_name("upstream") / script
    command = [sys.executable, "-X", "utf8", str(exporter), str(source), str(target)]
    try:
        result = subprocess.run(command, capture_output=True, timeout=120, check=False)
    except subprocess.TimeoutExpired as error:
        raise RuntimeError("HWP/HWPX to Word conversion timed out") from error
    if result.returncode:
        detail = (result.stderr or result.stdout).decode("utf-8", errors="replace").strip()
        raise RuntimeError(detail.splitlines()[-1] if detail else "HWP/HWPX to Word conversion failed")


def convert(source, target, pdf_objects=False):
    start, end = extension(source), extension(target)
    hwpx_layout = None
    if (start, end) == (".hwpx", ".docx"):
        from hwpx import convert as convert_hwpx

        hwpx_layout = convert_hwpx(source, target, export_native=partial(office, "hancom"))
    elif pdf_objects:
        from pdf_objects import reconstruct

        draft = target.with_name("positioned-legacy.docx")
        reconstruct(source, draft)
        office("word-modern", draft, target)
    elif start in HANCOM and end == ".docx":
        upstream_to_word(source, target)
    elif end == ".pdf":
        render(source, target, fit_table_images=True)
    elif start == ".pdf" and end in SLIDES:
        slides = target if end == ".pptx" else target.with_name("intermediate.pptx")
        pdf_to_slides(source, slides)
        if end == ".ppt":
            office("powerpoint", slides, target)
    elif start == ".pdf":
        if end in HANCOM:
            from pdf_objects import prefers_positioned_layout, try_positioned_hancom

            if prefers_positioned_layout(source):
                try:
                    positioned = try_positioned_hancom(source, target, office)
                except (RuntimeError, ValueError, OSError) as error:
                    print(f"Hancom positioned layout unavailable: {error}", file=sys.stderr)
                else:
                    if positioned:
                        validate(target)
                        return
        intermediate = target if end == ".docx" else target.with_name("intermediate.docx")
        pdf_to_word(source, intermediate)
        if end in HANCOM:
            from hancom_layout import fit

            fit(source, intermediate, target, office)
        elif end != ".docx":
            office("word", intermediate, target)
    elif start == ".docx" and end in HANCOM:
        from word_hancom import convert_word_to_hancom

        convert_word_to_hancom(source, target, office)
    else:
        office("hancom" if {start, end} & HANCOM else "word", source, target)
    validate(target)
    return hwpx_layout


def restore_hwpx_paper_box(page, render_size, paper_sizes):
    # Hancom's PDF exporter can crop a small strip from the declared sheet.
    # Restore the missing canvas without scaling or translating page content.
    candidates = [size for size in paper_sizes if all(
        -0.001 <= declared - rendered <= 2.001
        for declared, rendered in zip(size, render_size)
    )]
    if len(candidates) != 1 or candidates[0] == render_size:
        return
    width, height = candidates[0]
    if page.get_rotation() in (90, 270):
        width, height = height, width
    left, _, _, top = page.get_bbox()
    box = left, top - height, left + width, top
    page.set_mediabox(*box)
    page.set_cropbox(*box)


def compare(source_pdf, target_pdf, paper_sizes=(), measure_difference=False):
    from PIL import ImageChops
    import pypdfium2 as pdfium

    pages = []
    with closing(pdfium.PdfDocument(source_pdf)) as original, closing(pdfium.PdfDocument(target_pdf)) as result:
        for index in range(min(len(original), len(result))):
            with closing(original[index]) as left, closing(result[index]) as right:
                source_render_size = left.get_size()
                if paper_sizes:
                    restore_hwpx_paper_box(left, source_render_size, paper_sizes)
                source_size, output_size = left.get_size(), right.get_size()
                left_image, right_image = page_image(left), page_image(right)
                same_size = all(abs(a - b) <= 0.001 for a, b in zip(source_size, output_size))
                bounds = None
                difference = None
                if left_image.size == right_image.size:
                    difference = ImageChops.difference(left_image, right_image)
                    bounds = difference.getbbox()
                pages.append({
                    "page": index + 1,
                    "source_size_pt": source_size,
                    "output_size_pt": output_size,
                    "changed_pixel_bounds": bounds,
                    "match": same_size and left_image.size == right_image.size and bounds is None,
                })
                if source_render_size != source_size:
                    pages[-1]["source_render_size_pt"] = source_render_size
                if measure_difference:
                    pages[-1]["rgb_error"] = (sum((index % 256) * count for index, count in
                                                   enumerate(difference.histogram()))
                                                if difference is not None else None)
        return {
            "dpi": DPI,
            "source_pages": len(original),
            "output_pages": len(result),
            "visual_match": len(original) == len(result) and all(page["match"] for page in pages),
            "pages": pages,
        }


def refine_hwpx_layout(source, candidate, source_pdf, candidate_pdf, report, layout, separate_endnotes):
    from hwpx import verify_native_objects
    from hwpx_fidelity import position, refine

    if report["source_pages"] != report["output_pages"] or report["visual_match"]:
        return candidate, report
    before = [page["rgb_error"] for page in report["pages"]]
    if any(error is None for error in before):
        return candidate, report
    counts = layout.native_objects
    styled = candidate.with_name("styled.docx")
    if not refine(source, candidate, styled, source_pdf, candidate_pdf, counts, separate_endnotes):
        return candidate, report

    best_score = sum(before)
    for path in (styled, styled.with_name("positioned.docx")):
        preview = path.with_suffix(".pdf")
        if path != styled:
            position(styled, path, source_pdf, styled.with_suffix(".pdf"), counts, separate_endnotes)
        render(path, preview)
        result = compare(source_pdf, preview, layout.paper_sizes, measure_difference=True)
        if result["source_pages"] != result["output_pages"]:
            # The anchors are unreliable if styling changed the pagination.
            break
        errors = [page["rgb_error"] for page in result["pages"]]
        if all(error is not None for error in errors) and sum(errors) < best_score:
            verify_native_objects(path, counts)
            best_score = sum(errors)
            result["hwpx_rgb_error_before"] = before
            candidate, report = path, result
    return candidate, report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="?", type=Path)
    parser.add_argument("output", nargs="?", type=Path)
    parser.add_argument("--list", action="store_true")
    parser.add_argument("--report", type=Path)
    visual_modes = parser.add_mutually_exclusive_group()
    visual_modes.add_argument("--allow-differences", action="store_true",
                              help="Save a visual mismatch if the page count still matches.")
    visual_modes.add_argument("--require-visual-match", action="store_true",
                              help="Require pixel equality even for editable HWPX to DOCX output.")
    parser.add_argument("--allow-page-count-change", action="store_true",
                        help="Allow reflow to a different page count for editable document output.")
    parser.add_argument("--hwpx-structure", action="store_true",
                        help="Legacy HWPX to DOCX flag. Require an exact match unless differences are allowed.")
    pdf_modes = parser.add_mutually_exclusive_group()
    pdf_modes.add_argument("--pdf-reflow", action="store_true",
                           help="Reconstruct paragraphs and tables even for a designed PDF.")
    pdf_modes.add_argument("--pdf-objects", action="store_true",
                        help="Reconstruct PDF as positioned Word text boxes, pictures, and shapes.")
    args = parser.parse_args()
    if args.list:
        for start, end in sorted(PAIRS):
            print(f"{start[1:]} -> {end[1:]}")
        return 0
    if args.input is None or args.output is None:
        parser.error("Pass both input and output paths, or use --list.")
    source, target = args.input.resolve(), args.output.resolve()
    pair = extension(source), extension(target)
    if pair not in PAIRS:
        parser.error("Unsupported conversion pair. See --list.")
    is_hwpx_to_docx = pair == (".hwpx", ".docx")
    if args.allow_page_count_change and (
        not args.allow_differences or extension(target) not in HANCOM | WORD
    ):
        parser.error("--allow-page-count-change requires --allow-differences and editable document output.")
    if args.hwpx_structure and not is_hwpx_to_docx:
        parser.error("--hwpx-structure requires HWPX input and DOCX output.")
    is_pdf_to_docx = pair == (".pdf", ".docx")
    if (args.pdf_objects or args.pdf_reflow) and not is_pdf_to_docx:
        parser.error("PDF layout options require PDF input and DOCX output.")
    report_path = args.report.resolve() if args.report else None
    if target.exists() or (report_path and (report_path.exists() or report_path == target)):
        parser.error("Output and report paths must be distinct new files.")
    try:
        validate(source)
        pdf_objects = args.pdf_objects
        if is_pdf_to_docx and not (pdf_objects or args.pdf_reflow):
            from pdf_objects import prefers_positioned_layout

            pdf_objects = prefers_positioned_layout(source)
        with TemporaryDirectory(prefix="fileconvert-") as directory:
            scratch = Path(directory)
            candidate = scratch / ("result" + extension(target))
            hwpx_layout = convert(source, candidate, pdf_objects=pdf_objects)
            paper_sizes = hwpx_layout.paper_sizes if hwpx_layout else ()
            has_native_objects = hwpx_layout is not None and hwpx_layout.native_objects is not None
            source_pdf = scratch / "source.pdf"
            output_pdf = scratch / "output.pdf"
            render(source, source_pdf)
            render(candidate, output_pdf)
            report = compare(source_pdf, output_pdf, paper_sizes, measure_difference=has_native_objects)
            separate_endnotes = False
            if (hwpx_layout and hwpx_layout.endnotes_at_document_end
                    and report["output_pages"] < report["source_pages"]):
                from hwpx import separate_document_endnotes, verify_native_objects

                with_section = scratch / "endnotes-page.docx"
                with_section_pdf = scratch / "endnotes-page.pdf"
                separate_document_endnotes(candidate, with_section)
                render(with_section, with_section_pdf)
                adjusted = compare(source_pdf, with_section_pdf, paper_sizes, measure_difference=has_native_objects)
                if adjusted["output_pages"] == adjusted["source_pages"]:
                    verify_native_objects(with_section, hwpx_layout.native_objects)
                    candidate, output_pdf, report = with_section, with_section_pdf, adjusted
                    separate_endnotes = True
            if has_native_objects:
                try:
                    candidate, report = refine_hwpx_layout(
                        source, candidate, source_pdf, output_pdf, report, hwpx_layout, separate_endnotes,
                    )
                except (OSError, ValueError, RuntimeError, LookupError) as error:
                    report["hwpx_refinement_error"] = str(error)
                report["hwpx_native_objects"] = hwpx_layout.native_objects
                report["hwpx_endnotes_new_page"] = separate_endnotes
            report["hwpx_structure"] = bool(hwpx_layout and hwpx_layout.native_objects is None)
            report["pdf_objects"] = pdf_objects
            report["input_format"] = extension(source)
            report["output_format"] = extension(target)
            if report_path:
                with report_path.open("x", encoding="utf-8") as stream:
                    json.dump(report, stream, indent=2)
                    stream.write("\n")
            fixed_hancom_pagination = extension(source) == ".pdf" and extension(target) in HANCOM
            if report["source_pages"] != report["output_pages"] and (
                not args.allow_page_count_change or fixed_hancom_pagination
            ):
                print(
                    f"Page count changed from {report['source_pages']} to {report['output_pages']}. "
                    "Output was not saved.",
                    file=sys.stderr,
                )
                return 2
            editable_hwpx_default = is_hwpx_to_docx and not (args.require_visual_match or args.hwpx_structure)
            if not report["visual_match"] and not (editable_hwpx_default or args.allow_differences):
                print(f"Visual comparison failed at {DPI} DPI. Output was not saved.", file=sys.stderr)
                return 2
            with candidate.open("rb") as data, target.open("xb") as destination:
                shutil.copyfileobj(data, destination)
            print(f"{target} (visual match at {DPI} DPI: {report['visual_match']})")
    except Exception as error:
        print(f"Conversion failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
