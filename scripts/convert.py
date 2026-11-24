"""FileConvert local CLI."""
import argparse
from contextlib import closing
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
}
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


def office(engine, source, target):
    if sys.platform != "win32":
        raise RuntimeError(f"{engine} conversion requires a Windows desktop and the installed application")
    command = [
        sys.executable, str(Path(__file__).with_name("office.py")),
        engine, str(source), str(target),
        os.environ.get("FILECONVERT_HANCOM_PATH_MODULE", ""),
    ]
    try:
        result = subprocess.run(command, capture_output=True, timeout=120, check=False)
    except subprocess.TimeoutExpired as error:
        raise RuntimeError(f"{engine} timed out. Check the application for a file access or licensing prompt.") from error
    if result.returncode:
        diagnostic = result.stderr.decode(errors="replace").strip().splitlines()
        raise RuntimeError(f"{engine}: {diagnostic[-1] if diagnostic else 'conversion failed'}")
    validate(target)


def render(source, target):
    ext = extension(source)
    if ext == ".pdf":
        shutil.copyfile(source, target)
        return
    engine = "hancom" if ext in HANCOM else "word" if ext in WORD else "powerpoint"
    office(engine, source, target)


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
        document.convert(str(target))
    validate(target)


def convert(source, target, hwpx_structure=False):
    start, end = extension(source), extension(target)
    if hwpx_structure:
        from hwpx import convert as convert_structure

        convert_structure(source, target)
    elif end == ".pdf":
        render(source, target)
    elif start == ".pdf" and end in SLIDES:
        slides = target if end == ".pptx" else target.with_name("intermediate.pptx")
        pdf_to_slides(source, slides)
        if end == ".ppt":
            office("powerpoint", slides, target)
    elif start == ".pdf":
        intermediate = target if end == ".docx" else target.with_name("intermediate.docx")
        pdf_to_word(source, intermediate)
        if end != ".docx":
            office("hancom" if end in HANCOM else "word", intermediate, target)
    else:
        office("hancom" if {start, end} & HANCOM else "word", source, target)
    validate(target)


def compare(source_pdf, target_pdf):
    from PIL import ImageChops
    import pypdfium2 as pdfium

    pages = []
    with closing(pdfium.PdfDocument(source_pdf)) as original, closing(pdfium.PdfDocument(target_pdf)) as result:
        for index in range(min(len(original), len(result))):
            with closing(original[index]) as left, closing(result[index]) as right:
                left_image, right_image = page_image(left), page_image(right)
                same_size = all(abs(a - b) <= 0.001 for a, b in zip(left.get_size(), right.get_size()))
                bounds = None
                if left_image.size == right_image.size:
                    bounds = ImageChops.difference(left_image, right_image).getbbox()
                pages.append({
                    "page": index + 1,
                    "source_size_pt": left.get_size(),
                    "output_size_pt": right.get_size(),
                    "changed_pixel_bounds": bounds,
                    "match": same_size and left_image.size == right_image.size and bounds is None,
                })
        return {
            "dpi": DPI,
            "source_pages": len(original),
            "output_pages": len(result),
            "visual_match": len(original) == len(result) and all(page["match"] for page in pages),
            "pages": pages,
        }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="?", type=Path)
    parser.add_argument("output", nargs="?", type=Path)
    parser.add_argument("--list", action="store_true")
    parser.add_argument("--report", type=Path)
    parser.add_argument("--allow-differences", action="store_true")
    parser.add_argument("--hwpx-structure", action="store_true",
                        help="Use the direct HWPX XML importer for HWPX to DOCX.")
    args = parser.parse_args()
    if args.list:
        for start, end in sorted(PAIRS):
            print(f"{start[1:]} -> {end[1:]}")
        return 0
    if args.input is None or args.output is None:
        parser.error("Pass both input and output paths, or use --list.")
    source, target = args.input.resolve(), args.output.resolve()
    if (extension(source), extension(target)) not in PAIRS:
        parser.error("Unsupported conversion pair. See --list.")
    if args.hwpx_structure and (extension(source), extension(target)) != (".hwpx", ".docx"):
        parser.error("--hwpx-structure requires HWPX input and DOCX output.")
    report_path = args.report.resolve() if args.report else None
    if target.exists() or (report_path and (report_path.exists() or report_path == target)):
        parser.error("Output and report paths must be distinct new files.")
    try:
        validate(source)
        with TemporaryDirectory(prefix="fileconvert-") as directory:
            scratch = Path(directory)
            candidate = scratch / ("result" + extension(target))
            convert(source, candidate, args.hwpx_structure)
            source_pdf = scratch / "source.pdf"
            output_pdf = scratch / "output.pdf"
            render(source, source_pdf)
            render(candidate, output_pdf)
            report = compare(source_pdf, output_pdf)
            report["hwpx_structure"] = args.hwpx_structure
            report["input_format"] = extension(source)
            report["output_format"] = extension(target)
            if report_path:
                with report_path.open("x", encoding="utf-8") as stream:
                    json.dump(report, stream, indent=2)
                    stream.write("\n")
            if not report["visual_match"] and not args.allow_differences:
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
