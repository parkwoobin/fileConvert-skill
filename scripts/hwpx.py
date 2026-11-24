"""Map HWPX paragraphs, character styles and tables to editable Word elements."""
from io import BytesIO
from pathlib import PurePosixPath
import re
from zipfile import ZipFile

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from lxml import etree

SECTION = re.compile(r"Contents/section(\d+)\.xml")
PART_LIMIT = 24 * 1024 * 1024


def kind(element):
    return etree.QName(element).localname if isinstance(element.tag, str) else ""


def elements(parent, name):
    return (child for child in parent if kind(child) == name)


def first(parent, name):
    return next(elements(parent, name), None)


def part(package, name):
    info = package.getinfo(name)
    if info.file_size > PART_LIMIT:
        raise ValueError(f"HWPX part exceeds 24 MiB: {name}")
    return package.read(info)


def xml(package, name):
    parser = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False)
    root = etree.fromstring(part(package, name), parser)
    if root.getroottree().docinfo.doctype:
        raise ValueError("HWPX XML must not contain a DTD")
    return root


class Converter:
    def __init__(self, package):
        self.package = package
        self.document = Document()
        header = xml(package, "Contents/header.xml")
        self.character = {node.get("id"): node for node in header.iter() if kind(node) == "charPr"}
        self.paragraph = {node.get("id"): node for node in header.iter() if kind(node) == "paraPr"}
        self.fonts = {
            group.get("lang"): {font.get("id"): font.get("face") for font in elements(group, "font")}
            for group in header.iter() if kind(group) == "fontface"
        }
        manifest = xml(package, "Contents/content.hpf")
        self.images = {
            node.get("id"): node.get("href") for node in manifest.iter()
            if kind(node) == "item" and (node.get("href") or "").startswith("BinData/")
        }

    def format_run(self, run, style):
        if style is None:
            return
        height = int(style.get("height", "0"))
        if height > 0:
            run.font.size = Pt(height / 100)
        tracking = first(style, "spacing")
        if tracking is not None and height > 0:
            percentages = {int(tracking.get(key, "0")) for key in (
                "hangul", "latin", "hanja", "japanese", "other", "symbol", "user",
            )}
            if len(percentages) == 1:
                # Word uses one amount per run, in twentieths of a point.
                twips = round(height * percentages.pop() / 500)
                if twips:
                    spacing = OxmlElement("w:spacing")
                    spacing.set(qn("w:val"), str(twips))
                    run._element.get_or_add_rPr().insert_element_before(spacing, "w:sz", "w:szCs")
        color = style.get("textColor", "")
        if re.fullmatch(r"#[0-9A-Fa-f]{6}", color):
            run.font.color.rgb = RGBColor.from_string(color[1:])
        run.bold = first(style, "bold") is not None
        run.italic = first(style, "italic") is not None
        underline = first(style, "underline")
        run.underline = underline is not None and underline.get("type") not in {None, "NONE"}
        refs = first(style, "fontRef")
        if refs is not None:
            latin = self.fonts.get("LATIN", {}).get(refs.get("latin"))
            korean = self.fonts.get("HANGUL", {}).get(refs.get("hangul"))
            if latin:
                run.font.name = latin
            if korean:
                run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), korean)

    def add_picture(self, run, element):
        ref = next((node.get("binaryItemIDRef") for node in element.iter()
                    if node.get("binaryItemIDRef")), None)
        member = self.images.get(ref)
        if not member or PurePosixPath(member).suffix.lower() not in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}:
            run.add_text("[그림]")
            return
        size = first(element, "curSz")
        width = int(size.get("width", "0")) if size is not None else 0
        run.add_picture(BytesIO(part(self.package, member)), width=Pt(width / 100) if width else None)

    def add_paragraph(self, host, source, paragraph=None):
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        paragraph = host.add_paragraph() if paragraph is None else paragraph
        if source.get("pageBreak") == "1":
            paragraph.paragraph_format.page_break_before = True
        style = self.paragraph.get(source.get("paraPrIDRef"))
        alignment = first(style, "align") if style is not None else None
        if alignment is not None:
            paragraph.alignment = {
                "LEFT": WD_ALIGN_PARAGRAPH.LEFT, "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
                "CENTER": WD_ALIGN_PARAGRAPH.CENTER, "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
            }.get(alignment.get("horizontal"))
        line_spacing = first(style, "lineSpacing") if style is not None else None
        if line_spacing is not None and line_spacing.get("type") == "PERCENT":
            value = int(line_spacing.get("value", "0"))
            if value > 0:
                paragraph.paragraph_format.line_spacing = value / 100
        tables = []
        for source_run in elements(source, "run"):
            run = paragraph.add_run()
            self.format_run(run, self.character.get(source_run.get("charPrIDRef")))
            for item in source_run:
                tag = kind(item)
                if tag == "t":
                    run.add_text(item.text or "")
                    for marker in item:
                        if kind(marker) == "tab":
                            run.add_tab()
                        elif kind(marker) == "lineBreak":
                            run.add_break()
                        run.add_text(marker.tail or "")
                elif tag == "pic":
                    self.add_picture(run, item)
                elif tag == "tbl":
                    tables.append(item)
                elif tag == "equation":
                    run.add_text("[수식]")
        for table in tables:
            self.add_table(host, table)

    def add_table(self, host, source):
        rows, cols = int(source.get("rowCnt", "0")), int(source.get("colCnt", "0"))
        if rows < 1 or cols < 1 or rows * cols > 3000:
            raise ValueError("HWPX table dimensions are invalid or too large")
        table = host.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"
        for row in elements(source, "tr"):
            for entry in elements(row, "tc"):
                addr, span = first(entry, "cellAddr"), first(entry, "cellSpan")
                if addr is None:
                    raise ValueError("HWPX table cell has no address")
                r, c = int(addr.get("rowAddr")), int(addr.get("colAddr"))
                rs = int(span.get("rowSpan", "1")) if span is not None else 1
                cs = int(span.get("colSpan", "1")) if span is not None else 1
                if min(r, c) < 0 or min(rs, cs) < 1 or r + rs > rows or c + cs > cols:
                    raise ValueError("HWPX table cell is outside the table")
                cell = table.cell(r, c)
                if rs > 1 or cs > 1:
                    cell = cell.merge(table.cell(r + rs - 1, c + cs - 1))
                content = first(entry, "subList")
                if content is not None:
                    for index, paragraph in enumerate(elements(content, "p")):
                        self.add_paragraph(cell, paragraph, cell.paragraphs[0] if index == 0 else None)

    def configure_page(self, section, root):
        page = next((node for node in root.iter() if kind(node) == "pagePr"), None)
        if page is None:
            return
        for key in ("width", "height"):
            value = int(page.get(key, "0"))
            if value > 0:
                setattr(section, "page_" + key, Pt(value / 100))
        margins = first(page, "margin")
        if margins is None:
            return
        for key in ("top", "bottom", "left", "right"):
            if margins.get(key) is not None:
                setattr(section, key + "_margin", Pt(int(margins.get(key)) / 100))
        for key in ("header", "footer"):
            if margins.get(key) is not None:
                setattr(section, key + "_distance", Pt(int(margins.get(key)) / 100))

    def save(self, target):
        sections = sorted(
            (int(match.group(1)), member) for member in self.package.namelist()
            if (match := SECTION.fullmatch(member))
        )
        if not sections:
            raise ValueError("HWPX does not contain a body section")
        for index, (_, name) in enumerate(sections):
            root = xml(self.package, name)
            section = self.document.add_section(WD_SECTION_START.NEW_PAGE) if index else self.document.sections[0]
            self.configure_page(section, root)
            for paragraph in elements(root, "p"):
                self.add_paragraph(self.document, paragraph)
        self.document.save(target)


def convert(source, target):
    with ZipFile(source) as package:
        Converter(package).save(target)
