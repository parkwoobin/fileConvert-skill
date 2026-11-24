"""Preserve Word's saved page starts when importing editable documents into Hancom."""
from shutil import copyfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pymupdf


def normalized(text):
    return "".join(char for char in text if char.isalnum()).casefold()


def leading_cached_break(paragraph):
    if not paragraph.text.strip() or paragraph.paragraph_format.page_break_before is not None:
        return False
    for node in paragraph._p.iter():
        if node.tag == qn("w:lastRenderedPageBreak"):
            return True
        if node.tag == qn("w:t") and (node.text or "").strip():
            return False
        if node.tag in {qn("w:drawing"), qn("w:pict"), qn("w:br"), qn("w:object")}:
            return False
    return False


def compact_spacers(document):
    for paragraph in document.paragraphs:
        if paragraph.text.strip() or paragraph._p.xpath(
            ".//wp:inline | .//w:sectPr | .//w:br | .//w:fldChar | .//w:fldSimple | .//w:object | .//w:pict"
        ):
            continue
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = Pt(1)
        props = paragraph._p.get_or_add_pPr()
        rpr = props.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            props.insert_element_before(rpr, "w:sectPr", "w:pPrChange")
        size = rpr.find(qn("w:sz"))
        if size is None:
            size = OxmlElement("w:sz")
            rpr.insert_element_before(size, "w:szCs")
        size.set(qn("w:val"), "2")


def scan(preview):
    with pymupdf.open(preview) as pdf:
        texts = []
        blanks = 0
        for page in pdf:
            text = normalized(page.get_text())
            texts.append(text)
            if not text and not page.get_image_info() and not page.get_drawings():
                blanks += 1
    return texts, blanks


def find_page(texts, seed):
    pages = [index for index, body in enumerate(texts) if seed in body]
    return pages[0] if len(pages) == 1 else None


def preserve_page_starts(source, target, office):
    document = Document(source)
    candidates = [(index, normalized(paragraph.text)[:32])
                  for index, paragraph in enumerate(document.paragraphs)
                  if leading_cached_break(paragraph)]
    if not candidates:
        office("hancom", source, target)
        return

    workspace = target.parent
    source_preview = workspace / "word-source-layout.pdf"
    office("word", source, source_preview)
    source_pages, _ = scan(source_preview)
    anchors = []
    for paragraph, seed in candidates:
        if len(seed) < 3:
            continue
        page = find_page(source_pages, seed)
        if page is not None and page > 0 and source_pages[page].find(seed) < 80:
            anchors.append((paragraph, seed, page))
    if not anchors:
        office("hancom", source, target)
        return

    best = None
    for name, compact in (("forced-starts", False), ("compact", True)):
        draft = workspace / f"{name}.docx"
        candidate = workspace / f"{name}{target.suffix}"
        preview = workspace / f"{name}.pdf"
        document = Document(source)
        for index, _, _ in anchors:
            document.paragraphs[index].paragraph_format.page_break_before = True
        if compact:
            compact_spacers(document)
        document.save(draft)
        office("hancom", draft, candidate)
        office("hancom", candidate, preview)
        pages, blanks = scan(preview)
        moved = sum(find_page(pages, seed) != expected for _, seed, expected in anchors)
        rank = (moved, abs(len(pages) - len(source_pages)), blanks)
        if best is None or rank < best[0]:
            best = rank, candidate
        if rank[:2] == (0, 0) or len(pages) <= len(source_pages) or not blanks:
            break
    copyfile(best[1], target)


def convert_word_to_hancom(source, target, office):
    from word_borders import read_word_borders, repair_borders

    specs = read_word_borders(source)
    if not any(table and any(table[2].values()) for table in specs):
        preserve_page_starts(source, target, office)
        return

    imported = target.with_name("word-import.hwpx")
    preserve_page_starts(source, imported, office)
    repaired = target.with_name("word-table-borders.hwpx")
    candidate = repaired if repair_borders(specs, imported, repaired) else imported
    if target.suffix.lower() == ".hwp":
        office("hancom", candidate, target)
    else:
        copyfile(candidate, target)
