"""Refine native HWPX exports using source formatting and rendered anchors."""
from copy import deepcopy
from zipfile import ZipFile

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt
from lxml import etree
import pymupdf

from hwpx import SECTION, elements, first, kind, xml


def words(value):
    return "".join(value.split())


def rendered_lines(page):
    lines = [line["spans"] for block in page.get_text("rawdict")["blocks"] if block["type"] == 0
             for line in block["lines"] if line["spans"]]
    for spans in lines:
        for span in spans:
            span["text"] = "".join(character["c"] for character in span["chars"])
    return lines


def main_span(spans):
    return max(spans, key=lambda span: span["size"])


def baseline(spans):
    return main_span(spans)["origin"][1]


def text_line(page, text):
    matches = []
    for spans in rendered_lines(page):
        size = main_span(spans)["size"]
        if words("".join(span["text"] for span in spans if size - span["size"] < 0.4)) == text:
            matches.append(spans)
    return matches[0] if len(matches) == 1 else None


def raised_spans(spans):
    size, y = main_span(spans)["size"], baseline(spans)
    return [span for span in spans if span["text"].strip() and span["size"] < size - 0.4
            and span["origin"][1] < y - 0.5]


def set_run_value(properties, name, value):
    tag = "w:" + name
    element = properties.find(qn(tag))
    if element is None:
        element = OxmlElement(tag)
        trailing = ("w:szCs", "w:highlight", "w:vertAlign")
        if name == "position":
            trailing = ("w:sz", *trailing)
        properties.insert_element_before(element, *trailing)
    element.set(qn("w:val"), str(value))


def format_reference(run, span, line, fonts=None):
    properties = run.get_or_add_rPr()
    vertical = properties.find(qn("w:vertAlign"))
    if vertical is not None:
        properties.remove(vertical)
    set_run_value(properties, "position", round((baseline(line) - span["origin"][1]) * 2))
    set_run_value(properties, "sz", round(span["size"] * 2))
    if fonts is not None:
        existing = properties.find(qn("w:rFonts"))
        if existing is not None:
            properties.remove(existing)
        properties.insert_element_before(deepcopy(fonts), "w:position", "w:sz")


def body_matches(document, source_page, word_page):
    paragraphs = document.paragraphs[:2]
    if len(paragraphs) != 2:
        return ()
    pairs = []
    for target in paragraphs:
        key = words(target.text)
        if len(key) < 6:
            return ()
        left, right = text_line(source_page, key), text_line(word_page, key)
        if left is None or right is None:
            return ()
        pairs.append((left, right))
    return pairs


def position_body(document, pairs):
    if not pairs:
        return
    (left, right), (next_left, next_right) = pairs
    top_delta = baseline(left) - baseline(right)
    gap_delta = baseline(next_left) - baseline(next_right) - top_delta
    section = document.sections[0]
    if abs(top_delta) <= 12 and section.top_margin.pt + top_delta > section.header_distance.pt:
        section.top_margin = Pt(section.top_margin.pt + top_delta)
        first_paragraph = document.paragraphs[0]
        spacing = first_paragraph.paragraph_format.space_after
        if spacing is None:
            spacing = first_paragraph.style.paragraph_format.space_after
        if spacing is not None and abs(gap_delta) <= 12 and 0 <= spacing.pt + gap_delta <= 36:
            first_paragraph.paragraph_format.space_after = Pt(spacing.pt + gap_delta)


def style_body(document, pairs):
    for paragraph, (original, rendered) in zip(document.paragraphs, pairs):
        for run in paragraph.runs:
            properties = run._r.rPr
            tracking = properties.find(qn("w:spacing")) if properties is not None else None
            if tracking is None or len(run.text.strip()) < 6:
                continue
            source_spans = [span for span in original if span["text"].strip() == run.text.strip()]
            word_spans = [span for span in rendered if span["text"].strip() == run.text.strip()]
            if len(source_spans) != 1 or len(word_spans) != 1:
                continue
            left, right = source_spans[0], word_spans[0]
            if left["font"] != right["font"]:
                continue
            advances = []
            for span in (left, right):
                characters = [item for item in span["chars"] if item["c"].strip()]
                advances.append(characters[-1]["origin"][0] - characters[0]["origin"][0])
            delta = round((advances[0] - advances[1]) * 20 / (len(run.text.strip()) - 1))
            if 0 < abs(delta) <= 5:
                tracking.set(qn("w:val"), str(int(tracking.get(qn("w:val"))) + delta))
        tags = {qn("w:footnoteReference"), qn("w:endnoteReference")}
        references = [run for run in paragraph._p.iter(qn("w:r")) if any(child.tag in tags for child in run)]
        markers = raised_spans(original)
        if len(markers) != len(references):
            continue
        for run, marker in zip(references, markers):
            neighbor = run.getprevious()
            fonts = neighbor.find(qn("w:rPr") + "/" + qn("w:rFonts")) if neighbor is not None else None
            if marker["font"] != main_span(original)["font"]:
                fonts = None
            format_reference(run, marker, original, fonts)


def align_equations(document, source_paragraphs, styles):
    equations = [(paragraph, equation) for paragraph in source_paragraphs
                 for equation in paragraph.iter() if kind(equation) == "equation"]
    output = list(document._element.body.iter(qn("m:oMath")))
    if len(equations) != len(output):
        return
    for (paragraph, equation), math in zip(equations, output):
        position = first(equation, "pos")
        if position is None or position.get("treatAsChar") != "1":
            continue
        size = int(equation.get("baseUnit", "0"))
        if 100 <= size <= 9600:
            for run_properties in math.iter(qn("w:rPr")):
                set_run_value(run_properties, "sz", round(size / 50))
        parent = math.getparent()
        if parent.tag == qn("w:p") and all(child.tag in {qn("w:pPr"), qn("m:oMath")} for child in parent):
            style = styles.get(paragraph.get("paraPrIDRef"))
            alignment = first(style, "align") if style is not None else None
            value = {"LEFT": "left", "CENTER": "center", "RIGHT": "right"}.get(
                alignment.get("horizontal") if alignment is not None else None
            )
            if value is None:
                continue
            display, properties, justification = (OxmlElement(tag) for tag in ("m:oMathPara", "m:oMathParaPr", "m:jc"))
            justification.set(qn("m:val"), value)
            properties.append(justification)
            display.append(properties)
            parent.replace(math, display)
            display.append(math)


def note_line(pdf, last_page=False):
    pages = [pdf[-1]] if last_page else pdf
    matches = []
    for page in pages:
        for line in rendered_lines(page):
            raised = raised_spans(line)
            in_position = (baseline(line) < page.rect.height * 0.45 if last_page
                           else baseline(line) > page.rect.height * 0.6)
            if raised and line[0] is raised[0] and in_position:
                matches.append(line)
    return matches[0] if len(matches) == 1 else None


def note_part(document, note):
    part = next((part for part in document.part.related_parts.values()
                 if str(part.partname) == f"/word/{note}s.xml"), None)
    if part is None:
        raise ValueError(f"Word export has no {note} part")
    return part


def remove_endnote_separator(document, part, root):
    separator = next((entry for entry in root if entry.get(qn("w:type")) == "separator"), None)
    if separator is None:
        separator = OxmlElement("w:endnote")
        identifier = min(int(entry.get(qn("w:id"))) for entry in root) - 1
        separator.set(qn("w:type"), "separator")
        separator.set(qn("w:id"), str(identifier))
        root.insert(0, separator)
        settings = document.settings.element
        properties = settings.find(qn("w:endnotePr"))
        if properties is None:
            properties = OxmlElement("w:endnotePr")
            settings.insert_element_before(properties, "w:compat", "w:docVars", "w:rsids", "m:mathPr")
        reference = OxmlElement("w:endnote")
        reference.set(qn("w:id"), str(identifier))
        properties.append(reference)
    paragraph = parse_xml(
        f'<w:p {nsdecls("w")}><w:pPr>'
        '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
    )
    separator[:] = [paragraph]
    part._blob = serialize(root)
    # The new section exists only for endnotes. Give Word a minimal body line
    # so inheriting the main paragraph style does not push the notes down.
    leading = document.add_paragraph()
    leading.paragraph_format.space_before = Pt(0)
    leading.paragraph_format.space_after = Pt(0)
    leading.paragraph_format.line_spacing = Pt(1)
    properties = parse_xml(f'<w:rPr {nsdecls("w")}><w:sz w:val="2"/></w:rPr>')
    leading._p.get_or_add_pPr().append(properties)


def serialize(root):
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def style_notes(document, source_pdf, counts, separate_endnotes):
    for note in ("footnote", "endnote"):
        is_endnote = note == "endnote"
        if counts[note + "s"] != 1 or (is_endnote and not separate_endnotes):
            continue
        line = note_line(source_pdf, last_page=is_endnote)
        if line is None:
            continue
        part = note_part(document, note)
        root = parse_xml(part.blob)
        for reference in root.iter(qn(f"w:{note}Ref")):
            format_reference(reference.getparent(), raised_spans(line)[0], line)
        if is_endnote and not any(
            0 < baseline(line) - drawing["rect"].y0 <= 36 and drawing["rect"].width >= 20
            and drawing["rect"].height <= 2 for drawing in source_pdf[-1].get_drawings()
        ):
            remove_endnote_separator(document, part, root)
        else:
            part._blob = serialize(root)


def refine(source, candidate, target, source_pdf, word_pdf, counts, separate_endnotes):
    with ZipFile(source) as package:
        sections = [name for name in package.namelist() if SECTION.fullmatch(name)]
        if len(sections) != 1:
            return False
        body = xml(package, sections[0])
        header = xml(package, "Contents/header.xml")
    paragraphs = list(elements(body, "p"))
    styles = {node.get("id"): node for node in header.iter() if kind(node) == "paraPr"}
    document = Document(candidate)
    with pymupdf.open(source_pdf) as left, pymupdf.open(word_pdf) as right:
        style_body(document, body_matches(document, left[0], right[0]))
        align_equations(document, paragraphs, styles)
        style_notes(document, left, counts, separate_endnotes)
    document.save(target)
    return True


def position_notes(document, original, rendered, counts, separate_endnotes):
    for note in ("footnote", "endnote"):
        is_endnote = note == "endnote"
        if counts[note + "s"] != 1 or (is_endnote and not separate_endnotes):
            continue
        left, right = note_line(original, is_endnote), note_line(rendered, is_endnote)
        if left is None or right is None:
            continue
        delta = baseline(left) - baseline(right)
        if is_endnote:
            section = document.sections[-1]
            desired = max(section.header_distance.pt + 1, section.top_margin.pt + delta)
            if abs(delta) <= 45:
                section.top_margin = Pt(desired)
            continue
        part = note_part(document, note)
        root = parse_xml(part.blob)
        marker = next(root.iter(qn("w:footnoteRef")), None)
        if marker is None:
            continue
        paragraph = next(marker.iterancestors(qn("w:p")))
        properties = paragraph.get_or_add_pPr()
        spacing = properties.get_or_add_spacing()
        after = spacing.get(qn("w:after"))
        if after is None and properties.pStyle is not None:
            style = next((style for style in document.styles if style.style_id == properties.pStyle.val), None)
            inherited = style.paragraph_format.space_after if style is not None else None
            after = inherited.twips if inherited is not None else None
        if after is not None and 0 <= int(after) - round(delta * 20) <= 400 and abs(delta) <= 10:
            spacing.set(qn("w:after"), str(int(after) - round(delta * 20)))
            part._blob = serialize(root)


def position(candidate, target, source_pdf, word_pdf, counts, separate_endnotes):
    document = Document(candidate)
    with pymupdf.open(source_pdf) as left, pymupdf.open(word_pdf) as right:
        position_body(document, body_matches(document, left[0], right[0]))
        position_notes(document, left, right, counts, separate_endnotes)
    document.save(target)
