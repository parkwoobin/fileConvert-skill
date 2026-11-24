"""Refine native HWPX exports using source formatting and rendered anchors."""
from copy import deepcopy
import re
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
        restore_math_italics(math, equation)
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


def restore_math_italics(math, equation):
    if equation.get("font", "").casefold() != "hancomeqn":
        return
    script = first(equation, "script")
    expression = (script.text or "").replace("{", "").replace("}", "") if script is not None else ""
    if not re.fullmatch(r"[-+*/=<>a-zA-Z0-9(). \t]+", expression):
        return
    variables = re.findall(r"[a-zA-Z]+", expression)
    if not variables or any(len(name) != 1 for name in variables):
        return
    # HancomEQN uses private glyphs and lacks OpenType MATH. Preserve the Word
    # math font and restore Hancom's italic style for ordinary variables.
    for run in math.iter(qn("m:r")):
        if run.findtext(qn("m:t")) not in variables:
            continue
        properties = run.find(qn("m:rPr"))
        if properties is None:
            properties = OxmlElement("m:rPr")
            run.insert(0, properties)
        if properties.find(qn("m:nor")) is not None:
            continue
        style = properties.find(qn("m:sty"))
        if style is None:
            style = OxmlElement("m:sty")
            properties.insert_element_before(style, "m:brk", "m:aln")
        if style.get(qn("m:val")) in (None, "p"):
            style.set(qn("m:val"), "i")


def note_line(pdf, last_page=False):
    pages = [pdf[-1]] if last_page else pdf
    matches = []
    for page in pages:
        for line in rendered_lines(page):
            raised = raised_spans(line)
            in_position = (baseline(line) < page.rect.height * 0.45 if last_page
                           else baseline(line) > page.rect.height * 0.6)
            if raised and line[0] is raised[0] and in_position:
                matches.append((page, line))
    return matches[0] if len(matches) == 1 else None


def note_part(document, note):
    part = next((part for part in document.part.related_parts.values()
                 if str(part.partname) == f"/word/{note}s.xml"), None)
    if part is None:
        raise ValueError(f"Word export has no {note} part")
    return part


def set_note_separator(document, root, note, paragraph):
    separator = next((entry for entry in root if entry.get(qn("w:type")) == "separator"), None)
    if separator is None:
        separator = OxmlElement(f"w:{note}")
        identifier = min(int(entry.get(qn("w:id"))) for entry in root) - 1
        separator.set(qn("w:type"), "separator")
        separator.set(qn("w:id"), str(identifier))
        root.insert(0, separator)
        settings = document.settings.element
        properties = settings.find(qn(f"w:{note}Pr"))
        if properties is None:
            properties = OxmlElement(f"w:{note}Pr")
            successors = ("w:endnotePr",) if note == "footnote" else ()
            settings.insert_element_before(properties, *successors, "w:compat", "w:docVars", "w:rsids", "m:mathPr")
        reference = OxmlElement(f"w:{note}")
        reference.set(qn("w:id"), str(identifier))
        properties.append(reference)
    separator[:] = [paragraph]


def compact_endnotes(document, root):
    spacer = parse_xml(
        f'<w:p {nsdecls("w")}><w:pPr>'
        '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
    )
    set_note_separator(document, root, "endnote", spacer)
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


def note_rules(page, line, near_start=True):
    note_x = line[0]["origin"][0]
    line_y = baseline(line)
    return [drawing for drawing in page.get_drawings()
            if drawing["rect"].width >= 20 and drawing["rect"].height <= 2
            and (not near_start or abs(drawing["rect"].x0 - note_x) <= 12)
            and 0 < line_y - drawing["rect"].y0 <= 36]


def style_footnote_rule(document, root, drawing, line):
    stroke = drawing.get("color")
    width = drawing.get("width")
    if drawing["type"] != "s" or not stroke or len(stroke) != 3 or not width or not 0.25 <= width <= 2:
        return
    if len(drawing["items"]) != 1 or drawing["items"][0][0] != "l":
        return
    if drawing.get("dashes") not in (None, "[] 0"):
        return
    rect = drawing["rect"]
    section = document.sections[0]
    left = round((rect.x0 - section.left_margin.pt) * 20)
    right = round((section.page_width.pt - section.right_margin.pt - rect.x1) * 20)
    color = "".join(f"{round(value * 255):02X}" for value in stroke[:3])
    size = round(width * 8)
    separator = parse_xml(
        f'<w:p {nsdecls("w")}><w:pPr><w:pBdr>'
        f'<w:bottom w:val="single" w:sz="{size}" w:space="0" w:color="{color}"/>'
        '</w:pBdr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        f'<w:ind w:left="{left}" w:right="{right}"/>'
        '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
    )
    set_note_separator(document, root, "footnote", separator)
    marker = next(root.iter(qn("w:footnoteRef")), None)
    if marker is None:
        return
    paragraph = next(marker.iterancestors(qn("w:p")))
    if len(paragraph.getparent()) == 1 and not list(paragraph.iter(qn("w:br"))):
        spacing = paragraph.get_or_add_pPr().get_or_add_spacing()
        spacing.set(qn("w:lineRule"), "exact")
        spacing.set(qn("w:line"), str(round(main_span(line)["size"] * 1.5 * 20)))


def style_notes(document, source_pdf, counts, separate_endnotes):
    for note in ("footnote", "endnote"):
        is_endnote = note == "endnote"
        if counts[note + "s"] != 1 or (is_endnote and not separate_endnotes):
            continue
        anchor = note_line(source_pdf, last_page=is_endnote)
        if anchor is None:
            continue
        page, line = anchor
        part = note_part(document, note)
        root = parse_xml(part.blob)
        for reference in root.iter(qn(f"w:{note}Ref")):
            format_reference(reference.getparent(), raised_spans(line)[0], line)
        rules = note_rules(page, line, near_start=not is_endnote)
        if is_endnote and not rules:
            compact_endnotes(document, root)
        elif not is_endnote and len(rules) == 1:
            style_footnote_rule(document, root, rules[0], line)
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


def position_footnote_rule(root, spacing, source, rendered):
    separator = next((entry for entry in root if entry.get(qn("w:type")) == "separator"), None)
    if separator is None or separator.find(".//" + qn("w:bottom")) is None:
        return 0
    left = note_rules(*source)
    right = note_rules(*rendered)
    if len(left) != 1 or len(right) != 1:
        return 0
    original, current = left[0]["rect"], right[0]["rect"]
    indent = separator.find(".//" + qn("w:ind"))
    if indent is not None:
        for key, delta in (("left", original.x0 - current.x0), ("right", current.x1 - original.x1)):
            if abs(delta) <= 12:
                attribute = qn("w:" + key)
                indent.set(attribute, str(int(indent.get(attribute, "0")) + round(delta * 20)))
    if spacing.get(qn("w:lineRule")) != "exact":
        return 0
    original_gap = baseline(source[1]) - (original.y0 + original.y1) / 2
    current_gap = baseline(rendered[1]) - (current.y0 + current.y1) / 2
    line_adjustment = original_gap - current_gap
    if abs(line_adjustment) > 4:
        return 0
    line = int(spacing.get(qn("w:line"))) + round(line_adjustment * 20)
    if line / 20 < main_span(source[1])["size"]:
        return 0
    spacing.set(qn("w:line"), str(line))
    # Exact line height also moves the text baseline a fraction of that change.
    # Compensate in the note's trailing space without moving the whole section.
    return round(-line_adjustment * 4)


def position_notes(document, original, rendered, counts, separate_endnotes):
    for note in ("footnote", "endnote"):
        is_endnote = note == "endnote"
        if counts[note + "s"] != 1 or (is_endnote and not separate_endnotes):
            continue
        source, target = note_line(original, is_endnote), note_line(rendered, is_endnote)
        if source is None or target is None:
            continue
        delta = baseline(source[1]) - baseline(target[1])
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
        baseline_shift = position_footnote_rule(root, spacing, source, target)
        if after is not None and abs(delta) <= 10:
            trailing_space = int(after) - round(delta * 20) + baseline_shift
            if 0 <= trailing_space <= 400:
                spacing.set(qn("w:after"), str(trailing_space))
        part._blob = serialize(root)


def position(candidate, target, source_pdf, word_pdf, counts, separate_endnotes):
    document = Document(candidate)
    with pymupdf.open(source_pdf) as left, pymupdf.open(word_pdf) as right:
        position_body(document, body_matches(document, left[0], right[0]))
        position_notes(document, left, right, counts, separate_endnotes)
    document.save(target)
