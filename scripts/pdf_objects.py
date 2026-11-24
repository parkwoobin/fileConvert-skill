"""Reconstruct PDF paint objects as individually editable Word objects."""
import argparse
from io import BytesIO
from pathlib import Path
import re

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.transformPen import TransformPen
from fontTools.svgLib.path import parse_path
from lxml import etree
from PIL import Image
import pymupdf as pdf

V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
SVG = "{http://www.w3.org/2000/svg}"
XLINK = "{http://www.w3.org/1999/xlink}href"


def tint(values):
    if not values:
        return "000000"
    if len(values) == 4:
        c, m, y, k = values
        values = tuple(1 - min(v + k, 1) for v in (c, m, y))
    return "".join(f"{round(max(0, min(v, 1))*255):02X}" for v in (values * 3 if len(values) == 1 else values)[:3])


def intersect(a, b):
    return a & b


def svg_clips(page):
    root = etree.fromstring(page.get_svg_image(text_as_path=False).encode())
    clips = {}
    for clip in root.iter(SVG + "clipPath"):
        box = pdf.Rect()
        for path in clip.iter(SVG + "path"):
            pen = BoundsPen(None)
            matrix = path.get("transform", "")
            values = tuple(map(float, re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", matrix)))
            target = TransformPen(pen, values) if matrix.startswith("matrix(") and len(values) == 6 else pen
            parse_path(path.get("d", ""), target)
            if pen.bounds:
                box |= pdf.Rect(pen.bounds)
        if not box.is_empty:
            clips[clip.get("id")] = box
    definitions = {node.get("id"): node for node in root.iter() if node.get("id")}
    text_clips, image_clips = [], []

    def visit(node, active):
        if node.tag == SVG + "defs":
            return
        value = node.get("clip-path", "")
        if value.startswith("url(#") and value[5:-1] in clips:
            active = intersect(active, clips[value[5:-1]])
        if node.tag == SVG + "text":
            text_clips.append(active)
        if node.tag in (SVG + "use", SVG + "image"):
            referent = definitions.get((node.get(XLINK) or "")[1:])
            if node.tag == SVG + "image" or (referent is not None and any(x.tag == SVG + "image" for x in referent.iter())):
                image_clips.append(active)
        for child in node:
            visit(child, active)
    visit(root, page.rect)
    return text_clips, image_clips


def drawing_events(page):
    stack = []
    for drawing in page.get_drawings(extended=True):
        level = drawing["level"]
        while stack and stack[-1][0] >= level:
            stack.pop()
        kind = drawing["type"]
        if kind == "clip":
            parent = stack[-1][1] if stack else page.rect
            stack.append((level, parent & drawing["scissor"]))
        elif kind not in {"group"}:
            clip = stack[-1][1] if stack else page.rect
            if clip.is_empty or not (drawing["rect"] + (-0.5, -0.5, 0.5, 0.5)).intersects(clip):
                continue
            yield drawing["seqno"], drawing, clip


def visible_text(page, clips):
    traces = page.get_texttrace()
    if len(traces) != len(clips):
        clips = [page.rect] * len(traces)
    seen = set()
    for span, clip in reversed(list(zip(traces, clips))):
        if span["type"] != 0 or span["opacity"] < .05 or span["dir"][0] < .98:
            continue
        chars = []
        for char in span["chars"]:
            code, _, origin, bbox = char
            box = pdf.Rect(bbox)
            if 0 < code <= 0x10ffff and not 0xd800 <= code <= 0xdfff and (box & clip).get_area() >= box.get_area() * .70:
                chars.append(char)
        if not chars or not "".join(chr(c[0]) for c in chars).strip():
            continue
        start = chars[0][2]
        key = (round(start[0], 1), round(start[1], 1), tuple(c[0] for c in chars))
        if key in seen:
            continue
        seen.add(key)
        yield span["seqno"], span, chars


class Writer:
    def __init__(self):
        self.doc = Document()
        self.doc.settings.element.append(OxmlElement("w:doNotAutoCompressPictures"))
        self.count = 0
        self.page_size = None

    def shape(self, paragraph, box, seq, tag="shape"):
        self.count += 1
        node = etree.Element(V + tag)
        node.set("id", f"PDFObject{self.count}")
        node.set(O + "allowoverlap", "t")
        width, height = max(.1, box.width), max(.1, box.height)
        node.set("style", f"position:absolute;margin-left:{box.x0:.3f}pt;margin-top:{box.y0:.3f}pt;"
                 f"width:{width:.3f}pt;height:{height:.3f}pt;z-index:{251658240+seq};"
                 "mso-position-horizontal-relative:page;mso-position-vertical-relative:page;"
                 "mso-wrap-style:none;v-text-anchor:top")
        pict = OxmlElement("w:pict")
        pict.append(node)
        paragraph.add_run()._element.append(pict)
        return node

    def text(self, paragraph, seq, span, chars):
        from docx.oxml.ns import qn
        from docx.text.paragraph import Paragraph

        x, baseline = chars[0][2]
        size = span["size"]
        right = max(c[3][2] for c in chars)
        rendered_width = max(4, right-x)
        text = "".join(chr(c[0]) for c in chars)
        # Word lays out a line inside each box, leaving text individually editable.
        left, top = max(0, x), max(0, baseline - size * .94)
        node = self.shape(paragraph, pdf.Rect(left, top, left + rendered_width*1.32+3, top+size*1.7), seq, "rect")
        node.set("filled", "f")
        node.set("stroked", "f")
        textbox = etree.SubElement(node, V + "textbox", inset="0,0,0,0")
        content = OxmlElement("w:txbxContent")
        textbox.append(content)
        body = OxmlElement("w:p")
        content.append(body)
        line = Paragraph(body, paragraph)
        fmt = line.paragraph_format
        fmt.space_before = fmt.space_after = Pt(0)
        fmt.line_spacing = Pt(size * 1.12)
        run = line.add_run(text)
        source_face = span["font"].split("+")[-1]
        if source_face.startswith("Malgun"):
            face = "Malgun Gothic"
        elif source_face.startswith("NanumSquare"):
            face = "NanumSquare"
        else:
            face = re.sub(r"[-,](Regular|BoldItalic|Bold|Italic)$", "", source_face)
        run.font.name = face
        run.font.size = Pt(size)
        run.font.bold = span["font"].endswith(("B", "B00", "Bold"))
        run.font.color.rgb = RGBColor.from_string(tint(span["color"]))
        rpr = run._element.get_or_add_rPr()
        rpr.get_or_add_rFonts().set(qn("w:eastAsia"), face)
        if len(chars) > 1:
            natural = sum(max(0, c[3][2]-c[3][0]) for c in chars[:-1])
            tracking = round(((chars[-1][2][0] - x - natural)/(len(chars)-1))*20)
            if -30 <= tracking <= 30 and tracking:
                spacing = OxmlElement("w:spacing")
                spacing.set(qn("w:val"), str(tracking))
                rpr.insert_element_before(spacing, "w:sz", "w:szCs")

    def path(self, paragraph, seq, draw, clip):
        original = draw["rect"]
        rect = pdf.Rect(original)
        if rect.width <= 0:
            rect.x1 = rect.x0 + .1
        if rect.height <= 0:
            rect.y1 = rect.y0 + .1
        clipped = rect & clip
        if clipped.is_empty:
            return
        if rect.get_area() > 0 and clipped.get_area() < rect.get_area() * .96:
            if len(draw["items"]) != 1 or draw["items"][0][0] != "re":
                return
            rect = clipped
        node = self.shape(paragraph, rect, seq)

        def xy(p):
            return f"{round((p.x-rect.x0)*100)},{round((p.y-rect.y0)*100)}"

        path = []
        end = None
        for item in draw["items"]:
            kind = item[0]
            if kind == "re":
                rr = item[1] & rect
                path += ["m " + xy(rr.tl), "l " + " ".join(map(xy, (rr.tr, rr.br, rr.bl))), "x"]
                end = None
            elif kind in {"l", "c"}:
                if end != item[1]:
                    path.append("m " + xy(item[1]))
                path.append(kind + " " + " ".join(map(xy, item[2:])))
                end = item[-1]
            elif kind == "qu":
                q = item[1]
                path += ["m " + xy(q.ul), "l " + " ".join(map(xy, (q.ur, q.lr, q.ll))), "x"]
                end = None
        if draw.get("closePath") and end is not None:
            path.append("x")
        node.set("coordsize", f"{max(10,round(rect.width*100))},{max(10,round(rect.height*100))}")
        node.set("path", " ".join(path+["e"]))
        if draw.get("fill") is None:
            node.set("filled", "f")
        else:
            node.set("fillcolor", "#" + tint(draw["fill"]))
            if draw.get("fill_opacity", 1) < .999:
                etree.SubElement(node, V + "fill", opacity=f"{draw['fill_opacity']:.4f}")
        if draw.get("color") is None:
            node.set("stroked", "f")
        else:
            node.set("strokecolor", "#" + tint(draw["color"]))
            node.set("strokeweight", f"{draw.get('width') or .25:.3f}pt")

    def picture(self, paragraph, page, seq, box, clip, block):
        rect = box & clip
        if rect.is_empty or rect.width < 2 or rect.height < 2:
            return
        if block and box.contains(pdf.Rect(block["bbox"])) and block["transform"][1:3] == (0.0, 0.0):
            with Image.open(BytesIO(block["image"])) as decoded:
                picture = decoded.convert("RGBA")
            if block.get("mask"):
                with Image.open(BytesIO(block["mask"])) as decoded_mask:
                    picture.putalpha(decoded_mask.convert("L").resize(picture.size))
            left, top = (rect.x0 - box.x0) / box.width, (rect.y0 - box.y0) / box.height
            right, bottom = (rect.x1 - box.x0) / box.width, (rect.y1 - box.y0) / box.height
            px, py = picture.size
            picture = picture.crop((round(left*px), round(top*py), round(right*px), round(bottom*py)))
            if not picture.width or not picture.height:
                return
            stream = BytesIO()
            picture.save(stream, "PNG")
        else:
            crop = page.get_pixmap(matrix=pdf.Matrix(2, 2), clip=rect)
            stream = BytesIO(crop.tobytes("png"))
        stream.seek(0)
        rel, _ = self.doc.part.get_or_add_image(stream)
        node = self.shape(paragraph, rect, seq, "rect")
        node.set("stroked", "f")
        pic = etree.SubElement(node, V + "imagedata")
        pic.set(R + "id", rel)

    def add_page(self, page):
        size = (page.rect.width, page.rect.height)
        new_section = self.page_size is not None and size != self.page_size
        if new_section:
            section = self.doc.add_section(WD_SECTION_START.NEW_PAGE)
        else:
            section = self.doc.sections[-1]
        self.page_size = size
        section.page_width, section.page_height = Pt(page.rect.width), Pt(page.rect.height)
        for side in ("top", "bottom", "left", "right"):
            setattr(section, side + "_margin", Pt(0))
        section.header_distance = section.footer_distance = Pt(0)
        anchor = self.doc.add_paragraph()
        f = anchor.paragraph_format
        f.page_break_before = page.number > 0 and not new_section
        f.space_before = f.space_after = Pt(0)
        f.line_spacing = Pt(1)
        text_clips, image_clips = svg_clips(page)
        images = [(i, pdf.Rect(box)) for i, (kind, box) in enumerate(page.get_bboxlog()) if kind == "fill-image"]
        if len(images) != len(image_clips):
            image_clips = [page.rect] * len(images)
        events = [(seq, "path", (draw, clip)) for seq, draw, clip in drawing_events(page)]
        events += [(seq, "text", (span, chars)) for seq, span, chars in visible_text(page, text_clips)]
        blocks = [block for block in page.get_text("dict")["blocks"] if block["type"] == 1]
        if len(blocks) != len(images):
            blocks = [None] * len(images)
        events += [(seq, "picture", (page, box, clip, block))
                   for (seq, box), clip, block in zip(images, image_clips, blocks)]
        covers = [(seq, draw["rect"] & clip) for seq, kind, args in events if kind == "path"
                  for draw, clip in [args] if draw.get("fill") is not None
                  and draw.get("fill_opacity", 1) > .99 and len(draw["items"]) == 1 and draw["items"][0][0] == "re"]
        for seq, kind, args in sorted(events, key=lambda x: x[0]):
            if kind == "text":
                _, chars = args
                box = pdf.Rect()
                for c in chars:
                    box |= pdf.Rect(c[3])
                if any(later > seq and cover.contains(box) for later, cover in covers):
                    continue
                self.text(anchor, seq, *args)
            elif kind == "path":
                self.path(anchor, seq, *args)
            else:
                self.picture(anchor, args[0], seq, *args[1:])


def reconstruct(source, target, limit=None):
    writer = Writer()
    with pdf.open(source) as document:
        for i, page in enumerate(document):
            if limit is not None and i >= limit:
                break
            writer.add_page(page)
    with target.open("xb") as stream:
        writer.doc.save(stream)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", type=Path)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--pages", type=int)
    args = parser.parse_args()
    reconstruct(args.pdf, args.docx, args.pages)


if __name__ == "__main__":
    main()
