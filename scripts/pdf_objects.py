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


def split_fields(chars):
    """Keep white space reserved for handwriting at its original position."""
    start = 0
    index = 0
    while index < len(chars):
        if not chr(chars[index][0]).isspace():
            index += 1
            continue
        end = index + 1
        while end < len(chars) and chr(chars[end][0]).isspace():
            end += 1
        if end - index >= 2 or index == start or end == len(chars):
            if index > start:
                yield chars[start:index]
            start = end
        index = end
    if start < len(chars):
        yield chars[start:]


def text_rows(page, clips):
    traces = page.get_texttrace()
    if len(traces) != len(clips):
        clips = [page.rect] * len(traces)
    for span, clip in zip(traces, clips):
        if span["type"] != 0 or span["opacity"] < .05 or span["dir"][0] < .98:
            continue
        rows = {}
        for char in span["chars"]:
            code, _, origin, bbox = char
            box = pdf.Rect(bbox)
            if 0 < code <= 0x10ffff and not 0xd800 <= code <= 0xdfff and (box & clip).get_area() >= box.get_area() * .70:
                rows.setdefault(round(origin[1], 1), []).append(char)
        for chars in rows.values():
            for field in split_fields(chars):
                yield span["seqno"], span, field


def char_bounds(chars):
    box = pdf.Rect()
    for char in chars:
        box |= pdf.Rect(char[3])
    return box


def visible_text(page, clips, covers):
    import numpy as np

    raster = page.get_pixmap(matrix=pdf.Matrix(2, 2), alpha=False)
    pixels = np.frombuffer(raster.samples, dtype=np.uint8).reshape(raster.height,raster.width,3).astype(np.int16)
    rows = []
    for seq, span, chars in text_rows(page, clips):
        bbox = char_bounds(chars)
        if any(later > seq and cover.contains(bbox) for later, cover in covers):
            continue
        color = tuple(bytes.fromhex(tint(span["color"])))
        tested, hits = 0, 0
        for code, _, (x,y), limits in chars:
            if chr(code).isspace():
                continue
            # Sample the intended glyph rectangle, excluding adjacent baselines.
            left, right = max(0,int(2*x)), min(raster.width,int(2*limits[2])+1)
            top, bottom = max(0,int(2*(y-.86*span["size"]))), min(raster.height,int(2*y+2))
            sample = pixels[top:bottom,left:right]
            if sample.size:
                tested += 1
                hits += bool((np.abs(sample - color).max(axis=2) <= 53).any())
        if tested and hits >= max(1, tested*.38):
            rows.append((seq,span,chars,hits/tested))
    kept = {}
    for seq, span, chars, score in sorted(rows,key=lambda row:(row[3],row[0]),reverse=True):
        x,y=chars[0][2]
        key="".join(chr(c[0]) for c in chars).strip()
        if any(abs(x-xx)<=3 and abs(y-yy)<=5 for xx,yy in kept.get(key, ())):
            continue
        kept.setdefault(key,[]).append((x,y))
        yield seq,span,chars


def coalesce_rows(rows):
    bands = []
    for seq, span, chars in sorted(rows, key=lambda row: (row[2][0][2][1], row[2][0][2][0])):
        baseline = chars[0][2][1]
        if bands and abs(bands[-1][0] - baseline) <= .4:
            bands[-1][1].append((seq, span, chars))
        else:
            bands.append((baseline, [(seq, span, chars)]))
    for _, band in bands:
        current = None
        for seq, span, chars in sorted(band, key=lambda row: row[2][0][2][0]):
            if current:
                prev_seq, prev_style, prev_chars = current
                end = char_bounds(prev_chars).x1
                start, y = chars[0][2]
                gap = start - end
                if (
                    -.16 * span["size"] <= gap <= .46 * span["size"]
                    and span["font"] == prev_style["font"]
                    and abs(span["size"] - prev_style["size"]) < .08
                    and span["color"] == prev_style["color"]
                ):
                    spacer = []
                    if gap >= .12 * span["size"] and not chr(prev_chars[-1][0]).isspace() and not chr(chars[0][0]).isspace():
                        last = prev_chars[-1][3]
                        spacer = [(32, 0, (end, y), (end, last[1], start, last[3]))]
                    current = max(seq, prev_seq), prev_style, prev_chars + spacer + chars
                    continue
                yield current
            current = seq, span, chars
        if current:
            yield current


def clip_segment(a, b, rect):
    dx, dy = b.x-a.x, b.y-a.y
    lo, hi = 0., 1.
    for p,q in ((-dx,a.x-rect.x0),(dx,rect.x1-a.x),(-dy,a.y-rect.y0),(dy,rect.y1-a.y)):
        if p < 0:
            lo = max(lo,q/p)
        elif p > 0:
            hi = min(hi,q/p)
        elif q < 0:
            return None
    if hi <= lo:
        return None
    return pdf.Point(a.x+lo*dx,a.y+lo*dy), pdf.Point(a.x+hi*dx,a.y+hi*dy)


def text_blocks(rows):
    blocks = []
    for seq, span, chars in sorted(rows, key=lambda row: (row[2][0][2][1], row[2][0][2][0])):
        x,y=chars[0][2]
        size=span["size"]
        for candidate in reversed(blocks):
            _, previous, earlier = candidate[-1]
            x0,y0=earlier[0][2]
            gap=y-y0
            if (
                .90 * size <= gap <= 2.12 * size
                and abs(x-x0) <= 2.25
                and abs(size-previous["size"]) <= .18
                and span["font"] == previous["font"]
                and tint(span["color"]) == tint(previous["color"])
                and len(chars) >= 8 and len(earlier) >= 8
            ):
                candidate.append((seq,span,chars))
                break
        else:
            candidate = None
        if candidate is None or candidate[-1][2] is not chars:
            blocks.append([(seq,span,chars)])
    return blocks



def clip_polygon(points, rect):
    for axis, border, direction in ((0,rect.x0,1),(0,rect.x1,-1),(1,rect.y0,1),(1,rect.y1,-1)):
        if not points:
            break
        clipped=[]
        prev=points[-1]
        for point in points:
            inside=(point[axis]-border)*direction >= 0
            before=(prev[axis]-border)*direction >= 0
            if inside != before:
                ratio=(border-prev[axis])/(point[axis]-prev[axis])
                clipped.append(pdf.Point(*(a+(b-a)*ratio for a,b in zip(prev,point))))
            if inside:
                clipped.append(point)
            prev=point
        points=clipped
    return points


def contours(items):
    points=[]
    for item in items:
        kind=item[0]
        if kind in ("l","c"):
            if points and points[-1] != item[1]:
                yield points
                points=[]
            if not points:
                points=[item[1]]
            if kind=="l":
                points.append(item[-1])
            else:
                a,b,c,d=item[1:]
                for i in range(1,13):
                    t=i/12
                    u=1-t
                    points.append(pdf.Point(*(u**3*x+3*u*u*t*y+3*u*t*t*z+t**3*w
                                              for x,y,z,w in zip(a,b,c,d))))
        else:
            if points:
                yield points
                points=[]
            if kind=="re":
                r=item[1]
                yield [r.tl,r.tr,r.br,r.bl]
            elif kind=="qu":
                q=item[1]
                yield [q.ul,q.ur,q.lr,q.ll]
    if points:
        yield points


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

    def text(self, paragraph, seq, lines):
        from docx.text.paragraph import Paragraph
        from docx.oxml.ns import qn

        _, first_span, first_chars = lines[0]
        x, baseline = first_chars[0][2]
        size = first_span["size"]
        right = max(char_bounds(chars).x1 for _,_,chars in lines)
        bottom = lines[-1][2][0][2][1]
        left, top = max(0, x), max(0, baseline - size * .94)
        width = max(4, right-x)*1.24+4
        node = self.shape(paragraph, pdf.Rect(left, top, min(left+width,self.page_size[0]), bottom+size), seq, "rect")
        node.set("filled", "f")
        node.set("stroked", "f")
        textbox = etree.SubElement(node, V + "textbox", inset="0,0,0,0")
        content = OxmlElement("w:txbxContent")
        textbox.append(content)
        for index, (_, span, chars) in enumerate(lines):
            body = OxmlElement("w:p")
            content.append(body)
            line = Paragraph(body, paragraph)
            fmt = line.paragraph_format
            fmt.space_before = fmt.space_after = Pt(0)
            fmt.keep_with_next = False
            for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):
                option = OxmlElement(tag)
                option.set(qn("w:val"), "0")
                body.get_or_add_pPr().append(option)
            line_height = span["size"] * 1.12
            fmt.line_spacing = Pt(line_height)
            if index+1 < len(lines):
                gap = lines[index+1][2][0][2][1] - chars[0][2][1]
                fmt.space_after = Pt(max(0,gap-line_height))
            self.styled_run(line, span, chars)

    def styled_run(self, line, span, chars):
        from docx.oxml.ns import qn

        size = span["size"]
        x = chars[0][2][0]
        run = line.add_run("".join(chr(c[0]) for c in chars))
        source_face = span["font"].split("+")[-1]
        if source_face.startswith("Malgun"):
            face = "Malgun Gothic"
        elif source_face.startswith("NanumBarunGothic"):
            face = "NanumBarunGothic"
        elif source_face.startswith("NanumSquare"):
            face = "NanumSquare"
        else:
            face = re.sub(r"[-,](Regular|BoldItalic|Bold|Italic)$", "", source_face)
        run.font.name = face
        run.font.size = Pt(size)
        run.font.bold = span["font"].endswith(("B", "B00", "Bold"))
        run.font.color.rgb = RGBColor.from_string(tint(span["color"]))
        rpr = run._element.get_or_add_rPr()
        fonts = rpr.get_or_add_rFonts()
        fonts.set(qn("w:eastAsia"), face)
        fonts.set(qn("w:cs"), face)
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
        if clipped.get_area() < rect.get_area() * .995:
            self.clipped_path(paragraph, seq, draw, clip, clipped)
            return
        self.draw_path(paragraph, seq, draw, rect)

    def clipped_path(self, paragraph, seq, draw, clip, rect):
        paths = list(contours(draw["items"]))
        if draw.get("fill") is not None:
            polygons = [clip_polygon(points, clip) for points in paths]
            items = [("l", a, b) for points in polygons if len(points) >= 3
                     for a,b in zip(points, points[1:]+points[:1])]
            if items:
                self.draw_path(paragraph, seq, {**draw, "items": items, "color": None, "closePath": True}, rect)
        if draw.get("color") is not None:
            segments = []
            for points in paths:
                if draw.get("closePath") or (len(draw["items"]) == 1 and draw["items"][0][0] in {"re", "qu"}):
                    points = points + points[:1]
                for a, b in zip(points, points[1:]):
                    segment = clip_segment(a, b, clip)
                    if segment:
                        segments.append(("l", *segment))
            if segments:
                self.draw_path(paragraph, seq, {**draw, "items": segments, "fill": None, "closePath": False}, rect)

    def draw_path(self, paragraph, seq, draw, rect):
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
            width = draw.get("width") or .25
            node.set("strokeweight", f"{width:.3f}pt")
            dash_spec = re.search(r"\[([^\]]+)\]", draw.get("dashes") or "")
            if dash_spec:
                parts = [max(.1, float(length) / width) for length in dash_spec[1].split()]
                if any(parts):
                    etree.SubElement(node, V + "stroke", dashstyle=" ".join(f"{length:.3g}" for length in parts))

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
        run = paragraph.add_run()
        inline = run.add_picture(stream, width=Pt(rect.width), height=Pt(rect.height))._inline
        anchor = OxmlElement("wp:anchor")
        for name, value in {
            "distT": "0", "distB": "0", "distL": "0", "distR": "0",
            "simplePos": "0", "relativeHeight": str(251658240 + seq),
            "behindDoc": "0", "locked": "0", "layoutInCell": "1", "allowOverlap": "1",
        }.items():
            anchor.set(name, value)
        simple = OxmlElement("wp:simplePos")
        simple.set("x", "0")
        simple.set("y", "0")
        anchor.append(simple)
        for direction, position in (("H", rect.x0), ("V", rect.y0)):
            placement = OxmlElement("wp:position" + direction)
            placement.set("relativeFrom", "page")
            offset = OxmlElement("wp:posOffset")
            offset.text = str(int(Pt(position)))
            placement.append(offset)
            anchor.append(placement)
        anchor.append(inline.extent)
        anchor.append(OxmlElement("wp:wrapNone"))
        anchor.append(inline.docPr)
        frame = inline.xpath("wp:cNvGraphicFramePr")
        if frame:
            anchor.append(frame[0])
        anchor.append(inline.graphic)
        inline.getparent().replace(inline, anchor)

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
        events = [
            (seq, "path", (draw, clip))
            for seq, draw, clip in drawing_events(page)
            if not (seq == 0 and draw.get("color") is None and draw.get("fill") == (1., 1., 1.)
                    and (draw["rect"] & page.rect).get_area() >= .99 * page.rect.get_area())
        ]
        blocks = [block for block in page.get_text("dict")["blocks"] if block["type"] == 1]
        if len(blocks) != len(images):
            blocks = [None] * len(images)
        events += [(seq, "picture", (page, box, clip, block))
                   for (seq, box), clip, block in zip(images, image_clips, blocks)]
        covers = [(seq, draw["rect"] & clip) for seq, kind, args in events if kind == "path"
                  for draw, clip in [args] if draw.get("fill") is not None
                  and draw.get("fill_opacity", 1) > .99 and len(draw["items"]) == 1 and draw["items"][0][0] == "re"]
        blocks = text_blocks(coalesce_rows(visible_text(page, text_clips, covers)))
        events += [(max(row[0] for row in block), "text", block) for block in blocks]
        for seq, kind, args in sorted(events, key=lambda x: x[0]):
            if kind == "text":
                self.text(anchor, seq, args)
            elif kind == "path":
                self.path(anchor, seq, *args)
            else:
                self.picture(anchor, args[0], seq, *args[1:])


def prefers_positioned_layout(source):
    """Identify designed documents whose page backgrounds must stay behind editable text."""
    with pdf.open(source) as document:
        if not len(document):
            return False
        sample_count = min(12, len(document))
        indices = (range(len(document)) if sample_count == len(document) else
                   sorted({0, 1, *(round(i * (len(document) - 1) / (sample_count - 1))
                                    for i in range(sample_count))}))
        designed = 0
        for index in indices:
            page = document[index]
            area = page.rect.get_area()
            if not area or not page.get_text().strip():
                continue
            panels = {}
            full_background = False
            for draw in page.get_drawings():
                fill = draw.get("fill")
                if fill is None or draw.get("fill_opacity", 1) < .85:
                    continue
                bounds = draw["rect"] & page.rect
                fraction = bounds.get_area() / area
                red, green, blue = bytes.fromhex(tint(fill))
                chroma = max(red, green, blue) - min(red, green, blue)
                if fraction >= .70 and (chroma >= 25 or max(red, green, blue) < 190):
                    full_background = True
                    break
                if .018 <= fraction <= .28 and chroma >= 6:
                    panels[tuple(round(value) for value in bounds)] = fraction
            if full_background or (len(panels) >= 4 and sum(panels.values()) >= .10):
                designed += 1
        minimum = 1 if len(document) <= 3 else max(2, (len(indices) + 3) // 4)
        return designed >= minimum


def reconstruct(source, target, limit=None):
    writer = Writer()
    with pdf.open(source) as document:
        for i, page in enumerate(document):
            if limit is not None and i >= limit:
                break
            writer.add_page(page)
    with target.open("xb") as stream:
        writer.doc.save(stream)


def try_positioned_hancom(source, target, office):
    """Use editable positions for designed PDFs when Hancom retains the page content."""
    from collections import Counter
    from PIL import ImageChops, ImageStat
    from shutil import copyfile

    draft = target.with_name("positioned-input.docx")
    candidate = target.with_name("positioned-output" + target.suffix)
    preview = target.with_name("positioned-check.pdf")
    reconstruct(source, draft)
    office("hancom", draft, candidate)
    office("hancom", candidate, preview)
    with pdf.open(source) as original, pdf.open(preview) as result:
        if len(original) != len(result):
            return False
        for before, after in zip(original, result):
            if max(abs(before.rect.width - after.rect.width),
                   abs(before.rect.height - after.rect.height)) > 2:
                return False
            source_text = Counter(char for char in before.get_text() if char.isalnum())
            result_text = Counter(char for char in after.get_text() if char.isalnum())
            if sum((source_text & result_text).values()) < .88 * sum(source_text.values()):
                return False
            if before.get_image_info() and not after.get_image_info():
                return False
            thumbnails = [Image.frombytes("RGB", (p.width, p.height), p.samples)
                          for page in (before, after) for p in [page.get_pixmap(matrix=pdf.Matrix(.18, .18))]]
            left, right = thumbnails
            difference = ImageStat.Stat(ImageChops.difference(left, right.resize(left.size))).mean
            if sum(difference) / len(difference) > 18:
                return False
    copyfile(candidate, target)
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", type=Path)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--pages", type=int)
    args = parser.parse_args()
    reconstruct(args.pdf, args.docx, args.pages)


if __name__ == "__main__":
    main()
