"""Carry explicit Word table borders into Hancom's editable cell border styles."""
from copy import deepcopy
import re
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree


SECTION = re.compile(r"Contents/section(\d+)\.xml")
SIDES = ("top", "bottom", "left", "right")
TYPES = {
    "single": "SOLID",
    "double": "DOUBLE_SLIM",
    "dotted": "DOT",
    "dashed": "DASH",
    "dotDash": "DASH_DOT",
    "dotDotDash": "DASH_DOT_DOT",
    "none": "NONE",
    "nil": "NONE",
}
WIDTHS_MM = (.1, .12, .15, .2, .25, .3, .4, .5, .6, .7, .8, 1, 1.5, 2, 3, 4, 5)


def local(node):
    return etree.QName(node).localname if isinstance(node.tag, str) else ""


def child(parent, tag):
    return next((node for node in parent if local(node) == tag), None)


def normalized(text):
    return "".join(char for char in text if char.isalnum()).casefold()


def word_cell_specs(table):
    row_nodes = table.xpath("./w:tr")
    columns = len(table.xpath("./w:tblGrid/w:gridCol"))
    if not row_nodes or not columns:
        return None
    table_borders = {
        local(border): border for border in table.xpath("./w:tblPr/w:tblBorders/*")
    }
    cells = {}
    for row, row_node in enumerate(row_nodes):
        before = row_node.xpath("./w:trPr/w:gridBefore/@w:val")
        column = int(before[0]) if before else 0
        for node in row_node.xpath("./w:tc"):
            spans = node.xpath("./w:tcPr/w:gridSpan/@w:val")
            span = int(spans[0]) if spans else 1
            merge = node.xpath("./w:tcPr/w:vMerge")
            if not merge or merge[0].get(qn("w:val")) == "restart":
                explicit = {local(border): border for border in node.xpath("./w:tcPr/w:tcBorders/*")}
                specs = {}
                for side in SIDES:
                    outer = {
                        "top": row == 0,
                        "bottom": row == len(row_nodes) - 1,
                        "left": column == 0,
                        "right": column + span == columns,
                    }[side]
                    table_side = side if outer else "insideH" if side in {"top", "bottom"} else "insideV"
                    border = explicit.get(side)
                    # For merge starts, keep Hancom's bottom border unless Word specified it.
                    if border is None and not (merge and side == "bottom"):
                        border = table_borders.get(table_side)
                    if border is not None:
                        converted = to_hancom_border(border)
                        if converted:
                            specs[side + "Border"] = converted
                cells[(row, column)] = specs
            column += span
    seed = normalized("".join(table.xpath(".//w:t/text()")))[:32]
    return (len(row_nodes), columns), seed, cells


def to_hancom_border(border):
    style = TYPES.get(border.get(qn("w:val")))
    if style is None:
        return None
    info = {"type": style}
    size = border.get(qn("w:sz"))
    if style != "NONE" and size and size.isdigit() and int(size) > 0:
        mm = int(size) / 8 * 25.4 / 72
        nearest = min(WIDTHS_MM, key=lambda choice: abs(choice - mm))
        info["width"] = f"{nearest:g} mm"
    color = border.get(qn("w:color"))
    if color and re.fullmatch(r"[0-9a-fA-F]{6}", color):
        info["color"] = "#" + color.upper()
    return tuple(sorted(info.items()))


def read_word_borders(source):
    body = Document(source).element.body
    return [word_cell_specs(table) for table in body.xpath(".//w:tbl")]


def hancom_cells(table):
    cells = {}
    for row in table:
        if local(row) != "tr":
            continue
        for cell in row:
            if local(cell) != "tc":
                continue
            addr = child(cell, "cellAddr")
            if addr is not None:
                cells[(int(addr.get("rowAddr")), int(addr.get("colAddr")))] = cell
    return cells


def repair_borders(source_specs, imported, target):
    with ZipFile(imported) as package:
        header = etree.fromstring(package.read("Contents/header.xml"))
        border_list = next((node for node in header.iter() if local(node) == "borderFills"), None)
        if border_list is None:
            return False
        borders = {node.get("id"): node for node in border_list if local(node) == "borderFill"}
        if not borders:
            return False
        next_id = max(map(int, borders)) + 1
        sections = sorted(
            (int(match.group(1)), name) for name in package.namelist() if (match := SECTION.fullmatch(name))
        )
        roots = {name: etree.fromstring(package.read(name)) for _, name in sections}
        tables = [(name, node) for name, root in roots.items() for node in root.iter() if local(node) == "tbl"]
        if len(tables) != len(source_specs):
            return False
        clones = {}
        changed = set()
        for source, (name, table) in zip(source_specs, tables):
            if source is None:
                continue
            dimensions, seed, specs = source
            if dimensions != (int(table.get("rowCnt", 0)), int(table.get("colCnt", 0))):
                continue
            text = normalized("".join(node.text or "" for node in table.iter() if local(node) == "t"))
            if seed and not text.startswith(seed):
                continue
            cells = hancom_cells(table)
            if cells.keys() != specs.keys():
                continue
            for coordinate, sides in specs.items():
                cell = cells[coordinate]
                fill_id = cell.get("borderFillIDRef")
                if not sides or fill_id not in borders:
                    continue
                base = borders[fill_id]
                adjustments = {}
                for tag, attributes in sides.items():
                    line = child(base, tag)
                    if line is not None and any(line.get(key) != value for key, value in attributes):
                        adjustments[tag] = attributes
                if not adjustments:
                    continue
                key = fill_id, tuple(sorted(adjustments.items()))
                if key not in clones:
                    clone = deepcopy(base)
                    clone.set("id", str(next_id))
                    for tag, attributes in adjustments.items():
                        child(clone, tag).attrib.update(dict(attributes))
                    clones[key] = str(next_id)
                    next_id += 1
                    border_list.append(clone)
                cell.set("borderFillIDRef", clones[key])
                changed.add(name)
        if not changed:
            return False
        border_list.set("itemCnt", str(len(border_list)))
        changed_parts = {name: roots[name] for name in changed}
        changed_parts["Contents/header.xml"] = header
        with ZipFile(target, "w") as updated:
            for member in package.infolist():
                root = changed_parts.get(member.filename)
                content = (etree.tostring(root, encoding="UTF-8", xml_declaration=True)
                           if root is not None else package.read(member))
                updated.writestr(member, content)
    return True
