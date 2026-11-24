"""Keep each PDF page in one editable Hancom section and preserve its paper geometry."""
from copy import deepcopy
import re
import shutil
from zipfile import ZipFile

from docx import Document
from lxml import etree
import pymupdf

SECTION = re.compile(r"Contents/section(\d+)\.xml")
SIDES = ("top", "bottom", "left", "right")
FACTORS = (1.0, .85, .72, .60, .52)


def local(node):
    return etree.QName(node).localname if isinstance(node.tag, str) else ""


def serialize(root):
    return etree.tostring(root, encoding="utf-8", xml_declaration=True)


def normalize(text):
    return "".join(char for char in text if char.isalnum()).casefold()


def section_seed(root):
    for paragraph in root:
        text = normalize("".join(node.text or "" for node in paragraph.iter() if local(node) == "t"))
        if len(text) >= 10:
            return text[:24]
    return ""


def section_overflows(rendered, seeds):
    with pymupdf.open(rendered) as pdf:
        pages = [normalize(page.get_text()) for page in pdf]
    starts = []
    for seed in seeds:
        if not seed:
            return set(range(len(seeds)))
        hits = [i for i, text in enumerate(pages)
                if (not starts or i > starts[-1]) and seed[:20] in text]
        if not hits:
            hits = [i for i, text in enumerate(pages)
                    if (not starts or i > starts[-1]) and seed[:10] in text]
        if not hits:
            return set(range(len(seeds)))
        starts.append(hits[0])
    if not starts or starts[0] != 0:
        return set(range(len(seeds)))
    return {i for i, (a, b) in enumerate(zip(starts, starts[1:] + [len(pages)])) if b - a > 1}


def format_style(style, factor):
    for node in style.iter():
        name = local(node)
        value = node.get("value", "")
        if not value.lstrip("-").isdigit():
            continue
        amount = int(value)
        if name in {"prev", "next"} or (name == "lineSpacing" and node.get("type") == "FIXED"):
            node.set("value", str(round(amount * factor)))
        elif name == "lineSpacing" and node.get("type") == "PERCENT":
            node.set("value", str(min(amount, max(65, round(amount * factor)))))


def write_layout(raw, target, geometries, levels):
    with ZipFile(raw) as original:
        members = [
            (int(match.group(1)), info.filename)
            for info in original.infolist()
            if (match := SECTION.fullmatch(info.filename))
        ]
        members.sort()
        if len(members) != len(geometries):
            raise ValueError("Hancom import did not create one section for each PDF page")
        header = etree.fromstring(original.read("Contents/header.xml"))
        style_list = next(node for node in header.iter() if local(node) == "paraProperties")
        styles = {style.get("id"): style for style in style_list if local(style) == "paraPr"}
        next_id = max(map(int, styles), default=-1) + 1
        clones = {}
        updated = {}
        seeds = []
        for index, (_, member) in enumerate(members):
            root = etree.fromstring(original.read(member))
            seeds.append(section_seed(root))
            page = next(node for node in root.iter() if local(node) == "pagePr")
            margins = next(node for node in page if local(node) == "margin")
            width, height, sides = geometries[index]
            page.set("width", str(width))
            page.set("height", str(height))
            for name, amount in zip(SIDES, sides):
                margins.set(name, str(amount))
            for name in ("header", "footer", "gutter"):
                margins.set(name, "0")
            level = levels[index]
            if level:
                for node in root.iter():
                    style_id = node.get("paraPrIDRef")
                    if style_id not in styles:
                        continue
                    key = (style_id, level)
                    if key not in clones:
                        style = deepcopy(styles[style_id])
                        new_id = str(next_id)
                        next_id += 1
                        style.set("id", new_id)
                        format_style(style, FACTORS[level])
                        style_list.append(style)
                        clones[key] = new_id
                    node.set("paraPrIDRef", clones[key])
            updated[member] = serialize(root)
        if clones:
            style_list.set("itemCnt", str(len(styles) + len(clones)))
        updated["Contents/header.xml"] = serialize(header)
        with ZipFile(target, "w") as candidate:
            for info in original.infolist():
                data = updated[info.filename] if info.filename in updated else original.read(info)
                candidate.writestr(info, data)
    return seeds


def fit(source, intermediate, target, office):
    """Export a Hancom document only when a checked layout fits the PDF page count."""
    sections = Document(intermediate).sections
    with pymupdf.open(source) as pdf:
        if len(pdf) != len(sections):
            raise ValueError("Could not reconstruct one editable section per PDF page")
        geometries = [
            (round(page.rect.width * 100), round(page.rect.height * 100),
             [round(getattr(section, side + "_margin").pt * 100) for side in SIDES])
            for page, section in zip(pdf, sections)
        ]
    scratch = target.parent
    imported = scratch / "hancom-import.hwpx"
    office("hancom", intermediate, imported)
    levels = [0] * len(geometries)
    for attempt in range(len(FACTORS) + 1):
        candidate = scratch / f"hancom-fit-{attempt}.hwpx"
        preview = scratch / f"hancom-fit-{attempt}.pdf"
        seeds = write_layout(imported, candidate, geometries, levels)
        office("hancom", candidate, preview)
        with pymupdf.open(preview) as pdf:
            count = len(pdf)
        if count == len(geometries):
            if target.suffix.lower() == ".hwp":
                office("hancom", candidate, target)
            else:
                shutil.copyfile(candidate, target)
            return
        if attempt == len(FACTORS) - 1:
            if all(level == len(FACTORS) - 1 for level in levels):
                break
            levels = [len(FACTORS) - 1] * len(geometries)
        else:
            overflowing = section_overflows(preview, seeds) if count > len(geometries) else set()
            if not overflowing:
                break
            for index in overflowing:
                levels[index] = min(levels[index] + 1, len(FACTORS) - 1)
    raise ValueError(f"The editable Hancom layout could not fit {len(geometries)} PDF pages (last attempt: {count})")
