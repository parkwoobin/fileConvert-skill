"""The HWPX exporter's small runtime subset, with bounded XML reads."""
import re
import unicodedata
from zipfile import ZipFile

from lxml import etree


P = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
SECTION = re.compile(r"Contents/section(\d+)\.xml")
MAX_SECTION = 24 * 1024 * 1024


class NotHwpxError(ValueError):
    pass


def ensure_hwpx(path):
    with ZipFile(path) as package:
        if "Contents/content.hpf" not in package.namelist() or not section_names(package):
            raise NotHwpxError("HWPX body sections were not found")


def section_names(package):
    return sorted(
        (name for name in package.namelist() if SECTION.fullmatch(name)),
        key=lambda name: int(SECTION.fullmatch(name).group(1)),
    )


def read_section(package, name):
    if package.getinfo(name).file_size > MAX_SECTION:
        raise NotHwpxError("HWPX section exceeds 24 MiB")
    parser = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False)
    root = etree.fromstring(package.read(name), parser)
    if root.getroottree().docinfo.doctype:
        raise NotHwpxError("HWPX XML must not contain a DTD")
    return root


def strip_pua(text):
    return "".join(char for char in text if unicodedata.category(char) != "Co")
