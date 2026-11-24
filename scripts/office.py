"""Native Office conversion in a subprocess so the CLI can time out."""
import sys
from pathlib import Path


def hancom(source, target, module):
    from win32com.client import DispatchEx

    formats = {".hwp": "HWP", ".hwpx": "HWPX", ".docx": "OOXML", ".pdf": "PDF"}
    app = DispatchEx("HWPFrame.HwpObject")
    try:
        if module and not app.RegisterModule("FilePathCheckDLL", module):
            raise RuntimeError("Hancom did not register the configured file path module")
        if not app.Open(str(source), formats[source.suffix.lower()], ""):
            raise RuntimeError("Hancom could not open the input")
        if not app.SaveAs(str(target), formats[target.suffix.lower()], ""):
            raise RuntimeError("Hancom did not export the target format")
    finally:
        app.Quit()


def word(source, target):
    from win32com.client import DispatchEx

    app = DispatchEx("Word.Application")
    try:
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3
        app.Options.UpdateLinksAtOpen = False
        doc = app.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
        try:
            if target.suffix.lower() == ".pdf":
                doc.ExportAsFixedFormat(str(target), 17)
            else:
                doc.SaveAs2(str(target), FileFormat={".doc": 0, ".docx": 16}[target.suffix.lower()])
        finally:
            doc.Close(False)
    finally:
        app.Quit()


def powerpoint(source, target):
    from win32com.client import DispatchEx

    app = DispatchEx("PowerPoint.Application")
    # PowerPoint is a single-instance application. Close only the opened deck.
    old_security = app.AutomationSecurity
    try:
        app.AutomationSecurity = 3
        deck = app.Presentations.Open(str(source), True, False, False)
        try:
            deck.SaveAs(str(target), {".ppt": 1, ".pdf": 32}[target.suffix.lower()])
        finally:
            deck.Close()
    finally:
        app.AutomationSecurity = old_security


if __name__ == "__main__":
    engine, source_path, target_path, module_name = sys.argv[1:]
    source, target = Path(source_path), Path(target_path)
    if engine == "hancom":
        hancom(source, target, module_name)
    else:
        {"word": word, "powerpoint": powerpoint}[engine](source, target)
