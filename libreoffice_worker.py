"""Long-lived LibreOffice Calc worker for the Excel Preview plugin.

Starts one headless soffice with a private profile, then answers one JSON
request per stdin line with one JSON response per stdout line:

  {"id": 1, "op": "recalculate", "input": "/tmp/a.xlsx", "output": "/tmp/b.xlsx", "filter": "Calc MS Excel 2007 XML"}
  {"id": 2, "op": "edit", "input": ..., "output": ..., "filter": ..., "edits": [{"sheet": "Summary", "cell": "B5", "input": "=B6-1"}]}

Edits are applied the way typing into a cell does (Excel A1 syntax, en-US
separators), then the whole document is recalculated and stored.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException

# Clear values, dates, strings and formulas; keep formatting and notes.
CLEAR_CONTENT = 1 | 2 | 4 | 16
# org.openoffice.Office.Calc/Formula/Syntax/Grammar: 1 = Excel A1.
EXCEL_A1_GRAMMAR = 1


def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def start_office(soffice):
    profile = tempfile.mkdtemp(prefix="excel-preview-lo-")
    pipe = "excel_preview_" + str(os.getpid())
    process = subprocess.Popen(
        [
            soffice,
            "--headless",
            "--invisible",
            "--nologo",
            "--nodefault",
            "--norestore",
            "--nolockcheck",
            "-env:UserInstallation=" + uno.systemPathToFileUrl(profile),
            "--accept=pipe,name=" + pipe + ";urp;StarOffice.ComponentContext",
        ],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local
    )
    deadline = time.monotonic() + 60
    while True:
        try:
            context = resolver.resolve(
                "uno:pipe,name=" + pipe + ";urp;StarOffice.ComponentContext"
            )
            break
        except NoConnectException:
            if process.poll() is not None or time.monotonic() > deadline:
                raise RuntimeError("LibreOffice did not start.")
            time.sleep(0.1)
    manager = context.ServiceManager
    use_excel_formula_syntax(manager, context)
    desktop = manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    return process, profile, desktop


def use_excel_formula_syntax(manager, context):
    provider = manager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", context
    )
    access = provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess",
        (prop("nodepath", "/org.openoffice.Office.Calc/Formula/Syntax"),),
    )
    access.setPropertyValue("Grammar", EXCEL_A1_GRAMMAR)
    access.commitChanges()


def load(desktop, path):
    return desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(path),
        "_blank",
        0,
        (prop("Hidden", True), prop("ReadOnly", False), prop("UpdateDocMode", 0)),
    )


def apply_edits(document, edits):
    sheets = document.getSheets()
    for edit in edits:
        if not sheets.hasByName(edit["sheet"]):
            raise ValueError("No sheet named " + edit["sheet"] + ".")
        cell = sheets.getByName(edit["sheet"]).getCellRangeByName(edit["cell"])
        if edit["input"] == "":
            cell.clearContents(CLEAR_CONTENT)
        else:
            cell.FormulaLocal = edit["input"]


def handle(desktop, request):
    document = load(desktop, request["input"])
    if document is None:
        raise ValueError("LibreOffice could not open the file.")
    try:
        if request["op"] == "edit":
            apply_edits(document, request["edits"])
        document.calculateAll()
        document.storeToURL(
            uno.systemPathToFileUrl(request["output"]),
            (prop("FilterName", request["filter"]),),
        )
    finally:
        document.close(True)


def main():
    soffice = sys.argv[1] if len(sys.argv) > 1 else "soffice"
    process, profile, desktop = start_office(soffice)
    out = sys.stdout
    out.write(json.dumps({"ready": True}) + "\n")
    out.flush()
    try:
        for line in sys.stdin:
            if not line.strip():
                continue
            request = json.loads(line)
            try:
                handle(desktop, request)
                response = {"id": request["id"], "ok": True}
            except Exception as error:  # reported to the caller, worker keeps running
                message = getattr(error, "Message", None) or str(error) or type(error).__name__
                response = {"id": request["id"], "ok": False, "error": message}
            out.write(json.dumps(response) + "\n")
            out.flush()
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
        shutil.rmtree(profile, ignore_errors=True)


if __name__ == "__main__":
    main()
