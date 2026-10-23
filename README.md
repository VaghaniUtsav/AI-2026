# -*- coding: utf-8 -*-
"""
population_evidence_tool.py
===========================
One script for the whole GA-02-006_SEG sampling. The original
population_request_tool.py (Process 1 only) is kept unchanged in _Archive\\.

Process 1 - Population (same steps as population_request_tool.py):
  1. Ask for the period (posting date)
  2. Get an SAP session (via sap_connector.py)
  3. Run FBL4N with the selections and read the ALV grid directly
  4. Save the raw data as CSV (traceability for the audit)
  5. Back up the template to _Backup\\
  6. Clear Population (D28:X) and fill it with the SAP data. Buttons,
     macros and the other tabs are left untouched.
  7. Write the row count to "Main" C4, run the sample macro behind the button (C5)
  8. Filter "Population" on "Sample Selected ?" = "Selected", save

Process 2 - Evidences, for every "Selected" sample:
  9. FB03 'Data Entry View' of the document -> picture
 10. ZRFIZWF00001 with the Samsung Document Number -> 'Display Workflow
     History' -> picture
 11. FB03 again, "File attachment" -> the file whose title starts with C460
     -> download, page 1 -> picture. Without one: the PDF that looks most
     like the invoice; without a PDF: the link (e.g. 'GTNE EVIDENCE') ->
     picture of the whole website. Pictures and files go to
     _Evidences\\<date-time>\\
 12. "Evidences" sheet like the manual version: the selected rows under the
     header (row 10), then one block per sample - red label (number and
     document number), then document, workflow and invoice picture side by
     side from column C, large enough to read at 100% zoom.

Usage:
    python population_evidence_tool.py                    -> window to pick the year and quarter(s)
    python population_evidence_tool.py --from 01.07.2026 --to 30.09.2026
    python population_evidence_tool.py --columns          -> only show SAP columns,
                                                             write nothing to Excel
    python population_evidence_tool.py --evidences-only   -> only Process 2, for the
                                                             samples already drawn
    Add --template <file> to work on another workbook (e.g. a test copy).

Needs: pip install pywin32 keyring pymupdf playwright
"""

import argparse
import collections
import contextlib
import csv
import datetime as dt
import glob
import logging
import math
import os
import re
import shutil
import sys
import threading
import time

from playwright.sync_api import sync_playwright
import pymupdf
import win32com.client as win32
import win32con
import win32gui
from win32com.shell import shell, shellcon  # pyright: ignore[reportMissingModuleSource]

import sap_connector

try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:  # no window then - the period is typed in the console
    tk = None

LOGGER = logging.getLogger("population_evidence_tool")

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
VERSION = "2026-10-23b"  # printed at the start, so the log shows which file ran
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "population_evidence_tool.log")
BASE_DIR = (r"K:\Workgroups\Support\Business Management & Automation"
            r"\03_Projekte\Aktiv\[SEG] Population Request Tool")
TEMPLATE_PATTERN = "GA-02-006_SEG.xls*"  # extension doesn't matter (.xlsx / .xlsm / .xlsb)
BACKUP_DIR_NAME = "_Backup"
RAW_EXPORT_DIR_NAME = "_Rohdaten"  # existing folder name on the share
SAVE_RAW_EXPORT = True

SHEET_NAME = "Population"
HEADER_ROW = 27
FIRST_DATA_ROW = 28
# B (formula "Selected") and C (numbers 1-30000) are fixed in the template
# and are NOT touched - the tool only reads/clears/writes D:X.
SAMPLE_COL = "B"  # "Sample Selected ?" - only read/filtered
NUMBER_COL = "C"  # "Number" - only read
DATA_FIRST_COL = "D"
DATA_LAST_COL = "X"
MAX_ROWS = 30000  # how far the numbers/formulas in B and C go

# Expected headers D27:X27 - the SAP layout /POPU must have its columns
# in exactly this order.
TEMPLATE_HEADERS = [
    "Assignment",
    "Document Number",
    "Business Area",
    "Document Type",
    "Document Date",
    "Cost Center",
    "Posting Date",
    "Posting Key",
    "Amount in Local Currency",
    "Local Currency",
    "Tax Code",
    "G/L Account",
    "Offsetting Account",
    "Clearing Document",
    "Posting Period",
    "Text",
    "Net Due Date",
    "Samsung Document Number",
    "Document Header Text",
    "Trading Partner No.",
    "G/L Account",
]
MAIN_SHEET_NAME = "Main"
MAIN_BUTTON_CELL = "C5"
MAIN_COUNT_CELL = "C4"  # "Quantity of Population" - the macro draws the samples from rows 1..this number
MAIN_BUTTON_MACRO = None  # optionally hard-code, e.g. "Modul1.SamplesZiehen"
SELECTED_VALUE = "Selected"
MSO_OLE_CONTROL = 12  # shape type of an ActiveX button

DATE_HEADERS = {"Document Date", "Posting Date", "Net Due Date"}
AMOUNT_HEADERS = {"Amount in Local Currency"}
KEY_HEADER = "Document Number"  # rows without a document number = total rows -> ignore

# SAP selections (from the VBS recording)
SAP_DATE_FORMAT = "%d.%m.%Y"  # must match the SAP user profile (SU3 > Defaults > Date Format), e.g. 01.01.2026
SAP_LAYOUT = "/POPU"
DYN_FIELD_EXCLUDE = ["BA", "BK"]  # Product Group (dynamic selection DYN009): EXCLUDED
GL_ACCOUNTS = [
    "53510100", "53510300", "53510500", "53510700",
    "53510900", "53511300", "53511500", "53516900",
]

# SAP GUI IDs
GRID_ID = "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell"
DYN_SEL_TOGGLE_BTN = "wnd[0]/tbar[1]/btn[16]"
DYN009_BTN = ("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001"
              "/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000"
              "/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106"
              "/btn%_%%DYN009_%_APP_%-VALU_PUSH")
GL_BTN = "wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH"
# Tabs in the multiple selection dialog: (tab, subscreen, table, field)
MULTI_TABS = {
    "include": ("tabpSIVA", "SAPLALDB:3010", "tblSAPLALDBSINGLE", "RSCSEL_255-SLOW_I"),
    "exclude": ("tabpNOSV", "SAPLALDB:3030", "tblSAPLALDBSINGLE_E", "RSCSEL_255-SLOW_E"),
}

# Process 2 - "Evidences" sheet, laid out like block 1 of the manual version
EVIDENCES_SHEET_NAME = "Evidences"
EVIDENCES_HEADER_ROW = 10  # same columns B-X as the Population header
EVIDENCES_FIRST_ROW = 11
EVIDENCES_GAP_ROWS = 2  # empty rows between the table and the first block
BLOCK_ROWS = 43  # one sample: red label row, pictures, 2 empty rows
SLOT_ROWS = 40  # picture height in rows, starting under the label row
# On save Excel keeps about 220 pixels per inch of the size a picture is SHOWN at and throws the
# rest away - a picture shown smaller stays blurry even zoomed in. So every picture is shown at
# READABLE_SCALE points per pixel or larger, which also makes it readable at 100% zoom.
READABLE_SCALE = 72 / 220
# Picture -> width of its slot in points, side by side from column C. A slot holds at READABLE_SCALE:
# the 1765 px FB03 picture, an A4 invoice page at INVOICE_DPI, a website page in 3 strips.
EVIDENCE_SLOTS = {"document": 600, "workflow": 600, "invoice": 1330}
PICTURE_MARGIN = 5  # points between a picture and the border of its slot
# Fixed height for the rows below the header. Automatic heights differ between the tool's hidden
# Excel and a normal one (laptop 2: 14.76 vs 14.57pt), so the pictures slid off their labels.
# 15pt is a whole number of screen pixels at 100-200% scaling.
EVIDENCE_ROW_HEIGHT = 15
LABEL_COLOR = 255  # red #FF0000 (Excel colors are BGR)
MSO_PICTURES = (11, 13)  # shape types: linked picture, picture
XL_CENTER = -4108
XL_CONTINUOUS = 1
XL_MOVE = 2  # picture moves with its cells but keeps its size, like the manual ones
EVIDENCE_DIR_NAME = "_Evidences"
COMPANY_CODE = "C460"  # FBL4N population and FB03 documents - always C460 (user, 2026-10-23)

# FB03 (standard screen) and its "File attachment" box (company-specific)
FB03_DOC_FIELD = "wnd[0]/usr/txtRF05L-BELNR"
FB03_COMPANY_FIELD = "wnd[0]/usr/ctxtRF05L-BUKRS"
FB03_YEAR_FIELD = "wnd[0]/usr/txtRF05L-GJAHR"
FB03_VIEW_BTN = "wnd[0]/tbar[1]/btn[33]"  # 'General Ledger View' <-> 'Data Entry View'
FB03_ATTACH_BTN = "wnd[0]/tbar[1]/btn[31]"  # 'File attachment'
INVOICE_PREFIX = "C460"  # the invoice is the attachment whose title starts with this (the last one)
INVOICE_TYPES = (".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif")
INVOICE_DPI = 150  # resolution of the invoice picture
# Without a C460 file: of the PDFs, the one whose title and text look most like the invoice -
# points for the sample's Assignment and amount, 1 per invoice word (each word counted once)
INVOICE_DATA_POINTS = 10
INVOICE_WORDS = ("invoice", "rechnung", "rg.", "vat", "ust", "mwst", "total", "gesamt")
# Without a C460 file: a link attachment (e.g. 'GTNE EVIDENCE') -> picture of its whole website
LINK_ICON = "@8S@"  # icon of a link in the attachment list
LINK_TITLE_WORD = "EVIDENCE"  # several links: the (last) one with this in its title, else the last link
DETAILS_BTN = "&DETAIL"  # 'Details' of the selected attachment - its row 'URL' holds the address
WEBSITE_WIDTH = 1280  # browser width in pixels for the website picture
STRIP_GAP = 20  # grey pixels between the strips of a long website page
FOLDER_DIALOG_TITLE = "Browse for Files or Folders"  # Windows dialog opened by 'Download'
# Folder for 'Download': SAP GUI's own folder (Documents\SAP\SAP GUI), the one the dialog suggests
SAP_DOWNLOAD_DIR = os.path.join(shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0), "SAP", "SAP GUI")

# Workflow report (company-specific, from the recording): one request per Samsung Document Number
WF_TCODE = "ZRFIZWF00001"
WF_KEY_FIELD = "wnd[0]/usr/txtSO_OBJKY-LOW"  # 'Object Key' = Samsung Document Number
WF_READY_BOX = "wnd[0]/usr/chkPA_READY"  # ticked = only some requests -> untick
WF_DATE_FIELD = "wnd[0]/usr/ctxtSO_CPUDT-LOW"  # created from
WF_LIST = "wnd[0]/shellcont/shell/shellcont/shell"
WF_HISTORY_BTN = "STATUS"  # opens 'Display Workflow History'

Sample = collections.namedtuple("Sample", "row number doc_no fiscal_year ss_doc_no assignment amount",
                                defaults=("", None))


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def _quarter(year, q):
    """First and last day of quarter q (1-4)."""
    return dt.date(year, 3 * q - 2, 1), dt.date(year + q // 4, 3 * q % 12 + 1, 1) - dt.timedelta(days=1)


def _last_completed_quarter(today=None):
    today = today or dt.date.today()
    q = (today.month - 1) // 3  # quarters completed this year - 0 = Q1 is currently running
    return _quarter(today.year, q) if q else _quarter(today.year - 1, 4)


def _parse_user_date(text):
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Invalid date: '{text}' (expected DD.MM.YYYY)")


def _ask_date(prompt, default):
    while True:
        raw = input(f"{prompt} [{default:%d.%m.%Y}]: ").strip()
        if not raw:
            return default
        try:
            return _parse_user_date(raw)
        except ValueError as e:
            print(e)


def _ask_period(today=None):
    """Small window to pick the posting period: a year, then one or more of its
    quarters next to each other (one period from the first to the last).
    Without a window (no tkinter) the dates are asked in the console."""
    today = today or dt.date.today()
    default_from, default_to = _last_completed_quarter(today)
    try:
        root = tk.Tk()
    except Exception:  # tk is None, or Windows can't show the window
        return _ask_date("Posting date from", default_from), _ask_date("Posting date to", default_to)
    root.title("Period")  # short - the window is narrow
    root.resizable(False, False)
    root.attributes("-topmost", True)  # in front of the console it was started from
    year = tk.IntVar(root, default_from.year)
    ticks = [tk.BooleanVar(root, q == (default_from.month + 2) // 3) for q in range(1, 5)]
    period = []

    def chosen():
        qs = [q for q, tick in enumerate(ticks, 1) if tick.get()]
        if not qs:
            raise ValueError("Please tick at least one quarter.")
        if qs != list(range(qs[0], qs[-1] + 1)):
            raise ValueError("Please tick quarters that follow each other.")
        return _quarter(year.get(), qs[0])[0], _quarter(year.get(), qs[-1])[1]

    def show():
        try:
            dates.config(text="= {:%d.%m.%Y} – {:%d.%m.%Y}".format(*chosen()))
        except ValueError as e:
            dates.config(text=str(e))

    def grey_out():
        for q, box in enumerate(boxes, 1):
            box.config(state="normal" if _quarter(year.get(), q)[1] < today else "disabled")  # not over yet
        show()

    def new_year(_):
        for tick in ticks:
            tick.set(False)
        grey_out()

    def ok():
        try:
            period.extend(chosen())
        except ValueError as e:
            messagebox.showerror("Invalid period", str(e), parent=root)
            return
        root.destroy()

    tk.Label(root, text="Year").grid(row=0, column=0, padx=8, pady=4, sticky="w")
    tk.OptionMenu(root, year, *range(today.year, today.year - 5, -1), command=new_year).grid(
        row=0, column=1, columnspan=3, pady=4, sticky="w")
    boxes = [tk.Checkbutton(root, text=f"Q{q}", variable=tick, command=show) for q, tick in enumerate(ticks, 1)]
    for col, box in enumerate(boxes):
        box.grid(row=1, column=col, padx=4)
    dates = tk.Label(root, width=44)  # wide enough for the longest message - the window keeps its size
    dates.grid(row=2, column=0, columnspan=4, padx=8)
    buttons = tk.Frame(root)
    buttons.grid(row=3, column=0, columnspan=4, pady=8)
    tk.Button(buttons, text="OK", width=8, command=ok).pack(side="left", padx=4)
    tk.Button(buttons, text="Cancel", width=8, command=root.destroy).pack(side="left", padx=4)
    root.bind("<Return>", lambda e: ok())
    root.bind("<Escape>", lambda e: root.destroy())
    grey_out()
    root.mainloop()
    if not period:
        raise RuntimeError("No period chosen - nothing done.")
    return tuple(period)


def _parse_sap_date(text):
    if not text:
        return None
    return dt.datetime.strptime(text, SAP_DATE_FORMAT).date()


def _parse_amount(text):
    """Converts SAP amounts like '1.234,56-', '1,234.56' or '-500' to float."""
    s = text.strip().replace(" ", "")
    if not s:
        return None
    negative = s.endswith("-") or s.startswith("-")
    s = s.strip("-")
    if "," in s and "." in s:
        dec = "," if s.rfind(",") > s.rfind(".") else "."
    elif s.count(",") == 1 and len(s.split(",")[1]) != 3:
        dec = ","
    elif s.count(".") == 1 and len(s.split(".")[1]) != 3:
        dec = "."
    else:
        dec = None  # thousands separators only (e.g. currency without decimals)
    if dec:
        thousands = "." if dec == "," else ","
        s = s.replace(thousands, "").replace(dec, ".")
    else:
        s = s.replace(",", "").replace(".", "")
    value = float(s)
    return -value if negative else value


def _text(value):
    """Cell value as text (whole numbers without '.0')."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return "" if value is None else str(value).strip()


def _year_of(value):
    """Year of a posting date - an Excel date, or text in SAP or German format."""
    if hasattr(value, "year"):
        return value.year
    for fmt in (SAP_DATE_FORMAT, "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(str(value).strip(), fmt).year
        except ValueError:
            pass
    return None


def _wait_until_idle(session, timeout=300):
    for _ in range(timeout * 2):
        if not session.Busy:
            return
        time.sleep(0.5)
    raise RuntimeError(f"SAP was still busy after {timeout}s.")


# ---------------------------------------------------------------------------
# SAP: run FBL4N and read the grid (Process 1)
# ---------------------------------------------------------------------------

def _fill_multiselect(session, button_id, values, mode="include"):
    """Fills the multiple selection dialog. mode='include' -> tab
    'Select Single Values', mode='exclude' -> 'Exclude Single Values'."""
    tab, subscreen, table_name, field = MULTI_TABS[mode]
    session.findById(button_id).press()
    session.findById(f"wnd[1]/usr/tabsTAB_STRIP/{tab}").select()
    table_id = f"wnd[1]/usr/tabsTAB_STRIP/{tab}/ssubSCREEN_HEADER:{subscreen}/{table_name}"
    table = session.findById(table_id)
    if len(values) > table.VisibleRowCount:
        raise RuntimeError(
            f"{len(values)} values do not fit into the multiple selection dialog "
            f"(max. {table.VisibleRowCount} visible rows)."
        )
    for i, value in enumerate(values):
        session.findById(f"{table_id}/ctxt{field}[1,{i}]").text = value
    session.findById("wnd[1]/tbar[0]/btn[8]").press()


def _ensure_dynamic_selections(session):
    """Shows the dynamic selections only if they are not open yet
    (the button toggles - pressing it twice would close them again)."""
    try:
        session.findById(DYN009_BTN)
    except Exception:
        session.findById(DYN_SEL_TOGGLE_BTN).press()
        session.findById(DYN009_BTN)  # raises if the field is still missing


def run_fbl4n(session, date_from, date_to):
    LOGGER.info(f"Starting FBL4N for company code {COMPANY_CODE}, posting date "
                f"{date_from:%d.%m.%Y} - {date_to:%d.%m.%Y} ...")
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL4N"
    session.findById("wnd[0]").sendVKey(0)

    _ensure_dynamic_selections(session)
    _fill_multiselect(session, DYN009_BTN, DYN_FIELD_EXCLUDE, mode="exclude")
    _fill_multiselect(session, GL_BTN, GL_ACCOUNTS, mode="include")

    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = COMPANY_CODE  # else SAP takes the one it remembers
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = date_from.strftime(SAP_DATE_FORMAT)
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = date_to.strftime(SAP_DATE_FORMAT)
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = SAP_LAYOUT
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    _wait_until_idle(session)

    try:
        return session.findById(GRID_ID)
    except Exception:
        msg = session.findById("wnd[0]/sbar").Text
        raise RuntimeError(f"FBL4N returned no list. SAP message: {msg or '(none)'}")


def read_grid(grid):
    """Reads all data rows in the column order of the layout."""
    columns = [grid.ColumnOrder(i) for i in range(grid.ColumnCount)]
    titles = [grid.GetDisplayedColumnTitle(c) for c in columns]
    row_count = grid.RowCount
    visible = max(int(grid.VisibleRowCount), 1)
    LOGGER.info(f"Grid: {row_count} rows, {len(columns)} columns.")

    rows = []
    for start in range(0, row_count, visible):
        grid.firstVisibleRow = start  # force loading, otherwise cells are empty
        for r in range(start, min(start + visible, row_count)):
            rows.append([str(grid.GetCellValue(r, c)).strip() for c in columns])
    return columns, titles, rows


def leave_sap(session, we_logged_in):
    if we_logged_in:
        sap_connector.sap_logout(session)
        return
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)
    except Exception as e:
        LOGGER.warning(f"Could not leave the SAP transaction: {e}")


# ---------------------------------------------------------------------------
# SAP: evidence pictures (Process 2)
# ---------------------------------------------------------------------------

def capture_sample(session, s, stem):
    """Document, workflow and invoice picture of one sample, in that order.
    A picture that fails is logged and left out - its slot in the sheet
    stays empty."""
    files = {}
    for kind, capture in (("document", _capture_document), ("workflow", _capture_workflow),
                          ("invoice", _capture_invoice)):
        try:
            files[kind] = capture(session, s, f"{stem}_{kind}")
        except Exception as e:
            LOGGER.warning(f"Sample {s.number} (document {s.doc_no}): no {kind} picture - {e}")
        _close_popups(session)
    LOGGER.info(f"Sample {s.number} (document {s.doc_no}): {', '.join(files) or 'no'} picture(s).")
    return files


def _open_fb03(session, s):
    """Opens the document of the sample in FB03 and returns the window title."""
    if not s.fiscal_year:
        raise RuntimeError(f"no posting date in '{SHEET_NAME}'")
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFB03"
    session.findById("wnd[0]").sendVKey(0)
    session.findById(FB03_DOC_FIELD).text = s.doc_no
    session.findById(FB03_COMPANY_FIELD).text = COMPANY_CODE
    session.findById(FB03_YEAR_FIELD).text = str(s.fiscal_year)
    session.findById("wnd[0]").sendVKey(0)
    _wait_until_idle(session)
    title = session.findById("wnd[0]").Text
    if "Data Entry View" not in title and "General Ledger View" not in title:
        msg = session.findById("wnd[0]/sbar").Text
        raise RuntimeError(f"FB03 did not open the document. SAP message: {msg or '(none)'}")
    return title


def _capture_document(session, s, stem):
    """Opens the document in FB03 ('Data Entry View') and saves a picture of it."""
    if "General Ledger View" in _open_fb03(session, s):
        session.findById(FB03_VIEW_BTN).press()
        _wait_until_idle(session)
    png = stem + ".png"
    session.findById("wnd[0]").HardCopy(png, 2)  # 2 = PNG
    if not os.path.exists(png):  # not isfile() - on the K: share it says False for files that exist
        raise RuntimeError("SAP did not save the picture")
    return png


def _capture_invoice(session, s, stem):
    """Opens the document in FB03 again and makes the invoice picture from its
    'File attachment' box (the downloaded file is kept):
    1. the (last) file whose title starts with C460 -> page 1
    2. else, of the PDFs, the one that looks most like the invoice -> page 1
    3. else the link -> picture of its whole website"""
    _open_fb03(session, s)
    session.findById(FB03_ATTACH_BTN).press()
    _wait_until_idle(session)
    grid = _attachment_list(session)
    rows = _attachments(grid)
    if not rows:
        raise RuntimeError("the attachment section is empty")
    row = _invoice_row(grid)
    if row is not None:
        return _first_page_as_png(_download(session, grid, row, stem), stem + ".png")
    names = {r: title or f"row {r + 1}" for r, title, _ in rows}
    # file type from the title; only a title without an extension needs the slow 'Details' popup
    types = {r: _extension(title) or _details(session, grid, r).get("Extension", "").upper()
             for r, title, _ in rows}
    pdfs = [(r, text) for r, _, text in rows if types[r] == "PDF"]
    if pdfs:
        files = [_download(session, grid, r, f"{stem}_{i}") for i, (r, _) in enumerate(pdfs, 1)]
        best, scores = _pick_invoice([text + " " + _pdf_text(f) for (_, text), f in zip(pdfs, files)],
                                     s.assignment, s.amount)
        name = names[pdfs[best][0]]
        if scores[best]:
            listing = ", ".join(f"'{names[r]}'={score}" for (r, _), score in zip(pdfs, scores))
            LOGGER.info(f"Sample {s.number}: no {INVOICE_PREFIX} file - took '{name}' of {len(pdfs)} PDF(s) "
                        f"(scores: {listing})")
        else:
            LOGGER.warning(f"Sample {s.number}: no {INVOICE_PREFIX} file - no PDF looked like an invoice, "
                           f"took the last one '{name}' - please check it")
        for i, f in enumerate(files):
            if i != best:
                os.remove(f)
        target = stem + os.path.splitext(files[best])[1]
        os.replace(files[best], target)
        return _first_page_as_png(target, stem + ".png")
    link = _link_row(grid)
    if link is None:
        found = ", ".join(f"'{names[r]}' ({types[r] or '?'})" for r in names)
        raise RuntimeError(f"no {INVOICE_PREFIX} file, no PDF and no link. Found: {found}")
    url = _details(session, grid, link).get("URL", "")
    if not url.startswith(("http://", "https://")):
        raise RuntimeError("the link has no web address in its 'Details'")
    LOGGER.info(f"Sample {s.number}: no {INVOICE_PREFIX} file and no PDF - picture of the link {url}")
    return _website_as_png(url, stem + ".png")


def _capture_workflow(session, s, stem):
    """Opens the request of the sample (Samsung Document Number) in the workflow
    report and saves a picture of its 'Display Workflow History' window."""
    if not (s.ss_doc_no and s.fiscal_year):
        raise RuntimeError(f"no Samsung Document Number or posting date in '{SHEET_NAME}'")
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n" + WF_TCODE
    session.findById("wnd[0]").sendVKey(0)
    session.findById(WF_KEY_FIELD).text = s.ss_doc_no
    session.findById(WF_READY_BOX).selected = False
    # the request is made before the posting - a year earlier is surely early enough
    session.findById(WF_DATE_FIELD).text = dt.date(s.fiscal_year - 1, 1, 1).strftime(SAP_DATE_FORMAT)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    _wait_until_idle(session)
    try:
        requests = session.findById(WF_LIST)
    except Exception:
        msg = session.findById("wnd[0]/sbar").Text
        raise RuntimeError(f"no workflow request found. SAP message: {msg or '(none)'}")
    requests.selectedRows = "0"
    requests.pressToolbarButton(WF_HISTORY_BTN)
    _wait_until_idle(session)
    window = session.ActiveWindow
    if "Workflow History" not in window.Text:
        msg = session.findById("wnd[0]/sbar").Text
        raise RuntimeError(f"'Display Workflow History' did not open. SAP message: {msg or '(none)'}")
    png = stem + ".png"
    window.HardCopy(png, 2)  # 2 = PNG
    if not os.path.exists(png):  # not isfile() - see _capture_document
        raise RuntimeError("SAP did not save the picture")
    return png


def _attachment_list(session):
    """ALV list of the 'File attachment' popup (the grid with a 'Title' column), scrolled
    through once so that every row is loaded - also the ones below the visible part."""
    if session.Children.Count > 1:
        for grid in _grids(session.findById("wnd[1]/usr")):
            if "Title" in _column_titles(grid):
                for start in range(0, grid.RowCount, max(int(grid.VisibleRowCount), 1)):
                    grid.firstVisibleRow = start  # like read_grid: rows not shown yet would read as empty
                return grid
    msg = session.findById("wnd[0]/sbar").Text
    raise RuntimeError(f"the 'File attachment' list did not open. SAP message: {msg or '(none)'}")


def _grids(container):
    for i in range(container.Children.Count):
        el = container.Children(i)
        if el.Type == "GuiShell" and el.SubType == "GridView":
            yield el
        elif el.ContainerType:
            yield from _grids(el)


def _column_titles(grid):
    """{displayed column title: column name} of an ALV grid."""
    names = [grid.ColumnOrder(i) for i in range(grid.ColumnCount)]
    return {grid.GetDisplayedColumnTitle(n): n for n in names}


def _attachments(grid):
    """[(row, title, all its cells as one text)] of the attachment list - without the
    yellow total row at the end (no icon, no title)."""
    cols = _column_titles(grid)
    rows = []
    for r in range(grid.RowCount):
        cells = {t: grid.GetCellValue(r, n).strip() for t, n in cols.items()}
        if cells["Icon"] or cells["Title"]:
            rows.append((r, cells["Title"], " ".join(cells.values())))
    return rows


def _extension(title):
    """File type at the end of an attachment title ('Anlage 1.pdf' -> 'PDF'), '' if there is none."""
    m = re.search(r"\.([A-Za-z][A-Za-z0-9]{1,4})$", title)  # not 'Rg. 4711' or 'No.12345'
    return m.group(1).upper() if m else ""


def _invoice_row(grid):
    """Row of the invoice: the (last) attachment whose title starts with INVOICE_PREFIX, None if there is none."""
    title = _column_titles(grid)["Title"]
    rows = [r for r in range(grid.RowCount) if grid.GetCellValue(r, title).strip().startswith(INVOICE_PREFIX)]
    return rows[-1] if rows else None


def _link_row(grid):
    """Row of the link to use instead: the (last) one with LINK_TITLE_WORD in its title,
    else the last link, None if there is none."""
    cols = _column_titles(grid)
    links = [r for r in range(grid.RowCount) if grid.GetCellValue(r, cols["Icon"]).strip() == LINK_ICON]
    evidence = [r for r in links if LINK_TITLE_WORD in grid.GetCellValue(r, cols["Title"]).upper()]
    return (evidence or links or [None])[-1]


def _pick_invoice(texts, assignment, amount):
    """Index of the text (attachment row + PDF text) that looks most like the sample's
    invoice, and the score of every text. A tie, also all 0 -> the later one."""
    # ponytail: plain text search - a scan without text only scores with its title, OCR if that is too weak
    patterns = [rf"(?<![A-Za-z0-9]){re.escape(assignment)}(?![A-Za-z0-9])"] if assignment else []
    if amount:
        en = f"{abs(amount):,.2f}"  # 1,234.56
        forms = {en, en.translate(str.maketrans(",.", ".,")), en.replace(",", ""), en.replace(",", "").replace(".", ",")}
        patterns.append(r"(?<![\d.,])(" + "|".join(map(re.escape, forms)) + r")(?!\d)")  # not inside 1.234,567
    words = re.compile(r"\b(" + "|".join(map(re.escape, INVOICE_WORDS)) + ")", re.IGNORECASE)
    scores = [INVOICE_DATA_POINTS * sum(bool(re.search(p, t)) for p in patterns)
              + len({w.lower() for w in words.findall(t)}) for t in texts]
    return max(range(len(texts)), key=lambda i: (scores[i], i)), scores


def _details(session, grid, row):
    """{field: value} of an attachment, read from its 'Details' popup (closed again)."""
    grid.selectedRows = str(row)
    grid.pressToolbarButton(DETAILS_BTN)
    _wait_until_idle(session)
    popup = session.ActiveWindow
    details = popup.findById("usr/cntlGRID/shellcont/shell")  # rows: field name | value
    fields = {details.GetCellValue(r, "COLUMNTEXT").strip(): details.GetCellValue(r, "VALUE").strip()
              for r in range(details.RowCount)}
    popup.close()
    _wait_until_idle(session)
    return fields


def _download(session, grid, row, stem):
    """Downloads one attachment and moves it to stem + its extension."""
    grid.selectedRows = str(row)
    download = _toolbar_button(grid, "Download")
    os.makedirs(SAP_DOWNLOAD_DIR, exist_ok=True)
    started = time.time()
    done = threading.Event()
    threading.Thread(target=_answer_folder_dialog, args=(SAP_DOWNLOAD_DIR, done), daemon=True).start()
    try:
        grid.pressToolbarButton(download)  # SAP waits here until the folder dialog is answered
    finally:
        done.set()
    _wait_until_idle(session)
    _close_popups(session, keep=1)  # the attachment list stays open for the next download
    src = _downloaded_file(SAP_DOWNLOAD_DIR, started)
    target = stem + os.path.splitext(src)[1].lower()
    shutil.move(src, target)
    return target


def _toolbar_button(grid, text):
    for i in range(grid.ToolbarButtonCount):
        if grid.GetToolbarButtonText(i).strip() == text:
            return grid.GetToolbarButtonId(i)
    raise RuntimeError(f"no '{text}' button in the attachment list")


def _answer_folder_dialog(folder, done):
    """Helper thread: 'Download' opens a Windows folder dialog and SAP waits
    until it is answered - type the folder and press OK."""
    while not done.wait(0.2):
        try:
            dialog = win32gui.FindWindow("#32770", FOLDER_DIALOG_TITLE)
            edit = win32gui.FindWindowEx(dialog, 0, "Edit", None)
            if dialog and edit:
                win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, folder)
                win32gui.SendMessage(win32gui.GetDlgItem(dialog, win32con.IDOK), win32con.BM_CLICK, 0, 0)
                return
        except win32gui.error:
            pass  # not open yet, or it changed while we typed - next try


def _downloaded_file(folder, since, timeout=10):
    """The newest PDF/picture SAP saved into folder after 'since'."""
    deadline = time.time() + timeout
    while True:
        new = [e.path for e in os.scandir(folder) if e.is_file()
               and e.name.lower().endswith(INVOICE_TYPES) and e.stat().st_mtime >= since - 2]
        if new:
            return max(new, key=os.path.getmtime)
        if time.time() > deadline:
            raise RuntimeError(f"no downloaded PDF or picture found in {folder}")
        time.sleep(0.5)


def _first_page_as_png(path, png):
    """Picture of the invoice: page 1 of a PDF, or the file itself if it is one."""
    if not path.lower().endswith(".pdf"):
        return path
    # opened from memory: with a path PyMuPDF checks isfile(), which fails on the K: share
    with open(path, "rb") as f, pymupdf.open(stream=f.read(), filetype="pdf") as pdf:
        pdf[0].get_pixmap(dpi=INVOICE_DPI).save(png)
    return png


def _pdf_text(path):
    """Text of all pages of a PDF - empty for a scan, or a PDF that cannot be read."""
    try:
        with open(path, "rb") as f, pymupdf.open(stream=f.read(), filetype="pdf") as pdf:  # see _first_page_as_png
            return " ".join(page.get_text() for page in pdf)
    except Exception:
        return ""  # e.g. password-protected - it then only scores with its title


def _website_as_png(url, png):
    """Picture of the whole web page, top to bottom, taken with a hidden Edge
    (the installed one - Playwright downloads no browser). The page is kept as
    <name>_page.png; the picture for Excel shows it in strips side by side."""
    page_png = os.path.splitext(png)[0] + "_page.png"
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge")
        page = browser.new_page(viewport={"width": WEBSITE_WIDTH, "height": 900})
        response = page.goto(url, wait_until="networkidle", timeout=60000)
        if response and not response.ok:  # no error page as evidence
            raise RuntimeError(f"the website answered with error {response.status}")
        page_bytes = page.screenshot(path=page_png, full_page=True)
        browser.close()
    # from memory: reading the page back from K: right away was refused once ('Permission denied')
    return _as_strips(page_bytes, png)


def _as_strips(picture, png):
    """The picture (PNG bytes) cut into equal strips, side by side with a grey gap - each
    strip no taller than a slot shows at READABLE_SCALE, so a long page stays readable."""
    src = pymupdf.Pixmap(picture)
    if src.alpha:
        src = pymupdf.Pixmap(src, 0)  # the grey gap must not be see-through
    slot_px = (SLOT_ROWS * EVIDENCE_ROW_HEIGHT - 2 * PICTURE_MARGIN) / READABLE_SCALE
    n = math.ceil(src.height / slot_px)
    part = math.ceil(src.height / n)
    out = pymupdf.Pixmap(src.colorspace, pymupdf.IRect(0, 0, n * src.width + (n - 1) * STRIP_GAP, part), False)
    out.clear_with(200)
    for i in range(n):
        x = i * (src.width + STRIP_GAP)
        src.set_origin(x, -i * part)  # strip i of the page now lies at x, top
        out.copy(src, pymupdf.IRect(x, 0, x + src.width, part))
    out.save(png)
    return png


def _close_popups(session, keep=0):
    """Closes leftover popups, so the next step starts on the main window
    (or on the first `keep` popups, e.g. the attachment list)."""
    for _ in range(5):
        if session.Children.Count < 2 + keep:
            return
        session.findById(f"wnd[{session.Children.Count - 1}]").close()
        _wait_until_idle(session)


# ---------------------------------------------------------------------------
# Prepare data
# ---------------------------------------------------------------------------

def _log_column_mapping(columns, titles):
    LOGGER.info("Column comparison SAP layout <-> template:")
    for i in range(max(len(columns), len(TEMPLATE_HEADERS))):
        sap = f"{columns[i]} ({titles[i]})" if i < len(columns) else "-"
        tpl = TEMPLATE_HEADERS[i] if i < len(TEMPLATE_HEADERS) else "-"
        col = chr(ord(DATA_FIRST_COL) + i) if i < len(TEMPLATE_HEADERS) else "?"
        LOGGER.info(f"  {col}: {tpl:<26} <- {sap}")


def build_excel_rows(raw_rows):
    key_idx = TEMPLATE_HEADERS.index(KEY_HEADER)
    result, skipped, warnings = [], 0, 0

    for raw in raw_rows:
        if not raw[key_idx]:
            skipped += 1  # total/empty row
            continue
        values = []
        for header, text in zip(TEMPLATE_HEADERS, raw):
            try:
                if header in DATE_HEADERS:
                    d = _parse_sap_date(text)
                    # real date -> Excel takes it as a date without the
                    # cell format having to be changed (sheet protection!). In UTC: without a time zone
                    # pywin32 takes it as local time and stores it 1-2 h early - the cell showed the day before
                    values.append(dt.datetime(d.year, d.month, d.day, tzinfo=dt.timezone.utc) if d else None)
                elif header in AMOUNT_HEADERS:
                    values.append(_parse_amount(text))
                else:
                    # apostrophe = store as text, leading zeros are kept
                    # (the apostrophe itself is not part of the value)
                    values.append("'" + text if text else None)
            except ValueError:
                warnings += 1
                values.append(text)  # keep raw value instead of losing the row
        result.append(tuple(values))

    if skipped:
        LOGGER.info(f"{skipped} row(s) without document number skipped (total rows).")
    if warnings:
        LOGGER.warning(f"{warnings} value(s) could not be converted and "
                       f"were kept as text - please check.")
    return result


def save_raw_export(columns, titles, rows, date_from, date_to):
    target_dir = os.path.join(BASE_DIR, RAW_EXPORT_DIR_NAME)
    os.makedirs(target_dir, exist_ok=True)
    name = (f"FBL4N_{date_from:%Y%m%d}-{date_to:%Y%m%d}_"
            f"{dt.datetime.now():%Y-%m-%d_%H%M%S}.csv")
    target = os.path.join(target_dir, name)
    with open(target, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(columns)
        writer.writerow(titles)
        writer.writerows(rows)
    LOGGER.info(f"Raw data saved: {target}")


# ---------------------------------------------------------------------------
# Excel template
# ---------------------------------------------------------------------------

def find_template():
    pattern = os.path.join(glob.escape(BASE_DIR), TEMPLATE_PATTERN)  # [SEG] in the path!
    matches = [p for p in glob.glob(pattern) if not os.path.basename(p).startswith("~$")]
    if len(matches) != 1:
        raise RuntimeError(f"Template not found unambiguously ({len(matches)} matches): {pattern}")
    return matches[0]


def ensure_not_open(path):
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        raise RuntimeError(
            f"'{os.path.basename(path)}' is currently open (possibly by another "
            f"person). Please close it and start again."
        )


def backup_template(path):
    target_dir = os.path.join(os.path.dirname(path), BACKUP_DIR_NAME)
    os.makedirs(target_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    target = os.path.join(target_dir, f"{stem}_{dt.datetime.now():%Y-%m-%d_%H%M%S}{ext}")
    shutil.copy2(path, target)
    LOGGER.info(f"Backup created: {target}")


@contextlib.contextmanager
def _open_template(path):
    """The template in an own hidden Excel - closed without saving on errors."""
    excel = win32.DispatchEx("Excel.Application")  # own instance, doesn't disturb open Excel
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False  # don't trigger Workbook_Open macros when opening
    wb = None
    try:
        wb = excel.Workbooks.Open(path, 0)  # 0 = don't update links
        if wb.ReadOnly:
            raise RuntimeError("Template was opened read-only - aborting.")
        yield excel, wb
    except Exception:
        LOGGER.error("Aborted: the template was NOT saved and is unchanged.")
        raise
    finally:
        if wb is not None:
            with contextlib.suppress(Exception):
                wb.Close(False)  # a failed Close must not skip Quit (orphan EXCEL.EXE)
        excel.Quit()


def _check_headers(ws, row=HEADER_ROW):
    rng = ws.Range(f"{DATA_FIRST_COL}{row}:{DATA_LAST_COL}{row}")
    found = ["" if v is None else str(v).strip() for v in rng.Value[0]]
    if found != TEMPLATE_HEADERS:
        diffs = [f"{chr(ord(DATA_FIRST_COL) + i)}: expected '{e}', found '{f}'"
                 for i, (e, f) in enumerate(zip(TEMPLATE_HEADERS, found)) if e != f]
        raise RuntimeError(f"Headers in '{ws.Name}' row {row} do not match:\n  " + "\n  ".join(diffs))


def _last_data_row(ws):
    """End of the contiguous data block directly below D27:X27.
    The first completely empty row in D:X ends the block - anything below it
    (e.g. totals or notes) is never cleared. The search is also limited
    to the prepared area (max. MAX_ROWS rows)."""
    max_row = FIRST_DATA_ROW + MAX_ROWS - 1
    values = ws.Range(f"{DATA_FIRST_COL}{FIRST_DATA_ROW}:{DATA_LAST_COL}{max_row}").Value
    last = HEADER_ROW
    for offset, row in enumerate(values):
        if all(v is None or str(v).strip() == "" for v in row):
            break
        last = FIRST_DATA_ROW + offset
    return last


def _selected_rows(ws):
    """The samples marked 'Selected' in Population, in sheet order."""
    last = _last_data_row(ws)
    if last < FIRST_DATA_ROW:
        return []
    values = ws.Range(f"{SAMPLE_COL}{FIRST_DATA_ROW}:{DATA_LAST_COL}{last}").Value
    # positions in a B:X row: B=0, C=1, D=2 (first template header) ...
    doc, date, ss_doc, assignment, amount = (
        2 + TEMPLATE_HEADERS.index(h)
        for h in (KEY_HEADER, "Posting Date", "Samsung Document Number", "Assignment", "Amount in Local Currency"))
    return [Sample(FIRST_DATA_ROW + i, int(v[1]), _text(v[doc]), _year_of(v[date]), _text(v[ss_doc]),
                   _text(v[assignment]), v[amount] if isinstance(v[amount], (int, float)) else None)
            for i, v in enumerate(values) if str(v[0] or "").strip() == SELECTED_VALUE]


def _reset_filter(ws):
    """Remove the filter from the last run COMPLETELY (not just show all
    rows), so the workbook looks like before the first run when the macro
    starts. The filter on 'Selected' is only set again at the very end."""
    try:
        if ws.FilterMode:
            ws.ShowAllData()
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
            LOGGER.info("Filter from the last run removed in 'Population'.")
    except Exception as e:
        raise RuntimeError(
            f"Filter from the last run could not be removed ({e}). "
            f"Please remove it manually in the template and start again."
        )


def _remove_filter_names(wb):
    """Deletes orphaned '_FilterDatabase' names. Excel leaves them behind after
    a filter is removed; if the sample macro then copies a sheet, the popup
    'Name conflict - name cannot be the same as a built-in name' appears.
    Names of sheets with an active filter are kept."""
    removed = 0
    for i in range(wb.Names.Count, 0, -1):  # backwards, since we delete
        name = wb.Names(i)
        if not name.Name.endswith("_FilterDatabase"):
            continue
        parent = name.Parent
        sheet_has_filter = False
        try:
            sheet_has_filter = parent.Name != wb.Name and bool(parent.AutoFilterMode)
        except Exception:
            pass  # workbook-level name -> always orphaned
        if sheet_has_filter:
            continue
        try:
            name.Delete()
            removed += 1
        except Exception as e:
            LOGGER.warning(f"Name '{name.Name}' could not be deleted: {e}")
    if removed:
        LOGGER.info(f"{removed} orphaned '_FilterDatabase' name(s) removed.")


def _find_button_macro(wb):
    """Determines the macro behind the button located on MAIN_BUTTON_CELL."""
    if MAIN_BUTTON_MACRO:
        return MAIN_BUTTON_MACRO
    ws = wb.Worksheets(MAIN_SHEET_NAME)
    cell = ws.Range(MAIN_BUTTON_CELL)
    for i in range(1, ws.Shapes.Count + 1):
        shp = ws.Shapes(i)
        tl, br = shp.TopLeftCell, shp.BottomRightCell
        if not (tl.Row <= cell.Row <= br.Row and tl.Column <= cell.Column <= br.Column):
            continue
        if shp.Type == MSO_OLE_CONTROL:  # ActiveX: click event in the sheet module
            return f"{ws.CodeName}.{shp.Name}_Click"
        if shp.OnAction:  # form control / shape with macro
            return shp.OnAction
    raise RuntimeError(
        f"No button with a macro found on {MAIN_SHEET_NAME}!{MAIN_BUTTON_CELL}. "
        f"If needed, enter the macro name in MAIN_BUTTON_MACRO."
    )


def run_sample_button(excel, wb):
    macro = _find_button_macro(wb)
    if "!" not in macro:
        macro = f"'{wb.Name}'!{macro}"
    LOGGER.info(f"Starting sample selection (macro: {macro}) ...")
    wb.Worksheets(MAIN_SHEET_NAME).Activate()
    # Like a real click: events on, Excel visible - if the macro shows a
    # message, it can be clicked away instead of the tool hanging.
    excel.EnableEvents = True
    excel.ScreenUpdating = True
    excel.Visible = True
    excel.Run(macro)
    excel.Calculate()


def filter_selected(ws, n_rows):
    """Filters 'Sample Selected ?' on 'Selected'. An error here is not
    critical - the samples have still been drawn and saved."""
    last = FIRST_DATA_ROW + n_rows - 1
    try:
        ws.Activate()
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False  # remove old filter range (may have a different length)
        rng = ws.Range(f"{SAMPLE_COL}{HEADER_ROW}:{DATA_LAST_COL}{last}")
        rng.AutoFilter(1, SELECTED_VALUE)  # field 1 = column B
        LOGGER.info(f"'Population' filtered on '{SELECTED_VALUE}'.")
    except Exception as e:
        LOGGER.warning(f"Filter could not be set ({e}) - please filter manually.")


def process_template(path, rows):
    """Process 1 in Excel: fill Population, draw the samples, filter, save."""
    with _open_template(path) as (excel, wb):
        ws = wb.Worksheets(SHEET_NAME)
        _check_headers(ws)
        _reset_filter(ws)
        _remove_filter_names(wb)

        # 1. Clear and fill Population (only D28:X)
        last = _last_data_row(ws)
        if last >= FIRST_DATA_ROW:
            ws.Range(f"{DATA_FIRST_COL}{FIRST_DATA_ROW}:{DATA_LAST_COL}{last}").ClearContents()
            LOGGER.info(f"Old data in {DATA_FIRST_COL}:{DATA_LAST_COL} cleared "
                        f"(rows {FIRST_DATA_ROW}-{last}).")
        last_row = FIRST_DATA_ROW + len(rows) - 1
        ws.Range(f"{DATA_FIRST_COL}{FIRST_DATA_ROW}:{DATA_LAST_COL}{last_row}").Value = tuple(rows)
        LOGGER.info(f"{len(rows)} rows written to '{SHEET_NAME}'.")

        # 2. Draw samples (button in Main), from the real row count in C4
        wb.Worksheets(MAIN_SHEET_NAME).Range(MAIN_COUNT_CELL).Value = len(rows)
        LOGGER.info(f"{MAIN_SHEET_NAME}!{MAIN_COUNT_CELL} (Quantity of Population) set to {len(rows)}.")
        run_sample_button(excel, wb)
        numbers = [s.number for s in _selected_rows(ws)]
        if numbers:
            LOGGER.info(f"{len(numbers)} sample(s) drawn, numbers: {', '.join(map(str, numbers))}")
        else:
            LOGGER.warning("After the macro no row is marked as 'Selected' - please check.")

        # 3. Filter on Selected and save
        filter_selected(ws, len(rows))
        wb.Save()
        LOGGER.info("Template saved.")


# ---------------------------------------------------------------------------
# Excel: Evidences sheet (Process 2)
# ---------------------------------------------------------------------------

def _label_rows(n_samples):
    """Rows of the red labels: below the table, then one block every BLOCK_ROWS rows."""
    first = EVIDENCES_FIRST_ROW + n_samples + EVIDENCES_GAP_ROWS
    return [first + i * BLOCK_ROWS for i in range(n_samples)]


def _clear_evidences(ws):
    """Removes the previous evidence below the header row (pictures and cells);
    rows 1-10 with the instructions and the header stay as they are."""
    removed = 0
    for i in range(ws.Shapes.Count, 0, -1):  # backwards, since we delete
        shp = ws.Shapes(i)
        if shp.Type in MSO_PICTURES and shp.TopLeftCell.Row > EVIDENCES_HEADER_ROW:
            shp.Delete()
            removed += 1
    used = ws.UsedRange
    last = used.Row + used.Rows.Count - 1
    if last >= EVIDENCES_FIRST_ROW:
        ws.Range(f"{EVIDENCES_FIRST_ROW}:{last}").Clear()
    LOGGER.info(f"Previous evidence removed from '{ws.Name}' ({removed} picture(s)).")


def _write_label(ws, row, s):
    """Red label of a block: sample number in C, document number in E."""
    ws.Range(f"C{row}").Value = s.number
    ws.Range(f"E{row}").Value = "'" + s.doc_no
    ws.Range(f"C{row}:E{row}").Interior.Color = LABEL_COLOR
    number = ws.Range(f"C{row}")
    number.HorizontalAlignment = XL_CENTER
    number.Borders.LineStyle = XL_CONTINUOUS


def _insert_picture(ws, path, top_row, slot_left, slot_width):
    """Embeds the picture in its slot - slot_width points from slot_left, SLOT_ROWS
    rows from top_row - as large as fits, proportions kept."""
    left = slot_left + PICTURE_MARGIN
    right = slot_left + slot_width - PICTURE_MARGIN
    top = ws.Rows(top_row).Top + PICTURE_MARGIN
    bottom = ws.Rows(top_row + SLOT_ROWS).Top - PICTURE_MARGIN
    picture = ws.Shapes.AddPicture(path, False, True, left, top, -1, -1)  # embedded, not linked
    picture.LockAspectRatio = True
    picture.Width = picture.Width * min((right - left) / picture.Width, (bottom - top) / picture.Height, 1)
    picture.Placement = XL_MOVE


def fill_evidences(pop, ev, samples, pictures):
    """Evidences sheet like the manual version: the selected rows under the
    header, then one block per sample - red label and pictures side by side."""
    _clear_evidences(ev)
    last_row = _label_rows(len(samples))[-1] + BLOCK_ROWS - 1
    ev.Range(f"{EVIDENCES_FIRST_ROW}:{last_row}").RowHeight = EVIDENCE_ROW_HEIGHT  # before any position is taken
    for i, s in enumerate(samples):
        row = EVIDENCES_FIRST_ROW + i
        pop.Range(f"{SAMPLE_COL}{s.row}:{DATA_LAST_COL}{s.row}").Copy(ev.Range(f"{SAMPLE_COL}{row}"))
        result = ev.Range(f"{SAMPLE_COL}{row}:{NUMBER_COL}{row}")
        result.Value = result.Value  # B is a formula in Population - keep only its result
    # Pictures only while Evidences is the active sheet: on another one, Excel stores them under rows
    # rounded to the zoom of Evidences (115% in the template) - they slid ~9pt higher per block
    was_active = ev.Parent.ActiveSheet
    ev.Activate()
    for s, files, label in zip(samples, pictures, _label_rows(len(samples))):
        _write_label(ev, label, s)
        left = ev.Range("C1").Left
        for kind, width in EVIDENCE_SLOTS.items():
            if kind in files:  # a missing picture leaves its slot empty
                try:
                    _insert_picture(ev, files[kind], label + 1, left, width)
                except Exception as e:
                    LOGGER.warning(f"Sample {s.number}: {kind} picture could not be inserted - {e}")
            left += width
    was_active.Activate()
    LOGGER.info(f"'{ev.Name}' filled with {len(samples)} sample(s).")


def run_evidences(session, path):
    """Process 2: SAP pictures for every 'Selected' sample, then the Evidences sheet."""
    run_dir = os.path.join(os.path.dirname(path), EVIDENCE_DIR_NAME, f"{dt.datetime.now():%Y-%m-%d_%H%M%S}")
    with _open_template(path) as (_, wb):
        pop = wb.Worksheets(SHEET_NAME)
        ev = wb.Worksheets(EVIDENCES_SHEET_NAME)
        _check_headers(pop)  # _selected_rows reads doc no / dates by column position
        _check_headers(ev, EVIDENCES_HEADER_ROW)
        samples = _selected_rows(pop)
        if not samples:
            raise RuntimeError(f"No row in '{SHEET_NAME}' is marked '{SELECTED_VALUE}' - draw the samples first.")
        os.makedirs(run_dir)
        LOGGER.info(f"Taking SAP pictures for {len(samples)} sample(s) - please don't use SAP meanwhile ...")
        session.findById("wnd[0]").maximize()
        _close_popups(session)
        pictures = [capture_sample(session, s, os.path.join(run_dir, f"{s.number:03d}_{s.doc_no}"))
                    for s in samples]
        fill_evidences(pop, ev, samples, pictures)
        wb.Save()
    LOGGER.info(f"'{EVIDENCES_SHEET_NAME}' saved. Pictures and invoice files: {run_dir}")
    missing = [f"{s.number} ({kind})" for s, files in zip(samples, pictures)
               for kind in EVIDENCE_SLOTS if kind not in files]
    if missing:
        LOGGER.warning("Pictures missing - please add them manually: sample " + ", ".join(missing))


# ---------------------------------------------------------------------------
# Main program
# ---------------------------------------------------------------------------

def _template(args):
    """The workbook to work on (--template or the one on K:) - must not be open elsewhere."""
    path = os.path.abspath(args.template) if args.template else find_template()
    ensure_not_open(path)
    return path


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.StreamHandler(), logging.FileHandler(LOG_FILE, "a", "utf-8")])
    parser = argparse.ArgumentParser(description="Population + Evidences tool (FBL4N -> GA-02-006_SEG)")
    parser.add_argument("--from", dest="date_from", help="Posting date from (DD.MM.YYYY)")
    parser.add_argument("--to", dest="date_to", help="Posting date to (DD.MM.YYYY)")
    parser.add_argument("--columns", action="store_true",
                        help="Only show SAP columns, write nothing to Excel")
    parser.add_argument("--evidences-only", action="store_true",
                        help="Only Process 2 (Evidences) for the samples already drawn")
    parser.add_argument("--template", help="Work on this workbook instead of the one on K: (e.g. a test copy)")
    args = parser.parse_args()
    LOGGER.info(f"Population + Evidences tool, version {VERSION}")

    if args.evidences_only:
        template = _template(args)
        backup_template(template)
        session, we_logged_in = sap_connector.get_sap_session()
        try:
            run_evidences(session, template)
        finally:
            leave_sap(session, we_logged_in)
        LOGGER.info("Done.")
        return

    if args.date_from or args.date_to:  # a date not given: from/to of the last completed quarter
        default_from, default_to = _last_completed_quarter()
        date_from = _parse_user_date(args.date_from) if args.date_from else default_from
        date_to = _parse_user_date(args.date_to) if args.date_to else default_to
    else:
        date_from, date_to = _ask_period()
    if date_from > date_to:
        raise RuntimeError("'from' is after 'to'.")

    template = None if args.columns else _template(args)  # check it before the SAP work, not afterwards

    session, we_logged_in = sap_connector.get_sap_session()
    try:
        grid = run_fbl4n(session, date_from, date_to)
        columns, titles, raw_rows = read_grid(grid)
        _log_column_mapping(columns, titles)
        if args.columns:
            return

        if len(columns) != len(TEMPLATE_HEADERS):
            raise RuntimeError(
                f"Layout {SAP_LAYOUT} has {len(columns)} columns, the template expects "
                f"{len(TEMPLATE_HEADERS)}. Please adjust the layout (see column comparison)."
            )

        if SAVE_RAW_EXPORT:
            save_raw_export(columns, titles, raw_rows, date_from, date_to)

        rows = build_excel_rows(raw_rows)
        if not rows:
            raise RuntimeError("No data rows found - Population will NOT be cleared.")
        if len(rows) > MAX_ROWS:
            raise RuntimeError(
                f"{len(rows)} rows found, but the template is only prepared for {MAX_ROWS} "
                f"rows (numbers/formulas in B and C). Please shorten the period."
            )

        ensure_not_open(template)  # someone could have opened it in the meantime
        backup_template(template)
        process_template(template, rows)
        try:
            run_evidences(session, template)
        except Exception:
            LOGGER.error("Process 1 is saved. After fixing the problem below, run again with "
                         "--evidences-only - the samples stay the same.")
            raise
    finally:
        leave_sap(session, we_logged_in)
    LOGGER.info("Done.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        LOGGER.error(str(exc))
        sys.exit(1)
