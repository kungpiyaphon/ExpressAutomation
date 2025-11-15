#!/usr/bin/env python3
"""
tools/export_watcher_converter.py

Watch incoming_exports/ for .xls/.xlsx files, convert each to express template format
and save to excel_templates/{COMPANY}-{YEAR}-{SUFFIX}.xlsx after user selects company (EDS/FIX).

Run:
    python tools/export_watcher_converter.py
"""

import os
import time
import shutil
from pathlib import Path
import threading
import tkinter as tk
from tkinter import simpledialog, messagebox
from datetime import datetime

import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ---------------------------
# Config
# ---------------------------
PROJECT_ROOT = Path(__file__).resolve().parents[1]
INCOMING = PROJECT_ROOT / "incoming_exports"
INCOMING.mkdir(parents=True, exist_ok=True)

INCOMING_PROCESSED = INCOMING / "processed"
INCOMING_PROCESSED.mkdir(parents=True, exist_ok=True)

TEMPLATE_FOLDER = PROJECT_ROOT / "excel_templates"
TEMPLATE_FOLDER.mkdir(parents=True, exist_ok=True)

TEMPLATE_NAME_DEFAULT = "express_import_template.xlsx"  # fallback

# Mapping for Ship-to-Branch-Code -> Dept
BRANCH_MAP = {
    "0002198490": "BKK",
    "0006093962": "FPR",
    "0005785271": "TMB",
    "0002266232": "CSP",
    "0004374861": "RYY",
}

# Template columns and fixed values
TEMPLATE_COLUMNS = ["Dept", "Date", "Supplier", "Invoice", "Code", "Qty", "UnitCost"]
SUPPLIER_FIXED = "026959000"
CODE_FIXED = "001"
QTY_FIXED = 1

# Watcher behavior
MIN_STABLE_SECONDS = 1.0  # wait for file size to stabilize
CHECK_INTERVAL = 0.5

# ---------------------------
# Utilities
# ---------------------------
def wait_file_ready(path: Path, timeout=20.0) -> bool:
    """Wait until file size is stable for MIN_STABLE_SECONDS within timeout."""
    start = time.time()
    last_size = -1
    stable_since = None
    while time.time() - start < timeout:
        if not path.exists():
            return False
        try:
            size = path.stat().st_size
        except Exception:
            size = -1
        now = time.time()
        if size == last_size:
            if stable_since is None:
                stable_since = now
            elif now - stable_since >= MIN_STABLE_SECONDS:
                return True
        else:
            stable_since = None
        last_size = size
        time.sleep(CHECK_INTERVAL)
    return False

def parse_yyyymmdd_to_ddmmyy(s):
    s = str(s).strip()
    if not s:
        return ""
    # Expect either numeric yyyymmdd (8 digits) or other parseable format
    if s.isdigit() and len(s) == 8:
        try:
            dt = datetime.strptime(s, "%Y%m%d")
            return dt.strftime("%d/%m/%y")
        except Exception:
            pass
    # fallback: try pandas parse
    try:
        dt = pd.to_datetime(s, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%d/%m/%y")
    except Exception:
        pass
    # else return original
    return s

def map_row_to_template(row):
    """
    row: pandas Series with expected input columns, including:
     - 'Ship-to-Branch-Code'
     - 'Invoice Date'  (yyyymmdd)
     - 'Local Invoice No' (or 'Invoice No')
     - 'Amount'
    """
    code = str(row.get("Ship-to-Branch-Code", "")).strip()
    dept = BRANCH_MAP.get(code, "")
    invoice_date = row.get("Invoice Date", "")
    date_out = parse_yyyymmdd_to_ddmmyy(invoice_date)
    local_invoice = row.get("Local Invoice No", "") or row.get("Invoice No", "")
    amount = row.get("Amount", "")
    # normalize unitcost
    try:
        unitcost = float(str(amount).replace(",", "")) if amount != "" else ""
    except Exception:
        unitcost = amount
    return {
        "Dept": dept,
        "Date": date_out,
        "Supplier": SUPPLIER_FIXED,
        "Invoice": str(local_invoice),
        "Code": CODE_FIXED,
        "Qty": QTY_FIXED,
        "UnitCost": unitcost,
    }

# ---------------------------
# Robust sheet reader
# ---------------------------
def read_sheet_from_file(input_path: Path):
    """
    Read sheet named 'input' if exists; else read first sheet.
    Supports:
      - real .xls/.xlsx (via pandas.read_excel)
      - HTML-based Excel (sheet001.htm or .xls that is HTML) via pandas.read_html
      - if input_path is a directory (like LINE download), find .htm/.xls/.xlsx inside
    """
    # If user passed a directory (e.g., extracted from LINE), try to find a usable file inside
    if input_path.is_dir():
        # prefer sheet001.htm or first .htm, then xlsx/xls
        for candidate_name in ("sheet001.htm", "sheet001.html"):
            candidate = input_path / candidate_name
            if candidate.exists():
                input_path = candidate
                break
        else:
            # find any .htm/.html
            htm = next(input_path.glob("*.htm"), None) or next(input_path.glob("*.html"), None)
            if htm:
                input_path = htm
            else:
                # find any xls/xlsx
                excel = next(input_path.glob("*.xls"), None) or next(input_path.glob("*.xlsx"), None)
                if excel:
                    input_path = excel
                else:
                    raise FileNotFoundError(f"No usable sheet/html/xls found inside folder: {input_path}")

    suffix = input_path.suffix.lower()

    # Quick content sniff: check starting bytes to see if file is HTML
    try:
        with input_path.open("rb") as f:
            start = f.read(512)
        start_text = start.lstrip()[:10].lower()
        looks_like_html = b"<html" in start.lower() or start_text.startswith(b'<!doctype') or start_text.startswith(b'<html')
    except Exception:
        looks_like_html = False

    # If it's an HTML file (either .htm/.html or .xls that contains HTML), try read_html
    if suffix in (".htm", ".html") or looks_like_html:
        try:
            # pandas.read_html returns list of dataframes - take first
            tables = pd.read_html(input_path, header=0)
            if not tables:
                raise ValueError("No tables found in HTML")
            df = tables[0]
            # normalize headers
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            raise RuntimeError(f"Failed to parse HTML table from {input_path}: {e}")

    # Else, assume binary excel; try read_excel with best-effort engine selection
    try:
        # allow pandas to pick engine for xlsx; for xls prefer xlrd if available
        if suffix == ".xls":
            df_dict = pd.read_excel(input_path, sheet_name=None, engine="xlrd", dtype=str)
        else:
            df_dict = pd.read_excel(input_path, sheet_name=None, dtype=str)
        if "input" in df_dict:
            return df_dict["input"]
        return next(iter(df_dict.values()))
    except Exception as e:
        # If read_excel failed, but file content looked like HTML, we already tried above.
        # Provide actionable message for the user.
        raise RuntimeError(f"Failed reading Excel file {input_path}: {e}\n"
                           f"Hint: file may be HTML export or corrupt. Try opening in Excel and Save As .xlsx, "
                           f"or provide the extracted sheet (sheet001.htm).")

# ---------------------------
# Convert + write
# ---------------------------
def convert_and_write(input_path: Path, company_choice: str, year: str, suffix_tag: str) -> Path:
    # read
    df_in = read_sheet_from_file(input_path)
    # normalize headers (strip)
    df_in.columns = [str(c).strip() for c in df_in.columns]

    # Map rows
    rows = []
    for _, r in df_in.iterrows():
        mapped = map_row_to_template(r)
        rows.append(mapped)

    out_df = pd.DataFrame(rows, columns=TEMPLATE_COLUMNS)

    # assemble filename
    filename = f"{company_choice}-{year}-{suffix_tag}.xlsx"
    target_path = TEMPLATE_FOLDER / filename

    # atomic write: write to tmp then replace
    tmp = target_path.with_suffix(target_path.suffix + ".tmp")
    out_df.to_excel(tmp, index=False, engine="openpyxl")
    # ensure written then move
    tmp.replace(target_path)
    return target_path

# ---------------------------
# Simple GUI: sequential dialogs (more robust across threads)
# ---------------------------
def ask_user_choose_company(default_year=None):
    """
    Use simpledialog sequentially to collect:
      - company (EDS or FIX)
      - year (YYYY)
      - suffix (e.g. RR)
    Returns dict or None if cancelled.
    """
    root = tk.Tk()
    root.withdraw()

    # Company
    while True:
        company = simpledialog.askstring("Company", "Enter company (EDS or FIX):", initialvalue="EDS", parent=root)
        if company is None:
            root.destroy()
            return None
        company = company.strip().upper()
        if company in ("EDS", "FIX"):
            break
        messagebox.showerror("Invalid", "Please enter EDS or FIX.", parent=root)

    # Year
    default_y = str(default_year) if default_year else str(datetime.now().year)
    while True:
        year = simpledialog.askstring("Year", "Enter year (YYYY):", initialvalue=default_y, parent=root)
        if year is None:
            root.destroy()
            return None
        year = year.strip()
        if year.isdigit() and len(year) == 4:
            break
        messagebox.showerror("Invalid", "Please enter 4-digit year, e.g. 2025", parent=root)

    # Suffix
    suffix = simpledialog.askstring("Suffix", "Enter suffix (e.g. RR):", initialvalue="RR", parent=root)
    if suffix is None:
        root.destroy()
        return None
    suffix = suffix.strip() or "RR"

    root.destroy()
    return {"company": company, "year": year, "suffix": suffix}

# ---------------------------
# Watcher handler
# ---------------------------
class ExportHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self._lock = threading.Lock()

    def _process(self, src_path: str, event_name: str):
        path = Path(src_path)
        if not path.exists():
            return
        if path.is_dir():
            # if folder, attempt to handle contents (e.g., LINE zip extraction)
            # but we pass the folder to read_sheet_from_file which will search inside
            pass
        if path.is_dir() and not any(path.glob("*")):
            # empty folder, skip
            return

        if path.suffix.lower() not in (".xls", ".xlsx", ".htm", ".html") and not path.is_dir():
            print(f"[SKIP] Not an Excel/HTML file or folder: {path.name}")
            return

        # avoid concurrent processing
        if not self._lock.acquire(blocking=False):
            print("[SKIP] Converter busy; skipping event.")
            return

        try:
            print(f"[EVENT:{event_name}] Detected: {path.name}")
            ready = wait_file_ready(path, timeout=30.0)
            if not ready:
                print(f"[WARN] File not stable/ready: {path}")
                return

            # Quick try to read a small sample to validate
            try:
                df_sample = read_sheet_from_file(path)
            except Exception as e:
                print(f"[ERROR] Failed reading file {path}: {e}")
                return

            # Prompt user for company/year/suffix (sequential dialogs)
            choice = ask_user_choose_company(default_year=datetime.now().year)
            if not choice:
                print("[INFO] User cancelled conversion.")
                return

            company = choice['company']
            year = choice['year']
            suffix_tag = choice['suffix']

            # Convert and write
            try:
                out = convert_and_write(path, company, year, suffix_tag)
                print(f"[DONE] Converted to template: {out}")
            except Exception as e:
                print(f"[ERROR] Conversion failed: {e}")
                return

            # move original to processed
            try:
                dest = INCOMING_PROCESSED / path.name
                if dest.exists():
                    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
                    dest = INCOMING_PROCESSED / f"{path.stem}-{ts}{path.suffix}"
                shutil.move(str(path), str(dest))
                print(f"[INFO] Moved original to: {dest}")
            except Exception as e:
                print(f"[WARN] Could not move original: {e}")

        finally:
            self._lock.release()

    def on_created(self, event):
        if not event.is_directory:
            self._process(event.src_path, "created")
        else:
            # folder created - process folder too
            self._process(event.src_path, "created")

    def on_moved(self, event):
        if not event.is_directory:
            self._process(event.dest_path, "moved")
        else:
            self._process(event.dest_path, "moved")

# ---------------------------
# Main runner
# ---------------------------
def main():
    print(f"[WATCHING] {INCOMING}")
    observer = Observer()
    handler = ExportHandler()
    observer.schedule(handler, str(INCOMING), recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        print("Stopping watcher...")
        observer.stop()
    observer.join()

if __name__ == "__main__":
    main()
