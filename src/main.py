#!/usr/bin/env python3
"""
src/main.py

Single-process automation service:
 - Watches incoming_exports/ -> converts export -> writes template to excel_templates/
 - Immediately runs the Express workflow (via express_launcher.run_full_workflow)
 - Also watches excel_templates/ to handle templates placed manually (fallback)
 - Robust: tmp->move atomic write, waits for file ready, avoids double runs with RUN_LOCK

Usage:
    python src/main.py

Environment:
    NO_GUI=1    -> run in headless mode (no tkinter dialogs; uses defaults)
"""
from __future__ import annotations
import os
import sys
import time
import json
import shutil
import logging
import threading
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict

import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# try to import express workflow
try:
    from express_launcher import run_full_workflow
except Exception as e:
    run_full_workflow = None
    # we'll log later; ok for development if separate

# ---------------------------
# Config
# ---------------------------
ROOT = Path(__file__).resolve().parents[1]  # project root
INCOMING = ROOT / "incoming_exports"
INCOMING_PROCESSED = INCOMING / "processed"
TEMPLATE_DIR = ROOT / "excel_templates"
TEMPLATE_PROCESSED = TEMPLATE_DIR / "processed"
LOG_DIR = ROOT / "logs"

for d in (INCOMING, INCOMING_PROCESSED, TEMPLATE_DIR, TEMPLATE_PROCESSED, LOG_DIR):
    d.mkdir(parents=True, exist_ok=True)

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

# File behavior
MIN_STABLE_SECONDS = 1.0
CHECK_INTERVAL = 0.4

# Regex-ish filename expectation
import re
RE_FILENAME = re.compile(r"^([A-Za-z]+)-(\d{4})(?:-[A-Za-z0-9._-]+)?$", re.IGNORECASE)

# Headless?
HEADLESS = os.getenv("NO_GUI", "0") in ("1", "true", "True")

# ---------------------------
# Logging
# ---------------------------
LOG_FILE = LOG_DIR / "automation.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger("automation")

# ---------------------------
# Helpers
# ---------------------------
def wait_file_ready(path: Path, timeout: float = 20.0) -> bool:
    start = time.time()
    last = -1
    stable_since = None
    while time.time() - start < timeout:
        if not path.exists():
            return False
        try:
            size = path.stat().st_size
        except Exception:
            size = -1
        now = time.time()
        if size == last:
            if stable_since is None:
                stable_since = now
            elif now - stable_since >= MIN_STABLE_SECONDS:
                return True
        else:
            stable_since = None
        last = size
        time.sleep(CHECK_INTERVAL)
    return False

def parse_yyyymmdd_to_ddmmyy(s) -> str:
    s = str(s).strip()
    if not s:
        return ""
    if s.isdigit() and len(s) == 8:
        try:
            dt = datetime.strptime(s, "%Y%m%d")
            return dt.strftime("%d/%m/%y")
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%d/%m/%y")
    except Exception:
        pass
    return s

def parse_filename_search_key(name_no_ext: str) -> Optional[str]:
    m = RE_FILENAME.match(name_no_ext)
    if not m:
        return None
    return f"{m.group(1).upper()}{m.group(2)}"

def map_row_to_template(row: pd.Series) -> Dict[str, object]:
    code = str(row.get("Ship-to-Branch-Code", "")).strip()
    dept = BRANCH_MAP.get(code, "")
    invoice_date = row.get("Invoice Date", "")
    date_out = parse_yyyymmdd_to_ddmmyy(invoice_date)
    local_invoice = row.get("Local Invoice No", "") or row.get("Invoice No", "")
    amount = row.get("Amount", "")
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

# robust reader: supports .xls/.xlsx and html-based exports (sheet001.htm)
def read_sheet_from_file(p: Path) -> pd.DataFrame:
    if p.is_dir():
        # prefer sheet001.htm -> any .htm/.html -> any .xls/.xlsx
        for cand in ("sheet001.htm", "sheet001.html"):
            c = p / cand
            if c.exists():
                p = c
                break
        else:
            htm = next(p.glob("*.htm"), None) or next(p.glob("*.html"), None)
            if htm:
                p = htm
            else:
                excel = next(p.glob("*.xls"), None) or next(p.glob("*.xlsx"), None)
                if excel:
                    p = excel
                else:
                    raise FileNotFoundError(f"No usable file inside: {p}")
    suffix = p.suffix.lower()
    # sniff HTML
    looks_html = False
    try:
        with p.open("rb") as f:
            head = f.read(512).lower()
        if b"<html" in head or head.lstrip().startswith(b"<!doctype"):
            looks_html = True
    except Exception:
        looks_html = False
    if suffix in (".htm", ".html") or looks_html:
        tables = pd.read_html(p, header=0)
        if not tables:
            raise RuntimeError("No tables in HTML")
        df = tables[0]
        df.columns = [str(c).strip() for c in df.columns]
        return df
    # binary excel
    if suffix == ".xls":
        df_dict = pd.read_excel(p, sheet_name=None, engine="xlrd", dtype=str)
    else:
        df_dict = pd.read_excel(p, sheet_name=None, dtype=str)
    if "input" in df_dict:
        return df_dict["input"]
    return next(iter(df_dict.values()))

# ---------------------------
# Conversion & template write (atomic then move to final to trigger watchers)
# ---------------------------
def convert_export_to_template(src: Path, company: str, year: str, suffix_tag: str) -> Path:
    log.info("Converting export -> template: %s", src.name)
    df = read_sheet_from_file(src)
    df.columns = [str(c).strip() for c in df.columns]
    rows = [map_row_to_template(r) for _, r in df.iterrows()]
    out_df = pd.DataFrame(rows, columns=TEMPLATE_COLUMNS)
    filename = f"{company}-{year}-{suffix_tag}.xlsx"
    tmp = TEMPLATE_DIR / (filename + ".tmp")
    final = TEMPLATE_DIR / filename
    # write tmp
    out_df.to_excel(tmp, index=False, engine="openpyxl")
    # move tmp -> final to trigger file-created events reliably
    shutil.move(str(tmp), str(final))
    log.info("Wrote template: %s", final)
    return final

# ---------------------------
# RUN LOCK & processed registry
# ---------------------------
RUN_LOCK = threading.Lock()
_processed: Dict[str, float] = {}
_processing: set[str] = set()

def already_processed(path: Path) -> bool:
    k = str(path.resolve())
    try:
        m = path.stat().st_mtime
    except Exception:
        return False
    last = _processed.get(k)
    if last is not None and abs(last - m) < 1e-6:
        return True
    return False

def mark_processed(path: Path):
    try:
        _processed[str(path.resolve())] = path.stat().st_mtime
    except Exception:
        pass


def is_processing(path: Path) -> bool:
    return str(path.resolve()) in _processing


def mark_processing(path: Path):
    try:
        _processing.add(str(path.resolve()))
    except Exception:
        pass


def unmark_processing(path: Path):
    try:
        _processing.discard(str(path.resolve()))
    except Exception:
        pass

# ---------------------------
# Process template: validate & call express workflow & move processed
# ---------------------------
def validate_template(path: Path) -> bool:
    try:
        if not wait_file_ready(path):
            log.warning("Template not ready: %s", path)
            return False
        df = pd.read_excel(path, dtype=str)
        missing = [c for c in TEMPLATE_COLUMNS if c not in df.columns]
        if missing:
            log.error("Template missing columns: %s", missing)
            return False
        return True
    except Exception as e:
        log.exception("Template read error: %s", e)
        return False

def process_template(path: Path):
    if not path.exists():
        return
    # Avoid concurrent processing of the same file from multiple handlers
    if is_processing(path):
        log.debug("Template is already being processed: %s", path.name)
        return
    mark_processing(path)
    try:
        if already_processed(path):
            log.debug("Template already processed (mtime same): %s", path.name)
            return
        if not validate_template(path):
            log.warning("Template validation failed: %s", path.name)
            return

        # parse search_key from filename
        search_key = parse_filename_search_key(path.stem)
        log.info("Processing template %s (search_key=%s)", path.name, search_key)

        # run express workflow (if available)
        if run_full_workflow is None:
            log.warning("express_launcher.run_full_workflow not available in this process; skipping actual run.")
        else:
            # Acquire RUN_LOCK to avoid concurrent workflows
            if not RUN_LOCK.acquire(blocking=False):
                log.warning("Workflow is busy; skipping: %s", path.name)
                return
            try:
                try:
                    run_full_workflow(file_path=str(path), search_key=search_key)
                except TypeError:
                    # fallback older signature
                    run_full_workflow()
            finally:
                RUN_LOCK.release()

        # move to processed and then mark. Marking only after a successful move
        # prevents the file remaining in-place but being considered processed.
        dest = TEMPLATE_PROCESSED / path.name
        if dest.exists():
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            dest = TEMPLATE_PROCESSED / f"{path.stem}-{ts}{path.suffix}"
        try:
            shutil.move(str(path), str(dest))
            log.info("Moved template to processed: %s", dest)
            # mark processed using the final location's timestamp
            mark_processed(dest)
        except Exception:
            log.exception("Failed to move template to processed: %s", path)
    finally:
        # ensure processing flag is cleared so future events can handle the file
        unmark_processing(path)

# ---------------------------
# Handlers: incoming exports and templates watchers
# ---------------------------
class IncomingHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self._lock = threading.Lock()

    def _handle(self, src: Path, event_name: str):
        if not src.exists():
            return
        if src.is_dir() and not any(src.iterdir()):
            return
        # accept files or folder containers
        if not (src.is_dir() or src.suffix.lower() in (".xls", ".xlsx", ".htm", ".html")):
            log.debug("Incoming skip (not an accepted type): %s", src)
            return
        if not self._lock.acquire(blocking=False):
            log.debug("Incoming handler busy; skipping: %s", src)
            return
        try:
            log.info("[INCOMING:%s] %s", event_name, src.name)
            if not wait_file_ready(src, timeout=30.0):
                log.warning("Incoming file not stable: %s", src)
                return
            # Prevent duplicate handling of the same incoming path (watcher + poller)
            if is_processing(src):
                log.debug("Incoming already being processed: %s", src)
                return
            mark_processing(src)
            # quick read test
            try:
                _ = read_sheet_from_file(src)
            except Exception as e:
                log.exception("Cannot read incoming file: %s", e)
                return
            # ask user for company/year/suffix (unless headless)
            if HEADLESS:
                company = "EDS"
                year = str(datetime.now().year)
                suffix = "RR"
            else:
                # simple console-less prompt using input() fallback (tkinter on Windows is ok)
                try:
                    import tkinter as tk
                    from tkinter import simpledialog, messagebox
                    root = tk.Tk()
                    root.withdraw()
                    company = simpledialog.askstring("Company", "Enter company (EDS or FIX):", initialvalue="EDS", parent=root)
                    if company is None:
                        log.info("User cancelled conversion")
                        return
                    company = company.strip().upper()
                    year = simpledialog.askstring("Year", "Enter year (YYYY):", initialvalue=str(datetime.now().year), parent=root)
                    if year is None:
                        log.info("User cancelled conversion")
                        return
                    suffix = simpledialog.askstring("Suffix", "Enter suffix (e.g. RR):", initialvalue="RR", parent=root)
                    root.destroy()
                    if suffix is None:
                        log.info("User cancelled conversion")
                        return
                    company = company.strip().upper()
                    year = year.strip()
                    suffix = suffix.strip() or "RR"
                except Exception:
                    log.exception("GUI dialog failed; falling back to EDS/current-year/RR")
                    company = "EDS"; year = str(datetime.now().year); suffix = "RR"

            # convert
            try:
                template_path = convert_export_to_template(src, company, year, suffix)
            except Exception as e:
                log.exception("Conversion failed: %s", e)
                return

            # move original
            try:
                dest = INCOMING_PROCESSED / src.name
                if dest.exists():
                    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
                    dest = INCOMING_PROCESSED / f"{src.stem}-{ts}{src.suffix}"
                shutil.move(str(src), str(dest))
                log.info("Moved original to processed: %s", dest)
            except Exception:
                log.exception("Failed to move original: %s", src)

            # process the template immediately (don't rely only on fs events)
            try:
                process_template(template_path)
            finally:
                # clear incoming processing flag (template processing has its own flag)
                unmark_processing(src)
        finally:
            self._lock.release()

    def on_created(self, event):
        self._handle(Path(event.src_path), "created")

    def on_moved(self, event):
        self._handle(Path(event.dest_path), "moved")

class TemplateHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self._lock = threading.Lock()

    def _handle(self, src: Path, event_name: str):
        if not src.exists() or not src.is_file():
            return
        if src.suffix.lower() not in (".xlsx", ".xls"):
            return
        if not self._lock.acquire(blocking=False):
            return
        try:
            log.info("[TEMPLATE:%s] %s", event_name, src.name)
            # small wait to ensure ready
            if not wait_file_ready(src, timeout=15.0):
                log.warning("Template not stable yet: %s", src)
                return
            process_template(src)
        finally:
            self._lock.release()

    def on_created(self, event):
        self._handle(Path(event.src_path), "created")

    def on_moved(self, event):
        self._handle(Path(event.dest_path), "moved")

# ---------------------------
# Poller fallback (scans both folders) - optional but helpful
# ---------------------------
def poll_loop(interval: float = 2.0):
    while True:
        try:
            # scan templates folder
            for f in TEMPLATE_DIR.iterdir():
                try:
                    if f.is_file() and f.suffix.lower() in (".xlsx", ".xls") and not already_processed(f):
                        log.debug("[POLL] template found: %s", f.name)
                        process_template(f)
                except Exception:
                    log.exception("poll(template) error for %s", f)
            # scan incoming folder for manual dropped files
            for f in INCOMING.iterdir():
                try:
                    if f.is_file() and f.suffix.lower() in (".xls", ".xlsx", ".htm", ".html"):
                        log.debug("[POLL] incoming found: %s", f.name)
                        # use incoming handler directly
                        IncomingHandler()._handle(f, "polled")
                except Exception:
                    log.exception("poll(incoming) error for %s", f)
        except Exception:
            log.exception("Poll loop outer error")
        time.sleep(interval)

# ---------------------------
# Entrypoint
# ---------------------------
def main():
    log.info("Automation service starting...")
    # Observers
    obs = Observer()
    incoming_handler = IncomingHandler()
    template_handler = TemplateHandler()
    obs.schedule(incoming_handler, str(INCOMING), recursive=False)
    obs.schedule(template_handler, str(TEMPLATE_DIR), recursive=False)
    obs.start()

    # start poller thread (daemon)
    p = threading.Thread(target=poll_loop, args=(2.0,), daemon=True)
    p.start()

    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        log.info("Shutting down...")
        obs.stop()
    obs.join()

if __name__ == "__main__":
    main()
