import time
import re
import pandas as pd
from pathlib import Path
from typing import Optional, Tuple, Dict
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from tkinter import messagebox, Tk
import threading
import shutil

# ========================
# CONFIG
# ========================
PROJECT_ROOT = Path(__file__).resolve().parents[1]
WATCH_FOLDER = PROJECT_ROOT / "excel_templates"
WATCH_FOLDER.mkdir(parents=True, exist_ok=True)
PROCESSED_FOLDER = WATCH_FOLDER / "processed"
PROCESSED_FOLDER.mkdir(parents=True, exist_ok=True)

EXPECTED_COLUMNS = ["Dept", "Date", "Supplier", "Invoice", "Code", "Qty", "UnitCost"]

# ดีบ๊าวน์รวมไฟล์ (กัน spam เล็กน้อย แม้เรามี RUN_LOCK แล้ว)
MIN_INTERVAL_SECONDS = 5.0

RE_FILENAME = re.compile(r"^([A-Za-z]+)-(\d{4})(?:-[A-Za-z0-9._-]+)?$", re.IGNORECASE)

# ========================
# Popup helper
# ========================
def show_popup(title: str, message: str):
    try:
        root = Tk()
        root.withdraw()
        messagebox.showinfo(title, message)
        root.destroy()
    except Exception:
        print(f"[POPUP:{title}] {message}")

# ========================
# Utils
# ========================
def wait_file_ready(path: Path, timeout=10, interval=0.2) -> bool:
    start = time.time()
    last_size = -1
    while time.time() - start < timeout:
        if not path.exists():
            return False
        try:
            size = path.stat().st_size
            if size == last_size:
                with path.open("rb"):
                    return True
            last_size = size
        except Exception:
            pass
        time.sleep(interval)
    return False

def is_excel_file(path: Path) -> bool:
    return path.suffix.lower() in (".xlsx", ".xls")

def parse_filename_for_search_key(filename_no_ext: str):
    m = RE_FILENAME.match(filename_no_ext)
    if not m:
        return None, None, None
    company = m.group(1).upper()
    year = m.group(2)
    return company, year, f"{company}{year}"

def validate_excel_schema(path: Path) -> bool:
    try:
        if not wait_file_ready(path):
            show_popup("⚠️ File Busy", f"File '{path.name}' is not ready to read yet.")
            return False
        df = pd.read_excel(path)
        missing = [c for c in EXPECTED_COLUMNS if c not in df.columns]
        if missing:
            show_popup("❌ Template Error", f"Missing columns: {', '.join(missing)}")
            return False
        return True
    except PermissionError:
        show_popup("⚠️ File Locked", f"Cannot read '{path.name}'. Please close the file and try again.")
        return False
    except Exception as e:
        show_popup("❌ Read Error", f"Cannot read '{path.name}'\nError: {e}")
        return False

# ========================
# Debounce & processed registry & run-lock
# ========================
_last_run: Dict[str, float] = {}
_processed_by_mtime: Dict[str, float] = {}   # key=abs path, value=last mtime processed
RUN_LOCK = threading.Lock()                  # กันยิงซ้อนระหว่างกำลังรัน

def should_run_now(p: Path, min_interval=MIN_INTERVAL_SECONDS) -> bool:
    key = str(p.resolve())
    now = time.time()
    last = _last_run.get(key, 0.0)
    if (now - last) < min_interval:
        return False
    _last_run[key] = now
    return True

def already_processed(p: Path) -> bool:
    key = str(p.resolve())
    try:
        mtime = p.stat().st_mtime
    except FileNotFoundError:
        return False
    last_m = _processed_by_mtime.get(key)
    if last_m is not None and abs(last_m - mtime) < 1e-6:
        return True
    return False

def mark_processed(p: Path):
    key = str(p.resolve())
    try:
        _processed_by_mtime[key] = p.stat().st_mtime
    except FileNotFoundError:
        pass

# ========================
# Handler
# ========================
class ExcelHandler(FileSystemEventHandler):
    def _maybe_process(self, p: Path, event_name: str):
        if not is_excel_file(p):
            return

        # ข้ามถ้าประมวลผลไฟล์นี้ (mtime เดิม) ไปแล้ว
        if already_processed(p):
            return

        # กัน spam เบื้องต้น
        if not should_run_now(p):
            return

        print(f"[EVENT:{event_name}] {p}")

        if not validate_excel_schema(p):
            return

        company, year, search_key = parse_filename_for_search_key(p.stem)
        if search_key:
            print(f"[INFO] Parsed search_key from filename: {search_key}")
        else:
            show_popup(
                "ℹ️ Filename Hint",
                "Recommended pattern is COMPANY-YYYY-<anything>.xlsx\n"
                "Example: EDS-2025-RR.xlsx  → search_key = EDS2025\n"
                f"Received: {p.name}"
            )

        # ---- RUN LOCK: ถ้ากำลังรันอยู่ ให้ข้ามครั้งนี้ ----
        if not RUN_LOCK.acquire(blocking=False):
            print("[SKIP] Workflow is already running. This event will be ignored.")
            return

        try:
            # เรียก workflow
            print(f"[DONE] Sending function with parameters: file_path={p}, search_key={search_key}")
            try:
                from express_launcher import run_full_workflow
                run_full_workflow(file_path=str(p), search_key=search_key)
            except TypeError:
                from express_launcher import run_full_workflow
                run_full_workflow()

            # ทำเครื่องหมายว่าไฟล์นี้ (mtime นี้) ถูกประมวลผลแล้ว
            mark_processed(p)

            # (แนะนำ) ย้ายไฟล์เข้าโฟลเดอร์ processed/ เพื่อกัน event ซ้ำในอนาคต
            target = PROCESSED_FOLDER / p.name
            try:
                # ถ้าซ้ำชื่อ ให้เติม timestamp
                if target.exists():
                    ts = time.strftime("%Y%m%d-%H%M%S")
                    target = PROCESSED_FOLDER / f"{p.stem}-{ts}{p.suffix}"
                shutil.move(str(p), str(target))
                print(f"[INFO] Moved processed file to: {target}")
            except Exception as e:
                print(f"[WARN] Could not move file to processed/: {e}")

        finally:
            RUN_LOCK.release()

    def on_created(self, event):
        if not event.is_directory:
            self._maybe_process(Path(event.src_path), "created")

    # ⚠️ ปิด on_modified เพื่อตัดรอบซ้ำ
    # def on_modified(self, event):
    #     if not event.is_directory:
    #         self._maybe_process(Path(event.src_path), "modified")

    def on_moved(self, event):
        if not event.is_directory:
            self._maybe_process(Path(event.dest_path), "moved")

# ========================
# Main
# ========================
if __name__ == "__main__":
    print(f"[WATCHING] {WATCH_FOLDER}")
    observer = Observer()
    handler = ExcelHandler()
    observer.schedule(handler, str(WATCH_FOLDER), recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()