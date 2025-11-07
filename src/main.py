import time
import re
import pandas as pd
from pathlib import Path
from typing import Optional, Tuple, Dict
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from tkinter import messagebox, Tk

# ========================
# CONFIG
# ========================
PROJECT_ROOT = Path(__file__).resolve().parents[1]            # .../ExpressAutomation
WATCH_FOLDER = PROJECT_ROOT / "excel_templates"
WATCH_FOLDER.mkdir(parents=True, exist_ok=True)

EXPECTED_COLUMNS = ["Dept", "Date", "Supplier", "Invoice", "Code", "Qty", "UnitCost"]

# ดีบ๊าวน์อย่างน้อย (วินาที) ต่อไฟล์ เพื่อลดการเรียกซ้ำเมื่อแอป sync/เขียนหลายรอบ
MIN_INTERVAL_SECONDS = 3.0

# รูปแบบชื่อไฟล์ที่ "แนะนำ": COMPANY-YYYY(-anything).xlsx
# เช่น: EDS-2025-RR.xlsx, SH-2025-batch1.xlsx
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
        # เผื่อรันในสภาพแวดล้อมที่ไม่มี GUI
        print(f"[POPUP:{title}] {message}")

# ========================
# Utils
# ========================
def wait_file_ready(path: Path, timeout=10, interval=0.2) -> bool:
    """รอจนไฟล์นิ่งและเปิดอ่านได้ (กันเคสกำลังคัดลอก/เขียน)"""
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

def parse_filename_for_search_key(filename_no_ext: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    คืนค่า: (company_upper, year_str, search_key) หรือ (None, None, None) ถ้าไม่ตรงแพทเทิร์น
    ตัวอย่าง: 'EDS-2025-RR' -> ('EDS', '2025', 'EDS2025')
    """
    m = RE_FILENAME.match(filename_no_ext)
    if not m:
        return None, None, None
    company = m.group(1).upper()
    year = m.group(2)
    return company, year, f"{company}{year}"

def validate_excel_schema(path: Path) -> bool:
    """ตรวจหัวคอลัมน์ตาม EXPECTED_COLUMNS"""
    try:
        if not wait_file_ready(path):
            show_popup("⚠️ File Busy", f"File '{path.name}' is not ready to read yet.")
            return False
        df = pd.read_excel(path)  # ต้องมี openpyxl สำหรับ .xlsx
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
# Debounce / processed cache
# ========================
_last_run: Dict[str, float] = {}

def should_run_now(p: Path, min_interval=MIN_INTERVAL_SECONDS) -> bool:
    key = str(p.resolve())
    now = time.time()
    last = _last_run.get(key, 0.0)
    if (now - last) < min_interval:
        return False
    _last_run[key] = now
    return True

# ========================
# Handler
# ========================
class ExcelHandler(FileSystemEventHandler):
    def _maybe_process(self, p: Path):
        if not is_excel_file(p):
            return

        # กันยิงถี่ ๆ
        if not should_run_now(p):
            return

        # แจ้ง event
        print(f"[EVENT] {p}")

        # ตรวจ schema ของไฟล์
        if not validate_excel_schema(p):
            return

        # แยก search_key จากชื่อไฟล์ (ไม่บังคับ แต่จะแจ้งเตือนถ้าไม่ตรงรูปแบบ)
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

        # เรียก workflow (พยายามส่งพารามิเตอร์ก่อน หากยังไม่รองรับค่อย fallback)
        try:
            from express_launcher import run_full_workflow
            try:
                print(f"[DONE] Sending function with parameters: file_path={str(p)}, search_key={search_key}")
                # run_full_workflow(file_path=str(p), search_key=search_key)
            except TypeError:
                # รองรับเวอร์ชันเก่า
                print(f"[DONE] Sending function without parameters")
                # run_full_workflow()
        except Exception as e:
            show_popup("❌ Workflow Error", f"run_full_workflow failed:\n{e}")

    def on_created(self, event):
        if not event.is_directory:
            self._maybe_process(Path(event.src_path))

    def on_modified(self, event):
        if not event.is_directory:
            self._maybe_process(Path(event.src_path))

    def on_moved(self, event):
        if not event.is_directory:
            self._maybe_process(Path(event.dest_path))

# ========================
# Main
# ========================
if __name__ == "__main__":
    print(f"[WATCHING] {WATCH_FOLDER}")
    observer = Observer()
    observer.schedule(ExcelHandler(), str(WATCH_FOLDER), recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
