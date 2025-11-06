import time
import pandas as pd
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from tkinter import messagebox, Tk

# ========================
# CONFIGURATION
# ========================
PROJECT_ROOT = Path(__file__).resolve().parents[1]   # .../ExpressAutomation
WATCH_FOLDER = PROJECT_ROOT / "excel_templates"
WATCH_FOLDER.mkdir(parents=True, exist_ok=True)

EXPECTED_FILENAME = "express_import_template.xlsx"
EXPECTED_COLUMNS = ["Dept", "Date", "Supplier", "Invoice", "Code", "Qty", "UnitCost"]

# ========================
# Helper: Popup
# ========================
def show_popup(title, message):
    try:
        root = Tk()
        root.withdraw()
        messagebox.showinfo(title, message)
        root.destroy()
    except Exception:
        # เผื่อรันในสภาพแวดล้อมที่ไม่มี GUI (เช่น RDP บางแบบ / บริการ)
        print(f"[POPUP:{title}] {message}")

# ========================
# Utils
# ========================
def is_target_excel(path: Path) -> bool:
    return (path.suffix.lower() in [".xlsx", ".xls"]) and (path.name.lower() == EXPECTED_FILENAME.lower())

def wait_file_ready(path: Path, timeout=10, interval=0.2) -> bool:
    """รอจนไฟล์ไม่ถูกล็อก/เขียนอยู่"""
    start = time.time()
    last_size = -1
    while time.time() - start < timeout:
        if not path.exists():
            return False
        try:
            size = path.stat().st_size
            # ถ้าขนาดไฟล์นิ่งติดต่อกัน 1 รอบ และเปิดอ่านได้ แปลว่าพร้อม
            if size == last_size:
                with path.open("rb"):
                    return True
            last_size = size
        except Exception:
            pass
        time.sleep(interval)
    return False

def validate_excel(path: Path) -> bool:
    if path.name.lower() != EXPECTED_FILENAME.lower():
        show_popup("❌ Invalid Filename",
                   f"File name '{path.name}' is not allowed.\nExpected: {EXPECTED_FILENAME}")
        return False
    try:
        if not wait_file_ready(path):
            show_popup("⚠️ File Busy", f"File '{path.name}' is not ready to read yet.")
            return False
        df = pd.read_excel(path)  # ต้องมี openpyxl ติดตั้งไว้สำหรับ .xlsx
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
# Handler
# ========================
class ExcelHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self._processing = False
        self._last_processed_mtime = {}

    def _maybe_process(self, p: Path):
        if not is_target_excel(p):
            return
        print(f"[EVENT] {p}")

        try:
            mtime = p.stat().st_mtime
        except FileNotFoundError:
            return

        resolved = p.resolve()
        last_mtime = self._last_processed_mtime.get(resolved)
        if last_mtime and mtime <= last_mtime:
            print("[INFO] Duplicate file event ignored (no new changes detected).")
            return

        if self._processing:
            print("[INFO] Workflow already running; event skipped.")
            return

        if validate_excel(p):
            try:
                # Lazy import เพื่อลดปัญหา import-time crash จาก pyautogui/cv2
                from express_launcher import run_full_workflow
                self._processing = True
                run_full_workflow()
                self._last_processed_mtime[resolved] = mtime
            except Exception as e:
                show_popup("❌ Workflow Error", f"run_full_workflow failed:\n{e}")
            finally:
                self._processing = False

    def on_created(self, event):
        if not event.is_directory:
            self._maybe_process(Path(event.src_path))

    def on_modified(self, event):
        if not event.is_directory:
            self._maybe_process(Path(event.src_path))

    def on_moved(self, event):
        # move เข้ามาในโฟลเดอร์
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
