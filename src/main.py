import os
import time
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from tkinter import messagebox, Tk

# ========================
# CONFIGURATION
# ========================
WATCH_FOLDER = r"C:\Users\piyaphon.w\Documents\EDS_ExpressAutomation\excel_templates"  # Folder ที่ใช้เฝ้าดู
EXPECTED_FILENAME = "express_import_template.xlsx"  # ต้องเป็นชื่อไฟล์นี้เท่านั้น
EXPECTED_COLUMNS = [
    "Department",
    "Date",
    "Distributor",
    "BillNumber",
    "ProductCode",
    "Quantity",
    "PricePerUnit",
]

# ========================
# Helper: แสดง Popup
# ========================
def show_popup(title, message):
    root = Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()

# ========================
# ตรวจสอบชื่อไฟล์ + เทมเพลต
# ========================
def validate_excel(filepath):
    filename = os.path.basename(filepath)

    # ตรวจชื่อไฟล์ก่อน
    if filename.lower() != EXPECTED_FILENAME.lower():
        show_popup("❌ Invalid Filename", f"File name '{filename}' is not allowed.\nExpected: {EXPECTED_FILENAME}")
        return False

    # ตรวจหัวคอลัมน์
    try:
        # รอให้ไฟล์พร้อม (บางครั้งไฟล์เพิ่งถูกคัดลอกเข้ามาและยังไม่พร้อมอ่าน)
        time.sleep(1)

        # เปิดไฟล์ในโหมด read-only
        with open(filepath, 'rb') as f:
            df = pd.read_excel(f)

        missing = [col for col in EXPECTED_COLUMNS if col not in df.columns]
        if missing:
            show_popup("❌ Template Error", f"Missing columns: {', '.join(missing)}")
            return False
        else:
            show_popup("✅ Template OK", f"File '{filename}' passed validation.")
            return True
    except PermissionError:
        show_popup("⚠️ File Locked", f"Cannot read '{filename}' because it is open in Excel.\nPlease close the file and try again.")
        return False
    except Exception as e:
        show_popup("❌ Read Error", f"Cannot read '{filename}'\nError: {e}")
        return False

# ========================
# Handler ตรวจจับไฟล์ใหม่
# ========================
class ExcelHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith((".xlsx", ".xls")):
            print(f"[NEW FILE DETECTED] {event.src_path}")
            validate_excel(event.src_path)

# ========================
# Main Program
# ========================
if __name__ == "__main__":
    if not os.path.exists(WATCH_FOLDER):
        os.makedirs(WATCH_FOLDER)

    observer = Observer()
    event_handler = ExcelHandler()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()

    show_popup(
        "Express Automation Started",
        f"👀 Watching folder:\n{WATCH_FOLDER}\n\nExpected file name:\n{EXPECTED_FILENAME}"
    )

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
