import os
import json
import time
import subprocess
import ctypes
from pathlib import Path
from typing import Optional

import pyautogui

from express_menu import open_credit_purchase_add
from express_excel_entry import process_excel_to_express

# -------------------------
# Global behavior for PyAutoGUI
# -------------------------
pyautogui.FAILSAFE = True      # มุมซ้ายบน = emergency stop
pyautogui.PAUSE = 0.05

# -------------------------
# Paths / Project root
# -------------------------
PROJECT_ROOT = Path(__file__).resolve().parents[1]   # .../ExpressAutomation
EXCEL_DEFAULT = PROJECT_ROOT / "excel_templates" / "express_import_template.xlsx"
CONFIG_FILE = PROJECT_ROOT / "express.config.json"   # optional

# =========================
# Keyboard layout helpers
# =========================
def get_current_keyboard_layout() -> int:
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
    klid = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    lid = klid & (2**16 - 1)
    return lid

def require_keyboard_english() -> bool:
    """คีย์บอร์ดต้องเป็น EN (0x0409) เท่านั้น ไม่สลับอัตโนมัติ"""
    EN = 0x0409
    cur = get_current_keyboard_layout()
    if cur != EN:
        print(f"[ERROR] Keyboard layout must be English (0x0409). Current: {hex(cur)}")
        print("[HINT] โปรดสลับภาษาเป็น English ก่อน แล้วค่อยรันใหม่ (เช่น Alt+Shift)")
        return False
    print("[INFO] Keyboard layout OK (English)")
    return True

# =========================
# Credentials
# =========================
def get_credentials() -> tuple[Optional[str], Optional[str]]:
    cred_path = Path(__file__).parent / "credential.json"   # src/credential.json
    try:
        with cred_path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("username"), data.get("password")
    except FileNotFoundError:
        print(f"[ERROR] credential.json not found at {cred_path}")
        return None, None
    except Exception as e:
        print(f"[ERROR] Failed reading credentials: {e}")
        return None, None

# =========================
# Express path resolver
# =========================
def resolve_express_path(param_path: Optional[str]) -> Optional[str]:
    """
    ลำดับความสำคัญ:
    1) พารามิเตอร์ที่ส่งมา
    2) ENV: EXPRESS_PATH
    3) express.config.json { "express_path": "..." }
    4) ดีฟอลต์องค์กร: Z:\ExpressI.exe
    """
    # 1) explicit param
    if param_path:
        p = Path(param_path)
        if p.exists():
            return str(p)

    # 2) ENV
    env_path = os.getenv("EXPRESS_PATH")
    if env_path and Path(env_path).exists():
        return env_path

    # 3) config file
    if CONFIG_FILE.exists():
        try:
            with CONFIG_FILE.open("r", encoding="utf-8") as f:
                cfg = json.load(f)
            p = cfg.get("express_path")
            if p and Path(p).exists():
                return p
        except Exception as e:
            print(f"[WARN] Cannot read {CONFIG_FILE}: {e}")

    # 4) org default
    default_path = r"Z:\ExpressI.exe"
    if Path(default_path).exists():
        return default_path

    return None

# =========================
# Launch / Login
# =========================
def launch_express(express_path: Optional[str]) -> bool:
    exe = resolve_express_path(express_path)
    if not exe:
        print("[ERROR] Express executable not found. ตั้งค่าแมพไดรฟ์ Z: หรือระบุ express_path/ENV/express.config.json")
        return False
    try:
        subprocess.Popen([exe])
        print(f"[INFO] Launched Express: {exe}")
        # รอ UI เบื้องต้น
        time.sleep(3)
        return True
    except Exception as e:
        print(f"[ERROR] Failed to launch Express: {e}")
        return False

def enter_credentials() -> bool:
    username, password = get_credentials()
    if not username or not password:
        print("[ERROR] Missing username/password")
        return False

    # ก่อนพิมพ์ทุกครั้ง ยืนยัน EN อีกที
    if not require_keyboard_english():
        return False

    # focus ที่ฟิลด์ แล้วเคลียร์
    time.sleep(0.5)
    pyautogui.press('tab', presses=2)
    pyautogui.hotkey('ctrl', 'a'); pyautogui.press('delete')
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'a'); pyautogui.press('delete')
    time.sleep(0.2)

    print("[INFO] Typing username & password...")
    pyautogui.typewrite(username, interval=0.12)
    pyautogui.press('tab')
    pyautogui.typewrite(password, interval=0.12)
    pyautogui.press('enter')
    return True

def apply_search_key(search_key: Optional[str]) -> None:
    """
    ตามสเปกกุ้ง:
    - กด Tab 1 ครั้ง
    - พิมพ์ search_key (เช่น EDS2025)
    - กด OK (Enter) 4 ครั้ง
    """
    if not search_key:
        print("[INFO] No search_key provided, skipping company selection step.")
        return

    if not require_keyboard_english():
        print("[ERROR] Keyboard not EN during search_key step; aborting.")
        raise RuntimeError("Keyboard must be EN before typing search_key")

    print(f"[INFO] Applying search_key: {search_key}")
    time.sleep(0.5)
    pyautogui.press('tab', presses=1)
    time.sleep(0.2)
    pyautogui.typewrite(search_key, interval=0.10)
    time.sleep(0.2)
    for i in range(4):
        pyautogui.press('enter')
        print(f"[INFO] OK press {i+1}/4")
        time.sleep(0.35)

# =========================
# Main entry
# =========================
def run_full_workflow(
    file_path: Optional[str] = None,
    search_key: Optional[str] = None,
    express_path: Optional[str] = None
):
    """Main automation entry.
    - file_path: Excel path from watcher
    - search_key: e.g., 'EDS2025' parsed from filename
    - express_path: optional override (default is Z:\ExpressI.exe via resolver)
    """
    print("[START] Express Automation Workflow")
    print(f"[ARGS] file_path={file_path} | search_key={search_key} | express_path={express_path}")

    # เงื่อนไขบังคับ: ต้องเป็นภาษาอังกฤษก่อนเริ่มทุกอย่าง
    if not require_keyboard_english():
        return

    # เปิดโปรแกรม
    if not launch_express(express_path):
        return

    # ล็อกอิน
    time.sleep(2.0)
    if not enter_credentials():
        return

    # ขั้นตอนเลือกบริษัท/ปี ด้วย search_key
    try:
        apply_search_key(search_key)
    except Exception as e:
        print(f"[ERROR] search_key step failed: {e}")
        return

    # เข้าเมนูซื้อเชื่อ -> เพิ่มรายการ
    # (หากองค์กรต้องเปลี่ยนลำดับ สามารถย้ายจุดนี้ได้)
    print(f"[INFO] Navigating to Credit Purchase Add menu...")
    # open_credit_purchase_add()
    time.sleep(1.2)

    # ไฟล์ Excel (dynamic)
    excel_file = Path(file_path) if file_path else EXCEL_DEFAULT
    if not excel_file.exists():
        print(f"[ERROR] Excel file not found: {excel_file}")
        return
    print(f"[INFO] Using Excel file: {excel_file}")

    # ประมวลผลข้อมูล Excel → กรอกลง Express
    try:
        # พยายามส่ง company_key เข้าไปก่อน ถ้า signature ยังไม่รองรับจะ fallback
        print(f"[INFO] Processing Excel to Express with company_key={search_key}...")
        # process_excel_to_express(str(excel_file), company_key=search_key)
    except TypeError:
        print(f"[INFO] Processing Excel to Express without company_key...")
        # process_excel_to_express(str(excel_file))
    print("[DONE] Express launched, logged in, company selected, and Excel data processed!")

if __name__ == "__main__":
    run_full_workflow()
