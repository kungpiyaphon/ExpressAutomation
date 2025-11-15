import time
from decimal import Decimal, InvalidOperation
from pathlib import Path
import pandas as pd
import pyautogui
import ctypes

# =========================
# Global config
# =========================
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05

STEP_DELAY = 0.15     # หน่วงระหว่าง step
TYPE_INTERVAL = 0.07  # ระยะห่างในการพิมพ์ตัวอักษร
RETRY = 2             # จำนวนครั้งที่ลองซ้ำเมื่อกรอกฟิลด์สำคัญ

REQUIRED_COLS = ["Dept", "Date", "Supplier", "Invoice", "Code", "Qty", "UnitCost"]

# =========================
# Helpers (keyboard/layout)
# =========================
def _current_keyboard_layout_hex() -> int:
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
    klid = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    return klid & (2**16 - 1)

def _require_english_or_abort() -> bool:
    EN = 0x0409
    cur = _current_keyboard_layout_hex()
    if cur != EN:
        print(f"[ERROR] Keyboard must be English (0x0409). Current: {hex(cur)}")
        return False
    return True

# =========================
# Helpers (typing/pressing)
# =========================
def press(key, presses=1, delay=STEP_DELAY):
    pyautogui.press(key, presses=presses, interval=0.02)
    time.sleep(delay)

def hotkey(*keys, delay=STEP_DELAY):
    pyautogui.hotkey(*keys)
    time.sleep(delay)

def type_text(text: str, interval=TYPE_INTERVAL, delay=0.05):
    if not isinstance(text, str):
        text = str(text)
    pyautogui.typewrite(text, interval=interval)
    time.sleep(delay)

def clear_field():
    pyautogui.hotkey('ctrl', 'a'); press('delete', delay=0.05)

# =========================
# Data normalizers
# =========================
def _to_ddmmyy_from_digits(digits: str) -> str:
    """รับเฉพาะตัวเลขในสตริง แล้วคืนค่า DDMMYY (6 หลัก)
       - ถ้า len == 6: ถือว่าเป็น DDMMYY อยู่แล้ว → คืนกลับ
       - ถ้า len == 8: ถือว่าเป็น DDMMYYYY → ตัดปี 2 หลักท้าย
       - อื่น ๆ: พยายาม parse ด้วย pandas → แปลงเป็น DDMMYY
    """
    d = ''.join(ch for ch in digits if ch.isdigit())
    if len(d) == 6:
        # e.g. 101168
        return d
    if len(d) == 8:
        # e.g. 10112568 หรือ 10112025 → ddmmyy = d[:4] + d[-2:]
        return d[:4] + d[-2:]
    # fallback: ใช้ pandas เดาแล้ว format
    try:
        dt = pd.to_datetime(digits, dayfirst=True, errors='coerce')
        if pd.isna(dt):
            dt = pd.to_datetime(digits, errors='coerce')
        if not pd.isna(dt):
            return f"{dt.day:02d}{dt.month:02d}{dt.year % 100:02d}"
    except Exception:
        pass
    # ถ้าเดาไม่ได้จริง ๆ ให้คืนค่าเดิม (แต่โดยมาตรฐานควรเป็นตัวเลข 6 หลักอยู่แล้ว)
    return digits.strip()

def norm_date_to_ddmmyy(s: str) -> str:
    """แปลงคอลัมน์ Date ให้เป็นสตริง 6 หลัก DDMMYY (ไม่มี / หรือ -)
       รองรับ:
       - รูปแบบที่เป็นตัวเลขล้วน: 101168, 10112025, 10112568
       - รูปแบบมีตัวคั่น: 10/11/68, 10-11-2568, 10/11/2025 ฯลฯ
    """
    s = (s or "").strip()
    if not s:
        return s
    return _to_ddmmyy_from_digits(s)

def norm_qty(s: str) -> str:
    s = (s or '').replace(',', '').strip()
    if not s:
        return '0'
    try:
        q = int(Decimal(s).to_integral_value())
        return str(q)
    except (InvalidOperation, ValueError):
        return s

def norm_cost(s: str) -> str:
    s = (s or '').replace(',', '').strip()
    if not s:
        return '0.00'
    try:
        c = Decimal(s)
        return f"{c:.2f}"
    except (InvalidOperation, ValueError):
        return s

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # สตริปทุกคอลัมน์
    for c in df.columns:
        df[c] = df[c].astype(str).fillna('').map(lambda x: x.strip())

    # Date -> DDMMYY (6 หลัก) ตาม requirement ใหม่
    if "Date" in df.columns:
        df["Date"] = df["Date"].map(norm_date_to_ddmmyy)

    # Qty / UnitCost
    if "Qty" in df.columns:
        df["Qty"] = df["Qty"].map(norm_qty)
    if "UnitCost" in df.columns:
        df["UnitCost"] = df["UnitCost"].map(norm_cost)

    return df

def validate_required_columns(df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

# =========================
# Excel I/O
# =========================
def read_excel_data(file_path: str) -> pd.DataFrame:
    fp = Path(file_path)
    if not fp.exists():
        raise FileNotFoundError(f"Excel not found: {fp}")
    df = pd.read_excel(fp, dtype=str, engine="openpyxl")
    validate_required_columns(df)
    df = df.fillna('')
    df = normalize_dataframe(df)
    return df

# =========================
# Field groups
# =========================
def enter_header_fields(row):
    # Dept -> tab 2
    type_text(row['Dept'])
    press('tab', presses=2)

    # Date (ตอนนี้เป็น DDMMYY 6 หลักแล้ว) -> Enter
    type_text(row['Date'], interval=0.2)
    press('enter')

    # Supplier -> tab 1
    type_text(row['Supplier'])
    press('tab')

    # เดินต่ออีก 3 tab ไปยัง Invoice
    press('tab', presses=3)

    # Invoice -> enter 11
    type_text(row['Invoice'])
    press('enter', presses=11)

def enter_item_fields(row):
    # ไป Code
    time.sleep(0.6)
    press('tab')  # 1 tab ไปช่อง Code

    # Code + confirm
    code_value = row['Code']
    for attempt in range(1, RETRY + 1):
        clear_field()
        type_text(code_value, interval=0.5)
        time.sleep(0.3)
        press('enter', presses=2)
        break

    # เดินไป Qty
    press('tab', presses=2)

    # Qty + OK popup (generic enter)
    # qty = row['Qty']
    # for attempt in range(1, RETRY + 1):
    #     clear_field(); type_text(qty)
    #     press('enter', presses=3)
    #     break

    # UnitCost
    # unit_cost = row['UnitCost']
    # clear_field(); type_text(unit_cost)
    # press('tab', presses=3)

def save_line_and_prepare_next(has_next_row: bool):
    # Save (F9) -> Acquisition Basis (Enter)
    press('f9')
    time.sleep(0.8)
    print("[INFO] Saved line")
    press('enter')
    if has_next_row:
        pyautogui.hotkey('alt', 'a'); time.sleep(0.5)

# =========================
# Main row entry
# =========================
def enter_row_into_express(row, is_last_row: bool):
    if not _require_english_or_abort():
        raise RuntimeError("Keyboard not EN")
    enter_header_fields(row)
    enter_item_fields(row)
    # save_line_and_prepare_next(has_next_row=not is_last_row)

# =========================
# Full workflow
# =========================
def process_excel_to_express(file_path: str, company_key: str | None = None):
    if not _require_english_or_abort():
        print("[ERROR] Keyboard must be EN; aborting.")
        return

    df = read_excel_data(file_path)
    print(f"[INFO] {len(df)} rows detected in Excel")

    for idx, row in df.iterrows():
        is_last = (idx == len(df) - 1)
        try:
            print(f"[INFO] Processing row {idx + 1}/{len(df)}  (Date={row['Date']})")
            enter_row_into_express(row, is_last_row=is_last)
            time.sleep(0.4)
        except Exception as e:
            print(f"[ERROR] Row {idx + 1} failed: {e}")
            # raise  # ถ้าต้องการหยุดทั้งงานเมื่อเจอ error
            continue

    print("[DONE] Excel data entry completed.")
