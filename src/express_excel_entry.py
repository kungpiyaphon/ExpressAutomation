# express_excel_entry_simple.py
"""
Simple step-by-step automation for Excel -> Express UI (debug / validation)
This script implements the exact sequence of steps you supplied, one row at a time.
Purpose: slow, visible, and easy to adjust counts/delays so you can observe UI behaviour.

How to use:
- Put this file in `src/` and run it (or call `process_excel_to_express(path)`).
- By default it will actually send keystrokes. To only log actions without touching UI set
    DRY_RUN=1
  in the environment.
- To enable "step mode" (wait for you to press Enter in console between major steps) set
    STEP_MODE=1
- Tweak delays with env vars:
    TAP_DELAY (float seconds, default 1.0)  - delay after each tab/ok press to observe
    TYPE_INTERVAL (float seconds, default 0.08) - typing character interval
    ROW_DELAY (float seconds, default 0.6) - delay between rows
- The script expects the Excel template columns: Dept, Date, Supplier, Invoice, Code, Qty, UnitCost
"""

from __future__ import annotations
import os
import time
from pathlib import Path
from typing import Dict, Any

import pandas as pd
import pyautogui
import ctypes
from tkinter import messagebox, Tk

# ------------------------
# Config (env overrides)
# ------------------------
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.02

DRY_RUN = os.getenv("DRY_RUN", "0") in ("1", "true", "True")
STEP_MODE = os.getenv("STEP_MODE", "0") in ("1", "true", "True")

# Tunable speeds (defaults tuned to be faster but safe).
# These can be overridden via environment variables when running.
# - TAP_DELAY: pause after each Tab/Enter/step (seconds)
# - TYPE_INTERVAL: per-character typing interval used when filling form fields
# - ROW_DELAY: pause between rows
TAP_DELAY = float(os.getenv("TAP_DELAY", "0.4"))         # was 1.0 (slower), now 0.4s
TYPE_INTERVAL = float(os.getenv("TYPE_INTERVAL", "0.04"))  # was 0.08, now 0.04s/char
ROW_DELAY = float(os.getenv("ROW_DELAY", "0.2"))        # was 0.6, now 0.2s

REQUIRED_COLS = ["Dept", "Date", "Supplier", "Invoice", "Code", "Qty", "UnitCost"]

# ------------------------
# Helpers: keyboard/layout
# ------------------------
def _current_keyboard_layout_hex() -> int:
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
    klid = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    return klid & (2**16 - 1)

def require_english_or_abort() -> bool:
    EN = 0x0409
    cur = _current_keyboard_layout_hex()
    if cur != EN:
        print(f"[ERROR] Keyboard layout is not English (0x0409). Current: {hex(cur)}")
        return False
    return True

# ------------------------
# UI action wrappers (log + optionally execute)
# ------------------------
def _log(msg: str):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

def _maybe_exec(action_name: str, func=None, *args, **kwargs):
    """Log and optionally execute pyautogui action."""
    _log(f"ACTION -> {action_name} | args={args} kwargs={kwargs}")
    if DRY_RUN:
        _log("(DRY_RUN) skipping actual UI action")
        return None
    try:
        if func:
            return func(*args, **kwargs)
    except Exception as e:
        _log(f"[ERROR] action {action_name} failed: {e}")
        raise

def press_ok(times: int = 1):
    for i in range(times):
        _maybe_exec(f"press ENTER ({i+1}/{times})", pyautogui.press, "enter")
        time.sleep(TAP_DELAY)

def press_tab(times: int = 1):
    for i in range(times):
        _maybe_exec(f"press TAB ({i+1}/{times})", pyautogui.press, "tab")
        time.sleep(TAP_DELAY)

def type_text(text: str):
    s = "" if text is None else str(text)
    _log(f"typing -> '{s}'")
    _maybe_exec("typewrite", pyautogui.typewrite, s, interval=TYPE_INTERVAL)
    # tiny pause after typing so UI can react
    time.sleep(0.12)

def hotkey_alt_a():
    _maybe_exec("hotkey alt+a", pyautogui.hotkey, "alt", "a")
    time.sleep(TAP_DELAY)

def show_done_popup(msg: str):
    try:
        root = Tk()
        root.withdraw()
        messagebox.showinfo("Express Automation", msg)
        root.destroy()
    except Exception:
        # headless or blocked, just log
        _log(f"[POPUP] {msg}")

def step_pause(prompt: str = "Press Enter to continue..."):
    if STEP_MODE:
        input(f"[STEP MODE] {prompt}")

# ------------------------
# Validation + I/O
# ------------------------
def read_template_excel(path: str):
    fp = Path(path)
    if not fp.exists():
        raise FileNotFoundError(f"Template not found: {fp}")
    df = pd.read_excel(fp, dtype=str, engine="openpyxl").fillna("")
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    return df

# ------------------------
# Sequence of steps per your spec
# ------------------------
def enter_row_sequence(row: Dict[str, Any]):
    """
    Implements the exact sequence requested:
    1. Dept (type)
    2. Tab x2
    3. Date (type)
    4. Tab x1
    5. Supplier (type)
    6. Tab x4
    7. Invoice (type)
    8. Tab x10
    9. OK (Enter) x5
    10. Tab x1
    11. Code (type)
    12. Tab x3
    13. Qty (type)
    14. Tab x1
    15. OK x2
    16. UnitCost (type)
    17. OK x3
    18. F9 (save)
    19. OK x1
    20. Alt+A
    """
    # 1 Dept
    _log("STEP 1: Dept")
    type_text(row.get("Dept", ""))
    step_pause("Dept typed. ")

    # 2 Tab x2
    _log("STEP 2: Tab x2")
    press_tab(2)

    # 3 Date
    _log("STEP 3: Date")
    type_text(row.get("Date", ""))
    step_pause("Date typed. ")

    # 4 Tab x1
    _log("STEP 4: Tab x1")
    press_tab(1)

    # 5 Supplier
    _log("STEP 5: Supplier")
    type_text(row.get("Supplier", ""))
    step_pause("Supplier typed. ")

    # 6 Tab x4
    _log("STEP 6: Tab x4")
    press_tab(4)

    # 7 Invoice
    _log("STEP 7: Invoice")
    type_text(row.get("Invoice", ""))
    step_pause("Invoice typed. ")

    # 8 Tab x10
    _log("STEP 8: Tab x10")
    press_tab(10)

    # 9 OK x5
    _log("STEP 9: OK x5")
    press_ok(5)

    # 10 Tab x1
    _log("STEP 10: Tab x1")
    press_tab(1)

    # 11 Code
    _log("STEP 11: Code")
    type_text(row.get("Code", ""))
    step_pause("Code typed. ")

    # 12 Tab x3
    _log("STEP 12: Tab x3")
    press_tab(3)

    # 13 Qty
    _log("STEP 13: Qty")
    type_text(row.get("Qty", ""))
    step_pause("Qty typed. ")

    # 14 Tab x1
    _log("STEP 14: Tab x1")
    press_tab(1)

    # 15 OK x2
    _log("STEP 15: OK x2")
    press_ok(2)

    # 16 UnitCost
    _log("STEP 16: UnitCost")
    type_text(row.get("UnitCost", ""))
    step_pause("UnitCost typed. ")

    # 17 OK x3
    _log("STEP 17: OK x3")
    press_ok(3)

    # 18 F9
    _log("STEP 18: Press F9 to save")
    _maybe_exec("press F9", pyautogui.press, "f9")
    time.sleep(TAP_DELAY)

    # 19 OK x1
    _log("STEP 19: OK x1")
    press_ok(1)

    # 20 Alt+A (start new)
    _log("STEP 20: Alt+A (start next)")
    hotkey_alt_a()

# ------------------------
# High level process: read excel and iterate rows
# ------------------------
def process_excel_simple(file_path: str):
    """
    Read the template and execute the visible slow sequence for each row.
    """
    if not require_english_or_abort():
        _log("Keyboard not English; aborting.")
        return

    df = read_template_excel(file_path)
    total = len(df)
    _log(f"Loaded {total} rows from {file_path}")

    for idx, row in df.iterrows():
        rownum = idx + 1
        _log(f"--- Starting row {rownum}/{total} ---")
        try:
            enter_row_sequence(row.to_dict())
            _log(f"--- Finished row {rownum}/{total} ---")
        except Exception as e:
            _log(f"[ERROR] Row {rownum} failed: {e}")
            # continue to next row
        # small pause between rows
        time.sleep(ROW_DELAY)

    _log("All rows processed.")
    show_done_popup("All rows processed (simple run).")


# Compatibility wrapper for older import name used in `express_launcher`
def process_excel_to_express(file_path: str, company_key: str | None = None):
        """
        Backwards-compatible wrapper expected by `express_launcher`.
        - `company_key` is accepted for compatibility but currently unused by
            the simple implementation. Keeping the parameter avoids TypeError
            during calls from older/newer signatures.
        """
        _log(f"process_excel_to_express called (company_key={company_key})")
        # For now call the simple, visible runner. Later this can be expanded
        # to support more robust batch or headless operation.
        return process_excel_simple(file_path)

# ------------------------
# If executed directly, simple CLI
# ------------------------
if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Simple step-by-step Excel -> Express tester")
    p.add_argument("file", help="path to template Excel (xlsx)")
    args = p.parse_args()
    process_excel_simple(args.file)
