import pandas as pd
import pyautogui
import time

# =============================================================
# Function: Read Excel data
# =============================================================
def read_excel_data(file_path):
    df = pd.read_excel(file_path, dtype=str)  # ✅ อ่านทุกคอลัมน์เป็น string
    df = df.fillna('')  # กันกรณีมี NaN จะทำให้ typewrite error
    return df

# =============================================================
# Function: Enter a single row into Express
# =============================================================
def enter_row_into_express(row):
    # Dept -> tab 2
    pyautogui.typewrite(str(row['Dept']), interval=0.05)
    pyautogui.press('tab', presses=2, interval=0.1)

    # Date -> tab 1
    pyautogui.typewrite(str(row['Date']), interval=0.1)
    pyautogui.press('enter', presses=1, interval=0.5)

    # Supplier -> tab 1, then handle popup if exists
    pyautogui.typewrite(str(row['Supplier']), interval=0.05)
    pyautogui.press('tab', presses=1, interval=0.9)

    # ตรวจสอบว่ามี popup ปรากฏหรือไม่ (timeout ~3 วินาที)
    time.sleep(0.8)
    popup_found = False
    for i in range(6):  # loop check every 0.5s for up to 3s
        popup = pyautogui.locateOnScreen('ok_button.png', confidence=0.8) \
                 or pyautogui.locateOnScreen('select_window.png', confidence=0.8)
        if popup:
            popup_found = True
            print("[INFO] Supplier popup detected, selecting supplier...")
            break
        time.sleep(0.5)

    if popup_found:
        # มี popup → เลือก supplier ที่ 3 แล้วกด enter
        pyautogui.press('down', presses=3, interval=0.3)
        pyautogui.press('enter')
    else:
        # ไม่มี popup → เดินต่อเลย
        print("[INFO] No popup detected, skipping selection...")

    # หลังจาก popup หรือไม่ popup ให้ tab ต่อไปอีก 3 ครั้งเพื่อไปช่องถัดไป
    pyautogui.press('tab', presses=3, interval=0.9)

    # Invoice -> enter 11
    pyautogui.typewrite(str(row['Invoice']), interval=0.05)
    pyautogui.press('enter', presses=11, interval=0.9)

    # Code -> tab 1 + popup
    time.sleep(1)
    pyautogui.press('tab', presses=1, interval=0.9)

    code_value = str(row['Code']).strip()
    pyautogui.typewrite(code_value, interval=0.1)
    time.sleep(0.5)
    pyautogui.press('enter', presses=1, interval=0.9)

    # Qty -> OK popup
    # pyautogui.typewrite(str(row['Qty']), interval=0.05)
    # pyautogui.press('enter')  # handle popup automatically

    # UnitCost -> tab 3
    # pyautogui.typewrite(str(row['UnitCost']), interval=0.05)
    # pyautogui.press('tab', presses=3, interval=0.1)

    # # Save F9 -> Acquisition Basis (Enter)
    # pyautogui.press('f9')
    # time.sleep(0.5)
    # pyautogui.press('enter')

# =============================================================
# Function: Full workflow for Excel to Express
# =============================================================
def process_excel_to_express(file_path):
    df = read_excel_data(file_path)
    print(f"[INFO] {len(df)} rows detected in Excel")

    for idx, row in df.iterrows():
        print(f"[INFO] Processing row {idx + 1}")
        enter_row_into_express(row)
        time.sleep(0.5)  # small delay between rows

    print("[DONE] Excel data entry completed.")


# Example usage for direct testing (optional)
if __name__ == '__main__':
    # Only define excel_file here for manual test
    test_excel_file = r"C:\Users\piyaphon.w\Documents\Projects\ExpressAutomation\excel_templates\express_import_template.xlsx"
    process_excel_to_express(test_excel_file)
