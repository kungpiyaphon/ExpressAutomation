import os
import json
import time
import subprocess
import pyautogui
import ctypes
from express_menu import open_credit_purchase_add

# =============================================================
# Function: Detect current keyboard layout (hex)
# =============================================================
def get_current_keyboard_layout():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
    klid = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    lid = klid & (2**16 - 1)
    return lid

# =============================================================
# Function: Switch keyboard layout to English if not already
# =============================================================
def switch_keyboard_to_english(wait=True, timeout=5):
    ENGLISH_LAYOUT = 0x0409
    current_layout = get_current_keyboard_layout()
    if current_layout != ENGLISH_LAYOUT:
        print(f"[INFO] Current layout is {hex(current_layout)}, switching to English...")
        pyautogui.hotkey('alt', 'shift')
        if wait:
            total_wait = 0
            while get_current_keyboard_layout() != ENGLISH_LAYOUT and total_wait < timeout:
                time.sleep(0.3)
                total_wait += 0.3
            if get_current_keyboard_layout() == ENGLISH_LAYOUT:
                print("[INFO] Keyboard layout is now English")
                time.sleep(1)  # allow Windows to process
            else:
                print(f"[WARN] Could not switch keyboard layout to English in time, current: {hex(get_current_keyboard_layout())}")
    else:
        print("[INFO] Keyboard already in English layout")

# =============================================================
# Function: Read credentials
# =============================================================
def get_credentials():
    try:
        with open(os.path.join(os.path.dirname(__file__), 'credential.json'), 'r', encoding='utf-8') as f:
            creds = json.load(f)
        return creds.get('username'), creds.get('password')
    except FileNotFoundError:
        print("[ERROR] credential.json not found.")
        return None, None

# =============================================================
# Function: Launch Express
# =============================================================
def launch_express():
    express_path = r"Z:\\ExpressI.exe"
    if not os.path.exists(express_path):
        print(f"[ERROR] Cannot find Express at {express_path}")
        return False

    try:
        subprocess.Popen([express_path])
        print("[INFO] Express launched.")
        time.sleep(3)
        return True
    except Exception as e:
        print(f"[ERROR] Failed to launch Express: {e}")
        return False

# =============================================================
# Function: Input credentials
# =============================================================
def enter_credentials():
    # 1. Switch layout and ensure English
    # switch_keyboard_to_english()
    time.sleep(0.5)  # allow Windows to process

    # 2. Read credentials
    username, password = get_credentials()
    print(f"[DEBUG] Read credentials -> username: {username}, password: {password}")
    if not username or not password:
        print("[ERROR] Missing username or password.")
        return False

    # 3. Focus username field
    pyautogui.press('tab', presses=2)

    # 4. Clear fields
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('delete')
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('delete')
    time.sleep(0.3)

    # 5. Verify layout again before typing
    if get_current_keyboard_layout() != 0x0409:
        print("[WARN] Keyboard layout is not English before typing, retry switching...")
        switch_keyboard_to_english()
        time.sleep(0.3)

    # 6. Type credentials slowly
    print("[INFO] Typing username...")
    pyautogui.typewrite(username, interval=0.15)
    pyautogui.press('tab')
    print("[INFO] Typing password...")
    pyautogui.typewrite(password, interval=0.15)
    pyautogui.press('enter')

    print("[INFO] Credentials entered.")
    return True

# =============================================================
# Function: Handle login failure popup
# =============================================================
def handle_login_failure():
    print("[INFO] Checking for login failure popup...")
    time.sleep(2)
    try:
        # Popup detection without OpenCV fallback
        popup = pyautogui.locateOnScreen('ok_button.png')  # simple image match
        if popup:
            print("[WARN] Login failed detected. Closing popup...")
            pyautogui.press('enter')  # Close popup

            time.sleep(1)
            print("[INFO] Retrying login after popup...")
            enter_credentials()
    except Exception as e:
        print(f"[WARN] Could not detect popup: {e}")

# =============================================================
# Function: Click OK 3 times
# =============================================================
def click_ok_buttons():
    time.sleep(2)
    for i in range(3):
        pyautogui.press('enter')
        print(f"[INFO] Clicked OK button {i+1}/3")
        time.sleep(1)

# =============================================================
# Function: Full workflow (MAIN ENTRY)
# =============================================================
def run_full_workflow():
    print("[START] Express Automation Workflow...")

    # 0. สลับภาษาเป็นอังกฤษล่วงหน้า
    switch_keyboard_to_english()
    time.sleep(1)  # ให้ Windows process เสร็จ

    if not launch_express():
        return

    time.sleep(4)
    enter_credentials()
    handle_login_failure()
    time.sleep(1)
    click_ok_buttons()

    print("[DONE] Express launched and logged in successfully!")
    open_credit_purchase_add()

if __name__ == "__main__":
    run_full_workflow()
