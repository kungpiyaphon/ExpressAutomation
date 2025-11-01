import time
import pyautogui

# =============================================================
# Function: Navigate Express Menu to Credit Purchase Add Data
# =============================================================
def open_credit_purchase_add():
    """
    Automates opening the Credit Purchase Add screen in Express.
    Steps:
    1. Press Alt+1 to open 'Purchase' menu.
    2. Press 4 to select 'Credit Purchase' from dropdown.
    3. Press Alt+A to add new data.
    """
    print("[INFO] Navigating to Credit Purchase Add screen...")

    # Step 1: Open Purchase menu
    pyautogui.hotkey('alt', '1')
    time.sleep(0.5)  # allow menu to open

    # Step 2: Select Credit Purchase
    pyautogui.press('4')
    time.sleep(0.5)  # allow dropdown to process

    # Step 3: Press Alt+A to add data
    pyautogui.hotkey('alt', 'a')
    time.sleep(1)  # allow new screen to open

    print("[INFO] Ready to input data.")
