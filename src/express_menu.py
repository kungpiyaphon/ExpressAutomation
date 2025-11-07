import time
import pyautogui

# ปรับได้ตามเครื่อง/เครือข่าย
DEFAULT_KEY_INTERVAL = 0.05     # เวลาคั่นแต่ละคีย์
STEP_DELAY = 0.25               # เวลาคั่นแต่ละสเต็ป
RETRY = 3                       # จำนวนครั้งที่ลองซ้ำ

pyautogui.PAUSE = DEFAULT_KEY_INTERVAL
pyautogui.FAILSAFE = True  # มุมซ้ายบน = emergency stop

def _press_with_pause(*keys, delay=STEP_DELAY):
    """กดคีย์แล้วพักสั้น ๆ เพื่อให้ UI ทัน"""
    if len(keys) == 1:
        pyautogui.press(keys[0])
    else:
        pyautogui.hotkey(*keys)
    time.sleep(delay)

def open_credit_purchase_add():
    """
    เปิดหน้าจอ 'ซื้อเชื่อ -> เพิ่มรายการ' ด้วยลำดับคีย์ลัด:
      1) Alt+1  เปิดเมนูซื้อ
      2) 4      เลือก 'ซื้อเชื่อ'
      3) Alt+A  เพิ่มรายการใหม่
    มีกลไก retry เบา ๆ เผื่อ UI ช้า
    """
    print("[INFO] Navigating to Credit Purchase Add screen...")

    for attempt in range(1, RETRY + 1):
        try:
            # Step 1: Open Purchase menu
            _press_with_pause('alt', '1')

            # Step 2: Select Credit Purchase
            _press_with_pause('4')

            # Step 3: Add new record
            _press_with_pause('alt', 'a', delay=STEP_DELAY + 0.5)

            # เผื่อหน้าจอโหลด
            time.sleep(1.5)
            print("[INFO] Ready to input data.")
            return
        except Exception as e:
            print(f"[WARN] Menu navigation attempt {attempt}/{RETRY} failed: {e}")
            time.sleep(0.5)

    # ถ้าไม่สำเร็จใน RETRY ครั้ง
    print("[ERROR] Failed to navigate to Credit Purchase Add screen after retries.")
