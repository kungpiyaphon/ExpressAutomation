# 🧠 Excel Monitoring & Validation System

This project is a Python-based system designed to automatically monitor Excel files, validate their contents, and perform specific actions when changes occur. It leverages the `watchdog` library for real-time file event detection and a modular design for maintainability.

---

## 📂 Project Structure

```
ExpressAutomation/
├─ src/
│  ├─ main.py
│  ├─ express_excel_entry.py
│  ├─ express_launcher.py
│  └─ express_menu.py
│
├─ tools/
│  └─ export_watcher_converter.py    <-- สคริปต์ที่ผมให้
│
├─ incoming_exports/                 <-- **เพิ่ม**
│  └─ processed/                     <-- **เพิ่ม**
│
├─ excel_templates/
│  ├─ processed/
│  └─ EDS-2025-RR.xlsx
│
├─ venv/
├─ .gitignore
├─ express.config.json
├─ main.spec
├─ README.md
└─ requirements.txt
```

---

## ⚙️ Setup Instructions

### 1. Create Virtual Environment

```
python -m venv venv
```

### 2. Activate Virtual Environment

* **Windows (Command Prompt)**

  ```
  venv\Scripts\activate
  ```
* **Windows (PowerShell)**

  ```
  .\venv\Scripts\Activate.ps1
  ```
* **macOS / Linux**

  ```
  source venv/bin/activate
  ```

### 3. Install Dependencies

```
pip install -r requirements.txt
```

### 4. Run the Program

```
python src/main.py
```

---
## 💻 VS Code Setup (Recommended)

If you are using **Visual Studio Code**, a workspace setting has been included to automatically detect the correct Python interpreter from the virtual environment.

### 1. Open Command Palette
Press `Ctrl + Shift + P` → type `Python: Select Interpreter`

Then select:
Python 3.12.10 ('venv': venv)
[venv\Scripts\python.exe]

### 2. (Optional) Reload the window
If VS Code still shows import warnings (e.g. `pyautogui could not be resolved`), run:
Ctrl + Shift + P → Developer: Reload Window

## 🧩 Current Features

* ✅ File monitoring using **watchdog**
* ✅ Excel validation logic with custom rules
* ✅ Template-based structure for testing and expansion

--- 

## ⚠️ Known Limitations / Notes

* Currently, the automatic username/password entry assumes the keyboard layout is set to **English (US)**.
* If the keyboard is in another language (e.g., Thai), the script may type incorrect characters.
* Make sure to manually switch your keyboard to English before running the automation workflow for login.

---

## 🧱 Next Steps / Roadmap

* [ ] Add logging and error handling
* [ ] Create report summaries for validated files
* [ ] Add database or API integration for record tracking
* [ ] Develop a user interface for file upload & status monitoring

---

## 💬 Commit Log Summary

This section records commit messages for easy reference.

| Commit Type | Scope    | Message                                              |
| ----------- | -------- | ---------------------------------------------------- |
| feat        | watchdog | add Excel file monitoring and validation logic       |
| chore       | setup    | initialize project structure and virtual environment |
| docs        | readme   | add project documentation with setup guide           |

---

## 🧠 Tips & Best Practices

* Always activate your virtual environment before running the project.
* Use [Conventional Commits](https://www.conventionalcommits.org/) for clean commit history.
* Keep your `requirements.txt` updated after installing new packages.

---

## 🧑‍💻 Author

**KUNG ITEDS**
📧 *Internal IT Developer, EDS*
🚀 Focused on automation and internal process optimization.

