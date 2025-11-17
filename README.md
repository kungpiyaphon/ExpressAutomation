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
├─ incoming_exports/
│  └─ processed/
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

## ⚡ Tuning speeds (env vars)

To speed up or slow down the UI automation safely, set these environment variables before running `python src/main.py`.

- `TAP_DELAY` (float): pause (s) after each Tab/Enter step in `express_excel_entry.py`. Default `0.4`.
- `TYPE_INTERVAL` (float): per-character typing interval (s). Default `0.04`.
- `ROW_DELAY` (float): pause (s) between rows. Default `0.2`.
- `LAUNCH_WAIT` (float): wait (s) after launching Express. Default `3.0`.
- `LOGIN_WAIT` (float): brief wait before typing credentials. Default `1.0`.
- `LOGIN_TYPE_INTERVAL` (float): typing interval used during login. Default `0.06`.
- `SEARCH_TYPE_INTERVAL` (float): typing interval for company search key. Default `0.06`.
- `SEARCH_OK_DELAY` (float): pause (s) between OK presses during search. Default `0.25`.
- `KEY_INTERVAL` and `STEP_DELAY` (express menu): further fine-grain menu timing; defaults `0.03` and `0.20`.

Example (PowerShell):

```powershell
 $env:TAP_DELAY = '0.5'
 python src/main.py
```

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

---

**Building An Executable (.exe)**

- **Overview:** Use `PyInstaller` to package the Python app into a single Windows executable. You can either use a simple command or an existing `.spec` file (`main.spec`) included in the repo.
- **Install PyInstaller:**

  - Activate your virtual environment (bash on Windows):

    ```bash
    source venv/Scripts/activate
    ```

  - Install the package:

    ```bash
    pip install pyinstaller
    ```

- **Build command (single-file):**

  - Example (includes data folders `excel_templates`, `icons`, `incoming_exports`):

    ```bash
    pyinstaller --clean --onefile --name ExpressAutomation \
      --add-data "excel_templates;excel_templates" \
      --add-data "icons;icons" \
      --add-data "incoming_exports;incoming_exports" \
      src/main.py
    ```

  - Notes:
    - Use `--noconsole` if your app should not open a console window (GUI apps).
    - On Windows the `--add-data` separator is `;` (semicolor). On other platforms use `:`.
    - The produced executable will be in the `dist/` folder, for example `dist/ExpressAutomation.exe`.

- **Using the included spec file:**

  - If you prefer custom spec settings, run:

    ```bash
    pyinstaller main.spec
    ```

**Creating an Installer with Inno Setup (.iss)**

- **Overview:** Inno Setup creates a Windows installer (.exe) that wraps your built executable and any supporting files. Download and install Inno Setup from https://jrsoftware.org/ if not already installed.
- **Example `.iss` script (save as e.g. `build\\ExpressAutomation.iss`)**

  ```innosetup
  [Setup]
  AppName=ExpressAutomation
  AppVersion=1.0
  DefaultDirName={pf}\\ExpressAutomation
  DefaultGroupName=ExpressAutomation
  OutputBaseFilename=ExpressAutomationInstaller
  Compression=lzma
  SolidCompression=yes

  [Files]
  Source: "dist\\ExpressAutomation.exe"; DestDir: "{app}"; Flags: ignoreversion
  Source: "excel_templates\\*"; DestDir: "{app}\\excel_templates"; Flags: recursesubdirs createallsubdirs
  Source: "incoming_exports\\*"; DestDir: "{app}\\incoming_exports"; Flags: recursesubdirs createallsubdirs
  Source: "icons\\*"; DestDir: "{app}\\icons"; Flags: recursesubdirs createallsubdirs

  [Icons]
  Name: "{group}\\ExpressAutomation"; Filename: "{app}\\ExpressAutomation.exe"

  [Run]
  Filename: "{app}\\ExpressAutomation.exe"; Description: "Launch ExpressAutomation"; Flags: nowait postinstall skipifsilent
  ```

- **Compile the installer (command-line):**

  - Example (Inno Setup default install path):

    ```bash
    "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe" build\\ExpressAutomation.iss
    ```

  - Or open the `.iss` in the Inno Setup IDE and press Compile.

- **Tips & adjustments:**
  - Update `AppVersion`, `OutputBaseFilename`, and `DefaultDirName` as needed.
  - If your app requires configuration files or logs, add them under `[Files]` and set appropriate `Flags`.
  - Test the installer in a VM or clean machine before distribution.

---

If you want, I can now: build the `.exe` using `pyinstaller` from this environment, or compile the Inno Setup installer (you'll need Inno Setup installed). Which would you like me to do next?

