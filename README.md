# JewishGen Mass Search

Desktop application that automates searching across all JewishGen databases simultaneously,
filters results by keyword, and saves matching records to Word (.docx) and/or Excel (.xlsx).

Built with Python + Playwright (browser automation) + PySide6 (GUI).

---

## Project structure

```
JewishGen_python/
│
├── app.py                    ← entry point — run this
├── scraper.py                ← browser automation + file writing
├── requirements.txt
│
├── gui/
│   ├── __init__.py
│   └── main_window.py        ← PySide6 window
│
├── models/
│   ├── __init__.py
│   └── search_models.py      ← SearchProfile / SearchRow dataclasses
│
├── storage/
│   ├── __init__.py
│   ├── autosave.py           ← saves/loads last session
│   ├── autosave.json         ← written automatically (do not edit by hand)
│   ├── profiles.py           ← saves/loads named profiles
│   └── profiles.json         ← updated after every successful search
│
├── config/
│   ├── __init__.py
│   ├── constants.py          ← countries, data types, search types
│   └── JGlogo.png            ← logo shown in the GUI header
│
├── profile/
│   └── chromium_profile/     ← Playwright persistent browser profile
│                               (created on first run; stores your login cookies)
└── results/                  ← default output folder (created on first run)
```

---

## What it does

1. You fill in up to 4 search rows (Surname / Given Name / Town / Any Field),
   choose a region, and enter 1–3 filter keywords.
2. A real Chromium browser window opens at JewishGen — the browser is visible
   so you can watch progress and log in manually if needed.
3. The script clicks every **"List N records"** button on the results page,
   scrapes each database, filters rows by keyword (OR or AND logic), and saves
   matching records to your chosen output folder.
4. Results: one `.docx` per database, plus one multi-sheet `.xlsx` workbook.

### Autosave
All settings (email, password, country, search rows, keywords, output folder,
output format) are automatically saved to `storage/autosave.json` on search
start, search finish, and window close. They are restored on next launch.

### Profiles
`storage/profiles.json` is updated after every successful search. You can
manually add named profiles there for quick switching.

### 524 / rate-limit handling
JewishGen is behind Cloudflare. The script automatically:
- Waits 2 minutes when a 524 error is detected, then retries (up to 5 times)
- Adds random human-like pauses between databases (3–8 s) and between pages (2–4 s)
- Hides the Playwright automation fingerprint from Cloudflare

---

## Installation (development / run from source)

### Requirements
- Python 3.11 or 3.12 (3.13+ not yet tested with PySide6)
- Windows 10/11 or macOS 12+

### Steps

```
# 1. Clone or unzip the project
cd JewishGen_python

# 2. Create virtual environment
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

# 3. Install Python dependencies
pip install -r requirements.txt

# 4. Install Playwright's Chromium browser
playwright install chromium

# 5. Run
python app.py
```

### First run
On first run Playwright creates `profile/chromium_profile/` — this folder
stores your JewishGen login cookies so you only need to log in once.
Log in manually in the browser window that opens. On subsequent runs
the session is remembered.

---

## Building a standalone executable

### Windows — produces `JewishGenSearch.exe`

**Prerequisites:** do everything in the venv that already has all packages installed.

```
# 1. Install PyInstaller (already in requirements.txt)
pip install pyinstaller

# 2. Build
pyinstaller ^
  --onedir ^
  --windowed ^
  --name JewishGenSearch ^
  --add-data "config;config" ^
  --add-data "storage;storage" ^
  --add-data "models;models" ^
  --add-data "gui;gui" ^
  --hidden-import PySide6.QtSvg ^
  --hidden-import PySide6.QtSvgWidgets ^
  --hidden-import playwright ^
  --collect-all playwright ^
  app.py
```

After the build completes:

```
# 3. Copy the Playwright browser into the dist folder
#    (find your cache path first)
playwright install chromium
for /f "tokens=*" %i in ('playwright install chromium --dry-run 2^>^&1 ^| findstr /i "chromium"') do echo %i

# The browser lives in:
# C:\Users\<YOU>\AppData\Local\ms-playwright\chromium-XXXX\

# Copy it next to the exe:
xcopy /E /I "%LOCALAPPDATA%\ms-playwright\chromium-*" "dist\JewishGenSearch\ms-playwright\"
```

**Distribute** the entire `dist\JewishGenSearch\` folder — zip it up and send it.
The recipient just double-clicks `JewishGenSearch.exe`. No Python needed.

---

### macOS — produces `JewishGenSearch.app`

```
# 1. Build
pyinstaller \
  --onedir \
  --windowed \
  --name JewishGenSearch \
  --add-data "config:config" \
  --add-data "storage:storage" \
  --add-data "models:models" \
  --add-data "gui:gui" \
  --hidden-import PySide6.QtSvg \
  --hidden-import PySide6.QtSvgWidgets \
  --hidden-import playwright \
  --collect-all playwright \
  app.py

# 2. Copy the Playwright browser
# Find path:
python -c "import playwright; print(playwright.__file__)"
# Browser is in ~/Library/Caches/ms-playwright/

cp -r ~/Library/Caches/ms-playwright/chromium-* \
      dist/JewishGenSearch/JewishGenSearch.app/Contents/MacOS/ms-playwright/
```

**Distribute** by zipping `dist/JewishGenSearch/JewishGenSearch.app`.
On first launch macOS may show "unidentified developer" — right-click → Open to bypass.

---

## Notes

- **Free account**: only 2 search rows work. Rows 3 & 4 require a paid JewishGen subscription.
- **Password storage**: the password is saved to `storage/autosave.json` in plain text
  on your local machine only. It is never transmitted anywhere by this application.
- **Output files**: by default saved to `~/Downloads/JewishGen_results/`.
  Change in the GUI or browse to a different folder.
- **Slow searches**: JewishGen limits automated access. A search across many databases
  can take 30–90 minutes. The window shows live progress.

---

© Alla Khananashvili
