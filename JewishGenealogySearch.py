"""
JewishGenealogySearch.py — entry point

Opens the Launcher window, which lists all genealogy databases.
Integrated scrapers: JewishGen, MyHeritage, FamilySearch, Ancestry, and others.
All remaining sites open in the default browser.
"""
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

# Packaged build (PyInstaller): use the Playwright Chromium that build_installer.ps1
# copied next to the .exe (dist/.../ms-playwright) instead of the dev machine's cache.
# In a normal dev run (not frozen) this does nothing.
if getattr(sys, "frozen", False):
    _bundled_browsers = Path(sys.executable).resolve().parent / "ms-playwright"
    if _bundled_browsers.is_dir():
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(_bundled_browsers))

from PySide6.QtWidgets import QApplication
from gui.launcher import LauncherWindow


def main():
    app = QApplication(sys.argv)
    window = LauncherWindow()
    window.show()
    try:
        sys.exit(app.exec())
    except KeyboardInterrupt:        # Ctrl+C in the terminal → quit quietly
        sys.exit(0)


if __name__ == "__main__":
    main()
