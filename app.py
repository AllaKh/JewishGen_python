"""
app.py — entry point

Opens the Launcher window, which lists all genealogy databases.
Integrated scrapers: JewishGen, MyHeritage, FamilySearch.
All other sites open in the default browser.
"""
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

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
