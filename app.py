"""
app.py — entry point

Opens the Launcher window, which lists all genealogy databases.
JewishGen is the first and only fully integrated scraper so far;
all other sites open in the default browser until their own
scrapers are implemented.
"""
import sys
from pathlib import Path

# Make sure the project root is on sys.path so both  gui.*  and  scraper
# can be imported regardless of which directory the user runs from.
_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from PySide6.QtWidgets import QApplication
from gui.launcher import LauncherWindow


def main():
    app = QApplication(sys.argv)
    window = LauncherWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
