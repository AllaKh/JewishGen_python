"""
gui/launcher.py
---------------
Root window — shown when app.py starts.

Lists all genealogy databases. Clicking a site either:
  - opens its URL in the default browser (for direct-access sites), or
  - opens the matching PySide6 search window (for JewishGen and future
    integrated scrapers).

Logo images are loaded from  config/<key>.png  (same folder as JGlogo.png).
If a logo file is missing, a placeholder with the site name is shown instead.
"""

import sys
import webbrowser
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QScrollArea, QFrame, QApplication,
    QSizePolicy,
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QPixmap, QFont, QColor, QPalette

# ---------------------------------------------------------------------------
_HERE = Path(__file__).resolve().parent   # …/gui/
_ROOT = _HERE.parent                      # project root
_CONFIG = _ROOT / "config"


# ---------------------------------------------------------------------------
# Site catalogue
# ---------------------------------------------------------------------------
#  key        – used to find  config/<key>.png  for the logo
#  name       – display name
#  subtitle   – short description shown below the name
#  action     – "open_url"  → webbrowser.open(url)
#              "open_window" → launch a PySide6 window class
#  url        – for open_url sites
#  window_cls – dotted import path, e.g. "gui.jewishgen.JewishGenApp"
#               (resolved lazily so missing deps don't block the launcher)
# ---------------------------------------------------------------------------
SITES = [
    {
        "key":        "JGlogo",
        "name":       "JewishGen",
        "subtitle":   "Global Home of Jewish Genealogy — mass search across all databases",
        "action":     "open_window",
        "window_cls": "gui.jewishgen.JewishGenApp",
    },
    {
        "key":        "myheritage",
        "name":       "MyHeritage",
        "subtitle":   "Family trees, DNA matching, historical records",
        "action":     "open_url",
        "url":        "https://www.myheritage.com",
    },
    {
        "key":        "familysearch",
        "name":       "FamilySearch",
        "subtitle":   "Free genealogy records from The Church of Jesus Christ",
        "action":     "open_url",
        "url":        "https://www.familysearch.org",
    },
    {
        "key":        "ancestry",
        "name":       "Ancestry",
        "subtitle":   "World's largest online family history resource",
        "action":     "open_url",
        "url":        "https://www.ancestry.com",
    },
    {
        "key":        "pamyat-naroda",
        "name":       "Память народа",
        "subtitle":   "Обобщённый банк данных участников Великой Отечественной войны",
        "action":     "open_url",
        "url":        "https://pamyat-naroda.ru",
    },
    {
        "key":        "gwar",
        "name":       "Памяти героев Великой войны",
        "subtitle":   "Участники Первой мировой войны — gwar.mil.ru",
        "action":     "open_url",
        "url":        "https://gwar.mil.ru",
    },
    {
        "key":        "memorial",
        "name":       "Мемориал",
        "subtitle":   "База данных жертв политических репрессий — memsearch.org/ru",
        "action":     "open_url",
        "url":        "https://memsearch.org/ru",
    },
]


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
_LAUNCHER_STYLE = """
QMainWindow, QWidget#root {
    background: #f0f2f7;
}
QScrollArea {
    border: none;
    background: transparent;
}
QFrame#card {
    background: white;
    border: 1px solid #d4d9e8;
    border-radius: 10px;
}
QFrame#card:hover {
    border: 1px solid #4472c4;
    background: #f5f7fd;
}
QLabel#siteName {
    font-size: 15px;
    font-weight: bold;
    color: #1a2f5e;
}
QLabel#siteSub {
    font-size: 11px;
    color: #5a6a8a;
}
QLabel#logoPlaceholder {
    font-size: 11px;
    color: #8895b0;
    font-style: italic;
}
QPushButton#openBtn {
    background: #2a4a7f;
    color: white;
    font-weight: bold;
    font-size: 12px;
    border: none;
    border-radius: 5px;
    padding: 6px 18px;
    min-width: 90px;
}
QPushButton#openBtn:hover    { background: #3a5a9f; }
QPushButton#openBtn:pressed  { background: #1a3a6f; }
QPushButton#openBtn[type="url"] {
    background: #287a3c;
}
QPushButton#openBtn[type="url"]:hover { background: #359950; }
"""

_HEADER_STYLE = """
    font-size: 22px;
    font-weight: bold;
    color: #1a2f5e;
    padding: 4px 0 2px 0;
"""
_SUBHEADER_STYLE = "font-size: 12px; color: #5a6a8a;"


# ---------------------------------------------------------------------------
# Card widget
# ---------------------------------------------------------------------------
class SiteCard(QFrame):
    """One row in the launcher list."""

    def __init__(self, site: dict, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.site = site
        self._win = None      # keep a reference so the window isn't GC'd

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(14)

        # ── Logo ─────────────────────────────────────────────────────────── #
        logo_lbl = QLabel()
        logo_lbl.setFixedSize(QSize(120, 56))
        logo_lbl.setAlignment(Qt.AlignCenter)
        pix = QPixmap(str(_CONFIG / f"{site['key']}.png"))
        if not pix.isNull():
            logo_lbl.setPixmap(
                pix.scaled(120, 56, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            )
        else:
            logo_lbl.setText(site["name"])
            logo_lbl.setObjectName("logoPlaceholder")
            logo_lbl.setStyleSheet(
                "border: 1px dashed #c0c8d8; border-radius: 4px;"
                "font-size: 11px; color: #8895b0; font-style: italic;"
            )
        layout.addWidget(logo_lbl)

        # ── Text ─────────────────────────────────────────────────────────── #
        text_col = QVBoxLayout()
        text_col.setSpacing(2)
        name_lbl = QLabel(site["name"])
        name_lbl.setObjectName("siteName")
        sub_lbl  = QLabel(site.get("subtitle", ""))
        sub_lbl.setObjectName("siteSub")
        sub_lbl.setWordWrap(True)
        text_col.addWidget(name_lbl)
        text_col.addWidget(sub_lbl)
        layout.addLayout(text_col, 1)

        # ── Button ───────────────────────────────────────────────────────── #
        if site["action"] == "open_window":
            btn_label = "Open"
            btn_type  = "window"
        else:
            btn_label = "Open website"
            btn_type  = "url"

        btn = QPushButton(btn_label)
        btn.setObjectName("openBtn")
        btn.setProperty("type", btn_type)
        btn.setStyle(btn.style())   # force stylesheet re-evaluation for property
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self._on_click)
        layout.addWidget(btn)

    def _on_click(self):
        site = self.site
        if site["action"] == "open_url":
            webbrowser.open(site["url"])
        elif site["action"] == "open_window":
            cls_path = site.get("window_cls", "")
            try:
                module_path, cls_name = cls_path.rsplit(".", 1)
                import importlib
                mod = importlib.import_module(module_path)
                cls = getattr(mod, cls_name)
                # Keep a reference on self so the window survives
                self._win = cls()
                self._win.show()
            except Exception as exc:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.critical(
                    self,
                    "Could not open window",
                    f"Failed to open {site['name']}:\n\n{exc}\n\n"
                    "Make sure all dependencies are installed.",
                )


# ---------------------------------------------------------------------------
# Launcher main window
# ---------------------------------------------------------------------------
class LauncherWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Genealogy Search — Choose a Database")
        self.setMinimumWidth(700)
        self.resize(760, 560)
        self.setStyleSheet(_LAUNCHER_STYLE)
        self._build_ui()

    def _build_ui(self):
        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(24, 18, 24, 18)
        outer.setSpacing(10)

        # Header
        hdr = QLabel("Genealogy Search")
        hdr.setStyleSheet(_HEADER_STYLE)
        outer.addWidget(hdr)

        sub = QLabel("Select a database to search")
        sub.setStyleSheet(_SUBHEADER_STYLE)
        outer.addWidget(sub)

        # Divider
        line = QFrame(); line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color: #d4d9e8;")
        outer.addWidget(line)

        # Scrollable card list
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        container = QWidget()
        card_layout = QVBoxLayout(container)
        card_layout.setContentsMargins(0, 4, 0, 4)
        card_layout.setSpacing(8)

        for site in SITES:
            card = SiteCard(site)
            card_layout.addWidget(card)

        card_layout.addStretch(1)
        scroll.setWidget(container)
        outer.addWidget(scroll, 1)

        # Footer
        footer = QLabel("© Alla Khananashvili")
        footer.setAlignment(Qt.AlignRight)
        footer.setStyleSheet("color: #9aabcc; font-size: 10px;")
        outer.addWidget(footer)


# ---------------------------------------------------------------------------
# Standalone run
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = LauncherWindow()
    w.show()
    sys.exit(app.exec())
