"""
gui/launcher.py
---------------
Root window — shown when app.py starts.

Lists all genealogy databases loaded from  config/sites.json.
Styles are loaded from  config/styles.json.

Clicking a site either:
  - opens its URL in the default browser (for direct-access sites), or
  - opens the matching PySide6 search window (for JewishGen and future
    integrated scrapers).

Logo images are loaded from  config/<key>.png.
If a logo file is missing, a placeholder with the site name is shown instead.

The window auto-sizes to fit all cards without a scrollbar.
"""

import sys
import json
import webbrowser
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame, QApplication,
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QPixmap

# ---------------------------------------------------------------------------
_HERE   = Path(__file__).resolve().parent   # …/gui/
_ROOT   = _HERE.parent                      # project root
_CONFIG = _ROOT / "config"

# ---------------------------------------------------------------------------
# Load catalogue and styles from JSON
# ---------------------------------------------------------------------------

def _load_json(path: Path) -> object:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


SITES  = _load_json(_CONFIG / "sites.json")
STYLES = _load_json(_CONFIG / "styles.json")


# ---------------------------------------------------------------------------
# Card widget
# ---------------------------------------------------------------------------
class SiteCard(QFrame):
    """One row in the launcher list."""

    # Fixed geometry — used to calculate the exact window height.
    CARD_HEIGHT    = 76    # px  (including border)
    CARD_SPACING   = 8     # px  (spacing between cards in the layout)

    def __init__(self, site: dict, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.site = site
        self._win = None      # keep a reference so the child window isn't GC'd

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
            logo_lbl.setStyleSheet(STYLES["logo_placeholder"])
        layout.addWidget(logo_lbl)

        # ── Text ─────────────────────────────────────────────────────────── #
        text_col = QVBoxLayout()
        text_col.setSpacing(2)

        name_lbl = QLabel(site["name"])
        name_lbl.setObjectName("siteName")

        sub_lbl = QLabel(site.get("subtitle", ""))
        sub_lbl.setObjectName("siteSub")
        sub_lbl.setWordWrap(True)

        text_col.addWidget(name_lbl)
        text_col.addWidget(sub_lbl)
        layout.addLayout(text_col, 1)

        # ── Button ───────────────────────────────────────────────────────── #
        is_window  = site["action"] == "open_window"
        btn_label  = "Open" if is_window else "Open website"
        btn_type   = "window" if is_window else "url"

        btn = QPushButton(btn_label)
        btn.setObjectName("openBtn")
        btn.setProperty("type", btn_type)
        btn.setStyle(btn.style())          # force stylesheet re-evaluation
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self._on_click)
        layout.addWidget(btn)

    # ------------------------------------------------------------------
    def _on_click(self):
        site = self.site
        if site["action"] == "open_url":
            webbrowser.open(site["url"])
            return

        cls_path = site.get("window_cls", "")
        try:
            module_path, cls_name = cls_path.rsplit(".", 1)
            import importlib
            mod = importlib.import_module(module_path)
            cls = getattr(mod, cls_name)
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

    # Chrome around the card list (margins + header block + footer)
    _OUTER_V_MARGINS  = 18 + 18          # top + bottom outer margins
    _HEADER_BLOCK_H   = 22 + 6 + 12 + 8 + 1 + 10  # hdr + spacing + sub + spacing + divider + spacing
    _FOOTER_BLOCK_H   = 10 + 10          # footer label + bottom spacing
    _CARDS_TOP_PAD    = 4
    _CARDS_BOT_PAD    = 4

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Genealogy Search — Choose a Database")
        self.setMinimumWidth(700)
        self.setStyleSheet(STYLES["launcher"])
        self._build_ui()
        self._fit_to_cards()

    # ------------------------------------------------------------------
    def _build_ui(self):
        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)

        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(24, 18, 24, 18)
        self._outer.setSpacing(10)

        # Header
        hdr = QLabel("Genealogy Search")
        hdr.setStyleSheet(STYLES["header"])
        self._outer.addWidget(hdr)

        sub = QLabel("Select a database to search")
        sub.setStyleSheet(STYLES["subheader"])
        self._outer.addWidget(sub)

        # Divider
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet(STYLES["divider"])
        self._outer.addWidget(line)

        # Card list (plain widget, no scroll area)
        self._card_container = QWidget()
        card_layout = QVBoxLayout(self._card_container)
        card_layout.setContentsMargins(0, 4, 0, 4)
        card_layout.setSpacing(SiteCard.CARD_SPACING)

        for site in SITES:
            card = SiteCard(site)
            card_layout.addWidget(card)

        self._outer.addWidget(self._card_container)

        # Footer
        footer = QLabel("© Alla Khananashvili")
        footer.setAlignment(Qt.AlignRight)
        footer.setStyleSheet(STYLES["footer"])
        self._outer.addWidget(footer)

    # ------------------------------------------------------------------
    def _fit_to_cards(self):
        """Resize the window so all cards are visible without scrolling."""
        n          = len(SITES)
        cards_h    = (
            n * SiteCard.CARD_HEIGHT
            + max(n - 1, 0) * SiteCard.CARD_SPACING
            + self._CARDS_TOP_PAD
            + self._CARDS_BOT_PAD
        )
        total_h = (
            self._OUTER_V_MARGINS
            + self._HEADER_BLOCK_H
            + cards_h
            + self._FOOTER_BLOCK_H
        )
        # Let Qt do one pass so sizeHint() is accurate, then snap to it.
        self._card_container.adjustSize()
        hint = self.sizeHint()
        # Use the larger of our calculated height and Qt's own hint.
        final_h = max(total_h, hint.height())
        self.resize(760, final_h)
        self.setFixedHeight(final_h)   # prevent vertical resizing


# ---------------------------------------------------------------------------
# Standalone run
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = LauncherWindow()
    w.show()
    sys.exit(app.exec())
