"""
gui/launcher.py
---------------
Root window — shown when JewishGenealogySearch.py starts.

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
    QLabel, QPushButton, QFrame, QApplication, QSizePolicy, QScrollArea,
)
from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import QPixmap

from gui._app_icon import app_icon

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
    CARD_HEIGHT    = 60    # px  (including border)
    CARD_SPACING   = 6     # px  (spacing between cards in the layout)

    def __init__(self, site: dict, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.site = site
        self._win = None      # keep a reference so the child window isn't GC'd

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(12)

        # ── Logo ─────────────────────────────────────────────────────────── #
        logo_lbl = QLabel()
        logo_lbl.setFixedSize(QSize(104, 44))
        logo_lbl.setAlignment(Qt.AlignCenter)
        pix = QPixmap(str(_CONFIG / f"{site['key']}.png"))
        if not pix.isNull():
            logo_lbl.setPixmap(
                pix.scaled(104, 44, Qt.KeepAspectRatio, Qt.SmoothTransformation)
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
        sub_lbl.setWordWrap(False)             # one line → deterministic card height

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

        # Fixed, uniform card height → the launcher height is deterministic and
        # the fit never locks in a too-tall window (no empty band at the bottom).
        self.setFixedHeight(self.CARD_HEIGHT)

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
    _OUTER_V_MARGINS  = 10 + 10          # top + bottom outer margins
    _HEADER_BLOCK_H   = 16 + 4 + 8 + 6 + 1 + 8  # hdr + spacing + sub + spacing + divider + spacing
    _FOOTER_BLOCK_H   = 10 + 8           # footer label + bottom spacing
    _CARDS_TOP_PAD    = 4
    _CARDS_BOT_PAD    = 4

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Genealogy Search — Choose a Database")
        self.setMinimumWidth(900)        # wide enough for logo + big title
        self.setStyleSheet(STYLES["launcher"])
        self.setWindowIcon(app_icon())
        self._build_ui()
        # Sensible starting size; the real fit happens in showEvent (the layout
        # is only live once the window is actually shown — fitting before that
        # leaves the cards un-stretched → empty band on the right until you
        # nudge the window).
        self.resize(900, 660)

    def showEvent(self, ev):
        super().showEvent(ev)
        if not getattr(self, "_fitted_once", False):
            self._fitted_once = True
            # Fit now and again on several later ticks — on a slower / high-DPI
            # machine the layout (word-wrapped card subtitles) keeps settling past
            # 120 ms, and an early fit locks in a too-tall window → empty band at
            # the bottom. The last passes run after it has fully settled.
            self._fit_to_cards()
            for _ms in (0, 120, 250, 500, 900):
                QTimer.singleShot(_ms, self._fit_to_cards)

    # ------------------------------------------------------------------
    def _build_ui(self):
        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)

        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(22, 10, 22, 10)
        self._outer.setSpacing(8)

        # Header: logo flush to the LEFT edge; bilingual title centred in the
        # space between the logo and the right edge (stretch on both sides).
        hdr_row = QHBoxLayout()
        hdr_row.setSpacing(16)
        hdr_row.setContentsMargins(0, 2, 0, 2)

        # Compact header — the old 196px logo + 36px title pushed the cards off
        # screens below 4K.
        _logo_pix = QPixmap(str(_CONFIG / "app_logo.png"))
        if not _logo_pix.isNull():
            logo_lbl = QLabel()
            logo_lbl.setPixmap(
                _logo_pix.scaledToWidth(110, Qt.SmoothTransformation))
            logo_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            hdr_row.addWidget(logo_lbl)     # flush left, no spring before it

        title_col = QVBoxLayout()
        title_col.setSpacing(1)
        title_col.setAlignment(Qt.AlignVCenter | Qt.AlignHCenter)

        en_lbl = QLabel("Genealogy Search")
        en_lbl.setAlignment(Qt.AlignHCenter)
        en_lbl.setStyleSheet(
            "font-size:25px;font-weight:bold;color:#2a4a2a;"
            "font-family:'Palatino Linotype',Palatino,Georgia,serif;"
            "letter-spacing:1px;")

        ru_lbl = QLabel("Генеалогический поиск")
        ru_lbl.setAlignment(Qt.AlignHCenter)
        ru_lbl.setStyleSheet(
            "font-size:15px;font-weight:600;color:#5a7a4a;"
            "font-family:'Segoe UI',Arial,sans-serif;letter-spacing:0.5px;")

        sub_lbl = QLabel("Select a database  ·  Выберите базу данных")
        sub_lbl.setAlignment(Qt.AlignHCenter)
        sub_lbl.setStyleSheet(
            "font-size:11px;color:#777;font-style:italic;"
            "font-family:'Segoe UI',Arial,sans-serif;")

        title_col.addWidget(en_lbl)
        title_col.addWidget(ru_lbl)
        title_col.addWidget(sub_lbl)

        hdr_row.addStretch(1)              # spring between logo and title
        hdr_row.addLayout(title_col)
        hdr_row.addStretch(1)              # spring between title and right edge
        self._outer.addLayout(hdr_row)

        # Divider
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet(STYLES["divider"])
        self._outer.addWidget(line)

        # Card list inside a scroll area: on a normal screen everything fits and
        # no scrollbar shows; on a low-resolution screen the window is capped to
        # the screen height and the cards scroll instead of being cut off.
        self._card_container = QWidget()
        # Don't let the container grow past the cards' total height — otherwise the
        # extra vertical space lands as an empty band between the last card and the
        # footer (the «too big / empty space at the bottom» bug).
        self._card_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        card_layout = QVBoxLayout(self._card_container)
        card_layout.setContentsMargins(0, 4, 0, 4)
        card_layout.setSpacing(SiteCard.CARD_SPACING)

        for site in SITES:
            card = SiteCard(site)
            card_layout.addWidget(card)

        self._scroll = QScrollArea()
        self._scroll.setWidget(self._card_container)
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._scroll.setStyleSheet("QScrollArea{background:transparent;}")

        self._outer.addWidget(self._scroll)

        # Footer
        self._footer = QLabel("© 2026 Alla Khananashvili")
        self._footer.setAlignment(Qt.AlignRight)
        self._footer.setStyleSheet(STYLES["footer"])
        self._outer.addWidget(self._footer)

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    def _fit_to_cards(self, _scr_h=None):
        """Resize the window so all cards are visible, with NO empty gap and the
        cards stretched to the full width. On screens too small for the full
        list, the window shrinks to the screen and the cards scroll."""
        # Unlock any previously-pinned size so the layout can report its true,
        # fully-settled size (a stale fixed size would otherwise stick).
        self.setMinimumHeight(0)
        self.setMaximumHeight(16777215)
        scr = QApplication.primaryScreen().availableGeometry()
        scr_h = _scr_h if _scr_h else scr.height()   # override only in tests
        final_w = min(920, scr.width() - 16)   # snug width; shrink on narrow screens
        self.setFixedWidth(final_w)
        # A QScrollArea's own sizeHint is small and unrelated to its content, so
        # the height must be pinned explicitly: max = content height (taller
        # would re-open the «empty band at the bottom» bug), min = whatever fits
        # the screen — the full content on big screens (no scrollbar), the
        # remaining space on small ones (scrollbar appears).
        cards_h = self._card_container.sizeHint().height()
        self._scroll.setMaximumHeight(cards_h + 2)
        # header + divider + footer + margins
        chrome = 184
        avail  = (scr_h - 16) - chrome
        self._scroll.setMinimumHeight(max(120, min(cards_h + 2, avail)))
        self.centralWidget().resize(final_w, self.centralWidget().height())
        self._card_container.adjustSize()
        self.centralWidget().adjustSize()
        self._outer.activate()
        # Trim to the REAL bottom of the footer (sizeHint slightly over-reports,
        # leaving an empty band). footer.geometry() is in central-widget coords;
        # the central widget fills the window (no menu/tool/status bar).
        fb = self._footer.geometry().bottom()
        final_h = (fb + self._outer.contentsMargins().bottom() + 2) if fb > 50 \
            else max(self.centralWidget().sizeHint().height(), 360)
        # Never taller than the screen, and open it HIGH (top-aligned, centred)
        # so the bottom can't run off the lower edge.
        final_h = min(final_h, scr_h - 16)
        self.resize(final_w, final_h)
        self.setFixedHeight(final_h)        # prevent vertical resizing
        self.centralWidget().resize(final_w, final_h)
        self.move(scr.x() + max(0, (scr.width() - final_w) // 2), scr.y() + 8)
        self._outer.activate()


# ---------------------------------------------------------------------------
# Standalone run
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = LauncherWindow()
    w.show()
    sys.exit(app.exec())
