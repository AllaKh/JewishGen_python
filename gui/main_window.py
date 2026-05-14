"""
gui/main_window.py
------------------
Main application window for JewishGen Mass Search.

Features
--------
- SVG eye-icon toggle on the password field (real graphic, not emoji)
- Auto-save on every search start, search finish, and window close
  → restores ALL fields (email, password, country, rows, keywords,
    output folder, output format) on next launch
- profiles.json updated after every successful search
- AND / OR keyword mode toggle
- Paid-feature guard with friendly popup
"""

import base64
import sys
import asyncio
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QLineEdit, QPushButton,
    QCheckBox, QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox,
)
from PySide6.QtCore import QThread, Signal, Qt, QByteArray
from PySide6.QtGui import QPixmap, QIcon, QImage
from PySide6.QtSvgWidgets import QSvgWidget
from PySide6.QtSvg import QSvgRenderer

# ---------------------------------------------------------------------------
# SVG eye icons embedded as base64 — no external image files needed
# ---------------------------------------------------------------------------
_EYE_OPEN_SVG = b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
  fill="none" stroke="#555" stroke-width="2"
  stroke-linecap="round" stroke-linejoin="round">
  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
  <circle cx="12" cy="12" r="3"/>
</svg>"""

_EYE_CLOSED_SVG = b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
  fill="none" stroke="#555" stroke-width="2"
  stroke-linecap="round" stroke-linejoin="round">
  <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8
           a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4
           c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19
           m-6.72-1.07a3 3 0 1 1-4.24-4.24"/>
  <line x1="1" y1="1" x2="23" y2="23"/>
</svg>"""


def _svg_to_icon(svg_bytes: bytes, size: int = 20) -> QIcon:
    """Render an SVG byte-string to a QIcon at *size* × *size* px."""
    renderer = QSvgRenderer(QByteArray(svg_bytes))
    from PySide6.QtGui import QPixmap, QPainter
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    painter = QPainter(pix)
    renderer.render(painter)
    painter.end()
    return QIcon(pix)


# ---------------------------------------------------------------------------
# Constants — config package first, scraper's own lists as fallback
# ---------------------------------------------------------------------------
try:
    from config import constants as C
    DATA_TYPES   = C.DATA_TYPES
    SEARCH_TYPES = C.SEARCH_TYPES
    COUNTRIES    = C.COUNTRIES
except Exception:
    import scraper as _s
    DATA_TYPES   = _s.DATA_TYPES
    SEARCH_TYPES = _s.SEARCH_TYPES
    COUNTRIES    = _s.COUNTRIES

import scraper

# ---------------------------------------------------------------------------
# Storage layer — storage/ and models/ live at project root
# ---------------------------------------------------------------------------
_HERE = Path(__file__).resolve().parent   # …/gui/
_ROOT = _HERE.parent                      # …/project root

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    from storage import autosave as _autosave
    _AUTOSAVE_OK = True
except Exception:
    _AUTOSAVE_OK = False

try:
    from storage import profiles as _profiles
    _PROFILES_OK = True
except Exception:
    _PROFILES_OK = False

try:
    from models.search_models import SearchProfile, SearchRow
    _MODELS_OK = True
except Exception:
    _MODELS_OK = False


# ═══════════════════════════════════════════════════  WORKER THREAD  ════════ #

class Worker(QThread):
    progress = Signal(int, str)
    finished = Signal(dict)

    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload

    def run(self):
        def _cb(v, text):
            self.progress.emit(int(v), str(text))
        self.payload["progress"] = _cb
        result = asyncio.run(scraper.run_scraper(**self.payload))
        self.finished.emit(result)


# ═══════════════════════════════════════════════════  STYLE  ═════════════════#

_DEFAULT_FOLDER = str(Path.home() / "Downloads" / "JewishGen_results")

STYLE = """
QGroupBox {
    font-weight: bold;
    font-size: 11px;
    border: 1px solid #b0b8c8;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 6px;
    background: #f8f9fb;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 4px;
    color: #2a4a7f;
    background: #f8f9fb;
}
QLineEdit, QComboBox {
    padding: 4px 6px;
    border: 1px solid #c0c8d8;
    border-radius: 4px;
    background: white;
    min-height: 22px;
}
QLineEdit:focus, QComboBox:focus { border: 1px solid #4472c4; }
QPushButton {
    padding: 5px 14px;
    border-radius: 4px;
    border: 1px solid #b0b8c8;
    background: #eef1f7;
}
QPushButton:hover   { background: #dde3f0; }
QPushButton:pressed { background: #ccd3e8; }
QPushButton#startBtn {
    background: #2a4a7f; color: white;
    font-weight: bold; font-size: 13px;
    padding: 8px 20px; border: none; border-radius: 5px;
}
QPushButton#startBtn:hover    { background: #3a5a9f; }
QPushButton#startBtn:disabled { background: #9aabcc; }
QPushButton#eyeBtn {
    border: none; background: transparent; padding: 0;
}
QPushButton#eyeBtn:hover { background: #e0e4ef; border-radius: 3px; }
QProgressBar {
    border: 1px solid #c0c8d8; border-radius: 4px;
    text-align: center; min-height: 18px;
}
QProgressBar::chunk { background: #4472c4; border-radius: 3px; }
QCheckBox { spacing: 5px; }
QLabel#note { color: #666; font-size: 10px; font-style: italic; }
"""


# ═══════════════════════════════════════════════  PASSWORD WIDGET  ═══════════#

class PasswordLineEdit(QLineEdit):
    """Password field with a real SVG eye icon that toggles visibility."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setEchoMode(QLineEdit.Password)

        self._icon_open   = _svg_to_icon(_EYE_OPEN_SVG,   20)
        self._icon_closed = _svg_to_icon(_EYE_CLOSED_SVG, 20)

        self._eye = QPushButton(self)
        self._eye.setObjectName("eyeBtn")
        self._eye.setIcon(self._icon_closed)
        self._eye.setIconSize(self._eye.sizeHint())
        self._eye.setFixedSize(28, 28)
        self._eye.setCursor(Qt.PointingHandCursor)
        self._eye.setToolTip("Show / hide password")
        self._eye.setCheckable(True)
        self._eye.toggled.connect(self._toggle)

    def _toggle(self, visible: bool):
        if visible:
            self.setEchoMode(QLineEdit.Normal)
            self._eye.setIcon(self._icon_open)
        else:
            self.setEchoMode(QLineEdit.Password)
            self._eye.setIcon(self._icon_closed)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        b = 28
        self._eye.move(self.width() - b - 2, (self.height() - b) // 2)
        self._eye.resize(b, b)
        self.setTextMargins(0, 0, b + 4, 0)


# ═══════════════════════════════════════════════  MAIN WINDOW  ═══════════════#

class JewishGenApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("JewishGen Mass Search")
        self.setMinimumWidth(840)
        self.setStyleSheet(STYLE)
        self._build_ui()
        self._load_autosave()   # restore last session on startup

    # ── BUILD UI ─────────────────────────────────────────────────────────── #

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(14, 10, 14, 10)
        outer.setSpacing(10)

        # Logo
        logo = QLabel()
        pix = QPixmap(str(_ROOT / "config" / "JGlogo.png"))
        if not pix.isNull():
            logo.setPixmap(pix.scaledToWidth(180, Qt.SmoothTransformation))
        logo.setAlignment(Qt.AlignLeft)
        outer.addWidget(logo)

        # ── Credentials ──────────────────────────────────────────────────── #
        creds = QGroupBox("Account credentials  (optional — used for auto-login)")
        cl = QHBoxLayout(creds); cl.setSpacing(8)
        self.email    = QLineEdit(); self.email.setPlaceholderText("JewishGen email")
        self.password = PasswordLineEdit(); self.password.setPlaceholderText("Password")
        cl.addWidget(QLabel("Email:")); cl.addWidget(self.email, 2)
        cl.addWidget(QLabel("Password:")); cl.addWidget(self.password, 2)
        outer.addWidget(creds)

        # ── Country ──────────────────────────────────────────────────────── #
        country_box = QGroupBox("Search region / country")
        cl2 = QHBoxLayout(country_box)
        self.country = QComboBox(); self.country.addItems(COUNTRIES)
        cl2.addWidget(self.country); cl2.addStretch()
        outer.addWidget(country_box)

        # ── Search rows ──────────────────────────────────────────────────── #
        rows_box = QGroupBox("Search fields  (enter at least one)")
        rl = QVBoxLayout(rows_box); rl.setSpacing(6)
        self.rows = []
        for i in range(4):
            rh = QHBoxLayout(); rh.setSpacing(6)
            lbl = QLabel(f"Row {i+1}"); lbl.setFixedWidth(42)
            dt  = QComboBox(); dt.addItems(DATA_TYPES);   dt.setMinimumWidth(115)
            st  = QComboBox(); st.addItems(SEARCH_TYPES); st.setMinimumWidth(145)
            txt = QLineEdit(); txt.setPlaceholderText(f"Search text — row {i+1}")
            rh.addWidget(lbl); rh.addWidget(dt); rh.addWidget(st); rh.addWidget(txt, 1)
            rl.addLayout(rh)
            self.rows.append((dt, st, txt))
        note = QLabel("★  Free account: up to 2 search rows.  "
                       "Rows 3 & 4 require a paid subscription.")
        note.setObjectName("note")
        rl.addWidget(note)
        outer.addWidget(rows_box)

        # ── Filters ──────────────────────────────────────────────────────── #
        filters_box = QGroupBox("Filters  (result rows must match keywords)")
        fl = QVBoxLayout(filters_box); fl.setSpacing(8)

        mr = QHBoxLayout()
        mr.addWidget(QLabel("Match mode:"))
        self.mode_or  = QPushButton("OR");  self.mode_or.setCheckable(True)
        self.mode_and = QPushButton("AND"); self.mode_and.setCheckable(True)
        self.mode_or.setChecked(True)
        self.mode_or.setFixedWidth(54); self.mode_and.setFixedWidth(54)
        self.mode_or.clicked.connect( lambda: self._set_mode("OR"))
        self.mode_and.clicked.connect(lambda: self._set_mode("AND"))
        mr.addWidget(self.mode_or); mr.addWidget(self.mode_and)
        mr.addWidget(QLabel("  OR = any keyword matches,   AND = all must match"))
        mr.addStretch()
        fl.addLayout(mr)

        kr = QHBoxLayout(); kr.setSpacing(8)
        self.f1 = QLineEdit(); self.f1.setPlaceholderText("Keyword 1")
        self.f2 = QLineEdit(); self.f2.setPlaceholderText("Keyword 2  (optional)")
        self.f3 = QLineEdit(); self.f3.setPlaceholderText("Keyword 3  (optional)")
        kr.addWidget(self.f1, 1); kr.addWidget(self.f2, 1); kr.addWidget(self.f3, 1)
        fl.addLayout(kr)
        outer.addWidget(filters_box)

        # ── Output ───────────────────────────────────────────────────────── #
        output_box = QGroupBox("Output")
        ol = QVBoxLayout(output_box); ol.setSpacing(8)

        fmt_row = QHBoxLayout()
        self.docx_cb = QCheckBox("Word (.docx)  — one file per database")
        self.xlsx_cb = QCheckBox("Excel (.xlsx)  — one workbook, sheet per database")
        self.docx_cb.setChecked(True); self.xlsx_cb.setChecked(True)
        fmt_row.addWidget(self.docx_cb); fmt_row.addWidget(self.xlsx_cb)
        fmt_row.addStretch()
        ol.addLayout(fmt_row)

        folder_row = QHBoxLayout(); folder_row.setSpacing(6)
        self.folder = QLineEdit(); self.folder.setText(_DEFAULT_FOLDER)
        browse_btn  = QPushButton("Browse…"); browse_btn.setFixedWidth(80)
        browse_btn.clicked.connect(self._browse)
        folder_row.addWidget(QLabel("Save to:"))
        folder_row.addWidget(self.folder, 1)
        folder_row.addWidget(browse_btn)
        ol.addLayout(folder_row)
        outer.addWidget(output_box)

        # ── Progress ─────────────────────────────────────────────────────── #
        self.progress_bar = QProgressBar(); self.progress_bar.setValue(0)
        outer.addWidget(self.progress_bar)
        self.status_lbl = QLabel("Ready"); self.status_lbl.setAlignment(Qt.AlignLeft)
        outer.addWidget(self.status_lbl)

        # ── Start button ─────────────────────────────────────────────────── #
        btn_row = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        btn_row.addStretch(); btn_row.addWidget(self.start_btn); btn_row.addStretch()
        outer.addLayout(btn_row)

        outer.addWidget(QLabel("© Alla Khananashvili", alignment=Qt.AlignRight))

    # ── AUTOSAVE / PROFILES ──────────────────────────────────────────────── #

    def _current_profile(self) -> "SearchProfile":
        return SearchProfile(
            email=self.email.text().strip(),
            password=self.password.text(),          # intentionally saved
            country=self.country.currentText(),
            keywords=self._collect_keywords(),
            rows=[
                SearchRow(dt.currentText(), st.currentText(), txt.text().strip())
                for dt, st, txt in self.rows
                if txt.text().strip()
            ],
            output_folder=self.folder.text().strip(),
            output_format=self._output_format(),
        )

    def _apply_profile(self, p: "SearchProfile"):
        """Fill all GUI widgets from a SearchProfile."""
        if p.email:    self.email.setText(p.email)
        if p.password: self.password.setText(p.password)

        idx = self.country.findText(p.country or "ALL COUNTRIES")
        if idx >= 0: self.country.setCurrentIndex(idx)

        for i, (dt, st, txt) in enumerate(self.rows):
            txt.clear()
            if p.rows and i < len(p.rows):
                r = p.rows[i]
                di = dt.findText(r.data_type or "")
                if di >= 0: dt.setCurrentIndex(di)
                si = st.findText(r.search_type or "")
                if si >= 0: st.setCurrentIndex(si)
                txt.setText(r.text or "")

        kws = p.keywords or []
        self.f1.setText(kws[0] if len(kws) > 0 else "")
        self.f2.setText(kws[1] if len(kws) > 1 else "")
        self.f3.setText(kws[2] if len(kws) > 2 else "")

        fmt = (p.output_format or "both").lower()
        self.docx_cb.setChecked(fmt in ("docx", "both"))
        self.xlsx_cb.setChecked(fmt in ("xlsx", "both"))

        if p.output_folder:
            self.folder.setText(p.output_folder)

    def _load_autosave(self):
        if not (_AUTOSAVE_OK and _MODELS_OK):
            return
        try:
            p = _autosave.load()
            if p:
                self._apply_profile(p)
        except Exception:
            pass

    def _save_autosave(self):
        """Persist current state to storage/autosave.json."""
        if not (_AUTOSAVE_OK and _MODELS_OK):
            return
        try:
            _autosave.save(self._current_profile())
        except Exception:
            pass

    def _update_profiles(self):
        """Append (or update) current profile in storage/profiles.json."""
        if not (_PROFILES_OK and _MODELS_OK):
            return
        try:
            current = self._current_profile()
            all_profiles = _profiles.load_all()
            # Match by email; replace if found, otherwise prepend
            matched = False
            for i, p in enumerate(all_profiles):
                if p.email and p.email == current.email:
                    all_profiles[i] = current
                    matched = True
                    break
            if not matched:
                all_profiles.insert(0, current)
            _profiles.save_all(all_profiles)
        except Exception:
            pass

    def closeEvent(self, event):
        self._save_autosave()
        super().closeEvent(event)

    # ── HELPERS ──────────────────────────────────────────────────────────── #

    def _set_mode(self, mode):
        self.mode_or.setChecked(mode == "OR")
        self.mode_and.setChecked(mode == "AND")

    def _keyword_mode(self):
        return "AND" if self.mode_and.isChecked() else "OR"

    def _browse(self):
        path = QFileDialog.getExistingDirectory(
            self, "Select output folder", self.folder.text() or _DEFAULT_FOLDER)
        if path:
            self.folder.setText(path)

    def _collect_rows(self):
        return [(dt.currentText(), st.currentText(), txt.text().strip())
                for dt, st, txt in self.rows if txt.text().strip()]

    def _collect_keywords(self):
        return [x.strip() for x in
                [self.f1.text(), self.f2.text(), self.f3.text()] if x.strip()]

    def _output_format(self) -> str:
        d, x = self.docx_cb.isChecked(), self.xlsx_cb.isChecked()
        if d and x: return "both"
        if d: return "docx"
        if x: return "xlsx"
        return "both"

    def _build_payload(self) -> dict:
        return {
            "rows":          self._collect_rows(),
            "country":       self.country.currentText(),
            "keywords":      self._collect_keywords(),
            "keyword_mode":  self._keyword_mode(),
            "output_format": self._output_format(),
            "output_folder": Path(self.folder.text().strip() or _DEFAULT_FOLDER),
            "email":         self.email.text().strip() or None,
            "password":      self.password.text() or None,
            "log":           print,
            "cancel_event":  None,
        }

    # ── RUN ──────────────────────────────────────────────────────────────── #

    def _start(self):
        if not self._collect_rows():
            QMessageBox.warning(self, "No search rows",
                                "Please fill in at least one search row.")
            return
        if not self._collect_keywords():
            QMessageBox.warning(self, "No filter keywords",
                                "Please enter at least one filter keyword.")
            return
        if not self.docx_cb.isChecked() and not self.xlsx_cb.isChecked():
            QMessageBox.warning(self, "No output format",
                                "Select at least one output format.")
            return

        self._save_autosave()          # save before start (crash safety)
        self.start_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_lbl.setText("Starting…")

        self.worker = Worker(self._build_payload())
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_done)
        self.worker.start()

    def _on_progress(self, v: int, text: str):
        self.progress_bar.setValue(v)
        self.status_lbl.setText(text)

    def _on_done(self, result: dict):
        self.start_btn.setEnabled(True)

        if result.get("ok"):
            self._save_autosave()       # save final state
            self._update_profiles()     # update profiles.json

            parts = []
            if result.get("docx_count"): parts.append(f"{result['docx_count']} Word file(s)")
            if result.get("xlsx_path"):  parts.append("Excel workbook")
            summary = "\n".join(parts) if parts else "No matches found."
            if result.get("output_folder"):
                summary += f"\n\nSaved to:\n{result['output_folder']}"
            QMessageBox.information(self, "Done", summary)
            self.status_lbl.setText("Done.")

        elif result.get("error") == "paid_feature":
            QMessageBox.warning(self, "Paid feature",
                result.get("message",
                    "Rows 3 & 4 require a paid JewishGen account.\n"
                    "Clear those rows and try again."))
            self.status_lbl.setText("Stopped — paid feature.")

        else:
            QMessageBox.critical(self, "Error",
                "Search failed. Check the terminal window for details.")
            self.status_lbl.setText("Error — see terminal.")


# ── STANDALONE ───────────────────────────────────────────────────────────── #

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = JewishGenApp()
    w.show()
    sys.exit(app.exec())
