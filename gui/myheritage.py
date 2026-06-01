"""
gui/myheritage.py
-----------------
MyHeritage Search window.

Mirrors the JewishGen window style exactly:
  • logo at the top  (config/myheritage.png)
  • email / password credentials group
  • search fields: First/Patronymic name and Surname
  • filter selector: All Records / Historical Records / Family Trees
  • output format checkboxes (docx / xlsx) + folder picker
  • progress bar + status label
  • START SEARCH button

Only results with a match score ≥ 80 % are saved.
For each qualified result the scraper opens the detail page and collects:
  - full name
  - category (record collection name)
  - all data from the detail table
"""

import sys
import asyncio
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox, QComboBox,
)
from PySide6.QtCore import QThread, Signal, Qt, QByteArray
from PySide6.QtGui import QPixmap, QIcon

# ── SVG eye-icon for the password field (same as jewishgen.py) ─────────────── #
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
    from PySide6.QtSvg import QSvgRenderer
    from PySide6.QtGui import QPixmap, QPainter
    renderer = QSvgRenderer(QByteArray(svg_bytes))
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    painter = QPainter(pix)
    renderer.render(painter)
    painter.end()
    return QIcon(pix)


# ── Paths ──────────────────────────────────────────────────────────────────── #
_HERE = Path(__file__).resolve().parent
_ROOT = _HERE.parent
_CONFIG = _ROOT / "config"

# ── Scraper import ─────────────────────────────────────────────────────────── #
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    import myheritage_scraper as _scraper
    _SCRAPER_OK = True
except ImportError:
    _SCRAPER_OK = False

FILTER_OPTIONS = ["All Records", "Historical Records", "Family Trees"]
_DEFAULT_FOLDER = str(Path.home() / "Downloads" / "MyHeritage_results")

# ── Stylesheet (identical palette to JewishGen window) ─────────────────────── #
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


# ═══════════════════════════════════════════════  PASSWORD WIDGET  ══════════ #

class PasswordLineEdit(QLineEdit):
    """Password field with a real SVG eye icon — identical to JewishGen."""

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


# ═══════════════════════════════════════════════  WORKER THREAD  ═════════════#

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
        try:
            result = asyncio.run(_scraper.run_scraper(**self.payload))
        except Exception as exc:
            result = {
                "ok":      False,
                "error":   "exception",
                "message": f"{type(exc).__name__}: {exc}",
            }
        self.finished.emit(result)


# ═══════════════════════════════════════════════  MAIN WINDOW  ═══════════════#

class MyHeritageApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("MyHeritage Search")
        self.setMinimumWidth(780)
        self.setStyleSheet(STYLE)
        self._build_ui()

    # ── BUILD UI ─────────────────────────────────────────────────────────── #

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(14, 10, 14, 10)
        outer.setSpacing(10)

        # ── Logo ─────────────────────────────────────────────────────────── #
        logo_lbl = QLabel()
        pix = QPixmap(str(_CONFIG / "myheritage.png"))
        if not pix.isNull():
            logo_lbl.setPixmap(pix.scaledToWidth(200, Qt.SmoothTransformation))
        else:
            logo_lbl.setText("MyHeritage")
            logo_lbl.setStyleSheet(
                "font-size: 22px; font-weight: bold; color: #2a4a7f;"
            )
        logo_lbl.setAlignment(Qt.AlignLeft)
        outer.addWidget(logo_lbl)

        # ── Credentials ──────────────────────────────────────────────────── #
        creds = QGroupBox("Account credentials  (required for full access)")
        cl = QHBoxLayout(creds); cl.setSpacing(8)
        self.email    = QLineEdit()
        self.email.setPlaceholderText("MyHeritage email")
        self.password = PasswordLineEdit()
        self.password.setPlaceholderText("Password")
        cl.addWidget(QLabel("Email:"));    cl.addWidget(self.email,    2)
        cl.addWidget(QLabel("Password:")); cl.addWidget(self.password, 2)
        outer.addWidget(creds)

        # ── Search fields ────────────────────────────────────────────────── #
        search_box = QGroupBox("Search fields")
        sl = QVBoxLayout(search_box); sl.setSpacing(8)

        # First name + patronymic
        fn_row = QHBoxLayout(); fn_row.setSpacing(8)
        fn_lbl = QLabel("First name / Patronymic:"); fn_lbl.setFixedWidth(170)
        self.first_name = QLineEdit()
        self.first_name.setPlaceholderText("e.g.  Иван Иванович")
        fn_row.addWidget(fn_lbl)
        fn_row.addWidget(self.first_name, 1)
        sl.addLayout(fn_row)

        # Surname
        sn_row = QHBoxLayout(); sn_row.setSpacing(8)
        sn_lbl = QLabel("Surname:"); sn_lbl.setFixedWidth(170)
        self.surname = QLineEdit()
        self.surname.setPlaceholderText("e.g.  Иванов")
        sn_row.addWidget(sn_lbl)
        sn_row.addWidget(self.surname, 1)
        sl.addLayout(sn_row)

        outer.addWidget(search_box)

        # ── Record-type filter ───────────────────────────────────────────── #
        filter_box = QGroupBox("Record type filter")
        fl = QHBoxLayout(filter_box); fl.setSpacing(12)
        self.record_filter = QComboBox()
        self.record_filter.addItems(FILTER_OPTIONS)
        fl.addWidget(QLabel("Show:"))
        fl.addWidget(self.record_filter)
        fl.addStretch()
        note = QLabel("Only results with ≥ 80 % match will be saved.")
        note.setObjectName("note")
        fl.addWidget(note)
        outer.addWidget(filter_box)

        # ── Output ───────────────────────────────────────────────────────── #
        output_box = QGroupBox("Output")
        ol = QVBoxLayout(output_box); ol.setSpacing(8)

        fmt_row = QHBoxLayout()
        self.docx_cb = QCheckBox("Word (.docx)")
        self.xlsx_cb = QCheckBox("Excel (.xlsx)")
        self.docx_cb.setChecked(True)
        self.xlsx_cb.setChecked(True)
        fmt_row.addWidget(self.docx_cb)
        fmt_row.addWidget(self.xlsx_cb)
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
        self.status_lbl = QLabel("Ready")
        self.status_lbl.setAlignment(Qt.AlignLeft)
        outer.addWidget(self.status_lbl)

        # ── Start button ─────────────────────────────────────────────────── #
        btn_row = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        btn_row.addStretch(); btn_row.addWidget(self.start_btn); btn_row.addStretch()
        outer.addLayout(btn_row)

        outer.addWidget(QLabel("© Alla Khananashvili",
                               alignment=Qt.AlignRight))

    # ── HELPERS ──────────────────────────────────────────────────────────── #

    def _browse(self):
        path = QFileDialog.getExistingDirectory(
            self, "Select output folder",
            self.folder.text() or _DEFAULT_FOLDER,
        )
        if path:
            self.folder.setText(path)

    def _output_format(self) -> str:
        d, x = self.docx_cb.isChecked(), self.xlsx_cb.isChecked()
        if d and x: return "both"
        if d:       return "docx"
        if x:       return "xlsx"
        return "both"

    def _build_payload(self) -> dict:
        return {
            "first_name":     self.first_name.text().strip(),
            "surname":        self.surname.text().strip(),
            "record_filter":  self.record_filter.currentText(),
            "output_format":  self._output_format(),
            "output_folder":  Path(self.folder.text().strip() or _DEFAULT_FOLDER),
            "email":          self.email.text().strip() or None,
            "password":       self.password.text() or None,
            "log":            print,
            "cancel_event":   None,
        }

    # ── VALIDATION ───────────────────────────────────────────────────────── #

    def _validate(self) -> bool:
        if not self.first_name.text().strip() and not self.surname.text().strip():
            QMessageBox.warning(
                self, "Nothing to search",
                "Please enter at least a first name or a surname.",
            )
            return False
        if not self.docx_cb.isChecked() and not self.xlsx_cb.isChecked():
            QMessageBox.warning(
                self, "No output format",
                "Please select at least one output format (Word or Excel).",
            )
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(
                self, "Scraper not found",
                "myheritage_scraper.py could not be imported.\n"
                "Make sure it is in the project root directory.",
            )
            return False
        return True

    # ── RUN ──────────────────────────────────────────────────────────────── #

    def _start(self):
        if not self._validate():
            return

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
            n = result.get("n_records", 0)
            parts = []
            if result.get("docx_count"):
                parts.append("Word file")
            if result.get("xlsx_path"):
                parts.append("Excel workbook")
            summary = f"{n} record(s) saved"
            if parts:
                summary += " to " + " + ".join(parts)
            if result.get("output_folder"):
                summary += f"\n\nFolder:\n{result['output_folder']}"
            QMessageBox.information(self, "Done", summary)
            self.status_lbl.setText("Done.")

        else:
            err = result.get("message", "")
            msg = (
                f"The search could not be completed.\n\n{err}\n\n"
                "Check the terminal window for the full traceback."
                if err else
                "Search failed. Check the terminal window for details."
            )
            QMessageBox.critical(self, "Error", msg)
            self.status_lbl.setText("Error — see terminal.")


# ── STANDALONE ───────────────────────────────────────────────────────────── #

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MyHeritageApp()
    w.show()
    sys.exit(app.exec())
