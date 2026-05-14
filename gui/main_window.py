import sys
import asyncio
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QLineEdit, QPushButton,
    QCheckBox, QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox, QSizePolicy,
)
from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtGui import QPixmap, QIcon

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


# ─────────────────────────────────────────────  WORKER  ──────────────────── #

class Worker(QThread):
    progress = Signal(int, str)
    finished = Signal(dict)

    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload

    def run(self):
        def progress_cb(v, text):
            self.progress.emit(int(v), str(text))

        self.payload["progress"] = progress_cb
        result = asyncio.run(scraper.run_scraper(**self.payload))
        self.finished.emit(result)


# ─────────────────────────────────────────────  MAIN APP  ────────────────── #

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
QLineEdit:focus, QComboBox:focus {
    border: 1px solid #4472c4;
}
QPushButton {
    padding: 5px 14px;
    border-radius: 4px;
    border: 1px solid #b0b8c8;
    background: #eef1f7;
}
QPushButton:hover  { background: #dde3f0; }
QPushButton:pressed { background: #ccd3e8; }
QPushButton#startBtn {
    background: #2a4a7f;
    color: white;
    font-weight: bold;
    font-size: 13px;
    padding: 8px 20px;
    border: none;
    border-radius: 5px;
}
QPushButton#startBtn:hover  { background: #3a5a9f; }
QPushButton#startBtn:disabled { background: #9aabcc; }
QProgressBar {
    border: 1px solid #c0c8d8;
    border-radius: 4px;
    text-align: center;
    min-height: 18px;
}
QProgressBar::chunk { background: #4472c4; border-radius: 3px; }
QCheckBox { spacing: 5px; }
QLabel#note {
    color: #666;
    font-size: 10px;
    font-style: italic;
}
"""


class PasswordLineEdit(QLineEdit):
    """QLineEdit with a show/hide eye button."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setEchoMode(QLineEdit.Password)

        self._eye_btn = QPushButton("👁", self)
        self._eye_btn.setFixedSize(28, 28)
        self._eye_btn.setCursor(Qt.PointingHandCursor)
        self._eye_btn.setStyleSheet(
            "QPushButton { border: none; background: transparent; font-size: 14px; }"
            "QPushButton:hover { background: #e8eaf0; border-radius: 3px; }"
        )
        self._eye_btn.setCheckable(True)
        self._eye_btn.toggled.connect(self._toggle_visibility)
        self._update_padding()

    def _update_padding(self):
        self.setStyleSheet("QLineEdit { padding-right: 32px; }")

    def _toggle_visibility(self, checked):
        self.setEchoMode(QLineEdit.Normal if checked else QLineEdit.Password)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        h = self.height()
        btn_size = 26
        self._eye_btn.move(self.width() - btn_size - 3, (h - btn_size) // 2)
        self._eye_btn.resize(btn_size, btn_size)


class JewishGenApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("JewishGen Mass Search")
        self.setMinimumWidth(820)
        self.setStyleSheet(STYLE)
        self._build_ui()

    # ──────────────────────────────────────────────────────────  BUILD UI  ── #

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(14, 10, 14, 10)
        outer.setSpacing(10)

        # ── LOGO ──────────────────────────────────────────────────────────── #
        logo_lbl = QLabel()
        pix = QPixmap("config/JGlogo.png")
        if not pix.isNull():
            logo_lbl.setPixmap(pix.scaledToWidth(180, Qt.SmoothTransformation))
        logo_lbl.setAlignment(Qt.AlignLeft)
        outer.addWidget(logo_lbl)

        # ── CREDENTIALS ───────────────────────────────────────────────────── #
        creds_box = QGroupBox("Account credentials  (optional — used for auto-login)")
        creds_layout = QHBoxLayout(creds_box)
        creds_layout.setSpacing(8)

        self.email = QLineEdit()
        self.email.setPlaceholderText("JewishGen email")

        self.password = PasswordLineEdit()
        self.password.setPlaceholderText("Password")

        creds_layout.addWidget(QLabel("Email:"))
        creds_layout.addWidget(self.email, 2)
        creds_layout.addWidget(QLabel("Password:"))
        creds_layout.addWidget(self.password, 2)
        outer.addWidget(creds_box)

        # ── COUNTRY ───────────────────────────────────────────────────────── #
        country_box = QGroupBox("Search region / country")
        country_layout = QHBoxLayout(country_box)
        self.country = QComboBox()
        self.country.addItems(COUNTRIES)
        country_layout.addWidget(self.country)
        country_layout.addStretch()
        outer.addWidget(country_box)

        # ── SEARCH ROWS ───────────────────────────────────────────────────── #
        rows_box = QGroupBox("Search fields  (enter at least one)")
        rows_layout = QVBoxLayout(rows_box)
        rows_layout.setSpacing(6)

        self.rows = []
        for i in range(4):
            row_h = QHBoxLayout()
            row_h.setSpacing(6)

            lbl = QLabel(f"Row {i + 1}")
            lbl.setFixedWidth(42)

            dt = QComboBox()
            dt.addItems(DATA_TYPES)
            dt.setMinimumWidth(110)

            st = QComboBox()
            st.addItems(SEARCH_TYPES)
            st.setMinimumWidth(140)

            txt = QLineEdit()
            txt.setPlaceholderText(f"Search text for row {i + 1}")

            row_h.addWidget(lbl)
            row_h.addWidget(dt)
            row_h.addWidget(st)
            row_h.addWidget(txt, 1)
            rows_layout.addLayout(row_h)

            self.rows.append((dt, st, txt))

        note = QLabel(
            "★  Free JewishGen account: up to 2 search rows. "
            "Rows 3 & 4 require a paid subscription."
        )
        note.setObjectName("note")
        rows_layout.addWidget(note)
        outer.addWidget(rows_box)

        # ── FILTERS ───────────────────────────────────────────────────────── #
        filters_box = QGroupBox("Filters  (result rows must contain keywords)")
        filters_layout = QVBoxLayout(filters_box)
        filters_layout.setSpacing(8)

        # AND / OR mode toggle
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Match mode:"))
        self.mode_or  = QPushButton("OR")
        self.mode_and = QPushButton("AND")
        for btn in (self.mode_or, self.mode_and):
            btn.setCheckable(True)
            btn.setFixedWidth(54)
        self.mode_or.setChecked(True)
        self.mode_or.clicked.connect(lambda: self._set_mode("OR"))
        self.mode_and.clicked.connect(lambda: self._set_mode("AND"))
        mode_row.addWidget(self.mode_or)
        mode_row.addWidget(self.mode_and)
        mode_row.addWidget(QLabel("  OR = any keyword matches,   AND = all keywords must match"))
        mode_row.addStretch()
        filters_layout.addLayout(mode_row)

        kw_row = QHBoxLayout()
        kw_row.setSpacing(8)
        self.f1 = QLineEdit()
        self.f1.setPlaceholderText("Keyword 1")
        self.f2 = QLineEdit()
        self.f2.setPlaceholderText("Keyword 2  (optional)")
        self.f3 = QLineEdit()
        self.f3.setPlaceholderText("Keyword 3  (optional)")
        kw_row.addWidget(self.f1, 1)
        kw_row.addWidget(self.f2, 1)
        kw_row.addWidget(self.f3, 1)
        filters_layout.addLayout(kw_row)
        outer.addWidget(filters_box)

        # ── OUTPUT ────────────────────────────────────────────────────────── #
        output_box = QGroupBox("Output")
        output_layout = QVBoxLayout(output_box)
        output_layout.setSpacing(8)

        # Format checkboxes
        fmt_row = QHBoxLayout()
        self.docx_cb = QCheckBox("Word (.docx)  — one file per database")
        self.xlsx_cb = QCheckBox("Excel (.xlsx)  — one workbook, sheet per database")
        self.docx_cb.setChecked(True)
        self.xlsx_cb.setChecked(True)
        fmt_row.addWidget(self.docx_cb)
        fmt_row.addWidget(self.xlsx_cb)
        fmt_row.addStretch()
        output_layout.addLayout(fmt_row)

        # Folder row
        folder_row = QHBoxLayout()
        folder_row.setSpacing(6)
        self.folder = QLineEdit()
        self.folder.setPlaceholderText("Output folder")
        self.folder.setText(_DEFAULT_FOLDER)   # visible default

        browse_btn = QPushButton("Browse…")
        browse_btn.setFixedWidth(80)
        browse_btn.clicked.connect(self._browse)

        folder_row.addWidget(QLabel("Save to:"))
        folder_row.addWidget(self.folder, 1)
        folder_row.addWidget(browse_btn)
        output_layout.addLayout(folder_row)
        outer.addWidget(output_box)

        # ── PROGRESS ──────────────────────────────────────────────────────── #
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        outer.addWidget(self.progress_bar)

        self.status_lbl = QLabel("Ready")
        self.status_lbl.setAlignment(Qt.AlignLeft)
        outer.addWidget(self.status_lbl)

        # ── BUTTONS ───────────────────────────────────────────────────────── #
        btn_row = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        btn_row.addStretch()
        btn_row.addWidget(self.start_btn)
        btn_row.addStretch()
        outer.addLayout(btn_row)

        # ── COPYRIGHT ─────────────────────────────────────────────────────── #
        outer.addWidget(QLabel("© Alla Khananashvili", alignment=Qt.AlignRight))

    # ──────────────────────────────────────────────────────────  HELPERS  ─── #

    def _set_mode(self, mode):
        self.mode_or.setChecked(mode == "OR")
        self.mode_and.setChecked(mode == "AND")

    def _keyword_mode(self):
        return "AND" if self.mode_and.isChecked() else "OR"

    def _browse(self):
        path = QFileDialog.getExistingDirectory(
            self, "Select output folder",
            self.folder.text() or _DEFAULT_FOLDER
        )
        if path:
            self.folder.setText(path)

    def _collect_rows(self):
        """Return list of (data_type, search_type, text) for every non-empty row."""
        rows = []
        for dt, st, txt in self.rows:
            t = txt.text().strip()
            if t:
                rows.append((dt.currentText(), st.currentText(), t))
        return rows

    def _collect_keywords(self):
        return [
            x.strip()
            for x in [self.f1.text(), self.f2.text(), self.f3.text()]
            if x.strip()
        ]

    def _output_format(self) -> str:
        want_docx = self.docx_cb.isChecked()
        want_xlsx = self.xlsx_cb.isChecked()
        if want_docx and want_xlsx:
            return "both"
        if want_docx:
            return "docx"
        if want_xlsx:
            return "xlsx"
        return "both"   # fallback: never silently produce nothing

    def _build_payload(self) -> dict:
        folder_text = self.folder.text().strip() or _DEFAULT_FOLDER
        return {
            "rows":          self._collect_rows(),
            "country":       self.country.currentText(),
            "keywords":      self._collect_keywords(),
            "keyword_mode":  self._keyword_mode(),
            "output_format": self._output_format(),
            "output_folder": Path(folder_text),
            "email":         self.email.text().strip() or None,
            "password":      self.password.text() or None,
            "log":           print,
            "cancel_event":  None,
            # 'progress' injected by Worker.run()
        }

    # ────────────────────────────────────────────────────────────  RUN  ───── #

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
                                "Select at least one output format (Word or Excel).")
            return

        self.start_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_lbl.setText("Starting…")

        self.worker = Worker(self._build_payload())
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_done)
        self.worker.start()

    # ──────────────────────────────────────────────────────────  SIGNALS  ─── #

    def _on_progress(self, v: int, text: str):
        self.progress_bar.setValue(v)
        self.status_lbl.setText(text)

    def _on_done(self, result: dict):
        self.start_btn.setEnabled(True)

        if result.get("ok"):
            folder = result.get("output_folder", "")
            docx_n = result.get("docx_count", 0)
            xlsx   = result.get("xlsx_path")

            parts = []
            if docx_n:
                parts.append(f"{docx_n} Word file(s)")
            if xlsx:
                parts.append("Excel workbook")
            summary = "\n".join(parts) if parts else "No matches found."
            if folder:
                summary += f"\n\nSaved to:\n{folder}"

            QMessageBox.information(self, "Done", summary)
            self.status_lbl.setText("Done.")

        elif result.get("error") == "paid_feature":
            QMessageBox.warning(
                self, "Paid feature",
                result.get("message",
                           "Rows 3 & 4 require a paid JewishGen account.\n"
                           "Clear those rows and try again.")
            )
            self.status_lbl.setText("Stopped — paid feature.")
        else:
            QMessageBox.critical(
                self, "Error",
                "Search failed. Check the terminal window for details."
            )
            self.status_lbl.setText("Error — see terminal.")


# ──────────────────────────────────────────────  STANDALONE  ─────────────── #

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = JewishGenApp()
    w.show()
    sys.exit(app.exec())
