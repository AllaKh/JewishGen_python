"""
gui/pogroms.py
===============
Jewish Pogroms (jewishpogroms.info) search window.

Family name is typed (latin only); everything else is dropdowns whose options
come from config/pogroms_options.json (extracted from the live site). The
"Sounds like" checkbox mirrors the site: unchecked = exact family-name match.
Logo: Pogromslogo.png. Autosave: .pogroms_autosave.json
"""

import json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QComboBox,
    QFileDialog, QProgressBar, QMessageBox, QApplication, QGroupBox,
)
from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtGui import QRegularExpressionValidator
from PySide6.QtCore import QRegularExpression

from gui._app_icon import app_icon, make_header, make_cancel_button, autosave_path

# ── Paths ─────────────────────────────────────────────────────────────────── #
_HERE    = Path(__file__).resolve().parent
_ROOT    = _HERE.parent
_CONFIG  = _ROOT / "config"
_SAVE    = autosave_path(".pogroms_autosave.json")
_DEF_DIR = str(Path.home() / "Downloads" / "Pogroms_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
try:
    import pogroms_scraper as _scraper
    _SCRAPER_OK = True
except ImportError:
    _SCRAPER_OK = False

# dropdowns: (form field key, display label)
FILTERS = [
    ("region",      "Region"),
    ("city",        "City"),
    ("rec_type",    "Record Type"),
    ("incident",    "Incident"),
    ("notes",       "Notes"),
    ("found",       "Found"),
    ("register",    "Register"),
    ("case_number", "Case"),
    ("archive",     "Archive"),
    ("is_alive",    "Is Alive"),
]


def _load_options() -> dict:
    """Dropdown options extracted from the site (config/pogroms_options.json)."""
    try:
        return json.loads((_CONFIG / "pogroms_options.json")
                          .read_text(encoding="utf-8"))
    except Exception:
        return {}


STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #c8b878;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#fbf8ef;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#8B6914;background:#fbf8ef;}
QLineEdit,QComboBox{padding:4px 6px;border:1px solid #c8b878;
  border-radius:4px;background:white;min-height:22px;}
QLineEdit:focus,QComboBox:focus{border:1px solid #8B6914;}
QPushButton{padding:5px 14px;border-radius:4px;
  border:1px solid #c8b878;background:#f5efdd;}
QPushButton:hover{background:#ece1c2;}
QPushButton#startBtn{background:#8B6914;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#a37d18;}
QPushButton#startBtn:disabled{background:#c4ad6e;}
QProgressBar{border:1px solid #c8b878;border-radius:4px;
  text-align:center;min-height:18px;}
QProgressBar::chunk{background:#8B6914;border-radius:3px;}
"""


# ── Background worker thread ──────────────────────────────────────────────── #
class Worker(QThread):
    progress     = Signal(int, str)
    finished     = Signal(dict)
    request_file = Signal(str)          # emitted when output files already exist

    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload
        self._file_choice = "overwrite"
        self._file_ev     = threading.Event()

    def provide_file_choice(self, choice: str):
        self._file_choice = choice
        self._file_ev.set()

    def run(self):
        self.payload["progress"] = lambda v, t: self.progress.emit(int(v), str(t))

        def ask_file_conflict(names):
            self._file_choice = "overwrite"
            self._file_ev.clear()
            self.request_file.emit("\n".join(names))
            self._file_ev.wait(timeout=300)
            return self._file_choice or "overwrite"

        self.payload["ask_file_conflict"] = ask_file_conflict
        try:
            result = _scraper.run_scraper(**self.payload)
        except Exception as exc:
            result = {"ok": False, "error": "exception",
                      "message": f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)


# ── Main window ───────────────────────────────────────────────────────────── #
class PogromsApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Jewish Pogroms")
        self.setMinimumWidth(860)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._options = _load_options()
        self._build_ui()
        self._load()

    # ── Build UI ──────────────────────────────────────────────────────────── #
    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(18, 12, 18, 12)
        self._outer.setSpacing(8)

        self._outer.addLayout(
            make_header("Pogromslogo.png", "Jewish Pogroms", color="#8B6914"))
        sub = QLabel("Memorial site about Jewish pogroms in the Russian Empire "
                     "during the First World War and Civil War — jewishpogroms.info")
        sub.setAlignment(Qt.AlignCenter)
        self._outer.addWidget(sub)

        # Search
        sg = QGroupBox("Search")
        sl = QGridLayout(sg); sl.setSpacing(8)
        self.f_family = QLineEdit()
        self.f_family.setPlaceholderText("Family name (latin only)")
        # the site itself accepts only A-Z, space and *
        self.f_family.setValidator(QRegularExpressionValidator(
            QRegularExpression(r"[A-Za-z *]*")))
        self.f_sounds = QCheckBox("Sounds like")
        self.f_sounds.setChecked(True)
        self.f_sounds.setToolTip(
            "Checked: similar-sounding family names.\n"
            "Unchecked: exact family-name match only.")
        sl.addWidget(QLabel("Family name:"), 0, 0)
        sl.addWidget(self.f_family,          0, 1)
        sl.addWidget(self.f_sounds,          0, 2)
        sl.setColumnStretch(1, 1)

        # the 10 site dropdowns, options from config/pogroms_options.json
        self._filters: dict = {}
        for i, (key, label) in enumerate(FILTERS):
            r, c = 1 + i // 2, (i % 2) * 2
            cb = QComboBox()
            cb.setEditable(True)            # type-to-search in long lists
            cb.addItem("")                  # empty = any
            cb.addItems(self._options.get(key, []))
            cb.setCurrentIndex(0)
            cb.setMaxVisibleItems(25)
            # don't let 200-char option texts inflate the window width
            cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            cb.setMinimumContentsLength(24)
            self._filters[key] = cb
            sl.addWidget(QLabel(label + ":"), r, c)
            sl.addWidget(cb,                  r, c + 1)
        sl.setColumnStretch(1, 1); sl.setColumnStretch(3, 1)
        self._outer.addWidget(sg)

        # Output
        og = QGroupBox("Output")
        ol = QVBoxLayout(og); ol.setSpacing(6)
        fr = QHBoxLayout()
        self.f_docx = QCheckBox("Word (.docx)");  self.f_docx.setChecked(True)
        self.f_xlsx = QCheckBox("Excel (.xlsx)"); self.f_xlsx.setChecked(True)
        fr.addWidget(self.f_docx); fr.addWidget(self.f_xlsx); fr.addStretch()
        fr.addWidget(QLabel("(Person cards go to Word only)"))
        ol.addLayout(fr)
        dr = QHBoxLayout(); dr.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80)
        bb.clicked.connect(self._browse)
        dr.addWidget(QLabel("Save to:"))
        dr.addWidget(self.f_folder, 1)
        dr.addWidget(bb)
        ol.addLayout(dr)
        self._outer.addWidget(og)

        # Progress + start
        self.pbar  = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        self._outer.addWidget(self.pbar)
        self._outer.addWidget(self.stlbl)
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(
            QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        for w in self._all_widgets():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QComboBox): w.currentTextChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)

    # ── Autosave / load ───────────────────────────────────────────────────── #
    def _all_widgets(self) -> list:
        return ([self.f_family, self.f_sounds, self.f_folder,
                 self.f_docx, self.f_xlsx]
                + list(self._filters.values()))

    def _save(self, *_):
        d = {
            "family_name":   self.f_family.text(),
            "soundslike":    self.f_sounds.isChecked(),
            "output_folder": self.f_folder.text(),
            "fmt_docx":      self.f_docx.isChecked(),
            "fmt_xlsx":      self.f_xlsx.isChecked(),
        }
        for key, cb in self._filters.items():
            d[f"filter_{key}"] = cb.currentText()
        try:
            _SAVE.write_text(
                json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load(self):
        if not _SAVE.exists():
            return
        try:
            d = json.loads(_SAVE.read_text(encoding="utf-8"))
        except Exception:
            return
        self.f_family.setText(str(d.get("family_name", "")))
        self.f_sounds.setChecked(bool(d.get("soundslike", True)))
        self.f_folder.setText(str(d.get("output_folder", _DEF_DIR)) or _DEF_DIR)
        self.f_docx.setChecked(bool(d.get("fmt_docx", True)))
        self.f_xlsx.setChecked(bool(d.get("fmt_xlsx", True)))
        for key, cb in self._filters.items():
            v = str(d.get(f"filter_{key}", ""))
            if v:
                cb.setCurrentText(v)

    # ── Helpers ───────────────────────────────────────────────────────────── #
    def _browse(self):
        p = QFileDialog.getExistingDirectory(
            self, "Select output folder", self.f_folder.text() or _DEF_DIR)
        if p: self.f_folder.setText(p)

    def _fmt(self) -> str:
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _payload(self) -> dict:
        return {
            "family_name":   self.f_family.text().strip(),
            "soundslike":    self.f_sounds.isChecked(),
            "filters":       {k: cb.currentText().strip()
                              for k, cb in self._filters.items()
                              if cb.currentText().strip()},
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log":           print,
            "cancel_event": getattr(self, "_cancel_ev", None),
        }

    def _validate(self) -> bool:
        if not self.f_family.text().strip() \
                and not any(cb.currentText().strip()
                            for cb in self._filters.values()):
            QMessageBox.warning(self, "Nothing to search",
                                "Enter a family name or pick a filter.")
            return False
        if not self.f_docx.isChecked() and not self.f_xlsx.isChecked():
            QMessageBox.warning(self, "No output format",
                                "Select at least one output format.")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Scraper not found",
                                 "pogroms_scraper.py not found in project root.")
            return False
        return True

    # ── File-conflict dialog (existing output files) ──────────────────────── #
    def _show_file_conflict_dialog(self, names: str):
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Question)
        box.setWindowTitle("File already exists")
        box.setText("These output file(s) already exist:\n\n"
                    f"{names}\n\nWhat would you like to do?")
        b_over = box.addButton("Overwrite", QMessageBox.DestructiveRole)
        b_app  = box.addButton("Append new results", QMessageBox.AcceptRole)
        b_skip = box.addButton("Skip (don't save)", QMessageBox.RejectRole)
        box.setDefaultButton(b_app)
        box.exec()
        clicked = box.clickedButton()
        if clicked is b_app:
            choice = "append"
        elif clicked is b_skip:
            choice = "skip"
        else:
            choice = "overwrite"
        self.worker.provide_file_choice(choice)

    # ── Start / finish ────────────────────────────────────────────────────── #
    def _start(self):
        if not self._validate(): return
        self._cancel_ev = threading.Event()
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.pbar.setValue(0)
        self.stlbl.setText("Starting...")
        self.worker = Worker(self._payload())
        self.worker.progress.connect(
            lambda v, t: (self.pbar.setValue(v), self.stlbl.setText(t)))
        self.worker.request_file.connect(self._show_file_conflict_dialog)
        self.worker.finished.connect(self._done)
        self.worker.start()

    def _done(self, r: dict):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        if r.get("ok"):
            n = r.get("n_records", 0)
            parts = (["Word"]  if r.get("docx_count") else []) + \
                    (["Excel"] if r.get("xlsx_path")  else [])
            msg = f"{n} record(s) saved"
            if parts: msg += " → " + " + ".join(parts)
            if r.get("output_folder"):
                msg += f"\n\nFolder:\n{r['output_folder']}"
            QMessageBox.information(self, "Done", msg)
            self.stlbl.setText("Done.")
        else:
            QMessageBox.critical(
                self, "Error",
                f"Search failed.\n\n{r.get('message','')}\n\nCheck terminal.")
            self.stlbl.setText("Error — see terminal.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = PogromsApp()
    w.show()
    sys.exit(app.exec())
