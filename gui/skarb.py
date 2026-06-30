"""
gui/skarb.py — AIS Skarb search window (English UI; search terms stay Russian)
"""

import asyncio, json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QPushButton, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox, QRadioButton, QButtonGroup,
)
from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtGui import QPixmap
from gui._app_icon import app_icon, app_version, make_footer, make_header, make_cancel_button, autosave_path

_HERE   = Path(__file__).resolve().parent
_ROOT   = _HERE.parent
_CONFIG = _ROOT / "config"
_SAVE   = autosave_path(".skarb_autosave.json")
_DEF_DIR = str(Path.home() / "Downloads" / "Skarb_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    import skarb_scraper as _scraper
    _SCRAPER_OK = True
except ImportError:
    _SCRAPER_OK = False


STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b0b8c8;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f8f9fb;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#2a4a7f;background:#f8f9fb;}
QLineEdit,QTextEdit{padding:4px 6px;border:1px solid #c0c8d8;
  border-radius:4px;background:white;}
QLineEdit:focus,QTextEdit:focus{border:1px solid #4472c4;}
QPushButton{padding:5px 14px;border-radius:4px;
  border:1px solid #b0b8c8;background:#eef1f7;}
QPushButton:hover{background:#dde3f0;}
QPushButton#startBtn{background:#2a4a7f;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#3a5a9f;}
QPushButton#startBtn:disabled{background:#9aabcc;}
QProgressBar{border:1px solid #c0c8d8;border-radius:4px;
  text-align:center;min-height:18px;}
QProgressBar::chunk{background:#4472c4;border-radius:3px;}
"""


class Worker(QThread):
    progress = Signal(int, str)
    finished = Signal(dict)
    request_file = Signal(object)        # ask main thread about an existing file

    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload
        self._file_choice = None
        self._file_event = threading.Event()

    def provide_file_choice(self, choice: str):
        self._file_choice = choice
        self._file_event.set()

    def _ask_file_conflict(self, names):
        """Called from the scraper thread; blocks until the GUI answers."""
        self._file_event.clear()
        self.request_file.emit(names)
        self._file_event.wait(timeout=300)
        return self._file_choice or "overwrite"

    def run(self):
        self.payload["progress"] = lambda v, t: self.progress.emit(int(v), str(t))
        self.payload["ask_file_conflict"] = self._ask_file_conflict
        try:
            result = asyncio.run(_scraper.run_scraper(**self.payload))
        except Exception as exc:
            result = {"ok": False, "message": f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)


class SkarbApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AIS Skarb — Search")
        self.setMinimumWidth(700)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._build_ui()
        self._load()

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(16, 12, 16, 12)
        outer.setSpacing(10)

        # Logo + name (big letters next to the logo)
        outer.addLayout(make_header("Skarblogo.png", "AIS Skarb", color="#2a4a7f"))
        outer.addWidget(QLabel(
            "Automated Information System of the Belarusian Archives"))

        # Surnames (search is in Russian — enter Cyrillic surnames)
        sg = QGroupBox("Surnames to search")
        sl = QVBoxLayout(sg); sl.setSpacing(6)
        sl.addWidget(QLabel(
            "Enter one or more surnames — each on a new line (in Russian).\n"
            "Search uses OR (any of the surnames)."))
        self.f_surnames = QTextEdit()
        self.f_surnames.setPlaceholderText(
            "Шендерович\nШур\nЛевин")
        self.f_surnames.setFixedHeight(90)
        sl.addWidget(self.f_surnames)
        self.f_advanced = QCheckBox("Advanced search (site's «расширенный поиск»)")
        self.f_advanced.setToolTip("Tick the site's «расширенный поиск» box (name=\"add\") "
                                   "for a broader search.")
        sl.addWidget(self.f_advanced)
        outer.addWidget(sg)

        # Filters
        fg = QGroupBox("Result filter (keywords)")
        fl = QVBoxLayout(fg); fl.setSpacing(8)

        mr = QHBoxLayout()
        mr.addWidget(QLabel("Mode:"))
        self.rb_or  = QRadioButton("OR (any word)")
        self.rb_and = QRadioButton("AND (all words)")
        self.rb_or.setChecked(True)
        mr.addWidget(self.rb_or); mr.addWidget(self.rb_and); mr.addStretch()
        fl.addLayout(mr)

        kr = QHBoxLayout(); kr.setSpacing(8)
        self.f_kw1 = QLineEdit(); self.f_kw1.setPlaceholderText("Keyword 1")
        self.f_kw2 = QLineEdit(); self.f_kw2.setPlaceholderText("Keyword 2")
        self.f_kw3 = QLineEdit(); self.f_kw3.setPlaceholderText("Keyword 3")
        kr.addWidget(self.f_kw1, 1); kr.addWidget(self.f_kw2, 1); kr.addWidget(self.f_kw3, 1)
        fl.addLayout(kr)
        outer.addWidget(fg)

        # Output
        og = QGroupBox("Output")
        ol = QVBoxLayout(og); ol.setSpacing(8)
        fr = QHBoxLayout()
        self.f_docx = QCheckBox("Word (.docx)"); self.f_docx.setChecked(True)
        self.f_xlsx = QCheckBox("Excel (.xlsx)"); self.f_xlsx.setChecked(True)
        fr.addWidget(self.f_docx); fr.addWidget(self.f_xlsx); fr.addStretch()
        ol.addLayout(fr)
        dr = QHBoxLayout(); dr.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80)
        bb.clicked.connect(self._browse)
        dr.addWidget(QLabel("Save to:"))
        dr.addWidget(self.f_folder, 1); dr.addWidget(bb)
        ol.addLayout(dr)
        outer.addWidget(og)

        # Progress
        self.pbar  = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        outer.addWidget(self.pbar)
        outer.addWidget(self.stlbl)

        # Start button
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        outer.addLayout(br)

        outer.addWidget(make_footer())

        # Wire autosave
        for w in (self.f_surnames, self.f_kw1, self.f_kw2, self.f_kw3, self.f_folder):
            if hasattr(w, "textChanged"):
                w.textChanged.connect(self._save)
            elif hasattr(w, "document"):
                w.document().contentsChanged.connect(self._save)
        self.f_docx.stateChanged.connect(self._save)
        self.f_xlsx.stateChanged.connect(self._save)
        self.rb_or.toggled.connect(self._save)

    def _browse(self):
        p = QFileDialog.getExistingDirectory(
            self, "Select output folder", self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    def _keywords(self):
        return [x.strip() for x in
                [self.f_kw1.text(), self.f_kw2.text(), self.f_kw3.text()] if x.strip()]

    def _surnames(self):
        return [s.strip() for s in self.f_surnames.toPlainText().splitlines()
                if s.strip()]

    def _fmt(self):
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _payload(self):
        return {
            "surnames":      self._surnames(),
            "keywords":      self._keywords(),
            "keyword_mode":  "AND" if self.rb_and.isChecked() else "OR",
            "advanced":      self.f_advanced.isChecked(),
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log":           print,
            "cancel_event": getattr(self, "_cancel_ev", None),
        }

    def _validate(self):
        if not self._surnames():
            QMessageBox.warning(self, "No surnames",
                                "Enter at least one surname.")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Error",
                                 "skarb_scraper.py not found.")
            return False
        return True

    def _start(self):
        if not self._validate():
            return
        self._cancel_ev = threading.Event()
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.pbar.setValue(0)
        self.stlbl.setText("Starting...")
        self.worker = Worker(self._payload())
        self.worker.progress.connect(
            lambda v, t: (self.pbar.setValue(v), self.stlbl.setText(t)))
        self.worker.request_file.connect(self._show_file_conflict)
        self.worker.finished.connect(self._done)
        self.worker.start()

    def _show_file_conflict(self, names):
        """A result file already exists — ask Overwrite / Append (new file) / Skip."""
        nm = ", ".join(names) if isinstance(names, (list, tuple)) else str(names)
        box = QMessageBox(self)
        box.setWindowTitle("File already exists")
        box.setIcon(QMessageBox.Question)
        box.setText(f"File already exists:\n{nm}")
        ow = box.addButton("Overwrite", QMessageBox.AcceptRole)
        ap = box.addButton("Append (new file)", QMessageBox.ActionRole)
        sk = box.addButton("Skip", QMessageBox.RejectRole)
        box.exec()
        clicked = box.clickedButton()
        choice = "overwrite" if clicked is ow else "append" if clicked is ap else "skip"
        if self.worker:
            self.worker.provide_file_choice(choice)

    def _done(self, r: dict):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        if r.get("ok"):
            n = r.get("n_records", 0)
            msg = f"{n} record(s)"
            if r.get("output_folder"):
                msg += f"\n\nFolder:\n{r['output_folder']}"
            QMessageBox.information(self, "Done", msg)
            self.stlbl.setText("Done.")
        else:
            QMessageBox.critical(self, "Error",
                                 r.get("message", "Error — see terminal."))
            self.stlbl.setText("Error.")

    def _save(self, *_):
        try:
            d = {
                "surnames":      self.f_surnames.toPlainText(),
                "kw1": self.f_kw1.text(), "kw2": self.f_kw2.text(),
                "kw3": self.f_kw3.text(),
                "mode_and":      self.rb_and.isChecked(),
                "advanced":      self.f_advanced.isChecked(),
                "fmt_docx":      self.f_docx.isChecked(),
                "fmt_xlsx":      self.f_xlsx.isChecked(),
                "output_folder": self.f_folder.text(),
            }
            _SAVE.write_text(
                json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load(self):
        if not _SAVE.exists():
            return
        try:
            d = json.loads(_SAVE.read_text(encoding="utf-8"))
            if d.get("surnames"):
                self.f_surnames.setPlainText(d["surnames"])
            self.f_kw1.setText(d.get("kw1", ""))
            self.f_kw2.setText(d.get("kw2", ""))
            self.f_kw3.setText(d.get("kw3", ""))
            if d.get("mode_and"):
                self.rb_and.setChecked(True)
            self.f_advanced.setChecked(d.get("advanced", False))
            self.f_docx.setChecked(d.get("fmt_docx", True))
            self.f_xlsx.setChecked(d.get("fmt_xlsx", True))
            if d.get("output_folder"):
                self.f_folder.setText(d["output_folder"])
        except Exception:
            pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = SkarbApp()
    w.show()
    sys.exit(app.exec())
