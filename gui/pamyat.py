"""
gui/pamyat.py — «Память народа» (pamyat-naroda.ru) search window.

GUI is always English; a Site/Language selector (ru/en) picks which language
version of the site to search. Goes straight to the advanced search, collects
FIO-matching people, opens each, copies every document (with archive info) to
Word, saves photos/images.
"""

import asyncio, json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox, QApplication, QGroupBox,
)
from PySide6.QtCore import QThread, Signal, Qt
from gui._app_icon import app_icon, make_header

_HERE   = Path(__file__).resolve().parent
_ROOT   = _HERE.parent
_CONFIG = _ROOT / "config"
_SAVE   = _HERE / ".pamyat_autosave.json"
_DEF_DIR = str(Path.home() / "Downloads" / "Pamyat_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    import pamyat_scraper as _scraper
    _SCRAPER_OK = True
    BRANCH = _scraper.BRANCH
    AWARDS = _scraper.AWARDS
except ImportError:
    _SCRAPER_OK = False
    BRANCH = {"": ""}
    AWARDS = {"": ""}

SITE_LANG = {"Russian (ru)": "ru", "English (en)": "en"}

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b0b8c8;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f8f9fb;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#b03030;background:#f8f9fb;}
QLineEdit,QComboBox{padding:4px 6px;border:1px solid #c0c8d8;border-radius:4px;
  background:white;}
QLineEdit:focus{border:1px solid #d04040;}
QPushButton{padding:5px 14px;border-radius:4px;border:1px solid #b0b8c8;
  background:#eef1f7;}
QPushButton:hover{background:#dde3f0;}
QPushButton#startBtn{background:#b71c1c;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#c62828;}
QPushButton#startBtn:disabled{background:#d9a5a5;}
QProgressBar{border:1px solid #c0c8d8;border-radius:4px;text-align:center;
  min-height:18px;}
QProgressBar::chunk{background:#b71c1c;border-radius:3px;}
"""


class Worker(QThread):
    progress     = Signal(int, str)
    finished     = Signal(dict)
    request_file = Signal(str)

    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload
        self._file_choice = "overwrite"
        self._file_ev = threading.Event()

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
            result = asyncio.run(_scraper.run_scraper(**self.payload))
        except Exception as exc:
            result = {"ok": False, "message": f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)


class PamyatApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pamyat Naroda")
        self.setMinimumWidth(760)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._build_ui()
        self._load()

    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(16, 12, 16, 12); outer.setSpacing(10)

        outer.addLayout(make_header("Pamatlogo.png", "Pamyat Naroda", color="#b71c1c"))
        outer.addWidget(QLabel("WWII participant records (pamyat-naroda.ru) — "
                               "documents, awards, fate, scanned archives."))

        lg = QGroupBox("Site / Language")
        ll = QHBoxLayout(lg); ll.setSpacing(10)
        self.f_lang = QComboBox(); self.f_lang.addItems(list(SITE_LANG.keys()))
        ll.addWidget(QLabel("Search the site in:")); ll.addWidget(self.f_lang); ll.addStretch()
        outer.addWidget(lg)

        # Name
        ng = QGroupBox("Person")
        ngl = QGridLayout(ng); ngl.setSpacing(6)
        self.f_last  = QLineEdit(); self.f_first = QLineEdit(); self.f_mid = QLineEdit()
        self.f_byf   = QLineEdit(); self.f_byt = QLineEdit(); self.f_bplace = QLineEdit()
        self.f_byf.setPlaceholderText("e.g. 1921"); self.f_byt.setPlaceholderText("to (optional)")
        ngl.addWidget(QLabel("Surname:"), 0, 0);    ngl.addWidget(self.f_last, 0, 1)
        ngl.addWidget(QLabel("First name:"), 0, 2); ngl.addWidget(self.f_first, 0, 3)
        ngl.addWidget(QLabel("Patronymic:"), 1, 0); ngl.addWidget(self.f_mid, 1, 1)
        ngl.addWidget(QLabel("Birth year:"), 1, 2); ngl.addWidget(self.f_byf, 1, 3)
        ngl.addWidget(QLabel("Birth place:"), 2, 0); ngl.addWidget(self.f_bplace, 2, 1)
        ngl.addWidget(QLabel("…to year:"), 2, 2);   ngl.addWidget(self.f_byt, 2, 3)
        outer.addWidget(ng)

        # Service / awards
        sg = QGroupBox("Service & awards (optional)")
        sgl = QGridLayout(sg); sgl.setSpacing(6)
        self.f_rank = QLineEdit(); self.f_service = QLineEdit(); self.f_call = QLineEdit()
        self.f_branch = QComboBox(); self.f_branch.addItems(list(BRANCH.keys()))
        self.f_award  = QComboBox(); self.f_award.addItems(list(AWARDS.keys()))
        sgl.addWidget(QLabel("Rank:"), 0, 0);            sgl.addWidget(self.f_rank, 0, 1)
        sgl.addWidget(QLabel("Branch of service:"), 0, 2); sgl.addWidget(self.f_branch, 0, 3)
        sgl.addWidget(QLabel("Place of service:"), 1, 0); sgl.addWidget(self.f_service, 1, 1)
        sgl.addWidget(QLabel("Award:"), 1, 2);           sgl.addWidget(self.f_award, 1, 3)
        sgl.addWidget(QLabel("Place of conscription:"), 2, 0); sgl.addWidget(self.f_call, 2, 1)
        outer.addWidget(sg)

        self.f_group = QCheckBox("Group documents of one person")
        self.f_group.setChecked(True)
        outer.addWidget(self.f_group)

        og = QGroupBox("Output (Word)")
        ol = QHBoxLayout(og); ol.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        ol.addWidget(QLabel("Save to:")); ol.addWidget(self.f_folder, 1); ol.addWidget(bb)
        outer.addWidget(og)

        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        outer.addWidget(self.pbar); outer.addWidget(self.stlbl)

        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH"); self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn); br.addStretch()
        outer.addLayout(br)
        outer.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        for w in (self.f_last, self.f_first, self.f_mid, self.f_byf, self.f_byt,
                  self.f_bplace, self.f_rank, self.f_service, self.f_call,
                  self.f_folder):
            w.textChanged.connect(self._save)
        for cb in (self.f_lang, self.f_branch, self.f_award):
            cb.currentTextChanged.connect(self._save)
        self.f_group.stateChanged.connect(self._save)

    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Output folder",
                                             self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    def _payload(self):
        return {
            "lang":        SITE_LANG.get(self.f_lang.currentText(), "ru"),
            "last_name":   self.f_last.text().strip(),
            "first_name":  self.f_first.text().strip(),
            "middle_name": self.f_mid.text().strip(),
            "birth_from":  self.f_byf.text().strip(),
            "birth_to":    self.f_byt.text().strip(),
            "place_birth": self.f_bplace.text().strip(),
            "rank":        self.f_rank.text().strip(),
            "branch":      BRANCH.get(self.f_branch.currentText(), ""),
            "place_service": self.f_service.text().strip(),
            "place_conscription": self.f_call.text().strip(),
            "award":       AWARDS.get(self.f_award.currentText(), ""),
            "group_person": self.f_group.isChecked(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log":         print,
            "cancel_event": None,
        }

    def _validate(self):
        p = self._payload()
        if not any(p[k] for k in ("last_name", "first_name", "place_service",
                                  "rank")):
            QMessageBox.warning(self, "Nothing to search",
                                "Enter at least a surname (or another field).")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Error", "pamyat_scraper.py not found.")
            return False
        return True

    def _start(self):
        if not self._validate():
            return
        self.start_btn.setEnabled(False)
        self.pbar.setValue(0); self.stlbl.setText("Starting…")
        self._worker = Worker(self._payload())
        self._worker.progress.connect(
            lambda v, t: (self.pbar.setValue(v), self.stlbl.setText(t)))
        self._worker.finished.connect(self._done)
        self._worker.request_file.connect(self._show_file_conflict)
        self._worker.start()

    def _show_file_conflict(self, names: str):
        box = QMessageBox(self); box.setIcon(QMessageBox.Question)
        box.setWindowTitle("File already exists")
        box.setText(f"This file already exists:\n\n{names}\n\nWhat to do?")
        b_over = box.addButton("Overwrite", QMessageBox.DestructiveRole)
        b_app  = box.addButton("Append", QMessageBox.AcceptRole)
        b_skip = box.addButton("Skip", QMessageBox.RejectRole)
        box.setDefaultButton(b_app); box.exec()
        c = box.clickedButton()
        self._worker.provide_file_choice(
            "append" if c is b_app else "skip" if c is b_skip else "overwrite")

    def _done(self, r: dict):
        self.start_btn.setEnabled(True)
        if r.get("ok"):
            n = r.get("n_records", 0)
            msg = f"{n} person(s)"
            if r.get("output_folder"):
                msg += f"\n\nFolder:\n{r['output_folder']}"
            QMessageBox.information(self, "Done", msg)
            self.stlbl.setText("Done.")
        else:
            QMessageBox.critical(self, "Error",
                                 r.get("message", "Failed — see terminal."))
            self.stlbl.setText("Error.")

    def _save(self, *_):
        try:
            d = {"lang": self.f_lang.currentText(),
                 "last": self.f_last.text(), "first": self.f_first.text(),
                 "mid": self.f_mid.text(), "byf": self.f_byf.text(),
                 "byt": self.f_byt.text(), "bplace": self.f_bplace.text(),
                 "rank": self.f_rank.text(), "branch": self.f_branch.currentText(),
                 "service": self.f_service.text(), "call": self.f_call.text(),
                 "award": self.f_award.currentText(), "group": self.f_group.isChecked(),
                 "folder": self.f_folder.text()}
            _SAVE.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load(self):
        if not _SAVE.exists():
            return
        try:
            d = json.loads(_SAVE.read_text(encoding="utf-8"))
        except Exception:
            return
        self.f_last.setText(d.get("last", "")); self.f_first.setText(d.get("first", ""))
        self.f_mid.setText(d.get("mid", "")); self.f_byf.setText(d.get("byf", ""))
        self.f_byt.setText(d.get("byt", "")); self.f_bplace.setText(d.get("bplace", ""))
        self.f_rank.setText(d.get("rank", "")); self.f_service.setText(d.get("service", ""))
        self.f_call.setText(d.get("call", "")); self.f_folder.setText(d.get("folder", _DEF_DIR))
        self.f_group.setChecked(d.get("group", True))
        for cb, key in ((self.f_lang, "lang"), (self.f_branch, "branch"),
                        (self.f_award, "award")):
            i = cb.findText(d.get(key, ""))
            if i >= 0:
                cb.setCurrentIndex(i)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = PamyatApp(); w.show()
    sys.exit(app.exec())
