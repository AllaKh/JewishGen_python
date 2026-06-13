"""
gui/gwar.py — «Памяти героев Великой войны» (gwar.mil.ru/heroes/) search window.

WWI (1914–1918) participant records. The site is Russian-only; the GUI is
English. The form mirrors the site's own sections (Basic data / Place /
Sections / Additional parameters / Document storage). Fills the search form,
collects FIO-matching records, opens each and copies the fields to Word,
saving the scanned document image(s).
"""

import asyncio, json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QComboBox,
    QFileDialog, QProgressBar, QMessageBox, QApplication, QGroupBox,
)
from PySide6.QtCore import QThread, Signal, Qt
from gui._app_icon import app_icon, make_header, make_cancel_button

_HERE   = Path(__file__).resolve().parent
_ROOT   = _HERE.parent
_SAVE   = _HERE / ".gwar_autosave.json"
_DEF_DIR = str(Path.home() / "Downloads" / "Gwar_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    import gwar_scraper as _scraper
    _SCRAPER_OK = True
    SECTIONS = list(_scraper.SECTIONS.keys())
    EVENTS = _scraper.EVENTS
except ImportError:
    _SCRAPER_OK = False
    SECTIONS = ["Awards", "Losses", "Personal data", "Commanders", "Notable people"]
    EVENTS = {"": ""}

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #c9bd9a;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#faf8f2;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#8a6d1f;background:#faf8f2;}
QLineEdit,QComboBox{padding:4px 6px;border:1px solid #cdc4ac;border-radius:4px;
  background:white;}
QLineEdit:focus,QComboBox:focus{border:1px solid #b8860b;}
QPushButton{padding:5px 14px;border-radius:4px;border:1px solid #c9bd9a;
  background:#f1ecdd;}
QPushButton:hover{background:#e6dcc2;}
QPushButton#startBtn{background:#8a6d1f;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#a07d28;}
QPushButton#startBtn:disabled{background:#cdbf94;}
QProgressBar{border:1px solid #c9bd9a;border-radius:4px;text-align:center;
  min-height:18px;}
QProgressBar::chunk{background:#8a6d1f;border-radius:3px;}
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


class GwarApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Heroes of the Great War")
        self.setMinimumWidth(820)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._build_ui()
        self._load()

    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(16, 12, 16, 12); outer.setSpacing(10)

        outer.addLayout(make_header("Voinalogo.png", "Heroes of the Great War",
                                    color="#8a6d1f"))
        outer.addWidget(QLabel("WWI (1914–1918) participant records "
                               "(gwar.mil.ru) — the site is Russian only."))

        # 1) Basic data ──────────────────────────────────────────────────────
        bg = QGroupBox("Basic data")
        bgl = QGridLayout(bg); bgl.setSpacing(6)
        self.f_last = QLineEdit(); self.f_first = QLineEdit(); self.f_mid = QLineEdit()
        self.f_birth = QLineEdit(); self.f_birth.setPlaceholderText("year or DD.MM.YYYY")
        bgl.addWidget(QLabel("Surname:"), 0, 0);    bgl.addWidget(self.f_last, 0, 1)
        bgl.addWidget(QLabel("First name:"), 0, 2); bgl.addWidget(self.f_first, 0, 3)
        bgl.addWidget(QLabel("Patronymic:"), 1, 0); bgl.addWidget(self.f_mid, 1, 1)
        bgl.addWidget(QLabel("Birth date:"), 1, 2); bgl.addWidget(self.f_birth, 1, 3)
        self.f_exact = QCheckBox("Exact match (surname / first name / patronymic)")
        self.f_exact.setChecked(True)
        bgl.addWidget(self.f_exact, 2, 0, 1, 4)
        outer.addWidget(bg)

        # 2) Place of residence / conscription ───────────────────────────────
        pg = QGroupBox("Place of residence / conscription")
        pgl = QGridLayout(pg); pgl.setSpacing(6)
        self.f_gub = QLineEdit(); self.f_uezd = QLineEdit()
        self.f_vol = QLineEdit(); self.f_set = QLineEdit()
        pgl.addWidget(QLabel("Gubernia:"), 0, 0);   pgl.addWidget(self.f_gub, 0, 1)
        pgl.addWidget(QLabel("Uezd:"), 0, 2);       pgl.addWidget(self.f_uezd, 0, 3)
        pgl.addWidget(QLabel("Volost:"), 1, 0);     pgl.addWidget(self.f_vol, 1, 1)
        pgl.addWidget(QLabel("Settlement:"), 1, 2); pgl.addWidget(self.f_set, 1, 3)
        outer.addWidget(pg)

        # 3) Sections (Разделы) — all on by default ──────────────────────────
        secg = QGroupBox("Sections")
        secl = QHBoxLayout(secg); secl.setSpacing(12)
        self._sec_cbs = {}
        for name in SECTIONS:
            cb = QCheckBox(name); cb.setChecked(True)
            self._sec_cbs[name] = cb
            secl.addWidget(cb)
        secl.addStretch()
        outer.addWidget(secg)

        # 4) Additional search parameters ────────────────────────────────────
        ag = QGroupBox("Additional search parameters")
        agl = QGridLayout(ag); agl.setSpacing(6)
        self.f_rank = QLineEdit(); self.f_unit = QLineEdit(); self.f_evplace = QLineEdit()
        self.f_event = QComboBox(); self.f_event.addItems(list(EVENTS.keys()))
        self.f_evfrom = QLineEdit(); self.f_evfrom.setPlaceholderText("DD.MM.YYYY")
        self.f_evto   = QLineEdit(); self.f_evto.setPlaceholderText("DD.MM.YYYY")
        agl.addWidget(QLabel("Rank / position:"), 0, 0); agl.addWidget(self.f_rank, 0, 1)
        agl.addWidget(QLabel("Military unit:"), 0, 2);   agl.addWidget(self.f_unit, 0, 3)
        agl.addWidget(QLabel("Event:"), 1, 0);           agl.addWidget(self.f_event, 1, 1, 1, 3)
        agl.addWidget(QLabel("Event start:"), 2, 0);     agl.addWidget(self.f_evfrom, 2, 1)
        agl.addWidget(QLabel("Event end:"), 2, 2);       agl.addWidget(self.f_evto, 2, 3)
        agl.addWidget(QLabel("Place of event:"), 3, 0);  agl.addWidget(self.f_evplace, 3, 1, 1, 3)
        outer.addWidget(ag)

        # 5) Document storage ────────────────────────────────────────────────
        dg = QGroupBox("Document storage")
        dgl = QGridLayout(dg); dgl.setSpacing(6)
        self.f_fund = QLineEdit(); self.f_inv = QLineEdit(); self.f_file = QLineEdit()
        dgl.addWidget(QLabel("Fund:"), 0, 0);              dgl.addWidget(self.f_fund, 0, 1)
        dgl.addWidget(QLabel("Inventory/Cabinet:"), 0, 2); dgl.addWidget(self.f_inv, 0, 3)
        dgl.addWidget(QLabel("File/Box:"), 1, 0);          dgl.addWidget(self.f_file, 1, 1)
        outer.addWidget(dg)

        # Output ─────────────────────────────────────────────────────────────
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
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        outer.addLayout(br)
        outer.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        for w in (self.f_last, self.f_first, self.f_mid, self.f_birth,
                  self.f_gub, self.f_uezd, self.f_vol, self.f_set,
                  self.f_rank, self.f_unit, self.f_evfrom, self.f_evto, self.f_evplace,
                  self.f_fund, self.f_inv, self.f_file, self.f_folder):
            w.textChanged.connect(self._save)
        self.f_event.currentTextChanged.connect(self._save)
        self.f_exact.stateChanged.connect(self._save)
        for cb in self._sec_cbs.values():
            cb.stateChanged.connect(self._save)

    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Output folder",
                                             self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    def _payload(self):
        return {
            "last_name":   self.f_last.text().strip(),
            "first_name":  self.f_first.text().strip(),
            "middle_name": self.f_mid.text().strip(),
            "birth_date":  self.f_birth.text().strip(),
            "gubernia":    self.f_gub.text().strip(),
            "uezd":        self.f_uezd.text().strip(),
            "volost":      self.f_vol.text().strip(),
            "settlement":  self.f_set.text().strip(),
            "rank":        self.f_rank.text().strip(),
            "unit":        self.f_unit.text().strip(),
            "event":       EVENTS.get(self.f_event.currentText(), ""),
            "event_from":  self.f_evfrom.text().strip(),
            "event_to":    self.f_evto.text().strip(),
            "event_place": self.f_evplace.text().strip(),
            "fund":        self.f_fund.text().strip(),
            "inventory":   self.f_inv.text().strip(),
            "file":        self.f_file.text().strip(),
            "sections":    {n: cb.isChecked() for n, cb in self._sec_cbs.items()},
            "exact":       self.f_exact.isChecked(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log":         print,
            "cancel_event": getattr(self, "_cancel_ev", None),
        }

    def _validate(self):
        p = self._payload()
        if not any(p[k] for k in ("last_name", "first_name", "rank", "unit",
                                  "settlement", "gubernia", "uezd")):
            QMessageBox.warning(self, "Nothing to search",
                                "Enter at least a surname (or another field).")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Error", "gwar_scraper.py not found.")
            return False
        return True

    def _start(self):
        if not self._validate():
            return
        self._cancel_ev = threading.Event()
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
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
                                 r.get("message", "Failed — see terminal."))
            self.stlbl.setText("Error.")

    def _save(self, *_):
        try:
            d = {"last": self.f_last.text(), "first": self.f_first.text(),
                 "mid": self.f_mid.text(), "birth": self.f_birth.text(),
                 "gub": self.f_gub.text(), "uezd": self.f_uezd.text(),
                 "vol": self.f_vol.text(), "set": self.f_set.text(),
                 "rank": self.f_rank.text(), "unit": self.f_unit.text(),
                 "event": self.f_event.currentText(),
                 "evfrom": self.f_evfrom.text(), "evto": self.f_evto.text(),
                 "evplace": self.f_evplace.text(), "fund": self.f_fund.text(),
                 "inv": self.f_inv.text(), "file": self.f_file.text(),
                 "folder": self.f_folder.text(), "exact": self.f_exact.isChecked(),
                 "sections": {n: cb.isChecked() for n, cb in self._sec_cbs.items()}}
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
        self.f_mid.setText(d.get("mid", "")); self.f_birth.setText(d.get("birth", ""))
        self.f_gub.setText(d.get("gub", "")); self.f_uezd.setText(d.get("uezd", ""))
        self.f_vol.setText(d.get("vol", "")); self.f_set.setText(d.get("set", ""))
        self.f_rank.setText(d.get("rank", "")); self.f_unit.setText(d.get("unit", ""))
        self.f_evfrom.setText(d.get("evfrom", "")); self.f_evto.setText(d.get("evto", ""))
        self.f_evplace.setText(d.get("evplace", "")); self.f_fund.setText(d.get("fund", ""))
        self.f_inv.setText(d.get("inv", "")); self.f_file.setText(d.get("file", ""))
        self.f_folder.setText(d.get("folder", _DEF_DIR))
        i = self.f_event.findText(d.get("event", ""))
        if i >= 0:
            self.f_event.setCurrentIndex(i)
        self.f_exact.setChecked(bool(d.get("exact", True)))
        secs = d.get("sections", {})
        for n, cb in self._sec_cbs.items():
            cb.setChecked(bool(secs.get(n, True)))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = GwarApp(); w.show()
    sys.exit(app.exec())
