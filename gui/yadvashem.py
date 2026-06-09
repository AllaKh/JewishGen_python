"""
gui/yadvashem.py — Yad Vashem «Central Database of Shoah Victims' Names»
(collections.yadvashem.org) search window.

Open site, no login. The GUI mirrors the site's advanced search form (Name /
Date / Place / Family members / Submitter / Global) — all English. The search is
done by URL parameters; the scraper builds the search URL, collects the matching
victim records and writes them to Word.
"""

import asyncio, json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, QRadioButton,
    QButtonGroup, QScrollArea, QFrame,
    QFileDialog, QProgressBar, QMessageBox, QApplication, QGroupBox,
)
from PySide6.QtCore import QThread, Signal, Qt
from gui._app_icon import app_icon, make_header

_HERE   = Path(__file__).resolve().parent
_ROOT   = _HERE.parent
_SAVE   = _HERE / ".yadvashem_autosave.json"
_DEF_DIR = str(Path.home() / "Downloads" / "YadVashem_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    import yadvashem_scraper as _scraper
    _SCRAPER_OK = True
    SEARCH_TYPES = list(_scraper.SEARCH_TYPES.keys())
    YEAR_PREC = list(_scraper.YEAR_PREC.keys())
    LANGS = _scraper.LANGS
except ImportError:
    _SCRAPER_OK = False
    SEARCH_TYPES = ["YV synonyms", "Literal", "Phonetic"]
    YEAR_PREC = ["Exact", "± 2", "± 5"]
    LANGS = {"English": "en", "Hebrew": "he", "Russian": "ru",
             "Spanish": "es", "German": "de", "French": "fr"}

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b8c2d0;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f7f9fc;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#1f4e79;background:#f7f9fc;}
QLineEdit,QComboBox{padding:3px 6px;border:1px solid #c3ccda;border-radius:4px;
  background:white;}
QLineEdit:focus,QComboBox:focus{border:1px solid #1f6fc4;}
QPushButton{padding:5px 14px;border-radius:4px;border:1px solid #b8c2d0;
  background:#eef2f8;}
QPushButton:hover{background:#dde6f2;}
QPushButton#startBtn{background:#1f4e79;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#2a5e90;}
QPushButton#startBtn:disabled{background:#9fb3c8;}
QProgressBar{border:1px solid #c3ccda;border-radius:4px;text-align:center;
  min-height:18px;}
QProgressBar::chunk{background:#1f4e79;border-radius:3px;}
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


class YadVashemApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Yad Vashem — Shoah Victims' Names")
        self.setMinimumWidth(900)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._fields = {}           # key -> (QLineEdit, QComboBox|None)
        self._build_ui()
        self._load()

    # helper: a labelled text field + optional «search type» combo, into a grid
    def _row(self, grid, r, c, label, key, with_type=True, items=None):
        grid.addWidget(QLabel(label), r, c)
        ed = QLineEdit()
        grid.addWidget(ed, r, c + 1)
        combo = None
        if with_type:
            combo = QComboBox(); combo.addItems(items or SEARCH_TYPES)
            combo.setFixedWidth(120)
            grid.addWidget(combo, r, c + 2)
        self._fields[key] = (ed, combo)
        return ed, combo

    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        outerv = QVBoxLayout(central); outerv.setContentsMargins(0, 0, 0, 0)

        head = QWidget(); hv = QVBoxLayout(head)
        hv.setContentsMargins(16, 10, 16, 0)
        hv.addLayout(make_header("Yadvashemlogo.png",
                                 "Yad Vashem — Shoah Victims' Names", color="#1f4e79"))
        hv.addWidget(QLabel("collections.yadvashem.org — open database, no login. "
                            "The form mirrors the site's advanced search."))
        outerv.addWidget(head)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        body = QWidget(); outer = QVBoxLayout(body)
        outer.setContentsMargins(16, 4, 16, 8); outer.setSpacing(10)
        scroll.setWidget(body)
        outerv.addWidget(scroll, 1)

        # ── Search language ──────────────────────────────────────────────────
        lg = QGroupBox("Search language"); ll = QHBoxLayout(lg); ll.setSpacing(10)
        self.f_lang = QComboBox(); self.f_lang.addItems(list(LANGS.keys()))
        _i = self.f_lang.findText("Russian")
        if _i >= 0:
            self.f_lang.setCurrentIndex(_i)
        ll.addWidget(QLabel("Type the names in / search in this language:"))
        ll.addWidget(self.f_lang); ll.addStretch()
        outer.addWidget(lg)

        # ── Name ─────────────────────────────────────────────────────────────
        ng = QGroupBox("Name"); ngl = QGridLayout(ng); ngl.setSpacing(6)
        self._row(ngl, 0, 0, "Last name:",   "last_name")
        self._row(ngl, 1, 0, "First name:",  "first_name")
        self._row(ngl, 2, 0, "Maiden name:", "maiden_name")
        outer.addWidget(ng)

        # ── Date ─────────────────────────────────────────────────────────────
        dg = QGroupBox("Date"); dgl = QGridLayout(dg); dgl.setSpacing(6)
        self._row(dgl, 0, 0, "Birth year:", "birth_year", with_type=True, items=YEAR_PREC)
        self._row(dgl, 1, 0, "Death year:", "death_year", with_type=True, items=YEAR_PREC)
        outer.addWidget(dg)

        # ── Place ────────────────────────────────────────────────────────────
        pg = QGroupBox("Place"); pgl = QGridLayout(pg); pgl.setSpacing(6)
        self.rb_byfield = QRadioButton("Search by specific field"); self.rb_byfield.setChecked(True)
        self.rb_anyplace = QRadioButton("Search any place (not a specific field)")
        bgp = QButtonGroup(self); bgp.addButton(self.rb_byfield); bgp.addButton(self.rb_anyplace)
        pgl.addWidget(self.rb_byfield, 0, 0, 1, 3); pgl.addWidget(self.rb_anyplace, 0, 3, 1, 3)
        self._row(pgl, 1, 0, "Birth place:",          "birth_place")
        self._row(pgl, 2, 0, "Before the war:",       "place_before")
        self._row(pgl, 3, 0, "During the war / Shoah:", "place_during")
        self._row(pgl, 4, 0, "Death place:",          "death_place")
        outer.addWidget(pg)

        # ── Family members ───────────────────────────────────────────────────
        fg = QGroupBox("Family members"); fgl = QGridLayout(fg); fgl.setSpacing(6)
        self._row(fgl, 0, 0, "Father's name:",        "father_name")
        self._row(fgl, 1, 0, "Mother's name:",        "mother_name")
        self._row(fgl, 2, 0, "Mother's maiden name:", "mother_maiden")
        self._row(fgl, 3, 0, "Spouse's name:",        "spouse_name")
        self._row(fgl, 4, 0, "Spouse's maiden name:", "spouse_maiden")
        outer.addWidget(fg)

        # ── Submitter ────────────────────────────────────────────────────────
        sg = QGroupBox("Submitter (of Pages of Testimony)")
        sgl = QGridLayout(sg); sgl.setSpacing(6)
        self._row(sgl, 0, 0, "First name:", "submitter_first")
        self._row(sgl, 1, 0, "Last name:",  "submitter_last")
        outer.addWidget(sg)

        # ── Global ───────────────────────────────────────────────────────────
        gg = QGroupBox("Global search (applies only to this field)")
        ggl = QHBoxLayout(gg); ggl.setSpacing(6)
        self.f_global = QLineEdit()
        ggl.addWidget(QLabel("Any text:")); ggl.addWidget(self.f_global, 1)
        outer.addWidget(gg)

        # ── Output ───────────────────────────────────────────────────────────
        og = QGroupBox("Output (Word)"); ol = QHBoxLayout(og); ol.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        ol.addWidget(QLabel("Save to:")); ol.addWidget(self.f_folder, 1); ol.addWidget(bb)
        outer.addWidget(og)

        # ── Footer (outside the scroll area) ─────────────────────────────────
        foot = QWidget(); fv = QVBoxLayout(foot); fv.setContentsMargins(16, 0, 16, 10)
        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        fv.addWidget(self.pbar); fv.addWidget(self.stlbl)
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH"); self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn); br.addStretch()
        fv.addLayout(br)
        fv.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))
        outerv.addWidget(foot)

        for ed, combo in self._fields.values():
            ed.textChanged.connect(self._save)
            if combo is not None:
                combo.currentTextChanged.connect(self._save)
        self.f_global.textChanged.connect(self._save)
        self.f_folder.textChanged.connect(self._save)
        self.rb_byfield.toggled.connect(self._save)
        self.f_lang.currentTextChanged.connect(self._save)

        self.resize(940, 880)

    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Output folder",
                                             self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    def _payload(self):
        flds = {}
        for key, (ed, combo) in self._fields.items():
            flds[key] = ed.text().strip()
            if combo is not None:
                flds[key + "_type"] = combo.currentText()
        return {
            "fields":      flds,
            "place_mode":  "byfield" if self.rb_byfield.isChecked() else "anyplace",
            "global_text": self.f_global.text().strip(),
            "lang":        LANGS.get(self.f_lang.currentText(), "ru"),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log":         print,
            "cancel_event": None,
        }

    def _validate(self):
        p = self._payload()
        if not any(v for v in p["fields"].values() if v) and not p["global_text"]:
            QMessageBox.warning(self, "Nothing to search",
                                "Enter at least one field (e.g. a surname).")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Error", "yadvashem_scraper.py not found.")
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
            d = {"place_mode": "byfield" if self.rb_byfield.isChecked() else "anyplace",
                 "global": self.f_global.text(), "folder": self.f_folder.text(),
                 "lang": self.f_lang.currentText(), "fields": {}}
            for key, (ed, combo) in self._fields.items():
                d["fields"][key] = ed.text()
                if combo is not None:
                    d["fields"][key + "_type"] = combo.currentText()
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
        fl = d.get("fields", {})
        for key, (ed, combo) in self._fields.items():
            ed.setText(fl.get(key, ""))
            if combo is not None:
                i = combo.findText(fl.get(key + "_type", ""))
                if i >= 0:
                    combo.setCurrentIndex(i)
        self.f_global.setText(d.get("global", ""))
        self.f_folder.setText(d.get("folder", _DEF_DIR))
        _i = self.f_lang.findText(d.get("lang", "Russian"))
        if _i >= 0:
            self.f_lang.setCurrentIndex(_i)
        if d.get("place_mode") == "anyplace":
            self.rb_anyplace.setChecked(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = YadVashemApp(); w.show()
    sys.exit(app.exec())
