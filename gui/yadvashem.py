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
from gui._app_icon import app_icon, make_header, make_cancel_button

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
QPushButton#advBtn{text-align:left;font-weight:bold;color:#1f4e79;
  background:#e8eef7;border:1px solid #b8c2d0;padding:7px 12px;}
QPushButton#advBtn:hover{background:#dbe5f4;}
QPushButton#advBtn:checked{background:#d2e0f2;}
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

    # uniform grid: fixed-width label | stretching input | fixed «type» combo
    def _cfg_grid(self, grid):
        grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 175)
        grid.setColumnStretch(1, 1)

    def _row(self, grid, r, label, key, with_type=True, items=None):
        grid.addWidget(QLabel(label), r, 0)
        ed = QLineEdit()
        grid.addWidget(ed, r, 1)
        combo = None
        if with_type:
            combo = QComboBox(); combo.addItems(items or SEARCH_TYPES)
            combo.setFixedWidth(130)
            grid.addWidget(combo, r, 2)
        self._fields[key] = (ed, combo)
        return ed, combo

    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        self._outer = QVBoxLayout(central)
        self._outer.setContentsMargins(16, 10, 16, 10); self._outer.setSpacing(9)

        self._outer.addLayout(make_header(
            "Yadvashemlogo.png", "Yad Vashem — Shoah Victims' Names", color="#1f4e79"))
        self._outer.addWidget(QLabel(
            "collections.yadvashem.org — open database, no login."))

        # ── Search language ──────────────────────────────────────────────────
        lg = QGroupBox("Search language"); ll = QHBoxLayout(lg); ll.setSpacing(10)
        self.f_lang = QComboBox(); self.f_lang.addItems(list(LANGS.keys()))
        _i = self.f_lang.findText("Russian")
        if _i >= 0:
            self.f_lang.setCurrentIndex(_i)
        ll.addWidget(QLabel("Type the names in / search in this language:"))
        ll.addWidget(self.f_lang); ll.addStretch()
        self._outer.addWidget(lg)

        # ── Name (always visible) ────────────────────────────────────────────
        ng = QGroupBox("Name"); ngl = QGridLayout(ng); self._cfg_grid(ngl)
        self._row(ngl, 0, "Last name:",   "last_name")
        self._row(ngl, 1, "First name:",  "first_name")
        self._row(ngl, 2, "Maiden name:", "maiden_name")
        self._outer.addWidget(ng)

        # ── Date (always visible) ────────────────────────────────────────────
        dg = QGroupBox("Date"); dgl = QGridLayout(dg); self._cfg_grid(dgl)
        self._row(dgl, 0, "Birth year:", "birth_year", items=YEAR_PREC)
        self._row(dgl, 1, "Death year:", "death_year", items=YEAR_PREC)
        self._outer.addWidget(dg)

        # ── Advanced toggle ──────────────────────────────────────────────────
        self._adv_btn = QPushButton(
            "▶   Advanced search  (Place · Family members · Submitter · Global)")
        self._adv_btn.setObjectName("advBtn"); self._adv_btn.setCheckable(True)
        self._adv_btn.toggled.connect(self._toggle_adv)
        self._outer.addWidget(self._adv_btn)

        # ── Advanced panel (collapsible, scrollable) ─────────────────────────
        self._adv = QWidget(); av = QVBoxLayout(self._adv)
        av.setContentsMargins(0, 0, 0, 0); av.setSpacing(9)

        pg = QGroupBox("Place  (place names may be given in their original form)")
        pgl = QGridLayout(pg); self._cfg_grid(pgl)
        self.rb_byfield = QRadioButton("Search by specific field"); self.rb_byfield.setChecked(True)
        self.rb_anyplace = QRadioButton("Search any place")
        bgp = QButtonGroup(self); bgp.addButton(self.rb_byfield); bgp.addButton(self.rb_anyplace)
        rrow = QHBoxLayout(); rrow.addWidget(self.rb_byfield)
        rrow.addSpacing(20); rrow.addWidget(self.rb_anyplace); rrow.addStretch()
        pgl.addLayout(rrow, 0, 0, 1, 3)
        self._row(pgl, 1, "Birth place:",            "birth_place")
        self._row(pgl, 2, "Before the war:",         "place_before")
        self._row(pgl, 3, "During the war / Shoah:", "place_during")
        self._row(pgl, 4, "Death place:",            "death_place")
        av.addWidget(pg)

        fg = QGroupBox("Family members"); fgl = QGridLayout(fg); self._cfg_grid(fgl)
        self._row(fgl, 0, "Father's name:",        "father_name")
        self._row(fgl, 1, "Mother's name:",        "mother_name")
        self._row(fgl, 2, "Mother's maiden name:", "mother_maiden")
        self._row(fgl, 3, "Spouse's name:",        "spouse_name")
        self._row(fgl, 4, "Spouse's maiden name:", "spouse_maiden")
        av.addWidget(fg)

        sg = QGroupBox("Submitter  (of Pages of Testimony, survivor forms, memorial materials)")
        sgl = QGridLayout(sg); self._cfg_grid(sgl)
        self._row(sgl, 0, "First name:", "submitter_first")
        self._row(sgl, 1, "Last name:",  "submitter_last")
        av.addWidget(sg)

        gg = QGroupBox("Global search"); ggl = QGridLayout(gg); self._cfg_grid(ggl)
        ggl.addWidget(QLabel("Any text:"), 0, 0)
        self.f_global = QLineEdit(); ggl.addWidget(self.f_global, 0, 1, 1, 2)
        av.addWidget(gg)

        self._adv_scroll = QScrollArea(); self._adv_scroll.setWidgetResizable(True)
        self._adv_scroll.setFrameShape(QFrame.NoFrame)
        self._adv_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._adv_scroll.setWidget(self._adv)
        self._adv_scroll.setVisible(False)
        self._outer.addWidget(self._adv_scroll, 1)

        # ── Output ───────────────────────────────────────────────────────────
        og = QGroupBox("Output (Word)"); ol = QHBoxLayout(og); ol.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        ol.addWidget(QLabel("Save to:")); ol.addWidget(self.f_folder, 1); ol.addWidget(bb)
        self._outer.addWidget(og)

        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        self._outer.addWidget(self.pbar); self._outer.addWidget(self.stlbl)
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH"); self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        for ed, combo in self._fields.values():
            ed.textChanged.connect(self._save)
            if combo is not None:
                combo.currentTextChanged.connect(self._save)
        self.f_global.textChanged.connect(self._save)
        self.f_folder.textChanged.connect(self._save)
        self.rb_byfield.toggled.connect(self._save)
        self.f_lang.currentTextChanged.connect(self._save)

        self.setFixedWidth(960)
        self._fit()

    # ── Advanced toggle + height fit (never taller than the screen) ──────────
    def _toggle_adv(self, on):
        self._adv_scroll.setVisible(on)
        self._adv_btn.setText(("▼" if on else "▶")
                              + "   Advanced search  (Place · Family members · Submitter · Global)")
        self._fit()

    def _fit(self):
        self.setMinimumHeight(0); self.setMaximumHeight(16777215)
        avail = QApplication.primaryScreen().availableGeometry().height()
        if self._adv_scroll.isVisible():
            self._adv_scroll.setMaximumHeight(max(160, int(avail * 0.5)))
        self._outer.invalidate(); self._outer.activate()
        h = min(self.sizeHint().height(), avail - 48)   # stay on-screen
        self.resize(self.width(), h)
        self.setFixedHeight(h)

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
            "cancel_event": getattr(self, "_cancel_ev", None),
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
