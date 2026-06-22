"""
gui/wikisource.py — Вікіджерела (Ukrainian Wikisource) search window.

Archive scans of Jewish records on uk.wikisource.org. Two modes:
  • By place — walk «Архів:Єврейське містечко/<губернія>» and grab the documents
    of a повіт / місто / містечко, filtered by document type.
  • By archive code — e.g. «ДАКрО/185/1/49».
The scraper downloads the full-resolution PDFs from Wikimedia Commons.
"""

import asyncio, json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, QRadioButton,
    QButtonGroup, QSpinBox, QFrame, QGroupBox,
    QFileDialog, QProgressBar, QMessageBox, QApplication,
)
from PySide6.QtCore import QThread, Signal, Qt
from gui._app_icon import app_icon, app_version, make_footer, make_header, make_cancel_button, autosave_path

_HERE   = Path(__file__).resolve().parent
_ROOT   = _HERE.parent
_SAVE   = autosave_path(".wikisource_autosave.json")
_DEF_DIR = str(Path.home() / "Downloads" / "Wikisource_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    import wikisource_scraper as _scraper
    _SCRAPER_OK = True
    REGIONS = dict(_scraper.REGIONS)            # uk → (english, russian)
    DOC_TYPE_TR = dict(_scraper.DOC_TYPE_TR)    # uk → (english, russian)
except Exception:
    _SCRAPER_OK = False
    REGIONS = {"Волинська губернія": ("Volhynia Governorate", "Волынская губерния")}
    DOC_TYPE_TR = {"Ревізькі казки": ("Revision lists", "Ревизские сказки")}

_ALL_GUB = "— all regions / все регионы —"

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b6cfc7;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f5faf8;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#1f6f5c;background:#f5faf8;}
QLineEdit,QComboBox,QSpinBox{padding:3px 6px;border:1px solid #bcd0c9;
  border-radius:4px;background:white;}
QLineEdit:focus,QComboBox:focus{border:1px solid #1f9f80;}
QPushButton{padding:5px 14px;border-radius:4px;border:1px solid #b6cfc7;
  background:#eaf4f0;}
QPushButton:hover{background:#dbeee7;}
QPushButton#startBtn{background:#1f6f5c;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#26856e;}
QPushButton#startBtn:disabled{background:#9cbcb2;}
QProgressBar{border:1px solid #bcd0c9;border-radius:4px;text-align:center;
  min-height:18px;}
QProgressBar::chunk{background:#1f6f5c;border-radius:3px;}
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


class WikisourceApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ukrainian Wikisource Archives")
        self.setMinimumWidth(820)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._type_cb = {}
        self._build_ui()
        self._load()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        self._outer = QVBoxLayout(central)
        self._outer.setContentsMargins(16, 10, 16, 10); self._outer.setSpacing(9)

        self._outer.addLayout(make_header(
            "Vikijerelalogo.png", "Ukrainian Wikisource Archives", color="#1f6f5c"))
        self._outer.addWidget(QLabel(
            "uk.wikisource.org — archival scans (revision lists, vital records…). "
            "Downloads full-resolution PDFs from Wikimedia Commons."))

        # ── Search mode ──────────────────────────────────────────────────
        mg = QGroupBox("Search mode"); ml = QHBoxLayout(mg); ml.setSpacing(20)
        self.rb_place = QRadioButton("By place (Jewish shtetl tree)")
        self.rb_code  = QRadioButton("By archive code")
        self.rb_place.setChecked(True)
        bg = QButtonGroup(self); bg.addButton(self.rb_place); bg.addButton(self.rb_code)
        ml.addWidget(self.rb_place); ml.addWidget(self.rb_code); ml.addStretch()
        self._outer.addWidget(mg)

        # ── By place ─────────────────────────────────────────────────────
        self.pg = QGroupBox("By place"); pgl = QGridLayout(self.pg)
        pgl.setSpacing(6); pgl.setColumnMinimumWidth(0, 150); pgl.setColumnStretch(1, 1)
        pgl.addWidget(QLabel("Region:"), 0, 0)
        self.f_gub = QComboBox(); self.f_gub.addItem(_ALL_GUB, "")
        for uk, (en, ru) in REGIONS.items():
            self.f_gub.addItem(f"{en}   /   {ru}", uk)
        pgl.addWidget(self.f_gub, 0, 1)
        pgl.addWidget(QLabel("Place (county / town):"), 1, 0)
        self.f_place = QLineEdit()
        self.f_place.setPlaceholderText("e.g. Zhytomyr / Сураж  (English or Russian — matched to Ukrainian)")
        pgl.addWidget(self.f_place, 1, 1)
        pgl.addWidget(QLabel("Document types:"), 2, 0, Qt.AlignTop)
        tw = QWidget(); tv = QVBoxLayout(tw); tv.setContentsMargins(0, 0, 0, 0); tv.setSpacing(3)
        for uk, (en, ru) in DOC_TYPE_TR.items():
            cb = QCheckBox(f"{en}   /   {ru}"); cb.setChecked(True)
            self._type_cb[uk] = cb; tv.addWidget(cb)
        self.cb_all = QCheckBox("All documents of the place  (ignore type filter)")
        tv.addWidget(self.cb_all)
        pgl.addWidget(tw, 2, 1)
        self._outer.addWidget(self.pg)

        # ── By archive code ──────────────────────────────────────────────
        self.cg = QGroupBox("By archive code"); cgl = QGridLayout(self.cg)
        cgl.setSpacing(6); cgl.setColumnMinimumWidth(0, 150); cgl.setColumnStretch(1, 1)
        cgl.addWidget(QLabel("Archive code:"), 0, 0)
        self.f_code = QLineEdit(); self.f_code.setPlaceholderText("e.g. ДАКрО/185/1/49")
        cgl.addWidget(self.f_code, 0, 1)
        self._outer.addWidget(self.cg)

        # ── Options ──────────────────────────────────────────────────────
        og = QGroupBox("Options"); ol = QHBoxLayout(og); ol.setSpacing(14)
        self.cb_preview = QCheckBox("Cover preview only (JPG, fast)")
        ol.addWidget(self.cb_preview)
        ol.addWidget(QLabel("Max documents:"))
        self.sp_max = QSpinBox(); self.sp_max.setRange(1, 5000); self.sp_max.setValue(400)
        self.sp_max.setFixedWidth(80); ol.addWidget(self.sp_max)
        ol.addStretch()
        self._outer.addWidget(og)

        # ── Output ───────────────────────────────────────────────────────
        sg = QGroupBox("Output folder"); sl = QHBoxLayout(sg); sl.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        sl.addWidget(QLabel("Save to:")); sl.addWidget(self.f_folder, 1); sl.addWidget(bb)
        self._outer.addWidget(sg)

        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        self._outer.addWidget(self.pbar); self._outer.addWidget(self.stlbl)
        br = QHBoxLayout()
        self.start_btn = QPushButton("START"); self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(make_footer())

        self.rb_place.toggled.connect(self._update_mode)
        for w in (self.f_gub, self.f_place, self.f_code, self.f_folder):
            (w.currentTextChanged if isinstance(w, QComboBox) else w.textChanged).connect(self._save)
        for cb in list(self._type_cb.values()) + [self.cb_all, self.cb_preview]:
            cb.toggled.connect(self._save)
        self.sp_max.valueChanged.connect(self._save)
        self.rb_place.toggled.connect(self._save)

        self._update_mode()
        self.setFixedWidth(840)
        self._fit()

    def _update_mode(self, *_):
        place = self.rb_place.isChecked()
        self.pg.setEnabled(place)
        self.cg.setEnabled(not place)

    def _fit(self):
        self.setMinimumHeight(0); self.setMaximumHeight(16777215)
        self._outer.invalidate(); self._outer.activate()
        avail = QApplication.primaryScreen().availableGeometry().height()
        h = min(self.sizeHint().height(), avail - 48)
        self.resize(self.width(), h); self.setFixedHeight(h)

    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Output folder",
                                             self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    # ------------------------------------------------------------------ payload
    def _payload(self):
        return {
            "mode":          "locality" if self.rb_place.isChecked() else "code",
            "query":         self.f_code.text().strip(),
            "gubernia":      self.f_gub.currentData() or "",
            "locality":      self.f_place.text().strip(),
            "types":         [k for k, cb in self._type_cb.items() if cb.isChecked()],
            "all_docs":      self.cb_all.isChecked(),
            "jewish_only":   True,
            "preview_only":  self.cb_preview.isChecked(),
            "max_docs":      self.sp_max.value(),
            "output_folder": self.f_folder.text().strip() or _DEF_DIR,
            "log":           print,
            "cancel_event": getattr(self, "_cancel_ev", None),
        }

    def _validate(self):
        p = self._payload()
        if p["mode"] == "code" and not p["query"]:
            QMessageBox.warning(self, "Empty code", "Enter an archive code.")
            return False
        if p["mode"] == "locality" and not p["gubernia"] and not p["locality"]:
            QMessageBox.warning(self, "Specify a place",
                                "Choose a region or enter a place.")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Error", "wikisource_scraper.py not found.")
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
            n = r.get("n_records", 0); tot = r.get("n_total", n)
            msg = f"Downloaded {n} of {tot} documents."
            if r.get("output_folder"):
                msg += f"\n\nFolder:\n{r['output_folder']}"
            QMessageBox.information(self, "Done", msg)
            self.stlbl.setText("Done.")
        else:
            QMessageBox.critical(self, "Error", r.get("message", "Failed — see terminal."))
            self.stlbl.setText("Error.")

    # ------------------------------------------------------------------ autosave
    def _save(self, *_):
        try:
            d = {"mode": "locality" if self.rb_place.isChecked() else "code",
                 "gub": self.f_gub.currentData() or "", "place": self.f_place.text(),
                 "code": self.f_code.text(), "folder": self.f_folder.text(),
                 "all": self.cb_all.isChecked(), "preview": self.cb_preview.isChecked(),
                 "max": self.sp_max.value(),
                 "types": [k for k, cb in self._type_cb.items() if cb.isChecked()]}
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
        (self.rb_code if d.get("mode") == "code" else self.rb_place).setChecked(True)
        i = self.f_gub.findData(d.get("gub", ""))
        if i >= 0:
            self.f_gub.setCurrentIndex(i)
        self.f_place.setText(d.get("place", ""))
        self.f_code.setText(d.get("code", ""))
        if d.get("folder"):
            self.f_folder.setText(d["folder"])
        self.cb_all.setChecked(bool(d.get("all")))
        self.cb_preview.setChecked(bool(d.get("preview")))
        self.sp_max.setValue(int(d.get("max", 400)))
        if "types" in d:
            for k, cb in self._type_cb.items():
                cb.setChecked(k in d["types"])
        self._update_mode()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = WikisourceApp(); w.show()
    sys.exit(app.exec())
