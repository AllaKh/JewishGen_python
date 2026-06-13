"""
gui/ancestry.py
================
Ancestry.com search window. English GUI, green Ancestry theme.

Basic fields (First Names, Last Names, Place, Birth Year) each have their own
"Exact" checkbox (Ancestry's exact-match), all passed to the scraper. Advanced
section adds family members (Father / Mother / Spouse, each with First/Last +
Exact) and a Keyword. Login is required (persistent profile keeps the session).
Logo: Ancestry.png. Autosave: .ancestry_autosave.json
"""

import json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox, QFrame, QGridLayout, QScrollArea,
)
from gui._app_icon import app_icon, make_header
from PySide6.QtCore import QThread, Signal, Qt, QByteArray
from PySide6.QtGui import QIcon

_EYE_OPEN = b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
  fill="none" stroke="#555" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
  <circle cx="12" cy="12" r="3"/></svg>"""
_EYE_SHUT = b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
  fill="none" stroke="#555" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8
           a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4
           c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/>
  <line x1="1" y1="1" x2="23" y2="23"/></svg>"""


def _svg_icon(svg: bytes, size: int = 20) -> QIcon:
    from PySide6.QtSvg import QSvgRenderer
    from PySide6.QtGui import QPixmap, QPainter
    r = QSvgRenderer(QByteArray(svg))
    pix = QPixmap(size, size); pix.fill(Qt.transparent)
    p = QPainter(pix); r.render(p); p.end()
    return QIcon(pix)


# ── Paths ─────────────────────────────────────────────────────────────────── #
_HERE    = Path(__file__).resolve().parent
_ROOT    = _HERE.parent
_SAVE    = _HERE / ".ancestry_autosave.json"
_DEF_DIR = str(Path.home() / "Downloads" / "Ancestry_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
try:
    import ancestry_scraper as _scraper
    _SCRAPER_OK = True
except ImportError:
    _SCRAPER_OK = False

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b7cfa0;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f6faf0;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#4a6b1f;background:#f6faf0;}
QLineEdit,QComboBox{padding:4px 6px;border:1px solid #b7cfa0;
  border-radius:4px;background:white;min-height:22px;}
QLineEdit:focus,QComboBox:focus{border:1px solid #6B8E23;}
QPushButton{padding:5px 14px;border-radius:4px;
  border:1px solid #b7cfa0;background:#eef5e2;}
QPushButton:hover{background:#dcebc8;}
QPushButton#startBtn{background:#6B8E23;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#7da82b;}
QPushButton#startBtn:disabled{background:#aac178;}
QPushButton#eyeBtn{border:none;background:transparent;padding:0;}
QPushButton#advBtn{text-align:left;border:none;background:transparent;
  color:#4a6b1f;font-weight:bold;font-size:11px;padding:2px 0;}
QPushButton#advBtn:hover{color:#6B8E23;}
QProgressBar{border:1px solid #b7cfa0;border-radius:4px;
  text-align:center;min-height:18px;}
QProgressBar::chunk{background:#6B8E23;border-radius:3px;}
QLabel#sechead{font-weight:bold;color:#4a6b1f;font-size:10px;margin-top:4px;}
QFrame#div{background:#cfe0b5;}
"""


class PwdEdit(QLineEdit):
    def __init__(self):
        super().__init__()
        self.setEchoMode(QLineEdit.Password)
        self._io = _svg_icon(_EYE_OPEN, 20)
        self._is = _svg_icon(_EYE_SHUT, 20)
        self._btn = QPushButton(self)
        self._btn.setObjectName("eyeBtn")
        self._btn.setIcon(self._is)
        self._btn.setFixedSize(28, 28)
        self._btn.setCursor(Qt.PointingHandCursor)
        self._btn.setCheckable(True)
        self._btn.toggled.connect(lambda v: (
            self.setEchoMode(QLineEdit.Normal if v else QLineEdit.Password),
            self._btn.setIcon(self._io if v else self._is),
        ))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._btn.move(self.width()-30, (self.height()-28)//2)
        self.setTextMargins(0, 0, 32, 0)


class Worker(QThread):
    progress     = Signal(int, str)
    finished     = Signal(dict)
    request_file = Signal(str)

    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload
        self._file_choice = "overwrite"
        self._file_ev     = threading.Event()

    def provide_file_choice(self, choice: str):
        self._file_choice = choice
        self._file_ev.set()

    def run(self):
        import asyncio
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
            result = {"ok": False, "error": "exception",
                      "message": f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)


def _divider() -> QFrame:
    f = QFrame(); f.setObjectName("div")
    f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
    return f


def _sechead(text: str) -> QLabel:
    l = QLabel(text); l.setObjectName("sechead"); return l


class AncestryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ancestry")
        self.setMinimumWidth(900)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._build_ui()
        self._load()

    # ── Build UI ──────────────────────────────────────────────────────────── #
    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(18, 12, 18, 12)
        self._outer.setSpacing(8)

        self._outer.addLayout(
            make_header("Ancestry.png", "Ancestry", color="#4a6b1f"))

        # Credentials
        cg = QGroupBox("Account credentials")
        cl = QHBoxLayout(cg); cl.setSpacing(8)
        self.f_user = QLineEdit()
        self.f_user.setPlaceholderText("Ancestry email or username")
        self.f_pass = PwdEdit()
        self.f_pass.setPlaceholderText("Password")
        cl.addWidget(QLabel("Username:")); cl.addWidget(self.f_user, 2)
        cl.addWidget(QLabel("Password:")); cl.addWidget(self.f_pass, 2)
        self._outer.addWidget(cg)

        # Basic Search — each field has its own Exact checkbox
        bg = QGroupBox("Basic Search")
        bf = QGridLayout(bg); bf.setSpacing(8)
        self.f_first = QLineEdit(); self.f_first.setPlaceholderText("First & Middle Name(s)")
        self.f_last  = QLineEdit(); self.f_last.setPlaceholderText("Last Name")
        self.f_place = QLineEdit(); self.f_place.setPlaceholderText("City, county, state, country")
        self.f_byear = QLineEdit(); self.f_byear.setPlaceholderText("e.g. 1897")
        self.f_byear.setFixedWidth(120)
        self.f_first_exact = QCheckBox("Exact")
        self.f_last_exact  = QCheckBox("Exact")
        self.f_place_exact = QCheckBox("Exact")
        self.f_byear_exact = QCheckBox("Exact")
        bf.addWidget(QLabel("First Names:"), 0, 0)
        bf.addWidget(self.f_first,           0, 1)
        bf.addWidget(self.f_first_exact,     0, 2)
        bf.addWidget(QLabel("Last Names:"),  0, 3)
        bf.addWidget(self.f_last,            0, 4)
        bf.addWidget(self.f_last_exact,      0, 5)
        bf.addWidget(QLabel("Place:"),       1, 0)
        bf.addWidget(self.f_place,           1, 1)
        bf.addWidget(self.f_place_exact,     1, 2)
        bf.addWidget(QLabel("Birth Year:"),  1, 3)
        bf.addWidget(self.f_byear,           1, 4)
        bf.addWidget(self.f_byear_exact,     1, 5)
        bf.setColumnStretch(1, 2); bf.setColumnStretch(4, 1)
        self._outer.addWidget(bg)

        # Advanced toggle
        self._adv_btn = QPushButton("▶   Advanced Search")
        self._adv_btn.setObjectName("advBtn")
        self._adv_btn.setCheckable(True)
        self._adv_btn.toggled.connect(self._toggle_adv)
        self._outer.addWidget(self._adv_btn)

        self._adv = QGroupBox()
        av = QVBoxLayout(self._adv); av.setSpacing(6)
        av.addWidget(_sechead("FAMILY MEMBERS"))
        fg = QGridLayout(); fg.setSpacing(6)
        self._fam_fields: dict = {}
        fams = [("Father", "father"), ("Mother", "mother"), ("Spouse", "spouse")]
        fg.addWidget(QLabel("Member"),      0, 0)
        fg.addWidget(QLabel("First Names"), 0, 1)
        fg.addWidget(QLabel("Last Names"),  0, 2)
        fg.addWidget(QLabel("Exact"),       0, 3)
        for ri, (label, key) in enumerate(fams, 1):
            fg.addWidget(QLabel(label + ":"), ri, 0)
            ff = QLineEdit(); lf = QLineEdit(); cb = QCheckBox()
            cb.setToolTip("Exact match for this family member")
            self._fam_fields[key] = (ff, lf, cb)
            fg.addWidget(ff, ri, 1); fg.addWidget(lf, ri, 2); fg.addWidget(cb, ri, 3)
        fg.setColumnStretch(1, 2); fg.setColumnStretch(2, 2)
        av.addLayout(fg)
        av.addWidget(_divider())
        kr = QHBoxLayout(); kr.setSpacing(8)
        self.f_keyword = QLineEdit()
        self.f_keyword.setPlaceholderText("Keyword (occupation, etc.)")
        kr.addWidget(QLabel("Keyword:")); kr.addWidget(self.f_keyword, 1)
        av.addLayout(kr)

        self._adv_scroll = QScrollArea()
        self._adv_scroll.setWidget(self._adv)
        self._adv_scroll.setWidgetResizable(True)
        self._adv_scroll.setFrameShape(QFrame.NoFrame)
        self._adv_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._adv_scroll.setVisible(False)
        self._outer.addWidget(self._adv_scroll)

        # Output
        og = QGroupBox("Output")
        ol = QVBoxLayout(og); ol.setSpacing(6)
        fr = QHBoxLayout()
        self.f_docx = QCheckBox("Word (.docx)"); self.f_docx.setChecked(True)
        self.f_xlsx = QCheckBox("Excel (.xlsx)"); self.f_xlsx.setChecked(True)
        fr.addWidget(self.f_docx); fr.addWidget(self.f_xlsx); fr.addStretch()
        fr.addWidget(QLabel("(Images saved in sub-folder 'images')"))
        ol.addLayout(fr)
        dr = QHBoxLayout(); dr.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80)
        bb.clicked.connect(self._browse)
        dr.addWidget(QLabel("Save to:")); dr.addWidget(self.f_folder, 1); dr.addWidget(bb)
        ol.addLayout(dr)
        self._outer.addWidget(og)

        self.pbar  = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        self._outer.addWidget(self.pbar)
        self._outer.addWidget(self.stlbl)

        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn); br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(
            QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        for w in self._all_widgets():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)
        self._fit()

    # ── Advanced toggle / fit ─────────────────────────────────────────────── #
    def _toggle_adv(self, on: bool):
        self._adv_scroll.setVisible(on)
        if on:
            screen_h = QApplication.primaryScreen().availableGeometry().height()
            self._adv_scroll.setMaximumHeight(int(screen_h * 0.5))
        self._adv_btn.setText(("▼" if on else "▶") + "   Advanced Search")
        self._fit()

    def _fit(self):
        QApplication.processEvents()
        self.setMinimumHeight(0); self.setMaximumHeight(16777215)
        hint = self.centralWidget().sizeHint()
        self.resize(self.width(), hint.height() + 42)

    # ── Autosave ──────────────────────────────────────────────────────────── #
    def _all_widgets(self) -> list:
        ws = [self.f_user, self.f_pass, self.f_first, self.f_last,
              self.f_place, self.f_byear, self.f_first_exact, self.f_last_exact,
              self.f_place_exact, self.f_byear_exact, self.f_keyword,
              self.f_folder, self.f_docx, self.f_xlsx]
        for ff, lf, cb in self._fam_fields.values():
            ws += [ff, lf, cb]
        return ws

    def _save(self, *_):
        d = {
            "username": self.f_user.text(), "password": self.f_pass.text(),
            "first_names": self.f_first.text(), "last_names": self.f_last.text(),
            "place_lived": self.f_place.text(), "birth_year": self.f_byear.text(),
            "first_exact": self.f_first_exact.isChecked(),
            "last_exact":  self.f_last_exact.isChecked(),
            "place_exact": self.f_place_exact.isChecked(),
            "byear_exact": self.f_byear_exact.isChecked(),
            "keyword": self.f_keyword.text(),
            "output_folder": self.f_folder.text(),
            "fmt_docx": self.f_docx.isChecked(), "fmt_xlsx": self.f_xlsx.isChecked(),
            "adv_open": self._adv_btn.isChecked(),
        }
        for key, (ff, lf, cb) in self._fam_fields.items():
            d[f"{key}_first"] = ff.text()
            d[f"{key}_last"]  = lf.text()
            d[f"{key}_exact"] = cb.isChecked()
        try:
            _SAVE.write_text(json.dumps(d, ensure_ascii=False, indent=2),
                             encoding="utf-8")
        except Exception:
            pass

    def _load(self):
        if not _SAVE.exists():
            return
        try:
            d = json.loads(_SAVE.read_text(encoding="utf-8"))
        except Exception:
            return

        def _s(w, k):
            if k not in d: return
            v = d[k]
            if   isinstance(w, QLineEdit): w.setText(str(v))
            elif isinstance(w, QCheckBox): w.setChecked(bool(v))

        _s(self.f_user, "username"); _s(self.f_pass, "password")
        _s(self.f_first, "first_names"); _s(self.f_last, "last_names")
        _s(self.f_place, "place_lived"); _s(self.f_byear, "birth_year")
        _s(self.f_first_exact, "first_exact"); _s(self.f_last_exact, "last_exact")
        _s(self.f_place_exact, "place_exact"); _s(self.f_byear_exact, "byear_exact")
        _s(self.f_keyword, "keyword"); _s(self.f_folder, "output_folder")
        _s(self.f_docx, "fmt_docx"); _s(self.f_xlsx, "fmt_xlsx")
        for key, (ff, lf, cb) in self._fam_fields.items():
            _s(ff, f"{key}_first"); _s(lf, f"{key}_last"); _s(cb, f"{key}_exact")
        if d.get("adv_open"):
            self._adv_btn.setChecked(True)

    # ── Helpers ───────────────────────────────────────────────────────────── #
    def _browse(self):
        p = QFileDialog.getExistingDirectory(
            self, "Select output folder", self.f_folder.text() or _DEF_DIR)
        if p: self.f_folder.setText(p)

    def _fmt(self) -> str:
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _build_advanced(self) -> dict:
        adv = {"keyword": self.f_keyword.text().strip()}
        for key, (ff, lf, cb) in self._fam_fields.items():
            adv[f"{key}_first"] = ff.text().strip()
            adv[f"{key}_last"]  = lf.text().strip()
            adv[f"{key}_exact"] = cb.isChecked()
        return adv

    def _payload(self) -> dict:
        return {
            "first_names": self.f_first.text().strip(),
            "last_names":  self.f_last.text().strip(),
            "place_lived": self.f_place.text().strip(),
            "birth_year":  self.f_byear.text().strip(),
            "advanced":    self._build_advanced(),
            "exact": {
                "name":    self.f_first_exact.isChecked(),
                "surname": self.f_last_exact.isChecked(),
                "place":   self.f_place_exact.isChecked(),
                "year":    self.f_byear_exact.isChecked(),
            },
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "email":    self.f_user.text().strip() or None,
            "password": self.f_pass.text() or None,
            "log":      print,
            "cancel_event": None,
        }

    def _validate(self) -> bool:
        if not self.f_first.text().strip() and not self.f_last.text().strip():
            QMessageBox.warning(self, "Nothing to search",
                                "Enter at least a first or last name.")
            return False
        if not self.f_docx.isChecked() and not self.f_xlsx.isChecked():
            QMessageBox.warning(self, "No output format",
                                "Select at least one output format.")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Scraper not found",
                                 "ancestry_scraper.py not found in project root.")
            return False
        return True

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
        choice = ("append" if clicked is b_app else
                  "skip" if clicked is b_skip else "overwrite")
        self.worker.provide_file_choice(choice)

    # ── Start / finish ────────────────────────────────────────────────────── #
    def _start(self):
        if not self._validate(): return
        self.start_btn.setEnabled(False)
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
    w = AncestryApp()
    w.show()
    sys.exit(app.exec())
