"""
gui/familysearch.py  —  v5
===========================
FamilySearch search window.
All annotations in English.

Window resizes correctly when Advanced Search collapses/expands.
Logo: FSlogo.png
Autosave: .fs_autosave.json
"""

import json, sys
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox, QComboBox,
    QFrame, QGridLayout, QScrollArea,
)
from PySide6.QtCore import QThread, Signal, Qt, QByteArray
from PySide6.QtGui import QPixmap, QIcon

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
_CONFIG  = _ROOT / "config"
_SAVE    = _HERE / ".fs_autosave.json"
_DEF_DIR = str(Path.home() / "Downloads" / "FamilySearch_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
try:
    import familysearch_scraper as _scraper
    _SCRAPER_OK = True
except ImportError:
    _SCRAPER_OK = False

TAB_OPTIONS    = ["All", "Historical Records", "Family Tree Profiles", "Memories"]
GENDER_OPTIONS = ["Unspecified", "Male", "Female"]

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b0ccc8;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f5fafa;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#006B6B;background:#f5fafa;}
QLineEdit,QComboBox{padding:4px 6px;border:1px solid #b0ccc8;
  border-radius:4px;background:white;min-height:22px;}
QLineEdit:focus,QComboBox:focus{border:1px solid #006B6B;}
QPushButton{padding:5px 14px;border-radius:4px;
  border:1px solid #b0ccc8;background:#eef6f5;}
QPushButton:hover{background:#d5eeec;}
QPushButton:pressed{background:#bde0de;}
QPushButton#startBtn{background:#006B6B;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#008080;}
QPushButton#startBtn:disabled{background:#7ab5b5;}
QPushButton#eyeBtn{border:none;background:transparent;padding:0;}
QPushButton#eyeBtn:hover{background:#d5eeec;border-radius:3px;}
QPushButton#advBtn{text-align:left;border:none;background:transparent;
  color:#006B6B;font-weight:bold;font-size:11px;padding:2px 0;}
QPushButton#advBtn:hover{color:#008080;}
QProgressBar{border:1px solid #b0ccc8;border-radius:4px;
  text-align:center;min-height:18px;}
QProgressBar::chunk{background:#006B6B;border-radius:3px;}
QLabel#sechead{font-weight:bold;color:#006B6B;font-size:10px;margin-top:4px;}
QFrame#div{background:#c8dede;}
"""


# ── Password field with eye toggle ────────────────────────────────────────── #
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


# ── Background worker thread ──────────────────────────────────────────────── #
class Worker(QThread):
    progress = Signal(int, str)
    finished = Signal(dict)

    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload

    def run(self):
        import asyncio
        self.payload["progress"] = lambda v, t: self.progress.emit(int(v), str(t))
        try:
            result = asyncio.run(_scraper.run_scraper(**self.payload))
        except Exception as exc:
            result = {"ok": False, "error": "exception",
                      "message": f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)


# ── Small UI helpers ──────────────────────────────────────────────────────── #
def _divider() -> QFrame:
    f = QFrame(); f.setObjectName("div")
    f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
    return f


def _sechead(text: str) -> QLabel:
    l = QLabel(text); l.setObjectName("sechead"); return l


# ── Main window ───────────────────────────────────────────────────────────── #
class FamilySearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FamilySearch")
        self.setMinimumWidth(900)
        self.setStyleSheet(STYLE)
        self._build_ui()
        self._load()

    # ── Build UI ──────────────────────────────────────────────────────────── #
    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(18, 12, 18, 12)
        self._outer.setSpacing(8)

        # Logo
        lbl = QLabel()
        pix = QPixmap(str(_CONFIG / "FSlogo.png"))
        if not pix.isNull():
            lbl.setPixmap(pix.scaledToWidth(220, Qt.SmoothTransformation))
        else:
            lbl.setText("FamilySearch")
            lbl.setStyleSheet("font-size:22px;font-weight:bold;color:#006B6B;")
        lbl.setAlignment(Qt.AlignLeft)
        self._outer.addWidget(lbl)

        # Credentials
        cg = QGroupBox("Account credentials")
        cl = QHBoxLayout(cg); cl.setSpacing(8)
        self.f_user = QLineEdit()
        self.f_user.setPlaceholderText("FamilySearch username")
        self.f_pass = PwdEdit()
        self.f_pass.setPlaceholderText("Password")
        cl.addWidget(QLabel("Username:")); cl.addWidget(self.f_user, 2)
        cl.addWidget(QLabel("Password:")); cl.addWidget(self.f_pass, 2)
        self._outer.addWidget(cg)

        # Basic Search
        bg = QGroupBox("Basic Search")
        bf = QGridLayout(bg); bf.setSpacing(8)
        self.f_first = QLineEdit()
        self.f_first.setPlaceholderText("Ancestor's First and Middle Names")
        self.f_last  = QLineEdit()
        self.f_last.setPlaceholderText("Ancestor's Last or Maiden Names")
        self.f_place = QLineEdit()
        self.f_place.setPlaceholderText("City, County, State, Country")
        self.f_byear = QLineEdit()
        self.f_byear.setPlaceholderText("e.g. 1897")
        self.f_byear.setFixedWidth(120)
        bf.addWidget(QLabel("First Names:"), 0, 0)
        bf.addWidget(self.f_first,           0, 1)
        bf.addWidget(QLabel("Last Names:"),  0, 2)
        bf.addWidget(self.f_last,            0, 3)
        bf.addWidget(QLabel("Place Lived:"), 1, 0)
        bf.addWidget(self.f_place,           1, 1)
        bf.addWidget(QLabel("Birth Year:"),  1, 2)
        bf.addWidget(self.f_byear,           1, 3)
        bf.setColumnStretch(1, 2); bf.setColumnStretch(3, 1)
        self._outer.addWidget(bg)

        # Tab selector
        tg = QGroupBox("Search tab")
        tl = QHBoxLayout(tg); tl.setSpacing(10)
        self.f_tab = QComboBox(); self.f_tab.addItems(TAB_OPTIONS)
        tl.addWidget(QLabel("Tab:")); tl.addWidget(self.f_tab); tl.addStretch()
        self._outer.addWidget(tg)

        # Advanced Search toggle button
        self._adv_btn = QPushButton("▶   Advanced Search")
        self._adv_btn.setObjectName("advBtn")
        self._adv_btn.setCheckable(True)
        self._adv_btn.toggled.connect(self._toggle_adv)
        self._outer.addWidget(self._adv_btn)

        # Advanced Search panel (hidden by default)
        self._adv = QGroupBox()
        av = QVBoxLayout(self._adv); av.setSpacing(6)

        # Ancestor Info
        av.addWidget(_sechead("ANCESTOR INFORMATION"))
        ai = QHBoxLayout()
        self.f_alt_first = QLineEdit()
        self.f_alt_first.setPlaceholderText("Alternate First Name")
        self.f_alt_last  = QLineEdit()
        self.f_alt_last.setPlaceholderText("Alternate Last Name")
        self.f_sex = QComboBox(); self.f_sex.addItems(GENDER_OPTIONS)
        ai.addWidget(QLabel("Alt Name:"))
        ai.addWidget(self.f_alt_first)
        ai.addWidget(self.f_alt_last)
        ai.addWidget(QLabel("Sex:"))
        ai.addWidget(self.f_sex)
        av.addLayout(ai)
        av.addWidget(_divider())

        # Life Events grid
        av.addWidget(_sechead("LIFE EVENTS"))
        eg = QGridLayout(); eg.setSpacing(6)
        self._event_fields: dict = {}
        events_list = [
            ("Birth",     "birth"),
            ("Marriage",  "marriage"),
            ("Residence", "residence"),
            ("Death",     "death"),
            ("Any",       "any"),
        ]
        eg.addWidget(QLabel("Event"),    0, 0)
        eg.addWidget(QLabel("Place"),    0, 1)
        eg.addWidget(QLabel("Year"),     0, 2)
        eg.addWidget(QLabel("Exact+/-"), 0, 3)
        for ri, (label, key) in enumerate(events_list, 1):
            eg.addWidget(QLabel(label+":"), ri, 0)
            pf = QLineEdit()
            pf.setPlaceholderText("City, County, State, Province, or Country")
            yf = QLineEdit(); yf.setPlaceholderText("Year"); yf.setFixedWidth(80)
            cb = QCheckBox()
            self._event_fields[key] = (pf, yf, cb)
            eg.addWidget(pf, ri, 1)
            eg.addWidget(yf, ri, 2)
            eg.addWidget(cb, ri, 3)
        eg.setColumnStretch(1, 3)
        av.addLayout(eg)
        av.addWidget(_divider())

        # Family Members grid
        av.addWidget(_sechead("FAMILY MEMBERS"))
        fg = QGridLayout(); fg.setSpacing(6)
        self._fam_fields: dict = {}
        fams_list = [
            ("Father",       "father"),
            ("Mother",       "mother"),
            ("Spouse",       "spouse"),
            ("Other Person", "other"),
        ]
        fg.addWidget(QLabel("Member"),        0, 0)
        fg.addWidget(QLabel("First Names"),   0, 1)
        fg.addWidget(QLabel("Exact"),         0, 2)
        fg.addWidget(QLabel("Last Names"),    0, 3)
        fg.addWidget(QLabel("Exact"),         0, 4)
        for ri, (label, key) in enumerate(fams_list, 1):
            fg.addWidget(QLabel(label+":"), ri, 0)
            ff = QLineEdit(); cb_f = QCheckBox()
            lf = QLineEdit(); cb_l = QCheckBox()
            cb_f.setToolTip("Exact match for First Name")
            cb_l.setToolTip("Exact match for Last Name")
            self._fam_fields[key] = (ff, cb_f, lf, cb_l)
            fg.addWidget(ff,   ri, 1)
            fg.addWidget(cb_f, ri, 2)
            fg.addWidget(lf,   ri, 3)
            fg.addWidget(cb_l, ri, 4)
        fg.setColumnStretch(1, 2); fg.setColumnStretch(3, 2)
        av.addLayout(fg)
        av.addWidget(_divider())

        # Record Options
        av.addWidget(_sechead("RECORD OPTIONS (LOCATION)"))
        rr = QHBoxLayout(); rr.setSpacing(8)
        self.f_country = QLineEdit()
        self.f_country.setPlaceholderText("Country or Location")
        self.f_state   = QLineEdit()
        self.f_state.setPlaceholderText("State or Province")
        rr.addWidget(QLabel("Country:")); rr.addWidget(self.f_country)
        rr.addWidget(QLabel("State:"));   rr.addWidget(self.f_state)
        av.addLayout(rr)
        av.addWidget(_divider())

        # Keywords + Exact Search toggle
        kr = QHBoxLayout(); kr.setSpacing(8)
        self.f_keywords   = QLineEdit()
        self.f_keywords.setPlaceholderText("Keywords")
        self.f_show_exact = QCheckBox("Show Exact Search")
        kr.addWidget(QLabel("Keywords:"))
        kr.addWidget(self.f_keywords, 2)
        kr.addWidget(self.f_show_exact)
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
        dr.addWidget(QLabel("Save to:"))
        dr.addWidget(self.f_folder, 1)
        dr.addWidget(bb)
        ol.addLayout(dr)
        self._outer.addWidget(og)

        # Progress bar + status label
        self.pbar  = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        self._outer.addWidget(self.pbar)
        self._outer.addWidget(self.stlbl)

        # Start button
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn); br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(
            QLabel("© Alla Khananashvili", alignment=Qt.AlignRight))

        # Wire autosave to every interactive widget
        for w in self._all_widgets():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QComboBox): w.currentTextChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)

        self._fit()

    # ── Advanced Search toggle ────────────────────────────────────────────── #
    def _toggle_adv(self, on: bool):
        self._adv_scroll.setVisible(on)
        if on:
            screen_h = QApplication.primaryScreen().availableGeometry().height()
            self._adv_scroll.setMaximumHeight(int(screen_h * 0.55))
        self._adv_btn.setText(("▼" if on else "▶") + "   Advanced Search")
        self._fit()

    def _fit(self):
        """Resize window height to exactly fit visible content."""
        QApplication.processEvents()
        self.setMinimumHeight(0)
        self.setMaximumHeight(16777215)   # remove any fixed-height constraint
        hint = self.centralWidget().sizeHint()
        self.resize(self.width(), hint.height() + 42)  # +42 for title bar

    # ── All autosave-able widgets ─────────────────────────────────────────── #
    def _all_widgets(self) -> list:
        ws = [self.f_user, self.f_pass,
              self.f_first, self.f_last, self.f_place, self.f_byear,
              self.f_tab,
              self.f_alt_first, self.f_alt_last, self.f_sex,
              self.f_country, self.f_state,
              self.f_keywords, self.f_show_exact,
              self.f_folder, self.f_docx, self.f_xlsx]
        for pf, yf, cb in self._event_fields.values():
            ws += [pf, yf, cb]
        for ff, cb_f, lf, cb_l in self._fam_fields.values():
            ws += [ff, cb_f, lf, cb_l]
        return ws

    # ── Autosave / load ───────────────────────────────────────────────────── #
    def _save(self, *_):
        """Save all field values to .fs_autosave.json."""
        d = {
            "username":      self.f_user.text(),
            "password":      self.f_pass.text(),
            "first_names":   self.f_first.text(),
            "last_names":    self.f_last.text(),
            "place_lived":   self.f_place.text(),
            "birth_year":    self.f_byear.text(),
            "tab":           self.f_tab.currentText(),
            "alt_first":     self.f_alt_first.text(),
            "alt_last":      self.f_alt_last.text(),
            "sex":           self.f_sex.currentText(),
            "country":       self.f_country.text(),
            "state":         self.f_state.text(),
            "keywords":      self.f_keywords.text(),
            "show_exact":    self.f_show_exact.isChecked(),
            "output_folder": self.f_folder.text(),
            "fmt_docx":      self.f_docx.isChecked(),
            "fmt_xlsx":      self.f_xlsx.isChecked(),
            "adv_open":      self._adv_btn.isChecked(),
        }
        for key, (pf, yf, cb) in self._event_fields.items():
            d[f"{key}_place"] = pf.text()
            d[f"{key}_year"]  = yf.text()
            d[f"{key}_exact"] = cb.isChecked()
        for key, (ff, cb_f, lf, cb_l) in self._fam_fields.items():
            d[f"{key}_first"]       = ff.text()
            d[f"{key}_first_exact"] = cb_f.isChecked()
            d[f"{key}_last"]        = lf.text()
            d[f"{key}_last_exact"]  = cb_l.isChecked()
        try:
            _SAVE.write_text(
                json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load(self):
        """Restore field values from .fs_autosave.json on startup."""
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
            elif isinstance(w, QComboBox):
                i = w.findText(str(v))
                if i >= 0: w.setCurrentIndex(i)
            elif isinstance(w, QCheckBox): w.setChecked(bool(v))

        _s(self.f_user,  "username"); _s(self.f_pass,  "password")
        _s(self.f_first, "first_names"); _s(self.f_last,  "last_names")
        _s(self.f_place, "place_lived"); _s(self.f_byear, "birth_year")
        _s(self.f_tab,   "tab")
        _s(self.f_alt_first, "alt_first"); _s(self.f_alt_last, "alt_last")
        _s(self.f_sex,   "sex")
        _s(self.f_country, "country"); _s(self.f_state,   "state")
        _s(self.f_keywords, "keywords"); _s(self.f_show_exact, "show_exact")
        _s(self.f_folder, "output_folder")
        _s(self.f_docx,  "fmt_docx"); _s(self.f_xlsx, "fmt_xlsx")
        for key, (pf, yf, cb) in self._event_fields.items():
            _s(pf, f"{key}_place"); _s(yf, f"{key}_year"); _s(cb, f"{key}_exact")
        for key, (ff, cb_f, lf, cb_l) in self._fam_fields.items():
            _s(ff,   f"{key}_first")
            _s(cb_f, f"{key}_first_exact")
            _s(lf,   f"{key}_last")
            _s(cb_l, f"{key}_last_exact")
        if d.get("adv_open"):
            self._adv_btn.setChecked(True)

    # ── Helpers ───────────────────────────────────────────────────────────── #
    def _browse(self):
        p = QFileDialog.getExistingDirectory(
            self, "Select output folder",
            self.f_folder.text() or _DEF_DIR)
        if p: self.f_folder.setText(p)

    def _fmt(self) -> str:
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _build_advanced(self) -> dict:
        """Collect all advanced search field values into a dict."""
        adv = {
            "sex":        self.f_sex.currentText(),
            "alt_first":  self.f_alt_first.text().strip(),
            "alt_last":   self.f_alt_last.text().strip(),
            "country":    self.f_country.text().strip(),
            "state":      self.f_state.text().strip(),
            "keywords":   self.f_keywords.text().strip(),
            "show_exact": self.f_show_exact.isChecked(),
        }
        for key, (pf, yf, cb) in self._event_fields.items():
            adv[f"{key}_place"] = pf.text().strip()
            adv[f"{key}_year"]  = yf.text().strip()
            adv[f"{key}_exact"] = cb.isChecked()
        for key, (ff, cb_f, lf, cb_l) in self._fam_fields.items():
            adv[f"{key}_first"]       = ff.text().strip()
            adv[f"{key}_first_exact"] = cb_f.isChecked()
            adv[f"{key}_last"]        = lf.text().strip()
            adv[f"{key}_last_exact"]  = cb_l.isChecked()
        # Return only if something non-default was set
        has = any([
            adv["sex"] != "Unspecified",
            adv["alt_first"], adv["alt_last"],
            adv["country"], adv["state"], adv["keywords"],
        ] + [v for k, v in adv.items()
               if isinstance(v, str) and v and
                  any(k.endswith(s)
                      for s in ["_place","_year","_first","_last"])])
        return adv if has else {}

    def _payload(self) -> dict:
        return {
            "first_names":   self.f_first.text().strip(),
            "last_names":    self.f_last.text().strip(),
            "place_lived":   self.f_place.text().strip(),
            "birth_year":    self.f_byear.text().strip(),
            "tab":           self.f_tab.currentText(),
            "advanced":      self._build_advanced(),
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "email":         self.f_user.text().strip() or None,
            "password":      self.f_pass.text() or None,
            "log":           print,
            "cancel_event":  None,
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
                                 "familysearch_scraper.py not found in project root.")
            return False
        return True

    # ── Start / finish ────────────────────────────────────────────────────── #
    def _start(self):
        if not self._validate(): return
        self.start_btn.setEnabled(False)
        self.pbar.setValue(0)
        self.stlbl.setText("Starting...")
        self.worker = Worker(self._payload())
        self.worker.progress.connect(
            lambda v, t: (self.pbar.setValue(v), self.stlbl.setText(t)))
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
    w = FamilySearchApp()
    w.show()
    sys.exit(app.exec())
