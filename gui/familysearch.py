"""
gui/familysearch.py
-------------------
FamilySearch Search window.

Layout
------
• Logo + credentials (email / password)
• Basic search: First Names, Last Names, Place Lived, Birth Year
• Tab selector: All / Historical Records / Family Tree Profiles / Memories
• Advanced Search section (collapsible) — fields match the modal seen
  in the screenshots, organised into groups:
    Ancestor Info: Alternate Name (first+last), Sex
    Life Events:   Birth / Marriage / Residence / Death / Any
                   → each has Place, Year, Exact+/- checkbox
    Family Members: Father, Mother, Spouse, Other Person (first+last)
    Record Options: Country, State
    Keywords, Show Exact Search toggle
• Output: docx / xlsx checkboxes + folder picker
• Progress bar + Start button
• Autosave all fields to .fs_autosave.json
"""

import json, sys
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox, QComboBox, QSpinBox,
    QFrame, QGridLayout,
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

def _svg_icon(svg, size=20):
    from PySide6.QtSvg import QSvgRenderer
    from PySide6.QtGui import QPixmap, QPainter
    r = QSvgRenderer(QByteArray(svg))
    pix = QPixmap(size, size); pix.fill(Qt.transparent)
    p = QPainter(pix); r.render(p); p.end()
    return QIcon(pix)

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
QLineEdit,QComboBox,QSpinBox{padding:4px 6px;border:1px solid #b0ccc8;
  border-radius:4px;background:white;min-height:22px;}
QLineEdit:focus,QComboBox:focus,QSpinBox:focus{border:1px solid #006B6B;}
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
QLabel#sechead{font-weight:bold;color:#006B6B;font-size:10px;
  margin-top:4px;margin-bottom:2px;}
QFrame#divider{background:#c8dede;}
"""

# ── Password field ────────────────────────────────────────────────────────── #
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

# ── Worker ────────────────────────────────────────────────────────────────── #
class Worker(QThread):
    progress = Signal(int, str)
    finished = Signal(dict)
    def __init__(self, payload):
        super().__init__(); self.payload = payload
    def run(self):
        import asyncio
        self.payload["progress"] = lambda v,t: self.progress.emit(int(v), str(t))
        try:
            result = asyncio.run(_scraper.run_scraper(**self.payload))
        except Exception as exc:
            result = {"ok":False,"error":"exception",
                      "message":f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)

# ── Small helpers ─────────────────────────────────────────────────────────── #
def _divider():
    f = QFrame(); f.setObjectName("divider")
    f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
    return f

def _section_label(text):
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

    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(18, 12, 18, 12)
        self._outer.setSpacing(8)

        # ── Logo ─────────────────────────────────────────────────────────── #
        lbl = QLabel()
        pix = QPixmap(str(_CONFIG / "familysearch.png"))
        if not pix.isNull():
            lbl.setPixmap(pix.scaledToWidth(220, Qt.SmoothTransformation))
        else:
            lbl.setText("🌳  FamilySearch")
            lbl.setStyleSheet("font-size:22px;font-weight:bold;color:#006B6B;")
        lbl.setAlignment(Qt.AlignLeft)
        self._outer.addWidget(lbl)

        # ── Credentials ──────────────────────────────────────────────────── #
        cg = QGroupBox("Account credentials")
        cl = QHBoxLayout(cg); cl.setSpacing(8)
        self.f_email = QLineEdit(); self.f_email.setPlaceholderText("Email / Username")
        self.f_pass  = PwdEdit();   self.f_pass.setPlaceholderText("Password")
        cl.addWidget(QLabel("Email:")); cl.addWidget(self.f_email, 2)
        cl.addWidget(QLabel("Password:")); cl.addWidget(self.f_pass, 2)
        self._outer.addWidget(cg)

        # ── Basic search ─────────────────────────────────────────────────── #
        bg = QGroupBox("Basic Search")
        bf = QGridLayout(bg); bf.setSpacing(8)
        self.f_first  = QLineEdit(); self.f_first.setPlaceholderText("Ancestor's First and Middle Names")
        self.f_last   = QLineEdit(); self.f_last.setPlaceholderText("Ancestor's Last or Maiden Names")
        self.f_place  = QLineEdit(); self.f_place.setPlaceholderText("City, County, State, Country")
        self.f_byear  = QLineEdit(); self.f_byear.setPlaceholderText("e.g. 1897")
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

        # ── Tab selector ─────────────────────────────────────────────────── #
        tg = QGroupBox("Search tab")
        tl = QHBoxLayout(tg); tl.setSpacing(10)
        self.f_tab = QComboBox(); self.f_tab.addItems(TAB_OPTIONS)
        tl.addWidget(QLabel("Tab:")); tl.addWidget(self.f_tab); tl.addStretch()
        self._outer.addWidget(tg)

        # ── Advanced Search (collapsible) ─────────────────────────────────── #
        self._adv_btn = QPushButton("▶   Advanced Search")
        self._adv_btn.setObjectName("advBtn")
        self._adv_btn.setCheckable(True)
        self._adv_btn.toggled.connect(self._toggle_adv)
        self._outer.addWidget(self._adv_btn)

        self._adv = QGroupBox()
        av = QVBoxLayout(self._adv); av.setSpacing(6)

        # — Ancestor Info —
        av.addWidget(_section_label("ANCESTOR INFORMATION"))
        ai_row = QHBoxLayout()
        self.f_alt_first = QLineEdit(); self.f_alt_first.setPlaceholderText("Alternate First Name")
        self.f_alt_last  = QLineEdit(); self.f_alt_last.setPlaceholderText("Alternate Last Name")
        self.f_sex       = QComboBox(); self.f_sex.addItems(GENDER_OPTIONS)
        ai_row.addWidget(QLabel("Alt Name:")); ai_row.addWidget(self.f_alt_first)
        ai_row.addWidget(self.f_alt_last)
        ai_row.addWidget(QLabel("Sex:")); ai_row.addWidget(self.f_sex)
        av.addLayout(ai_row)
        av.addWidget(_divider())

        # — Life Events —
        av.addWidget(_section_label("LIFE EVENTS"))
        events_grid = QGridLayout(); events_grid.setSpacing(6)
        self._event_fields = {}
        event_labels = [
            ("Birth",     "birth"),
            ("Marriage",  "marriage"),
            ("Residence", "residence"),
            ("Death",     "death"),
            ("Any",       "any"),
        ]
        events_grid.addWidget(QLabel("Event"),    0, 0)
        events_grid.addWidget(QLabel("Place"),    0, 1)
        events_grid.addWidget(QLabel("Year"),     0, 2)
        events_grid.addWidget(QLabel("Exact+/-"), 0, 3)
        for row_i, (label, key) in enumerate(event_labels, 1):
            events_grid.addWidget(QLabel(label + ":"), row_i, 0)
            place_f = QLineEdit(); place_f.setPlaceholderText("City, County, State, Province, or Country")
            year_f  = QLineEdit(); year_f.setPlaceholderText("Year"); year_f.setFixedWidth(80)
            exact_cb = QCheckBox()
            self._event_fields[key] = (place_f, year_f, exact_cb)
            events_grid.addWidget(place_f,   row_i, 1)
            events_grid.addWidget(year_f,    row_i, 2)
            events_grid.addWidget(exact_cb,  row_i, 3)
        events_grid.setColumnStretch(1, 3)
        av.addLayout(events_grid)
        av.addWidget(_divider())

        # — Family Members —
        av.addWidget(_section_label("FAMILY MEMBERS"))
        fam_grid = QGridLayout(); fam_grid.setSpacing(6)
        self._fam_fields = {}
        fam_labels = [
            ("Father",      "father"),
            ("Mother",      "mother"),
            ("Spouse",      "spouse"),
            ("Other Person","other"),
        ]
        fam_grid.addWidget(QLabel("Member"), 0, 0)
        fam_grid.addWidget(QLabel("First Names"), 0, 1)
        fam_grid.addWidget(QLabel("Last Names"), 0, 2)
        for row_i, (label, key) in enumerate(fam_labels, 1):
            fam_grid.addWidget(QLabel(label + ":"), row_i, 0)
            first_f = QLineEdit()
            last_f  = QLineEdit()
            self._fam_fields[key] = (first_f, last_f)
            fam_grid.addWidget(first_f, row_i, 1)
            fam_grid.addWidget(last_f,  row_i, 2)
        fam_grid.setColumnStretch(1, 2); fam_grid.setColumnStretch(2, 2)
        av.addLayout(fam_grid)
        av.addWidget(_divider())

        # — Record Options —
        av.addWidget(_section_label("RECORD OPTIONS (LOCATION)"))
        rec_row = QHBoxLayout(); rec_row.setSpacing(8)
        self.f_country = QLineEdit(); self.f_country.setPlaceholderText("Country or Location")
        self.f_state   = QLineEdit(); self.f_state.setPlaceholderText("State or Province")
        rec_row.addWidget(QLabel("Country:")); rec_row.addWidget(self.f_country)
        rec_row.addWidget(QLabel("State:"));   rec_row.addWidget(self.f_state)
        av.addLayout(rec_row)
        av.addWidget(_divider())

        # — Keywords + Exact Search —
        kw_row = QHBoxLayout(); kw_row.setSpacing(8)
        self.f_keywords   = QLineEdit(); self.f_keywords.setPlaceholderText("Keywords")
        self.f_show_exact = QCheckBox("Show Exact Search")
        kw_row.addWidget(QLabel("Keywords:")); kw_row.addWidget(self.f_keywords, 2)
        kw_row.addWidget(self.f_show_exact)
        av.addLayout(kw_row)

        self._adv.setVisible(False)
        self._outer.addWidget(self._adv)

        # ── Output ───────────────────────────────────────────────────────── #
        og = QGroupBox("Output")
        ol = QVBoxLayout(og); ol.setSpacing(6)
        fr = QHBoxLayout()
        self.f_docx = QCheckBox("Word (.docx)"); self.f_docx.setChecked(True)
        self.f_xlsx = QCheckBox("Excel (.xlsx)"); self.f_xlsx.setChecked(True)
        fr.addWidget(self.f_docx); fr.addWidget(self.f_xlsx); fr.addStretch()
        fr.addWidget(QLabel("(Document images saved in sub-folder 'images')"))
        ol.addLayout(fr)
        dr = QHBoxLayout(); dr.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80)
        bb.clicked.connect(self._browse)
        dr.addWidget(QLabel("Save to:")); dr.addWidget(self.f_folder, 1); dr.addWidget(bb)
        ol.addLayout(dr)
        self._outer.addWidget(og)

        # ── Progress ─────────────────────────────────────────────────────── #
        self.pbar  = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        self._outer.addWidget(self.pbar)
        self._outer.addWidget(self.stlbl)

        # ── Start ─────────────────────────────────────────────────────────── #
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn); br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(
            QLabel("© Alla Khananashvili", alignment=Qt.AlignRight))

        # Autosave wiring
        for w in self._autosave_widgets():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QComboBox): w.currentTextChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)

        self._fit()

    # ── Advanced toggle ───────────────────────────────────────────────────── #
    def _toggle_adv(self, on):
        self._adv.setVisible(on)
        self._adv_btn.setText(("▼" if on else "▶") + "   Advanced Search")
        self._fit()

    def _fit(self):
        self.adjustSize()
        self.setFixedHeight(self.sizeHint().height())

    # ── Widget list for autosave ──────────────────────────────────────────── #
    def _autosave_widgets(self):
        ws = [self.f_email, self.f_pass, self.f_first, self.f_last,
              self.f_place, self.f_byear, self.f_tab,
              self.f_alt_first, self.f_alt_last, self.f_sex,
              self.f_country, self.f_state, self.f_keywords,
              self.f_show_exact, self.f_folder, self.f_docx, self.f_xlsx]
        for place_f, year_f, exact_cb in self._event_fields.values():
            ws += [place_f, year_f, exact_cb]
        for first_f, last_f in self._fam_fields.values():
            ws += [first_f, last_f]
        return ws

    # ── Autosave ──────────────────────────────────────────────────────────── #
    def _save(self, *_):
        d = {
            "email": self.f_email.text(), "password": self.f_pass.text(),
            "first_names": self.f_first.text(), "last_names": self.f_last.text(),
            "place_lived": self.f_place.text(), "birth_year": self.f_byear.text(),
            "tab": self.f_tab.currentText(),
            "alt_first": self.f_alt_first.text(), "alt_last": self.f_alt_last.text(),
            "sex": self.f_sex.currentText(),
            "country": self.f_country.text(), "state": self.f_state.text(),
            "keywords": self.f_keywords.text(),
            "show_exact": self.f_show_exact.isChecked(),
            "output_folder": self.f_folder.text(),
            "fmt_docx": self.f_docx.isChecked(), "fmt_xlsx": self.f_xlsx.isChecked(),
            "adv_open": self._adv_btn.isChecked(),
        }
        for key, (place_f, year_f, exact_cb) in self._event_fields.items():
            d[f"{key}_place"] = place_f.text()
            d[f"{key}_year"]  = year_f.text()
            d[f"{key}_exact"] = exact_cb.isChecked()
        for key, (first_f, last_f) in self._fam_fields.items():
            d[f"{key}_first"] = first_f.text()
            d[f"{key}_last"]  = last_f.text()
        try:
            _SAVE.write_text(json.dumps(d, ensure_ascii=False, indent=2),
                             encoding="utf-8")
        except Exception:
            pass

    def _load(self):
        if not _SAVE.exists(): return
        try: d = json.loads(_SAVE.read_text(encoding="utf-8"))
        except Exception: return
        def s(w, k):
            if k not in d: return
            v = d[k]
            if   isinstance(w, QLineEdit): w.setText(str(v))
            elif isinstance(w, QComboBox):
                i = w.findText(str(v))
                if i >= 0: w.setCurrentIndex(i)
            elif isinstance(w, QCheckBox): w.setChecked(bool(v))
        s(self.f_email,"email"); s(self.f_pass,"password")
        s(self.f_first,"first_names"); s(self.f_last,"last_names")
        s(self.f_place,"place_lived"); s(self.f_byear,"birth_year")
        s(self.f_tab,"tab")
        s(self.f_alt_first,"alt_first"); s(self.f_alt_last,"alt_last")
        s(self.f_sex,"sex")
        s(self.f_country,"country"); s(self.f_state,"state")
        s(self.f_keywords,"keywords"); s(self.f_show_exact,"show_exact")
        s(self.f_folder,"output_folder")
        s(self.f_docx,"fmt_docx"); s(self.f_xlsx,"fmt_xlsx")
        for key, (place_f, year_f, exact_cb) in self._event_fields.items():
            s(place_f, f"{key}_place"); s(year_f, f"{key}_year")
            s(exact_cb, f"{key}_exact")
        for key, (first_f, last_f) in self._fam_fields.items():
            s(first_f, f"{key}_first"); s(last_f, f"{key}_last")
        if d.get("adv_open"):
            self._adv_btn.setChecked(True)

    # ── Helpers ───────────────────────────────────────────────────────────── #
    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Select output folder",
                                              self.f_folder.text() or _DEF_DIR)
        if p: self.f_folder.setText(p)

    def _fmt(self):
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _build_advanced(self):
        adv = {}
        adv["sex"]          = self.f_sex.currentText()
        adv["alt_first"]    = self.f_alt_first.text().strip()
        adv["alt_last"]     = self.f_alt_last.text().strip()
        adv["country"]      = self.f_country.text().strip()
        adv["state"]        = self.f_state.text().strip()
        adv["keywords"]     = self.f_keywords.text().strip()
        adv["show_exact"]   = self.f_show_exact.isChecked()
        for key, (place_f, year_f, exact_cb) in self._event_fields.items():
            adv[f"{key}_place"] = place_f.text().strip()
            adv[f"{key}_year"]  = year_f.text().strip()
            adv[f"{key}_exact"] = exact_cb.isChecked()
        for key, (first_f, last_f) in self._fam_fields.items():
            adv[f"{key}_first"] = first_f.text().strip()
            adv[f"{key}_last"]  = last_f.text().strip()
        # Only return if anything non-default was set
        has_data = any([
            adv["sex"] != "Unspecified",
            adv["alt_first"], adv["alt_last"],
            adv["country"], adv["state"], adv["keywords"],
        ] + [v for k,v in adv.items()
               if k.endswith("_place") or k.endswith("_year") or
                  (k.endswith("_first") or k.endswith("_last"))])
        return adv if has_data else {}

    def _payload(self):
        return {
            "first_names":   self.f_first.text().strip(),
            "last_names":    self.f_last.text().strip(),
            "place_lived":   self.f_place.text().strip(),
            "birth_year":    self.f_byear.text().strip(),
            "tab":           self.f_tab.currentText(),
            "advanced":      self._build_advanced(),
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "email":         self.f_email.text().strip() or None,
            "password":      self.f_pass.text() or None,
            "log":           print,
            "cancel_event":  None,
        }

    def _validate(self):
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

    def _start(self):
        if not self._validate(): return
        self.start_btn.setEnabled(False)
        self.pbar.setValue(0); self.stlbl.setText("Starting…")
        self.worker = Worker(self._payload())
        self.worker.progress.connect(
            lambda v,t: (self.pbar.setValue(v), self.stlbl.setText(t)))
        self.worker.finished.connect(self._done)
        self.worker.start()

    def _done(self, r):
        self.start_btn.setEnabled(True)
        if r.get("ok"):
            n = r.get("n_records", 0)
            parts = (["Word"] if r.get("docx_count") else []) + \
                    (["Excel"] if r.get("xlsx_path") else [])
            msg = f"{n} record(s) saved"
            if parts: msg += " → " + " + ".join(parts)
            if r.get("output_folder"):
                msg += f"\n\nFolder:\n{r['output_folder']}"
            QMessageBox.information(self, "Done", msg)
            self.stlbl.setText("Done.")
        else:
            QMessageBox.critical(self, "Error",
                f"Search failed.\n\n{r.get('message','')}\n\nCheck terminal.")
            self.stlbl.setText("Error — see terminal.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = FamilySearchApp()
    w.show()
    sys.exit(app.exec())
