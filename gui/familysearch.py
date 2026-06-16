"""
gui/familysearch.py  —  v5
===========================
FamilySearch search window.
All annotations in English.

Window resizes correctly when Advanced Search collapses/expands.
Logo: FSlogo.png
Autosave: .fs_autosave.json
"""

import json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox, QComboBox,
    QFrame, QGridLayout, QScrollArea,
)
from gui._app_icon import app_icon, make_header, make_cancel_button
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

# Life-event year range (like the site: Exact to / this year / ±1 / ±2 / ±5 / ±10)
YEAR_RANGE = [("Exact (this year)", 0), ("± 1 year", 1), ("± 2 years", 2),
              ("± 5 years", 5), ("± 10 years", 10)]
EVENTS = [("Birth", "birth"), ("Marriage", "marriage"), ("Residence", "residence"),
          ("Death", "death"), ("Any", "any")]
# Family members — tabs/order exactly like the site; max 3 of each
FAMILY = [("Spouse", "spouse"), ("Father", "father"),
          ("Mother", "mother"), ("Other Person", "other")]
_FAM_LABEL = {k: l for l, k in FAMILY}      # key → display label
FAM_MAX = 3
ALT_MAX = 3                                  # up to 3 alternate names (like the site)
# Record Type checkboxes → f.recordType=<index>, in the site's order
RECORD_TYPES = ["Birth, Baptism, and Christening", "Marriage", "Death",
                "Census, Residence, and Lists", "Immigration and Naturalization",
                "Military", "Probate", "Other"]

def _load_locations() -> dict:
    try:
        return json.loads((_CONFIG / "familysearch_locations.json").read_text("utf-8"))
    except Exception:
        return {}
FS_LOCATIONS = _load_locations()
FS_COUNTRIES = list(FS_LOCATIONS.keys())

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
QPushButton#advBtn:disabled{color:#aebdbd;}
QPushButton#rmBtn{color:#a33;background:transparent;border:none;font-weight:bold;
  max-width:26px;padding:0;}
QPushButton#rmBtn:hover{color:#d33;}
QWidget#famCard,QWidget#evCard{background:#eef6f5;border:1px solid #cfe3e1;
  border-radius:5px;}
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
    progress     = Signal(int, str)
    finished     = Signal(dict)
    request_file = Signal(str)          # emitted when output files already exist

    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload
        self._file_choice = "overwrite"
        self._file_ev     = threading.Event()

    def provide_file_choice(self, choice: str):
        """Called from the main thread after the user picks overwrite/append/skip."""
        self._file_choice = choice
        self._file_ev.set()

    def run(self):
        import asyncio
        self.payload["progress"] = lambda v, t: self.progress.emit(int(v), str(t))

        def ask_file_conflict(names):
            """Called from the scraper when output files already exist. Emits
            request_file → main thread shows a dialog; blocks for the choice."""
            self._file_choice = "overwrite"
            self._file_ev.clear()
            self.request_file.emit("\n".join(names))
            self._file_ev.wait(timeout=300)     # up to 5 min for the choice
            return self._file_choice or "overwrite"

        self.payload["ask_file_conflict"] = ask_file_conflict

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
        self.setWindowIcon(app_icon())
        self._build_ui()
        self._load()

    # ── Build UI ──────────────────────────────────────────────────────────── #
    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(18, 12, 18, 12)
        self._outer.setSpacing(8)

        # Logo + name
        self._outer.addLayout(
            make_header("FSlogo.png", "FamilySearch", color="#006B6B"))

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

        # Basic Search — NAME only (place / birth are Life Events, like the site)
        bg = QGroupBox("Basic Search")
        bf = QGridLayout(bg); bf.setSpacing(8)
        self.f_first = QLineEdit()
        self.f_first.setPlaceholderText("Ancestor's First and Middle Names")
        self.f_last  = QLineEdit()
        self.f_last.setPlaceholderText("Ancestor's Last or Maiden Names")
        # Per-field "Exact" checkboxes — revealed by "Show Exact Search"
        self.f_first_exact = QCheckBox("Exact")
        self.f_last_exact  = QCheckBox("Exact")
        self.f_first_exact.setToolTip("Exact match for the first/middle name")
        self.f_last_exact.setToolTip("Exact match for the last/maiden name")
        self._exact_boxes = [self.f_first_exact, self.f_last_exact]
        bf.addWidget(QLabel("First Names:"), 0, 0)
        bf.addWidget(self.f_first,           0, 1)
        bf.addWidget(self.f_first_exact,     0, 2)
        bf.addWidget(QLabel("Last Names:"),  0, 3)
        bf.addWidget(self.f_last,            0, 4)
        bf.addWidget(self.f_last_exact,      0, 5)
        bf.setColumnStretch(1, 2); bf.setColumnStretch(4, 2)
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
        av.setContentsMargins(10, 9, 22, 9)   # reserve room for the scrollbar

        # Ancestor Info — Alternate Names (up to 3, ✕ to remove, like the site)
        # and Sex.
        av.addWidget(_sechead("ANCESTOR INFORMATION"))
        self._alt = []                                  # list of alt-name row dicts
        # Sex — its OWN row (it's a property of the ancestor, NOT of the alt name).
        sx = QHBoxLayout(); sx.setSpacing(8)
        sx_lbl = QLabel("Sex:"); sx_lbl.setFixedWidth(95)
        self.f_sex = QComboBox(); self.f_sex.addItems(GENDER_OPTIONS)
        self.f_sex.setFixedWidth(170)
        sx.addWidget(sx_lbl); sx.addWidget(self.f_sex); sx.addStretch()
        av.addLayout(sx)
        # «+ Alternate Name» on its own row, then the alt-name rows underneath.
        self._alt_link = QPushButton("+ Alternate Name")
        self._alt_link.setObjectName("advBtn")
        self._alt_link.clicked.connect(lambda _=0: self._add_alt_row())
        al = QHBoxLayout(); al.addWidget(self._alt_link); al.addStretch()
        av.addLayout(al)
        self._alt_box = QVBoxLayout(); self._alt_box.setSpacing(3)
        av.addLayout(self._alt_box)
        av.addWidget(_divider())

        # Life Events: Birth / Marriage / Residence / Death / Any — each a CARD
        # like the site: Place (+ own Exact tick) on the left, Year + range
        # dropdown (+ own "Exact +/−" tick) on the right. The Exact ticks are
        # revealed by "Show Exact Search".
        av.addWidget(_sechead("LIFE EVENTS"))
        self._event_fields: dict = {}
        for label, key in EVENTS:
            card = QWidget(); card.setObjectName("evCard")
            cg = QGridLayout(card); cg.setContentsMargins(8, 6, 8, 6); cg.setSpacing(6)
            pf = QLineEdit()
            pf.setPlaceholderText("City, County, State, Province, or Country")
            pcb = QCheckBox("Exact"); pcb.setToolTip(f"Exact match for the {label} place")
            yf = QLineEdit(); yf.setPlaceholderText("Year"); yf.setFixedWidth(70)
            rc = QComboBox()
            for lab, _v in YEAR_RANGE:
                rc.addItem(lab)
            rc.setFixedWidth(150)
            dcb = QCheckBox("Exact +/−")
            dcb.setToolTip(f"Search this exact {label} date range")
            self._event_fields[key] = (pf, pcb, yf, rc, dcb)
            self._exact_boxes += [pcb, dcb]
            cg.addWidget(QLabel(f"{label} Place"), 0, 0)
            cg.addWidget(QLabel(f"{label} Year"),  0, 2, 1, 2)
            cg.addWidget(pf,  1, 0); cg.addWidget(pcb, 1, 1)
            cg.addWidget(yf,  1, 2); cg.addWidget(rc,  1, 3)
            cg.addWidget(dcb, 2, 2, 1, 2)
            cg.setColumnStretch(0, 3)
            av.addWidget(card)
        av.addWidget(_divider())

        # Family Members — tabs add rows (max 3 each), ✕ to remove (like the site)
        av.addWidget(_sechead("FAMILY MEMBERS  (up to 3 of each)"))
        flinks = QHBoxLayout(); flinks.setSpacing(14)
        self._fam = {k: [] for _l, k in FAMILY}        # key → list of row dicts
        self._fam_links = {}
        for label, key in FAMILY:
            lk = QPushButton("+ " + label); lk.setObjectName("advBtn")
            lk.clicked.connect(lambda _=0, k=key: self._add_fam_row(k))
            self._fam_links[key] = lk
            flinks.addWidget(lk)
        flinks.addStretch()
        av.addLayout(flinks)
        self._fam_box = QVBoxLayout(); self._fam_box.setSpacing(3)
        av.addLayout(self._fam_box)
        av.addWidget(_divider())

        # Record Options — Location (Country + State dropdowns)
        av.addWidget(_sechead("RECORD OPTIONS — LOCATION"))
        rr = QHBoxLayout(); rr.setSpacing(8)
        self.f_country = QComboBox(); self.f_country.setEditable(True)
        self.f_country.setInsertPolicy(QComboBox.NoInsert)
        self.f_country.addItem(""); self.f_country.addItems(FS_COUNTRIES)
        self.f_state = QComboBox(); self.f_state.setEditable(True)
        self.f_state.setInsertPolicy(QComboBox.NoInsert); self.f_state.addItem("")
        self.f_country.currentTextChanged.connect(self._reload_states)
        rr.addWidget(QLabel("Country or Location:")); rr.addWidget(self.f_country, 2)
        rr.addWidget(QLabel("State or Province:"));   rr.addWidget(self.f_state, 2)
        av.addLayout(rr)

        # Record Type checkboxes (f.recordType=0..7)
        av.addWidget(_sechead("RECORD TYPE"))
        self.f_rectypes = []
        rtg = QGridLayout(); rtg.setSpacing(4)
        for i, name in enumerate(RECORD_TYPES):
            cb = QCheckBox(name); self.f_rectypes.append(cb)
            rtg.addWidget(cb, i // 2, i % 2)
        av.addLayout(rtg)

        # Batch / Film / Principal
        br2 = QHBoxLayout(); br2.setSpacing(8)
        self.f_batch = QLineEdit(); self.f_batch.setPlaceholderText("Batch Number")
        self.f_film  = QLineEdit()
        self.f_film.setPlaceholderText("Film/Fiche/Image Group Number (DGS)")
        self.f_principal = QCheckBox("Principal on Record")
        br2.addWidget(QLabel("Batch:")); br2.addWidget(self.f_batch)
        br2.addWidget(QLabel("Film/DGS:")); br2.addWidget(self.f_film)
        br2.addWidget(self.f_principal)
        av.addLayout(br2)
        av.addWidget(_divider())

        # Keywords + Show Exact Search (reveals the exact checkboxes for the search)
        kr = QHBoxLayout(); kr.setSpacing(8)
        self.f_keywords   = QLineEdit()
        self.f_keywords.setPlaceholderText("Keywords")
        self.f_show_exact = QCheckBox("Show Exact Search")
        self.f_show_exact.toggled.connect(self._toggle_exact)
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
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(
            QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        # Wire autosave to every interactive widget
        for w in self._all_widgets():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QComboBox): w.currentTextChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)

        self._toggle_exact(self.f_show_exact.isChecked())   # hide exact by default
        self._fit()

    # ── Family-member rows (max 3 of each, ✕ to remove — like the site) ────── #
    def _add_fam_row(self, key, first="", last="", first_exact=False, last_exact=False):
        if len(self._fam[key]) >= FAM_MAX:
            return None
        label = _FAM_LABEL[key]
        row_w = QWidget(); row_w.setObjectName("famCard")
        rl = QHBoxLayout(row_w); rl.setContentsMargins(8, 6, 8, 6); rl.setSpacing(6)
        lbl = QLabel(label + ":"); lbl.setFixedWidth(95)
        ff = QLineEdit(); ff.setPlaceholderText(label + "'s First Names"); ff.setText(first)
        cbf = QCheckBox("Exact"); cbf.setChecked(bool(first_exact))
        lf = QLineEdit(); lf.setPlaceholderText(label + "'s Last Names"); lf.setText(last)
        cbl = QCheckBox("Exact"); cbl.setChecked(bool(last_exact))
        rm = QPushButton("✕"); rm.setObjectName("rmBtn"); rm.setFixedWidth(26)
        rl.addWidget(lbl); rl.addWidget(ff, 2); rl.addWidget(cbf)
        rl.addWidget(lf, 2); rl.addWidget(cbl); rl.addWidget(rm)
        rec = {"w": row_w, "first": ff, "last": lf, "fe": cbf, "le": cbl}
        self._fam[key].append(rec)
        self._fam_box.addWidget(row_w)
        for cb in (cbf, cbl):
            self._exact_boxes.append(cb)
            cb.setVisible(self.f_show_exact.isChecked())
        ff.textChanged.connect(self._save); lf.textChanged.connect(self._save)
        cbf.stateChanged.connect(self._save); cbl.stateChanged.connect(self._save)
        rm.clicked.connect(lambda _=0: self._remove_fam_row(key, rec))
        if len(self._fam[key]) >= FAM_MAX:
            self._fam_links[key].setEnabled(False)
        self._save(); self._fit()
        return rec

    def _remove_fam_row(self, key, rec):
        try:
            self._fam[key].remove(rec)
            for cb in (rec["fe"], rec["le"]):
                if cb in self._exact_boxes:
                    self._exact_boxes.remove(cb)
            rec["w"].setParent(None); rec["w"].deleteLater()
        except Exception:
            pass
        self._fam_links[key].setEnabled(True)
        self._save(); self._fit()

    # ── Alternate-name rows (up to 3, ✕ to remove — like the site) ─────────── #
    def _add_alt_row(self, first="", last="", first_exact=False, last_exact=False):
        if len(self._alt) >= ALT_MAX:
            return None
        row_w = QWidget(); row_w.setObjectName("famCard")
        rl = QHBoxLayout(row_w); rl.setContentsMargins(8, 6, 8, 6); rl.setSpacing(6)
        lbl = QLabel("Alt Name:"); lbl.setFixedWidth(95)
        ff = QLineEdit(); ff.setPlaceholderText("Alternate First Name"); ff.setText(first)
        cbf = QCheckBox("Exact"); cbf.setChecked(bool(first_exact))
        lf = QLineEdit(); lf.setPlaceholderText("Alternate Last Name"); lf.setText(last)
        cbl = QCheckBox("Exact"); cbl.setChecked(bool(last_exact))
        rm = QPushButton("✕"); rm.setObjectName("rmBtn"); rm.setFixedWidth(26)
        rl.addWidget(lbl); rl.addWidget(ff, 2); rl.addWidget(cbf)
        rl.addWidget(lf, 2); rl.addWidget(cbl); rl.addWidget(rm)
        rec = {"w": row_w, "first": ff, "last": lf, "fe": cbf, "le": cbl}
        self._alt.append(rec)
        self._alt_box.addWidget(row_w)
        for cb in (cbf, cbl):
            self._exact_boxes.append(cb)
            cb.setVisible(self.f_show_exact.isChecked())
        ff.textChanged.connect(self._save); lf.textChanged.connect(self._save)
        cbf.stateChanged.connect(self._save); cbl.stateChanged.connect(self._save)
        rm.clicked.connect(lambda _=0: self._remove_alt_row(rec))
        if len(self._alt) >= ALT_MAX:
            self._alt_link.setEnabled(False)
        self._save(); self._fit()
        return rec

    def _remove_alt_row(self, rec):
        try:
            self._alt.remove(rec)
            for cb in (rec["fe"], rec["le"]):
                if cb in self._exact_boxes:
                    self._exact_boxes.remove(cb)
            rec["w"].setParent(None); rec["w"].deleteLater()
        except Exception:
            pass
        self._alt_link.setEnabled(True)
        self._save(); self._fit()

    def _reload_states(self, *_):
        country = self.f_country.currentText().strip()
        cur = self.f_state.currentText()
        self.f_state.blockSignals(True)
        self.f_state.clear(); self.f_state.addItem("")
        for st in (FS_LOCATIONS.get(country) or []):
            self.f_state.addItem(st)
        i = self.f_state.findText(cur)
        self.f_state.setCurrentIndex(i if i >= 0 else 0)
        self.f_state.blockSignals(False)
        self._save()

    def _toggle_exact(self, on: bool):
        for cb in self._exact_boxes:
            cb.setVisible(bool(on))
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
              self.f_first, self.f_last,
              self.f_first_exact, self.f_last_exact,
              self.f_tab, self.f_sex,
              self.f_country, self.f_state, self.f_keywords, self.f_show_exact,
              self.f_batch, self.f_film, self.f_principal,
              self.f_folder, self.f_docx, self.f_xlsx]
        ws += self.f_rectypes
        for pf, pcb, yf, rc, dcb in self._event_fields.values():
            ws += [pf, pcb, yf, rc, dcb]
        return ws

    # ── Autosave / load ───────────────────────────────────────────────────── #
    def _save(self, *_):
        d = {
            "username": self.f_user.text(), "password": self.f_pass.text(),
            "first_names": self.f_first.text(), "last_names": self.f_last.text(),
            "first_exact": self.f_first_exact.isChecked(),
            "last_exact": self.f_last_exact.isChecked(),
            "tab": self.f_tab.currentText(),
            "sex": self.f_sex.currentText(),
            "country": self.f_country.currentText(),
            "state": self.f_state.currentText(),
            "keywords": self.f_keywords.text(),
            "show_exact": self.f_show_exact.isChecked(),
            "batch": self.f_batch.text(), "film": self.f_film.text(),
            "principal": self.f_principal.isChecked(),
            "rectypes": [i for i, cb in enumerate(self.f_rectypes) if cb.isChecked()],
            "output_folder": self.f_folder.text(),
            "fmt_docx": self.f_docx.isChecked(), "fmt_xlsx": self.f_xlsx.isChecked(),
            "adv_open": self._adv_btn.isChecked(),
        }
        for key, (pf, pcb, yf, rc, dcb) in self._event_fields.items():
            d[f"{key}_place"] = pf.text()
            d[f"{key}_place_exact"] = pcb.isChecked()
            d[f"{key}_year"]  = yf.text()
            d[f"{key}_range_i"] = rc.currentIndex()
            d[f"{key}_date_exact"] = dcb.isChecked()
        d["family"] = {k: [[r["first"].text(), r["last"].text(),
                            r["fe"].isChecked(), r["le"].isChecked()]
                           for r in rows] for k, rows in self._fam.items()}
        d["alt_names"] = [[r["first"].text(), r["last"].text(),
                           r["fe"].isChecked(), r["le"].isChecked()]
                          for r in self._alt]
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
        _s(self.f_first_exact, "first_exact"); _s(self.f_last_exact,  "last_exact")
        _s(self.f_tab,   "tab")
        _s(self.f_sex,   "sex")
        _s(self.f_country, "country"); self._reload_states(); _s(self.f_state, "state")
        _s(self.f_keywords, "keywords"); _s(self.f_show_exact, "show_exact")
        _s(self.f_batch, "batch"); _s(self.f_film, "film")
        _s(self.f_principal, "principal")
        for i in (d.get("rectypes") or []):
            if 0 <= i < len(self.f_rectypes): self.f_rectypes[i].setChecked(True)
        _s(self.f_folder, "output_folder")
        _s(self.f_docx,  "fmt_docx"); _s(self.f_xlsx, "fmt_xlsx")
        for key, (pf, pcb, yf, rc, dcb) in self._event_fields.items():
            _s(pf, f"{key}_place"); _s(pcb, f"{key}_place_exact")
            _s(yf, f"{key}_year"); _s(dcb, f"{key}_date_exact")
            if isinstance(d.get(f"{key}_range_i"), int):
                rc.setCurrentIndex(d[f"{key}_range_i"])
        for key, rows in (d.get("family") or {}).items():
            if key in self._fam:
                for r in rows:
                    r = (list(r) + ["", "", False, False])[:4]
                    self._add_fam_row(key, r[0], r[1], bool(r[2]), bool(r[3]))
        for r in (d.get("alt_names") or []):
            r = (list(r) + ["", "", False, False])[:4]
            self._add_alt_row(r[0], r[1], bool(r[2]), bool(r[3]))
        self._toggle_exact(self.f_show_exact.isChecked())
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
        """Collect all advanced search field values for the URL builder."""
        adv = {
            "sex":        self.f_sex.currentText(),
            "country":    self.f_country.currentText().strip(),
            "state":      self.f_state.currentText().strip(),
            "keywords":   self.f_keywords.text().strip(),
            "batch":      self.f_batch.text().strip(),
            "film":       self.f_film.text().strip(),
            "principal":  self.f_principal.isChecked(),
            "record_types": [i for i, cb in enumerate(self.f_rectypes)
                             if cb.isChecked()],
            "show_exact": self.f_show_exact.isChecked(),
        }
        # life events (incl. Birth): {key:{place,place_exact,year,range,date_exact}}
        events = {}
        for key, (pf, pcb, yf, rc, dcb) in self._event_fields.items():
            pl, yr = pf.text().strip(), yf.text().strip()
            if pl or yr:
                events[key] = {"place": pl, "place_exact": pcb.isChecked(),
                               "year": yr,
                               "range": YEAR_RANGE[rc.currentIndex()][1],
                               "date_exact": dcb.isChecked()}
        adv["events"] = events
        # family members: {key: [{first,last,first_exact,last_exact}, …]}
        fam = {}
        for key, rows in self._fam.items():
            people = [{"first": r["first"].text().strip(),
                       "last":  r["last"].text().strip(),
                       "first_exact": r["fe"].isChecked(),
                       "last_exact":  r["le"].isChecked()}
                      for r in rows
                      if r["first"].text().strip() or r["last"].text().strip()]
            if people:
                fam[key] = people
        adv["family"] = fam
        # alternate names: [{first,last,first_exact,last_exact}, …] (up to 3)
        alts = [{"first": r["first"].text().strip(),
                 "last":  r["last"].text().strip(),
                 "first_exact": r["fe"].isChecked(),
                 "last_exact":  r["le"].isChecked()}
                for r in self._alt
                if r["first"].text().strip() or r["last"].text().strip()]
        adv["alt_names"] = alts
        has = any([adv["sex"] != "Unspecified", alts,
                   adv["country"], adv["state"], adv["keywords"], adv["batch"],
                   adv["film"], adv["principal"], adv["record_types"],
                   events, fam])
        return adv if has else {}

    def _payload(self) -> dict:
        return {
            "first_names":   self.f_first.text().strip(),
            "last_names":    self.f_last.text().strip(),
            "place_lived":   "",     # place / birth are now Life Events
            "birth_year":    "",
            "year_range":    0,
            "tab":           self.f_tab.currentText(),
            "advanced":      self._build_advanced(),
            "exact": {
                "name":    self.f_first_exact.isChecked(),
                "surname": self.f_last_exact.isChecked(),
            },
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "email":         self.f_user.text().strip() or None,
            "password":      self.f_pass.text() or None,
            "log":           print,
            "cancel_event":  getattr(self, "_cancel_ev", None),
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

    # ── File-conflict dialog (existing output files) ──────────────────────── #
    def _show_file_conflict_dialog(self, names: str):
        """Existing Word/Excel files were found — ask what to do (always)."""
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
    w = FamilySearchApp()
    w.show()
    sys.exit(app.exec())
