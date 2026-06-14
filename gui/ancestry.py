"""
gui/ancestry.py
================
Ancestry.com search window. English GUI, green Ancestry theme. Mirrors the live
advanced search (which is MyHeritage-like):

- Basic: First Names / Last Names / Place — each with its own "Exact" checkbox
  (default OFF); Birth Year with a ± dropdown (Exact / ±1 / ±2 / ±5 / ±10).
- Advanced: family members (Father / Mother / Spouse / Sibling / Child) — add as
  many as you like with the "+ Add" buttons, each with First/Last + Exact;
  Keyword; Gender (Male/Female); Race/Nationality (text); Collection Focus
  dropdown; and the four result-type filter checkboxes.

Login required (persistent profile keeps the session). Logo: Ancestry.png.
Autosave: .ancestry_autosave.json
"""

import json, re, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QComboBox, QToolButton,
    QFileDialog, QProgressBar, QMessageBox,
    QApplication, QGroupBox, QFrame, QGridLayout, QScrollArea,
)
from gui._app_icon import app_icon, make_header, make_cancel_button
from PySide6.QtCore import QThread, Signal, Qt, QByteArray, QEvent
from PySide6.QtGui import QIcon, QStandardItem, QStandardItemModel

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

# Birth-year ± dropdown → year_range (birth_x=<N>-0-0)
YEAR_OPTIONS = [("Exact (this year)", 0), ("± 1 year", 1), ("± 2 years", 2),
                ("± 5 years", 5), ("± 10 years", 10)]
GENDER_OPTIONS = ["—", "Male", "Female"]
COLLECTION_OPTIONS = ["All Collections", "USA", "UK & Ireland", "Europe",
                      "Canada", "Australia & New Zealand", "Jewish Family History"]
# Name exactness — the site's slider levels (first name 5, surname 4). Last item
# "Exact" maps to name_x=1; everything else is broad.
FIRST_EXACT_OPTIONS   = ["Broad", "Exact + similar + sounds + initials",
                         "Exact + sounds + similar", "Exact + similar", "Exact"]
SURNAME_EXACT_OPTIONS = ["Broad", "Exact + sounds + similar",
                         "Exact + similar", "Exact"]

# Result filters (left panel) — loaded from the per-section JSONs in config/
# (single source of truth; edit those to add/correct filters). Record-type & date
# selections become /search/categories/<code>/ passes; the GUI carries each code
# with its selection (collision-proof for the "1800s" century-vs-decade label).
_CFG_DIR = _ROOT / "config"

def _load_json(name, default):
    try:
        return json.loads((_CFG_DIR / name).read_text("utf-8"))
    except Exception:
        return default

# The filters are an EXPANDABLE TREE, like the site (arrow opens each layer, a
# dropdown only at the deepest leaf). Each section is converted to a uniform
# spec tuple (label, code, children, leaf):
#   • children = list of child spec tuples (recurse) — None at a leaf
#   • leaf     = {name: code} → a checkable dropdown at the end (collections /
#                states / decades). None when the node has children instead.
# Sources (config/, single source of truth — edit to add/correct filters):
#   record types: category → subcategory → collections   (Census → decade → colls)
#   locations:    continent → country → states            (each a record_f code)
#   Record Date:  century → decades                       (separate bottom facet)
_RT  = _load_json("ancestry_record_types.json", {})
_RD  = _load_json("ancestry_record_dates.json", {})
LOCATIONS = _load_json("ancestry_locations.json", {})

def _num_key(s):                       # numeric-aware sort (1700s < 1800s < …)
    m = re.match(r"(\d+)", str(s))
    return (0, int(m.group(1))) if m else (1, str(s).lower())

# continents in the site's order (most records first); countries A→Z inside
CONTINENT_ORDER = ["North America", "Europe", "Asia", "South America",
                   "Oceania", "Africa", "Antarctica"]

def _rt_spec():
    # categories in document (site) order; census decades chronological, other
    # children kept in their JSON order
    def node(label, d):
        kids = d.get("children")
        if kids:
            keys = (sorted(kids, key=_num_key) if d.get("code") == "35"
                    else list(kids))
            return (label, d.get("code"), [node(k, kids[k]) for k in keys], None)
        return (label, d.get("code"), None, d.get("collections") or None)
    return [node(c, _RT[c]) for c in _RT]

def _loc_spec():
    # continent (site order) → country (A→Z) → state (A→Z) → places (A→Z).
    # NO dropdowns: places are checkbox children opened by an arrow.
    out = []
    for cont in (CONTINENT_ORDER + [c for c in LOCATIONS if c not in CONTINENT_ORDER]):
        if cont not in LOCATIONS:
            continue
        cd = LOCATIONS[cont]
        countries = []
        for ctry in sorted(cd.get("countries", {})):
            ccd = cd["countries"][ctry]
            states = []
            for st in sorted(ccd.get("states", {})):
                sd = ccd["states"][st]
                places = [(p, pc, None, None)
                          for p, pc in sorted((sd.get("places") or {}).items())]
                states.append((st, sd.get("code"), places or None, None))
            countries.append((ctry, ccd.get("code"), states or None, None))
        out.append((cont, cd.get("code"), countries, None))
    return out

def _rd_spec():
    # century → decades. NO dropdown: decades are checkbox children (arrow opens).
    out = []
    for c in sorted(_RD, key=_num_key):
        decs = sorted((_RD[c].get("decades") or {}).items(),
                      key=lambda kv: _num_key(kv[0]))
        children = [(d, code, None, None) for d, code in decs]
        out.append((c, _RD[c].get("code"), children or None, None))
    return out
FAMILY_TYPES = [("Father", "father"), ("Mother", "mother"),
                ("Sibling", "sibling"), ("Spouse", "spouse"), ("Child", "child")]
_FAM_LABEL = {k: l for l, k in FAMILY_TYPES}   # key → display label
_FAM_SINGLE = {"father", "mother"}             # only ONE of each allowed
# Add event: links → URL key. Birth stays the basic Birth Year up top.
EVENT_TYPES = [("Marriage", "marriage"), ("Death", "death"),
               ("Lived In", "residence"), ("Any Event", "any")]
_EVENT_LABEL = {k: l for l, k in EVENT_TYPES}

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
QPushButton#addBtn{color:#4a6b1f;background:transparent;border:none;
  font-weight:bold;text-align:left;padding:1px 0;}
QPushButton#addBtn:hover{color:#6B8E23;}
QPushButton#rmBtn{color:#a33;background:transparent;border:none;font-weight:bold;
  max-width:24px;padding:0;}
QPushButton#eyeBtn{border:none;background:transparent;padding:0;}
QPushButton#advBtn{text-align:left;border:none;background:transparent;
  color:#4a6b1f;font-weight:bold;font-size:11px;padding:2px 0;}
QPushButton#advBtn:hover{color:#6B8E23;}
QProgressBar{border:1px solid #b7cfa0;border-radius:4px;
  text-align:center;min-height:18px;}
QProgressBar::chunk{background:#6B8E23;border-radius:3px;}
QLabel#sechead{font-weight:bold;color:#4a6b1f;font-size:10px;margin-top:4px;}
QFrame#div{background:#cfe0b5;}
QToolButton#treeTog{border:none;background:transparent;color:#4a6b1f;
  font-weight:bold;font-size:11px;padding:0;}
QToolButton#treeTog:hover{color:#6B8E23;}
"""


class CheckableComboBox(QComboBox):
    """A dropdown whose items each have a checkbox (multi-select). No "any" entry —
    nothing checked = nothing selected. Each item carries its own code. The popup
    stays open while ticking; the line shows the checked labels."""
    changed = Signal()

    def __init__(self, placeholder="— none —"):
        super().__init__()
        self._ph = placeholder
        self.setModel(QStandardItemModel(self))
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self.lineEdit().setPlaceholderText(placeholder)
        self.setInsertPolicy(QComboBox.NoInsert)
        self._codes = {}
        self.view().viewport().installEventFilter(self)
        self.model().dataChanged.connect(lambda *a: (self._refresh(),
                                                      self.changed.emit()))

    def add_item(self, label, code=None):
        it = QStandardItem(label)
        it.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        it.setData(Qt.Unchecked, Qt.CheckStateRole)
        self.model().appendRow(it)
        self._codes[label] = code

    def add_items(self, pairs):
        for label, code in pairs:
            self.add_item(label, code)

    def clear_items(self):
        self.model().clear()
        self._codes = {}
        self._refresh()

    def checked(self) -> list:
        out = []
        for i in range(self.model().rowCount()):
            it = self.model().item(i)
            if it.checkState() == Qt.Checked:
                out.append((it.text(), self._codes.get(it.text())))
        return out

    def set_checked(self, labels):
        s = set(labels or [])
        for i in range(self.model().rowCount()):
            it = self.model().item(i)
            it.setCheckState(Qt.Checked if it.text() in s else Qt.Unchecked)

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.MouseButtonRelease and obj is self.view().viewport():
            idx = self.view().indexAt(ev.pos())
            if idx.isValid():
                it = self.model().itemFromIndex(idx)
                it.setCheckState(Qt.Unchecked if it.checkState() == Qt.Checked
                                 else Qt.Checked)
            return True                     # consume → popup stays open
        return super().eventFilter(obj, ev)

    def _refresh(self):
        labels = [t for t, _c in self.checked()]
        le = self.lineEdit()
        le.setText(", ".join(labels))
        le.setToolTip("\n".join(labels))


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
        self.setMinimumWidth(920)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._fam: dict = {k: [] for _l, k in FAMILY_TYPES}   # rel → [row dicts]
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
        self.f_user = QLineEdit(); self.f_user.setPlaceholderText("Ancestry email or username")
        self.f_pass = PwdEdit(); self.f_pass.setPlaceholderText("Password")
        cl.addWidget(QLabel("Username:")); cl.addWidget(self.f_user, 2)
        cl.addWidget(QLabel("Password:")); cl.addWidget(self.f_pass, 2)
        self._outer.addWidget(cg)

        # Basic Search
        bg = QGroupBox("Basic Search")
        bf = QGridLayout(bg); bf.setSpacing(8)
        self.f_first = QLineEdit(); self.f_first.setPlaceholderText("First & Middle Name(s)")
        self.f_last  = QLineEdit(); self.f_last.setPlaceholderText("Last Name")
        self.f_place = QLineEdit(); self.f_place.setPlaceholderText("City, county, state, country")
        self.f_byear = QLineEdit(); self.f_byear.setPlaceholderText("e.g. 1897")
        self.f_byear.setFixedWidth(90)
        self.f_year_range = QComboBox()
        for lbl, _v in YEAR_OPTIONS:
            self.f_year_range.addItem(lbl)
        self.f_year_range.setCurrentIndex(1)        # ± 1 year (default)
        self.f_first_exact = QComboBox(); self.f_first_exact.addItems(FIRST_EXACT_OPTIONS)
        self.f_first_exact.setToolTip("Name match exactness (like the site's slider)")
        self.f_first_exact.setMaximumWidth(150)
        self.f_last_exact  = QComboBox(); self.f_last_exact.addItems(SURNAME_EXACT_OPTIONS)
        self.f_last_exact.setToolTip("Surname match exactness")
        self.f_last_exact.setMaximumWidth(150)
        self.f_place_exact = QCheckBox("Exact")
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
        yr = QHBoxLayout(); yr.setSpacing(4)
        yr.addWidget(self.f_byear); yr.addWidget(self.f_year_range, 1)
        bf.addLayout(yr, 1, 4, 1, 2)
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

        # Events — "Add event: Marriage Death Lived In Any Event" (link row)
        ev_head = QHBoxLayout(); ev_head.setSpacing(10)
        ev_head.addWidget(QLabel("Add event:"))
        for label, etype in EVENT_TYPES:
            lk = QPushButton(label); lk.setObjectName("addBtn")
            lk.clicked.connect(lambda _=False, t=etype: (self._add_event_row(t), self._save()))
            ev_head.addWidget(lk)
        ev_head.addStretch()
        av.addLayout(ev_head)
        self._events: list = []
        self._events_box = QVBoxLayout(); self._events_box.setSpacing(3)
        av.addLayout(self._events_box)
        av.addWidget(_divider())

        # Family — "Add family member: Father Mother Sibling Spouse Child" (link row,
        # exactly like the site). Father/Mother single; Sibling/Spouse/Child unlimited.
        fam_head = QHBoxLayout(); fam_head.setSpacing(10)
        fam_head.addWidget(QLabel("Add family member:"))
        self._fam_links = {}
        for label, key in FAMILY_TYPES:
            lk = QPushButton(label); lk.setObjectName("addBtn")
            lk.clicked.connect(lambda _=False, k=key: (self._add_fam_row(k), self._save()))
            self._fam_links[key] = lk
            fam_head.addWidget(lk)
        fam_head.addStretch()
        av.addLayout(fam_head)
        self._fam_box = QVBoxLayout(); self._fam_box.setSpacing(3)
        self._fam_containers = {}
        for _label, key in FAMILY_TYPES:
            box = QVBoxLayout(); box.setSpacing(3)
            self._fam_containers[key] = box
            self._fam_box.addLayout(box)
        av.addLayout(self._fam_box)
        av.addWidget(_divider())

        # Keyword / Gender / Race
        gr = QGridLayout(); gr.setSpacing(6)
        self.f_keyword = QLineEdit()
        self.f_keyword.setPlaceholderText("Occupation, street address, etc.")
        self.f_gender = QComboBox(); self.f_gender.addItems(GENDER_OPTIONS)
        self.f_race = QLineEdit(); self.f_race.setPlaceholderText("Race / Nationality")
        gr.addWidget(QLabel("Keyword:"), 0, 0); gr.addWidget(self.f_keyword, 0, 1, 1, 3)
        gr.addWidget(QLabel("Gender:"),  1, 0); gr.addWidget(self.f_gender, 1, 1)
        gr.addWidget(QLabel("Race/Nat.:"), 1, 2); gr.addWidget(self.f_race, 1, 3)
        gr.setColumnStretch(1, 1); gr.setColumnStretch(3, 1)
        av.addLayout(gr)
        av.addWidget(_divider())

        # Collection focus + result-type filters
        cr = QHBoxLayout(); cr.setSpacing(8)
        self.f_collection = QComboBox(); self.f_collection.addItems(COLLECTION_OPTIONS)
        cr.addWidget(QLabel("Collection Focus:")); cr.addWidget(self.f_collection)
        cr.addStretch()
        av.addLayout(cr)
        fr2 = QGridLayout(); fr2.setSpacing(4)
        self.f_hist   = QCheckBox("Historical Records"); self.f_hist.setChecked(True)
        self.f_trees  = QCheckBox("Family Trees");       self.f_trees.setChecked(True)
        self.f_stories= QCheckBox("Stories & Publications"); self.f_stories.setChecked(True)
        self.f_photos = QCheckBox("Photos & Maps");      self.f_photos.setChecked(True)
        fr2.addWidget(self.f_hist, 0, 0);   fr2.addWidget(self.f_trees, 0, 1)
        fr2.addWidget(self.f_stories, 1, 0); fr2.addWidget(self.f_photos, 1, 1)
        av.addLayout(fr2)

        self._adv_scroll = QScrollArea()
        self._adv_scroll.setWidget(self._adv)
        self._adv_scroll.setWidgetResizable(True)
        self._adv_scroll.setFrameShape(QFrame.NoFrame)
        self._adv_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._adv_scroll.setVisible(False)
        self._outer.addWidget(self._adv_scroll)

        # ── Result filters (left-panel; applied AFTER the search) ──────────── #
        self._flt_btn = QPushButton("▶   Filters  (Record type / Location / Date — multi-pass)")
        self._flt_btn.setObjectName("advBtn")
        self._flt_btn.setCheckable(True)
        self._flt_btn.toggled.connect(self._toggle_flt)
        self._outer.addWidget(self._flt_btn)

        self._flt = QGroupBox()
        fv = QVBoxLayout(self._flt); fv.setSpacing(4)
        hdr = QHBoxLayout()
        hdr.addWidget(QLabel(
            "Click ▶ to open each layer; tick any value (each becomes its own pass "
            "/ document). Location is combined (AND) with each."))
        hdr.addStretch()
        self._clear_flt_btn = QPushButton("Clear all filters")
        self._clear_flt_btn.setObjectName("addBtn")
        self._clear_flt_btn.clicked.connect(self._clear_filters)
        hdr.addWidget(self._clear_flt_btn)
        fv.addLayout(hdr)
        # each section keeps its ticked-node checkboxes, leaf dropdowns, and a
        # path→widgets registry (so saved selections can be rebuilt lazily)
        self._rt_checks,  self._rt_combos,  self._rt_reg  = [], [], {}
        self._loc_checks, self._loc_combos, self._loc_reg = [], [], {}
        self._rd_checks,  self._rd_combos,  self._rd_reg  = [], [], {}

        fv.addWidget(_sechead("RECORD TYPE"))
        for spec in _rt_spec():
            fv.addLayout(self._tree_node(spec, self._rt_checks, self._rt_combos,
                                         self._rt_reg))
        fv.addWidget(_divider())
        fv.addWidget(_sechead("RECORD LOCATION"))
        for spec in _loc_spec():
            fv.addLayout(self._tree_node(spec, self._loc_checks, self._loc_combos,
                                         self._loc_reg))
        fv.addWidget(_divider())
        fv.addWidget(_sechead("RECORD DATE"))
        fv.addLayout(self._rd_columns(_rd_spec(), self._rd_checks, self._rd_combos,
                                      self._rd_reg))

        self._flt_scroll = QScrollArea()
        self._flt_scroll.setWidget(self._flt)
        self._flt_scroll.setWidgetResizable(True)
        self._flt_scroll.setFrameShape(QFrame.NoFrame)
        self._flt_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._flt_scroll.setVisible(False)
        self._outer.addWidget(self._flt_scroll)

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
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(
            QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        for w in self._static_widgets():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QComboBox): w.currentTextChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)
        # filter checkboxes/combos are wired to _save inside their builders
        self._fit()

    def _toggle_flt(self, on: bool):
        self._flt_scroll.setVisible(on)
        if on:
            sh = QApplication.primaryScreen().availableGeometry().height()
            self._flt_scroll.setMaximumHeight(int(sh * 0.5))
        self._flt_btn.setText(("▼" if on else "▶") +
                              "   Filters  (Record type / Location / Date — multi-pass)")
        self._fit()

    def _tree_node(self, spec, checks, combos, reg, path=(), depth=0) -> QVBoxLayout:
        """One expandable tree node. Header = [arrow | aligned spacer] + [checkbox].
        Body is built LAZILY on first expand (fast startup). `reg` maps a node's
        path (tuple of labels) → {cb, open, combo} so saved selections can be
        re-opened and re-checked across the lazy tree."""
        label, code, children, leaf = spec
        path2 = path + (label,)
        box = QVBoxLayout(); box.setSpacing(1); box.setContentsMargins(0, 0, 0, 0)
        head = QHBoxLayout(); head.setSpacing(3)
        head.setContentsMargins(depth * 16, 0, 0, 0)
        has_body = bool(children or leaf)
        tog = QToolButton(); tog.setObjectName("treeTog"); tog.setFixedWidth(16)
        tog.setText("▶" if has_body else "")
        tog.setEnabled(has_body); head.addWidget(tog)      # fixed col → aligned cb
        cb = QCheckBox(label.replace("&", "&&")); cb.stateChanged.connect(self._save)
        checks.append((cb, label, code)); head.addWidget(cb); head.addStretch()
        box.addLayout(head)
        entry = reg.setdefault(path2, {}); entry["cb"] = cb
        if has_body:
            body = QWidget(); bl = QVBoxLayout(body)
            bl.setContentsMargins(0, 0, 0, 0); bl.setSpacing(1)
            body.setVisible(False); box.addWidget(body)
            st = {"built": False}

            def build_once(sp=spec, lay=bl, dp=depth, p=path2):
                if not st["built"]:
                    self._build_body(lay, sp, checks, combos, reg, p, dp)
                    st["built"] = True

            def toggle(_=0, b=body, t=tog):
                build_once(); vis = not b.isVisible(); b.setVisible(vis)
                t.setText("▼" if vis else "▶")

            def open_node(b=body, t=tog):       # for restore: build + reveal
                build_once(); b.setVisible(True); t.setText("▼")
            tog.clicked.connect(toggle)
            entry["open"] = open_node
        return box

    def _build_body(self, bl, spec, checks, combos, reg, path, depth):
        """Build a node's children/leaf on demand (lazy)."""
        label, code, children, leaf = spec
        if code == "35" and children:           # Census → THREE century columns
            self._census_columns(bl, children, checks, combos, reg, path, depth)
            return
        for ch in (children or []):
            bl.addLayout(self._tree_node(ch, checks, combos, reg, path, depth + 1))
        if leaf:
            combo = CheckableComboBox()
            combo.add_items([(n, leaf[n]) for n in sorted(leaf)])
            combo.changed.connect(self._save); combos.append(combo)
            combo.setMinimumWidth(320)
            combo.setMaximumWidth(480)        # keep the arrow inside the window
            reg.setdefault(path, {})["combo"] = combo
            w = QHBoxLayout(); w.addSpacing((depth + 1) * 16)
            w.addWidget(combo); w.addStretch(); bl.addLayout(w)

    def _census_columns(self, bl, decade_specs, checks, combos, reg, path, depth):
        """Census decades as three columns by century (1700s/1800s/1900s), the
        whole block shifted right (years = 2nd level). Each century is a CHECKBOX
        node (cen_century1700/1800/1900 — select the whole century), expandable to
        its decades."""
        cols = {}
        for ch in decade_specs:                 # ch label like "1890s"
            cent = ch[0][:2] + "00s"
            cols.setdefault(cent, []).append(ch)
        row = QHBoxLayout(); row.setSpacing(20)
        row.setContentsMargins(24, 0, 0, 0)
        for cent in sorted(cols, key=_num_key):
            colw = QWidget(); col = QVBoxLayout(colw)
            col.setContentsMargins(0, 0, 0, 0); col.setSpacing(1)
            century = (cent, "cen_century" + cent[:4], cols[cent], None)
            col.addLayout(self._tree_node(century, checks, combos, reg, path, 0))
            col.addStretch()
            row.addWidget(colw, 0, Qt.AlignTop)
        row.addStretch()
        bl.addLayout(row)

    def _rd_columns(self, century_specs, checks, combos, reg) -> QHBoxLayout:
        """Record Date in THREE columns (no dropdowns): 1500s+1600s | 1700s+1800s |
        1900s+2000s. Each century opens by arrow into its decades (checkboxes)."""
        cols = {0: [], 1: [], 2: []}
        for spec in century_specs:
            c = int(re.match(r"\d+", spec[0]).group())
            cols[min((c - 1500) // 200, 2)].append(spec)
        row = QHBoxLayout(); row.setSpacing(24)
        for ci in (0, 1, 2):
            colw = QWidget(); col = QVBoxLayout(colw)
            col.setContentsMargins(0, 0, 0, 0); col.setSpacing(1)
            for spec in cols[ci]:
                col.addLayout(self._tree_node(spec, checks, combos, reg))
            col.addStretch()
            row.addWidget(colw, 0, Qt.AlignTop)
        row.addStretch()
        return row

    # ── filter persistence + clear ────────────────────────────────────────── #
    def _filter_state(self) -> dict:
        """Selected filters as paths (so they survive the lazy tree)."""
        out = {}
        for k, reg in (("rt", self._rt_reg), ("loc", self._loc_reg),
                       ("rd", self._rd_reg)):
            out[k + "_nodes"] = ["||".join(p) for p, e in reg.items()
                                 if e.get("cb") and e["cb"].isChecked()]
            cm = {}
            for p, e in reg.items():
                c = e.get("combo")
                if c:
                    ch = [l for l, _x in c.checked()]
                    if ch:
                        cm["||".join(p)] = ch
            out[k + "_combos"] = cm
        return out

    def _restore_filters(self, d):
        for k, reg in (("rt", self._rt_reg), ("loc", self._loc_reg),
                       ("rd", self._rd_reg)):
            for pj in (d.get(k + "_nodes") or []):
                path = tuple(pj.split("||"))
                self._open_to(reg, path[:-1])
                e = reg.get(path)
                if e and e.get("cb"):
                    e["cb"].setChecked(True)
            for pj, labels in (d.get(k + "_combos") or {}).items():
                path = tuple(pj.split("||"))
                self._open_to(reg, path)
                e = reg.get(path)
                if e and e.get("combo"):
                    e["combo"].set_checked(labels)

    @staticmethod
    def _open_to(reg, path):
        """Build + reveal every node along `path` so deep widgets exist."""
        for i in range(1, len(path) + 1):
            e = reg.get(path[:i])
            if e and e.get("open"):
                e["open"]()

    def _clear_filters(self):
        for reg in (self._rt_reg, self._loc_reg, self._rd_reg):
            for e in reg.values():
                if e.get("cb"):
                    e["cb"].setChecked(False)
                if e.get("combo"):
                    e["combo"].set_checked([])
        self._save()

    @staticmethod
    def _collect(checks, combos) -> list:
        """Selections in one section as [{label, code}] — every ticked node
        checkbox + every ticked leaf-dropdown value."""
        out = [{"label": l, "code": c} for cb, l, c in checks if cb.isChecked()]
        for combo in combos:
            for l, c in combo.checked():
                out.append({"label": l, "code": c})
        return out

    def _type_filters(self):
        return self._collect(self._rt_checks, self._rt_combos)

    def _location_filters(self):
        return self._collect(self._loc_checks, self._loc_combos)

    def _date_filters(self):
        return self._collect(self._rd_checks, self._rd_combos)

    # ── Dynamic family-member rows ────────────────────────────────────────── #
    def _add_fam_row(self, key, first="", last="", first_exact=False, last_exact=False):
        if key in _FAM_SINGLE and self._fam[key]:
            return None                       # only ONE father / mother
        box = self._fam_containers[key]
        row_w = QWidget()
        rl = QHBoxLayout(row_w); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(6)
        lbl = QLabel(_FAM_LABEL[key] + ":"); lbl.setFixedWidth(60)
        ff = QLineEdit(); ff.setPlaceholderText("First & Middle Name(s)"); ff.setText(first)
        cbf = QCheckBox("Exact"); cbf.setChecked(bool(first_exact))
        lf = QLineEdit(); lf.setPlaceholderText("Last Name"); lf.setText(last)
        cbl = QCheckBox("Exact"); cbl.setChecked(bool(last_exact))
        rm = QPushButton("✕"); rm.setObjectName("rmBtn"); rm.setFixedWidth(24)
        rl.addWidget(lbl); rl.addWidget(ff, 2); rl.addWidget(cbf)
        rl.addWidget(lf, 2); rl.addWidget(cbl); rl.addWidget(rm)
        rec = {"w": row_w, "first": ff, "last": lf,
               "first_exact": cbf, "last_exact": cbl}
        self._fam[key].append(rec)
        box.addWidget(row_w)
        ff.textChanged.connect(self._save); lf.textChanged.connect(self._save)
        cbf.stateChanged.connect(self._save); cbl.stateChanged.connect(self._save)
        rm.clicked.connect(lambda _=False: self._remove_fam_row(key, rec))
        if key in _FAM_SINGLE and key in getattr(self, "_fam_links", {}):
            self._fam_links[key].setEnabled(False)   # disable «Father/Mother» link
        self._fit()
        return rec

    def _remove_fam_row(self, key, rec):
        try:
            self._fam[key].remove(rec)
            rec["w"].setParent(None)
            rec["w"].deleteLater()
        except Exception:
            pass
        if key in _FAM_SINGLE and key in getattr(self, "_fam_links", {}):
            self._fam_links[key].setEnabled(True)
        self._save(); self._fit()

    # ── Dynamic event rows ────────────────────────────────────────────────── #
    def _add_event_row(self, etype, year="", rng=1, place=""):
        row_w = QWidget()
        rl = QHBoxLayout(row_w); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(6)
        lbl = QLabel(_EVENT_LABEL.get(etype, etype) + ":"); lbl.setFixedWidth(75)
        yf = QLineEdit(); yf.setPlaceholderText("Year"); yf.setFixedWidth(64); yf.setText(year)
        rngc = QComboBox()
        for lbl2, _v in YEAR_OPTIONS:
            rngc.addItem(lbl2)
        rngc.setCurrentIndex(next((i for i, (_l, v) in enumerate(YEAR_OPTIONS)
                                   if v == rng), 1))
        pf = QLineEdit(); pf.setPlaceholderText("City, County, State, Country"); pf.setText(place)
        rm = QPushButton("✕"); rm.setObjectName("rmBtn"); rm.setFixedWidth(24)
        rl.addWidget(lbl); rl.addWidget(yf); rl.addWidget(rngc)
        rl.addWidget(pf, 2); rl.addWidget(rm)
        rec = {"w": row_w, "type": etype, "year": yf, "range": rngc, "place": pf}
        self._events.append(rec)
        self._events_box.addWidget(row_w)
        yf.textChanged.connect(self._save); pf.textChanged.connect(self._save)
        rngc.currentTextChanged.connect(self._save)
        rm.clicked.connect(lambda _=False: self._remove_event_row(rec))
        self._fit()
        return rec

    def _remove_event_row(self, rec):
        try:
            self._events.remove(rec)
            rec["w"].setParent(None)
            rec["w"].deleteLater()
        except Exception:
            pass
        self._save(); self._fit()

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
        scr = QApplication.primaryScreen().availableGeometry()
        self.resize(self.width(), min(hint.height() + 42, scr.height() - 24))

    # ── Autosave ──────────────────────────────────────────────────────────── #
    def _static_widgets(self) -> list:
        return [self.f_user, self.f_pass, self.f_first, self.f_last, self.f_place,
                self.f_byear, self.f_year_range, self.f_first_exact,
                self.f_last_exact, self.f_place_exact, self.f_keyword,
                self.f_gender, self.f_race, self.f_collection,
                self.f_hist, self.f_trees, self.f_stories, self.f_photos,
                self.f_folder, self.f_docx, self.f_xlsx]

    def _save(self, *_):
        d = {
            "username": self.f_user.text(), "password": self.f_pass.text(),
            "first_names": self.f_first.text(), "last_names": self.f_last.text(),
            "place_lived": self.f_place.text(), "birth_year": self.f_byear.text(),
            "year_range_i": self.f_year_range.currentIndex(),
            "first_exact_i": self.f_first_exact.currentIndex(),
            "last_exact_i":  self.f_last_exact.currentIndex(),
            "place_exact": self.f_place_exact.isChecked(),
            "keyword": self.f_keyword.text(), "gender": self.f_gender.currentText(),
            "race": self.f_race.text(), "collection": self.f_collection.currentText(),
            "hist": self.f_hist.isChecked(), "trees": self.f_trees.isChecked(),
            "stories": self.f_stories.isChecked(), "photos": self.f_photos.isChecked(),
            "output_folder": self.f_folder.text(),
            "fmt_docx": self.f_docx.isChecked(), "fmt_xlsx": self.f_xlsx.isChecked(),
            "adv_open": self._adv_btn.isChecked(),
            "flt_open": self._flt_btn.isChecked(),
            "filters": self._filter_state(),     # full path-based filter selection
            "fam": {k: [[r["first"].text(), r["last"].text(),
                         r["first_exact"].isChecked(), r["last_exact"].isChecked()]
                        for r in rows] for k, rows in self._fam.items()},
            "events": [[e["type"], e["year"].text(),
                        YEAR_OPTIONS[e["range"].currentIndex()][1], e["place"].text()]
                       for e in self._events],
        }
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
        if isinstance(d.get("year_range_i"), int):
            self.f_year_range.setCurrentIndex(d["year_range_i"])
        if isinstance(d.get("first_exact_i"), int):
            self.f_first_exact.setCurrentIndex(d["first_exact_i"])
        if isinstance(d.get("last_exact_i"), int):
            self.f_last_exact.setCurrentIndex(d["last_exact_i"])
        _s(self.f_place_exact, "place_exact"); _s(self.f_keyword, "keyword")
        i = self.f_gender.findText(str(d.get("gender", "—")))
        if i >= 0: self.f_gender.setCurrentIndex(i)
        _s(self.f_race, "race")
        i = self.f_collection.findText(str(d.get("collection", "All Collections")))
        if i >= 0: self.f_collection.setCurrentIndex(i)
        _s(self.f_hist, "hist"); _s(self.f_trees, "trees")
        _s(self.f_stories, "stories"); _s(self.f_photos, "photos")
        _s(self.f_folder, "output_folder")
        _s(self.f_docx, "fmt_docx"); _s(self.f_xlsx, "fmt_xlsx")
        for key, rows in (d.get("fam") or {}).items():
            if key in self._fam:
                for r in rows:
                    self._add_fam_row(key, *(list(r) + ["", "", False, False])[:4])
        for e in (d.get("events") or []):
            e = (list(e) + ["", "", 1, ""])[:4]
            self._add_event_row(e[0], str(e[1]), int(e[2] or 1), str(e[3]))
        # restore the full filter selection (rebuilds the needed lazy branches)
        if d.get("filters"):
            self._restore_filters(d["filters"])
        if d.get("adv_open"):
            self._adv_btn.setChecked(True)
        if d.get("flt_open"):
            self._flt_btn.setChecked(True)

    # ── Helpers ───────────────────────────────────────────────────────────── #
    def _browse(self):
        p = QFileDialog.getExistingDirectory(
            self, "Select output folder", self.f_folder.text() or _DEF_DIR)
        if p: self.f_folder.setText(p)

    def _fmt(self) -> str:
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _build_advanced(self) -> dict:
        adv = {"keyword": self.f_keyword.text().strip(),
               "gender": self.f_gender.currentText().strip(),
               "race": self.f_race.text().strip(),
               "collection": self.f_collection.currentText().strip(),
               "filters": {"historical": self.f_hist.isChecked(),
                           "trees": self.f_trees.isChecked(),
                           "stories": self.f_stories.isChecked(),
                           "photos": self.f_photos.isChecked()}}
        for key, rows in self._fam.items():
            people = [{"first": r["first"].text().strip(),
                       "last": r["last"].text().strip(),
                       "first_exact": r["first_exact"].isChecked(),
                       "last_exact": r["last_exact"].isChecked()}
                      for r in rows
                      if r["first"].text().strip() or r["last"].text().strip()]
            if people:
                adv[key] = people
        adv["events"] = [
            {"type": e["type"], "year": e["year"].text().strip(),
             "range": YEAR_OPTIONS[e["range"].currentIndex()][1],
             "place": e["place"].text().strip()}
            for e in self._events
            if e["year"].text().strip() or e["place"].text().strip()]
        return adv

    def _payload(self) -> dict:
        return {
            "first_names": self.f_first.text().strip(),
            "last_names":  self.f_last.text().strip(),
            "place_lived": self.f_place.text().strip(),
            "birth_year":  self.f_byear.text().strip(),
            "year_range":  YEAR_OPTIONS[self.f_year_range.currentIndex()][1],
            "advanced":    self._build_advanced(),
            "exact": {
                "name":    self.f_first_exact.currentText() == "Exact",
                "surname": self.f_last_exact.currentText() == "Exact",
                "place":   self.f_place_exact.isChecked(),
                "name_level":    self.f_first_exact.currentText(),
                "surname_level": self.f_last_exact.currentText(),
            },
            "filters": {
                "types":     self._type_filters(),       # [{label, code}]
                "dates":     self._date_filters(),       # [{label, code}]
                "locations": self._location_filters(),   # [{label, code}] record_f
            },
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "email":    self.f_user.text().strip() or None,
            "password": self.f_pass.text() or None,
            "log":      print,
            "cancel_event": getattr(self, "_cancel_ev", None),
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

    def closeEvent(self, ev):
        # closing mid-run: cancel the scraper and let it wind down (no
        # "QThread destroyed while running" / Playwright noise in the log)
        w = getattr(self, "worker", None)
        if w is not None and w.isRunning():
            cev = getattr(self, "_cancel_ev", None)
            if cev is not None:
                cev.set()
            w.wait(8000)
        super().closeEvent(ev)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = AncestryApp()
    w.show()
    try:
        sys.exit(app.exec())
    except KeyboardInterrupt:
        sys.exit(0)
