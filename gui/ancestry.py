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
    QApplication, QGroupBox, QFrame, QGridLayout, QScrollArea, QTabWidget, QTabBar, QListWidget,
    QMenu, QToolButton, QRadioButton, QButtonGroup, QWidgetAction,
)
from gui._app_icon import app_icon, make_footer, make_header, make_cancel_button, autosave_path, clamp_on_screen
from PySide6.QtCore import QThread, Signal, Qt, QByteArray, QEvent, QTimer, QRegularExpression
from PySide6.QtGui import QIcon, QStandardItem, QStandardItemModel, QRegularExpressionValidator

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
_SAVE    = autosave_path(".ancestry_autosave.json")
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
# Name match FORMS — exactly like the site: «Exact and…» on top, the fuzzy forms
# indented under it (tick any combination). Code letters feed Ancestry's name_x:
# p=sounds-like, s=similar, i=initials, 1=exact only.
FIRST_FORMS   = [("Exact and…", "1"), ("    Sounds like", "p"),
                 ("    Similar", "s"), ("    Initials", "i")]
SURNAME_FORMS = [("Exact and…", "1"), ("    Sounds like", "p"), ("    Similar", "s")]

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

# Per-tab «Narrow by Category» links (from the user's Census.docx): tab title →
# [[display name, code], …]. Ticking one runs a pass drilled into that category.
NARROW_CATEGORIES = _load_json("ancestry_narrow_categories.json", {})

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

def _place_node(label, val):
    # a place is either a bare code (leaf) or {code, places:{…}} that nests to ANY
    # depth (county → city → locality …). Arrows open each level; no dropdowns.
    if isinstance(val, dict):
        kids = [_place_node(p, v) for p, v in sorted((val.get("places") or {}).items())]
        return (label, val.get("code"), kids or None, None)
    return (label, val, None, None)                  # leaf: val is the record_f code

def _loc_spec():
    # continent (site order) → country (A→Z) → state (A→Z) → places (A→Z, nested).
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
                places = [_place_node(p, v)
                          for p, v in sorted((sd.get("places") or {}).items())]
                states.append((st, sd.get("code"), places or None, None))
            countries.append((ctry, ccd.get("code"), states or None, None))
        out.append((cont, cd.get("code"), countries, None))
    return out

def _rd_spec():
    # century → decade → (year). NO dropdown: every level is a checkbox child that
    # an arrow opens. A decade value is either a bare code (old schema) or a dict
    # {code, years:{year:code}} (new schema from the filter crawler).
    out = []
    for c in sorted(_RD, key=_num_key):
        decs = sorted((_RD[c].get("decades") or {}).items(),
                      key=lambda kv: _num_key(kv[0]))
        children = []
        for d, dv in decs:
            if isinstance(dv, dict):
                yrs = sorted((dv.get("years") or {}).items(),
                             key=lambda kv: _num_key(kv[0]))
                ych = [(y, yc, None, None) for y, yc in yrs]
                children.append((d, dv.get("code"), ych or None, None))
            else:
                children.append((d, dv, None, None))      # old schema: code string
        out.append((c, _RD[c].get("code"), children or None, None))
    return out
FAMILY_TYPES = [("Father", "father"), ("Mother", "mother"),
                ("Sibling", "sibling"), ("Spouse", "spouse"), ("Child", "child")]
_FAM_LABEL = {k: l for l, k in FAMILY_TYPES}   # key → display label
_FAM_SINGLE = {"father", "mother"}             # only ONE of each allowed
# «Add event» — flat row, ALL of them (like the site). Each row = Day/Month/Year +
# Add Range + Location + «Exact to…».
EVENT_TYPES = [("Birth", "birth"), ("Marriage", "marriage"), ("Death", "death"),
               ("Lived In", "residence"), ("Any Event", "any"),
               ("Arrival", "arrival"), ("Departure", "departure"), ("Military", "military")]
_EVENT_LABEL = {k: l for l, k in EVENT_TYPES}
# Day/Month shown for these; Lived In / Any Event are year-only.
_EVENT_DM = {"birth": "full", "marriage": "full", "death": "full",
             "arrival": "full", "departure": "full", "military": "full",
             "residence": None, "any": None}
# Birth / Marriage / Death once; the rest up to 10.
_EVENT_CAP = {"birth": 1, "marriage": 1, "death": 1, "residence": 10, "any": 10,
              "arrival": 10, "departure": 10, "military": 10}

MONTHS = [("Month", "")] + [(m, str(i)) for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July", "August",
     "September", "October", "November", "December"], 1)]

# Category search TABS — shown IN ADDITION to the main Search tab, each mirroring
# the matching category page on ancestry.com. `code` → /search/categories/<code>/
# (or "collection:<id>" → /search/collections/<id>/). `sections` = the field blocks
# to render top-to-bottom, exactly like the live form. `family` = the relative rows
# that category shows. `narrow` = the sidebar "Narrow by Category" links (info only).
CATEGORY_SPECS = [
    {"title": "America 250", "code": "2026_us_a250",
     "sections": ["birth_full", "death_full", "lived_in", "marriage_full",
                  "arrival", "departure", "military", "any_event", "family",
                  "publication", "keyword", "gender_race", "residence_date",
                  "father_birthplace", "mother_birthplace"],
     "family": ["father", "mother", "sibling", "spouse", "child"]},
    {"title": "Census & Voter Lists", "code": "35",
     "sections": ["birth_year", "lived_in", "any_event", "family",
                  "keyword", "gender_race"],
     "family": ["father", "mother", "sibling", "spouse", "child"],
     "narrow": ["U.S. Federal Census Collection", "U.K. Census Collection",
                "Canadian Census Collection", "1700s Censuses",
                "1800s Censuses", "1900s Censuses"]},
    {"title": "Birth, Marriage & Death", "code": "34",
     "sections": ["birth_full", "death_full", "marriage_full", "any_event",
                  "family", "keyword", "gender"],
     "family": ["father", "mother", "spouse", "child"],
     "narrow": ["Birth, Baptism & Christening", "Marriage & Divorce",
                "Death, Burial, Cemetery & Obituaries"]},
    {"title": "Military", "code": "39",
     "sections": ["birth_full", "death_full", "lived_in", "military",
                  "any_event_month", "keyword", "race"],
     "narrow": ["Draft, Enlistment and Service", "Casualties",
                "Soldier, Veteran & Prisoner Rolls & Lists", "Pensions",
                "Histories", "Awards & Decorations of Honor", "News",
                "Disciplinary Actions", "Photos"]},
    {"title": "Immigration & Emigration", "code": "40",
     "sections": ["birth_year", "lived_in", "arrival", "departure",
                  "any_event", "origin", "keyword", "gender_race"],
     "narrow": ["Passenger Lists", "Citizenship & Naturalization",
                "Border Crossings & Passports", "Crew Lists",
                "Immigration & Emigration Books", "Ship & Port Pictures"]},
    {"title": "Directories & Member Lists", "code": "37",
     "sections": ["birth_full", "lived_in", "any_event", "family",
                  "keyword", "gender"],
     "family": ["spouse"]},
    {"title": "Public Member Trees", "code": "collection:1030",
     "sections": ["birth_full", "death_full", "marriage_full", "family", "keyword"],
     "family": ["father", "mother", "spouse", "child", "sibling"]},
]

# Collection-focus dropdown — EXACT site options (label, URL code). `None` code =
# group header (rendered disabled). Father/Mother are single; the rest up to 10.
COLLECTION_FOCUS = [
    ("All Collections", ""),
    ("— Country —", None),
    ("Australia", "australian"), ("Canada", "canada"), ("England", "england"),
    ("France", "france"), ("Germany", "german"), ("Ireland", "ireland"),
    ("Italy", "italy"), ("Mexico", "mexico"), ("Netherlands", "netherlands"),
    ("New Zealand", "new-zealand"), ("Norway", "norway"), ("Scotland", "scotland"),
    ("Sweden", "sweden"), ("United Kingdom", "united-kingdom"),
    ("United States", "usa"), ("Wales", "wales"),
    ("— Ethnicity —", None),
    ("African American", "african-american"), ("Jewish", "jewish"),
    ("Native American", "native-american"),
]

# "Included data collections" lists (category code → list, or a "ref" string to
# reuse another key). Loaded from config/ancestry_category_collections.json.
_CAT_COLLECTIONS = {}
try:
    _ccf = _ROOT / "config" / "ancestry_category_collections.json"
    if _ccf.exists():
        _CAT_COLLECTIONS = json.loads(_ccf.read_text(encoding="utf-8"))
except Exception:
    _CAT_COLLECTIONS = {}

class _StretchTabBar(QTabBar):
    """EVERY tab is the SAME width — the bar width split evenly across the tabs, floored
    at the widest label so no text is clipped. The width is IDENTICAL for every index and
    does NOT depend on which tab is selected, so switching tabs never reflows the bar
    (no "jumping tabs"). A small remainder gap on the right is left as-is — the window
    keeps its margin (no flush-to-edge over-fill)."""
    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.updateGeometry()                 # re-lay-out tabs to the new (settled) width

    def tabSizeHint(self, index):
        sz = super().tabSizeHint(index)
        n = self.count()
        if n <= 0:
            return sz
        naturals = [super(_StretchTabBar, self).tabSizeHint(i).width() for i in range(n)]
        widest = max(naturals) if naturals else sz.width()
        # Basis = the TOP-LEVEL WINDOW width (capped by the screen), NOT the bar's/parent's
        # own width. Reading the parent feeds back into its own size hint → tabs inflate
        # without bound. The window width is fixed by _fit/the screen, so no feedback loop.
        win = self.window()
        basis = win.width() if (win is not None and win.width() > 0) else self.width()
        avail = max(0, basis - 48)            # window minus side margins / scroll chrome
        # equal split, but never narrower than the widest label. Same value for ALL
        # indices → no per-tab variation → no jump on select.
        w = max(avail // n, widest)
        sz.setWidth(w)
        return sz


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
QPushButton#addBtn:disabled{color:#b4c2a4;}
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
QFrame#div{background:#cfe0b5;border:none;}
QToolButton#treeTog{border:none;background:transparent;color:#4a6b1f;
  font-weight:bold;font-size:11px;padding:0;}
QToolButton#treeTog:hover{color:#6B8E23;}
QTabWidget::pane{border:none;background:#f6faf0;top:0;}    /* no frame → tabs flush to window edge */
QWidget#tabPage{background:#f6faf0;}            /* ALL tabs the same light-green as the main */
QScrollArea{background:#f6faf0;border:none;}
QGroupBox#narrowBox{background:#eef5e2;}
QTabBar::tab{background:#e3ecd6;color:#5a7030;padding:5px 9px;margin-right:1px;
  border:1px solid #b7cfa0;border-bottom:none;
  border-top-left-radius:6px;border-top-right-radius:6px;}
QTabBar::tab:selected{background:#cfe0b5;color:#2f4612;}  /* NO bold → label width is selection-independent → tabs don't jump on switch */
QTabBar::tab:hover{background:#dcebc8;}
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

    def add_info(self, label):
        """A non-checkable info/link row (e.g. «About these settings») — shown but
        never toggled, never part of checked()."""
        it = QStandardItem(label)
        it.setFlags(Qt.ItemIsEnabled)            # NOT user-checkable
        it.setForeground(Qt.gray)
        self.model().appendRow(it)

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
                if it.isCheckable():        # skip non-checkable info/link rows
                    it.setCheckState(Qt.Unchecked if it.checkState() == Qt.Checked
                                     else Qt.Checked)
            return True                     # consume → popup stays open
        return super().eventFilter(obj, ev)

    def _refresh(self):
        labels = [t.strip() for t, _c in self.checked()]   # drop indent in the line
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
    f.setFrameShape(QFrame.NoFrame); f.setFixedHeight(2)   # single dark-green bar (no HLine)
    return f


def _sechead(text: str) -> QLabel:
    l = QLabel(text); l.setObjectName("sechead"); return l


# Canonical family-row order (the site order): Father, Mother, Spouse, Child, Sibling.
_FAM_ORDER = ["father", "mother", "spouse", "child", "sibling"]


# Place-exactness dropdown ("Exact to…" on the site → radios). label, URL code.
PLACE_EXACT = [("— any —", ""), ("This place", "1"), ("County", "PCO"),
               ("County and adjacent counties", "PACO"), ("State", "PS"),
               ("State and adjacent states", "PAS"), ("Country", "PC")]

# Section key → Ancestry URL key (from Alla's live result URLs).
_DATE_KEY = {"birth": "birth", "death": "death", "marriage": "marriage",
             "arrival": "arrival", "departure": "departure", "military": "military",
             "any": "event", "publication": "e--Publication",
             "residence": "residence", "residence_date": "residence",
             "origin": "e-Self-Origin", "father_birthplace": "e-Father-Birth",
             "mother_birthplace": "e-Mother-Birth"}


def _uv(s: str) -> str:
    """URL value: trim, spaces → «+» (Ancestry uses literal «+» in the path query)."""
    return (s or "").strip().replace(" ", "+")


def _mk_year_edit(width=84):
    """A Year input that only accepts up to 4 digits (no letters / no «1», «11» —
    those are caught on Search by bad_years())."""
    e = QLineEdit(); e.setPlaceholderText("Year"); e.setFixedWidth(width)
    e.setMaxLength(4)
    # keep digits only via a live filter — NO QValidator, so Qt never paints the field
    # red for an «intermediate» value (e.g. while «1870» is being typed). bad_years()
    # still rejects a non-4-digit year on Search (a dialog, not a red field).
    def _keep_digits(t, _e=e):
        d = "".join(ch for ch in t if ch.isdigit())[:4]
        if d != t:
            _e.blockSignals(True); _e.setText(d)
            _e.setCursorPosition(len(d)); _e.blockSignals(False)
    e.textEdited.connect(_keep_digits)
    return e


class _ExactTo(QToolButton):
    """The site's «Exact to…» place control: a button that opens a popup with a
    master checkbox + radios (This place / State & adjacent states / Country).
    code() → '' (off) or 1 / PAS / PC for the URL. Used on EVERY place/location."""
    changed = Signal()
    _OPTS = [("This place", "1"), ("County", "PCO"),
             ("County and adjacent counties", "PACO"), ("State", "PS"),
             ("State and adjacent states", "PAS"), ("Country", "PC")]

    def __init__(self):
        super().__init__()
        self.setObjectName("addBtn")
        self.setPopupMode(QToolButton.InstantPopup)
        self.setToolButtonStyle(Qt.ToolButtonTextOnly)
        menu = QMenu(self)
        cw = QWidget(); cl = QVBoxLayout(cw); cl.setContentsMargins(8, 6, 8, 6); cl.setSpacing(3)
        self._cb = QCheckBox("Exact to…"); cl.addWidget(self._cb)
        self._grp = QButtonGroup(self); self._radios = []
        for i, (lbl, code) in enumerate(self._OPTS):
            rb = QRadioButton(lbl); rb.setProperty("code", code)
            if i == 0:
                rb.setChecked(True)
            self._grp.addButton(rb); self._radios.append(rb)
            rr = QHBoxLayout(); rr.setContentsMargins(0, 0, 0, 0)
            rr.addSpacing(16); rr.addWidget(rb); cl.addLayout(rr)
        wa = QWidgetAction(menu); wa.setDefaultWidget(cw); menu.addAction(wa)
        self.setMenu(menu)
        self._cb.toggled.connect(self._refresh)
        for rb in self._radios:
            rb.toggled.connect(self._refresh)
        self._refresh()

    def _refresh(self, *_):
        if self._cb.isChecked():
            sel = next((r for r in self._radios if r.isChecked()), self._radios[0])
            self.setText("✓ Exact to: " + sel.text())
        else:
            self.setText("Exact to…")
        self.changed.emit()

    def code(self) -> str:
        if not self._cb.isChecked():
            return ""
        sel = next((r for r in self._radios if r.isChecked()), self._radios[0])
        return sel.property("code") or ""

    def set_code(self, code):
        code = code or ""
        self._cb.setChecked(bool(code))
        for r in self._radios:
            if r.property("code") == code:
                r.setChecked(True)
        self._refresh()


class _YearCalc(QToolButton):
    """Birth-Year calculator (site sfs_calc): «He/she was about [Age] years old in
    [Year]» → sets the bound Year field to Year − Age."""
    def __init__(self, year_edit):
        super().__init__()
        self._yedit = year_edit
        self.setObjectName("addBtn"); self.setText("⊞")
        self.setToolTip("Birth Year Calculator")
        self.setPopupMode(QToolButton.InstantPopup)
        menu = QMenu(self)
        cw = QWidget(); cl = QHBoxLayout(cw); cl.setContentsMargins(8, 6, 8, 6); cl.setSpacing(4)
        _digv = lambda n: QRegularExpressionValidator(QRegularExpression(r"\d{0,%d}" % n))
        cl.addWidget(QLabel("He/she was about"))
        self._age = QLineEdit(); self._age.setPlaceholderText("Age")
        self._age.setFixedWidth(48); self._age.setMaxLength(3); self._age.setValidator(_digv(3))
        cl.addWidget(self._age); cl.addWidget(QLabel("years old in"))
        self._yr = QLineEdit(); self._yr.setPlaceholderText("Year")
        self._yr.setFixedWidth(58); self._yr.setMaxLength(4); self._yr.setValidator(_digv(4))
        cl.addWidget(self._yr)
        calc = QPushButton("Calculate"); calc.clicked.connect(self._calc)
        cl.addWidget(calc)
        wa = QWidgetAction(menu); wa.setDefaultWidget(cw); menu.addAction(wa)
        self.setMenu(menu); self._menu = menu

    def _calc(self):
        try:
            self._yedit.setText(str(int(self._yr.text()) - int(self._age.text())))
        except Exception:
            pass
        self._menu.hide()


class _CategoryTab(QWidget):
    """A category-scoped Ancestry search tab, built from a spec to mirror the LIVE
    category form: First/Last with the SAME match-form dropdowns as the main search;
    an "Add Range" toggle on every date (single Year ↔ Start/End year); family members
    (Father/Mother once, Spouse/Child/Sibling up to 10) each removable with an Exact
    box; an Exact box on every Location/Keyword; a Collection-focus dropdown; and the
    included-collections list. payload() → run_scraper kwargs with category=<code>.
    Touches none of the main-search code."""

    MAX_FAM = 10
    _DAYW, _MONW, _YRW, _LBLW, _TTLW = 62, 104, 84, 72, 92

    def __init__(self, spec):
        super().__init__()
        self.setObjectName("tabPage")          # same light-green bg as the main tab
        self.spec = spec
        self.fields: dict = {}
        self._fam: dict = {}          # role → [row dicts]
        self._fam_boxes: dict = {}    # role → QVBoxLayout
        self._fam_links: dict = {}    # role → "+ add" button
        v = QVBoxLayout(self); v.setContentsMargins(10, 8, 10, 8); v.setSpacing(7)

        head = QHBoxLayout()
        ttl = QLabel(spec["title"]); ttl.setObjectName("sechead")
        self.match_all = QCheckBox("Match all terms exactly")
        clr = QPushButton("Clear search"); clr.setObjectName("addBtn")
        clr.clicked.connect(self._clear)
        head.addWidget(ttl); head.addStretch()
        head.addWidget(self.match_all); head.addSpacing(10); head.addWidget(clr)
        v.addLayout(head)

        # First/Last with the SAME match-form dropdowns as the main search (every tab)
        ng = QGridLayout(); ng.setSpacing(6)
        self.first = QLineEdit(); self.first.setPlaceholderText("First / Middle Name(s)")
        self.first_forms = CheckableComboBox(placeholder="— broad —")
        self.first_forms.add_items(FIRST_FORMS); self.first_forms.setMaximumWidth(150)
        self.last = QLineEdit(); self.last.setPlaceholderText("Last Name")
        self.last_forms = CheckableComboBox(placeholder="— broad —")
        self.last_forms.add_items(SURNAME_FORMS); self.last_forms.setMaximumWidth(150)
        ng.addWidget(QLabel("First / Middle Name(s):"), 0, 0)
        ng.addWidget(self.first, 0, 1); ng.addWidget(self.first_forms, 0, 2)
        ng.addWidget(QLabel("Last Name:"), 0, 3)
        ng.addWidget(self.last, 0, 4); ng.addWidget(self.last_forms, 0, 5)
        ng.setColumnStretch(1, 1); ng.setColumnStretch(4, 1)
        v.addLayout(ng)

        for sec in spec["sections"]:
            self._section(v, sec)

        self._narrow_by_category(v)                  # «Narrow by Category» (one pass per tick)

        if "collection" not in self.fields:        # Collection focus on EVERY tab
            self._collection_focus(v)
        self._collections_dropdown(v)               # "Included data collections" (if known)
        v.addStretch(1)

    # ── date block: ONE LINE — «Title [Day][Month][Year][Exact-to ▾] Add Range».
    # Add Range → HIDE Day/Month/exactness, show Start/End year. Like the live form. #
    def _date_block(self, v, title, key, dm="full", location=True):
        rec: dict = {"ranged": [False]}
        row = QHBoxLayout(); row.setSpacing(5); row.setAlignment(Qt.AlignLeft)
        lab = QLabel(title); lab.setObjectName("sechead"); lab.setMinimumWidth(self._TTLW)
        row.addWidget(lab)
        dmw = []
        if dm == "full":
            d = QComboBox(); d.addItem("Day", "")
            for i in range(1, 32): d.addItem(str(i), str(i))
            d.setFixedWidth(self._DAYW); rec["day"] = d; row.addWidget(d); dmw.append(d)
        if dm in ("full", "month"):
            m = QComboBox()
            for lbl, val in MONTHS: m.addItem(lbl, val)
            m.setFixedWidth(self._MONW); rec["month"] = m; row.addWidget(m); dmw.append(m)
        y  = _mk_year_edit(self._YRW)
        ys = _mk_year_edit(self._YRW); ys.setPlaceholderText("Start"); ys.setVisible(False)
        ye = _mk_year_edit(self._YRW); ye.setPlaceholderText("End"); ye.setVisible(False)
        rng = QComboBox()                              # «Exact to…» year exactness
        for lbl, _c in YEAR_OPTIONS: rng.addItem(lbl)
        rng.setFixedWidth(118)
        link = QPushButton("Add Range"); link.setObjectName("addBtn")
        rec.update(year=y, start=ys, end=ye, range=rng, link=link, dmw=dmw)
        link.clicked.connect(lambda _=0, r=rec: self._toggle_range(r))
        for w in (y, ys, ye, rng, link):
            row.addWidget(w)
        row.addStretch()
        v.addLayout(row)
        if location:
            self._loc_row(v, rec, key)
        self.fields[key] = rec
        v.addWidget(_divider())                        # separator between sections

    def _resized(self):
        """Tell the parent window to re-fit (collapse the empty footer) after the
        tab's content grows or shrinks (add/remove a family row, toggle a range)."""
        cb = getattr(self, "_on_resize", None)
        if cb:
            QTimer.singleShot(0, cb)      # after the removed row leaves the layout

    def _toggle_range(self, rec):
        on = not rec["ranged"][0]; rec["ranged"][0] = on
        rec["year"].setVisible(not on); rec["range"].setVisible(not on)
        for w in rec.get("dmw", []):                   # Day/Month hidden in range mode
            w.setVisible(not on)
        rec["start"].setVisible(on); rec["end"].setVisible(on)
        rec["link"].setText("Single Year" if on else "Add Range")
        self._resized()

    def _loc_row(self, v, rec, key=""):
        r = QHBoxLayout(); r.setAlignment(Qt.AlignLeft)
        lab = QLabel("Location"); lab.setMinimumWidth(self._TTLW)
        lp = QLineEdit(); lp.setPlaceholderText("City, County, State, Country")
        ex = _ExactTo()
        r.addWidget(lab); r.addWidget(lp, 1); r.addWidget(ex)
        v.addLayout(r); rec["place"] = lp; rec["place_exact"] = ex

    def _loc_block(self, v, title, key):
        r = QHBoxLayout(); r.setAlignment(Qt.AlignLeft)
        lab = QLabel(title); lab.setObjectName("sechead"); lab.setMinimumWidth(self._TTLW)
        lp = QLineEdit(); lp.setPlaceholderText("City, County, State, Country")
        ex = _ExactTo()
        r.addWidget(lab); r.addWidget(lp, 1); r.addWidget(ex)
        v.addLayout(r)
        self.fields[key] = {"place": lp, "place_exact": ex}
        v.addWidget(_divider())                        # separator between sections

    def _section(self, v, sec):
        if   sec == "birth_full":      self._date_block(v, "Birth", "birth", "full")
        elif sec == "birth_year":      self._date_block(v, "Birth", "birth", None)
        elif sec == "death_full":      self._date_block(v, "Death", "death", "full")
        elif sec == "marriage_full":   self._date_block(v, "Marriage", "marriage", "full")
        elif sec == "arrival":         self._date_block(v, "Arrival", "arrival", "full")
        elif sec == "departure":       self._date_block(v, "Departure", "departure", "full")
        elif sec == "military":        self._date_block(v, "Military", "military", "full")
        elif sec == "any_event":       self._date_block(v, "Any Event", "any", None)
        elif sec == "any_event_month": self._date_block(v, "Any Event", "any", "month")
        elif sec == "publication":     self._date_block(v, "Publication Info", "publication", "full", location=False)
        elif sec == "residence_date":  self._date_block(v, "Residence Date", "residence_date", None, location=False)
        elif sec == "lived_in":        self._loc_block(v, "Lived In", "residence")
        elif sec == "origin":          self._loc_block(v, "Origin", "origin")
        elif sec == "father_birthplace": self._loc_block(v, "Father's Birthplace", "father_birthplace")
        elif sec == "mother_birthplace": self._loc_block(v, "Mother's Birthplace", "mother_birthplace")
        elif sec == "family":          self._family_block(v)
        elif sec == "keyword":
            r = QHBoxLayout()
            lab = QLabel("Keyword:"); lab.setFixedWidth(self._LBLW)
            kw = QLineEdit(); kw.setPlaceholderText('e.g. pilot or "Flying Tigers"')
            ex = QCheckBox("Exact")
            r.addWidget(lab); r.addWidget(kw, 1); r.addWidget(ex)
            v.addLayout(r); self.fields["keyword"] = {"w": kw, "exact": ex}
        elif sec == "gender":
            r = QHBoxLayout()
            lab = QLabel("Gender:"); lab.setFixedWidth(self._LBLW)
            g = QComboBox(); g.addItems(GENDER_OPTIONS); g.setFixedWidth(140)
            r.addWidget(lab); r.addWidget(g); r.addStretch()
            v.addLayout(r); self.fields["gender"] = {"w": g}
        elif sec == "race":
            r = QHBoxLayout()
            lab = QLabel("Race/Nationality:"); lab.setFixedWidth(self._LBLW)
            rc = QLineEdit(); rc.setPlaceholderText("Race / Nationality")
            r.addWidget(lab); r.addWidget(rc, 1)
            v.addLayout(r); self.fields["race"] = {"w": rc}
        elif sec == "gender_race":     # gender + race on ONE row (like the site)
            r = QHBoxLayout()
            lab = QLabel("Gender:"); lab.setFixedWidth(self._LBLW)
            g = QComboBox(); g.addItems(GENDER_OPTIONS); g.setFixedWidth(140)
            rc = QLineEdit(); rc.setPlaceholderText("Race / Nationality")
            r.addWidget(lab); r.addWidget(g); r.addSpacing(16)
            r.addWidget(QLabel("Race/Nationality:")); r.addWidget(rc, 1)
            v.addLayout(r)
            self.fields["gender"] = {"w": g}; self.fields["race"] = {"w": rc}

    # ── family: order Father/Mother/Spouse/Child/Sibling; Father/Mother max 1, the
    # rest max 10; «+» disables AT the cap (10, not 11); EVERY row removable (✕) ── #
    def _fam_cap(self, role):
        return 1 if role in _FAM_SINGLE else self.MAX_FAM

    def _family_block(self, v):
        # NO leading divider here — the previous section already ends with one
        hdr = QHBoxLayout(); hdr.addWidget(QLabel("Add family member:"))
        roles = [r for r in _FAM_ORDER if r in self.spec.get("family", [])]
        for role in roles:
            lk = QPushButton("+ " + _FAM_LABEL.get(role, role.title())); lk.setObjectName("addBtn")
            lk.clicked.connect(lambda _=0, r=role: self._add_fam(r))
            self._fam_links[role] = lk; hdr.addWidget(lk)
        hdr.addStretch(); v.addLayout(hdr)
        for role in roles:
            box = QVBoxLayout(); box.setSpacing(3)
            self._fam_boxes[role] = box; self._fam[role] = []
            v.addLayout(box)
            self._add_fam(role)            # one starting row each, like the site

    def _add_fam(self, role):
        rows = self._fam.setdefault(role, [])
        if len(rows) >= self._fam_cap(role):
            return
        w = QWidget(); h = QHBoxLayout(w); h.setContentsMargins(0, 0, 0, 0); h.setSpacing(4)
        lab = QLabel(_FAM_LABEL.get(role, role.title())); lab.setFixedWidth(self._LBLW)
        first = QLineEdit(); first.setPlaceholderText("First Name")
        fex = QCheckBox("Exact")
        last  = QLineEdit(); last.setPlaceholderText("Last Name")
        lex = QCheckBox("Exact")
        rm = QPushButton("✕"); rm.setObjectName("rmBtn")
        h.addWidget(lab); h.addWidget(first, 1); h.addWidget(fex)
        h.addWidget(last, 1); h.addWidget(lex); h.addWidget(rm)
        rec = {"first": first, "last": last, "first_exact": fex, "last_exact": lex, "w": w}
        rm.clicked.connect(lambda _=0, r=role, rc=rec: self._del_fam(r, rc))
        self._fam_boxes[role].addWidget(w); rows.append(rec)
        self._update_fam_link(role)
        self._resized()

    def _del_fam(self, role, rec):
        try:
            self._fam[role].remove(rec)
            rec["w"].setParent(None); rec["w"].deleteLater()
        except Exception:
            pass
        self._update_fam_link(role)
        self._resized()

    def _update_fam_link(self, role):
        lk = self._fam_links.get(role)
        if lk is not None:
            lk.setEnabled(len(self._fam.get(role, [])) < self._fam_cap(role))

    # ── collection focus (exact site options) + included-collections list ──── #
    def _collection_focus(self, v):
        v.addWidget(_divider())
        r = QHBoxLayout(); cf = QComboBox()
        for lbl, code in COLLECTION_FOCUS:
            if code is None:                       # group header (disabled)
                cf.addItem(lbl)
                it = cf.model().item(cf.count() - 1)
                if it is not None: it.setEnabled(False)
            else:
                cf.addItem(lbl, code)
        r.addWidget(QLabel("Collection focus:")); r.addWidget(cf); r.addStretch()
        v.addLayout(r); self.fields["collection"] = {"w": cf}

    def _collections_dropdown(self, v):
        items = _CAT_COLLECTIONS.get(str(self.spec["code"]))
        if isinstance(items, str):
            items = _CAT_COLLECTIONS.get(items)
        if not items:
            return
        btn = QPushButton(f"▶  Included data collections ({len(items)})")
        btn.setObjectName("advBtn"); btn.setCheckable(True)
        lst = QListWidget(); lst.addItems(items)
        lst.setMinimumWidth(560); lst.setMaximumHeight(300); lst.setVisible(False)
        btn.toggled.connect(lambda on, b=btn, l=lst, n=len(items): (
            l.setVisible(on),
            b.setText(("▼" if on else "▶") + f"  Included data collections ({n})")))
        v.addWidget(btn); v.addWidget(lst)

    def _narrow_by_category(self, v):
        """«Narrow by Category» — the per-tab category links from the site's sidebar
        (Census.docx). Each checkbox = its filter code; ticking any runs a SEPARATE
        drilled-in pass (its own document), exactly like ticking a filter-tree node."""
        self._narrow_cbs = []
        cats = NARROW_CATEGORIES.get(self.spec["title"], [])
        if not cats:
            return
        box = QGroupBox("Narrow by Category"); box.setObjectName("narrowBox")
        g = QGridLayout(box); g.setSpacing(4); g.setContentsMargins(8, 6, 8, 6)
        for i, (name, code) in enumerate(cats):
            cb = QCheckBox(name.replace("&", "&&"))     # Qt eats a single & as a mnemonic
            cb._narrow_code = code
            g.addWidget(cb, i // 2, i % 2)              # two columns to stay compact
            self._narrow_cbs.append(cb)
        v.addWidget(box)

    def narrow_codes(self) -> list:
        """Ticked «Narrow by Category» as (code, display name) pairs → drilled passes."""
        return [(cb._narrow_code, cb.text().replace("&&", "&"))
                for cb in getattr(self, "_narrow_cbs", []) if cb.isChecked()]

    def _clear(self):
        for le in self.findChildren(QLineEdit): le.clear()
        for cb in self.findChildren(QCheckBox): cb.setChecked(False)
        for co in self.findChildren(QComboBox):
            if isinstance(co, CheckableComboBox):
                co.set_checked([])
            elif co.count():
                co.setCurrentIndex(0)

    # ── payload ───────────────────────────────────────────────────────────── #
    def has_name(self) -> bool:
        return bool(self.first.text().strip() or self.last.text().strip())

    @staticmethod
    def _forms_code(combo) -> str:
        """Per-field name-match code for Ancestry's <field>_x: fuzzy letters p/s/i
        (sounds/similar/initials), else «1» for exact-only, else «»."""
        forms = [c for _l, c in combo.checked()]
        fuzzy = "".join(l for l in ("p", "s", "i") if l in forms)
        return fuzzy or ("1" if "1" in forms else "")

    def _date_value(self, rec) -> str:
        """Ancestry date token: «Y-M-D» (month/day optional) or a range «start---end»."""
        if rec.get("ranged", [False])[0]:
            s = rec["start"].text().strip(); e = rec["end"].text().strip()
            return f"{s}---{e}" if (s or e) else ""
        y = rec["year"].text().strip()
        if not y:
            return ""
        mo = rec["month"].currentData() if rec.get("month") else ""
        da = rec["day"].currentData() if rec.get("day") else ""
        out = y
        if mo:
            out += f"-{mo}"
        if da:
            out += (f"-{da}" if mo else f"-0-{da}")
        return out

    def bad_years(self) -> list:
        """Filled year fields that aren't a full 4-digit year (1, 11, …)."""
        bad = []
        for rec in self.fields.values():
            for k in ("year", "start", "end"):
                ed = rec.get(k)
                if ed is not None:
                    t = ed.text().strip()
                    if t and (not t.isdigit() or len(t) != 4):
                        bad.append(t)
        return bad

    def payload(self) -> dict:
        """Build the category-search URL params (k, v) EXACTLY in Alla's live-URL
        format: name=First_Last & name_x=<first>_<last>; dates <key>=<Y-M-D|a---b>_place;
        family <rel>=First_Last & <rel>_x=<first>_<last>; place-exactness, priority…"""
        f = self.fields
        P = []
        fn = self.first.text().strip(); ln = self.last.text().strip()
        P.append(("name", f"{_uv(fn)}_{_uv(ln)}".strip("_")))
        fc = self._forms_code(self.first_forms); lc = self._forms_code(self.last_forms)
        if self.match_all.isChecked():
            fc = fc or "1"; lc = lc or "1"
        if fc or lc:
            P.append(("name_x", f"{fc}_{lc}"))

        # «Lived In» place + «Residence Date» year → ONE residence=<year>_<place>
        res_place, res_pe = "", ""
        if "residence" in f and f["residence"].get("place"):
            res_place = f["residence"]["place"].text().strip()
            res_pe = f["residence"]["place_exact"].code()

        for key, rec in f.items():
            if key in ("keyword", "gender", "race", "collection", "residence"):
                continue
            is_date = ("year" in rec or "start" in rec)
            place = rec["place"].text().strip() if rec.get("place") else ""
            pe = rec["place_exact"].code() if rec.get("place_exact") else ""
            val = self._date_value(rec) if is_date else ""
            if key == "residence_date":              # carries the year; place from Lived In
                urlkey, place, pe = "residence", res_place, res_pe
            else:
                urlkey = _DATE_KEY.get(key, key)
            if not (val or place):
                continue
            P.append((urlkey, val + (f"_{_uv(place)}" if place else "")))
            # year exactness — emit <key>_x=<span>-0-0 whenever a single year was entered,
            # INCLUDING «Exact (this year)» → 0-0-0 (the site sends it; «if sv» dropped it)
            span = ""
            if is_date and val and "---" not in val and not rec.get("ranged", [False])[0] and rec.get("range"):
                span = f"{YEAR_OPTIONS[rec['range'].currentIndex()][1]}-0-0"
            if span or pe:
                P.append((f"{urlkey}_x", f"{span}_{pe}" if pe else span))
        if res_place and "residence_date" not in f:   # Lived In without a Residence Date row
            P.append(("residence", f"_{_uv(res_place)}"))
            if res_pe:
                P.append(("residence_x", f"_{res_pe}"))

        for role in _FAM_ORDER:                        # Father, Mother, Spouse, Child, Sibling
            for r in self._fam.get(role, []):
                rf = r["first"].text().strip(); rl = r["last"].text().strip()
                if not (rf or rl):
                    continue
                P.append((role, f"{_uv(rf)}_{_uv(rl)}".strip("_")))
                fe = "1" if r["first_exact"].isChecked() else ""
                le = "1" if r["last_exact"].isChecked() else ""
                if fe or le:
                    P.append((f"{role}_x", f"{fe}_{le}"))

        if "keyword" in f and f["keyword"]["w"].text().strip():
            P.append(("keyword", _uv(f["keyword"]["w"].text())))
            if f["keyword"]["exact"].isChecked():
                P.append(("keyword_x", "1"))
        if "gender" in f:
            g = f["gender"]["w"].currentText().strip().lower()
            if g.startswith("m"):   P.append(("gender", "m"))
            elif g.startswith("f"): P.append(("gender", "f"))
        if "race" in f and f["race"]["w"].text().strip():
            P.append(("race", _uv(f["race"]["w"].text()))); P.append(("race_x", "1"))
        if "collection" in f and f["collection"]["w"].currentData():
            P.append(("priority", f["collection"]["w"].currentData()))
        return {
            "category":        self.spec["code"],
            "category_params": P,
            "narrow":          self.narrow_codes(),   # «Narrow by Category» → drilled passes
            "first_names":     fn,
            "last_names":      ln,
            "advanced":        {},
            "exact":           {},
            "filters":         {},
        }

    # ── autosave: remember EVERY field of this category tab ──────────────────── #
    def state(self) -> dict:
        st = {"first": self.first.text(), "last": self.last.text(),
              "match": self.match_all.isChecked(),
              "ff": [l for l, _c in self.first_forms.checked()],
              "lf": [l for l, _c in self.last_forms.checked()],
              "nc": [c for c, _n in self.narrow_codes()], "fields": {}, "fam": {}}
        for k, r in self.fields.items():
            d = {}
            for key in ("year", "start", "end"):
                if r.get(key) is not None: d[key] = r[key].text()
            if "ranged" in r: d["ranged"] = r["ranged"][0]
            if r.get("range") is not None: d["ri"] = r["range"].currentIndex()
            if r.get("day") is not None: d["day"] = r["day"].currentData()
            if r.get("month") is not None: d["mon"] = r["month"].currentData()
            if r.get("place") is not None: d["place"] = r["place"].text()
            pe = r.get("place_exact")
            if pe is not None and hasattr(pe, "code"): d["pe"] = pe.code()
            w = r.get("w")
            if isinstance(w, QLineEdit): d["w"] = w.text()
            elif isinstance(w, QComboBox): d["wi"] = w.currentIndex()
            ex = r.get("exact")
            if ex is not None: d["ex"] = ex.isChecked()
            st["fields"][k] = d
        for role, rows in self._fam.items():
            st["fam"][role] = [[x["first"].text(), x["last"].text(),
                                x["first_exact"].isChecked(), x["last_exact"].isChecked()]
                               for x in rows]
        return st

    def restore(self, st: dict):
        if not st:
            return
        self.first.setText(st.get("first", "")); self.last.setText(st.get("last", ""))
        self.match_all.setChecked(bool(st.get("match")))
        self.first_forms.set_checked(st.get("ff", [])); self.last_forms.set_checked(st.get("lf", []))
        for cb in getattr(self, "_narrow_cbs", []):     # «Narrow by Category» ticks
            cb.setChecked(cb._narrow_code in (st.get("nc") or []))
        for k, d in (st.get("fields") or {}).items():
            r = self.fields.get(k)
            if not r:
                continue
            for key in ("year", "start", "end"):
                if r.get(key) is not None and key in d: r[key].setText(d[key] or "")
            if r.get("range") is not None and "ri" in d: r["range"].setCurrentIndex(int(d["ri"]))
            if r.get("day") is not None and d.get("day"): r["day"].setCurrentIndex(max(0, r["day"].findData(d["day"])))
            if r.get("month") is not None and d.get("mon"): r["month"].setCurrentIndex(max(0, r["month"].findData(d["mon"])))
            if r.get("place") is not None and "place" in d: r["place"].setText(d["place"] or "")
            pe = r.get("place_exact")
            if pe is not None and hasattr(pe, "set_code") and "pe" in d: pe.set_code(d["pe"])
            w = r.get("w")
            if isinstance(w, QLineEdit) and "w" in d: w.setText(d["w"] or "")
            elif isinstance(w, QComboBox) and "wi" in d: w.setCurrentIndex(int(d["wi"]))
            ex = r.get("exact")
            if ex is not None and "ex" in d: ex.setChecked(bool(d["ex"]))
            if d.get("ranged") and not r.get("ranged", [False])[0]:
                self._toggle_range(r)
        for role, rows in (st.get("fam") or {}).items():
            existing = self._fam.get(role, [])
            for i, rd in enumerate(rows):
                rec = existing[i] if i < len(existing) else self._add_fam(role)
                if not rec:
                    continue
                rec["first"].setText(rd[0] or ""); rec["last"].setText(rd[1] or "")
                rec["first_exact"].setChecked(bool(rd[2])); rec["last_exact"].setChecked(bool(rd[3]))


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
        # Top-level content scroll + a FIXED bottom bar: the form scrolls, Start/Cancel
        # stay always visible, and the window never grows past the screen.
        outer_root = QWidget(); self.setCentralWidget(outer_root)
        _ol = QVBoxLayout(outer_root); _ol.setContentsMargins(0, 0, 0, 0); _ol.setSpacing(0)
        self._content_scroll = QScrollArea()
        self._content_scroll.setWidgetResizable(True)
        self._content_scroll.setFrameShape(QFrame.NoFrame)
        self._content_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        _ol.addWidget(self._content_scroll, 1)
        root = QWidget(); self._content_scroll.setWidget(root)
        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(18, 12, 18, 12)   # keep a right/left field (no flush-to-edge)
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

        # ── Tabs: the existing Search (basic + advanced + filters) is tab 0; each
        # category search is its own tab (assembled lower down). Credentials (above)
        # and Output + Start/Cancel (below) are SHARED across every tab. The main
        # Search tab and all its code are left exactly as before.
        self._tabs = QTabWidget()
        self._tabs.setTabBar(_StretchTabBar())         # tabs stretch to the window's right edge
        self._tabs.tabBar().setUsesScrollButtons(False)
        self._tab0 = QWidget(); self._tab0.setObjectName("tabPage")
        self._tab0_lay = QVBoxLayout(self._tab0)
        self._tab0_lay.setContentsMargins(6, 6, 6, 6); self._tab0_lay.setSpacing(8)

        # Basic Search
        bg = QGroupBox("Basic Search")
        bf = QGridLayout(bg); bf.setSpacing(8)
        self.f_first = QLineEdit(); self.f_first.setPlaceholderText("First / Middle Name(s)")
        self.f_last  = QLineEdit(); self.f_last.setPlaceholderText("Last Name")
        self.f_place = QLineEdit(); self.f_place.setPlaceholderText("City, county, state, country")
        self.f_byear = QLineEdit(); self.f_byear.setPlaceholderText("e.g. 1897")
        self.f_byear.setFixedWidth(90)
        self.f_year_range = QComboBox()
        for lbl, _v in YEAR_OPTIONS:
            self.f_year_range.addItem(lbl)
        self.f_year_range.setCurrentIndex(1)        # ± 1 year (default)
        # multi-select match forms (checkboxes in the dropdown) — like the site
        self.f_first_exact = CheckableComboBox(placeholder="— broad —")
        self.f_first_exact.add_items(FIRST_FORMS)
        self.f_first_exact.setToolTip("Tick any combination of name match forms")
        self.f_first_exact.setMaximumWidth(170)
        self.f_last_exact  = CheckableComboBox(placeholder="— broad —")
        self.f_last_exact.add_items(SURNAME_FORMS)
        self.f_last_exact.setToolTip("Tick any combination of surname match forms")
        self.f_last_exact.setMaximumWidth(170)
        self.f_first_exact.changed.connect(self._save)
        self.f_last_exact.changed.connect(self._save)
        self.f_place_exact = _ExactTo()              # site «Exact to…» popup (radios)
        self.f_place_exact.changed.connect(self._save)
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
        yr.addWidget(self.f_byear); yr.addWidget(_YearCalc(self.f_byear))
        yr.addWidget(self.f_year_range, 1)
        bf.addLayout(yr, 1, 4, 1, 2)
        bf.setColumnStretch(1, 2); bf.setColumnStretch(4, 1)
        self._tab0_lay.addWidget(bg)
        # Advanced toggle (left) + «Clear search» (right) on ONE line (no staircase)
        _adv_row = QHBoxLayout()
        self._adv_btn = QPushButton("▶   Advanced Search")
        self._adv_btn.setObjectName("advBtn")
        self._adv_btn.setCheckable(True)
        self._adv_btn.toggled.connect(self._toggle_adv)
        _clr_btn = QPushButton("Clear search"); _clr_btn.setObjectName("addBtn")
        _clr_btn.clicked.connect(self._clear_filters)
        _adv_row.addWidget(self._adv_btn); _adv_row.addStretch(); _adv_row.addWidget(_clr_btn)
        self._tab0_lay.addLayout(_adv_row)

        self._adv = QGroupBox()
        av = QVBoxLayout(self._adv); av.setSpacing(6)

        # Events — "Add event: Marriage Death Lived In Any Event" (link row)
        ev_head = QHBoxLayout(); ev_head.setSpacing(10)
        ev_head.addWidget(QLabel("Add event:"))
        self._event_links = {}
        for label, etype in EVENT_TYPES:
            lk = QPushButton(label); lk.setObjectName("addBtn")
            lk.clicked.connect(lambda _=False, t=etype: (self._add_event_row(t), self._save()))
            self._event_links[etype] = lk
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
            lk.setFocusPolicy(Qt.NoFocus)            # no focus rectangle after a click
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
        self.f_collection = QComboBox()              # same options as the category tabs
        for _lbl, _code in COLLECTION_FOCUS:
            if _code is None:
                self.f_collection.addItem(_lbl)
                _it = self.f_collection.model().item(self.f_collection.count() - 1)
                if _it is not None: _it.setEnabled(False)
            else:
                self.f_collection.addItem(_lbl, _code)
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

        self._adv.setVisible(False)            # the single top-level scroll handles overflow
        self._tab0_lay.addWidget(self._adv)

        # ── Result filters (left-panel; applied AFTER the search) ──────────── #
        self._flt_btn = QPushButton("▶   Filters  (Record type / Location / Date — multi-pass)")
        self._flt_btn.setObjectName("advBtn")
        self._flt_btn.setCheckable(True)
        self._flt_btn.toggled.connect(self._toggle_flt)
        self._tab0_lay.addWidget(self._flt_btn)

        self._flt = QGroupBox()
        fv = QVBoxLayout(self._flt); fv.setSpacing(4)
        hdr = QHBoxLayout()
        hdr.addWidget(QLabel(
            "Click ▶ to open each layer; tick any value (each becomes its own pass "
            "/ document). Location is combined (AND) with each."))
        hdr.addStretch()
        self._collapse_flt_btn = QPushButton("Collapse all filters")
        self._collapse_flt_btn.setObjectName("addBtn")
        self._collapse_flt_btn.clicked.connect(self._collapse_all_filters)
        hdr.addWidget(self._collapse_flt_btn)
        self._clear_flt_btn = QPushButton("Clear all filters")
        self._clear_flt_btn.setObjectName("addBtn")
        self._clear_flt_btn.clicked.connect(self._clear_filters)
        hdr.addWidget(self._clear_flt_btn)
        fv.addLayout(hdr)
        # each section keeps its ticked-node checkboxes, leaf dropdowns, and a
        # path→widgets registry (so saved selections can be rebuilt lazily)
        self._tree_bodies = []     # (body, toggle) for every expandable node → collapse-all
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

        self._flt.setVisible(False)            # the single top-level scroll handles overflow
        self._tab0_lay.addWidget(self._flt)
        self._tab0_lay.addStretch(1)           # absorb spare height → no giant gap inside Basic Search

        # tab 0 = the existing Search; EVERY tab (incl. this one) lives in its OWN
        # scroll so a tall form scrolls INSIDE the tab — never squashing its rows and
        # never inheriting the tallest tab's height (which left an empty footer)
        sc0 = QScrollArea(); sc0.setWidgetResizable(True); sc0.setFrameShape(QFrame.NoFrame)
        sc0.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        sc0.setWidget(self._tab0)
        self._tabs.addTab(sc0, "All Collections")
        self._cat_tabs = []
        for spec in CATEGORY_SPECS:
            ct = _CategoryTab(spec)
            sc = QScrollArea(); sc.setWidgetResizable(True); sc.setFrameShape(QFrame.NoFrame)
            sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            sc.setWidget(ct)                          # own scroll → a tall tab scrolls INSIDE the tab
            self._tabs.addTab(sc, spec["title"].replace("&", "&&"))   # «&&» → visible «&»
            self._cat_tabs.append(ct)
        # autosave + re-fit on layout change (kills the empty footer after add/remove)
        for ct in self._cat_tabs:
            ct._on_resize = self._fit
            for le in ct.findChildren(QLineEdit):  le.textChanged.connect(self._save)
            for co in ct.findChildren(QComboBox):  co.currentIndexChanged.connect(self._save)
            for cb in ct.findChildren(QCheckBox):  cb.stateChanged.connect(self._save)
            for ex in ct.findChildren(_ExactTo):   ex.changed.connect(self._save)
        self._tabs.currentChanged.connect(lambda *_: self._shrink_tabs())
        self._outer.addWidget(self._tabs)

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

        # ── Fixed bottom bar (OUTSIDE the scroll → progress + Start/Cancel ALWAYS visible) ──
        self._bottom = QWidget()
        _bl = QVBoxLayout(self._bottom)
        _bl.setContentsMargins(18, 4, 18, 8); _bl.setSpacing(6)
        self.pbar  = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        _bl.addWidget(self.pbar)
        _bl.addWidget(self.stlbl)
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        _bl.addLayout(br)
        _bl.addWidget(make_footer())
        _ol.addWidget(self._bottom, 0)

        for w in self._static_widgets():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QComboBox): w.currentTextChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)
        # filter checkboxes/combos are wired to _save inside their builders
        self._fit()
        self._shrink_tabs()

    def _toggle_flt(self, on: bool):
        self._flt.setVisible(on)
        self._flt_btn.setText(("▼" if on else "▶") +
                              "   Filters  (Record type / Location / Date — multi-pass)")
        self._fit()
        QTimer.singleShot(0, self._fit)        # re-fit after the layout settles

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
            self._tree_bodies.append((body, tog))   # for «Collapse all filters»
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

    def _collapse_all_filters(self):
        """Collapse every expanded node in the filter tree (hides the bodies, resets
        the arrows) and recompact the window. Folds the whole tree in one click."""
        for body, tog in getattr(self, "_tree_bodies", []):
            try:
                body.setVisible(False)
                tog.setText("▶")
            except Exception:
                pass
        self._fit()

    def _clear_filters(self):
        """Reset EVERY search constraint to broad (keeps the name being searched):
        Birth Year, Gender, name/place exactness, place, spouse/family members,
        events, keyword, race, collection focus + checkboxes, and the filter tree.
        Without this the stale Birth Year / Gender keep wiping the results."""
        # basic constraints
        self.f_place.clear()
        self.f_byear.clear()
        self.f_year_range.setCurrentIndex(0)
        self.f_first_exact.set_checked([])           # broad (nothing ticked)
        self.f_last_exact.set_checked([])
        self.f_place_exact.set_code("")
        # advanced
        self.f_gender.setCurrentIndex(0)             # «—» = no gender
        self.f_race.clear()
        self.f_keyword.clear()
        self.f_collection.setCurrentIndex(0)         # All Collections
        for cb in (self.f_hist, self.f_trees, self.f_stories, self.f_photos):
            cb.setChecked(True)                      # site default: all four on
        # remove every spouse/family + event row
        for key in list(self._fam):
            for rec in list(self._fam[key]):
                self._remove_fam_row(key, rec)
        for rec in list(self._events):
            self._remove_event_row(rec)
        # the record-type / location / date filter tree
        for reg in (self._rt_reg, self._loc_reg, self._rd_reg):
            for e in reg.values():
                if e.get("cb"):
                    e["cb"].setChecked(False)
                if e.get("combo"):
                    e["combo"].set_checked([])
        self._save(); self._fit()

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
        cap = 1 if key in _FAM_SINGLE else 10        # Father/Mother once; the rest up to 10
        if len(self._fam[key]) >= cap:
            return None
        box = self._fam_containers[key]
        row_w = QWidget()
        rl = QHBoxLayout(row_w); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(6)
        lbl = QLabel(_FAM_LABEL[key] + ":"); lbl.setFixedWidth(60)
        ff = QLineEdit(); ff.setPlaceholderText("First / Middle Name(s)"); ff.setText(first)
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
        self._update_fam_link_main(key)              # grey out «+» at the cap
        self._fit()
        return rec

    def _update_fam_link_main(self, key):
        lk = getattr(self, "_fam_links", {}).get(key)
        if lk is not None:
            cap = 1 if key in _FAM_SINGLE else 10
            lk.setEnabled(len(self._fam[key]) < cap)

    def _remove_fam_row(self, key, rec):
        try:
            self._fam[key].remove(rec)
            rec["w"].setParent(None)
            rec["w"].deleteLater()
        except Exception:
            pass
        self._update_fam_link_main(key)
        self._save(); self._fit()

    # ── Dynamic event rows ────────────────────────────────────────────────── #
    def _add_event_row(self, etype, data=None):
        if sum(1 for e in self._events if e["type"] == etype) >= _EVENT_CAP.get(etype, 10):
            return None
        data = data or {}
        dm = None                                    # All-Collections events: YEAR only (no Day/Month)
        row_w = QWidget()
        rl = QHBoxLayout(row_w); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(5)
        rl.setAlignment(Qt.AlignLeft)
        lbl = QLabel(_EVENT_LABEL.get(etype, etype)); lbl.setObjectName("sechead")
        lbl.setMinimumWidth(78); rl.addWidget(lbl)
        rec = {"w": row_w, "type": etype, "ranged": [False]}
        dmw = []
        if dm == "full":
            d = QComboBox(); d.addItem("Day", "")
            for i in range(1, 32): d.addItem(str(i), str(i))
            d.setFixedWidth(58)
            if data.get("day"): d.setCurrentIndex(max(0, d.findData(data["day"])))
            rec["day"] = d; rl.addWidget(d); dmw.append(d)
            m = QComboBox()
            for ml, mv in MONTHS: m.addItem(ml, mv)
            m.setFixedWidth(96)
            if data.get("month"): m.setCurrentIndex(max(0, m.findData(data["month"])))
            rec["month"] = m; rl.addWidget(m); dmw.append(m)
        y  = _mk_year_edit(80); y.setText(data.get("year", ""))
        ys = _mk_year_edit(80); ys.setPlaceholderText("Start"); ys.setText(data.get("start", "")); ys.setVisible(False)
        ye = _mk_year_edit(80); ye.setPlaceholderText("End"); ye.setText(data.get("end", "")); ye.setVisible(False)
        rngc = QComboBox()
        for rl2, _rv in YEAR_OPTIONS: rngc.addItem(rl2)
        rngc.setCurrentIndex(int(data["range_i"]) if str(data.get("range_i", "")).isdigit() else 1)
        rngc.setFixedWidth(118)
        link = QPushButton("Add Range"); link.setObjectName("addBtn")
        rec.update(year=y, start=ys, end=ye, range=rngc, link=link, dmw=dmw)
        link.clicked.connect(lambda _=0, r=rec: (self._toggle_event_range(r), self._save()))
        for wdg in (y, ys, ye, rngc, link): rl.addWidget(wdg)
        pf = QLineEdit(); pf.setPlaceholderText("City, County, State, Country"); pf.setText(data.get("place", ""))
        ex = _ExactTo(); ex.set_code(data.get("place_exact", ""))
        rl.addWidget(QLabel("Loc:")); rl.addWidget(pf, 2); rl.addWidget(ex)
        rec["place"] = pf; rec["place_exact"] = ex
        rm = QPushButton("✕"); rm.setObjectName("rmBtn"); rm.setFixedWidth(24); rl.addWidget(rm)
        self._events.append(rec); self._events_box.addWidget(row_w)
        if data.get("ranged"): self._toggle_event_range(rec)
        for wdg in (y, ys, ye, pf): wdg.textChanged.connect(self._save)
        rngc.currentTextChanged.connect(self._save); ex.changed.connect(self._save)
        for cb in dmw: cb.currentTextChanged.connect(self._save)
        rm.clicked.connect(lambda _=False: self._remove_event_row(rec))
        self._update_event_link(etype)
        self._fit()
        return rec

    def _toggle_event_range(self, rec):
        on = not rec["ranged"][0]; rec["ranged"][0] = on
        rec["year"].setVisible(not on); rec["range"].setVisible(not on)
        for w in rec.get("dmw", []): w.setVisible(not on)
        rec["start"].setVisible(on); rec["end"].setVisible(on)
        rec["link"].setText("Single Year" if on else "Add Range")

    def _event_date(self, rec):
        """Ancestry date token: «Y-M-D» (month/day optional) or «start---end»."""
        if rec["ranged"][0]:
            s = rec["start"].text().strip(); e = rec["end"].text().strip()
            return f"{s}---{e}" if (s or e) else ""
        y = rec["year"].text().strip()
        if not y:
            return ""
        mo = rec["month"].currentData() if rec.get("month") else ""
        da = rec["day"].currentData() if rec.get("day") else ""
        out = y
        if mo: out += f"-{mo}"
        if da: out += (f"-{da}" if mo else f"-0-{da}")
        return out

    def _update_event_link(self, etype):
        lk = getattr(self, "_event_links", {}).get(etype)
        if lk is not None:
            lk.setEnabled(sum(1 for e in self._events if e["type"] == etype)
                          < _EVENT_CAP.get(etype, 10))

    def _remove_event_row(self, rec):
        et = rec.get("type", "")
        try:
            self._events.remove(rec)
            rec["w"].setParent(None)
            rec["w"].deleteLater()
        except Exception:
            pass
        self._update_event_link(et)
        self._save(); self._fit()

    # ── Advanced toggle / fit ─────────────────────────────────────────────── #
    def _toggle_adv(self, on: bool):
        self._adv.setVisible(on)
        self._adv_btn.setText(("▼" if on else "▶") + "   Advanced Search")
        self._fit()
        QTimer.singleShot(0, self._fit)        # re-fit after the layout settles

    def _shrink_tabs(self):
        """Re-fit on tab switch: each category tab is in its own scroll, so the
        stack's minimum stays small and the window shrinks to a SHORT tab (collapsing
        empty space); a tall tab scrolls inside its own area, buttons stay visible."""
        self._fit()

    def _fit(self):
        """Open at the tab-bar WIDTH and cap the tab area to the CURRENT tab's content
        (so a short tab has NO empty footer; a tall tab scrolls inside its own area).
        The fixed bottom bar keeps Start/Cancel visible. No move() (launcher centers)."""
        sw = self.screen() or QApplication.primaryScreen()
        scr = sw.availableGeometry()
        self.setMinimumHeight(0); self.setMaximumHeight(16777215)
        tabs = getattr(self, "_tabs", None)
        tab_w = (tabs.tabBar().sizeHint().width() + 36) if tabs is not None else 0
        self.setMinimumWidth(min(max(920, tab_w), scr.width() - 16))    # never narrower than the tabs
        if tabs is not None:
            tabs.setMinimumHeight(0); tabs.setMaximumHeight(16777215)   # release before measuring
        cw = self._content_scroll.widget()
        self._outer.invalidate(); self._outer.activate(); cw.adjustSize()
        bottom_h = self._bottom.sizeHint().height() if hasattr(self, "_bottom") else 0
        if tabs is not None:
            tabbar = tabs.tabBar().sizeHint().height()
            chrome = cw.sizeHint().height() - tabs.sizeHint().height()  # header / creds / output
            cur = tabs.currentWidget()
            is_scroll = isinstance(cur, QScrollArea)
            page = cur.widget() if is_scroll and cur.widget() else cur
            page_h = page.sizeHint().height() if page is not None else 0
            avail = scr.height() - 48 - chrome - bottom_h - tabbar - 12  # screen room for the tab area
            if is_scroll:
                # a category tab scrolls INSIDE its own area → CAP the tab block to the
                # current tab's content (kills the empty footer; a tall tab scrolls in sc)
                tab_h = min(page_h, max(140, avail))
                tabs.setMinimumHeight(tab_h + tabbar + 4)
                tabs.setMaximumHeight(tab_h + tabbar + 4)
            else:
                # «All Collections» has NO inner scroll → DON'T constrain the tab block
                # (full natural height); the OUTER page scroll handles overflow so the
                # event / family / filter rows never get squashed or overlapped
                tabs.setMinimumHeight(0); tabs.setMaximumHeight(16777215)
                tab_h = page_h
            hint = chrome + tabbar + tab_h + 16 + bottom_h
        else:
            hint = cw.sizeHint().height() + bottom_h + 8
        want_w = max(self.width(), tab_w, cw.sizeHint().width() + 4)
        w = min(want_w, scr.width() - 16)
        self.resize(w, min(hint, scr.height() - 48))
        self.setMaximumHeight(scr.height() - 48)
        clamp_on_screen(self)                            # keep on-screen after the resize
        # Nudge the bar to re-lay-out its equal-width tabs at the settled width —
        # updateGeometry ONLY, never setFixedWidth (forcing a fixed width fights the bar
        # on every tab switch → the "jumping tabs"). Equal widths are selection-stable.
        if tabs is not None:
            QTimer.singleShot(0, lambda: tabs.tabBar().updateGeometry())

    # ── Autosave ──────────────────────────────────────────────────────────── #
    def _static_widgets(self) -> list:
        # f_first_exact / f_last_exact are CheckableComboBox → wired via .changed
        return [self.f_user, self.f_pass, self.f_first, self.f_last, self.f_place,
                self.f_byear, self.f_year_range, self.f_keyword,
                self.f_gender, self.f_race, self.f_collection,
                self.f_hist, self.f_trees, self.f_stories, self.f_photos,
                self.f_folder, self.f_docx, self.f_xlsx]

    def _save(self, *_):
        d = {
            "username": self.f_user.text(), "password": self.f_pass.text(),
            "first_names": self.f_first.text(), "last_names": self.f_last.text(),
            "place_lived": self.f_place.text(), "birth_year": self.f_byear.text(),
            "year_range_i": self.f_year_range.currentIndex(),
            "first_forms": [l for l, _c in self.f_first_exact.checked()],
            "last_forms":  [l for l, _c in self.f_last_exact.checked()],
            "place_exact": self.f_place_exact.code(),
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
            "events": [{"type": e["type"], "year": e["year"].text(),
                        "start": e["start"].text(), "end": e["end"].text(),
                        "ranged": e["ranged"][0], "range_i": e["range"].currentIndex(),
                        "day": e["day"].currentData() if e.get("day") else "",
                        "month": e["month"].currentData() if e.get("month") else "",
                        "place": e["place"].text(), "place_exact": e["place_exact"].code()}
                       for e in self._events],
            "cat_states": [t.state() for t in getattr(self, "_cat_tabs", [])],
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
        self.f_first_exact.set_checked(d.get("first_forms") or [])
        self.f_last_exact.set_checked(d.get("last_forms") or [])
        self.f_place_exact.set_code(d.get("place_exact", "")); _s(self.f_keyword, "keyword")
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
            if isinstance(e, dict):
                self._add_event_row(e.get("type", "any"), e)
            else:                                        # legacy [type, year, range, place]
                e = (list(e) + ["", "", 1, ""])[:4]
                self._add_event_row(e[0], {"year": str(e[1]), "place": str(e[3])})
        # restore the full filter selection (rebuilds the needed lazy branches)
        if d.get("filters"):
            self._restore_filters(d["filters"])
        if d.get("adv_open"):
            self._adv_btn.setChecked(True)
        if d.get("flt_open"):
            self._flt_btn.setChecked(True)
        for t, stt in zip(getattr(self, "_cat_tabs", []), d.get("cat_states") or []):
            try: t.restore(stt)
            except Exception: pass

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
               "collection": (self.f_collection.currentData() or ""),
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
        adv["events"] = []
        for e in self._events:
            date = self._event_date(e)
            place = e["place"].text().strip()
            if date or place:
                adv["events"].append({"type": e["type"], "date": date, "place": place,
                                      "place_exact": e["place_exact"].code()})
        return adv

    def _search_payload(self) -> dict:
        """The main Search tab → run_scraper search keys (basic + advanced + filters).
        Unchanged from before — the category tabs build their own payload separately."""
        return {
            "first_names": self.f_first.text().strip(),
            "last_names":  self.f_last.text().strip(),
            "place_lived": self.f_place.text().strip(),
            "birth_year":  self.f_byear.text().strip(),
            "year_range":  YEAR_OPTIONS[self.f_year_range.currentIndex()][1],
            "advanced":    self._build_advanced(),
            "exact": {
                "place":         self.f_place_exact.code(),
                "name_forms":    [c for _l, c in self.f_first_exact.checked()],
                "surname_forms": [c for _l, c in self.f_last_exact.checked()],
            },
            "filters": {
                "types":     self._type_filters(),       # [{label, code}]
                "dates":     self._date_filters(),       # [{label, code}]
                "locations": self._location_filters(),   # [{label, code}] record_f
            },
        }

    def _payload(self) -> dict:
        # SHARED fields (credentials / output) + the ACTIVE tab's search fields:
        # tab 0 = the main Search; any other tab = that category's _CategoryTab.
        base = {
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "email":    self.f_user.text().strip() or None,
            "password": self.f_pass.text() or None,
            "log":      print,
            "cancel_event": getattr(self, "_cancel_ev", None),
        }
        idx = self._tabs.currentIndex()
        if idx <= 0:
            base.update(self._search_payload())
        else:
            base.update(self._cat_tabs[idx - 1].payload())
        return base

    def _validate(self) -> bool:
        idx = self._tabs.currentIndex()
        has_name = (bool(self.f_first.text().strip() or self.f_last.text().strip())
                    if idx <= 0 else self._cat_tabs[idx - 1].has_name())
        if not has_name:
            QMessageBox.warning(self, "Nothing to search",
                                "Enter at least a first or last name.")
            return False
        bad = []
        if idx <= 0:
            by = self.f_byear.text().strip()
            if by and (not by.isdigit() or len(by) != 4):
                bad = [by]
        else:
            bad = self._cat_tabs[idx - 1].bad_years()
        if bad:
            QMessageBox.warning(self, "Invalid year",
                                "Enter a full 4-digit year (e.g. 1897).\nInvalid: "
                                + ", ".join(bad))
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

    def showEvent(self, ev):
        super().showEvent(ev)
        if not getattr(self, "_shown_once", False):
            self._shown_once = True
            # re-fit on several ticks: the launcher's center_window() runs at
            # singleShot(0) and would shrink us — the LAST tick re-applies the
            # tab-bar width so the window always opens wide enough for all tabs.
            for _d in (0, 60, 180):
                QTimer.singleShot(_d, self._fit)
            # re-fit when the window is dragged to ANOTHER screen (e.g. a lower
            # resolution) so it shrinks to that screen and the Start/Cancel bar
            # never falls off the bottom.
            wh = self.windowHandle()
            if wh is not None:
                wh.screenChanged.connect(lambda *_: QTimer.singleShot(0, self._fit))

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
