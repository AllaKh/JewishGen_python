"""
gui/myheritage.py  —  v6
─────────────────────────
Changes:
* Site selector (combo) in GUI: Israel (.co.il), English, Russian, Hebrew, etc.
  — passed straight to the scraper, not hardcoded.
* 2FA support: when the scraper detects the verification-code dialog it calls
  ask_2fa_code() which triggers a Qt dialog asking the user to enter the code.
  The dialog stays open until the user clicks OK or Cancel.
* No QScrollArea — window auto-fits to content.
* Autosave / restore all fields including site selection.
"""

import json, re, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox, QInputDialog,
    QApplication, QGroupBox, QComboBox, QSpinBox,
    QScrollArea, QRadioButton, QButtonGroup, QFrame, QGridLayout,
    QToolButton, QMenu, QWidgetAction,
)
from PySide6.QtCore import QThread, Signal, Qt, QByteArray
from PySide6.QtGui import QPixmap, QIcon, QValidator, QAction
from gui._app_icon import app_icon, make_header, make_cancel_button


class _YearSpin(QSpinBox):
    """Year field that can be CLEARED back to empty (value 0 → «—»).

    A plain QSpinBox refuses empty text and reverts to the last value, so an
    entered year could only be *changed*, never *deleted* (user-reported bug).
    Accepting empty/«—» as valid lets the user backspace the field to clear it.
    """
    def validate(self, text, pos):
        if text.strip() in ("", "—"):
            return (QValidator.Acceptable, text, pos)
        return super().validate(text, pos)

    def valueFromText(self, text):
        digits = re.sub(r"\D", "", text or "")
        return int(digits) if digits else 0

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
_SAVE    = _HERE / ".mh_autosave.json"
_DEF_DIR = str(Path.home() / "Downloads" / "MyHeritage_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
try:
    import myheritage_scraper as _scraper
    _SCRAPER_OK = True
    SITE_OPTIONS = list(_scraper.SITE_PRESETS.keys())
except ImportError:
    _SCRAPER_OK  = False
    SITE_OPTIONS = ["Israel (.co.il)", "English (.com EN)", "Russian (.com RU)"]

FILTER_OPTIONS = ["All Records", "Historical Records", "Family Trees"]
GENDER_OPTIONS = ["Any", "Male", "Female"]

# Record-type radio options (refine by record type). The GUI is ALWAYS English;
# the scraper translates these canonical values to the SITE language (RU/EN/HE)
# chosen in the "Site / Language" selector before sending them to MyHeritage.
RECORD_TYPE_OPTIONS = ["All records", "Historical records", "Family trees"]

# Category filter (restrict search by category) — canonical English labels.
CATEGORY_OPTIONS = [
    "All collections",
    "Public records",
    "Schools & universities",
    "Census & voter lists",
    "Stories, memories & histories",
    "Birth, marriage & death",
    "Immigration & travel",
    "Books & publications",
    "Photos",
    "Family trees",
    "Newspapers",
    "Government, land, court & wills",
]

# The full «narrow by category» tree (site labels, nested) — built by
# myheritage_filter_crawler.py. {label: {count, children:{…}}}.
def _load_categories():
    try:
        return json.loads(
            (Path(__file__).resolve().parent.parent / "config"
             / "myheritage_categories.json").read_text("utf-8"))
    except Exception:
        return {}
CATEGORIES = _load_categories()

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b0b8c8;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f8f9fb;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#2a4a7f;background:#f8f9fb;}
QLineEdit,QComboBox,QSpinBox{padding:4px 6px;border:1px solid #c0c8d8;
  border-radius:4px;background:white;min-height:22px;}
QLineEdit:focus,QComboBox:focus,QSpinBox:focus{border:1px solid #4472c4;}
QPushButton{padding:5px 14px;border-radius:4px;border:1px solid #b0b8c8;background:#eef1f7;}
QPushButton:hover{background:#dde3f0;}
QPushButton:pressed{background:#ccd3e8;}
QPushButton#startBtn{background:#2a4a7f;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#3a5a9f;}
QPushButton#startBtn:disabled{background:#9aabcc;}
QPushButton#eyeBtn{border:none;background:transparent;padding:0;}
QPushButton#eyeBtn:hover{background:#e0e4ef;border-radius:3px;}
QPushButton#advBtn{text-align:left;border:none;background:transparent;
  color:#2a4a7f;font-weight:bold;font-size:11px;padding:2px 0;}
QPushButton#advBtn:hover{color:#4472c4;}
QProgressBar{border:1px solid #c0c8d8;border-radius:4px;
  text-align:center;min-height:18px;}
QProgressBar::chunk{background:#4472c4;border-radius:3px;}
QLabel#note{color:#777;font-size:10px;font-style:italic;}
QLabel#site_note{color:#e07000;font-size:10px;}
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
    progress    = Signal(int, str)
    finished    = Signal(dict)
    request_2fa = Signal()          # emitted when scraper needs the 2FA code
    request_file = Signal(str)      # emitted when output files already exist

    def __init__(self, payload, window):
        super().__init__()
        self.payload  = payload
        self._window  = window
        self._code    = None
        self._code_ev = threading.Event()
        self._file_choice = "overwrite"
        self._file_ev = threading.Event()

    def provide_code(self, code: str):
        """Called from main thread after user enters code."""
        self._code = code
        self._code_ev.set()

    def provide_file_choice(self, choice: str):
        """Called from main thread after the user chooses overwrite/append/skip."""
        self._file_choice = choice
        self._file_ev.set()

    def run(self):
        import asyncio

        self.payload["progress"] = lambda v, t: self.progress.emit(int(v), str(t))

        def ask_2fa():
            """
            Called from the scraper (async, in Worker thread).
            Emits request_2fa signal → main thread shows a dialog.
            Blocks until the main thread calls provide_code().
            """
            self._code    = None
            self._code_ev.clear()
            self.request_2fa.emit()
            self._code_ev.wait(timeout=120)   # wait up to 2 min
            return self._code or ""

        self.payload["ask_2fa_code"] = ask_2fa

        def ask_file_conflict(names):
            """Called from the scraper when output files already exist.
            Emits request_file → main thread shows a dialog; blocks for the
            choice ("overwrite" / "append" / "skip")."""
            self._file_choice = "overwrite"
            self._file_ev.clear()
            self.request_file.emit("\n".join(names))
            self._file_ev.wait(timeout=300)   # wait up to 5 min for the choice
            return self._file_choice or "overwrite"

        self.payload["ask_file_conflict"] = ask_file_conflict

        try:
            result = asyncio.run(_scraper.run_scraper(**self.payload))
        except Exception as exc:
            result = {"ok": False, "error": "exception",
                      "message": f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)

# ── Main window ───────────────────────────────────────────────────────────── #
class MyHeritageApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MyHeritage Search")
        self.setMinimumWidth(860)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._build_ui()
        self._load()

    def _match_btn(self, label, options):
        """Per-field dropdown («match ▾»). Each menu item is a REAL QCheckBox (a
        visible square ☐/☑), via QWidgetAction — not a checkmark-only QAction.
        `options`: (key, text, default). Stored in self._match_actions[key] as
        QCheckBox (same isChecked/setChecked/toggled, so save/load/payload work)."""
        btn = QToolButton()
        btn.setText(label)
        btn.setPopupMode(QToolButton.InstantPopup)
        btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(
            "QToolButton{padding:4px 10px;border:1px solid #9aa4b2;border-radius:5px;}"
            "QToolButton::menu-indicator{image:none;}")
        menu = QMenu(btn)
        for key, text, default in options:
            cb = QCheckBox(text)
            cb.setChecked(default)
            cb.setStyleSheet("QCheckBox{padding:5px 12px;}")
            cb.toggled.connect(self._save)
            wa = QWidgetAction(menu)
            wa.setDefaultWidget(cb)
            menu.addAction(wa)
            self._match_actions[key] = cb
        btn.setMenu(menu)
        return btn

    def _build_ui(self):
        self._match_actions = {}
        root = QWidget(); self.setCentralWidget(root)
        self._outer = QVBoxLayout(root)
        self._outer.setContentsMargins(18, 12, 18, 12)
        self._outer.setSpacing(8)

        # Logo + name
        self._outer.addLayout(
            make_header("MHlogo.png", "MyHeritage", color="#2a4a7f"))

        # ── Site / language selector ─────────────────────────────────────── #
        sg = QGroupBox("Site / Language")
        sl = QHBoxLayout(sg); sl.setSpacing(10)
        self.f_site = QComboBox()
        self.f_site.addItems(SITE_OPTIONS)
        idx = self.f_site.findText("Israel (.co.il)")
        if idx >= 0:
            self.f_site.setCurrentIndex(idx)
        sl.addWidget(QLabel("Site:")); sl.addWidget(self.f_site); sl.addStretch()
        self._outer.addWidget(sg)

        # ── Credentials ──────────────────────────────────────────────────── #
        cg = QGroupBox("Account credentials")
        cl = QVBoxLayout(cg); cl.setSpacing(6)

        mh_row = QHBoxLayout()
        self.f_email = QLineEdit(); self.f_email.setPlaceholderText("Email")
        self.f_pass  = PwdEdit();   self.f_pass.setPlaceholderText("Password")
        mh_row.addWidget(QLabel("Email:"))
        mh_row.addWidget(self.f_email, 2)
        mh_row.addWidget(QLabel("Password:")); mh_row.addWidget(self.f_pass, 2)
        cl.addLayout(mh_row)

        imap_row = QHBoxLayout()
        self.f_imap_pass = PwdEdit()
        self.f_imap_pass.setPlaceholderText(
            "Password of the email above — to read the 2FA code automatically")
        imap_row.addWidget(QLabel("Mail password:"))
        imap_row.addWidget(self.f_imap_pass, 3)
        cl.addLayout(imap_row)

        self._outer.addWidget(cg)

        # ── Basic search (always visible) ────────────────────────────────── #
        # Each row is a HBox: fixed-width label | stretching input | checkbox.
        # The stretching input absorbs the slack so the checkbox always stays
        # INSIDE the group border (the old QGridLayout let them spill past it).
        ms = QGroupBox("Basic Search")
        mv = QVBoxLayout(ms); mv.setSpacing(10); mv.setContentsMargins(12, 10, 12, 10)
        self.f_first   = QLineEdit(); self.f_first.setPlaceholderText("e.g.  Ivan Ivanovich")
        self.f_surname = QLineEdit(); self.f_surname.setPlaceholderText("e.g.  Ivanov")
        _LBLW = 170
        # Each field has its own «match ▾» dropdown (with visible checkbox squares).
        nm_btn = self._match_btn("match ▾", [
            ("name_strict",     "Strict — exact name", False),
            ("name_variants",   "Spelling variants",   True),
            ("name_initials",   "Initial matching",    True),
            ("name_startswith", "Starts with letters", False),
        ])
        sn_btn = self._match_btn("match ▾", [
            ("surname_strict",  "Strict — exact surname", False),
        ])
        r1 = QHBoxLayout(); r1.setSpacing(8)
        _l1 = QLabel("First name / Patronymic:"); _l1.setFixedWidth(_LBLW)
        r1.addWidget(_l1); r1.addWidget(self.f_first, 1); r1.addWidget(nm_btn)
        r2 = QHBoxLayout(); r2.setSpacing(8)
        _l2 = QLabel("Surname:"); _l2.setFixedWidth(_LBLW)
        r2.addWidget(_l2); r2.addWidget(self.f_surname, 1); r2.addWidget(sn_btn)
        mv.addLayout(r1); mv.addLayout(r2)
        self._outer.addWidget(ms)

        # ── Advanced toggle ──────────────────────────────────────────────── #
        self._adv_btn = QPushButton("▶   Advanced Search")
        self._adv_btn.setObjectName("advBtn")
        self._adv_btn.setCheckable(True)
        self._adv_btn.toggled.connect(self._toggle_adv)
        self._outer.addWidget(self._adv_btn)

        # Advanced panel — wrapped in a QScrollArea (like FamilySearch)
        self._adv = QGroupBox()
        af = QFormLayout(self._adv); af.setSpacing(8)
        self.f_by  = _YearSpin(); self.f_by.setRange(0,2025); self.f_by.setValue(0)
        self.f_by.setSpecialValueText("—"); self.f_by.setFixedWidth(110)
        self.f_bp  = QLineEdit(); self.f_bp.setPlaceholderText("City, country…")
        self.f_fa  = QLineEdit(); self.f_fa.setPlaceholderText("Father's name")
        self.f_fa_last = QLineEdit(); self.f_fa_last.setPlaceholderText("Father's surname")
        self.f_mo  = QLineEdit(); self.f_mo.setPlaceholderText("Mother's name")
        self.f_mo_last = QLineEdit(); self.f_mo_last.setPlaceholderText("Mother's surname")
        self.f_sp  = QLineEdit(); self.f_sp.setPlaceholderText("Spouse's name")
        self.f_sp_last = QLineEdit(); self.f_sp_last.setPlaceholderText("Spouse's surname")
        self.f_dy  = _YearSpin(); self.f_dy.setRange(0,2025); self.f_dy.setValue(0)
        self.f_dy.setSpecialValueText("—"); self.f_dy.setFixedWidth(110)
        self.f_dp  = QLineEdit(); self.f_dp.setPlaceholderText("City, country…")
        self.f_res = QLineEdit(); self.f_res.setPlaceholderText("City / region")
        self.f_mil = QLineEdit(); self.f_mil.setPlaceholderText("Unit, branch…")
        self.f_imm = QLineEdit(); self.f_imm.setPlaceholderText("Destination / year")
        self.f_kw  = QLineEdit(); self.f_kw.setPlaceholderText("Any keywords")
        self.f_gen = QComboBox(); self.f_gen.addItems(GENDER_OPTIONS)
        self.f_ex  = QCheckBox("Exact match for all parameters")
        # Per-field match options: birth-year tolerance (combo) + a visible
        # «place must match» checkbox next to the place field.
        self.f_ym  = QComboBox()
        self.f_ym.addItems(["—", "Exact", "± 1", "± 2", "± 5", "± 10", "± 20"])
        self.f_ym.setCurrentText("± 5")
        self.f_ym.setFixedWidth(110)
        pl_btn = self._match_btn("match ▾", [
            ("place_match", "Location must match", False),
        ])
        _byrow = QHBoxLayout(); _byrow.setSpacing(8)
        _byrow.addWidget(self.f_by); _byrow.addWidget(QLabel("match:"))
        _byrow.addWidget(self.f_ym); _byrow.addStretch()
        _bprow = QHBoxLayout(); _bprow.setSpacing(8)
        _bprow.addWidget(self.f_bp, 1); _bprow.addWidget(pl_btn)
        af.addRow("Birth year:",  _byrow)
        af.addRow("Birth place:", _bprow)
        af.addRow("Father (name):",    self.f_fa)
        af.addRow("Father (surname):", self.f_fa_last)
        af.addRow("Mother (name):",    self.f_mo)
        af.addRow("Mother (surname):", self.f_mo_last)
        af.addRow("Spouse (name):",    self.f_sp)
        af.addRow("Spouse (surname):", self.f_sp_last)
        af.addRow("Death year:",  self.f_dy)
        af.addRow("Death place:", self.f_dp)
        af.addRow("Residence:",   self.f_res)
        af.addRow("Military:",    self.f_mil)
        af.addRow("Immigration:", self.f_imm)
        af.addRow("Keywords:",    self.f_kw)
        af.addRow("Gender:",      self.f_gen)
        af.addRow("",             self.f_ex)

        self._adv_scroll = QScrollArea()
        self._adv_scroll.setWidget(self._adv)
        self._adv_scroll.setWidgetResizable(True)
        self._adv_scroll.setFrameShape(QFrame.NoFrame)
        self._adv_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._adv_scroll.setVisible(False)
        self._outer.addWidget(self._adv_scroll)

        # ── Refine by record type (radio buttons) ────────────────────────── #
        fg = QGroupBox("Refine by record type")
        fl = QHBoxLayout(fg); fl.setSpacing(12)
        self._rt_group = QButtonGroup(self)
        self._rt_buttons = {}
        for opt in RECORD_TYPE_OPTIONS:
            rb = QRadioButton(opt)
            self._rt_group.addButton(rb)
            self._rt_buttons[opt] = rb
            fl.addWidget(rb)
        self._rt_buttons[RECORD_TYPE_OPTIONS[0]].setChecked(True)
        fl.addStretch()
        self._outer.addWidget(fg)

        # ── Restrict search by category — collapsible tree (arrow opens it) ── #
        self._cat_btn = QPushButton("▶   Narrow down by category")
        self._cat_btn.setObjectName("advBtn")
        self._cat_btn.setCheckable(True)
        self._cat_btn.toggled.connect(self._toggle_cat)
        self._outer.addWidget(self._cat_btn)

        self._cat_host = QWidget()
        chl = QVBoxLayout(self._cat_host)
        chl.setContentsMargins(8, 0, 0, 0); chl.setSpacing(1)
        self._cat_checks = []                      # [(checkbox, label)]
        for name, data in (CATEGORIES or {}).items():
            self._cat_node(name, data, chl, 0)
        chl.addStretch()
        self._cat_scroll = QScrollArea()
        self._cat_scroll.setWidget(self._cat_host)
        self._cat_scroll.setWidgetResizable(True)
        self._cat_scroll.setFrameShape(QFrame.NoFrame)
        self._cat_scroll.setVisible(False)
        self._outer.addWidget(self._cat_scroll)

        # ── Output ───────────────────────────────────────────────────────── #
        og = QGroupBox("Output")
        ol = QVBoxLayout(og); ol.setSpacing(6)
        fr = QHBoxLayout()
        self.f_docx = QCheckBox("Word (.docx)"); self.f_docx.setChecked(True)
        self.f_xlsx = QCheckBox("Excel (.xlsx)"); self.f_xlsx.setChecked(True)
        fr.addWidget(self.f_docx); fr.addWidget(self.f_xlsx); fr.addStretch()
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
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        self._outer.addLayout(br)
        self._outer.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        # Autosave wiring
        for w in self._all_fields():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QComboBox): w.currentTextChanged.connect(self._save)
            elif isinstance(w, QSpinBox):  w.valueChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)
        # «Refine by record type» radios weren't autosaved → the choice was
        # forgotten between runs. Persist it too.
        self._rt_group.buttonToggled.connect(self._save)

        self._fit()

    # ── Advanced toggle ───────────────────────────────────────────────────── #
    def _toggle_adv(self, on):
        self._adv_scroll.setVisible(on)
        if on:
            screen_h = QApplication.primaryScreen().availableGeometry().height()
            self._adv_scroll.setMaximumHeight(int(screen_h * 0.55))
        self._adv_btn.setText(("▼" if on else "▶") + "   Advanced Search")
        self._fit()

    # ── Category tree (Restrict search by category) ───────────────────────── #
    def _toggle_cat(self, on):
        self._cat_scroll.setVisible(on)
        if on:
            screen_h = QApplication.primaryScreen().availableGeometry().height()
            self._cat_scroll.setMaximumHeight(int(screen_h * 0.45))
        self._cat_btn.setText(("▼" if on else "▶") + "   Narrow down by category")
        self._fit()

    def _cat_node(self, name, data, parent_layout, depth):
        """One category row: optional ▶ arrow (if it has children) + checkbox.
        Children are built lazily the first time the arrow is opened."""
        data = data or {}
        children = data.get("children") or {}
        count = (data.get("count") or "").strip()
        row = QHBoxLayout()
        row.setContentsMargins(depth * 16, 0, 0, 0); row.setSpacing(4)
        if children:
            arrow = QPushButton("▶"); arrow.setObjectName("advBtn")
            arrow.setFixedWidth(16); arrow.setCheckable(True)
            row.addWidget(arrow)
        cb = QCheckBox(name + (f"   ({count})" if count else ""))
        cb.stateChanged.connect(self._save)
        self._cat_checks.append((cb, name))
        row.addWidget(cb); row.addStretch()
        parent_layout.addLayout(row)
        if children:
            holder = QWidget(); hb = QVBoxLayout(holder)
            hb.setContentsMargins(0, 0, 0, 0); hb.setSpacing(1)
            holder.setVisible(False)
            parent_layout.addWidget(holder)
            built = {"done": False}

            def _toggle(on, _arrow=arrow, _hb=hb, _ch=children, _d=depth, _h=holder):
                _arrow.setText("▼" if on else "▶")
                if on and not built["done"]:
                    for cn, cd in _ch.items():
                        self._cat_node(cn, cd, _hb, _d + 1)
                    built["done"] = True
                _h.setVisible(on); self._fit()
            arrow.toggled.connect(_toggle)

    def _category_filters(self) -> list:
        """Ticked category labels (excluding «Все коллекции» = no narrowing)."""
        out = []
        for cb, name in self._cat_checks:
            if cb.isChecked() and name.strip().lower() not in (
                    "все коллекции", "all collections"):
                out.append(name)
        return out

    def _fit(self):
        # Unlock height, force the layout to RECOMPUTE synchronously, then lock to
        # the fresh size. Without invalidate()/activate() the sizeHint right after
        # collapsing the Advanced panel is still the old (taller) value, so the
        # window keeps an empty «подвал» whose extra height inflates the groups.
        self.setMinimumHeight(0)
        self.setMaximumHeight(16777215)
        self._outer.invalidate()
        self._outer.activate()
        h = self.sizeHint().height()
        self.resize(self.width(), h)
        self.setFixedHeight(h)

    def _record_type(self) -> str:
        for opt, rb in self._rt_buttons.items():
            if rb.isChecked():
                return opt
        return RECORD_TYPE_OPTIONS[0]

    def _year_match(self) -> str:
        """Combo text → scraper value: «—»→«» (don't set), «Exact»→«exact»,
        «± 5»→«5»."""
        t = self.f_ym.currentText().lower()
        if t.strip() in ("—", "-", ""):
            return ""
        if "exact" in t or "точ" in t:
            return "exact"
        m = "".join(ch for ch in t if ch.isdigit())
        return m or ""

    # ── Field list ────────────────────────────────────────────────────────── #
    def _all_fields(self):
        return [self.f_site, self.f_email, self.f_pass, self.f_imap_pass,
                self.f_first, self.f_surname,
                self.f_by, self.f_ym, self.f_bp,
                self.f_fa, self.f_fa_last, self.f_mo, self.f_mo_last,
                self.f_sp, self.f_sp_last,
                self.f_dy, self.f_dp, self.f_res, self.f_mil, self.f_imm,
                self.f_kw, self.f_gen, self.f_ex,
                self.f_folder, self.f_docx, self.f_xlsx]

    # ── Autosave ──────────────────────────────────────────────────────────── #
    def _save(self, *_):
        d = {
            "site":          self.f_site.currentText(),
            "email":         self.f_email.text(),
            "password":      self.f_pass.text(),
            "imap_password": self.f_imap_pass.text(),
            "first_name":    self.f_first.text(),
            "surname":       self.f_surname.text(),
            "name_strict":     self._match_actions["name_strict"].isChecked(),
            "name_variants":   self._match_actions["name_variants"].isChecked(),
            "name_initials":   self._match_actions["name_initials"].isChecked(),
            "name_startswith": self._match_actions["name_startswith"].isChecked(),
            "surname_strict":  self._match_actions["surname_strict"].isChecked(),
            "year_match":      self.f_ym.currentText(),
            "place_match":     self._match_actions["place_match"].isChecked(),
            "birth_year":    self.f_by.value(),
            "birth_place":   self.f_bp.text(),
            "father":        self.f_fa.text(),
            "father_last":   self.f_fa_last.text(),
            "mother":        self.f_mo.text(),
            "mother_last":   self.f_mo_last.text(),
            "spouse":        self.f_sp.text(),
            "spouse_last":   self.f_sp_last.text(),
            "death_year":    self.f_dy.value(),
            "death_place":   self.f_dp.text(),
            "residence":     self.f_res.text(),
            "military":      self.f_mil.text(),
            "immigration":   self.f_imm.text(),
            "keywords":      self.f_kw.text(),
            "gender":        self.f_gen.currentText(),
            "exact_match":   self.f_ex.isChecked(),
            "record_type":   self._record_type(),
            "categories":    self._category_filters(),
            "output_folder": self.f_folder.text(),
            "fmt_docx":      self.f_docx.isChecked(),
            "fmt_xlsx":      self.f_xlsx.isChecked(),
            "adv_open":      self._adv_btn.isChecked(),
            "cat_open":      self._cat_btn.isChecked(),
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
        def s(w, k):
            if k not in d:
                return
            v = d[k]
            if   isinstance(w, QLineEdit): w.setText(str(v))
            elif isinstance(w, QSpinBox):  w.setValue(int(v) if v else 0)
            elif isinstance(w, QComboBox):
                i = w.findText(str(v))
                if i >= 0: w.setCurrentIndex(i)
            elif isinstance(w, QCheckBox): w.setChecked(bool(v))
        s(self.f_site,   "site")
        s(self.f_email,     "email");    s(self.f_pass,      "password")
        s(self.f_imap_pass, "imap_password")
        s(self.f_first,  "first_name"); s(self.f_surname, "surname")
        for _k in ("name_strict", "name_variants", "name_initials",
                   "name_startswith", "surname_strict", "place_match"):
            if _k in d and _k in self._match_actions:
                self._match_actions[_k].setChecked(bool(d[_k]))
        s(self.f_ym, "year_match")
        s(self.f_by,     "birth_year"); s(self.f_bp,  "birth_place")
        s(self.f_fa,     "father");     s(self.f_fa_last, "father_last")
        s(self.f_mo,     "mother");     s(self.f_mo_last, "mother_last")
        s(self.f_sp,     "spouse");     s(self.f_sp_last, "spouse_last")
        s(self.f_dy,     "death_year")
        s(self.f_dp,     "death_place"); s(self.f_res,"residence")
        s(self.f_mil,    "military");   s(self.f_imm, "immigration")
        s(self.f_kw,     "keywords");   s(self.f_gen, "gender")
        s(self.f_ex,     "exact_match")
        cats = set(d.get("categories") or [])
        for cb, name in self._cat_checks:
            cb.setChecked(name in cats)
        rt = d.get("record_type")
        if rt in self._rt_buttons:
            self._rt_buttons[rt].setChecked(True)
        s(self.f_folder, "output_folder")
        s(self.f_docx,   "fmt_docx");  s(self.f_xlsx, "fmt_xlsx")
        if d.get("adv_open"):
            self._adv_btn.setChecked(True)
        if d.get("cat_open"):
            self._cat_btn.setChecked(True)

    # ── Helpers ───────────────────────────────────────────────────────────── #
    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Select output folder",
                                              self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    def _fmt(self):
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _payload(self):
        return {
            "site_preset":   self.f_site.currentText(),
            "first_name":    self.f_first.text().strip(),
            "surname":       self.f_surname.text().strip(),
            "name_strict":     self._match_actions["name_strict"].isChecked(),
            "name_variants":   self._match_actions["name_variants"].isChecked(),
            "name_initials":   self._match_actions["name_initials"].isChecked(),
            "name_startswith": self._match_actions["name_startswith"].isChecked(),
            "surname_strict":  self._match_actions["surname_strict"].isChecked(),
            "year_match":      self._year_match(),
            "place_match":     self._match_actions["place_match"].isChecked(),
            "birth_year":    str(self.f_by.value()) if self.f_by.value() else "",
            "birth_place":   self.f_bp.text().strip(),
            "father":        self.f_fa.text().strip(),
            "father_last":   self.f_fa_last.text().strip(),
            "mother":        self.f_mo.text().strip(),
            "mother_last":   self.f_mo_last.text().strip(),
            "spouse":        self.f_sp.text().strip(),
            "spouse_last":   self.f_sp_last.text().strip(),
            "death_year":    str(self.f_dy.value()) if self.f_dy.value() else "",
            "death_place":   self.f_dp.text().strip(),
            "residence":     self.f_res.text().strip(),
            "military":      self.f_mil.text().strip(),
            "immigration":   self.f_imm.text().strip(),
            "keywords":      self.f_kw.text().strip(),
            "gender":        self.f_gen.currentText(),
            "exact_match":   self.f_ex.isChecked(),
            "record_type":   self._record_type(),
            "categories":    self._category_filters(),   # restrict-by-category passes
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "email":         self.f_email.text().strip() or None,
            "password":      self.f_pass.text() or None,
            "imap_password": self.f_imap_pass.text() or None,
            "log":           print,
            "cancel_event": getattr(self, "_cancel_ev", None),
            # ask_2fa_code injected by Worker
        }

    def _validate(self):
        if not self.f_first.text().strip() and not self.f_surname.text().strip():
            QMessageBox.warning(self, "Nothing to search",
                                "Please enter at least a first name or surname.")
            return False
        if not self.f_docx.isChecked() and not self.f_xlsx.isChecked():
            QMessageBox.warning(self, "No output format",
                                "Select at least one output format.")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Scraper not found",
                                 "myheritage_scraper.py could not be imported.\n"
                                 "Place it in the project root directory.")
            return False
        return True

    # ── 2FA dialog (called from main thread via signal) ───────────────────── #
    def _show_2fa_dialog(self):
        """
        Show a modal input dialog asking for the verification code.
        The code is then passed back to the Worker via provide_code().
        """
        code, ok = QInputDialog.getText(
            self,
            "Two-Factor Authentication",
            "A verification code was sent to your email.\n\n"
            "Please check your inbox (e.g. Yandex Mail),\n"
            "copy the code, and paste it here:",
            QLineEdit.Normal,
            "",
        )
        if ok and code.strip():
            self._worker.provide_code(code.strip())
        else:
            self._worker.provide_code("")   # cancelled → scraper will abort

    # ── File-conflict dialog (existing output files) ──────────────────────── #
    def _show_file_conflict_dialog(self, names: str):
        """Existing Word/Excel files were found — ask what to do."""
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
        self._worker.provide_file_choice(choice)

    # ── Start / finish ────────────────────────────────────────────────────── #
    def _start(self):
        if not self._validate():
            return
        self._cancel_ev = threading.Event()
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.pbar.setValue(0)
        self.stlbl.setText("Starting…")
        self._worker = Worker(self._payload(), self)
        self._worker.progress.connect(
            lambda v, t: (self.pbar.setValue(v), self.stlbl.setText(t)))
        self._worker.finished.connect(self._done)
        self._worker.request_2fa.connect(self._show_2fa_dialog)
        self._worker.request_file.connect(self._show_file_conflict_dialog)
        self._worker.start()

    def _done(self, r):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        if r.get("ok"):
            n = r.get("n_records", 0)
            parts = (["Word"] if r.get("docx_count") else []) + \
                    (["Excel"] if r.get("xlsx_path") else [])
            msg = f"{n} record(s) saved"
            if parts: msg += " → " + " + ".join(parts)
            if r.get("output_folder"): msg += f"\n\nFolder:\n{r['output_folder']}"
            QMessageBox.information(self, "Done", msg)
            self.stlbl.setText("Done.")
        else:
            QMessageBox.critical(self, "Error",
                f"Search failed.\n\n{r.get('message','')}\n\nCheck terminal for details.")
            self.stlbl.setText("Error — see terminal.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MyHeritageApp()
    w.show()
    sys.exit(app.exec())
