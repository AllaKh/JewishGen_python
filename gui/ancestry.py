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

import json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QComboBox,
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

# Birth-year ± dropdown → year_range (birth_x=<N>-0-0)
YEAR_OPTIONS = [("Exact (this year)", 0), ("± 1 year", 1), ("± 2 years", 2),
                ("± 5 years", 5), ("± 10 years", 10)]
GENDER_OPTIONS = ["—", "Male", "Female"]
COLLECTION_OPTIONS = ["All Collections", "USA", "UK & Ireland", "Europe",
                      "Canada", "Australia & New Zealand", "Jewish Family History"]
FAMILY_TYPES = [("Father", "father"), ("Mother", "mother"),
                ("Spouse", "spouse"), ("Sibling", "sibling"), ("Child", "child")]
_FAM_LABEL = {k: l for l, k in FAMILY_TYPES}   # key → display label

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
        self.f_first_exact = QCheckBox("Exact")
        self.f_last_exact  = QCheckBox("Exact")
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

        # Family members — dynamic rows
        av.addWidget(_sechead("FAMILY MEMBERS  (add as many as you like)"))
        self._fam_box = QVBoxLayout(); self._fam_box.setSpacing(4)
        self._fam_containers = {}
        for label, key in FAMILY_TYPES:
            head = QHBoxLayout()
            add = QPushButton(f"+ Add {label}"); add.setObjectName("addBtn")
            add.clicked.connect(lambda _=False, k=key: (self._add_fam_row(k), self._save()))
            head.addWidget(add); head.addStretch()
            self._fam_box.addLayout(head)
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

        for w in self._static_widgets():
            if   isinstance(w, QLineEdit): w.textChanged.connect(self._save)
            elif isinstance(w, QComboBox): w.currentTextChanged.connect(self._save)
            elif isinstance(w, QCheckBox): w.stateChanged.connect(self._save)
        self._fit()

    # ── Dynamic family-member rows ────────────────────────────────────────── #
    def _add_fam_row(self, key, first="", last="", exact=False):
        box = self._fam_containers[key]
        row_w = QWidget()
        rl = QHBoxLayout(row_w); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(6)
        lbl = QLabel(_FAM_LABEL[key] + ":"); lbl.setFixedWidth(60)
        ff = QLineEdit(); ff.setPlaceholderText("First & Middle Name(s)"); ff.setText(first)
        lf = QLineEdit(); lf.setPlaceholderText("Last Name"); lf.setText(last)
        cb = QCheckBox("Exact"); cb.setChecked(bool(exact))
        rm = QPushButton("✕"); rm.setObjectName("rmBtn"); rm.setFixedWidth(24)
        rl.addWidget(lbl); rl.addWidget(ff, 2); rl.addWidget(lf, 2); rl.addWidget(cb)
        rl.addWidget(rm)
        rec = {"w": row_w, "first": ff, "last": lf, "exact": cb}
        self._fam[key].append(rec)
        box.addWidget(row_w)
        ff.textChanged.connect(self._save); lf.textChanged.connect(self._save)
        cb.stateChanged.connect(self._save)
        rm.clicked.connect(lambda _=False: self._remove_fam_row(key, rec))
        self._fit()
        return rec

    def _remove_fam_row(self, key, rec):
        try:
            self._fam[key].remove(rec)
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
            "first_exact": self.f_first_exact.isChecked(),
            "last_exact":  self.f_last_exact.isChecked(),
            "place_exact": self.f_place_exact.isChecked(),
            "keyword": self.f_keyword.text(), "gender": self.f_gender.currentText(),
            "race": self.f_race.text(), "collection": self.f_collection.currentText(),
            "hist": self.f_hist.isChecked(), "trees": self.f_trees.isChecked(),
            "stories": self.f_stories.isChecked(), "photos": self.f_photos.isChecked(),
            "output_folder": self.f_folder.text(),
            "fmt_docx": self.f_docx.isChecked(), "fmt_xlsx": self.f_xlsx.isChecked(),
            "adv_open": self._adv_btn.isChecked(),
            "fam": {k: [[r["first"].text(), r["last"].text(), r["exact"].isChecked()]
                        for r in rows] for k, rows in self._fam.items()},
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
        _s(self.f_first_exact, "first_exact"); _s(self.f_last_exact, "last_exact")
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
                    self._add_fam_row(key, *(r + ["", "", False])[:3])
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
                       "exact": r["exact"].isChecked()}
                      for r in rows
                      if r["first"].text().strip() or r["last"].text().strip()]
            if people:
                adv[key] = people
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
                "name":    self.f_first_exact.isChecked(),
                "surname": self.f_last_exact.isChecked(),
                "place":   self.f_place_exact.isChecked(),
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
