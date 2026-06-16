"""
gui/pamyat.py — «Память народа» (pamyat-naroda.ru) search window.

GUI is always English; a Site/Language selector (ru/en) picks which language
version of the site to search. Goes straight to the advanced search, collects
FIO-matching people, opens each, copies every document (with archive info) to
Word, saves photos/images.
"""

import asyncio, json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox,
    QFileDialog, QProgressBar, QMessageBox, QApplication, QGroupBox,
    QScrollArea, QFrame,
)
from PySide6.QtCore import QThread, Signal, Qt
from gui._app_icon import app_icon, make_header, make_cancel_button

_HERE   = Path(__file__).resolve().parent
_ROOT   = _HERE.parent
_CONFIG = _ROOT / "config"
_SAVE   = _HERE / ".pamyat_autosave.json"
_DEF_DIR = str(Path.home() / "Downloads" / "Pamyat_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    import pamyat_scraper as _scraper
    _SCRAPER_OK = True
    BRANCH = _scraper.BRANCH
    AWARDS = _scraper.AWARDS
except ImportError:
    _SCRAPER_OK = False
    BRANCH = {"": ""}
    AWARDS = {"": ""}

from gui.gwar import _FacetChecks            # shared searchable checkbox widget


def _load_pm_sources():
    """«Источники информации» groups (EN label ↔ RU site value). GUI shows EN."""
    try:
        return json.loads((_CONFIG / "pamyat_sources.json")
                          .read_text("utf-8")).get("groups", [])
    except Exception:
        return []
PM_SOURCES = _load_pm_sources()

# Result-page filter tags (EN label ↔ RU site text). Order as on the site.
PM_TAGS = [
    ("All results",            "Все результаты"),
    ("Commanders",             "Командующие"),
    ("Awarded",                "Награжденные"),
    ("Killed & missing",       "Погибшие и пропавшие без вести"),
    ("Books of Memory",        "Книги Памяти"),
    ("Personnel information",   "Сведения о личном составе"),
]

SITE_LANG = {"Russian (ru)": "ru", "English (en)": "en"}

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b0b8c8;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f8f9fb;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#b03030;background:#f8f9fb;}
QLineEdit,QComboBox{padding:4px 6px;border:1px solid #c0c8d8;border-radius:4px;
  background:white;}
QLineEdit:focus{border:1px solid #d04040;}
QPushButton{padding:5px 14px;border-radius:4px;border:1px solid #b0b8c8;
  background:#eef1f7;}
QPushButton:hover{background:#dde3f0;}
QPushButton#startBtn{background:#b71c1c;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#c62828;}
QPushButton#startBtn:disabled{background:#d9a5a5;}
QProgressBar{border:1px solid #c0c8d8;border-radius:4px;text-align:center;
  min-height:18px;}
QProgressBar::chunk{background:#b71c1c;border-radius:3px;}
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


class PamyatApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pamyat Naroda")
        self.setMinimumWidth(760)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._build_ui()
        self._load()
        # Initial fit + every time the layout settles after show.
        self._fit()
        from PySide6.QtCore import QTimer
        for _ms in (0, 120, 300, 600):
            QTimer.singleShot(_ms, self._fit)

    def _fit(self):
        """Dynamic sizing. Window grows to the form's natural width/height when the
        screen has room (no scroll). When the screen is too short, height is capped
        and the vertical scroll bar appears. Horizontal scroll is NEVER shown:
        the window widens to the form's natural width, capped to the screen."""
        if not hasattr(self, "_body"):
            return
        scr = QApplication.primaryScreen().availableGeometry()
        # Unlock everything so the layout reports honest hints.
        self.setMinimumHeight(0); self.setMaximumHeight(16777215)
        self._scroll.setMinimumHeight(0); self._scroll.setMaximumHeight(16777215)
        self._body.adjustSize()
        body_w = self._body.sizeHint().width()
        body_h = self._body.sizeHint().height()
        # Settle the layout once to measure non-scroll chrome (header + groupboxes
        # + buttons + footer + margins).
        self.adjustSize()
        chrome_h = max(150, self.height() - self._scroll.height())
        target_w = min(max(self.minimumWidth(), body_w + 24), scr.width() - 16)
        avail_h  = scr.height() - 16
        if chrome_h + body_h <= avail_h:                  # everything fits
            scroll_h = body_h + 4
            target_h = chrome_h + scroll_h
        else:                                             # need vertical scroll
            target_h = avail_h
            scroll_h = max(120, target_h - chrome_h)
        self._scroll.setMinimumHeight(scroll_h)
        self._scroll.setMaximumHeight(scroll_h)
        self.resize(target_w, target_h)
        self.move(scr.x() + max(0, (scr.width() - self.width()) // 2),
                  scr.y() + 8)

    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        main = QVBoxLayout(root)
        main.setContentsMargins(16, 12, 16, 12); main.setSpacing(8)

        main.addLayout(make_header("Pamatlogo.png", "Pamyat Naroda", color="#b71c1c"))
        main.addWidget(QLabel("WWII participant records (pamyat-naroda.ru) — "
                              "documents, awards, fate, scanned archives."))

        lg = QGroupBox("Site / Language")
        ll = QHBoxLayout(lg); ll.setSpacing(10)
        self.f_lang = QComboBox(); self.f_lang.addItems(list(SITE_LANG.keys()))
        ll.addWidget(QLabel("Search the site in:")); ll.addWidget(self.f_lang); ll.addStretch()
        main.addWidget(lg)

        # ── Scrollable form mirroring the site's advanced search ─────────────
        body = QWidget(); outer = QVBoxLayout(body)
        outer.setContentsMargins(0, 0, 8, 0); outer.setSpacing(10)

        # Top checkboxes (use_main_string / use_person / use_collection). Ignore-main
        # is ON by default: this form fills the ADVANCED fields, not the top string,
        # so the empty top string must be ignored (matches the proven scraper).
        self.f_ignore_main = QCheckBox("Ignore the main search string")
        self.f_ignore_main.setChecked(True)
        self.f_group       = QCheckBox("Group documents of one person")
        self.f_group.setChecked(True)
        self.f_victory     = QCheckBox("Search documents for the “Victory Lesson”")
        topcb = QHBoxLayout(); topcb.setSpacing(20)
        for cb in (self.f_ignore_main, self.f_group, self.f_victory):
            topcb.addWidget(cb)
        topcb.addStretch()
        outer.addLayout(topcb)

        # Person
        ng = QGroupBox("Person")
        ngl = QGridLayout(ng); ngl.setSpacing(6)
        self.f_last = QLineEdit(); self.f_first = QLineEdit(); self.f_mid = QLineEdit()
        self.f_byf = QLineEdit(); self.f_byf.setPlaceholderText("Date / year")
        self.f_birth_period = QCheckBox("Search by period")
        self.f_byt = QLineEdit(); self.f_byt.setPlaceholderText("End"); self.f_byt.setEnabled(False)
        self.f_birth_period.toggled.connect(self.f_byt.setEnabled)
        self.f_bplace = QLineEdit()
        self.f_rank = QLineEdit()
        self.f_branch = QComboBox(); self.f_branch.addItems(list(BRANCH.keys()))
        self.f_service = QLineEdit()
        self.f_call = QLineEdit(); self.f_ids = QLineEdit()
        ngl.addWidget(QLabel("Surname:"), 0, 0);    ngl.addWidget(self.f_last, 0, 1)
        ngl.addWidget(QLabel("First name:"), 0, 2); ngl.addWidget(self.f_first, 0, 3)
        ngl.addWidget(QLabel("Patronymic:"), 0, 4); ngl.addWidget(self.f_mid, 0, 5)
        ngl.addWidget(QLabel("Date of birth:"), 1, 0); ngl.addWidget(self.f_byf, 1, 1)
        ngl.addWidget(self.f_birth_period, 1, 2);   ngl.addWidget(self.f_byt, 1, 3)
        ngl.addWidget(QLabel("Place of birth:"), 1, 4); ngl.addWidget(self.f_bplace, 1, 5)
        ngl.addWidget(QLabel("Military rank:"), 2, 0); ngl.addWidget(self.f_rank, 2, 1)
        ngl.addWidget(QLabel("Branch of service:"), 2, 2); ngl.addWidget(self.f_branch, 2, 3)
        ngl.addWidget(QLabel("Place of service:"), 2, 4); ngl.addWidget(self.f_service, 2, 5)
        ngl.addWidget(QLabel("Place of conscription:"), 3, 0); ngl.addWidget(self.f_call, 3, 1)
        ngl.addWidget(QLabel("Record ID:"), 3, 2);  ngl.addWidget(self.f_ids, 3, 3)
        outer.addWidget(ng)

        # Departure / death info
        dg = QGroupBox("Search with departure / death info")
        dgl = QGridLayout(dg); dgl.setSpacing(6)
        self.f_dvf = QLineEdit(); self.f_dvf.setPlaceholderText("Date")
        self.f_dv_period = QCheckBox("Search by period")
        self.f_dvt = QLineEdit(); self.f_dvt.setPlaceholderText("End"); self.f_dvt.setEnabled(False)
        self.f_dv_period.toggled.connect(self.f_dvt.setEnabled)
        self.f_dvplace = QLineEdit()
        self.f_hospital = QLineEdit(); self.f_camp = QLineEdit(); self.f_capture = QLineEdit()
        dgl.addWidget(QLabel("Date of departure/death:"), 0, 0); dgl.addWidget(self.f_dvf, 0, 1)
        dgl.addWidget(self.f_dv_period, 0, 2);      dgl.addWidget(self.f_dvt, 0, 3)
        dgl.addWidget(QLabel("Place of departure/death:"), 0, 4); dgl.addWidget(self.f_dvplace, 0, 5)
        dgl.addWidget(QLabel("Hospital:"), 1, 0);   dgl.addWidget(self.f_hospital, 1, 1)
        dgl.addWidget(QLabel("POW camp:"), 1, 2);   dgl.addWidget(self.f_camp, 1, 3)
        dgl.addWidget(QLabel("Place of capture:"), 1, 4); dgl.addWidget(self.f_capture, 1, 5)
        outer.addWidget(dg)

        # Award info
        ag = QGroupBox("Search with award info")
        agl = QGridLayout(ag); agl.setSpacing(6)
        self.f_award = QComboBox(); self.f_award.addItems(list(AWARDS.keys()))
        self.f_award_doc = QLineEdit(); self.f_award_date = QLineEdit()
        agl.addWidget(QLabel("Award:"), 0, 0);      agl.addWidget(self.f_award, 0, 1)
        agl.addWidget(QLabel("Award document №:"), 0, 2); agl.addWidget(self.f_award_doc, 0, 3)
        agl.addWidget(QLabel("Award date:"), 0, 4); agl.addWidget(self.f_award_date, 0, 5)
        outer.addWidget(ag)

        # Archive details
        arg = QGroupBox("Archive details")
        argl = QGridLayout(arg); argl.setSpacing(6)
        self.f_fund = QLineEdit(); self.f_opis = QLineEdit(); self.f_delo = QLineEdit()
        argl.addWidget(QLabel("Fund №:"), 0, 0);    argl.addWidget(self.f_fund, 0, 1)
        argl.addWidget(QLabel("Inventory №:"), 0, 2); argl.addWidget(self.f_opis, 0, 3)
        argl.addWidget(QLabel("File №:"), 0, 4);    argl.addWidget(self.f_delo, 0, 5)
        outer.addWidget(arg)

        # Result filters (the site's coloured result tags)
        fg = QGroupBox("Result filters")
        fgl = QHBoxLayout(fg); fgl.setSpacing(10)
        self._tag_cbs = {}
        for en, ru in PM_TAGS:
            cb = QCheckBox(en); self._tag_cbs[ru] = cb
            cb.stateChanged.connect(self._save); fgl.addWidget(cb)
        fgl.addStretch()
        outer.addWidget(fg)

        # Information sources (collapsible) — English labels; tick to narrow.
        self._src_btn = QPushButton("▶  Information sources")
        self._src_btn.setObjectName("advBtn"); self._src_btn.setCheckable(True)
        self._src_btn.toggled.connect(self._toggle_src)
        outer.addWidget(self._src_btn)
        self._src_box = QWidget()
        sgl2 = QGridLayout(self._src_box); sgl2.setSpacing(8)
        self._src_fc = []
        for gi, grp in enumerate(PM_SOURCES):
            sgl2.addWidget(QLabel(grp.get("en", "") + ":"), 0, gi)
            fc = _FacetChecks(grp.get("items", []), searchable=False, scroll=False,
                              on_change=self._save)
            self._src_fc.append(fc)
            sgl2.addWidget(fc, 1, gi)
        self._src_box.setVisible(False)
        outer.addWidget(self._src_box)
        outer.addStretch()

        self._body   = body
        self._scroll = QScrollArea()
        self._scroll.setWidget(body)
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        # Horizontal scroll = «порнография» — NEVER show it. The window resizes
        # horizontally to the form's natural width (capped to the screen).
        # Vertical scroll appears ONLY when the screen is too short to fit it.
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        main.addWidget(self._scroll, 1)

        og = QGroupBox("Output (Word)")
        ol = QHBoxLayout(og); ol.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        ol.addWidget(QLabel("Save to:")); ol.addWidget(self.f_folder, 1); ol.addWidget(bb)
        main.addWidget(og)

        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        main.addWidget(self.pbar); main.addWidget(self.stlbl)

        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH"); self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        main.addLayout(br)
        main.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        for w in (self.f_last, self.f_first, self.f_mid, self.f_byf, self.f_byt,
                  self.f_bplace, self.f_rank, self.f_service, self.f_call, self.f_ids,
                  self.f_dvf, self.f_dvt, self.f_dvplace, self.f_hospital, self.f_camp,
                  self.f_capture, self.f_award_doc, self.f_award_date,
                  self.f_fund, self.f_opis, self.f_delo, self.f_folder):
            w.textChanged.connect(self._save)
        for cb in (self.f_lang, self.f_branch, self.f_award):
            cb.currentTextChanged.connect(self._save)
        for cb in (self.f_group, self.f_ignore_main, self.f_victory,
                   self.f_birth_period, self.f_dv_period):
            cb.stateChanged.connect(self._save)

    def _toggle_src(self, on):
        self._src_box.setVisible(on)
        self._src_btn.setText(("▼" if on else "▶") + "  Information sources")
        # The form just grew (or shrunk) — re-fit so the window expands to fit
        # the new content if the screen has room, or scrolls if it doesn't.
        from PySide6.QtCore import QTimer
        QTimer.singleShot(0, self._fit)

    def _src_checked(self):
        return [ru for fc in self._src_fc for ru in fc.checked()]

    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Output folder",
                                             self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    def _payload(self):
        return {
            "lang":        SITE_LANG.get(self.f_lang.currentText(), "ru"),
            "last_name":   self.f_last.text().strip(),
            "first_name":  self.f_first.text().strip(),
            "middle_name": self.f_mid.text().strip(),
            "birth_from":  self.f_byf.text().strip(),
            "birth_to":    self.f_byt.text().strip(),
            "birth_period": self.f_birth_period.isChecked(),
            "place_birth": self.f_bplace.text().strip(),
            "rank":        self.f_rank.text().strip(),
            "branch":      BRANCH.get(self.f_branch.currentText(), ""),
            "place_service": self.f_service.text().strip(),
            "place_conscription": self.f_call.text().strip(),
            "record_id":   self.f_ids.text().strip(),
            "ignore_main": self.f_ignore_main.isChecked(),
            "group_person": self.f_group.isChecked(),
            "victory_lesson": self.f_victory.isChecked(),
            # departure / death info
            "dep_from":    self.f_dvf.text().strip(),
            "dep_to":      self.f_dvt.text().strip(),
            "dep_period":  self.f_dv_period.isChecked(),
            "place_departure": self.f_dvplace.text().strip(),
            "hospital":    self.f_hospital.text().strip(),
            "pow_camp":    self.f_camp.text().strip(),
            "place_capture": self.f_capture.text().strip(),
            # award info
            "award":       AWARDS.get(self.f_award.currentText(), ""),
            "award_doc":   self.f_award_doc.text().strip(),
            "award_date":  self.f_award_date.text().strip(),
            # archive details
            "fund":        self.f_fund.text().strip(),
            "inventory":   self.f_opis.text().strip(),
            "file":        self.f_delo.text().strip(),
            # filters
            "result_tags": [ru for ru, cb in self._tag_cbs.items() if cb.isChecked()],
            "info_sources": self._src_checked(),     # RU site values, ticked only
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log":         print,
            "cancel_event": getattr(self, "_cancel_ev", None),
        }

    def _validate(self):
        p = self._payload()
        if not any(p[k] for k in ("last_name", "first_name", "place_service",
                                  "rank")):
            QMessageBox.warning(self, "Nothing to search",
                                "Enter at least a surname (or another field).")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Error", "pamyat_scraper.py not found.")
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
            msg = f"{n} person(s)"
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
            d = {"lang": self.f_lang.currentText(),
                 "last": self.f_last.text(), "first": self.f_first.text(),
                 "mid": self.f_mid.text(), "byf": self.f_byf.text(),
                 "byt": self.f_byt.text(), "bplace": self.f_bplace.text(),
                 "rank": self.f_rank.text(), "branch": self.f_branch.currentText(),
                 "service": self.f_service.text(), "call": self.f_call.text(),
                 "award": self.f_award.currentText(), "group": self.f_group.isChecked(),
                 "folder": self.f_folder.text(),
                 "ids": self.f_ids.text(), "ignore_main": self.f_ignore_main.isChecked(),
                 "victory": self.f_victory.isChecked(),
                 "birth_period": self.f_birth_period.isChecked(),
                 "dvf": self.f_dvf.text(), "dvt": self.f_dvt.text(),
                 "dv_period": self.f_dv_period.isChecked(),
                 "dvplace": self.f_dvplace.text(), "hospital": self.f_hospital.text(),
                 "camp": self.f_camp.text(), "capture": self.f_capture.text(),
                 "award_doc": self.f_award_doc.text(), "award_date": self.f_award_date.text(),
                 "fund": self.f_fund.text(), "opis": self.f_opis.text(),
                 "delo": self.f_delo.text(),
                 "tags": [ru for ru, cb in self._tag_cbs.items() if cb.isChecked()],
                 "info_sources": self._src_checked()}
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
        self.f_last.setText(d.get("last", "")); self.f_first.setText(d.get("first", ""))
        self.f_mid.setText(d.get("mid", "")); self.f_byf.setText(d.get("byf", ""))
        self.f_byt.setText(d.get("byt", "")); self.f_bplace.setText(d.get("bplace", ""))
        self.f_rank.setText(d.get("rank", "")); self.f_service.setText(d.get("service", ""))
        self.f_call.setText(d.get("call", "")); self.f_folder.setText(d.get("folder", _DEF_DIR))
        self.f_group.setChecked(d.get("group", True))
        self.f_ids.setText(d.get("ids", ""))
        self.f_ignore_main.setChecked(bool(d.get("ignore_main", False)))
        self.f_victory.setChecked(bool(d.get("victory", False)))
        self.f_birth_period.setChecked(bool(d.get("birth_period", False)))
        self.f_dvf.setText(d.get("dvf", "")); self.f_dvt.setText(d.get("dvt", ""))
        self.f_dv_period.setChecked(bool(d.get("dv_period", False)))
        self.f_dvplace.setText(d.get("dvplace", "")); self.f_hospital.setText(d.get("hospital", ""))
        self.f_camp.setText(d.get("camp", "")); self.f_capture.setText(d.get("capture", ""))
        self.f_award_doc.setText(d.get("award_doc", "")); self.f_award_date.setText(d.get("award_date", ""))
        self.f_fund.setText(d.get("fund", "")); self.f_opis.setText(d.get("opis", ""))
        self.f_delo.setText(d.get("delo", ""))
        for cb, key in ((self.f_lang, "lang"), (self.f_branch, "branch"),
                        (self.f_award, "award")):
            i = cb.findText(d.get(key, ""))
            if i >= 0:
                cb.setCurrentIndex(i)
        saved_tags = set(d.get("tags", []))
        for ru, cb in self._tag_cbs.items():
            cb.setChecked(ru in saved_tags)
        saved_src = set(d.get("info_sources", []))
        if saved_src:
            for fc in self._src_fc:
                fc.set_checked(saved_src)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = PamyatApp(); w.show()
    sys.exit(app.exec())
