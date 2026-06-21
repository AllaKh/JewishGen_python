"""
gui/memsearch.py — Memsearch (memsearch.org) search window.

The GUI is ALWAYS English. A "Site / Language" selector chooses which language
version of memsearch.org to search (ru / en); the scraper translates the tab
names, field labels and dropdown options to that site language itself.

One search box + an entity-type tab (All types / People / Places / Objects /
Documents). The advanced fields for the chosen tab appear below it (stacked).
Each result card's external source page is opened and its full info copied to
Word (with photos when present).
"""

import asyncio, json, sys, threading
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QSpinBox,
    QFileDialog, QProgressBar, QMessageBox, QApplication,
    QGroupBox, QRadioButton, QButtonGroup, QStackedWidget, QSizePolicy,
    QCheckBox, QScrollArea, QFrame,
)
from PySide6.QtCore import QThread, Signal, Qt, QTimer
from PySide6.QtGui import QIntValidator
from gui._app_icon import (app_icon, make_header, make_cancel_button, autosave_path,
                           clamp_on_screen)

_HERE   = Path(__file__).resolve().parent
_ROOT   = _HERE.parent
_CONFIG = _ROOT / "config"
_SAVE   = autosave_path(".memsearch_autosave.json")
_DEF_DIR = str(Path.home() / "Downloads" / "Memsearch_results")

if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    import memsearch_scraper as _scraper
    _SCRAPER_OK = True
    TABS = _scraper.TABS
    REGION_TYPE = _scraper.REGION_TYPE
    PLACE_AMONG = _scraper.PLACE_AMONG
    OBJECT_AMONG = _scraper.OBJECT_AMONG
except ImportError:
    _SCRAPER_OK = False
    TABS = ["All types", "People", "Places", "Objects", "Documents"]
    REGION_TYPE = ["", "Any", "Place of birth", "Place of residence/repression"]
    PLACE_AMONG = ["", "Everywhere", "Burial places", "Places of imprisonment", "Monuments"]
    OBJECT_AMONG = ["", "Everywhere", "Photographs", "Museum objects",
                    "Accompanying texts", "Monuments"]

# Site / language selector → lang code sent to the scraper.
SITE_LANG = {"Russian (ru)": "ru", "English (en)": "en"}


def _load_sources():
    """The site's «ГДЕ ИЩЕМ?» (WHERE WE SEARCH) database list — {key, en, ru}."""
    try:
        return json.loads((_CONFIG / "memsearch_sources.json").read_text("utf-8"))
    except Exception:
        return []
SOURCES = _load_sources()

STYLE = """
QMainWindow,QWidget{font-family:Segoe UI,Arial,sans-serif;font-size:11px;}
QGroupBox{font-weight:bold;font-size:11px;border:1px solid #b0b8c8;
  border-radius:6px;margin-top:10px;padding-top:6px;background:#f8f9fb;}
QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;
  left:10px;padding:0 4px;color:#b03030;background:#f8f9fb;}
QLineEdit,QComboBox,QSpinBox{padding:4px 6px;border:1px solid #c0c8d8;
  border-radius:4px;background:white;}
QLineEdit:focus{border:1px solid #d04040;}
QPushButton{padding:5px 14px;border-radius:4px;border:1px solid #b0b8c8;
  background:#eef1f7;}
QPushButton:hover{background:#dde3f0;}
QPushButton#startBtn{background:#c0392b;color:white;font-weight:bold;
  font-size:13px;padding:8px 20px;border:none;border-radius:5px;}
QPushButton#startBtn:hover{background:#d04030;}
QPushButton#startBtn:disabled{background:#d9a59f;}
QRadioButton{font-weight:bold;}
QProgressBar{border:1px solid #c0c8d8;border-radius:4px;text-align:center;
  min-height:18px;}
QProgressBar::chunk{background:#c0392b;border-radius:3px;}
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


class MemsearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Memorial")
        self.setMinimumWidth(780)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._build_ui()
        self._load()
        self._fit()
        for _ms in (0, 130, 320):
            QTimer.singleShot(_ms, self._fit)

    def _fit(self):
        """Grow the window to the form's content (capped to the screen) so the form
        scrolls past the cap and the fixed bottom Start/Cancel bar is ALWAYS visible,
        on any resolution. No move — the launcher positions it (no jumping)."""
        if not hasattr(self, "_content_scroll"):
            return
        scr = (self.screen() or QApplication.primaryScreen()).availableGeometry()
        max_w, max_h = scr.width() - 16, scr.height() - 48
        self.setMinimumHeight(0); self.setMaximumHeight(16777215)
        cw = self._content_scroll.widget(); cw.adjustSize()
        bottom_h = self._bottom.sizeHint().height() if hasattr(self, "_bottom") else 0
        hint = cw.sizeHint().height() + bottom_h + 8
        self.resize(min(max(self.width(), self.minimumWidth()), max_w), min(hint, max_h))
        self.setMaximumHeight(max_h)
        clamp_on_screen(self)

    # ── UI ────────────────────────────────────────────────────────────────── #
    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        self._ol = QVBoxLayout(root)
        self._ol.setContentsMargins(0, 0, 0, 0); self._ol.setSpacing(0)

        # Scrollable form (so a tall form scrolls instead of pushing the buttons off
        # the screen) + a FIXED bottom bar that always shows Start/Cancel.
        self._content_scroll = QScrollArea()
        self._content_scroll.setWidgetResizable(True)
        self._content_scroll.setFrameShape(QFrame.NoFrame)
        self._content_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        content = QWidget(); outer = QVBoxLayout(content)
        outer.setContentsMargins(16, 12, 16, 12); outer.setSpacing(10)
        self._content_scroll.setWidget(content)
        self._ol.addWidget(self._content_scroll, 1)

        outer.addLayout(make_header("Memlogo.png", "Memorial", color="#c0392b"))
        outer.addWidget(QLabel(
            "Search across Soviet-repression victim databases (Memorial / GULAG.CZ)."))

        # Site / language
        lg = QGroupBox("Site / Language")
        ll = QHBoxLayout(lg); ll.setSpacing(10)
        self.f_lang = QComboBox(); self.f_lang.addItems(list(SITE_LANG.keys()))
        ll.addWidget(QLabel("Search the site in:")); ll.addWidget(self.f_lang); ll.addStretch()
        outer.addWidget(lg)

        # Query
        qg = QGroupBox("Search")
        ql = QVBoxLayout(qg); ql.setSpacing(6)
        self.f_query = QLineEdit(); self.f_query.setPlaceholderText("Surname, e.g. Shenderovich / Шендерович")
        ql.addWidget(self.f_query)
        outer.addWidget(qg)

        # Entity type tabs
        tg = QGroupBox("Record type")
        tl = QHBoxLayout(tg); tl.setSpacing(10)
        self._tab_group = QButtonGroup(self)
        self._tab_buttons = {}
        for i, t in enumerate(TABS):
            rb = QRadioButton(t)
            self._tab_group.addButton(rb, i)
            self._tab_buttons[t] = rb
            tl.addWidget(rb)
            rb.toggled.connect(lambda on, idx=i:
                               on and hasattr(self, "_stack")
                               and self._stack.setCurrentIndex(idx))
        tl.addStretch()
        outer.addWidget(tg)

        # Where to search (databases) — the site's «ГДЕ ИЩЕМ?» (WHERE WE SEARCH) list.
        wg = QGroupBox("Where to search (databases)")
        wl = QVBoxLayout(wg); wl.setSpacing(4)
        topr = QHBoxLayout()
        self.f_all_src = QCheckBox("Select all")
        self.f_all_src.setChecked(True)
        self.f_all_src.stateChanged.connect(self._toggle_all_sources)
        topr.addWidget(self.f_all_src); topr.addStretch()
        wl.addLayout(topr)
        host = QWidget(); hl = QVBoxLayout(host)
        hl.setContentsMargins(2, 2, 2, 2); hl.setSpacing(1)
        self._src_checks = []                                   # (checkbox, source key)
        for s in SOURCES:
            cb = QCheckBox((s.get("en") or s.get("ru") or s["key"]).replace("&", "&&"))
            cb.setChecked(True)
            if s.get("ru"):
                cb.setToolTip(s["ru"])
            cb.stateChanged.connect(self._on_src_toggle)
            self._src_checks.append((cb, s["key"]))
            hl.addWidget(cb)
        sc = QScrollArea(); sc.setWidget(host); sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.NoFrame)
        sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        sc.setMaximumHeight(170)
        wl.addWidget(sc)
        outer.addWidget(wg)

        # Per-tab filter fields. Shown only for a SPECIFIC record type — on «All types»
        # there are no type-specific filters, so the box is HIDDEN (no empty box).
        ag = QGroupBox("Filters for the selected type")
        self._adv_group = ag
        al = QVBoxLayout(ag)
        self._stack = QStackedWidget()
        self._stack.addWidget(self._page_all())
        self._stack.addWidget(self._page_people())
        self._stack.addWidget(self._page_places())
        self._stack.addWidget(self._page_objects())
        self._stack.addWidget(self._page_docs())
        al.addWidget(self._stack)
        # Collapse the stack to the CURRENT page's height (a QStackedWidget
        # otherwise reserves the tallest page → empty space for small tabs).
        self._stack.currentChanged.connect(self._shrink_stack)
        outer.addWidget(ag)
        for rb in self._tab_buttons.values():
            rb.toggled.connect(lambda *_: self._update_adv())
        self._tab_buttons[TABS[0]].setChecked(True)
        self._update_adv()

        # Output
        og = QGroupBox("Output (Word)")
        ol = QHBoxLayout(og); ol.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        ol.addWidget(QLabel("Save to:")); ol.addWidget(self.f_folder, 1); ol.addWidget(bb)
        outer.addWidget(og)

        # ── Fixed bottom bar (progress + Start/Cancel) — ALWAYS visible ──────
        self._bottom = QWidget()
        bl = QVBoxLayout(self._bottom)
        bl.setContentsMargins(16, 6, 16, 8); bl.setSpacing(6)
        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        bl.addWidget(self.pbar); bl.addWidget(self.stlbl)
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        bl.addLayout(br)
        bl.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))
        self._ol.addWidget(self._bottom, 0)

        # Autosave wiring
        for w in (self.f_query, self.f_last, self.f_first, self.f_patr,
                  self.f_region, self.f_place, self.f_object, self.f_doc,
                  self.f_folder, self.f_byear, self.f_repr_from, self.f_repr_to,
                  self.f_doc_from, self.f_doc_to):
            w.textChanged.connect(self._save)
        for cb in (self.f_lang, self.f_rtype, self.f_pamong, self.f_oamong):
            cb.currentTextChanged.connect(self._save)
        for rb in self._tab_buttons.values():
            rb.toggled.connect(self._save)

    def _page_all(self):
        w = QWidget(); l = QVBoxLayout(w)
        l.addWidget(QLabel("Searches all record types at once. To refine, pick a "
                           "specific type above."))
        return w

    def _year_edit(self):
        """A clearable year field (plain text — a QSpinBox can't be emptied, which is why
        the year «reappeared» when the user tried to delete it)."""
        e = QLineEdit(); e.setPlaceholderText("year")
        e.setMaxLength(4); e.setFixedWidth(72)
        e.setValidator(QIntValidator(0, 2026, self))
        return e

    def _wrap(self, layout):
        c = QWidget(); layout.setContentsMargins(0, 0, 0, 0); c.setLayout(layout); return c

    def _page_people(self):
        w = QWidget(); f = QFormLayout(w)
        self.f_last = QLineEdit(); self.f_first = QLineEdit(); self.f_patr = QLineEdit()
        self.f_byear = self._year_edit()
        self.f_region = QLineEdit()
        self.f_rtype = QComboBox(); self.f_rtype.addItems(REGION_TYPE)
        self.f_repr_from = self._year_edit(); self.f_repr_to = self._year_edit()
        f.addRow("Surname:", self.f_last)
        f.addRow("First name:", self.f_first)
        f.addRow("Patronymic:", self.f_patr)
        f.addRow("Year of birth:", self.f_byear)
        f.addRow("Region:", self.f_region)
        f.addRow("Type of region:", self.f_rtype)
        rr = QHBoxLayout()
        rr.addWidget(self.f_repr_from); rr.addWidget(QLabel("–"))
        rr.addWidget(self.f_repr_to); rr.addStretch()
        f.addRow("Date of repression:", self._wrap(rr))
        return w

    def _page_places(self):
        w = QWidget(); f = QFormLayout(w)
        self.f_place = QLineEdit()
        self.f_pamong = QComboBox(); self.f_pamong.addItems(PLACE_AMONG)
        f.addRow("Place name:", self.f_place)
        f.addRow("Search among:", self.f_pamong)
        return w

    def _page_objects(self):
        w = QWidget(); f = QFormLayout(w)
        self.f_object = QLineEdit()
        self.f_oamong = QComboBox(); self.f_oamong.addItems(OBJECT_AMONG)
        f.addRow("Object name:", self.f_object)
        f.addRow("Search among:", self.f_oamong)
        return w

    def _page_docs(self):
        w = QWidget(); f = QFormLayout(w)
        self.f_doc = QLineEdit()
        self.f_doc_from = self._year_edit(); self.f_doc_to = self._year_edit()
        f.addRow("Title of document:", self.f_doc)
        dr = QHBoxLayout()
        dr.addWidget(self.f_doc_from); dr.addWidget(QLabel("–"))
        dr.addWidget(self.f_doc_to); dr.addStretch()
        f.addRow("Date of issue:", self._wrap(dr))
        return w

    def _shrink_stack(self, idx):
        """Size the advanced-fields stack to the CURRENT page only, then shrink
        the window so there's no empty band under a small tab."""
        for i in range(self._stack.count()):
            p = self._stack.widget(i)
            p.setSizePolicy(QSizePolicy.Preferred,
                            QSizePolicy.Preferred if i == idx else QSizePolicy.Ignored)
        cur = self._stack.widget(idx)
        if cur:
            cur.adjustSize()
            self._stack.setFixedHeight(max(cur.sizeHint().height(), 10))
        QTimer.singleShot(0, self._fit)

    def _update_adv(self):
        """Show the per-type filter box only for a SPECIFIC record type — on «All types»
        there are no type-specific filters, so it's hidden (no empty «nightmare» box)."""
        if hasattr(self, "_adv_group"):
            self._adv_group.setVisible(self._tab() != TABS[0])
        QTimer.singleShot(0, self._fit)

    def _toggle_all_sources(self, state):
        on = bool(state)
        for cb, _k in self._src_checks:
            cb.blockSignals(True); cb.setChecked(on); cb.blockSignals(False)
        self._save()

    def _on_src_toggle(self, *_):
        all_on = bool(self._src_checks) and all(cb.isChecked() for cb, _k in self._src_checks)
        self.f_all_src.blockSignals(True); self.f_all_src.setChecked(all_on)
        self.f_all_src.blockSignals(False)
        self._save()

    def _selected_sources(self):
        """Checked source keys; [] when ALL are checked (= no filter, search every base)."""
        sel = [k for cb, k in self._src_checks if cb.isChecked()]
        return [] if len(sel) == len(self._src_checks) else sel

    # ── helpers ─────────────────────────────────────────────────────────────#
    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Output folder",
                                             self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    def _tab(self):
        for t, rb in self._tab_buttons.items():
            if rb.isChecked():
                return t
        return TABS[0]

    def _payload(self):
        return {
            "query":       self.f_query.text().strip(),
            "lang":        SITE_LANG.get(self.f_lang.currentText(), "ru"),
            "entity_tab":  self._tab(),
            "sources":     self._selected_sources(),
            "last_name":   self.f_last.text().strip(),
            "first_name":  self.f_first.text().strip(),
            "patronymic":  self.f_patr.text().strip(),
            "birth_year":  self.f_byear.text().strip(),
            "region":      self.f_region.text().strip(),
            "region_type": self.f_rtype.currentText().strip(),
            "repress_from": self.f_repr_from.text().strip(),
            "repress_to":  self.f_repr_to.text().strip(),
            "place_name":  self.f_place.text().strip(),
            "place_among": self.f_pamong.currentText().strip(),
            "object_name": self.f_object.text().strip(),
            "object_among": self.f_oamong.currentText().strip(),
            "doc_name":    self.f_doc.text().strip(),
            "doc_from":    self.f_doc_from.text().strip(),
            "doc_to":      self.f_doc_to.text().strip(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log":         print,
            "cancel_event": getattr(self, "_cancel_ev", None),
        }

    def _validate(self):
        p = self._payload()
        if not p["query"] and not any(
                p[k] for k in ("last_name", "first_name", "place_name",
                               "object_name", "doc_name")):
            QMessageBox.warning(self, "Nothing to search",
                                "Enter a surname or at least one field.")
            return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Error",
                                 "memsearch_scraper.py not found.")
            return False
        return True

    # ── run ─────────────────────────────────────────────────────────────────#
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
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Question)
        box.setWindowTitle("File already exists")
        box.setText(f"This file already exists:\n\n{names}\n\nWhat to do?")
        b_over = box.addButton("Overwrite", QMessageBox.DestructiveRole)
        b_app  = box.addButton("Append", QMessageBox.AcceptRole)
        b_skip = box.addButton("Skip", QMessageBox.RejectRole)
        box.setDefaultButton(b_app)
        box.exec()
        c = box.clickedButton()
        self._worker.provide_file_choice(
            "append" if c is b_app else "skip" if c is b_skip else "overwrite")

    def _done(self, r: dict):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        if r.get("ok"):
            n = r.get("n_records", 0)
            msg = f"{n} record(s)"
            if r.get("output_folder"):
                msg += f"\n\nFolder:\n{r['output_folder']}"
            QMessageBox.information(self, "Done", msg)
            self.stlbl.setText("Done.")
        else:
            QMessageBox.critical(self, "Error",
                                 r.get("message", "Failed — see terminal."))
            self.stlbl.setText("Error.")

    # ── autosave ────────────────────────────────────────────────────────────#
    def _save(self, *_):
        try:
            d = {
                "query": self.f_query.text(), "lang": self.f_lang.currentText(),
                "tab": self._tab(),
                "last": self.f_last.text(), "first": self.f_first.text(),
                "patr": self.f_patr.text(), "byear": self.f_byear.text(),
                "region": self.f_region.text(), "rtype": self.f_rtype.currentText(),
                "repr_from": self.f_repr_from.text(), "repr_to": self.f_repr_to.text(),
                "place": self.f_place.text(), "pamong": self.f_pamong.currentText(),
                "object": self.f_object.text(), "oamong": self.f_oamong.currentText(),
                "doc": self.f_doc.text(), "folder": self.f_folder.text(),
                "doc_from": self.f_doc_from.text(), "doc_to": self.f_doc_to.text(),
                "sources": [k for cb, k in self._src_checks if cb.isChecked()],
            }
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
        self.f_query.setText(d.get("query", ""))
        self.f_last.setText(d.get("last", "")); self.f_first.setText(d.get("first", ""))
        self.f_patr.setText(d.get("patr", "")); self.f_byear.setText(str(d.get("byear", "") or ""))
        self.f_repr_from.setText(str(d.get("repr_from", "") or ""))
        self.f_repr_to.setText(str(d.get("repr_to", "") or ""))
        self.f_doc_from.setText(str(d.get("doc_from", "") or ""))
        self.f_doc_to.setText(str(d.get("doc_to", "") or ""))
        self.f_region.setText(d.get("region", ""))
        self.f_place.setText(d.get("place", "")); self.f_object.setText(d.get("object", ""))
        self.f_doc.setText(d.get("doc", "")); self.f_folder.setText(d.get("folder", _DEF_DIR))
        for cb, key in ((self.f_lang, "lang"), (self.f_rtype, "rtype"),
                        (self.f_pamong, "pamong"), (self.f_oamong, "oamong")):
            i = cb.findText(d.get(key, ""))
            if i >= 0:
                cb.setCurrentIndex(i)
        t = d.get("tab")
        if t in self._tab_buttons:
            self._tab_buttons[t].setChecked(True)
        # restore the database selection (default: all checked)
        if "sources" in d and self._src_checks:
            keep = set(d.get("sources") or [])
            for cb, k in self._src_checks:
                cb.blockSignals(True); cb.setChecked(k in keep); cb.blockSignals(False)
            self._on_src_toggle()
        self._update_adv()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MemsearchApp(); w.show()
    sys.exit(app.exec())
