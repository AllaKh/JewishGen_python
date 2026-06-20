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
)
from PySide6.QtCore import QThread, Signal, Qt, QTimer
from gui._app_icon import app_icon, make_header, make_cancel_button, autosave_path

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

    # ── UI ────────────────────────────────────────────────────────────────── #
    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(16, 12, 16, 12); outer.setSpacing(10)

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

        # Per-tab advanced fields (stacked)
        ag = QGroupBox("Advanced search (for the selected type)")
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
        self._tab_buttons[TABS[0]].setChecked(True)

        # Output
        og = QGroupBox("Output (Word)")
        ol = QHBoxLayout(og); ol.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        ol.addWidget(QLabel("Save to:")); ol.addWidget(self.f_folder, 1); ol.addWidget(bb)
        outer.addWidget(og)

        # Progress + start
        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        outer.addWidget(self.pbar); outer.addWidget(self.stlbl)

        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH")
        self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br)
        br.addStretch()
        outer.addLayout(br)
        outer.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        # Autosave wiring
        for w in (self.f_query, self.f_last, self.f_first, self.f_patr,
                  self.f_region, self.f_place, self.f_object, self.f_doc,
                  self.f_folder):
            w.textChanged.connect(self._save)
        self.f_byear.valueChanged.connect(self._save)
        for cb in (self.f_lang, self.f_rtype, self.f_pamong, self.f_oamong):
            cb.currentTextChanged.connect(self._save)
        for rb in self._tab_buttons.values():
            rb.toggled.connect(self._save)

    def _page_all(self):
        w = QWidget(); l = QVBoxLayout(w)
        l.addWidget(QLabel("Searches all record types at once. To refine, pick a "
                           "specific type above."))
        return w

    def _page_people(self):
        w = QWidget(); f = QFormLayout(w)
        self.f_last = QLineEdit(); self.f_first = QLineEdit(); self.f_patr = QLineEdit()
        self.f_byear = QSpinBox(); self.f_byear.setRange(0, 2026); self.f_byear.setSpecialValueText(" ")
        self.f_region = QLineEdit()
        self.f_rtype = QComboBox(); self.f_rtype.addItems(REGION_TYPE)
        f.addRow("Surname:", self.f_last)
        f.addRow("Name:", self.f_first)
        f.addRow("Patronymic:", self.f_patr)
        f.addRow("Year of birth:", self.f_byear)
        f.addRow("Region:", self.f_region)
        f.addRow("Region type:", self.f_rtype)
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
        f.addRow("Document name:", self.f_doc)
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
        QTimer.singleShot(0, self.adjustSize)

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
        by = self.f_byear.value()
        return {
            "query":       self.f_query.text().strip(),
            "lang":        SITE_LANG.get(self.f_lang.currentText(), "ru"),
            "entity_tab":  self._tab(),
            "last_name":   self.f_last.text().strip(),
            "first_name":  self.f_first.text().strip(),
            "patronymic":  self.f_patr.text().strip(),
            "birth_year":  str(by) if by else "",
            "region":      self.f_region.text().strip(),
            "region_type": self.f_rtype.currentText().strip(),
            "place_name":  self.f_place.text().strip(),
            "place_among": self.f_pamong.currentText().strip(),
            "object_name": self.f_object.text().strip(),
            "object_among": self.f_oamong.currentText().strip(),
            "doc_name":    self.f_doc.text().strip(),
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
                "patr": self.f_patr.text(), "byear": self.f_byear.value(),
                "region": self.f_region.text(), "rtype": self.f_rtype.currentText(),
                "place": self.f_place.text(), "pamong": self.f_pamong.currentText(),
                "object": self.f_object.text(), "oamong": self.f_oamong.currentText(),
                "doc": self.f_doc.text(), "folder": self.f_folder.text(),
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
        self.f_patr.setText(d.get("patr", "")); self.f_byear.setValue(int(d.get("byear", 0) or 0))
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MemsearchApp(); w.show()
    sys.exit(app.exec())
