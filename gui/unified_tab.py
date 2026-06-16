"""
gui/unified_tab.py — the «Unified Search» tab for the launcher.
Two fields (First / Last name) → searches every site with a personal search
(except Wikisource) and writes ONE grouped Word/Excel. MyHeritage needs login,
so it has its own credential fields; no other site needs credentials.
"""
import json
import sys
import threading
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox,
    QLabel, QLineEdit, QPushButton, QCheckBox, QProgressBar, QFileDialog,
    QMessageBox,
)
from PySide6.QtCore import QThread, Signal, Qt

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
_DEF_DIR = str(Path.home() / "Downloads" / "Unified_search_results")


class _Worker(QThread):
    progress = Signal(int, str)
    finished = Signal(dict)
    request_file = Signal(str)          # output files already exist → ask

    def __init__(self, payload):
        super().__init__()
        self.payload = payload
        self._file_choice = "overwrite"
        self._file_ev = threading.Event()

    def provide_file_choice(self, choice):
        self._file_choice = choice
        self._file_ev.set()

    def run(self):
        import asyncio
        import unified_search
        self.payload["progress"] = lambda v, t: self.progress.emit(int(v), str(t))

        def ask_file_conflict(names):
            """Called from the engine when output files exist; blocks for the
            choice (overwrite / append / skip)."""
            self._file_choice = "overwrite"
            self._file_ev.clear()
            self.request_file.emit("\n".join(names))
            self._file_ev.wait(timeout=300)
            return self._file_choice or "overwrite"
        self.payload["ask_file_conflict"] = ask_file_conflict

        try:
            result = asyncio.run(unified_search.run_unified(**self.payload))
        except Exception as exc:
            result = {"ok": False, "message": f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)


class UnifiedSearchTab(QWidget):
    def __init__(self):
        super().__init__()
        self._build()

    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 12, 16, 12)
        outer.setSpacing(10)

        intro = QLabel(
            "Enter a name and surname once — every site with a personal search "
            "is searched (Wikisource excluded). The name is transliterated to each "
            "site's language; results pages are collected (records are NOT opened) "
            "and saved to ONE Word/Excel, grouped by site.")
        intro.setWordWrap(True)
        outer.addWidget(intro)

        # Name
        ng = QGroupBox("Search")
        nl = QGridLayout(ng); nl.setSpacing(8)
        self.f_first = QLineEdit(); self.f_first.setPlaceholderText("First name")
        self.f_last  = QLineEdit(); self.f_last.setPlaceholderText("Last name")
        nl.addWidget(QLabel("First name:"), 0, 0); nl.addWidget(self.f_first, 0, 1)
        nl.addWidget(QLabel("Last name:"),  0, 2); nl.addWidget(self.f_last, 0, 3)
        nl.setColumnStretch(1, 1); nl.setColumnStretch(3, 1)
        self.f_strict = QCheckBox("Exact / strict match (ignored where a site has no such option)")
        nl.addWidget(self.f_strict, 1, 0, 1, 4)
        outer.addWidget(ng)

        # MyHeritage credentials (only site that needs login)
        mg = QGroupBox("MyHeritage login (only this site needs credentials)")
        ml = QHBoxLayout(mg); ml.setSpacing(8)
        self.f_mh_email = QLineEdit(); self.f_mh_email.setPlaceholderText("MyHeritage e-mail")
        self.f_mh_pass  = QLineEdit(); self.f_mh_pass.setPlaceholderText("Password")
        self.f_mh_pass.setEchoMode(QLineEdit.Password)
        ml.addWidget(QLabel("E-mail:")); ml.addWidget(self.f_mh_email, 2)
        ml.addWidget(QLabel("Password:")); ml.addWidget(self.f_mh_pass, 2)
        outer.addWidget(mg)

        # Output
        og = QGroupBox("Output")
        ol = QVBoxLayout(og); ol.setSpacing(6)
        fr = QHBoxLayout()
        self.f_docx = QCheckBox("Word (.docx)"); self.f_docx.setChecked(True)
        self.f_xlsx = QCheckBox("Excel (.xlsx)"); self.f_xlsx.setChecked(True)
        fr.addWidget(self.f_docx); fr.addWidget(self.f_xlsx); fr.addStretch()
        ol.addLayout(fr)
        dr = QHBoxLayout(); dr.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        dr.addWidget(QLabel("Save to:")); dr.addWidget(self.f_folder, 1); dr.addWidget(bb)
        ol.addLayout(dr)
        outer.addWidget(og)

        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        outer.addWidget(self.pbar); outer.addWidget(self.stlbl)

        br = QHBoxLayout()
        self.start_btn = QPushButton("START UNIFIED SEARCH")
        self.start_btn.setObjectName("openBtn")
        self.start_btn.clicked.connect(self._start)
        self.cancel_btn = QPushButton("Cancel"); self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel)
        br.addStretch(); br.addWidget(self.start_btn); br.addWidget(self.cancel_btn)
        br.addStretch()
        outer.addLayout(br)
        outer.addStretch()

        # Remember the search parameters between runs (NOT the password).
        self._load_auto()
        for w in (self.f_first, self.f_last, self.f_mh_email, self.f_folder):
            w.textChanged.connect(self._save_auto)
        for w in (self.f_strict, self.f_docx, self.f_xlsx):
            w.toggled.connect(self._save_auto)

    _AUTO = Path(__file__).resolve().parent / ".unified_autosave.json"

    def _save_auto(self, *_):
        try:
            self._AUTO.write_text(json.dumps({
                "first": self.f_first.text(), "last": self.f_last.text(),
                "strict": self.f_strict.isChecked(),
                "mh_email": self.f_mh_email.text(),
                "folder": self.f_folder.text(),
                "docx": self.f_docx.isChecked(), "xlsx": self.f_xlsx.isChecked(),
            }, ensure_ascii=False), "utf-8")
        except Exception:
            pass

    def _load_auto(self):
        try:
            d = json.loads(self._AUTO.read_text("utf-8"))
        except Exception:
            return
        self.f_first.setText(d.get("first", ""))
        self.f_last.setText(d.get("last", ""))
        self.f_strict.setChecked(bool(d.get("strict", False)))
        self.f_mh_email.setText(d.get("mh_email", ""))
        if d.get("folder"):
            self.f_folder.setText(d["folder"])
        self.f_docx.setChecked(bool(d.get("docx", True)))
        self.f_xlsx.setChecked(bool(d.get("xlsx", True)))

    def _browse(self):
        p = QFileDialog.getExistingDirectory(self, "Select output folder",
                                             self.f_folder.text() or _DEF_DIR)
        if p:
            self.f_folder.setText(p)

    def _fmt(self):
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _start(self):
        if not self.f_first.text().strip() and not self.f_last.text().strip():
            QMessageBox.warning(self, "Nothing to search",
                                "Enter a first or last name.")
            return
        self._cancel_ev = threading.Event()
        self.start_btn.setEnabled(False); self.cancel_btn.setEnabled(True)
        self.pbar.setValue(0); self.stlbl.setText("Starting…")
        self.worker = _Worker({
            "first_name": self.f_first.text().strip(),
            "last_name":  self.f_last.text().strip(),
            "strict":     self.f_strict.isChecked(),
            "mh_email":   self.f_mh_email.text().strip() or None,
            "mh_password": self.f_mh_pass.text() or None,
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log": print,
            "cancel_event": self._cancel_ev,
        })
        self.worker.progress.connect(
            lambda v, t: (self.pbar.setValue(v), self.stlbl.setText(t)))
        self.worker.finished.connect(self._done)
        self.worker.request_file.connect(self._show_file_conflict_dialog)
        self.worker.start()

    def _show_file_conflict_dialog(self, names: str):
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Question)
        box.setWindowTitle("File already exists")
        box.setText("These output file(s) already exist:\n\n"
                    f"{names}\n\nWhat would you like to do?")
        b_app  = box.addButton("Append new results", QMessageBox.AcceptRole)
        b_skip = box.addButton("Skip (don't save)", QMessageBox.RejectRole)
        box.addButton("Overwrite", QMessageBox.DestructiveRole)
        box.setDefaultButton(b_app)
        box.exec()
        clicked = box.clickedButton()
        choice = ("append" if clicked is b_app else
                  "skip" if clicked is b_skip else "overwrite")
        self.worker.provide_file_choice(choice)

    def _cancel(self):
        ev = getattr(self, "_cancel_ev", None)
        if ev:
            ev.set()
        self.cancel_btn.setEnabled(False)
        self.stlbl.setText("Cancelling — finishing the current site…")

    def _done(self, r):
        self.start_btn.setEnabled(True); self.cancel_btn.setEnabled(False)
        if r.get("ok"):
            if r.get("skipped"):
                msg = "Files were not saved (skipped)."
            else:
                msg = f"{r.get('n_records', 0)} result row(s)"
                if r.get("output_folder"):
                    msg += f"\n\nFolder:\n{r['output_folder']}"
            QMessageBox.information(self, "Done", msg)
            self.stlbl.setText("Done.")
        else:
            QMessageBox.critical(self, "Error",
                                 r.get("message", "Error — see terminal."))
            self.stlbl.setText("Error — see terminal.")
