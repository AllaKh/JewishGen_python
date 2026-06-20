"""gui/hryc.py — search window for hryc.by (archives, newspapers and books).

Login (email + password) → search box → tick any of the 164 sources (grouped tree,
with a «Select all» toggle the site itself lacks) → results to Word + Excel.
Opening/saving the actual documents needs a paid account and comes later.
GUI is English; source names are shown «English / Русский» (the checkbox identity is
the site's own name — search uses the source id, the label is display only)."""
import sys, json, threading
from pathlib import Path

from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QGroupBox, QProgressBar,
    QFileDialog, QMessageBox, QScrollArea, QFrame, QToolButton)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from gui._app_icon import app_icon, make_header, make_cancel_button

try:
    import hryc_scraper as _scraper
    _SCRAPER_OK = True
except Exception:
    _SCRAPER_OK = False

_HERE    = Path(__file__).resolve().parent
_CONFIG  = _HERE.parent / "config"
_SRCFILE = _CONFIG / "hryc_sources.json"
_SAVE    = _HERE / ".hryc_autosave.json"
_DEF_DIR = str(_HERE.parent / "results")

STYLE = """
QMainWindow, QWidget { background:#f4f6f9; font-size:13px; }
QGroupBox { font-weight:bold; border:1px solid #c4cdd8; border-radius:6px;
            margin-top:10px; padding:8px; }
QGroupBox::title { subcontrol-origin:margin; left:10px; padding:0 4px; color:#2a4a7f; }
QLineEdit { padding:5px; border:1px solid #b6c0cc; border-radius:4px; background:white; }
QPushButton#startBtn { background:#2a6f4a; color:white; font-weight:bold;
            font-size:14px; padding:9px 22px; border:none; border-radius:5px; }
QPushButton#startBtn:hover { background:#34855a; }
QPushButton#startBtn:disabled { background:#9cbfac; }
QToolButton { border:none; background:transparent; }
"""


# ── password field with show/hide eye ───────────────────────────────────────── #
class PwdEdit(QLineEdit):
    """Password field with a show/hide eye button (no external icons)."""
    def __init__(self, *a):
        super().__init__(*a)
        self.setEchoMode(QLineEdit.Password)
        self._btn = QPushButton("\U0001F441", self)        # 👁
        self._btn.setFixedSize(26, 24)
        self._btn.setCheckable(True)
        self._btn.setCursor(Qt.PointingHandCursor)
        self._btn.setToolTip("Show / hide password")
        self._btn.setStyleSheet("border:none;background:transparent;")
        self._btn.toggled.connect(lambda v: self.setEchoMode(
            QLineEdit.Normal if v else QLineEdit.Password))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._btn.move(self.width() - 28, (self.height() - 24) // 2)
        self.setTextMargins(0, 0, 30, 0)


# ── background worker ───────────────────────────────────────────────────────── #
class Worker(QThread):
    progress     = Signal(int, str)
    finished     = Signal(dict)
    request_file = Signal(str)

    def __init__(self, payload):
        super().__init__()
        self.payload = payload
        self._file_choice = "overwrite"
        self._file_ev = threading.Event()

    def provide_file_choice(self, choice):
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
            result = _scraper.run_scraper(**self.payload)
        except Exception as exc:
            result = {"ok": False, "error": "exception",
                      "message": f"{type(exc).__name__}: {exc}"}
        self.finished.emit(result)


# ── main window ─────────────────────────────────────────────────────────────── #
class HrycApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("hryc.by")
        self.setMinimumWidth(820)
        self.setStyleSheet(STYLE)
        self.setWindowIcon(app_icon())
        self._src_checks = []          # [(checkbox, source_id)]
        self._build_ui()
        self._load()

    # ── build ─────────────────────────────────────────────────────────────── #
    def _build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(18, 12, 18, 12); outer.setSpacing(8)
        self._outer = outer

        outer.addLayout(make_header("hryclogo.png", "hryc.by", color="#5a3a1f"))
        sub = QLabel("Search in archives, newspapers and books — hryc.by")
        sub.setAlignment(Qt.AlignCenter)
        outer.addWidget(sub)

        # credentials
        cg = QGroupBox("Account (login required for results)")
        cl = QGridLayout(cg); cl.setSpacing(8)
        self.f_email = QLineEdit(); self.f_email.setPlaceholderText("Email")
        self.f_pass  = PwdEdit();   self.f_pass.setPlaceholderText("Password")
        cl.addWidget(QLabel("Email:"),    0, 0); cl.addWidget(self.f_email, 0, 1)
        cl.addWidget(QLabel("Password:"), 1, 0); cl.addWidget(self.f_pass,  1, 1)
        cl.setColumnStretch(1, 1)
        outer.addWidget(cg)

        # search query
        sg = QGroupBox("Search")
        sl = QGridLayout(sg); sl.setSpacing(8)
        self.f_query = QLineEdit()
        self.f_query.setPlaceholderText("Surname, place, etc. (фамилия, населённый пункт…)")
        sl.addWidget(QLabel("Query:"), 0, 0); sl.addWidget(self.f_query, 0, 1)
        sl.setColumnStretch(1, 1)
        # search options (the site's R.NoStemming / R.NoFuzziness / R.ShowExperts)
        opt = QHBoxLayout()
        self.f_nostem  = QCheckBox("No stemming")
        self.f_nofuzz  = QCheckBox("No fuzziness")
        self.f_experts = QCheckBox("Experts")
        self.f_nostem.setToolTip("Без стемминга — match the exact word form")
        self.f_nofuzz.setToolTip("Без ошибок — no fuzzy/typo matching")
        self.f_experts.setToolTip("Эксперты — include expert-level sources")
        for c in (self.f_nostem, self.f_nofuzz, self.f_experts):
            opt.addWidget(c)
        opt.addStretch()
        sl.addLayout(opt, 1, 0, 1, 2)
        outer.addWidget(sg)

        # where to search — tree of sources + Select all
        wg = QGroupBox("Where to search")
        wl = QVBoxLayout(wg); wl.setSpacing(4)
        toprow = QHBoxLayout()
        self.f_all = QCheckBox("Select all sources")
        self.f_all.stateChanged.connect(self._toggle_all)
        toprow.addWidget(self.f_all); toprow.addStretch()
        self.lbl_count = QLabel("0 selected")
        toprow.addWidget(self.lbl_count)
        wl.addLayout(toprow)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.StyledPanel)
        scroll.setMaximumHeight(int(self.screen().availableGeometry().height() * 0.42))
        host = QWidget(); self._tree = QVBoxLayout(host)
        self._tree.setContentsMargins(6, 4, 6, 4); self._tree.setSpacing(1)
        self._build_tree()
        self._tree.addStretch()
        scroll.setWidget(host)
        wl.addWidget(scroll)
        outer.addWidget(wg, 1)

        # output
        og = QGroupBox("Output")
        ol = QVBoxLayout(og); ol.setSpacing(6)
        fr = QHBoxLayout()
        self.f_docx = QCheckBox("Word (.docx)");  self.f_docx.setChecked(True)
        self.f_xlsx = QCheckBox("Excel (.xlsx)"); self.f_xlsx.setChecked(True)
        fr.addWidget(self.f_docx); fr.addWidget(self.f_xlsx); fr.addStretch()
        fr.addWidget(QLabel("(documents: later, with a paid account)"))
        ol.addLayout(fr)
        dr = QHBoxLayout(); dr.setSpacing(6)
        self.f_folder = QLineEdit(); self.f_folder.setText(_DEF_DIR)
        bb = QPushButton("Browse…"); bb.setFixedWidth(80); bb.clicked.connect(self._browse)
        dr.addWidget(QLabel("Save to:")); dr.addWidget(self.f_folder, 1); dr.addWidget(bb)
        ol.addLayout(dr)
        outer.addWidget(og)

        # progress + start/cancel
        self.pbar = QProgressBar(); self.pbar.setValue(0)
        self.stlbl = QLabel("Ready")
        outer.addWidget(self.pbar); outer.addWidget(self.stlbl)
        br = QHBoxLayout()
        self.start_btn = QPushButton("START SEARCH"); self.start_btn.setObjectName("startBtn")
        self.start_btn.clicked.connect(self._start)
        br.addStretch(); br.addWidget(self.start_btn)
        self.cancel_btn = make_cancel_button(self, br); br.addStretch()
        outer.addLayout(br)
        outer.addWidget(QLabel("© 2026 Alla Khananashvili", alignment=Qt.AlignRight))

        for w in (self.f_email, self.f_pass, self.f_query, self.f_folder,
                  self.f_docx, self.f_xlsx, self.f_nostem, self.f_nofuzz, self.f_experts):
            (w.textChanged if isinstance(w, QLineEdit) else w.stateChanged).connect(self._save)

    # ── source tree ───────────────────────────────────────────────────────── #
    def _build_tree(self):
        try:
            tree = json.loads(_SRCFILE.read_text(encoding="utf-8"))
        except Exception as e:
            self._tree.addWidget(QLabel(f"(could not load sources: {e})"))
            return
        for node in tree:
            self._add_node(node, self._tree, 0)

    def _add_node(self, node, parent_layout, depth):
        """Recursive: a source → checkbox; a group → arrow that shows/hides its
        children, plus a group checkbox that ticks the whole subtree."""
        if "idx" in node:                                  # leaf source
            cb = QCheckBox(self._dual(node))
            cb.stateChanged.connect(self._on_check)
            self._src_checks.append((cb, node["id"]))
            row = QHBoxLayout(); row.setContentsMargins(depth * 18 + 18, 0, 0, 0)
            row.addWidget(cb); row.addStretch()
            parent_layout.addLayout(row)
            return [cb]

        # group node
        row = QHBoxLayout(); row.setContentsMargins(depth * 18, 0, 0, 0); row.setSpacing(2)
        arrow = QToolButton(); arrow.setText("▶"); arrow.setCheckable(True)
        arrow.setFixedWidth(16)
        gcb = QCheckBox(self._dual(node))
        gcb.setStyleSheet("font-weight:bold;")
        row.addWidget(arrow); row.addWidget(gcb); row.addStretch()
        parent_layout.addLayout(row)

        holder = QWidget(); hb = QVBoxLayout(holder)
        hb.setContentsMargins(0, 0, 0, 0); hb.setSpacing(1)
        holder.setVisible(False)
        parent_layout.addWidget(holder)

        kids = []
        for child in node.get("children", []):
            kids += self._add_node(child, hb, depth + 1)
        arrow.toggled.connect(lambda on, a=arrow, h=holder:
                              (a.setText("▼" if on else "▶"), h.setVisible(on)))
        # group checkbox ticks/unticks its whole subtree
        gcb.stateChanged.connect(
            lambda st, ks=kids: [c.setChecked(st == Qt.Checked.value) for c in ks])
        return kids

    def _dual(self, node) -> str:
        en = (node.get("en") or "").strip()
        ru = (node.get("label") or node.get("name") or "").strip()
        lbl = f"{en} / {ru}" if en and en != ru else (en or ru)
        return lbl.replace("&", "&&")                      # Qt eats a lone &

    def _toggle_all(self, st):
        on = st == Qt.Checked.value
        for cb, _ in self._src_checks:
            cb.blockSignals(True); cb.setChecked(on); cb.blockSignals(False)
        self._on_check()

    def _on_check(self, *_):
        n = sum(1 for cb, _ in self._src_checks if cb.isChecked())
        self.lbl_count.setText(f"{n} selected")
        self._save()

    def _selected_ids(self) -> list:
        return [sid for cb, sid in self._src_checks if cb.isChecked()]

    # ── helpers ───────────────────────────────────────────────────────────── #
    def _browse(self):
        p = QFileDialog.getExistingDirectory(
            self, "Select output folder", self.f_folder.text() or _DEF_DIR)
        if p: self.f_folder.setText(p)

    def _fmt(self) -> str:
        d, x = self.f_docx.isChecked(), self.f_xlsx.isChecked()
        return "both" if d and x else ("docx" if d else "xlsx" if x else "both")

    def _payload(self) -> dict:
        return {
            "email":         self.f_email.text().strip(),
            "password":      self.f_pass.text(),
            "query":         self.f_query.text().strip(),
            "sources":       self._selected_ids(),
            "no_stemming":   self.f_nostem.isChecked(),
            "no_fuzziness":  self.f_nofuzz.isChecked(),
            "show_experts":  self.f_experts.isChecked(),
            "output_format": self._fmt(),
            "output_folder": Path(self.f_folder.text().strip() or _DEF_DIR),
            "log":           print,
            "cancel_event":  getattr(self, "_cancel_ev", None),
        }

    def _validate(self) -> bool:
        if not self.f_query.text().strip():
            QMessageBox.warning(self, "Nothing to search", "Enter a search query."); return False
        if not self._selected_ids():
            QMessageBox.warning(self, "No sources",
                                "Tick at least one source (or «Select all»)."); return False
        if not self.f_docx.isChecked() and not self.f_xlsx.isChecked():
            QMessageBox.warning(self, "No output", "Select at least one output format."); return False
        if not _SCRAPER_OK:
            QMessageBox.critical(self, "Scraper not found",
                                 "hryc_scraper.py not found in project root."); return False
        return True

    # ── save / load ───────────────────────────────────────────────────────── #
    def _save(self, *_):
        try:
            _SAVE.write_text(json.dumps({
                "email":  self.f_email.text(),
                "query":  self.f_query.text(),
                "folder": self.f_folder.text(),
                "docx":   self.f_docx.isChecked(),
                "xlsx":   self.f_xlsx.isChecked(),
                "nostem": self.f_nostem.isChecked(),
                "nofuzz": self.f_nofuzz.isChecked(),
                "experts": self.f_experts.isChecked(),
                "sources": self._selected_ids(),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load(self):
        if not _SAVE.exists():
            return
        try:
            d = json.loads(_SAVE.read_text(encoding="utf-8"))
        except Exception:
            return
        self.f_email.setText(str(d.get("email", "")))
        self.f_query.setText(str(d.get("query", "")))
        self.f_folder.setText(str(d.get("folder", _DEF_DIR)) or _DEF_DIR)
        self.f_docx.setChecked(bool(d.get("docx", True)))
        self.f_xlsx.setChecked(bool(d.get("xlsx", True)))
        self.f_nostem.setChecked(bool(d.get("nostem", False)))
        self.f_nofuzz.setChecked(bool(d.get("nofuzz", False)))
        self.f_experts.setChecked(bool(d.get("experts", False)))
        want = set(d.get("sources", []))
        if want:
            for cb, sid in self._src_checks:
                if sid in want:
                    cb.blockSignals(True); cb.setChecked(True); cb.blockSignals(False)
            self._on_check()

    # ── file-conflict dialog ──────────────────────────────────────────────── #
    def _show_file_conflict_dialog(self, names: str):
        box = QMessageBox(self); box.setIcon(QMessageBox.Question)
        box.setWindowTitle("File already exists")
        box.setText(f"These output file(s) already exist:\n\n{names}\n\nWhat to do?")
        b_over = box.addButton("Overwrite", QMessageBox.DestructiveRole)
        b_app  = box.addButton("Append new results", QMessageBox.AcceptRole)
        b_skip = box.addButton("Skip (don't save)", QMessageBox.RejectRole)
        box.setDefaultButton(b_app); box.exec()
        c = box.clickedButton()
        self.worker.provide_file_choice(
            "append" if c is b_app else "skip" if c is b_skip else "overwrite")

    # ── start / finish ────────────────────────────────────────────────────── #
    def _start(self):
        if not self._validate():
            return
        self._cancel_ev = threading.Event()
        self.start_btn.setEnabled(False); self.cancel_btn.setEnabled(True)
        self.pbar.setValue(0); self.stlbl.setText("Starting…")
        self.worker = Worker(self._payload())
        self.worker.progress.connect(
            lambda v, t: (self.pbar.setValue(v), self.stlbl.setText(t)))
        self.worker.request_file.connect(self._show_file_conflict_dialog)
        self.worker.finished.connect(self._done)
        self.worker.start()

    def _done(self, r: dict):
        self.start_btn.setEnabled(True); self.cancel_btn.setEnabled(False)
        if r.get("ok"):
            n = r.get("n_records", 0)
            parts = (["Word"] if r.get("docx_count") else []) + \
                    (["Excel"] if r.get("xlsx_path") else [])
            msg = f"{n} record(s) saved"
            if parts: msg += " → " + " + ".join(parts)
            if r.get("denied"):
                msg += f"\n\n{r['denied']} source(s) need a paid account."
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
    w = HrycApp(); w.show()
    sys.exit(app.exec())
