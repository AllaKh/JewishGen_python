"""gui/_app_icon.py — shared app icon + window-header helpers."""
import threading
from pathlib import Path
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Qt, QObject, QEvent
from PySide6.QtWidgets import (QLabel, QHBoxLayout, QPushButton, QApplication,
                               QWidget, QComboBox, QAbstractSpinBox,
                               QAbstractItemView)

_CONFIG = Path(__file__).resolve().parent.parent / "config"


# ── Mouse-wheel guard ──────────────────────────────────────────────────────── #
# Scrolling the mouse over a dropdown / spin-box must NOT change its value — you
# could silently change settings you never meant to. The combo opens only by click;
# the OPEN popup list still scrolls normally (its view is allowed below).
class _WheelGuard(QObject):
    def eventFilter(self, obj, ev):
        # Only widgets have parentWidget(); app-wide we also get QWindow/QObject
        # events → guard with isinstance(QWidget) (a QWindow has no parentWidget).
        if ev.type() == QEvent.Type.Wheel and isinstance(obj, QWidget):
            w, depth = obj, 0
            while isinstance(w, QWidget) and depth < 4:
                if isinstance(w, QAbstractItemView):
                    return False                    # open popup → allow scrolling
                if isinstance(w, (QComboBox, QAbstractSpinBox)):
                    return True                     # closed combo/spin → block change
                w = w.parentWidget(); depth += 1
        return False


_wheel_guard = None


def install_wheel_guard():
    """Install the wheel guard on the QApplication once (idempotent)."""
    global _wheel_guard
    app = QApplication.instance()
    if app is not None and _wheel_guard is None:
        _wheel_guard = _WheelGuard(app)
        app.installEventFilter(_wheel_guard)


def make_cancel_button(window, row_layout) -> QPushButton:
    """Add a «Cancel» button to row_layout and wire it to stop a running search.

    Contract for the window (kept uniform across all scraper GUIs):
      • `_payload()` passes `cancel_event=getattr(self, "_cancel_ev", None)`,
      • `_start()` sets `self._cancel_ev = threading.Event()` and calls
        `self.cancel_btn.setEnabled(True)`,
      • `_done()` calls `self.cancel_btn.setEnabled(False)`.
    The scrapers poll cancel_event and stop, saving whatever they have."""
    btn = QPushButton("Cancel")
    btn.setObjectName("cancelBtn")
    btn.setEnabled(False)
    btn.setStyleSheet(
        "QPushButton#cancelBtn{background:#b23b3b;color:white;font-weight:bold;"
        "font-size:13px;padding:8px 18px;border:none;border-radius:5px;}"
        "QPushButton#cancelBtn:hover{background:#c85050;}"
        "QPushButton#cancelBtn:disabled{background:#d9a9a9;}")

    def _cancel():
        ev = getattr(window, "_cancel_ev", None)
        if ev is not None:
            ev.set()
        btn.setEnabled(False)
        lbl = getattr(window, "stlbl", None)
        if lbl is not None:
            try: lbl.setText("Cancelling — finishing the current item…")
            except Exception: pass

    btn.clicked.connect(_cancel)
    row_layout.addWidget(btn)
    return btn


def app_icon() -> QIcon:
    """Return the application window icon (config/app_icon.png)."""
    install_wheel_guard()      # every GUI calls app_icon() → guard installed once
    pix = QPixmap(str(_CONFIG / "app_icon.png"))
    if not pix.isNull():
        return QIcon(pix)
    return QIcon()


def center_window(window, screen=None):
    """Cap a window to the screen and center it there, so EVERY GUI opens in the same
    spot, never off-screen, on any resolution. `screen` defaults to the window's current
    screen — the launcher passes its OWN screen so windows don't hop between monitors.
    Call it deferred (after the window's own dynamic sizing has settled)."""
    app = QApplication.instance()
    sc = screen or window.screen() or (app.primaryScreen() if app else None)
    if sc is None:
        return
    g = sc.availableGeometry()
    # title-bar + border thickness, so the WHOLE window (frame included) stays on screen
    dw = max(0, window.frameGeometry().width()  - window.geometry().width())
    dh = max(0, window.frameGeometry().height() - window.geometry().height())
    avail_w = g.width()  - dw - 8
    avail_h = g.height() - dh - 8
    # let the window shrink to the screen even if it set a larger minimum size
    window.setMinimumSize(min(window.minimumWidth(), max(avail_w, 200)),
                          min(window.minimumHeight(), max(avail_h, 200)))
    window.resize(min(window.width(), avail_w), min(window.height(), avail_h))
    w, h = window.width(), window.height()      # actual size after constraints
    fw, fh = w + dw, h + dh                      # full size incl. frame
    window.move(g.x() + max(0, (g.width()  - fw) // 2),
                g.y() + max(0, (g.height() - fh) // 2))


def make_header(logo_file: str, title: str,
                color: str = "#2a4a7f", logo_w: int = 110) -> QHBoxLayout:
    """
    Build a header row: logo flush-left + big title name CENTERED in the
    space between the logo and the right edge.
    `logo_file` is a filename inside config/ (e.g. 'MHlogo.png').
    Returns a QHBoxLayout ready to add to the window's outer layout.
    """
    row = QHBoxLayout()
    row.setSpacing(16)
    row.setContentsMargins(0, 2, 0, 2)

    pix = QPixmap(str(_CONFIG / logo_file))
    if not pix.isNull():
        logo = QLabel()
        logo.setPixmap(pix.scaledToWidth(logo_w, Qt.SmoothTransformation))
        logo.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        row.addWidget(logo)

    name = QLabel(title)
    name.setAlignment(Qt.AlignCenter)
    name.setStyleSheet(
        f"font-size:36px;font-weight:bold;color:{color};"
        "font-family:'Palatino Linotype',Palatino,Georgia,'Times New Roman',serif;"
        "letter-spacing:2px;")

    row.addStretch(1)        # spring between logo and name
    row.addWidget(name)
    row.addStretch(1)        # spring between name and right edge → name centered
    return row
