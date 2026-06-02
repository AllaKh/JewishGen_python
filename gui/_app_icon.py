"""gui/_app_icon.py — shared app icon + window-header helpers."""
from pathlib import Path
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel, QHBoxLayout

_CONFIG = Path(__file__).resolve().parent.parent / "config"


def app_icon() -> QIcon:
    """Return the application window icon (config/app_icon.png)."""
    pix = QPixmap(str(_CONFIG / "app_icon.png"))
    if not pix.isNull():
        return QIcon(pix)
    return QIcon()


def make_header(logo_file: str, title: str,
                color: str = "#2a4a7f", logo_w: int = 110) -> QHBoxLayout:
    """
    Build a header row: logo flush-left + big title name next to it.
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
    name.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
    name.setStyleSheet(
        f"font-size:34px;font-weight:bold;color:{color};"
        "font-family:'Segoe UI',Arial,sans-serif;letter-spacing:1px;")
    row.addWidget(name)
    row.addStretch(1)
    return row
