"""gui/_app_icon.py — shared app icon loader."""
from pathlib import Path
from PySide6.QtGui import QIcon, QPixmap

_CONFIG = Path(__file__).resolve().parent.parent / "config"


def app_icon() -> QIcon:
    """Return the application window icon (config/app_icon.png)."""
    pix = QPixmap(str(_CONFIG / "app_icon.png"))
    if not pix.isNull():
        return QIcon(pix)
    return QIcon()
