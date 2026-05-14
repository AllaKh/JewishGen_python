"""
storage/autosave.py
-------------------
Saves the FULL last-used SearchProfile (including email, password, all
search rows, keywords, output folder, output format) so the GUI can
restore the previous session on next launch.

Password is stored in plain text in autosave.json — this is intentional
(the file lives locally and is never sent anywhere).
"""
import json
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent   # storage/
_ROOT = _HERE.parent                      # project root
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from models.search_models import SearchProfile  # noqa: E402

FILE = _HERE / "autosave.json"


def save(profile: SearchProfile) -> None:
    """Write profile to autosave.json (creates the file if absent)."""
    FILE.write_text(
        json.dumps(profile.to_dict(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def load() -> SearchProfile | None:
    """Return the saved SearchProfile, or None if missing / empty / corrupt."""
    if not FILE.exists():
        return None
    try:
        text = FILE.read_text(encoding="utf-8").strip()
        if not text:
            return None
        return SearchProfile.from_dict(json.loads(text))
    except Exception:
        return None
