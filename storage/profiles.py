"""
storage/profiles.py
-------------------
Load / save all named SearchProfiles to storage/profiles.json.
"""

import json
from pathlib import Path

import sys
_HERE = Path(__file__).resolve().parent          # …/JewishGen_python/storage
_ROOT = _HERE.parent                             # …/JewishGen_python
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from models.search_models import SearchProfile   # noqa: E402

FILE = _HERE / "profiles.json"


def load_all() -> list[SearchProfile]:
    """Return all saved profiles, or [] if none exist."""
    if not FILE.exists():
        return []
    try:
        data = json.loads(FILE.read_text(encoding="utf-8"))
        return [SearchProfile.from_dict(p) for p in data]
    except Exception:
        return []


def save_all(profiles: list[SearchProfile]) -> None:
    """Overwrite profiles.json with the given list."""
    FILE.write_text(
        json.dumps([p.to_dict() for p in profiles], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
