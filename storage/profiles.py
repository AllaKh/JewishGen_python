"""
storage/profiles.py
-------------------
Persist named SearchProfiles to storage/profiles.json.
Updated automatically after every successful search run.
"""
import json
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_ROOT = _HERE.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from models.search_models import SearchProfile  # noqa: E402

FILE = _HERE / "profiles.json"


def load_all() -> list[SearchProfile]:
    if not FILE.exists():
        return []
    try:
        data = json.loads(FILE.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
        return [SearchProfile.from_dict(p) for p in data]
    except Exception:
        return []


def save_all(profiles: list[SearchProfile]) -> None:
    FILE.write_text(
        json.dumps([p.to_dict() for p in profiles], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
