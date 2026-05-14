"""
storage/autosave.py
-------------------
Save / load the last-used SearchProfile to storage/autosave.json so the
GUI can restore the previous session on startup.
"""

import json
from pathlib import Path

# ---------------------------------------------------------------------------
# The models/ package lives one level above storage/, so we need to resolve
# the project root before importing.
# ---------------------------------------------------------------------------
import sys
_HERE = Path(__file__).resolve().parent          # …/JewishGen_python/storage
_ROOT = _HERE.parent                             # …/JewishGen_python
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from models.search_models import SearchProfile   # noqa: E402 (after sys.path fix)

FILE = _HERE / "autosave.json"


def save(profile: SearchProfile) -> None:
    """Serialise *profile* to autosave.json."""
    FILE.write_text(
        json.dumps(profile.to_dict(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def load() -> SearchProfile | None:
    """Return the saved SearchProfile, or None if the file is missing /
    empty / corrupt."""
    if not FILE.exists():
        return None
    try:
        text = FILE.read_text(encoding="utf-8").strip()
        if not text:
            return None
        return SearchProfile.from_dict(json.loads(text))
    except Exception:
        return None
