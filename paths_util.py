"""Writable per-user data directory.

Browser profiles (.hryc_profile, .mh_profile, …) and autosaves must live somewhere WRITABLE.
Next to the source files is fine in development, but a PACKAGED install often sits in a
read-only place (C:\\Program Files\\…), where Chrome cannot create its --user-data-dir and the
browser launch fails with TargetClosedError. So in a frozen build we use
%APPDATA%\\JewishGenealogySearch instead; in dev we keep the project folder (paths unchanged).
"""
import os
import sys
from pathlib import Path


def user_data_dir() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("APPDATA") or Path.home()) / "JewishGenealogySearch"
    else:
        base = Path(__file__).resolve().parent
    try:
        base.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return base
