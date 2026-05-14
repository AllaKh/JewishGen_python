"""
config/constants.py
-------------------
Single source of truth for all lists and paths used by both the GUI and
the scraper. The scraper also defines its own copies of DATA_TYPES /
SEARCH_TYPES / COUNTRIES as a fallback — keep them in sync manually if
you add a new option.
"""
from pathlib import Path

# ── project layout ─────────────────────────────────────────────────────── #
ROOT_DIR    = Path(__file__).resolve().parent.parent
PROFILE_DIR = ROOT_DIR / "profile" / "chromium_profile"
RESULTS_DIR = ROOT_DIR / "results"

# ── JewishGen home URL ─────────────────────────────────────────────────── #
HOME_URL = "https://www.jewishgen.org/databases/"

# ── search form options ────────────────────────────────────────────────── #
COUNTRIES = [
    "ALL COUNTRIES", "Algeria", "Argentina", "Austria / Czechia", "Belarus",
    "Bulgaria", "Canada", "CryptoJews", "Egypt", "France", "Germany",
    "Hungary / Slovakia", "India", "Iraq", "Ireland", "Israel", "Italy",
    "Latvia", "Latin America", "Lebanon", "Libya", "Lithuania", "Morocco",
    "Netherlands", "Poland", "Portugal", "Romania / Moldova", "Sephardic",
    "Scandinavia", "Spain", "South Africa", "Syria", "Tunisia", "Turkey",
    "Ukraine", "United Kingdom", "United States", "Venezuela", "Yugoslavia",
    "Holocaust records", "Cemetery records", "Search by Cemetery",
]

DATA_TYPES = [
    "Surname",
    "Given Name",
    "Town",
    "Any Field",
]

SEARCH_TYPES = [
    "Phonetically Like",
    "Sounds Like",
    "Starts with",
    "is Exactly",
    "Fuzzy Match",
    "Fuzzier Match",
    "Fuzziest Match",
]
