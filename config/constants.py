from pathlib import Path

HOME_URL = "https://www.jewishgen.org/databases/"

ROOT_DIR = Path(__file__).resolve().parent.parent
PROFILE_DIR = ROOT_DIR / "profile" / "chromium_profile"
RESULTS_DIR = ROOT_DIR / "results"

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
    "Exact Match",
    "Contains",
    "Starts With",
]

OUTPUT_FORMATS = [
    "DOCX",
    "XLSX",
    "BOTH",
]