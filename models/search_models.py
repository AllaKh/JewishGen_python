"""
models/search_models.py
-----------------------
Data classes shared between the GUI, storage layer, and scraper.
"""

from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import List


@dataclass
class SearchRow:
    data_type: str
    search_type: str
    text: str

    def valid(self) -> bool:
        return bool(self.data_type and self.search_type and self.text.strip())


@dataclass
class SearchProfile:
    email: str = ""
    password: str = ""

    country: str = "ALL COUNTRIES"
    keywords: List[str] = field(default_factory=list)
    rows: List[SearchRow] = field(default_factory=list)

    output_folder: str = ""
    # Must match what scraper.run_scraper expects: 'both' | 'docx' | 'xlsx'
    output_format: str = "both"

    paid: bool = False

    def to_dict(self) -> dict:
        return asdict(self)

    @staticmethod
    def from_dict(d: dict) -> SearchProfile:
        # Normalise output_format to lowercase so old JSON saved with 'BOTH'
        # still works after the fix.
        fmt = d.get("output_format", "both")
        if isinstance(fmt, str):
            fmt = fmt.lower()
        if fmt not in ("both", "docx", "xlsx"):
            fmt = "both"

        return SearchProfile(
            email=d.get("email", ""),
            password=d.get("password", ""),
            country=d.get("country", "ALL COUNTRIES"),
            keywords=d.get("keywords") or [],
            rows=[SearchRow(**r) for r in (d.get("rows") or [])],
            output_folder=d.get("output_folder", ""),
            output_format=fmt,
            paid=d.get("paid", False),
        )
