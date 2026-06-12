# -*- coding: utf-8 -*-
"""Shared docx helpers used by every scraper's Word writer."""
import re

# A 4-digit year glued DIRECTLY (no space) to a following capital letter is the
# common "слипшиеся строки" symptom in family / event lists, e.g.
#   "Моисей … 1858 - 1931Сора-Лея … 1860 - 1936"
# A space after the year ("1950 Detroit") is left alone — that's a real date+place.
_YEAR_GLUE = re.compile(r"(1[5-9]\d\d|20\d\d)(?=[A-ZА-ЯЁ])")


def split_glued(value: str) -> list:
    """Return the value as a list of lines: split on existing newlines AND right
    after a year that is stuck to a capital letter (one person / event per line)."""
    text = _YEAR_GLUE.sub(lambda m: m.group(1) + "\n", str(value or ""))
    lines = [ln.strip() for ln in text.split("\n")]
    lines = [ln for ln in lines if ln]
    return lines or [""]


def set_cell_lines(cell, value):
    """Write `value` into a python-docx table cell, one item per LINE (not all
    glued into one run). First line replaces the cell text, the rest become extra
    paragraphs in the same cell."""
    lines = split_glued(value)
    cell.text = lines[0]
    for extra in lines[1:]:
        cell.add_paragraph(extra)
