"""
build_mh_categories.py — turn the user's MyHeritage filter-structure docx
(results/MH.docx) into config/myheritage_categories.json (the GUI's category tree).

Run:
    python build_mh_categories.py [path-to-MH.docx]

The docx she produced marks the tree like this (highlight + level labels):
  • YELLOW highlight = the level markers «Первый/Второй/Третий уровень:» AND the
    example names she tagged with <div class="name">…</div>.
  • BRIGHT_GREEN highlight = collection names she did NOT tag (the bulk) — we treat
    their whole line as the name.

LEVEL RULES (learned the hard way — see CLAUDE.md):
  - The name right after «Первый уровень:» is an L1 category.
  - State machine: «Второй уровень:» → following names are L2; «Третий уровень:» →
    following names are L3.
  - A name IMMEDIATELY FOLLOWED BY «Третий уровень:» is an L2 heading — sibling L2s
    usually have NO «Второй уровень:» marker of their own (only «Третий уровень:»
    after them). This is why «Записей о рождении», «Смерть…», «Бракосочетание…» were
    wrongly nested under «Церковные реестры» before.
  - Some L1 (Газеты, Книги, Публичные отчёты) are a FLAT list of collections at the
    second level — no L3 at all.
  - Record counts live on SEPARATE lines or trailing the name («10 000+», «3 689»,
    «997», «… N records»). declean() strips trailing grouped counts; a line that is a
    BARE number is skipped entirely. Years and «1879–1934» ranges are KEPT.

Labels stay in the docx language (Russian here) — the scraper clicks the facet .name
by this exact text. English/Hebrew display is layered on later by matching a flat
same-order export position-by-position (see _ordered_names()).
"""
from __future__ import annotations
import json
import re
import sys
from pathlib import Path

from docx import Document

_HERE = Path(__file__).resolve().parent
_DOCX = _HERE / "results" / "MH.docx"
_JSON = _HERE / "config" / "myheritage_categories.json"

_NAME_RE = re.compile(r'<div class="name">(.*?)</div>', re.I | re.S)
_MARK_RE = re.compile(r'^(перв|втор|трет|четверт)\w*\s+уровень', re.I)
_LVL = {"перв": 1, "втор": 2, "трет": 3, "четверт": 4}
_NUMONLY = re.compile(r'^[\d\s.,+]+$')      # bare count line (no dash → year-ranges survive)


def _hl(p):
    for r in p.runs:
        try:
            if r.font.highlight_color:
                return str(r.font.highlight_color)
        except Exception:
            pass
    return None


def declean(s: str) -> str:
    """Strip a trailing record count but keep years / «1879–1934» / «часть N»."""
    s = re.sub(r'\s*\d[\d.,\s]*(?:records?|записей|запис[иь])\b.*$', '', s, flags=re.I)
    s = re.sub(r'\s*\b\d{1,3}(?:[.,\s]\d{3})+\+?\s*$', '', s)   # «1,526,476», «10 000+»
    s = re.sub(r'\s*\b\d{1,3}\+\s*$', '', s)                    # «100+»
    return re.sub(r'\s+', ' ', s).strip(' –,;')


def _ordered_names(doc) -> list:
    """The cleaned (kind, text) sequence in document order. kind ∈
    {MARK1..MARK4, NAME}. Used for parsing AND for position-matching a flat
    English/Hebrew export later."""
    seq = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if "<br>" in t or len(t) > 160:               # the stray collection-description blob
            if not _NAME_RE.search(t):
                continue
        mk = _MARK_RE.match(t.lower())
        if mk:
            for k, v in _LVL.items():
                if t.lower().startswith(k):
                    seq.append(("MARK%d" % v, ""))
                    break
            continue
        nm = _NAME_RE.search(t)
        raw = nm.group(1).strip() if nm else (t if _hl(p) == "BRIGHT_GREEN (4)" else None)
        if raw is None:
            continue
        name = declean(raw)
        if not name or _NUMONLY.match(name):          # empty / bare-count line → skip
            continue
        seq.append(("NAME", name))
    return seq


def build_tree(seq) -> dict:
    tree, last, cur, pending_L1 = {}, {1: None, 2: None}, None, False
    kind_at = lambda i: seq[i][0] if 0 <= i < len(seq) else ""
    for i, (kind, txt) in enumerate(seq):
        if kind == "MARK1":
            pending_L1 = True
            continue
        if kind in ("MARK2", "MARK3", "MARK4"):
            cur = int(kind[-1])
            continue
        if pending_L1:
            lvl, pending_L1, cur = 1, False, 2
        elif kind_at(i + 1) == "MARK3":               # heads an L3 block → it's an L2
            lvl = 2
        elif cur in (2, 3):
            lvl = cur
        else:
            lvl = 2
        if lvl == 1:
            node = tree.setdefault(txt, {})
            last[1], last[2] = node, None
        elif lvl == 2:
            par = last[1] if last[1] is not None else tree.setdefault("(прочее)", {})
            node = par.setdefault("children", {}).setdefault(txt, {})
            last[2] = node
        else:
            par = last[2] if last[2] is not None else (
                last[1] if last[1] is not None else tree.setdefault("(прочее)", {}))
            par.setdefault("children", {}).setdefault(txt, {})
    return tree


def _count(t) -> int:
    n = len(t)
    for v in t.values():
        if v.get("children"):
            n += _count(v["children"])
    return n


def main(argv):
    docx = Path(argv[1]) if len(argv) > 1 else _DOCX
    doc = Document(str(docx))
    seq = _ordered_names(doc)
    tree = build_tree(seq)
    _JSON.write_text(json.dumps(tree, ensure_ascii=False, indent=2), "utf-8")
    print(f"{docx.name}: {len(tree)} top categories, {_count(tree)} nodes → {_JSON.name}")
    for l1, v in tree.items():
        ch = v.get("children", {})
        l3 = sum(len(x.get("children", {})) for x in ch.values())
        print(f"  {l1}: L2={len(ch)} L3={l3}")


if __name__ == "__main__":
    main(sys.argv)
