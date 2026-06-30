"""
archive_sweep.py — the ONE hryc.by bulk-download script (no GUI, headless).

What it does
------------
For the YEAR(S) and the SOURCE you set just below, it fires, one after another:
  * the 500 most useful Russian letter-combinations (_make_combos),
  * the whole alphabet as wildcards  «*а*», «*б*», … «*я*»,
  * the digits as wildcards          «*0*», … «*9*»,
and saves every NEW document's scans. Documents already on disk are SKIPPED (keyed on the
document id / base HAID read from the result URL — NOT the folder title, so two different
documents that share a title are never confused). One login, one browser.

Set it up (edit the four lines in the CONFIG block below):
  * YEAR_FROM / YEAR_TO — the year or the descending range to sweep.
  * SOURCE              — the source to limit to (path words, e.g. "губернские могилевские").
                          Leave "" to search ALL sources (huge — usually not what you want).
  * OUT_DIR             — where the scans go.

Run it
------
    python archive_sweep.py                 # sweep YEAR_FROM..YEAR_TO of SOURCE, headless
    python archive_sweep.py --show          # show the window (FIRST login / solving a CAPTCHA)
    python archive_sweep.py --from 1905 --to 1900            # override the years
    python archive_sweep.py --source "губернские киевские"   # override the source
    python archive_sweep.py --instance 2    # run a 2nd copy in parallel (its own profile)
    python archive_sweep.py --list-sources  # print every source path and exit

Headless can't solve a CAPTCHA or do the first login — do those ONCE with --show; the profile
remembers, then run headless. Several copies at once: give each a different --instance N.
"""

import argparse
import asyncio
import re
import sys
import json
from pathlib import Path

import hryc_scraper as H

# ═════════════════════════════════════════════════════════════════════════════
#  CONFIG — edit these four lines.
# ═════════════════════════════════════════════════════════════════════════════
YEAR_FROM = 1910          # first year to sweep
YEAR_TO   = 1914          # last year — set equal to YEAR_FROM for a SINGLE year;
                          #   for a whole range put the newer year first, e.g. 1914 / 1838
SOURCE    = "губернские могилевские"   # source path words (e.g. "губернские киевские").
                                        #   "" = ALL sources (huge — usually not wanted)
OUT_DIR   = r"D:\Archives"             # where the scans are written
# ═════════════════════════════════════════════════════════════════════════════

# The Windows console may be cp1251, and output may be piped/redirected to a file → printing
# our →/⟪/⟨/⟩/⏭/«» characters would crash. Force UTF-8 with a safe fallback.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    from playwright.async_api import async_playwright
    _PW_OK = True
except Exception:
    _PW_OK = False

# own profile (a Chrome user-data-dir can be open in only one process at a time). The same
# «.hryc_pages_profile» the old scripts used — so the existing login is reused, no re-login.
PAGES_PROFILE = Path(__file__).resolve().parent / ".hryc_pages_profile"

# Headless by default (logs only, no visible browser). --show flips it for the two cases that
# NEED a window: first-time login and solving a CAPTCHA.
_HEADLESS = True

MAX_PAGES_DEFAULT = 25            # result pages scanned per query


def _log(m):
    print(m, flush=True)


def _creds():
    try:
        d = json.loads((Path(__file__).resolve().parent / "gui" / ".hryc_autosave.json")
                       .read_text("utf-8"))
        return d.get("email", ""), d.get("password", "")
    except Exception:
        return "", ""


def _aid(url: str) -> str:
    m = re.search(r"(?:aid|id)=(HAID[\w]+)", url or "")
    return m.group(1) if m else ""


def _base_haid(aid: str) -> str:
    """Document id without the trailing `_<page>` — so all pages of ONE document share it,
    and we can tell when ⟨/⟩ crossed into the NEXT document. E.g.
    HAID2_17C8…D8_1 → HAID2_17C8…D8."""
    return re.sub(r"_\d+$", "", aid or "")


def _instance_profile(instance: int, shared: bool = False) -> Path:
    """Chrome user-data-dir for THIS run. SEVERAL copies of the script can run at once only if
    each uses its OWN profile (a Chrome user-data-dir is open in one process at a time).
      • instance 0 (default) → the master profile («.hryc_pages_profile»). Log in here ONCE.
      • instance N≥1        → a PRIVATE copy «.hryc_pages_profile_N», seeded from the master
                              on first use so it's already logged in. Each N is independent →
                              you can run them in parallel."""
    master = H.PROFILE_DIR if shared else PAGES_PROFILE
    if instance <= 0:
        master.mkdir(parents=True, exist_ok=True)
        return master
    p = master.parent / f"{master.name}_{instance}"
    if not p.exists():
        if master.exists() and any(master.iterdir()):
            import shutil
            ig = shutil.ignore_patterns(
                "Singleton*", "*.lock", "lockfile",
                "Cache", "Code Cache", "GPUCache", "ShaderCache", "DawnCache",
                "GraphiteDawnCache", "Service Worker", "CacheStorage")
            try:
                shutil.copytree(master, p, ignore=ig, dirs_exist_ok=True)
                _log(f"  → профиль #{instance} создан из «{master.name}» (логин скопирован)")
            except Exception as e:
                _log(f"  !! копия профиля не удалась ({type(e).__name__}) — войдёшь в окне вручную")
                p.mkdir(parents=True, exist_ok=True)
        else:
            p.mkdir(parents=True, exist_ok=True)
            _log(f"  → профиль #{instance} новый (master пуст) — войдёшь в окне вручную")
    return p


def _saved_bases(out_root: Path, year: str) -> set:
    """Base HAIDs of documents whose scans are ALREADY on disk for this year — read straight
    from the scan filenames («<aid>_<year>.<ext>»). Keyed on the document id (unique), so a
    re-run with another query skips exactly the documents already downloaded; a folder you
    DELETE is re-fetched (self-correcting — no stale index, no title-collision)."""
    bases, suffix = set(), f"_{year}"
    try:
        for f in out_root.rglob(f"*{suffix}.*"):
            if not f.is_file():
                continue
            stem = f.stem
            if stem.endswith(suffix):
                aid = stem[:-len(suffix)]
                if aid.startswith("HAID"):
                    bases.add(_base_haid(aid))
    except Exception:
        pass
    return bases


def _source_paths():
    """[(id, 'Group / Subgroup / Label')] for every leaf source — the full path so short
    leaf labels (three different «Могилевские»!) can be told apart by their group."""
    tree = json.loads(H.SRC_FILE.read_text("utf-8"))
    out = []

    def walk(ns, trail):
        for n in ns:
            if "idx" in n:                                  # leaf source
                lbl = n.get("label") or n.get("en") or n.get("id")
                out.append((n["id"], " / ".join(trail + [lbl])))
            else:                                           # group
                gl = n.get("name") or n.get("en") or ""
                walk(n.get("children", []), trail + ([gl] if gl else []))

    walk(tree, [])
    return out


def _resolve_sources(phrases):
    """Return (ids, paths). A source matches a phrase when its full path contains ALL of the
    phrase's words (case-insensitive, order-independent); multiple --source = union.
    No phrases → all sources."""
    paths = _source_paths()
    if not phrases:
        return [sid for sid, _ in paths], ["(все источники)"]
    ids, chosen = [], []
    for sid, path in paths:
        pl = path.lower()
        if any(all(w in pl for w in ph.lower().split()) for ph in phrases):
            ids.append(sid); chosen.append(path)
    return ids, chosen


async def _title(page) -> str:
    try:
        t = await page.evaluate(
            "() => { const h = document.querySelector('h1,h2,h3.title,.title,.collab-title');"
            "        return h ? h.textContent : document.title; }")
        return re.sub(r"\s+", " ", (t or "")).strip(" «»\"")
    except Exception:
        return ""


async def _nav(page) -> dict:
    """Resolved absolute URLs of the page-nav arrows + the current url:
      first «⟪» (jump straight to page 1), prev «⟨», next «⟩», last «⟫».
    Detected by glyph (unambiguous), with parent-div-class fallback."""
    return await page.evaluate(r"""() => {
        const r = {first: '', prev: '', next: '', last: '', cur: location.href};
        for (const a of document.querySelectorAll('a.collab-page-nav')) {
            const t = (a.textContent || '').trim();
            if      (t.indexOf('⟪') >= 0) r.first = a.href;   // ⟪ to first
            else if (t.indexOf('⟨') >= 0) r.prev  = a.href;   // ⟨ one back
            else if (t.indexOf('⟩') >= 0) r.next  = a.href;   // ⟩ one forward
            else if (t.indexOf('⟫') >= 0) r.last  = a.href;   // ⟫ to last
        }
        if (!r.prev)  { const p = document.querySelector('.collab-page-nav-prev a.collab-page-nav');  if (p) r.prev  = p.href; }
        if (!r.next)  { const n = document.querySelector('.collab-page-nav-next a.collab-page-nav');  if (n) r.next  = n.href; }
        if (!r.first) { const f = document.querySelector('.collab-page-nav-first a.collab-page-nav'); if (f) r.first = f.href; }
        return r;
    }""")


async def _goto(page, url):
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=40000)
    except Exception as e:
        _log(f"    !! не открыть {url[:70]} ({type(e).__name__})")
        return False
    # headless can't show the captcha to solve → wait only briefly, then move on (logged);
    # visible (--show) → full 5-min wait so you can solve it.
    await H._wait_if_captcha(page, _log, timeout=(6 if _HEADLESS else 300))
    return True


async def _capture_folder(page, start_url, year, out_root, saved, log, folder=None):
    """Capture the ONE document start_url belongs to: ⟪/⟨ to its first page, then ⟩ to its
    last page, saving every scan. The walk is BOUNDED to this document — it stops the moment
    an arrow points to a DIFFERENT base HAID (the next document), so we don't bleed into the
    rest of the collection. Scans are saved under their ORIGINAL name (<aid>_<year>).
    `folder` overrides the output directory (e.g. one «<gazette> <year>» dir for the whole
    year); None → a per-document «<title>_<year>» dir. Returns the number of scans saved."""
    if not await _goto(page, start_url):
        return 0
    doc = _base_haid(_aid(start_url) or _aid(page.url))     # this document's base id
    same = lambda a: bool(a) and _base_haid(a) == doc

    # 1) go to THIS document's first page: «⟪» in one jump if it points inside this doc,
    #    else step back via «⟨» while staying inside this doc.
    nav = await _nav(page)
    if same(_aid(nav.get("first", ""))):
        log("    ⟪ к началу документа (одним прыжком)")
        await _goto(page, nav["first"])
    else:
        walked = set()
        while True:
            nav = await _nav(page)
            walked.add(_aid(nav["cur"]))
            prev = nav["prev"]
            if not same(_aid(prev)) or _aid(prev) in walked:
                break
            if not await _goto(page, prev):
                break

    # 2) folder: caller-supplied «<gazette> <year>» dir, else per-document «<title>_<year>»
    if folder is None:
        name = await _title(page) or doc
        folder = out_root / H.safe_fn(f"{name}_{year}")

    # 3) walk ⟩ forward to the last page of THIS document, saving each scan
    seq, fwd = 0, set()
    while True:
        nav = await _nav(page)
        cur = _aid(nav["cur"])
        if cur in fwd:                                     # loop guard
            break
        fwd.add(cur)
        if cur and cur not in saved:
            seq += 1
            # ORIGINAL scan name (the aid, which already carries the page number) + year;
            # if some aid lacks a trailing number, append the folder sequence number.
            stem = cur if re.search(r"_\d+$", cur) else f"{cur}_{seq}"
            files = await H._save_doc_scan(page, folder, f"{stem}_{year}", log)
            saved.add(cur)
            if not files:
                log(f"    !! страница {seq}: скан не сохранён")
        nxt = _aid(nav["next"])
        if not same(nxt) or nxt in fwd:                    # next document / end → stop
            break
        if not await _goto(page, nav["next"]):
            break
    return seq


# ── the query pool: 500 letter-combos + the whole alphabet + digits as wildcards ──────────
_LEAD = [
    # most common Russian bigrams — broadest coverage, tried first
    "ст", "но", "то", "на", "по", "ен", "ов", "ни", "ра", "ко", "ро", "не", "ли", "во",
    "ка", "ер", "ет", "ал", "ор", "ри", "ан", "ос", "ом", "ва", "ла", "ле", "та", "ре",
    "ес", "ил", "де", "те", "се", "ме", "ть", "ин", "ит", "им", "од", "ой", "он", "от",
    "об", "ед", "ел", "ек", "ис", "из", "ас", "ат", "ач", "ам", "ум", "ия", "ые", "ых",
    # consonant clusters
    "пр", "тр", "ск", "сл", "кр", "гр", "бр", "сн", "зн", "дн", "тв", "мн", "кв", "пл",
    "гл", "бл", "вл", "сп", "нн", "ль",
    # softer / less frequent bigrams + the user's examples (ке, си, ад)
    "бе", "ве", "ге", "же", "пе", "че", "ши", "жи", "ци", "ди", "ги", "би", "ви", "ки",
    "ке", "си", "ад", "лю", "ня", "тя", "дя", "ча", "ща", "ло", "лы", "мы", "ры", "ты",
    # common trigrams
    "про", "при", "пре", "пер", "ста", "сто", "сте", "стр", "ост", "ого", "ова", "ние",
    "ени", "ный", "ной", "ров", "тор", "ско", "ска", "ные", "ест", "тра", "ива", "ани",
    "енн", "нос", "ред", "раз", "под", "пол", "кон", "ком", "гор", "дер", "нов", "ник",
    "тел", "тер", "чес", "пра", "ход", "мер", "вер", "лен",
    # surname-friendly fragments (genealogy corpus)
    "вич", "ева", "ина", "ций", "ский", "цкий", "енко", "ман", "берг", "штейн",
    # digits
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
]

_ALPHABET = "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
_DIGITS   = "0123456789"


def _make_combos(target=500):
    """The first `target` DISTINCT Russian substrings, deterministic: the curated common-first
    _LEAD, then ALL frequent-letter bigrams, then trigrams. Pool ≈ 22 900 (168 + 28² + 28³)."""
    out, seen = [], set()
    def add(c):
        if c not in seen:
            seen.add(c); out.append(c)
    for c in _LEAD:
        add(c)
        if len(out) >= target:
            return out
    freq = "оеаинтсрвлкмдпуяыьгзбйчхжшюцщ"
    for x in freq:                                   # bigrams
        for y in freq:
            if len(out) >= target:
                return out
            add(x + y)
    for x in freq:                                   # trigrams
        for y in freq:
            for z in freq:
                if len(out) >= target:
                    return out
                add(x + y + z)
    return out


def _sweep_pool():
    """The full per-year query list: 500 letter-combos, THEN the whole alphabet «*а*»…«*я*»,
    THEN the digits «*0*»…«*9*» as wildcards. De-duplicated, order preserved."""
    combos   = _make_combos(500)
    alphabet = [f"*{c}*" for c in _ALPHABET]
    digits   = [f"*{c}*" for c in _DIGITS]
    pool, seen = [], set()
    for q in combos + alphabet + digits:
        if q not in seen:
            seen.add(q); pool.append(q)
    return pool


async def _search_capture(page, query, sources, year, out_root, done, max_pages,
                          no_stemming, no_fuzziness, show_experts, folder=None):
    """One query: paginate the results, capture every NEW document (base HAID not yet in
    `done`). `done` carries over between queries of the same year so a document found by one
    substring isn't re-walked by another. `folder` (if given) is the one «<gazette> <year>»
    dir all of the year's scans go into. Returns the number of scans saved this query."""
    urls, seen = [], set()
    for pno in range(1, max_pages + 1):
        u = H.build_search_url(query, sources, pno, doc_dates=str(year),
                               no_stemming=no_stemming, no_fuzziness=no_fuzziness,
                               show_experts=show_experts)
        if not await _goto(page, u):
            break
        res = H.parse_results(await page.content(), _log)
        new = 0
        for r in res["rows"]:
            uu = r.get("url")
            if uu and uu not in seen:
                seen.add(uu); urls.append(uu); new += 1
        if new == 0:
            break
    scans, saved = 0, set()
    for u in urls:
        base = _base_haid(_aid(u))
        if base in done:
            continue
        done.add(base)
        scans += await _capture_folder(page, u, str(year), out_root, saved, _log, folder=folder)
    return scans


async def run_auto(out_root, years, sources, pool, max_pages, instance, shared,
                   no_stemming, no_fuzziness, show_experts, headless, gazette=None):
    if not _PW_OK:
        _log("Playwright не установлен."); return
    out_root = Path(out_root); out_root.mkdir(parents=True, exist_ok=True)
    global _HEADLESS
    _HEADLESS = headless
    profile = _instance_profile(instance, shared)
    _log(f"Автозагрузка hryc.by → {out_root}")
    if gazette:
        _log(f"Ведомости: {gazette} — на каждый год своя папка «{gazette} <год>»")
    _log(f"Годы: {years[0]} … {years[-1]}  ({len(years)} лет), запросов на год: {len(pool)} "
         f"(500 буквосочетаний + алфавит + цифры), источников: {len(sources)}, "
         f"профиль: {profile.name}, браузер: {'headless' if headless else 'видимый (--show)'}")
    _log("Каждый год прогоняется ВЕСЬ список (без раннего стопа). "
         "Капча/первый вход — только с --show.")

    def _clear_locks():
        for n in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            try: (profile / n).unlink()
            except Exception: pass

    async with async_playwright() as pw:
        ctx = None
        for _attempt in range(1, 4):
            _clear_locks()
            for _channel in ("chrome", None):     # real Chrome, else bundled Chromium
                try:
                    kw = dict(headless=headless, accept_downloads=True, no_viewport=True,
                              args=["--disable-blink-features=AutomationControlled",
                                    "--window-size=1680,1050" if headless else "--start-maximized"])
                    if _channel:
                        kw["channel"] = _channel
                    ctx = await pw.chromium.launch_persistent_context(str(profile), **kw)
                    _log(f"  → браузер: {'Google Chrome' if _channel else 'Chromium'} "
                         f"({'headless' if headless else 'видимый'})")
                    break
                except Exception as e:
                    _log(f"  !! {_channel or 'chromium'}: {type(e).__name__}")
            if ctx:
                break
            await asyncio.sleep(3)
        if ctx is None:
            _log(f"  !! Не удалось открыть браузер (профиль «{profile.name}» занят?). "
                 "Закрой другие копии и запусти снова."); return
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for extra in list(ctx.pages):
            if extra is not page:
                try: await extra.close()
                except Exception: pass

        grand = 0
        try:
            # ── login once ───────────────────────────────────────────────────
            await _goto(page, H.BASE_URL + "/")
            await H._accept_cookies(page, _log)
            if not await H._is_logged_in(page):
                email, password = _creds()
                if email and password:
                    await H._login(page, email, password, _log)
                if not await H._is_logged_in(page):
                    if headless:
                        _log("  !! НЕ залогинен, браузер скрытый — войти нельзя. Запусти ОДИН раз "
                             "с --show и войди; профиль запомнит, дальше работай headless.")
                        return
                    _log("  → Войди в открытом окне вручную (жду до 3 минут)…")
                    for _ in range(180):
                        if await H._is_logged_in(page):
                            break
                        await asyncio.sleep(1)
            if not await H._is_logged_in(page):
                _log("  !! НЕ залогинен — документы будут недоступны. Останавливаюсь."); return
            _log("  → Статус: залогинен")

            # ── year loop (descending) — run the FULL pool each year ──────────
            for year in years:
                done = _saved_bases(out_root, str(year))     # docs already on disk for this year
                yfolder = out_root / H.safe_fn(f"{gazette} {year}") if gazette else None
                where = f"→ «{yfolder.name}»" if yfolder else "(папка по документу)"
                _log(f"\n===== ГОД {year} =====  на диске уже {len(done)} документов  {where}")
                total = 0
                for qi, combo in enumerate(pool, 1):
                    n = await _search_capture(page, combo, sources, year, out_root, done,
                                              max_pages, no_stemming, no_fuzziness, show_experts,
                                              yfolder)
                    total += n
                    _log(f"  [{year}] «{combo}» ({qi}/{len(pool)}) → новых сканов: {n}  "
                         f"(за год: {total})")
                grand += total
                _log(f"===== ГОД {year} готов: новых сканов {total}, запросов {len(pool)} =====")
            _log(f"\nВСЁ. Прошёл годы {years[0]}…{years[-1]}. Всего новых сканов за прогон: {grand}.")
            _log(f"Папка: {out_root}")
        finally:
            try: await ctx.close()
            except Exception: pass


def main():
    ap = argparse.ArgumentParser(
        description="hryc.by ONE-script bulk downloader: for the year(s) and the source set at "
                    "the top of the file (or overridden below), sweeps the 500 letter-combos + "
                    "the whole alphabet + digits, saving every NEW document's scans (skips ones "
                    "already on disk).")
    ap.add_argument("--from", dest="y_from", type=int, default=YEAR_FROM,
                    help=f"first year (default {YEAR_FROM})")
    ap.add_argument("--to",   dest="y_to",   type=int, default=YEAR_TO,
                    help=f"last year (default {YEAR_TO})")
    ap.add_argument("--source", action="append", default=[], metavar="PHRASE",
                    help="limit to sources whose full path contains ALL these words; repeat for "
                         f"several; omit = the file's SOURCE ({SOURCE!r}).")
    ap.add_argument("--gazette", default=None, metavar="NAME",
                    help="put each year's scans in ONE folder «<NAME> <year>» (default: a "
                         "per-document folder named after the document title).")
    ap.add_argument("--list-sources", action="store_true",
                    help="print every source path (optionally filtered by --source) and exit")
    ap.add_argument("--out", default=OUT_DIR)
    ap.add_argument("--max-pages", type=int, default=MAX_PAGES_DEFAULT, help="result pages per query")
    ap.add_argument("--shared", action="store_true",
                    help="reuse the app's .hryc_profile (close the app first)")
    ap.add_argument("--instance", type=int, default=0, metavar="N",
                    help="run SEVERAL copies at once: give each a different N (1, 2, 3…). Each "
                         "gets its own Chrome profile copied from the master. Practical max ~3-4.")
    ap.add_argument("--show", action="store_true",
                    help="show the browser window (default = headless). Use it for the FIRST "
                         "login and for solving a CAPTCHA.")
    ap.add_argument("--no-stemming",  action="store_true", help="«Без стемминга» — exact word form only")
    ap.add_argument("--no-fuzziness", action="store_true", help="«Без ошибок» — no typo tolerance")
    ap.add_argument("--no-experts",   action="store_true", help="«Эксперты» OFF (default ON → more)")
    a = ap.parse_args()

    phrases = a.source or ([SOURCE] if SOURCE else [])
    source_ids, chosen = _resolve_sources(phrases)
    if a.list_sources:
        for sid, path in _source_paths():
            if not phrases or sid in source_ids:
                print(f"{sid:24} {path}")
        return
    if phrases and not source_ids:
        ap.error(f"no source matches {phrases!r}. Run --list-sources to see them.")
    if phrases:
        print("Источники:"); [print("  •", p) for p in chosen]
    else:
        print("ВНИМАНИЕ: источник не задан — ищу по ВСЕМ источникам (это очень много). "
              "Задай SOURCE в начале файла или --source \"<слова пути>\".")

    step = -1 if a.y_from >= a.y_to else 1            # descending by default
    years = list(range(a.y_from, a.y_to + step, step))
    pool = _sweep_pool()
    asyncio.run(run_auto(a.out, years, source_ids, pool, a.max_pages, a.instance, a.shared,
                         a.no_stemming, a.no_fuzziness, not a.no_experts,
                         headless=not a.show, gazette=a.gazette))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
