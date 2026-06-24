"""
hryc_pages_resume.py — AUTOMATIC bulk downloader for hryc.by (no GUI, headless).

What it does, unattended
------------------------
* Walks YEARS by itself: from 1914 down to 1838 by default (override with --from / --to).
* For each year it fires ~500 DIFFERENT common Russian 2- and 3-letter substrings (ст, но, на,
  по, … then rarer ones and digits) as search queries — these OCR fragments appear in almost
  every document, so many of them surface the whole year.
* Every NEW document's scans are saved; documents already on disk are SKIPPED (keyed on the
  document id / base HAID read from the result URL — NOT the folder title, so different
  documents that share a title are never confused). First query of a year saves everything.
* It stops a year after TEN queries in a row return 0 new scans, then moves to the next year.
  One login, one browser, headless by default.
* --gazette "<NAME>" puts each year's scans in ONE folder «<NAME> <year>» (used by the gazette
  presets hryc_minsk.py / hryc_mogilev.py); without it, each document gets its own folder.

Run it
------
    python hryc_pages_resume.py                      # 1914 → 1838, all sources, per-document folders
    python hryc_pages_resume.py --show               # show the window (first login / captcha)
    python hryc_pages_resume.py --from 1905 --to 1900
    python hryc_pages_resume.py --source "губернские могилевские" --gazette "Могилевские"
    python hryc_pages_resume.py --instance 2         # run a 2nd copy in parallel (own profile)

For a single gazette, prefer the presets: hryc_minsk.py (Минские, 1838-1847) and hryc_mogilev.py
(Могилевские, 1917-1838). Headless can't solve a CAPTCHA or do the first login — do those once
with --show; the profile remembers. Several copies at once: give each a different --instance N.
"""

import argparse
import asyncio
import sys
from pathlib import Path

import hryc_scraper as H
import hryc_pages as P
from hryc_pages import (_aid, _base_haid, _goto, _capture_folder, _saved_bases,
                        _resolve_sources, _source_paths, _creds, _instance_profile,
                        _log, DEFAULT_OUT)

try:
    from playwright.async_api import async_playwright
    _PW_OK = True
except Exception:
    _PW_OK = False

YEAR_HI, YEAR_LO = 1914, 1838     # default span (descending): 1914 → 1838
ZERO_STOP        = 10             # stop a year after this many 0-new-scan queries in a row
COMBO_BATCH      = 500            # if a year hasn't hit the zero-streak, add this many more combos
MAX_Q_PER_YEAR   = 3000           # hard safety cap so a year can never loop forever

# Russian 2- and 3-letter substrings (+ digits). Searching these OCR fragments surfaces a
# year's documents; MANY different ones + the on-disk skip cover the whole year. MOST-COMMON
# FIRST (ст, но, на, по, …) so a year that has ANYTHING gets a hit right away and can't be
# falsely declared empty; the user's examples (ке, си, ад), rarer fragments and surname endings
# (вич, ский, штейн) come later. _make_combos() tops this curated list up to ~500 DISTINCT
# variants with a systematic frequent-letter bigram fill.
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


def _make_combos(target=500):
    """The first `target` DISTINCT Russian substrings, deterministic and EXTENSIBLE: the curated
    common-first _LEAD, then ALL frequent-letter bigrams, then trigrams. Deterministic order means
    _make_combos(500) is a prefix of _make_combos(1000) — so a year that hasn't hit its zero-streak
    after 500 can ask for 500 more and just keep going. Pool size ≈ 22 900 (168 + 28² + 28³)."""
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
    for x in freq:                                   # trigrams (the extra batches)
        for y in freq:
            for z in freq:
                if len(out) >= target:
                    return out
                add(x + y + z)
    return out


COMBOS = _make_combos(500)


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


async def run_auto(out_root, years, sources, max_pages, instance, shared,
                   no_stemming, no_fuzziness, show_experts, headless, zero_stop, gazette=None,
                   combo_offset=0, only=None):
    if not _PW_OK:
        _log("Playwright не установлен."); return
    out_root = Path(out_root); out_root.mkdir(parents=True, exist_ok=True)
    P._HEADLESS = headless
    profile = _instance_profile(instance, shared)
    _log(f"Автозагрузка hryc.by → {out_root}")
    if gazette:
        _log(f"Ведомости: {gazette} — на каждый год своя папка «{gazette} <год>»")
    _log(f"Годы: {years[0]} … {years[-1]}  ({len(years)} лет), запросов-вариантов: {len(COMBOS)}, "
         f"источников: {len(sources)}, профиль: {profile.name}, "
         f"браузер: {'headless' if headless else 'видимый (--show)'}")
    _log(f"Стоп года: {zero_stop} нулевых запросов подряд. Капча/первый вход — только с --show.")
    if combo_offset and not only:
        _log(f"  ВТОРОЙ НАБОР: начинаю с сочетания #{combo_offset} (совершенно другие 500), не с начала.")
    if only:
        _log(f"  ТОЛЬКО заданные запросы: {only} — без 500-набора и без авто-добора, по одному на год.")

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

            # ── year loop (descending) ───────────────────────────────────────
            for year in years:
                done = _saved_bases(out_root, str(year))     # docs already on disk for this year
                yfolder = out_root / H.safe_fn(f"{gazette} {year}") if gazette else None
                where = f"→ «{yfolder.name}»" if yfolder else "(папка по документу)"
                _log(f"\n===== ГОД {year} =====  на диске уже {len(done)} документов  {where}")
                zero = total = q = 0
                if only:                                      # ровно эти запросы, без авто-добора
                    pool, ci, extend = list(only), 0, False
                else:                                         # обычный режим: ~500 + авто-добор
                    pool, ci, extend = _make_combos(combo_offset + COMBO_BATCH), combo_offset, True
                why = ""
                while True:
                    if extend and zero >= zero_stop:          # zero-stop only in auto mode; --only/--alphabet runs ALL given queries
                        why = f"{zero_stop} нулей подряд"; break
                    if q >= MAX_Q_PER_YEAR:
                        why = f"предел {MAX_Q_PER_YEAR} запросов"; break
                    if ci >= len(pool):                       # ran out WITHOUT 10 zeros
                        if not extend:                        # --only: no auto-extend
                            why = f"перебрал заданные запросы ({len(pool)})"; break
                        bigger = _make_combos(len(pool) + COMBO_BATCH)
                        if len(bigger) <= len(pool):
                            why = f"кончились сочетания ({len(pool)})"; break
                        _log(f"  [{year}] {ci} сочетаний пройдено, {zero_stop} нулей подряд ещё НЕ "
                             f"набрано → добавляю +{len(bigger) - len(pool)} сочетаний (всего {len(bigger)})")
                        pool = bigger
                    combo = pool[ci]; ci += 1
                    q += 1
                    n = await _search_capture(page, combo, sources, year, out_root, done,
                                              max_pages, no_stemming, no_fuzziness, show_experts,
                                              yfolder)
                    total += n
                    zero = zero + 1 if n == 0 else 0
                    _log(f"  [{year}] «{combo}» → новых сканов: {n}  "
                         f"(нулей подряд: {zero}/{zero_stop}, за год: {total})")
                grand += total
                _log(f"===== ГОД {year} готов: новых сканов {total}, запросов {q} ({why}) =====")
            _log(f"\nВСЁ. Прошёл годы {years[0]}…{years[-1]}. Всего новых сканов за прогон: {grand}.")
            _log(f"Папка: {out_root}")
        finally:
            try: await ctx.close()
            except Exception: pass


def main(default_from=YEAR_HI, default_to=YEAR_LO, default_gazette=None, default_source=None,
         default_out=DEFAULT_OUT):
    ap = argparse.ArgumentParser(
        description="hryc.by AUTO bulk downloader: walks years (default 1914→1838), fires ~500 of "
                    "the most common Russian 2-3 letter substrings + digits, saves every NEW "
                    "document's scans (skips ones already on disk), stops a year after 10 "
                    "zero-result queries in a row.")
    ap.add_argument("--from", dest="y_from", type=int, default=default_from,
                    help=f"first year (default {default_from})")
    ap.add_argument("--to",   dest="y_to",   type=int, default=default_to,
                    help=f"last year (default {default_to})")
    ap.add_argument("--gazette", default=default_gazette, metavar="NAME",
                    help="put each year's scans in ONE folder «<NAME> <year>» (default: a "
                         "per-document folder named after the document title).")
    ap.add_argument("--source", action="append", default=[], metavar="PHRASE",
                    help="limit to sources whose full path contains ALL these words; repeat for "
                         "several; omit = the script's default source (or all).")
    ap.add_argument("--list-sources", action="store_true",
                    help="print every source path (optionally filtered by --source) and exit")
    ap.add_argument("--out", default=default_out)
    ap.add_argument("--max-pages", type=int, default=25, help="result pages per query")
    ap.add_argument("--zero-stop", type=int, default=ZERO_STOP,
                    help=f"stop a year after this many 0-new-scan queries in a row (default {ZERO_STOP})")
    ap.add_argument("--combo-offset", type=int, default=0, metavar="N",
                    help="start from combo #N instead of #0 — a SECOND pass with a COMPLETELY "
                         "different set of letter combos (e.g. --combo-offset 500 = the next 500).")
    ap.add_argument("--only", default="", metavar="QUERIES",
                    help="run ONLY these exact queries (comma-separated), nothing else — skips the "
                         "500-combo set and auto-extend. E.g. --only ъ (the hard sign, in almost "
                         "every pre-1918 document → one query catches nearly everything).")
    ap.add_argument("--alphabet", action="store_true",
                    help="run the whole Russian alphabet + digits as wildcard queries «*а*», «*б*», "
                         "…, «*я*», «*0*»…«*9*» (43 queries/year, all of them) — thorough sweep.")
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

    phrases = a.source or ([default_source] if default_source else [])
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

    step = -1 if a.y_from >= a.y_to else 1            # descending by default (1914 → 1838)
    years = list(range(a.y_from, a.y_to + step, step))
    only = [c for c in a.only.split(",") if c] or None
    if a.alphabet:                                   # whole alphabet + digits as «*x*» wildcards
        only = [f"*{c}*" for c in "абвгдеёжзийклмнопрстуфхцчшщъыьэюя0123456789"]
    asyncio.run(run_auto(a.out, years, source_ids, a.max_pages, a.instance, a.shared,
                         a.no_stemming, a.no_fuzziness, not a.no_experts,
                         headless=not a.show, zero_stop=a.zero_stop, gazette=a.gazette,
                         combo_offset=a.combo_offset, only=only))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
