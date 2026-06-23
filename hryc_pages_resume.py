"""
hryc_pages_resume.py — AUTOMATIC bulk downloader for hryc.by (no GUI, headless).

What it does, unattended
------------------------
* Walks YEARS by itself: from 1914 down to 1838 (1914, 1913, … 1838).
* For each year it fires a long list of the most common Russian 2- and 3-letter substrings
  (ке, си, ро, ад, про, при, ст, но, …) plus digits as search queries — these OCR fragments
  appear in almost every document, so cycling them surfaces the whole year.
* Every NEW document's scans are saved; documents already on disk are SKIPPED (keyed on the
  document id / base HAID read from the result URL — NOT the folder title, so different
  documents that share a title are never confused).
* It keeps querying a year until FIVE queries in a row return 0 new scans, then moves to the
  next year. One login, one browser, headless by default.

Run it
------
    python hryc_pages_resume.py                      # 1914 → 1838, all combos, all sources
    python hryc_pages_resume.py --show               # show the window (first login / captcha)
    python hryc_pages_resume.py --from 1905 --to 1900
    python hryc_pages_resume.py --source "губернские могилевские"
    python hryc_pages_resume.py --instance 2         # run a 2nd copy in parallel (own profile)

Headless can't solve a CAPTCHA or do the first login — do those once with --show; the profile
remembers, then run headless. Several copies at once: give each a different --instance N.
"""

import argparse
import asyncio
import itertools
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
ZERO_STOP        = 5              # stop a year after this many 0-new-scan queries in a row
MAX_Q_PER_YEAR   = 600           # safety cap so a year can never loop forever

# The most common Russian 2- and 3-letter substrings (+ digits). Searching these OCR fragments
# surfaces the bulk of a year's documents; cycling them + the on-disk skip covers the year.
COMBOS = [
    # user's examples first
    "ке", "си", "ро", "ад", "про", "при",
    # frequent bigrams (consonant+vowel / vowel+consonant)
    "ст", "но", "то", "на", "по", "ен", "ов", "ни", "ра", "во", "ко", "не", "ли", "ка",
    "ер", "ет", "ал", "ри", "ан", "ом", "ос", "ор", "ва", "ле", "ть", "ре", "ме", "де",
    "те", "се", "ла", "ло", "ил", "им", "ин", "ит", "ес", "од", "ой", "он", "от", "об",
    "из", "ис", "ас", "ат", "ач", "ел", "ед", "ек", "бе", "ве", "ге", "же", "пе", "че",
    "ши", "жи", "ци", "ди", "ги", "би", "ви", "ки", "пр", "тр", "ск", "сл", "кр", "гр",
    "бр", "сн", "зн", "дн", "тв", "мн",
    # frequent trigrams
    "пре", "пер", "ста", "сто", "сте", "стр", "ост", "ого", "ова", "ние", "ени", "ный",
    "ной", "ров", "тор", "ско", "ска", "ные", "ест", "тра", "ива", "ани", "енн", "нос",
    "ред", "раз", "под", "пол", "кон", "ком", "гор", "дер", "нов", "ник", "тел", "тер",
    "чес", "пра",
    # digits
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
]


async def _search_capture(page, query, sources, year, out_root, done, max_pages,
                          no_stemming, no_fuzziness, show_experts):
    """One query: paginate the results, capture every NEW document (base HAID not yet in
    `done`). `done` carries over between queries of the same year so a document found by one
    substring isn't re-walked by another. Returns the number of scans saved this query."""
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
        scans += await _capture_folder(page, u, str(year), out_root, saved, _log)
    return scans


async def run_auto(out_root, years, sources, max_pages, instance, shared,
                   no_stemming, no_fuzziness, show_experts, headless, zero_stop):
    if not _PW_OK:
        _log("Playwright не установлен."); return
    out_root = Path(out_root); out_root.mkdir(parents=True, exist_ok=True)
    P._HEADLESS = headless
    profile = _instance_profile(instance, shared)
    _log(f"Автозагрузка hryc.by → {out_root}")
    _log(f"Годы: {years[0]} … {years[-1]}  ({len(years)} лет), запросов-комбинаций: {len(COMBOS)}, "
         f"источников: {len(sources)}, профиль: {profile.name}, "
         f"браузер: {'headless' if headless else 'видимый (--show)'}")
    _log(f"Стоп года: {zero_stop} нулевых запросов подряд. Capтча/первый вход — только с --show.")

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
                _log(f"\n===== ГОД {year} =====  на диске уже {len(done)} документов")
                zero = total = q = 0
                for combo in itertools.cycle(COMBOS):
                    if zero >= zero_stop or q >= MAX_Q_PER_YEAR:
                        break
                    q += 1
                    n = await _search_capture(page, combo, sources, year, out_root, done,
                                              max_pages, no_stemming, no_fuzziness, show_experts)
                    total += n
                    zero = zero + 1 if n == 0 else 0
                    _log(f"  [{year}] «{combo}» → новых сканов: {n}  "
                         f"(нулей подряд: {zero}/{zero_stop}, за год: {total})")
                grand += total
                why = "лимит запросов" if q >= MAX_Q_PER_YEAR else f"{zero_stop} нулей подряд"
                _log(f"===== ГОД {year} готов: новых сканов {total}, запросов {q} ({why}) =====")
            _log(f"\nВСЁ. Прошёл годы {years[0]}…{years[-1]}. Всего новых сканов за прогон: {grand}.")
            _log(f"Папка: {out_root}")
        finally:
            try: await ctx.close()
            except Exception: pass


def main(default_from=YEAR_HI, default_to=YEAR_LO):
    ap = argparse.ArgumentParser(
        description="hryc.by AUTO bulk downloader: walks years (default 1914→1838), fires the most "
                    "common Russian 2-3 letter substrings + digits, saves every NEW document's "
                    "scans (skips ones already on disk), stops a year after 5 zero-result queries.")
    ap.add_argument("--from", dest="y_from", type=int, default=default_from,
                    help=f"first year (default {default_from})")
    ap.add_argument("--to",   dest="y_to",   type=int, default=default_to,
                    help=f"last year (default {default_to})")
    ap.add_argument("--source", action="append", default=[], metavar="PHRASE",
                    help="limit to sources whose full path contains ALL these words; repeat for "
                         "several; omit = all sources.")
    ap.add_argument("--list-sources", action="store_true",
                    help="print every source path (optionally filtered by --source) and exit")
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--max-pages", type=int, default=25, help="result pages per query")
    ap.add_argument("--zero-stop", type=int, default=ZERO_STOP,
                    help="stop a year after this many 0-new-scan queries in a row (default 5)")
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

    source_ids, chosen = _resolve_sources(a.source)
    if a.list_sources:
        for sid, path in _source_paths():
            if not a.source or sid in source_ids:
                print(f"{sid:24} {path}")
        return
    if a.source and not source_ids:
        ap.error(f"no source matches {a.source!r}. Run --list-sources to see them.")
    if a.source:
        print("Источники:"); [print("  •", p) for p in chosen]

    step = -1 if a.y_from >= a.y_to else 1            # descending by default (1914 → 1838)
    years = list(range(a.y_from, a.y_to + step, step))
    asyncio.run(run_auto(a.out, years, source_ids, a.max_pages, a.instance, a.shared,
                         a.no_stemming, a.no_fuzziness, not a.no_experts,
                         headless=not a.show, zero_stop=a.zero_stop))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
