"""
hryc_pages.py — personal CLI tool (no GUI) for hryc.by.

What it does
------------
You give it a WORD and a YEAR on the command line. It searches hryc.by (all sources) for
that word + year, then for EACH result it captures the WHOLE document folder the result
belongs to:

  * open the result, then follow the «⟨» arrow (collab-page-nav-prev) to the START of the
    folder, then follow the «⟩» arrow (collab-page-nav-next) to the very END, saving every
    scan on the way as  <name>_<seq>_<year>.<ext>  in  <out>/<name>_<year>/ ;
  * before capturing the next result, check whether that scan was ALREADY saved this run
    (same scan id). If yes — skip it (it's a page inside a folder we already walked). If it
    belongs to a NEW folder — walk that folder start→end too.

So you get every distinct folder once, with no duplicate scans.

Run it yourself (needs your logged-in «Эксперт» account):
    python hryc_pages.py "Шендерович" 1859
    python hryc_pages.py --query "*а*" --year 1859 --out "D:\\Archives"
    python hryc_pages.py "Левин" 1900 --max-folders 5      # stop after 5 folders (testing)
    python hryc_pages.py "Левин" 1900 --shared             # reuse the app's .hryc_profile

CAPTCHA: hryc may show «Я не робот». The script NEVER solves it (that's off-limits) — it
brings the window forward and WAITS for you to click it, then continues.
"""

import argparse
import asyncio
import re
import sys
import json
from pathlib import Path

import hryc_scraper as H

try:
    from playwright.async_api import async_playwright
    _PW_OK = True
except Exception:
    _PW_OK = False

DEFAULT_OUT = r"D:\Archives"
# own profile so we don't clash with the app's .hryc_profile (a Chrome user-data-dir can be
# open in only one process at a time → TargetClosedError otherwise)
PAGES_PROFILE = Path(__file__).resolve().parent / ".hryc_pages_profile"


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
    await H._wait_if_captcha(page, _log)
    return True


async def _capture_folder(page, start_url, year, out_root, saved, log):
    """Capture the ONE document start_url belongs to: ⟪/⟨ to its first page, then ⟩ to its
    last page, saving every scan. The walk is BOUNDED to this document — it stops the moment
    an arrow points to a DIFFERENT base HAID (the next document), so we don't bleed into the
    rest of the collection. Scans are saved under their ORIGINAL name (<aid>_<year>).
    Returns the number of scans saved."""
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

    # 2) folder name = document title + year
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


async def run(query, year, out_root, max_pages, max_folders, shared, source_ids):
    if not _PW_OK:
        _log("Playwright не установлен."); return
    out_root = Path(out_root); out_root.mkdir(parents=True, exist_ok=True)
    _log(f"Запрос: {query!r}, год: {year}, источников: {len(source_ids)}, "
         f"складываю в: {out_root}")

    profile = H.PROFILE_DIR if shared else PAGES_PROFILE
    profile.mkdir(parents=True, exist_ok=True)

    def _clear_locks():
        for n in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            try: (profile / n).unlink()
            except Exception: pass

    async with async_playwright() as pw:
        ctx = None
        for attempt in range(1, 4):
            _clear_locks()
            try:
                ctx = await pw.chromium.launch_persistent_context(
                    str(profile), headless=False, accept_downloads=True, no_viewport=True,
                    args=["--start-maximized", "--disable-blink-features=AutomationControlled"])
                break
            except Exception as e:
                _log(f"  !! браузер не запустился (попытка {attempt}/3): {type(e).__name__}")
                await asyncio.sleep(3)
        if ctx is None:
            _log("  !! Не удалось открыть браузер. Закрой окна Chrome этого приложения "
                 f"(профиль «{profile.name}») и запусти снова."); return
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for extra in list(ctx.pages):
            if extra is not page:
                try: await extra.close()
                except Exception: pass

        try:
            # ── login + cookies ──────────────────────────────────────────────
            await _goto(page, H.BASE_URL + "/")
            await H._accept_cookies(page, _log)
            if not await H._is_logged_in(page):
                email, password = _creds()
                if email and password:
                    await H._login(page, email, password, _log)
                if not await H._is_logged_in(page):
                    _log("  → Войди в открытом окне вручную (жду до 3 минут)…")
                    for _ in range(180):
                        if await H._is_logged_in(page):
                            break
                        await asyncio.sleep(1)
            _log("  → Статус: залогинен" if await H._is_logged_in(page)
                 else "  !! НЕ залогинен — документы будут недоступны")

            # ── collect result document urls across pages ────────────────────
            sources = source_ids
            urls, seen = [], set()
            for pno in range(1, max_pages + 1):
                u = H.build_search_url(query, sources, pno, doc_dates=str(year))
                if not await _goto(page, u):
                    break
                res = H.parse_results(await page.content(), _log)
                if pno == 1:
                    _log(f"  Total на сайте: {res.get('total', 0)}")
                new = 0
                for r in res["rows"]:
                    if r.get("url") and r["url"] not in seen:
                        seen.add(r["url"]); urls.append(r["url"]); new += 1
                if new == 0:
                    break
            _log(f"  Результатов со ссылками: {len(urls)}")

            # ── go through results one by one; capture each NEW document once ──
            # saved   = scan aids already written this run (within-folder dedup)
            # done    = documents (base HAID) already captured → skip later results in them
            saved, done, folders, n_scans = set(), set(), 0, 0
            for i, u in enumerate(urls, 1):
                base = _base_haid(_aid(u))
                if base in done:
                    _log(f"  [{i}/{len(urls)}] · документ уже сохранён — следующий результат")
                    continue
                folders += 1
                _log(f"  [{i}/{len(urls)}] новый документ {base[-8:]} — иду к началу, сохраняю всё…")
                n_scans += await _capture_folder(page, u, str(year), out_root, saved, _log)
                done.add(base)
                if max_folders and folders >= max_folders:
                    _log(f"  → достигнут лимит --max-folders {max_folders}"); break
            _log(f"\nГотово. Документов: {folders}, сканов сохранено: {n_scans}.")
            _log(f"Папка: {out_root}")
        finally:
            try: await ctx.close()
            except Exception: pass


def main():
    ap = argparse.ArgumentParser(description="hryc.by: save all scans of every document folder "
                                             "matching a word + year.")
    ap.add_argument("query", nargs="?", default="", help="search word (e.g. Шендерович or *а*)")
    ap.add_argument("year", nargs="?", default="", help="year (Doc dates), e.g. 1859")
    ap.add_argument("--query", dest="query_opt", default="")
    ap.add_argument("--year", dest="year_opt", default="")
    ap.add_argument("--source", action="append", default=[], metavar="PHRASE",
                    help="limit to sources whose full path contains ALL these words, e.g. "
                         "--source \"губернские могилевские\". Repeat for several. Omit = all.")
    ap.add_argument("--list-sources", action="store_true",
                    help="print every source path (optionally filtered by --source) and exit")
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--max-pages", type=int, default=25, help="result pages to scan")
    ap.add_argument("--max-folders", type=int, default=0, help="stop after N folders (0 = all)")
    ap.add_argument("--shared", action="store_true", help="reuse the app's .hryc_profile (close the app first)")
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

    query = a.query_opt or a.query
    year = a.year_opt or a.year
    if not query or not year:
        ap.error("give a search word and a year, e.g.:  python hryc_pages.py \"Шендерович\" 1859")
    asyncio.run(run(query, year, a.out, a.max_pages, a.max_folders, a.shared, source_ids))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
