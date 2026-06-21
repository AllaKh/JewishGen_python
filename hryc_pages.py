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


async def _title(page) -> str:
    try:
        t = await page.evaluate(
            "() => { const h = document.querySelector('h1,h2,h3.title,.title,.collab-title');"
            "        return h ? h.textContent : document.title; }")
        return re.sub(r"\s+", " ", (t or "")).strip(" «»\"")
    except Exception:
        return ""


async def _nav(page) -> dict:
    """Resolved absolute URLs of the prev («⟨», toward the start) and next («⟩», toward the
    end) page-nav arrows, plus the current url. Uses the parent div classes, with an arrow-
    glyph fallback."""
    return await page.evaluate(r"""() => {
        const r = {prev: '', next: '', cur: location.href};
        let p = document.querySelector('.collab-page-nav-prev a.collab-page-nav');
        let n = document.querySelector('.collab-page-nav-next a.collab-page-nav');
        if (!p || !n) {                                   // fallback: by arrow glyph
            for (const a of document.querySelectorAll('a.collab-page-nav')) {
                const t = (a.textContent || '').trim();
                if (!p && t.indexOf('⟨') >= 0) p = a;     // ⟨
                if (!n && t.indexOf('⟩') >= 0) n = a;     // ⟩
            }
        }
        if (p) r.prev = p.href;
        if (n) r.next = n.href;
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
    """Walk the folder that start_url belongs to: ⟨ to the start, then ⟩ to the end, saving
    every scan. Returns the number of scans saved."""
    if not await _goto(page, start_url):
        return 0
    # 1) walk LEFT (⟨) to the start of the folder
    walked = set()
    while True:
        nav = await _nav(page)
        cur = _aid(nav["cur"])
        if cur:
            walked.add(cur)
        prev = nav["prev"]
        if not prev or _aid(prev) in walked:           # at start / loop guard
            break
        if not await _goto(page, prev):
            break
    # 2) name the folder from the start page
    name = await _title(page) or "hryc"
    folder = out_root / H.safe_fn(f"{name}_{year}")
    # 3) walk RIGHT (⟩) to the end, saving each scan
    seq, fwd = 0, set()
    while True:
        nav = await _nav(page)
        cur = _aid(nav["cur"])
        if cur and cur in fwd:                          # loop guard
            break
        if cur:
            fwd.add(cur)
        if cur and cur not in saved:
            seq += 1
            base = f"{H.safe_fn(name)}_{seq}_{year}"
            files = await H._save_doc_scan(page, folder, base, log)
            saved.add(cur)
            if not files:
                log(f"    !! страница {seq}: скан не сохранён")
        nxt = nav["next"]
        if not nxt or _aid(nxt) in fwd:                 # at end / loop guard
            break
        if not await _goto(page, nxt):
            break
    return seq


async def run(query, year, out_root, max_pages, max_folders, shared):
    if not _PW_OK:
        _log("Playwright не установлен."); return
    out_root = Path(out_root); out_root.mkdir(parents=True, exist_ok=True)
    _log(f"Запрос: {query!r}, год: {year}, складываю в: {out_root}")

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
            sources = [s["id"] for s in H.load_sources()]
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

            # ── capture each NEW folder (skip results already saved) ──────────
            saved, folders, n_scans = set(), 0, 0
            for u in urls:
                if _aid(u) in saved:
                    _log(f"  · скан уже сохранён — пропускаю ({_aid(u)[-8:]})")
                    continue
                folders += 1
                _log(f"  [{folders}] новая папка — иду к началу и сохраняю всё…")
                n_scans += await _capture_folder(page, u, str(year), out_root, saved, _log)
                if max_folders and folders >= max_folders:
                    _log(f"  → достигнут лимит --max-folders {max_folders}"); break
            _log(f"\nГотово. Папок: {folders}, сканов сохранено: {n_scans}.")
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
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--max-pages", type=int, default=25, help="result pages to scan")
    ap.add_argument("--max-folders", type=int, default=0, help="stop after N folders (0 = all)")
    ap.add_argument("--shared", action="store_true", help="reuse the app's .hryc_profile (close the app first)")
    a = ap.parse_args()
    query = a.query_opt or a.query
    year = a.year_opt or a.year
    if not query or not year:
        ap.error("give a search word and a year, e.g.:  python hryc_pages.py \"Шендерович\" 1859")
    asyncio.run(run(query, year, a.out, a.max_pages, a.max_folders, a.shared))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
