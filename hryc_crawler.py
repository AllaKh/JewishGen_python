"""
hryc_crawler.py — download EVERY available document from hryc.by «Закрома».

Run it yourself (the site needs a logged-in «Эксперт» account):

    python hryc_crawler.py
    python hryc_crawler.py --out "D:\\Archives"          # output root (default D:\\Archives)
    python hryc_crawler.py --start https://hryc.by/zakroma  # start page
    python hryc_crawler.py --max 50                       # stop after N documents (testing)

What it does
------------
Starts at https://hryc.by/zakroma and walks the whole tree by following links:
  category → periodical → year-grid → issue-grid → issue page (the «Локальная копия №1 / №2
  / №3» tabs) → /document/HAID… (the scan).
On an issue page it picks **«Локальная копия №1»**, opens its /document/HAID page and saves
the scan (`#mainCanvas`, the no-referrer Yandex-Disk image) — exactly like the main scraper.

Files go to:  <OUT>/<periodical name>/<year>/<issue>.jpg
(each document TYPE and YEAR in its own folder, descriptive names). Already-downloaded files
are skipped, so you can stop and re-run (resume). It is polite (small delays) and never dies
on a single broken page.

It reuses the main scraper's login / cookies / scan-saving and its persistent browser
profile (.hryc_profile), so if you've logged in once in the app, it's already logged in here.
"""

import argparse
import asyncio
import random
import re
import sys
import json
from pathlib import Path

import hryc_scraper as H            # reuse login / cookies / scan-saving / profile

try:
    from playwright.async_api import async_playwright
    _PW_OK = True
except Exception:
    _PW_OK = False

DEFAULT_OUT = r"D:\Archives"
START_URL   = H.BASE_URL + "/zakroma"

# Links worth following into the document tree (skip site nav / account / external).
_FOLLOW_RE = re.compile(r"(/zakroma|/documents?/|periodical\?|/document/HAID)", re.I)
# A document scan lives behind one of these «Локальная копия …» buttons.
_COPY_RE   = re.compile(r"локальн\w*\s+копи", re.I)


def _log(msg):
    print(msg, flush=True)


def _creds():
    """Best-effort email/password from the app's hryc autosave (so the crawler can log in
    if the persistent profile isn't already authenticated). Falls back to manual login."""
    for p in (Path(__file__).resolve().parent / "gui" / ".hryc_autosave.json",):
        try:
            d = json.loads(p.read_text("utf-8"))
            return d.get("email", ""), d.get("password", "")
        except Exception:
            continue
    return "", ""


async def _title(page) -> str:
    try:
        return (await page.evaluate(
            """() => { const h = document.querySelector('h1,h2,h3.title,.page-title');
                       return h ? h.textContent : document.title; }""") or "").strip()
    except Exception:
        return ""


async def _copy_one_href(page) -> str:
    """The href of «Локальная копия №1» (else the first «Локальная копия …») on an issue
    page, or '' if this page isn't an issue page."""
    return await page.evaluate(r"""() => {
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const links = [...document.querySelectorAll('a[href]')].filter(
            a => /локальн\w*\s+копи/i.test(norm(a.textContent)) && /\/document\//i.test(a.href));
        if (!links.length) return '';
        const one = links.find(a => /№?\s*1\b|N?\s*1\b/.test(norm(a.textContent)));
        return (one || links[0]).href;
    }""")


async def _child_links(page) -> list:
    """[(href, link-text)] for grid cells / sub-pages worth descending into."""
    out = await page.evaluate(r"""() => {
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const seen = new Set(); const out = [];
        for (const a of document.querySelectorAll('a[href]')) {
            const href = a.href || '';
            if (!href || seen.has(href)) continue;
            seen.add(href);
            out.push([href, norm(a.textContent)]);
        }
        return out;
    }""")
    res = []
    for href, text in out:
        if not href.startswith(H.BASE_URL):
            continue
        if not _FOLLOW_RE.search(href):
            continue
        if re.search(r"/(Account|Identity)/|logout", href, re.I):
            continue
        res.append((href, text))
    return res


def _clean_name(title: str) -> str:
    """Periodical name without the «за 1840 год» suffix and the « » quotes."""
    t = re.sub(r"\s*за\s*\d{4}\s*год.*$", "", title or "").strip()
    return t.strip("«»\" ").strip()


def _year(title: str) -> str:
    m = re.search(r"за\s*(\d{4})\s*год", title or "") or re.search(r"\b(1[5-9]\d\d|20\d\d)\b", title or "")
    return m.group(1) if m else ""


async def _save_scan(page, doc_href, out_root, name, year, label, seen_files, log):
    """Open /document/HAID and save its scan into <out_root>\\<name>\\<year>\\<label>.jpg."""
    haid = ""
    m = re.search(r"/document/(HAID[\w]+)", doc_href) or re.search(r"[?&]id=(HAID[\w]+)", doc_href)
    if m:
        haid = m.group(1)
    folder = Path(out_root) / (H.safe_fn(name) or "hryc") / (H.safe_fn(year) or "no_year")
    base = H.safe_fn(label) or (haid[:14] if haid else "doc")
    if haid:
        base = f"{base}_{haid[-8:]}"
    # resume: skip if any file with this base already exists
    if folder.exists() and any(folder.glob(base + ".*")):
        log(f"      ⏩ уже скачано: {folder.name}/{base}")
        return True
    try:
        await page.goto(doc_href, wait_until="domcontentloaded", timeout=40000)
    except Exception as e:
        log(f"      !! не открыть документ: {e}")
        return False
    await H._wait_if_captcha(page, log)                 # pause for a challenge before the scan
    saved = await H._save_doc_scan(page, folder, base, log)
    return bool(saved)


async def crawl():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--start", default=START_URL)
    ap.add_argument("--max", type=int, default=0, help="stop after N documents (0 = no limit)")
    ap.add_argument("--depth", type=int, default=10)
    ap.add_argument("--delay", type=float, default=4.0,
                    help="seconds between documents (+ random jitter). BIGGER = fewer CAPTCHAs; "
                         "try 15-30 if the site challenges every document.")
    args = ap.parse_args()

    if not _PW_OK:
        _log("Playwright не установлен."); return

    out_root = Path(args.out)
    out_root.mkdir(parents=True, exist_ok=True)
    _log(f"Складываю в: {out_root}")

    H.PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    H._clear_singleton_locks()

    seen_pages, seen_files = set(), set()
    n_docs = n_ok = 0

    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            str(H.PROFILE_DIR), headless=False, accept_downloads=True, no_viewport=True,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"])
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for extra in list(ctx.pages):
            if extra is not page:
                try: await extra.close()
                except Exception: pass

        # ── login (persistent profile usually already authenticated) ──────────
        await page.goto(H.BASE_URL + "/", wait_until="domcontentloaded", timeout=30000)
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
             else "  !! НЕ залогинен — большинство документов будут недоступны")
        await H._wait_if_captcha(page, _log)

        # ── DFS over the «Закрома» tree ──────────────────────────────────────
        stack = [(args.start, {"name": "", "year": "", "label": ""}, 0)]
        while stack:
            url, c, depth = stack.pop()
            if url in seen_pages or depth > args.depth:
                continue
            seen_pages.add(url)
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=40000)
            except Exception as e:
                _log(f"  !! пропускаю {url[:70]} ({type(e).__name__})")
                continue
            await H._wait_if_captcha(page, _log)        # pause for «я не робот» if it appears
            title = await _title(page)
            name = _clean_name(title) or c["name"]
            year = _year(title) or c["year"]

            # Issue page? → grab «Локальная копия №1» and save the scan.
            try:
                doc_href = await _copy_one_href(page)
            except Exception:
                doc_href = ""
            if doc_href:
                n_docs += 1
                _log(f"  [{n_docs}] {name} {year} — выпуск «{c['label']}»")
                ok = await _save_scan(page, doc_href, out_root, name, year, c["label"],
                                      seen_files, _log)
                n_ok += 1 if ok else 0
                # human-like pause — bigger --delay = fewer CAPTCHAs (the only honest lever;
                # we never bypass the challenge, only space out requests).
                await asyncio.sleep(args.delay + random.uniform(0, args.delay))
                if args.max and n_docs >= args.max:
                    _log(f"  → достигнут лимit --max {args.max}"); break
                continue

            # Otherwise it's a grid / category page → descend into its cells.
            try:
                kids = await _child_links(page)
            except Exception:
                kids = []
            # push in reverse so the first cell is processed first (DFS)
            for href, text in reversed(kids):
                if href not in seen_pages:
                    stack.append((href, {"name": name, "year": year, "label": text or c["label"]},
                                  depth + 1))

        try: await ctx.close()
        except Exception: pass

    _log(f"\nГотово. Документов найдено: {n_docs}, сканов сохранено/уже было: {n_ok}.")
    _log(f"Папка: {out_root}")


if __name__ == "__main__":
    try:
        asyncio.run(crawl())
    except KeyboardInterrupt:
        print("\nПрервано пользователем — скачанное сохранено.")
