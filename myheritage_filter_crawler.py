"""
myheritage_filter_crawler.py — standalone (NO GUI) crawler for MyHeritage's
«Ограничить поиск по категории» (narrow-by-category) filter tree → JSON.
=============================================================================
Run it yourself:

    python myheritage_filter_crawler.py                 # full recursion → JSON
    python myheritage_filter_crawler.py --dry-run       # crawl + print, write nothing
    python myheritage_filter_crawler.py --depth 3       # cap the depth
    python myheritage_filter_crawler.py --user … --password …   # if not logged in

What it does
------------
 1. Opens MyHeritage with the SAME persistent profile as myheritage_scraper.py
    (.mh_profile) so the login/cookies are reused; signs in once if creds given.
 2. Searches the surname «Smith» (reusing the scraper's own search) to land on a
    results page that shows the category facet.
 3. Walks the «narrow by category» facet to FULL DEPTH: reads the categories
    (`.narrow_down_link` → `.name` + `.count`), clicks each to narrow, reads the
    sub-categories that appear, recurses, and backtracks (re-search + click the
    path) — collecting the whole tree.
 4. Writes it to config/myheritage_categories.json (+ a raw dump alongside).

Same shape as ancestry_filter_crawler.py (a template for more sites).

NOTE: MyHeritage is login-walled, anti-bot and CLICK-based (no category URL
params), and is not reachable from the dev box — so the facet selectors below are
best-effort, inferred from the page HTML. Every node is logged; if a level comes
back empty, run with the window visible and adjust SITE["cat_*"] selectors. The
raw dump preserves whatever was collected.
"""
from __future__ import annotations
import argparse
import asyncio
import json
import sys
from pathlib import Path

from playwright.async_api import async_playwright

import myheritage_scraper as M     # reuse login / search / presets / profile

_HERE   = Path(__file__).resolve().parent
_CONFIG = _HERE / "config"

# label (in ANY site language) → English canonical, so the JSON stays ENGLISH no
# matter what language the account UI happens to render the facet in.
_EN_BY_VARIANT = {}
for _en, _vars in getattr(M, "_CAT_I18N", {}).items():
    _EN_BY_VARIANT[_en.strip().lower()] = _en
    for _v in _vars:
        _EN_BY_VARIANT[_v.strip().lower()] = _en


def _to_english(label: str) -> str:
    return _EN_BY_VARIANT.get((label or "").strip().lower(), label)


# ── SITE CONFIG (everything MyHeritage-specific) ─────────────────────────────── #
SITE = {
    "name":     "myheritage",
    "preset":   "English (.com EN)",         # ENGLISH .com site → English labels + the
                                             # .com login cookies apply (NOT Hebrew .co.il!)
    "surname":  "Smith",
    "json":     "myheritage_categories.json",
    # MyHeritage's «narrow by category» is a FLAT list (15 categories, no children) —
    # depth 1 reads them all from the first results page and STOPS. Without this the
    # crawler re-searched once per category looking for children that don't exist
    # (the «search → results → form reopens» loop). Override with --depth N to probe.
    "max_depth": 1,
    "all_label": "All Collections",          # the «no narrowing» row — recorded, not drilled

    # the category facet rows (from the live HTML the user provided):
    #   <span class="button_action_text narrow_down_link …" data-automations="action_text">
    #     <div class="name">Газеты</div><span class="count">10 000+</span></span>
    "read_js": r"""() => {
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const out = [], seen = new Set();
        const sel = '[class*="narrow_down_link"], [data-automations="action_text"]';
        for (const n of document.querySelectorAll(sel)) {
            const nm  = n.querySelector('.name');
            const cnt = n.querySelector('.count');
            const label = norm(nm ? nm.textContent : n.textContent);
            if (!label || label.length > 80 || seen.has(label)) continue;
            seen.add(label);
            out.push({label: label, count: norm(cnt ? cnt.textContent : '')});
        }
        return out;
    }""",
    # click the category whose .name equals the label → narrows the results
    "click_js": r"""(label) => {
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const sel = '[class*="narrow_down_link"], [data-automations="action_text"]';
        for (const n of document.querySelectorAll(sel)) {
            const nm = n.querySelector('.name');
            if (norm(nm ? nm.textContent : n.textContent) === label) {
                try { n.click(); } catch (e) {}
                try { (nm || n).click(); } catch (e) {}
                return true;
            }
        }
        return false;
    }""",
}


def log(msg: str):
    print(msg, flush=True)


def _params(lang):
    """A minimal search payload (only a surname) compatible with M._search."""
    keys = ("first_name surname name_strict name_variants name_initials "
            "name_startswith surname_strict year_match place_match birth_year "
            "birth_place father father_last mother mother_last spouse spouse_last "
            "death_year death_place residence military immigration keywords gender "
            "exact_match record_filter category").split()
    p = {k: "" for k in keys}
    p.update(surname=SITE["surname"], gender="Any", record_filter="All records",
             category="All", exact_match=False, lang=lang)
    return p


async def _wait_results(page, secs=20):
    """Wait until the result cards (showRecord links) are present."""
    for _ in range(secs * 2):
        try:
            if await page.evaluate(
                    "() => document.querySelectorAll('a[href*=\"showRecord\"]').length"):
                return True
        except Exception:
            pass
        await asyncio.sleep(0.5)
    return False


async def _first_record(page):
    try:
        return await page.evaluate(
            "() => (document.querySelector('a[href*=\"showRecord\"]')||{}).href || ''")
    except Exception:
        return ""


async def _fresh_results(form_page, search_url, params, has_cookies):
    """Re-run the search to get a clean results page (closing stale result tabs)."""
    ctx = form_page.context
    for p in list(ctx.pages):
        if p is not form_page and "research" in (p.url or "").lower():
            try: await p.close()
            except Exception: pass
    res = await M._search(form_page, search_url, params, has_cookies, log)
    if res:
        await _wait_results(res)
    return res


async def _click_path(results, path) -> bool:
    """Click each category label in `path` in order, waiting for the narrow to take
    effect between clicks. Returns False if any label couldn't be clicked."""
    for label in path:
        before = await _first_record(results)
        ok = False
        for _ in range(12):
            try:
                ok = await results.evaluate(SITE["click_js"], label)
            except Exception:
                ok = False
            if ok:
                break
            try: await results.evaluate("() => window.scrollBy(0, 300)")
            except Exception: pass
            await asyncio.sleep(0.6)
        if not ok:
            log(f"      !! не нашёл категорию для клика: «{label}»")
            return False
        # wait for the result set to refresh after narrowing
        for _ in range(20):
            await asyncio.sleep(0.4)
            if await _first_record(results) != before:
                break
    return True


async def crawl(args):
    preset = args.site or SITE["preset"]
    login_url, search_url, has_cookies = M.SITE_PRESETS.get(
        preset, list(M.SITE_PRESETS.values())[0])
    lang = M._site_lang(preset)
    params = _params(lang)
    dump_path = _CONFIG / SITE["json"].replace(".json", "_crawl_raw.json")
    log(f"site: {preset}   (login {login_url})")

    visited: set = set()
    nodes: list = []          # {path: [...], label, count, depth}

    async with async_playwright() as pw:
        # SAME anti-detect persistent context as the main scraper (profile +
        # webdriver/plugins spoof + FB blocking) — no bare context, no bot flag.
        ctx, page = await M.make_browser_context(pw)

        try:
            # Land on the site + ACCEPT COOKIES first (the banner blocks the form),
            # then sign in if creds were given (the persistent profile usually
            # already holds the session from the main scraper).
            try:
                await page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
            except Exception as e:
                log(f"  goto login: {type(e).__name__}: {e}")
            await M._accept_cookies(page, log)
            if args.user and args.password:
                try:
                    await M._login(page, login_url, has_cookies,
                                   args.user, args.password, log)
                except Exception as e:
                    log(f"  login: {type(e).__name__}: {e}")

            # CRITICAL: the generic «research/search/all/all» URL is «Page not found»
            # until a family site is chosen — the main scraper picks one and gets the
            # real «research?s=<id>» URL. Do the SAME here, else every goto 404s.
            try:
                site_url = await M._handle_select_site(page, None, log)
                if site_url:
                    search_url = site_url
                    log(f"  research URL: {search_url}")
            except Exception as e:
                log(f"  select-site: {type(e).__name__}: {e}")

            async def visit(path, depth):
                results = await _fresh_results(page, search_url, params, has_cookies)
                if results is None:
                    log("  !! нет страницы результатов — стоп ветки")
                    return
                if path and not await _click_path(results, path):
                    return
                try:
                    opts = await results.evaluate(SITE["read_js"])
                except Exception as e:
                    log(f"  !! read_js: {type(e).__name__}: {e}")
                    opts = []
                indent = "  " * depth
                kids = [o for o in opts
                        if tuple(path + [o["label"]]) not in visited]
                log(f"{indent}[{path[-1] if path else 'ROOT'}] → {len(kids)} категори(й)")
                for o in kids:
                    visited.add(tuple(path + [o["label"]]))
                    nodes.append({"path": list(path), "label": o["label"],
                                  "count": o.get("count", ""), "depth": depth + 1})
                try:
                    dump_path.write_text(
                        json.dumps(nodes, ensure_ascii=False, indent=2), "utf-8")
                except Exception:
                    pass
                if depth + 1 < (args.depth or SITE["max_depth"]):
                    for o in kids:
                        if o["label"] == SITE["all_label"]:
                            continue            # «Все коллекции» = no narrowing
                        await visit(path + [o["label"]], depth + 1)

            log("==== crawling MyHeritage categories ====")
            await visit([], 0)
            log(f"==== done: {len(nodes)} nodes (raw dump → {dump_path.name}) ====")
        finally:
            if nodes:
                _merge(nodes, args.dry_run)
            try: await ctx.close()
            except Exception: pass


def _merge(nodes: list, dry: bool):
    """nodes (path + label + count) → nested {label:{count, children:{…}}}."""
    jp = _CONFIG / SITE["json"]
    tree = json.loads(jp.read_text("utf-8")) if jp.exists() else {}
    added = 0
    for n in sorted(nodes, key=lambda x: x["depth"]):     # parents before children
        cur = tree
        for p in n["path"]:                                # walk to the parent (English)
            p = _to_english(p)
            cur = cur.setdefault(p, {"count": "", "children": {}})["children"]
        label = _to_english(n["label"])                    # English canonical label
        if label not in cur:
            cur[label] = {"count": n.get("count", ""), "children": {}}
            added += 1
        elif n.get("count"):
            cur[label]["count"] = n["count"]
    log(f"categories: +{added} new")
    if not dry:
        jp.write_text(json.dumps(tree, ensure_ascii=False, indent=2), "utf-8")
        log(f"  wrote {jp}")


def main():
    ap = argparse.ArgumentParser(
        description="Crawl MyHeritage's narrow-by-category filter tree into the JSON.")
    ap.add_argument("--depth", type=int, default=0, help="max depth (0 = full)")
    ap.add_argument("--dry-run", action="store_true", help="crawl + print, no write")
    ap.add_argument("--site", default=None,
                    help="site preset (default «English (.com EN)»); NOT Hebrew .co.il")
    ap.add_argument("--user", default=None, help="MyHeritage e-mail (if not logged in)")
    ap.add_argument("--password", default=None, help="MyHeritage password")
    args = ap.parse_args()
    try:
        asyncio.run(crawl(args))
    except KeyboardInterrupt:
        log("\ninterrupted — partial results were merged/dumped.")
        sys.exit(0)


if __name__ == "__main__":
    main()
