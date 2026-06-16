"""
myheritage_filter_crawler.py — standalone (NO GUI) crawler for MyHeritage's
«narrow by category» filter tree → JSON.
=============================================================================
Run it yourself:

    python myheritage_filter_crawler.py              # recurse to depth 4 (default)
    python myheritage_filter_crawler.py --depth 6    # go deeper
    python myheritage_filter_crawler.py --dry-run    # crawl + print, write nothing
    python myheritage_filter_crawler.py --user … --password …   # if not logged in

What it does
------------
 1. Opens MyHeritage with the SAME persistent profile as myheritage_scraper.py
    (.mh_profile) so the login/cookies are reused; signs in once if creds given.
    Picks the family site so the generic /research URL doesn't 404.
 2. Searches the surname «Smith» (reusing the scraper's own search) to land on a
    results page that shows the «narrow by category» facet.
 3. RECURSIVELY walks every category and every sub-category by clicking each one
    to narrow the result set and re-reading the facet. Modelled on
    ancestry_filter_crawler.py — same DFS + resume + raw dump.
 4. Writes labels (no counts — those are filter-result numbers and unstable) into
    config/myheritage_categories.json as a nested Ancestry-style tree:
        { "Category": { "children": { "Sub-cat": { "children": {…} } } } }

Resume
------
After every visited node the script saves config/myheritage_categories_crawl_raw.json
(a flat list of {path, label, depth}). Stop with Ctrl-C and re-run — any path
whose children are already in the dump is NOT refetched (just traversed). Delete
the raw dump file to force a fresh crawl.

NOTE: MyHeritage is login-walled, anti-bot and CLICK-based (no category URL
params). The facet selectors are best-effort; every node is logged. If a level
comes back empty, run with the window visible and adjust SITE["…"].
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
    "preset":   "English (.com EN)",         # ENGLISH .com site → English labels +
                                             # .com login cookies apply (NOT .co.il!)
    "surname":  "Smith",
    "json":     "myheritage_categories.json",
    # Recurse to all available levels. Default 4 matches the Ancestry crawler;
    # bump it with --depth N if MH later exposes deeper sub-categories.
    "max_depth": 4,
    "all_label": "All Collections",          # the «no narrowing» row — recorded, not drilled

    # Facet rows (live HTML):
    #   <span class="button_action_text narrow_down_link …" data-automations="action_text">
    #     <div class="name">Newspapers</div><span class="count">10,000+</span></span>
    # We read JUST the label; counts are filter results, not part of the taxonomy.
    "read_js": r"""() => {
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const out = [], seen = new Set();
        const sel = '[class*="narrow_down_link"], [data-automations="action_text"]';
        for (const n of document.querySelectorAll(sel)) {
            const nm  = n.querySelector('.name');
            const label = norm(nm ? nm.textContent : n.textContent);
            if (!label || label.length > 80 || seen.has(label)) continue;
            seen.add(label);
            out.push(label);
        }
        return out;
    }""",
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

    max_depth = args.depth or SITE["max_depth"]

    # nodes: flat list of {path: [str…], label: str, depth: int}
    # path[i] = parent's English label at level i+1; depth = the node's own level.
    nodes: list = []
    visited: set = set()        # (tuple(path), label) keys — drops re-recordings

    # ── Resume from dump ───────────────────────────────────────────────────── #
    resumed_dropped = 0
    if dump_path.exists():
        try:
            prev = json.loads(dump_path.read_text("utf-8"))
        except Exception as e:
            log(f"  !! could not parse resume dump ({type(e).__name__}) — starting fresh")
            prev = []
        for n in prev:
            d = n.get("depth", 99)
            label = _to_english(n.get("label") or "")
            path  = [_to_english(p) for p in n.get("path") or []]
            if d <= max_depth and label:
                key = (tuple(path), label)
                if key not in visited:
                    nodes.append({"path": path, "label": label, "depth": d})
                    visited.add(key)
            else:
                resumed_dropped += 1
        if nodes:
            log(f"  resume: loaded {len(nodes)} nodes "
                f"(dropped {resumed_dropped} beyond depth {max_depth})")

    # A parent path is "done" if at least one child was recorded under it.
    # We then skip the re-search/click for that path and just recurse into the
    # known children. ROOT (empty path) is added if anything was resumed.
    done_paths: set = set()
    kids_of: dict = {}
    for n in nodes:
        parent_key = tuple(n["path"])
        done_paths.add(parent_key)
        kids_of.setdefault(parent_key, []).append(n["label"])
    if nodes:
        done_paths.add(())                  # ROOT had its top-level read

    async with async_playwright() as pw:
        ctx, page = await M.make_browser_context(pw)

        try:
            try:
                await page.goto(login_url, wait_until="domcontentloaded",
                                timeout=45000)
            except Exception as e:
                log(f"  goto login: {type(e).__name__}: {e}")
            await M._accept_cookies(page, log)
            if args.user and args.password:
                try:
                    await M._login(page, login_url, has_cookies,
                                   args.user, args.password, log)
                except Exception as e:
                    log(f"  login: {type(e).__name__}: {e}")

            # Pick the family site so /research?s=<id>… works (the generic URL 404s).
            try:
                site_url = await M._handle_select_site(page, None, log)
                if site_url:
                    search_url = site_url
                    log(f"  research URL: {search_url}")
            except Exception as e:
                log(f"  select-site: {type(e).__name__}: {e}")

            async def visit(path, depth):
                """Read the facet at `path` and recurse into each child."""
                indent = "  " * depth
                key = tuple(path)
                if key in done_paths:
                    # already fetched on a previous run → reuse known children
                    kids = kids_of.get(key, [])
                    log(f"{indent}[{path[-1] if path else 'ROOT'}] → {len(kids)} "
                        f"категори(й) (resumed)")
                else:
                    results = await _fresh_results(page, search_url, params,
                                                   has_cookies)
                    if results is None:
                        log("  !! нет страницы результатов — стоп ветки")
                        return
                    if path and not await _click_path(results, path):
                        return
                    try:
                        labels = await results.evaluate(SITE["read_js"])
                    except Exception as e:
                        log(f"  !! read_js: {type(e).__name__}: {e}")
                        labels = []
                    kids = []
                    for raw_label in labels:
                        en = _to_english(raw_label)
                        ckey = (tuple(path), en)
                        if ckey in visited:
                            continue
                        visited.add(ckey)
                        nodes.append({"path": list(path), "label": en,
                                      "depth": depth + 1})
                        kids.append(en)
                    done_paths.add(key)
                    kids_of[key] = kids
                    log(f"{indent}[{path[-1] if path else 'ROOT'}] → {len(kids)} "
                        f"категори(й)")
                    try:
                        dump_path.write_text(
                            json.dumps(nodes, ensure_ascii=False, indent=2), "utf-8")
                    except Exception:
                        pass
                if depth + 1 < max_depth:
                    for child in kids:
                        if child == SITE["all_label"]:
                            continue       # «All Collections» = no narrowing
                        await visit(path + [child], depth + 1)

            log(f"==== crawling MyHeritage categories (max depth {max_depth}) ====")
            await visit([], 0)
            log(f"==== done: {len(nodes)} nodes (raw dump → {dump_path.name}) ====")
        finally:
            if nodes:
                _merge(nodes, args.dry_run)
            try: await ctx.close()
            except Exception: pass


def _merge(nodes: list, dry: bool):
    """nodes [{path, label, depth}] → nested {label:{children:{…}}}.

    Ancestry-style: no counts, no codes — just labels and nested children. Keeps
    every existing entry; only adds new ones."""
    jp = _CONFIG / SITE["json"]
    tree = json.loads(jp.read_text("utf-8")) if jp.exists() else {}

    def _walk_to(parent_path):
        """Walk the tree to the dict that should hold this node's siblings, adding
        empty parents along the way if missing."""
        cur = tree
        for p in parent_path:
            p_en = _to_english(p)
            if p_en not in cur:
                cur[p_en] = {"children": {}}
            elif "children" not in cur[p_en]:
                cur[p_en]["children"] = {}
            cur = cur[p_en]["children"]
        return cur

    added = 0
    for n in sorted(nodes, key=lambda x: x["depth"]):     # parents before children
        host = _walk_to(n["path"])
        label = _to_english(n["label"])
        if label not in host:
            host[label] = {"children": {}}
            added += 1
        elif "children" not in host[label]:
            host[label]["children"] = {}
    log(f"categories: +{added} new")
    if not dry:
        jp.write_text(json.dumps(tree, ensure_ascii=False, indent=2), "utf-8")
        log(f"  wrote {jp}")


def main():
    ap = argparse.ArgumentParser(
        description="Crawl MyHeritage's narrow-by-category filter tree into the JSON.")
    ap.add_argument("--depth", type=int, default=0,
                    help=f"max depth (0 = SITE default = {SITE['max_depth']})")
    ap.add_argument("--dry-run", action="store_true", help="crawl + print, no write")
    ap.add_argument("--site", default=None,
                    help="site preset (default «English (.com EN)»); NOT .co.il")
    ap.add_argument("--user", default=None,
                    help="MyHeritage e-mail (if not logged in)")
    ap.add_argument("--password", default=None, help="MyHeritage password")
    args = ap.parse_args()
    try:
        asyncio.run(crawl(args))
    except KeyboardInterrupt:
        log("\ninterrupted — partial results were merged/dumped.")
        sys.exit(0)


if __name__ == "__main__":
    main()
