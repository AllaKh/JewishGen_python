"""
myheritage_filter_crawler.py — standalone (NO GUI) crawler for MyHeritage's
category filter tree → JSON.  URL/code-based, exactly like ancestry_filter_crawler.py.
=============================================================================
Run it yourself:

    python myheritage_filter_crawler.py              # recurse to depth 4 (default)
    python myheritage_filter_crawler.py --depth 6    # go deeper
    python myheritage_filter_crawler.py --dry-run    # crawl + print, write nothing
    python myheritage_filter_crawler.py --user … --password …   # if not logged in

How it works (the RIGHT way — see CLAUDE.md):
 MyHeritage category links ARE real URLs with a code, just like Ancestry's
 record_f=:   /research/category-<CODE>/<slug>?lang=…&action=query&qname=…
 So we crawl by URL NAVIGATION + a GLOBAL dedup by code:
  1. login (persistent .mh_profile) + accept cookies + pick the family site +
     force lang=EN, then run ONE «Smith» search to land on a results page that
     shows the category facet.
  2. read every <a href*="/category-<code>"> on the page → (label, code, href).
  3. DFS: for each NOT-yet-seen code, goto its href, read its children, recurse.
     A global `visited` set of CODES drops the breadcrumb / the current category
     / «All Collections» / siblings already seen — so there is no infinite
     «Maps → Maps → Maps» and no re-search (re-searching tripped MH's anti-bot,
     spawning the endless «verify you're not a robot» window).
  4. write a nested ENGLISH tree:  { "Category": { "children": { … } } }

Resume: after every node we save myheritage_categories_crawl_raw.json (flat list
of {code,label,depth,parent,href}). Stop with Ctrl-C and re-run — codes whose
children are already saved are NOT refetched. Delete the dump for a fresh crawl.
"""
from __future__ import annotations
import argparse
import asyncio
import json
import re
import sys
from pathlib import Path

from playwright.async_api import async_playwright

import myheritage_scraper as M     # reuse login / search / presets / profile

_HERE   = Path(__file__).resolve().parent
_CONFIG = _HERE / "config"

# raw site label (any language) → English canonical, so the JSON stays English.
_EN_BY_VARIANT = {}
for _en, _vars in getattr(M, "_CAT_I18N", {}).items():
    _EN_BY_VARIANT[_en.strip().lower()] = _en
    for _v in _vars:
        _EN_BY_VARIANT[_v.strip().lower()] = _en


def _to_english(label: str) -> str:
    return _EN_BY_VARIANT.get((label or "").strip().lower(), label)


SITE = {
    "name":      "myheritage",
    "preset":    "English (.com EN)",
    "surname":   "Smith",
    "json":      "myheritage_categories.json",
    "max_depth": 4,                              # override with --depth N
}

# Category links carry the code in the path: /research/category-<digits>/…
_CODE_RE = r"/category-(\d+)"

# A row of category links on the page → [(label, code, href)].
_READ_JS = r"""() => {
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const seen = new Set(), out = [];
    for (const a of document.querySelectorAll('a[href*="/category-"]')) {
        const href = a.href || '';
        if (!/^https?:/i.test(href)) continue;          // skip javascript:/relative
        const m = href.match(/\/category-(\d+)/);
        if (!m) continue;
        const code = m[1];
        if (seen.has(code)) continue;
        const nm = a.querySelector('.name');
        let label = norm(nm ? nm.textContent : a.textContent);
        // strip the trailing result count: «Birth Records1,526,476,780 records» →
        // «Birth Records», «Maps5,232» → «Maps», «Photos10,000+» → «Photos».
        // Only strips «N records» or a comma-grouped/«+» number (so a real trailing
        // digit like «World War 2» is kept).
        label = label.replace(/\s*\d[\d.,\s]*records?\b.*$/i, '')
                     .replace(/\s*\d{1,3}(?:,\d{3})+\+?\s*$/, '')
                     .replace(/\s+/g, ' ').trim();
        if (!label || label.length > 90) continue;
        seen.add(code);
        out.push([label, code, href]);
    }
    return out;
}"""


def log(msg: str):
    print(msg, flush=True)


def _params(lang):
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
    for _ in range(secs * 2):
        try:
            if await page.evaluate(
                    "() => document.querySelectorAll('a[href*=\"showRecord\"], "
                    "a[href*=\"/category-\"]').length"):
                return True
        except Exception:
            pass
        await asyncio.sleep(0.5)
    return False


async def _bootstrap_search(form_page, search_url, params, has_cookies):
    """Run the surname search ONCE to land on a results page with category links."""
    res = await M._search(form_page, search_url, params, has_cookies, log)
    if res:
        await _wait_results(res)
    return res


async def _is_blocked(page) -> bool:
    """Detect MyHeritage's «verify you're not a robot» / CAPTCHA / block page."""
    try:
        u = (page.url or "").lower()
        if any(s in u for s in ("captcha", "px-captcha", "/blocked", "challenge",
                                 "are-you-human")):
            return True
        t = (await page.title() or "").lower()
        if any(s in t for s in ("robot", "captcha", "just a moment", "access denied",
                                "verify", "are you human")):
            return True
        txt = (await page.evaluate(
            "() => (document.body && document.body.innerText || '').slice(0, 600)")
        ).lower()
        return any(s in txt for s in ("not a robot", "press & hold", "verify you",
                                      "unusual activity", "are you a human"))
    except Exception:
        return False


class _Blocked(Exception):
    pass


async def read_options(page):
    try:
        return await page.evaluate(_READ_JS)
    except Exception as e:
        log(f"    !! read_options: {type(e).__name__}: {e}")
        return []


_SEE_ALL_RE = re.compile(
    r"^(show all|see all|show more|view all|more categories|all categories|more)\b",
    re.I)


async def expand_see_all(page):
    """Click «Show more / See all» inside the category facet so the FULL list is in
    the DOM before we read it — the results facet shows only the first few + a «more»
    control, which is why ROOT was coming back with 5 of the 15 categories."""
    for _ in range(5):
        clicked = False
        try:
            els = await page.query_selector_all(
                "a, button, span[role='button'], [class*='see' i], "
                "[class*='show-more' i], [class*='showmore' i], [class*='more' i]")
            for el in els:
                try:
                    txt = (await el.inner_text() or "").strip()
                except Exception:
                    continue
                if txt and _SEE_ALL_RE.match(txt):
                    try:
                        if await el.is_visible():
                            await el.click()
                            clicked = True
                            await asyncio.sleep(0.7)
                            break
                    except Exception:
                        pass
        except Exception:
            pass
        if not clicked:
            break


async def crawl(args):
    preset = args.site or SITE["preset"]
    login_url, search_url, has_cookies = M.SITE_PRESETS.get(
        preset, list(M.SITE_PRESETS.values())[0])
    lang = M._site_lang(preset)
    params = _params(lang)
    dump_path = _CONFIG / SITE["json"].replace(".json", "_crawl_raw.json")
    max_depth = args.depth or SITE["max_depth"]
    log(f"site: {preset}   (login {login_url})")

    nodes: list = []            # {code, label(raw), depth, parent, href}
    visited: set = set()        # GLOBAL set of codes — the anti-loop
    href_of: dict = {}

    # ── resume ─────────────────────────────────────────────────────────────── #
    if dump_path.exists():
        try:
            prev = json.loads(dump_path.read_text("utf-8"))
        except Exception:
            prev = []
        for n in prev:
            c = n.get("code")
            if c and n.get("depth", 99) <= max_depth and c not in visited:
                nodes.append(n); visited.add(c)
                if n.get("href"):
                    href_of[c] = n["href"]
        if nodes:
            log(f"  resume: loaded {len(nodes)} nodes")
    done_parents = {n.get("parent") for n in nodes}      # codes already drilled
    kids_of: dict = {}
    for n in nodes:
        kids_of.setdefault(n.get("parent"), []).append(n["code"])
    if nodes:
        done_parents.add(None)                            # ROOT was read

    def _save():
        try:
            dump_path.write_text(json.dumps(nodes, ensure_ascii=False, indent=2),
                                 "utf-8")
        except Exception:
            pass

    async with async_playwright() as pw:
        ctx, page = await M.make_browser_context(pw)
        try:
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

            # family site → real research URL, forced to English
            try:
                site_url = await M._handle_select_site(page, None, log)
                if site_url:
                    if re.search(r"[?&]lang=", site_url):
                        site_url = re.sub(r"([?&]lang=)[A-Za-z]+", r"\1EN", site_url)
                    else:
                        site_url += ("&" if "?" in site_url else "?") + "lang=EN"
                    search_url = site_url
                    log(f"  research URL: {search_url}")
            except Exception as e:
                log(f"  select-site: {type(e).__name__}: {e}")

            # ONE search to bootstrap the results page (the only form-submit)
            results = await _bootstrap_search(page, search_url, params, has_cookies)
            if results is None:
                log("  !! нет страницы результатов — стоп")
                return
            page = results
            if await _is_blocked(page):
                raise _Blocked()

            async def visit(parent_code, parent_label, depth):
                indent = "  " * depth
                if parent_code in done_parents:
                    kids = [(c, None) for c in kids_of.get(parent_code, [])]
                    log(f"{indent}[{parent_label or 'ROOT'}] → {len(kids)} "
                        f"категори(й) (resumed)")
                else:
                    if parent_code is not None:        # navigate into this category
                        href = href_of.get(parent_code)
                        if not href:
                            return
                        try:
                            await page.goto(href, wait_until="domcontentloaded",
                                            timeout=45000)
                        except Exception as e:
                            log(f"    !! goto {type(e).__name__}")
                            return
                        await asyncio.sleep(args.delay / 1000)
                        if await _is_blocked(page):
                            raise _Blocked()
                    await expand_see_all(page)         # reveal the full category list
                    opts = await read_options(page)
                    kids = []
                    for label, code, href in opts:
                        if code in visited:            # global dedup → no Maps→Maps
                            continue
                        visited.add(code); href_of[code] = href
                        nodes.append({"code": code, "label": label,
                                      "depth": depth + 1, "parent": parent_code,
                                      "href": href})
                        kids.append((code, label))
                    done_parents.add(parent_code)
                    kids_of[parent_code] = [c for c, _l in kids]
                    log(f"{indent}[{parent_label or 'ROOT'}] → {len(kids)} категори(й)")
                    _save()
                if depth + 1 < max_depth:
                    for code, label in kids:
                        lab = label or next((n["label"] for n in nodes
                                             if n["code"] == code), "")
                        await visit(code, lab, depth + 1)

            log(f"==== crawling MyHeritage categories (max depth {max_depth}) ====")
            await visit(None, None, 0)
            log(f"==== done: {len(nodes)} nodes (raw dump → {dump_path.name}) ====")
        except _Blocked:
            log("  ⛔ MyHeritage показал анти-бот/CAPTCHA — СТОП (не спамлю). "
                "Подожди, пройди проверку вручную, запусти снова (resume сохранён).")
        finally:
            if nodes:
                _merge(nodes, args.dry_run)
            try:
                await ctx.close()
            except Exception:
                pass


def _merge(nodes: list, dry: bool):
    """Flat {code,label,depth,parent} → nested ENGLISH {label:{children:{…}}}.
    PRESERVES the existing JSON (loads it, only ADDS) so a partial / interrupted
    crawl never deletes categories — the «JSON shrank from 15 to 5» bug. To start
    clean, delete config/myheritage_categories.json yourself."""
    jp = _CONFIG / SITE["json"]
    tree = json.loads(jp.read_text("utf-8")) if jp.exists() else {}
    dict_of: dict = {}                          # code → its node dict in the tree
    added = 0
    for n in sorted(nodes, key=lambda x: x["depth"]):    # parents before children
        label = _to_english(n["label"])
        parent = n.get("parent")
        container = (dict_of[parent].setdefault("children", {})
                     if parent in dict_of else tree)
        if label not in container:
            container[label] = {"children": {}}
            added += 1
        container[label].setdefault("children", {})
        dict_of[n["code"]] = container[label]
    log(f"categories: +{added} new ({len(nodes)} nodes crawled, "
        f"{len(tree)} roots total)")
    if not dry:
        jp.write_text(json.dumps(tree, ensure_ascii=False, indent=2), "utf-8")
        log(f"  wrote {jp}")


def main():
    ap = argparse.ArgumentParser(
        description="Crawl MyHeritage's category filter tree into the JSON.")
    ap.add_argument("--depth", type=int, default=0,
                    help=f"max depth (0 = default {SITE['max_depth']})")
    ap.add_argument("--delay", type=int, default=1500,
                    help="ms to wait after each navigation (politeness)")
    ap.add_argument("--dry-run", action="store_true", help="crawl + print, no write")
    ap.add_argument("--site", default=None, help="site preset (default English .com)")
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
