"""
ancestry_filter_crawler.py — standalone (NO GUI) filter-tree crawler for Ancestry
=================================================================================
Run it yourself from the command line, e.g.:

    python ancestry_filter_crawler.py                 # crawl locations + dates, merge
    python ancestry_filter_crawler.py --dry-run       # crawl + show, do NOT write JSON
    python ancestry_filter_crawler.py --only dates     # one axis only
    python ancestry_filter_crawler.py --depth 4        # locations down to cities too
    python ancestry_filter_crawler.py --headless       # no visible window

What it does
------------
 1. Opens Ancestry with the SAME persistent login profile as ancestry_scraper.py
    (.ancestry_profile) and signs in once if needed.
 2. Types «Smith» into the Last Name field and presses Search (falls back to the
    canonical results URL if the form fields can't be found).
 3. On the results page it walks the LEFT filter panel by NAVIGATING to each
    filter's own link (so it never loses the level) and reading the child options:
        • Location    : continent → country → state (→ city)  — links carry record_f=<code>
        • Record Date : century  → decade  → year             — links carry /categories/cen_<code>
    At every node it records each child's label + code, then drills into it.
 4. Merges everything into:
        config/ancestry_locations.json     (continent → country → state[/city])
        config/ancestry_record_dates.json  (century → decade → year)
    Existing entries, codes and US city lists are PRESERVED; only new nodes are
    added and missing codes filled. A raw dump is also written next to each JSON
    (…_crawl_raw.json) so you can inspect exactly what was found.

Adapting to the other three sites
---------------------------------
Everything site-specific lives in the SITE dict below: the URLs, the Last-Name
field, the search button, how a filter link encodes its code, and how a code turns
into a results URL. Copy this file, edit SITE (and, if a site needs clicks instead
of link-navigation, the two helpers read_options/child_goto), and the crawl/merge
machinery stays the same.

NOTE: the live Ancestry DOM is not reachable from the dev box, so the facet
selectors are best-effort. If a level comes back empty, run with the window
visible and adjust SITE["…"]["code_in_href"] / the see-all selectors — the script
logs what it finds at every node.
"""

from __future__ import annotations
import argparse
import asyncio
import json
import re
import sys
from pathlib import Path

from playwright.async_api import async_playwright

# reuse the scraper's persistent login profile + sign-in helper
import ancestry_scraper as A

_HERE   = Path(__file__).resolve().parent
_CONFIG = _HERE / "config"


# ── SITE CONFIG (everything site-specific — copy & edit for the next 3 sites) ── #
SITE = {
    "name":         "ancestry",
    "home":         A.HOME_URL,                       # https://www.ancestry.com/
    "search_page":  "https://www.ancestry.com/search/",
    "surname":      "Smith",
    # canonical results URL used as the drill base (reliable; ?query so we can append)
    "results_base": "https://www.ancestry.com/search/?name=_Smith&count=50",

    # Last-Name field + Search button on the search form (best-effort; URL fallback)
    "surname_selectors": ["#lastName", "input[name='lastName']",
                          "input[name='ln']", "input[aria-label*='Last' i]"],
    "search_button":     ["button[type='submit']", "#searchButton",
                          "button:has-text('Search')", "input[type='submit']"],

    # ONE axis: locations, crawled to FULL DEPTH (continent → country → state →
    # county → city → … → last level). Record Date is NOT crawled (not needed).
    # Locations use `record_f=_35`; Record Date uses a RANGE `record_f=1850-1859`,
    # so `exclude` drops those ranges to keep the date filters out of locations.
    "axes": {
        "locations": {
            "json":         "ancestry_locations.json",
            "code_in_href": r"[?&]record_f=([^&#\"']+)",
            "exclude":      r"^\d{4}-\d{4}",          # skip date ranges (those are dates)
            "child_url":    lambda base, code: base + ("&" if "?" in base else "?")
                                               + "record_f=" + code,
            # Crawl 4 levels: continent → country → state → county.
            # Level 5 (city) is INTENTIONALLY skipped (city lists balloon the JSON
            # without adding useful filtering — there are 17k+ US cities alone).
            # Override at the command line with --depth N to go deeper temporarily.
            "max_depth":    4,
            "levels":       ["continent", "country", "state", "county",
                             "city", "locality", "area"],
        },
    },
}

# "see all / show more" controls inside a facet (click to reveal the full list)
SEE_ALL_RE = re.compile(r"^(see all|show all|show more|view all|more)\b", re.I)


def log(msg: str):
    print(msg, flush=True)


def _clean_label(s: str) -> str:
    """Strip the result count and tidy whitespace: «New York (12,345)» → «New York»."""
    s = re.sub(r"\s+", " ", s or "").strip()
    s = re.sub(r"\s*[\(\[][\d.,\s]+[\)\]]\s*$", "", s).strip()
    return s


# ── reading a facet's child options at the current page ───────────────────────── #
async def read_options(page, code_re: str, exclude: str = None):
    """Return [(label, code, href)] for every facet link on the page that carries a
    code (and does NOT match `exclude` — used to keep the date ranges out of the
    location axis and vice-versa). Generic on purpose: the global-visited set in
    crawl_axis() drops the breadcrumb / sibling links, leaving the real children."""
    js = r"""([reSrc, exclSrc]) => {
        const re = new RegExp(reSrc);
        const excl = exclSrc ? new RegExp(exclSrc) : null;
        const seen = new Set(), out = [];
        for (const a of document.querySelectorAll('a[href]')) {
            const href = a.href || '';
            if (!/\/search\//.test(href)) continue;          // facet links are /search/ URLs
            const m = href.match(re);
            if (!m) continue;
            const code = m[1];
            if (excl && excl.test(code)) continue;           // belongs to the other axis
            if (seen.has(code)) continue;
            const label = (a.innerText || a.textContent || '').replace(/\s+/g,' ').trim();
            if (!label || label.length > 60) continue;
            seen.add(code);
            out.push([label, code, href]);
        }
        return out;
    }"""
    try:
        return await page.evaluate(js, [code_re, exclude])
    except Exception as e:
        log(f"    !! read_options failed: {type(e).__name__}: {e}")
        return []


async def expand_see_all(page):
    """Click any «See all / Show more» control inside the filter panel so the full
    option list is in the DOM before we read it. Best-effort, a few passes."""
    for _ in range(4):
        clicked = False
        try:
            links = await page.query_selector_all(
                "a, button, span[role='button'], [data-testid*='see' i]")
            for el in links:
                try:
                    txt = _clean_label(await el.inner_text())
                except Exception:
                    continue
                if txt and SEE_ALL_RE.match(txt) and await el.is_visible():
                    await el.click()
                    clicked = True
                    await page.wait_for_timeout(700)
                    break
        except Exception:
            pass
        if not clicked:
            break


# ── DFS over one axis (locations | dates) ─────────────────────────────────────── #
async def crawl_axis(page, axis: str, args) -> list:
    """Depth-first walk of one filter axis. Returns a flat node list:
       [{code, label, depth, parent}]  (depth 1 = top level).

    RESUMES from the saved raw dump: any parent whose children are already in the
    dump is NOT refetched — we just recurse into them. Deleting the *_crawl_raw.json
    file forces a fresh crawl."""
    cfg = SITE["axes"][axis]
    base = SITE["results_base"]
    max_depth = args.depth or cfg["max_depth"]   # 0 → use cfg default
    visited: set[str] = set()
    nodes: list[dict] = []
    dump_path = _CONFIG / cfg["json"].replace(".json", "_crawl_raw.json")

    # ── Resume from the saved dump ────────────────────────────────────────── #
    # Keep only nodes within the current depth cap; drop deeper ones (e.g. depth-5
    # cities from an earlier deeper run) — the user wants level 5 omitted.
    resumed_dropped = 0
    if dump_path.exists():
        try:
            prev = json.loads(dump_path.read_text("utf-8"))
        except Exception as e:
            log(f"  !! could not parse resume dump ({type(e).__name__}) — starting fresh")
            prev = []
        for n in prev:
            if n.get("depth", 99) <= max_depth and n.get("code"):
                nodes.append(n)
                visited.add(n["code"])
            else:
                resumed_dropped += 1
        if nodes:
            log(f"  resume: loaded {len(nodes)} nodes "
                f"(dropped {resumed_dropped} beyond depth {max_depth})")

    # A parent counts as "done" if its children are already in `nodes`. This skips
    # the goto/read but still recurses, so any sub-tree not yet explored continues.
    done_parents: set = {n.get("parent") for n in nodes if n.get("parent")}
    if nodes:                                    # if anything resumed, ROOT is too
        done_parents.add(None)
    kids_by_parent: dict = {}
    for n in nodes:
        kids_by_parent.setdefault(n.get("parent"), []).append((n["label"], n["code"]))

    async def visit(parent_code, parent_label, depth):
        indent = "  " * depth
        lvl = cfg["levels"][min(depth - 1, len(cfg["levels"]) - 1)]
        # Already fetched this parent on an earlier run → reuse its kids
        if parent_code in done_parents:
            kids = kids_by_parent.get(parent_code, [])
            log(f"{indent}[{axis}] {parent_label or 'ROOT'} → {len(kids)} {lvl}(s) (resumed)")
        else:
            url = cfg["child_url"](base, parent_code) if parent_code else base
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=45000)
            except Exception as e:
                log(f"    !! goto failed ({type(e).__name__}) {url[:90]}")
                return
            await page.wait_for_timeout(args.delay)
            await expand_see_all(page)
            opts = await read_options(page, cfg["code_in_href"], cfg.get("exclude"))
            kids = [(lbl, code) for (lbl, code, _h) in opts if code not in visited]
            log(f"{indent}[{axis}] {parent_label or 'ROOT'} → {len(kids)} {lvl}(s)")
            for lbl, code in kids:
                visited.add(code)
                nodes.append({"code": code, "label": _clean_label(lbl),
                              "depth": depth, "parent": parent_code})
            done_parents.add(parent_code)
            kids_by_parent[parent_code] = kids
            # persist progress so a crash/Ctrl-C still leaves the raw dump
            try:
                dump_path.write_text(
                    json.dumps(nodes, ensure_ascii=False, indent=2), "utf-8")
            except Exception:
                pass
        if depth < max_depth:
            for lbl, code in kids:
                await visit(code, _clean_label(lbl), depth + 1)

    log(f"\n==== crawling {axis} (max depth {max_depth}) ====")
    await visit(None, None, 1)
    log(f"==== {axis}: {len(nodes)} nodes (raw dump → {dump_path.name}) ====")
    return nodes


# ── merge crawled nodes into the JSON files ───────────────────────────────────── #
_CONTINENTS = {"north america", "south america", "europe", "asia", "africa",
               "oceania", "antarctica"}


def merge_locations(nodes: list, dry: bool):
    """nodes depth 1/2/3/4 → continent/country/state/city. Adds new, fills codes,
    NEVER drops existing entries or US city lists. Date-range codes (record_f like
    1850-1859) are skipped — those are Record Date, not locations."""
    jp = _CONFIG / "ancestry_locations.json"
    loc = json.loads(jp.read_text("utf-8")) if jp.exists() else {}
    nodes = [n for n in nodes if not re.match(r"^\d{4}-\d{4}", n.get("code", ""))]
    by_code = {n["code"]: n for n in nodes}
    ctry_cont = {ctry.lower(): cont for cont, cv in loc.items()
                 for ctry in cv.get("countries", {})}
    added = filled = 0

    def chain(n):                                    # root → node  (continent … leaf)
        seq, cur, g = [], n, 0
        while cur and g < 20:
            seq.append(cur); cur = by_code.get(cur.get("parent")); g += 1
        return list(reversed(seq))

    for n in nodes:
        seq = chain(n)
        if not seq:
            continue
        # ── continent (level 1) ──
        cont = seq[0]
        ckey = next((k for k in loc if k.lower() == cont["label"].lower()), cont["label"])
        cv = loc.setdefault(ckey, {"code": cont.get("code"), "countries": {}})
        cv.setdefault("countries", {})
        if cont.get("code") and not cv.get("code"):
            cv["code"] = cont["code"]
        if len(seq) == 1:
            continue
        # ── country (level 2) — keep an existing country under its own continent ──
        country = seq[1]
        host = ctry_cont.get(country["label"].lower(), ckey)
        hv = loc.setdefault(host, {"code": None, "countries": {}}); hv.setdefault("countries", {})
        cd = hv["countries"].get(country["label"])
        if cd is None:
            hv["countries"][country["label"]] = {"code": country.get("code"), "states": {}}
            cd = hv["countries"][country["label"]]
            ctry_cont[country["label"].lower()] = host; added += 1
        elif country.get("code") and not cd.get("code"):
            cd["code"] = country["code"]; filled += 1
        cd.setdefault("states", {})
        if len(seq) == 2:
            continue
        # ── state (level 3) ──
        state = seq[2]
        sd = cd["states"].get(state["label"])
        if sd is None:
            cd["states"][state["label"]] = {"code": state.get("code"), "places": {}}
            sd = cd["states"][state["label"]]; added += 1
        elif state.get("code") and not sd.get("code"):
            sd["code"] = state["code"]; filled += 1
        sd.setdefault("places", {})
        # ── places (level 4+) — nest recursively to ANY depth ──
        node = sd
        for p in seq[3:]:
            places = node.setdefault("places", {})
            v = places.get(p["label"])
            if isinstance(v, dict):
                node = v
                if p.get("code") and not v.get("code"):
                    v["code"] = p["code"]
            elif isinstance(v, str):                 # existing flat city → allow nesting
                places[p["label"]] = {"code": v, "places": {}}; node = places[p["label"]]
            else:
                places[p["label"]] = {"code": p.get("code"), "places": {}}
                node = places[p["label"]]; added += 1

    log(f"locations: +{added} nodes, {filled} codes filled")
    if not dry:
        jp.write_text(json.dumps(loc, ensure_ascii=False, indent=2), "utf-8")
        log(f"  wrote {jp}")


def merge_dates(nodes: list, dry: bool):
    """Record Date nodes carry record_f RANGE codes: century 1500-1599, decade
    1500-1509, year 1851-1851. The level is derived from the code span (robust to
    the ambiguous labels «1500»). Existing century/decade `cen_*` codes are KEPT;
    only individual years (range codes) are added under each decade."""
    jp = _CONFIG / "ancestry_record_dates.json"
    rd = json.loads(jp.read_text("utf-8")) if jp.exists() else {}

    def span(code):
        m = re.match(r"^(\d{4})-(\d{4})$", code or "")
        return (int(m.group(1)), int(m.group(2))) if m else (None, None)

    def decade_dict(century_label, decade_label, fallback_code):
        decs = rd.setdefault(century_label, {"code": None, "decades": {}})\
                 .setdefault("decades", {})
        v = decs.get(decade_label)
        if isinstance(v, str):                      # keep the verified cen_* code
            decs[decade_label] = {"code": v, "years": {}}
        elif v is None:
            decs[decade_label] = {"code": fallback_code, "years": {}}
        return decs[decade_label]

    added_cent = added_dec = added_yr = 0
    for n in nodes:
        s, e = span(n["code"])
        if s is None:
            continue
        width = e - s
        if width >= 90:                              # century
            lab = f"{s}s"
            if lab not in rd:
                rd[lab] = {"code": n["code"], "decades": {}}; added_cent += 1
        elif width >= 9:                             # decade
            clab = f"{(s // 100) * 100}s"
            dlab = f"{s}s"
            decs = rd.setdefault(clab, {"code": None, "decades": {}})\
                     .setdefault("decades", {})
            if dlab not in decs:
                decs[dlab] = {"code": n["code"], "years": {}}; added_dec += 1
            elif isinstance(decs[dlab], str):
                decs[dlab] = {"code": decs[dlab], "years": {}}
        else:                                        # individual year
            clab = f"{(s // 100) * 100}s"
            dlab = f"{(s // 10) * 10}s"
            dd = decade_dict(clab, dlab, None)
            if str(s) not in dd["years"]:
                dd["years"][str(s)] = n["code"]; added_yr += 1

    log(f"dates: +{added_cent} centuries, +{added_dec} decades, +{added_yr} years")
    if not dry:
        jp.write_text(json.dumps(rd, ensure_ascii=False, indent=2), "utf-8")
        log(f"  wrote {jp}")


# ── search (fill Last Name = Smith, press Search; URL fallback) ───────────────── #
async def do_search(page):
    log("Opening search form…")
    try:
        await page.goto(SITE["search_page"], wait_until="domcontentloaded", timeout=45000)
        await page.wait_for_timeout(2500)
    except Exception:
        pass
    filled = False
    for sel in SITE["surname_selectors"]:
        try:
            el = await page.query_selector(sel)
            if el:
                await el.fill(SITE["surname"])
                filled = True
                log(f"  typed «{SITE['surname']}» into {sel}")
                break
        except Exception:
            continue
    if filled:
        for sel in SITE["search_button"]:
            try:
                b = await page.query_selector(sel)
                if b and await b.is_visible():
                    await b.click()
                    log(f"  clicked search ({sel})")
                    break
            except Exception:
                continue
        try:
            await page.wait_for_load_state("domcontentloaded", timeout=30000)
        except Exception:
            pass
        await page.wait_for_timeout(3000)
    # make sure we're on a Smith results page; otherwise use the canonical URL
    if "/search/" not in page.url or "name=" not in page.url:
        log("  form path unreliable → using canonical results URL")
        try:
            await page.goto(SITE["results_base"], wait_until="domcontentloaded", timeout=45000)
            await page.wait_for_timeout(3000)
        except Exception:
            pass
    log(f"  results page: {page.url[:100]}")


async def main_async(args):
    profile = A.ANC_PROFILE_DIR
    async with async_playwright() as pw:
        for lk in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            try: (profile / lk).unlink()
            except Exception: pass
        profile.mkdir(parents=True, exist_ok=True)
        ctx = await pw.chromium.launch_persistent_context(
            str(profile),
            headless=args.headless,
            no_viewport=not args.headless,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for p in list(ctx.pages)[1:]:
            try: await p.close()
            except Exception: pass

        all_nodes = {}
        try:
            await page.goto(SITE["home"], wait_until="domcontentloaded", timeout=45000)
            await page.wait_for_timeout(3000)
            logged = [False]
            await A._sign_in_if_needed(page, args.user or "", args.password or "",
                                       logged, log)
            await do_search(page)

            all_nodes["locations"] = await crawl_axis(page, "locations", args)
        finally:
            # merge whatever was collected, even on Ctrl-C / error
            if all_nodes.get("locations"):
                merge_locations(all_nodes["locations"], args.dry_run)
            try: await ctx.close()
            except Exception: pass


def main():
    ap = argparse.ArgumentParser(description="Crawl Ancestry's full location filter tree into the JSON.")
    ap.add_argument("--depth", type=int, default=0,
                    help="max location depth (0 = full recursion, the default)")
    ap.add_argument("--dry-run", action="store_true", help="crawl + print, do NOT write JSON")
    ap.add_argument("--headless", action="store_true", help="run without a visible window")
    ap.add_argument("--delay", type=int, default=1800, help="ms to wait after each page load")
    ap.add_argument("--user", default=None, help="Ancestry username (if not already logged in)")
    ap.add_argument("--password", default=None, help="Ancestry password (if not already logged in)")
    args = ap.parse_args()
    try:
        asyncio.run(main_async(args))
    except KeyboardInterrupt:
        log("\ninterrupted — partial results were merged/dumped.")
        sys.exit(0)


if __name__ == "__main__":
    main()
