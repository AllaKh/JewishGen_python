#!/usr/bin/env python3
"""
familysearch_scraper.py — v11
================================
Flow:
1.  Open home page, fill First/Last Names (+ Place, Birth Year), click SEARCH.
2.  Wait for results. Click "Historical Records" tab (data-testid="hr-tab").
3.  Set results per page = 60.
4.  Collect rows, filter name-match >= 80%.
5.  Click FIRST qualifying result.
    FamilySearch redirects to sign-in — complete login — navigate back to record.
6.  Scrape first record: all text fields + full JPG image download via viewer.
7.  Go back to results page.
8a. NO advanced fields: continue scraping remaining qualifying results (>=80%).
8b. HAS advanced fields: REFRESH PAGE (mandatory!), click Advanced Search
    (data-testid="advanced-search-form-button"), fill modal, search,
    re-collect results, scrape ALL.
9.  Write Word (.docx) named "{FirstNames} {LastNames}.docx" with text + image.
    Write Excel (.xlsx) named "{FirstNames} {LastNames}.xlsx" with text only.
"""

import asyncio, difflib, io, json, os, re, shutil, sys
from pathlib import Path

if getattr(sys, "frozen", False):
    bd = Path(sys.executable).resolve().parent / "ms-playwright"
    if bd.exists():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(bd)

try:
    from playwright.async_api import async_playwright
except ImportError:
    sys.stderr.write("pip install playwright && playwright install chromium\n")
    sys.exit(1)

try:
    from docx import Document
    from docx.shared import Mm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# ── Config ───────────────────────────────────────────────────────────────────── #
_HERE     = Path(__file__).resolve().parent
_CFG_PATH = _HERE / "config" / "familysearch.json"
try:
    _CFG = json.loads(_CFG_PATH.read_text(encoding="utf-8"))
except Exception:
    _CFG = {}

HOME_URL      = _CFG.get("home_url", "https://www.familysearch.org/en/global")
MIN_MATCH     = int(_CFG.get("min_match", 80))
BAD_PATHS     = _CFG.get("bad_paths", [
    "/records/images", "/search/linker", "/linker",
    "/en/tree/", "/tree/person/", "/tree/",
    "/catalog", "/wiki", "/books", "/films",
])
_dl           = _CFG.get("downloads_dir", "")
DOWNLOADS_DIR = Path(_dl) if _dl else Path.home() / "Downloads"
HYPERLINK_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)


# ── Helpers ──────────────────────────────────────────────────────────────────── #

def _sim(a: str, b: str) -> float:
    a = re.sub(r"\s+", " ", a.strip().lower())
    b = re.sub(r"\s+", " ", b.strip().lower())
    if not a or not b:
        return 0.0
    wa, wb = set(a.split()), set(b.split())
    return max(len(wa & wb) / max(len(wa), 1),
               difflib.SequenceMatcher(None, a, b).ratio()) * 100


def safe_fn(s: str, n: int = 100) -> str:
    s = re.sub(r'[\\/*?:"<>|]', "_", s.strip())
    return re.sub(r"\s+", " ", s)[:n].strip() or "document"


def _is_record(href: str) -> bool:
    if not href:
        return False
    for bad in BAD_PATHS:
        if bad in href:
            return False
    return "/ark:" in href


# ── Type into a form field ───────────────────────────────────────────────────── #

async def _type(page, sels: list, val: str, label: str, log) -> bool:
    if not val:
        return True
    for sel in sels:
        try:
            el = page.locator(sel).first
            if not await el.count():
                continue
            await el.scroll_into_view_if_needed(timeout=4000)
            await page.evaluate(
                "s => { const e=document.querySelector(s); if(e) e.focus(); }", sel)
            await asyncio.sleep(0.1)
            await el.click(timeout=3000)
            await page.keyboard.press("Control+a")
            await page.keyboard.press("Delete")
            await asyncio.sleep(0.1)
            await page.keyboard.type(val, delay=40)
            await asyncio.sleep(0.2)
            if (await el.input_value(timeout=2000)).strip():
                log(f"  OK  {label}")
                return True
        except Exception:
            continue
    log(f"  !!  field not found: {label}")
    return False


# ── Step 1: Search from home page ────────────────────────────────────────────── #

async def _search_from_home(page, first_names, last_names,
                             place_lived, birth_year, log) -> None:
    log(f"  Opening: {HOME_URL}")
    await page.goto(HOME_URL, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(3)

    # Wait for the search form to appear
    for sel in ['input[placeholder*="First and Middle" i]',
                'input[id*="givenName" i]']:
        try:
            await page.locator(sel).first.wait_for(state="visible", timeout=10000)
            break
        except Exception:
            continue

    await _type(page,
        ['input[placeholder*="First and Middle" i]',
         'input[id*="givenName" i]', 'input[name*="givenName" i]'],
        first_names, "First Names", log)

    await _type(page,
        ['input[placeholder*="Last or Maiden" i]',
         'input[id*="surname" i]', 'input[name*="surname" i]'],
        last_names, "Last Names", log)

    if place_lived:
        await _type(page,
            ['input[placeholder*="City, County, State" i]'],
            place_lived, "Place Lived", log)

    if birth_year:
        await _type(page,
            ['input[placeholder*="Birth Year" i]',
             'input[id*="birthYear" i]'],
            birth_year, "Birth Year", log)

    for sel in ['button:has-text("SEARCH")', 'button:has-text("Search")',
                'button[type="submit"]']:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=6000)
                log("  SEARCH clicked")
                break
        except Exception:
            continue

    try:
        await page.wait_for_url(
            lambda u: "search" in u and "discovery" in u, timeout=20000)
    except Exception:
        pass
    try:
        await page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        await asyncio.sleep(4)
    log(f"  Results page: {page.url}")


# ── Click Historical Records tab ─────────────────────────────────────────────── #

async def _click_historical_tab(page, log) -> None:
    # Prefer exact data-testid (from user's HTML inspection)
    for sel in ['[data-testid="hr-tab"]', 'div[data-testid="hr-tab"]']:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=5000)
                await asyncio.sleep(2)
                log("  Historical Records tab clicked (data-testid=hr-tab)")
                return
        except Exception:
            continue

    # Fallback: by role or text
    for label in ("Historical Records", "Historical records"):
        try:
            el = page.get_by_role("tab", name=re.compile(re.escape(label), re.I)).first
            if not await el.count():
                el = page.locator(
                    f'[role="tab"]:has-text("{label}"), '
                    f'a:has-text("{label}"), button:has-text("{label}")'
                ).first
            if await el.count():
                await el.click(timeout=5000)
                await asyncio.sleep(2)
                log(f"  Historical Records tab clicked (text fallback)")
                return
        except Exception:
            pass
    log("  (Historical Records tab not found — continuing)")


# ── Set 60 per page ──────────────────────────────────────────────────────────── #

async def _set_60(page, log) -> None:
    for sel in ['select[aria-label*="result" i]', 'select[name*="result" i]',
                'select[id*="result" i]', 'select']:
        try:
            el = page.locator(sel).last
            if await el.count():
                opts = await el.evaluate(
                    "e => Array.from(e.options).map(o => o.value)")
                if "60" in opts:
                    await el.select_option(value="60")
                    await asyncio.sleep(2)
                    log("  Results per page: 60")
                    return
        except Exception:
            continue
    try:
        el = page.get_by_text(re.compile(r"^60$")).first
        if await el.count():
            await el.click(timeout=3000)
            await asyncio.sleep(2)
    except Exception:
        pass


# ── Collect result rows ──────────────────────────────────────────────────────── #

async def _collect(page, qname: str, log) -> list:
    await asyncio.sleep(2)
    results, seen = [], set()
    rows = await page.query_selector_all("tbody tr")
    log(f"  Rows found: {len(rows)}")
    for row in rows:
        try:
            lnk = None
            for a in await row.query_selector_all("a[href]"):
                href = await a.get_attribute("href") or ""
                if _is_record(href):
                    lnk = a
                    break
            if not lnk:
                continue
            href = await lnk.get_attribute("href") or ""
            name = (await lnk.text_content() or "").strip()
            if not href or not name:
                continue
            url = ("https://www.familysearch.org" + href
                   if href.startswith("/") else href)
            if url in seen:
                continue
            seen.add(url)
            cells = await row.query_selector_all("td")
            results.append({
                "url":   url,
                "name":  name,
                "coll":  (await cells[1].text_content() or "").strip() if len(cells) > 1 else "",
                "evts":  (await cells[2].text_content() or "").strip() if len(cells) > 2 else "",
                "rels":  (await cells[3].text_content() or "").strip() if len(cells) > 3 else "",
                "score": round(_sim(qname, name), 1),
            })
        except Exception:
            continue
    log(f"  Candidates: {len(results)}")
    return results


# ── Sign-in ──────────────────────────────────────────────────────────────────── #

async def _do_login(page, username: str, password: str, log) -> bool:
    log("  Filling sign-in form...")
    await asyncio.sleep(1.5)

    found = False
    for attempt in range(3):
        try:
            await page.locator('#userName').wait_for(state="visible", timeout=8000)
            found = True
            break
        except Exception:
            log(f"  Waiting for sign-in form... attempt {attempt + 1}")
            await asyncio.sleep(2)

    if not found:
        log("  !! Sign-in form (#userName) not found")
        return False

    # Use JavaScript to set values reliably (React controlled inputs)
    try:
        await page.evaluate("""([u, p]) => {
            function setVal(sel, val) {
                const el = document.querySelector(sel);
                if (!el) return false;
                const nativeSetter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value').set;
                nativeSetter.call(el, val);
                el.dispatchEvent(new Event('input',  {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
                return true;
            }
            setVal('#userName', u);
            setVal('#password', p);
        }""", [username, password])
        log("  Credentials set via JS")
        await asyncio.sleep(0.5)
    except Exception as exc:
        log(f"  JS set failed ({exc}), using keyboard...")
        for sel, val, lbl in [
            ('#userName', username, 'username'),
            ('#password', password, 'password'),
        ]:
            try:
                el = page.locator(sel).first
                await el.click(timeout=3000)
                await page.keyboard.press("Control+a")
                await page.keyboard.press("Delete")
                await page.keyboard.type(val, delay=50)
                log(f"  Typed {lbl}")
            except Exception as e2:
                log(f"  !! Could not type {lbl}: {e2}")
                return False

    # Click SIGN IN
    clicked = False
    for sel in ['button:has-text("SIGN IN")', 'button:has-text("Sign In")',
                'button:has-text("Sign in")', 'button[type="submit"]']:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=6000)
                log(f"  Sign-in submitted ({sel})")
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        await page.keyboard.press("Enter")
        log("  Sign-in via Enter")

    try:
        await page.wait_for_url(
            lambda u: (
                "familysearch.org" in u
                and "/identity/login" not in u
                and "/auth/familysearch/login" not in u
            ),
            timeout=30000,
        )
    except Exception:
        await asyncio.sleep(5)

    cur = page.url
    ok  = ("familysearch.org" in cur
           and "/identity/login" not in cur
           and "/auth/familysearch/login" not in cur)
    log(f"  {'Sign-in OK' if ok else 'SIGN-IN FAILED'}. URL: {cur}")
    return ok


# ── Advanced Search modal ────────────────────────────────────────────────────── #

async def _advanced_search(page, adv: dict, log) -> None:
    log("  Opening Advanced Search...")

    # Exact data-testid provided by user, with text fallbacks
    opened = False
    for sel in [
        '[data-testid="advanced-search-form-button"]',
        'button:has-text("Advanced Search")',
        'button:has-text("ADVANCED SEARCH")',
        'a:has-text("Advanced Search")',
    ]:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=5000)
                await asyncio.sleep(1.5)
                log(f"  Advanced Search modal opened ({sel[:60]})")
                opened = True
                break
        except Exception:
            continue

    if not opened:
        log("  !! Advanced Search button not found")
        return

    # Life Events (tabs: BIRTH, DEATH, MARRIAGE, RESIDENCE, ANY)
    event_map = {
        "birth_place":     ("BIRTH",     'input[placeholder*="City, County, State, Province"]'),
        "birth_year":      ("BIRTH",     'input[placeholder="Year"]'),
        "death_place":     ("DEATH",     'input[placeholder*="City, County, State, Province"]'),
        "death_year":      ("DEATH",     'input[placeholder="Year"]'),
        "marriage_place":  ("MARRIAGE",  'input[placeholder*="City, County, State, Province"]'),
        "marriage_year":   ("MARRIAGE",  'input[placeholder="Year"]'),
        "residence_place": ("RESIDENCE", 'input[placeholder*="City, County, State, Province"]'),
        "residence_year":  ("RESIDENCE", 'input[placeholder="Year"]'),
        "any_place":       ("ANY",       'input[placeholder*="City, County, State, Province"]'),
        "any_year":        ("ANY",       'input[placeholder="Year"]'),
    }
    opened_tabs: set = set()
    for key, (tab_label, input_sel) in event_map.items():
        val = adv.get(key, "")
        if not val:
            continue
        if tab_label not in opened_tabs:
            try:
                t = page.get_by_text(
                    re.compile(rf"^{re.escape(tab_label)}$", re.I)).first
                if await t.count():
                    await t.click(timeout=3000)
                    await asyncio.sleep(0.7)
                    opened_tabs.add(tab_label)
            except Exception:
                pass
        await _type(page, [input_sel], val, key, log)

    # Family Members (tabs: SPOUSE, FATHER, MOTHER, OTHER PERSON)
    fam_map = {
        "spouse": ("SPOUSE",       "Spouse's First Names",       "Spouse's Last Names"),
        "father": ("FATHER",       "Father's First Names",       "Father's Last Names"),
        "mother": ("MOTHER",       "Mother's First Names",       "Mother's Last Names"),
        "other":  ("OTHER PERSON", "Other Person's First Names", "Other Person's Last Names"),
    }
    for key, (tab_label, first_ph, last_ph) in fam_map.items():
        fv = adv.get(f"{key}_first", "")
        lv = adv.get(f"{key}_last",  "")
        if not fv and not lv:
            continue
        try:
            t = page.get_by_text(
                re.compile(rf"^{re.escape(tab_label)}$", re.I)).first
            if await t.count():
                await t.click(timeout=3000)
                await asyncio.sleep(0.7)
        except Exception:
            pass
        if fv:
            await _type(page, [f'input[placeholder="{first_ph}"]'], fv, f"{key} first", log)
        if lv:
            await _type(page, [f'input[placeholder="{last_ph}"]'], lv, f"{key} last", log)

    # Location (country / state)
    if adv.get("country"):
        try:
            t = page.get_by_text(re.compile(r"^LOCATION$", re.I)).first
            if await t.count():
                await t.click(timeout=3000)
                await asyncio.sleep(0.5)
        except Exception:
            pass
        await _type(page, ['input[placeholder="Country or Location"]'],
                   adv["country"], "country", log)
    if adv.get("state"):
        await _type(page, ['input[placeholder="State or Province"]'],
                   adv["state"], "state", log)

    # Keywords
    if adv.get("keywords"):
        await _type(page, ['input[placeholder*="keyword" i]'],
                   adv["keywords"], "keywords", log)

    # Submit inside modal (use .last to target modal button, not main search bar)
    for sel in ['button:has-text("SEARCH")', 'button:has-text("Search")']:
        try:
            el = page.locator(sel).last
            if await el.count() and await el.is_visible():
                await el.click(timeout=5000)
                break
        except Exception:
            continue

    try:
        await page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        await asyncio.sleep(3)
    log("  Advanced Search submitted.")


# ── Thumbnail bytes (for Word embed, no viewer) ──────────────────────────────── #

async def _get_thumb_bytes(ctx, page) -> bytes | None:
    """Fetch thumbnail image bytes for embedding in Word document."""
    SKIP = ("icon", "logo", "sprite", "avatar", "pixel", "button", "badge")
    for sel in [
        'img[alt="Thumbnail"]',
        'img[class*="imageThumb"]',
        'img[src*="dz/v1"]',
        'img[src*="/image/"]',
        '.imageViewer img',
        '[class*="thumbnail" i] img',
        'main img[src*="http"]',
    ]:
        try:
            for el in await page.query_selector_all(sel):
                src = (await el.get_attribute("src") or "").strip()
                if not src.startswith("http"):
                    continue
                if any(b in src.lower() for b in SKIP):
                    continue
                ip = await ctx.new_page()
                try:
                    r = await ip.goto(src, timeout=10000)
                    if r and r.ok:
                        body = await r.body()
                        if len(body) > 5000:
                            return body
                finally:
                    await ip.close()
        except Exception:
            continue
    return None


# ── Download full-resolution JPG from viewer ─────────────────────────────────── #

async def _download_full_image(ctx, page, dest_dir: Path,
                                title: str, log) -> str | None:
    """
    1. Click thumbnail → viewer opens (same page or new tab).
    2. Click the download arrow button in the viewer toolbar.
    3. Select JPG Only in popup.
    4. Click DOWNLOAD (data-testid="full-text-confirm-download").
    Falls back to saving thumbnail bytes if any step fails.
    """
    dest_dir.mkdir(parents=True, exist_ok=True)
    fname  = safe_fn(title) + ".jpg"
    dest   = dest_dir / fname
    before = set(DOWNLOADS_DIR.glob("*.jpg")) | set(DOWNLOADS_DIR.glob("*.jpeg"))
    pages_before = set(ctx.pages)

    # Click the thumbnail image to open full viewer
    thumb_clicked = False
    for sel in [
        'img[alt="Thumbnail"]',
        'img[class*="imageThumb"]',
        'img[src*="dz/v1"]',
        'img[src*="/image/"]',
        'a > img[src*="familysearch"]',
        '.imageViewer img',
        '[class*="thumbnail" i] img',
        'main img[src*="http"]',
    ]:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=5000)
                await asyncio.sleep(3)
                log(f"    Thumbnail clicked ({sel})")
                thumb_clicked = True
                break
        except Exception:
            continue

    if not thumb_clicked:
        log("    No thumbnail found — skipping full image download")
        return None

    # Viewer may open in a new tab
    viewer = page
    await asyncio.sleep(1)
    new_pg = set(ctx.pages) - pages_before
    if new_pg:
        viewer = list(new_pg)[0]
        await asyncio.sleep(2)
        log("    Viewer opened in new tab")

    # Click the download arrow button in the viewer toolbar
    dl_clicked = False
    for sel in [
        'button[aria-label*="Download" i]',
        'button[title*="Download" i]',
        '[data-testid*="download" i]:not([data-testid="full-text-confirm-download"])',
        'button[class*="download" i]',
        '[class*="toolbar"] button:nth-last-child(3)',
        '[class*="toolbar"] button:nth-last-child(2)',
        '[class*="tools"] button:nth-last-child(3)',
        '[class*="tools"] button:nth-last-child(2)',
    ]:
        try:
            el = viewer.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=4000)
                await asyncio.sleep(1.5)
                dl_clicked = True
                log(f"    Download button clicked ({sel})")
                break
        except Exception:
            continue

    if not dl_clicked:
        log("    Download button not found — saving thumbnail preview instead")
        if viewer is not page:
            try: await viewer.close()
            except Exception: pass
        return await _save_thumb_preview(ctx, page, dest_dir, title, log)

    # Select "JPG Only" in the popup
    for lbl in ("JPG Only", "JPG only", "JPG"):
        try:
            el = viewer.get_by_text(lbl, exact=True).first
            if await el.count():
                await el.click(timeout=3000)
                await asyncio.sleep(0.5)
                log("    JPG Only selected")
                break
        except Exception:
            continue

    # Click DOWNLOAD button — use exact data-testid first
    downloaded = None
    try:
        async with viewer.expect_download(timeout=30000) as dl_info:
            for sel in [
                '[data-testid="full-text-confirm-download"]',
                'button:has-text("DOWNLOAD")',
                'button:has-text("Download")',
            ]:
                try:
                    btn = viewer.locator(sel).first
                    if await btn.count() and await btn.is_visible():
                        await btn.click(timeout=5000)
                        log(f"    Download confirmed ({sel})")
                        break
                except Exception:
                    continue
        dl = await dl_info.value
        await dl.save_as(str(dest))
        downloaded = str(dest)
        log(f"    Saved: {fname}")
    except Exception as exc:
        log(f"    expect_download failed ({exc}) — watching Downloads folder...")
        for _ in range(15):
            await asyncio.sleep(1)
            after = set(DOWNLOADS_DIR.glob("*.jpg")) | set(DOWNLOADS_DIR.glob("*.jpeg"))
            new_f = after - before
            if new_f:
                src_f = max(new_f, key=lambda p: p.stat().st_mtime)
                shutil.move(str(src_f), str(dest))
                downloaded = str(dest)
                log(f"    Moved from Downloads: {fname}")
                break

    # Close stray non-FamilySearch tabs (Adobe Express, Photos, etc.)
    await asyncio.sleep(1)
    for pg in list(ctx.pages):
        if pg not in (page, viewer) and "familysearch" not in pg.url:
            try: await pg.close()
            except Exception: pass
    if viewer is not page:
        try: await viewer.close()
        except Exception: pass

    if downloaded:
        return downloaded

    # Last resort: save thumbnail bytes
    return await _save_thumb_preview(ctx, page, dest_dir, title, log)


async def _save_thumb_preview(ctx, page, dest_dir: Path,
                               title: str, log) -> str | None:
    """Download thumbnail and save as *_preview.jpg (fallback)."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / (safe_fn(title) + "_preview.jpg")
    SKIP = ("icon", "logo", "sprite", "avatar", "pixel", "button", "badge")
    for sel in [
        'img[alt="Thumbnail"]',
        'img[class*="imageThumb"]',
        'img[src*="dz/v1"]',
        'img[src*="/image/"]',
        '.imageViewer img',
        '[class*="thumbnail" i] img',
        'main img[src*="http"]',
    ]:
        try:
            for el in await page.query_selector_all(sel):
                src = (await el.get_attribute("src") or "").strip()
                if not src.startswith("http"):
                    continue
                if any(b in src.lower() for b in SKIP):
                    continue
                ip = await ctx.new_page()
                try:
                    r = await ip.goto(src, timeout=15000)
                    if r and r.ok:
                        body = await r.body()
                        if len(body) > 5000:
                            dest.write_bytes(body)
                            log(f"    Preview saved: {dest.name} ({len(body)//1024}KB)")
                            return str(dest)
                finally:
                    await ip.close()
        except Exception:
            continue
    return None


# ── Scrape one record page ───────────────────────────────────────────────────── #

async def _scrape_record(ctx, url: str, name_hint: str,
                          images_root: Path, log) -> dict:
    for bad in BAD_PATHS:
        if bad in url:
            log(f"  Skip bad URL: {url[:80]}")
            return _empty(url, name_hint)

    rec  = _empty(url, name_hint)
    page = await ctx.new_page()
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=25000)
        await asyncio.sleep(2)

        # Page title
        for sel in ["h1", '[class*="title" i]', "h2"]:
            try:
                t = (await page.locator(sel).first
                     .text_content(timeout=3000) or "").strip()
                if 3 < len(t) < 300:
                    rec["title"] = t
                    break
            except Exception:
                pass

        # Structured data: dl/dt/dd first, then table rows
        td: dict = {}
        try:
            dts = await page.query_selector_all("dl dt")
            dds = await page.query_selector_all("dl dd")
            for dt, dd in zip(dts, dds):
                k = (await dt.text_content() or "").strip().rstrip(":")
                v = (await dd.text_content() or "").strip()
                if k and v:
                    td[k] = v
        except Exception:
            pass
        if not td:
            try:
                for row in await page.query_selector_all("table tr"):
                    cells = await row.query_selector_all("td, th")
                    if len(cells) >= 2:
                        k = (await cells[0].text_content() or "").strip().rstrip(":")
                        v = (await cells[1].text_content() or "").strip()
                        if k and v:
                            td[k] = v
            except Exception:
                pass
        rec["table_data"] = td

        # Links on page
        links: list = []
        try:
            for a in await page.query_selector_all("a[href]"):
                href = (await a.get_attribute("href") or "").strip()
                text = (await a.text_content() or "").strip()
                if href.startswith("http") and text and len(text) < 200:
                    full = ("https://www.familysearch.org" + href
                            if href.startswith("/") else href)
                    links.append(f"{text}: {full}")
        except Exception:
            pass
        rec["links"] = links

        # Build meaningful image/folder name from record data
        img_label = name_hint
        if td:
            parts = [name_hint]
            for k in ("Event Type", "Type", "Event"):
                if td.get(k):
                    parts.append(td[k])
                    break
            for k in ("Event Date", "Date", "Death Date", "Birth Date",
                      "Marriage Date", "Naturalization Date"):
                if td.get(k):
                    parts.append(td[k])
                    break
            if len(parts) > 1:
                img_label = " — ".join(parts)

        img_dir = images_root / safe_fn(img_label)

        # Grab thumbnail bytes for embedding in Word
        rec["thumb_bytes"] = await _get_thumb_bytes(ctx, page)

        # Download full-resolution JPG via viewer
        img_path = await _download_full_image(ctx, page, img_dir, img_label, log)
        rec["images"] = [img_path] if img_path else []

    except Exception as exc:
        log(f"    !! Record error: {exc}")
    finally:
        await page.close()
    return rec


def _empty(url, name):
    return {"url": url, "title": name, "name": name,
            "table_data": {}, "links": [], "images": [],
            "thumb_bytes": None,
            "events": "", "relationships": "", "collection": ""}


# ── Word output ──────────────────────────────────────────────────────────────── #

def _add_hyperlink(para, text, url):
    rid = para.part.relate_to(url, HYPERLINK_REL, is_external=True)
    hl  = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), rid)
    run = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    c   = OxmlElement("w:color"); c.set(qn("w:val"), "0563C1")
    u   = OxmlElement("w:u");     u.set(qn("w:val"), "single")
    rPr.append(c); rPr.append(u); run.append(rPr)
    t   = OxmlElement("w:t"); t.text = text or url
    t.set(qn("xml:space"), "preserve")
    run.append(t); hl.append(run); para._p.append(hl)


def write_docx(path: Path, records: list, qlines: list) -> None:
    if not _DOCX_OK:
        raise RuntimeError("python-docx not installed")
    doc = Document()
    sec = doc.sections[0]
    sec.page_width  = Mm(297); sec.page_height = Mm(210)
    sec.left_margin = sec.right_margin = Mm(18)
    sec.top_margin  = sec.bottom_margin = Mm(15)
    h = doc.add_heading("FamilySearch Search Results", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Search parameters:")
    for ln in qlines:
        doc.add_paragraph(ln, style="List Bullet")
    doc.add_paragraph(f"Records found: {len(records)}  (match >= {MIN_MATCH}%)")
    doc.add_paragraph("")

    for i, rec in enumerate(records, 1):
        title = rec.get("title") or rec.get("name", "—")
        doc.add_heading(f"{i}. {title}", level=2)

        p = doc.add_paragraph()
        p.add_run("Match: ").bold = True
        p.add_run(f"{rec.get('score', '?')}%")

        for label, key in [("Collection",     "collection"),
                            ("Events",         "events"),
                            ("Relationships",  "relationships")]:
            if rec.get(key):
                pp = doc.add_paragraph()
                pp.add_run(f"{label}: ").bold = True
                pp.add_run(rec[key])

        if rec.get("url"):
            pp = doc.add_paragraph()
            pp.add_run("Source: ").bold = True
            _add_hyperlink(pp, rec["url"], rec["url"])

        td = rec.get("table_data", {})
        if td:
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            hdr[0].text = "Field"; hdr[1].text = "Value"
            for cell in hdr:
                for r in cell.paragraphs[0].runs:
                    r.bold = True
            for f, v in td.items():
                row = tbl.add_row().cells
                row[0].text = str(f)
                row[1].text = str(v)

        links = rec.get("links", [])
        if links:
            doc.add_paragraph("")
            pl = doc.add_paragraph()
            pl.add_run("Page links:").bold = True
            for lnk in links[:20]:
                parts = lnk.split(": ", 1)
                p_l = doc.add_paragraph(style="List Bullet")
                if len(parts) == 2:
                    _add_hyperlink(p_l, parts[0], parts[1])
                else:
                    p_l.add_run(lnk)

        # Image: prefer downloaded full-res JPG, fall back to thumbnail bytes
        doc.add_paragraph("")
        imgs = rec.get("images", [])
        tb   = rec.get("thumb_bytes")

        if imgs and Path(imgs[0]).exists():
            doc.add_paragraph("Document image:").runs[0].bold = True
            doc.add_paragraph(Path(imgs[0]).name, style="List Bullet")
            try:
                if Path(imgs[0]).suffix.lower() in (".jpg", ".jpeg", ".png"):
                    doc.add_picture(imgs[0], width=Inches(5))
            except Exception:
                pass
        elif tb:
            doc.add_paragraph("Document preview (thumbnail):").runs[0].bold = True
            try:
                doc.add_picture(io.BytesIO(tb), width=Inches(5))
            except Exception:
                pass

        doc.add_paragraph("")
    doc.save(path)


# ── Excel output ─────────────────────────────────────────────────────────────── #

def write_xlsx(path: Path, records: list, qlines: list) -> None:
    if not _OPENPYXL_OK:
        raise RuntimeError("openpyxl not installed")
    wb = Workbook(); ws = wb.active; ws.title = "FamilySearch"
    HF = PatternFill("solid", fgColor="006B6B")
    HN = Font(bold=True, color="FFFFFF", size=11)
    TS = Side(style="thin", color="B0C8C8")
    T  = Border(left=TS, right=TS, top=TS, bottom=TS)

    # Collect all unique table_data field names across all records
    aff: list = []
    for rec in records:
        for k in rec.get("table_data", {}):
            if k not in aff:
                aff.append(k)

    cols = ["#", "Title", "Match %", "Collection", "Events",
            "Relationships", "Image file", "URL"] + aff

    for ci, cn in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=cn)
        c.font = HN; c.fill = HF; c.border = T
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for ri, rec in enumerate(records, 2):
        td   = rec.get("table_data", {})
        imgs = "\n".join(Path(p).name for p in rec.get("images", []))
        vals = [
            ri - 1,
            rec.get("title", rec.get("name", "")),
            rec.get("score", ""),
            rec.get("collection", ""),
            rec.get("events", ""),
            rec.get("relationships", ""),
            imgs,
            rec.get("url", ""),
        ] + [td.get(f, "") for f in aff]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = T
            c.alignment = Alignment(wrap_text=True, vertical="top")

    for ci in range(1, len(cols) + 1):
        letter = get_column_letter(ci)
        mw = max(
            len(str(cols[ci - 1])),
            *(len(str(ws.cell(row=r, column=ci).value or "").split("\n")[0])
              for r in range(2, ws.max_row + 1)),
            8,
        )
        ws.column_dimensions[letter].width = min(mw + 4, 60)
    wb.save(path)


# ── Main entry point ─────────────────────────────────────────────────────────── #

async def run_scraper(
    *,
    first_names:   str       = "",
    last_names:    str       = "",
    place_lived:   str       = "",
    birth_year:    str       = "",
    tab:           str       = "Historical Records",
    advanced:      dict|None = None,
    output_format: str       = "both",
    output_folder            = Path("."),
    email:         str|None  = None,
    password:      str|None  = None,
    log                      = print,
    progress                 = None,
    cancel_event             = None,
) -> dict:

    def _prog(pct, txt):
        log(txt)
        if progress:
            progress(pct, txt)

    def _done():
        return bool(cancel_event and cancel_event.is_set())

    adv       = advanced or {}
    has_adv   = any(v for v in adv.values() if v and v not in (False, "Unspecified"))
    want_docx = output_format in ("docx", "both")
    want_xlsx = output_format in ("xlsx", "both")
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    images_root = output_folder / "images"

    qname  = " ".join(p for p in (first_names, last_names) if p)
    qlines = [ln for ln in [
        f"First Names: {first_names}",
        f"Last Names: {last_names}",
        f"Place Lived: {place_lived}",
        f"Birth Year: {birth_year}",
    ] if not ln.endswith(": ")]
    summary = {"ok": False}

    # Output file base name: "FirstNames LastNames" (not "familysearch_...")
    file_base = safe_fn(qname) if qname else "familysearch_results"

    _prog(0, "Launching browser...")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=["--start-maximized",
                  "--disable-blink-features=AutomationControlled"],
        )
        ctx  = await browser.new_context(no_viewport=True, accept_downloads=True)
        page = await ctx.new_page()

        try:
            # 1. Search from home page
            _prog(5, "Searching on home page...")
            await _search_from_home(page, first_names, last_names,
                                    place_lived, birth_year, log)
            if _done(): return summary

            # 2. Click Historical Records tab
            _prog(10, "Selecting Historical Records tab...")
            await _click_historical_tab(page, log)
            await asyncio.sleep(2)
            if _done(): return summary

            # 3. Set 60 results per page
            _prog(13, "Setting 60 results per page...")
            await _set_60(page, log)

            # 4. Collect preliminary results
            _prog(17, "Collecting results...")
            raw       = await _collect(page, qname, log)
            qualified = [r for r in raw if r["score"] >= MIN_MATCH]
            log(f"  Qualified (>= {MIN_MATCH}%): {len(qualified)}")

            if not qualified:
                _prog(100, "No results above match threshold.")
                summary.update({"ok": True, "n_records": 0,
                                "message": f"No records >= {MIN_MATCH}%."})
                return summary

            # 5. Click first result — may redirect to sign-in
            _prog(22, "Opening first result...")
            first_url  = qualified[0]["url"]
            first_name = qualified[0]["name"]
            await page.goto(first_url, wait_until="domcontentloaded", timeout=25000)
            await asyncio.sleep(2)

            cur = page.url
            if "/identity/login" in cur or "/auth/familysearch/login" in cur:
                _prog(26, "Signing in...")
                if not email or not password:
                    summary.update({"error":   "no_credentials",
                                    "message": "Login required but no credentials provided."})
                    return summary
                if not await _do_login(page, email, password, log):
                    summary.update({"error":   "login_failed",
                                    "message": "Sign-in failed — check credentials."})
                    return summary
                await asyncio.sleep(2)
                # After login FamilySearch may redirect to home; go back to first record
                if first_url not in page.url:
                    _prog(28, "Navigating to first record after login...")
                    await page.goto(first_url, wait_until="domcontentloaded", timeout=25000)
                    await asyncio.sleep(2)
            else:
                log(f"  Already signed in. URL: {cur}")

            if _done(): return summary

            # 6. Scrape first record
            _prog(30, f"[1] Scraping: {first_name[:60]}...")
            records: list = []
            det = await _scrape_record(ctx, first_url, first_name, images_root, log)
            det["score"]         = qualified[0]["score"]
            det["collection"]    = qualified[0].get("coll", "")
            det["events"]        = qualified[0].get("evts", "")
            det["relationships"] = qualified[0].get("rels", "")
            records.append(det)
            log(f"    OK  {det['title'][:70]}  {det['score']}%")
            await asyncio.sleep(0.8)

            # 7. Return to results page
            _prog(35, "Returning to results page...")
            await page.go_back()
            await asyncio.sleep(2)

            if has_adv:
                # ── Advanced search branch ── #
                # MANDATORY: refresh page before opening Advanced Search
                _prog(38, "Refreshing page (required before Advanced Search)...")
                await page.reload(wait_until="domcontentloaded", timeout=20000)
                await asyncio.sleep(3)

                # Re-click Historical Records tab after refresh
                await _click_historical_tab(page, log)
                await asyncio.sleep(2)

                # Open Advanced Search and fill fields
                _prog(40, "Opening Advanced Search...")
                await _advanced_search(page, adv, log)
                await asyncio.sleep(2)

                _prog(44, "Setting 60 results after advanced search...")
                await _set_60(page, log)

                _prog(47, "Re-collecting results...")
                raw_adv   = await _collect(page, qname, log)
                qualified = [r for r in raw_adv if r["score"] >= MIN_MATCH]
                log(f"  Qualified after advanced search: {len(qualified)}")

                # Scrape ALL results from scratch (first record included again)
                records = []
                remaining = qualified
            else:
                # ── Simple branch: continue from result #2 ── #
                remaining = qualified[1:]
                log(f"  Remaining results to scrape: {len(remaining)}")

            if not remaining and not records:
                _prog(100, "No qualifying results.")
                summary.update({"ok": True, "n_records": 0,
                                "message": "No qualifying results."})
                return summary

            n_total = len(remaining) + (0 if has_adv else 1)
            for i, r in enumerate(remaining, 2 if not has_adv else 1):
                if _done(): break
                pct = 47 + int(38 * i / max(n_total, 1))
                _prog(pct, f"[{i}/{n_total}] {r['name'][:60]}...")
                det = await _scrape_record(ctx, r["url"], r["name"],
                                           images_root, log)
                det["score"]         = r["score"]
                det["collection"]    = r.get("coll", "")
                det["events"]        = r.get("evts", "")
                det["relationships"] = r.get("rels", "")
                records.append(det)
                log(f"    OK  {det['title'][:70]}  {det['score']}%")
                await asyncio.sleep(0.8)

            # 8. Save output files named by person's name
            _prog(88, "Saving output files...")
            docx_p = output_folder / f"{file_base}.docx"
            xlsx_p = output_folder / f"{file_base}.xlsx"
            sd = sx = False
            if want_docx and records:
                write_docx(docx_p, records, qlines)
                sd = True
                log(f"  Word saved: {docx_p.name}")
            if want_xlsx and records:
                write_xlsx(xlsx_p, records, qlines)
                sx = True
                log(f"  Excel saved: {xlsx_p.name}")

            _prog(100, f"Done — {len(records)} record(s).")
            summary.update({
                "ok":            True,
                "docx_count":    1 if sd else 0,
                "xlsx_path":     str(xlsx_p) if sx else None,
                "output_folder": str(output_folder),
                "n_records":     len(records),
            })

        except Exception as exc:
            summary.update({"error":   "exception",
                            "message": f"{type(exc).__name__}: {exc}"})
            log(f"  !! {exc}")
        finally:
            try: await ctx.close()
            except Exception: pass
            try: await browser.close()
            except Exception: pass

    return summary
