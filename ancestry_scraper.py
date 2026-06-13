#!/usr/bin/env python3
"""
ancestry_scraper.py
===================
Ancestry.com scraper — built on the SAME pattern as familysearch_scraper.py
(persistent login profile, sign in ONCE before opening records, per-field exact
match, one scrape pass, Word + Excel output with the site name, overwrite/append/
skip dialog, never crash on a broken page).

Ancestry is a paid, login-walled, anti-bot site (like MyHeritage / FamilySearch):
- PERSISTENT profile `.ancestry_profile` keeps the login between runs — sign in
  once, not every run. NO hardcoded user_agent (stale UA = bot tell).
- Sign in ONCE up front, before opening any record.

Flow:
  1. Home page form → #firstName, #lastName, #place, #birthYear → Search.
  2. The results URL (/search/?name=First+Middle_Last&birth=YEAR…) is then
     augmented for the requested EXACT flags / spouse-parent filters and reloaded
     — this narrows the result set BEFORE scraping (like FS advanced search).
  3. Collect result rows (record links), keep name-matches ≥ MIN_MATCH.
  4. Open each record on ONE page, copy every field, save the document image
     (best-effort via the image viewer).

NOTE (live tuning): the home-form fields, the login fields, the results/edit/
more-options buttons and the search-URL scheme are confirmed; the result-row,
record-detail and image-viewer selectors are defensive/generic and dump a
diagnostic on 0 results — refine them from the first live run's log.
"""

import asyncio, difflib, io, json, os, re, shutil, sys, time
from pathlib import Path
from urllib.parse import urlparse, parse_qsl, urlencode, quote
from docx_util import set_cell_lines

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
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# ── Config ────────────────────────────────────────────────────────────────── #
_HERE = Path(__file__).resolve().parent
try:
    _CFG = json.loads((_HERE / "config" / "ancestry.json").read_text("utf-8"))
except Exception:
    _CFG = {}

HOME_URL   = _CFG.get("home_url", "https://www.ancestry.com/")
SIGNIN_URL = _CFG.get("signin_url", "https://www.ancestry.com/account/signin")
SEARCH_URL = "https://www.ancestry.com/search/"
ANC_BASE   = "https://www.ancestry.com"
SITE_NAME  = "Ancestry"
MIN_MATCH  = int(_CFG.get("min_match", 80))
MAX_PAGES  = int(_CFG.get("max_pages", 5))     # result pages to walk
MAX_SCRAPE = int(_CFG.get("max_scrape", 40))   # cap records actually opened
ANC_PROFILE_DIR = _HERE / ".ancestry_profile"  # persistent login/cookies
_dl           = _CFG.get("downloads_dir", "")
DOWNLOADS_DIR = Path(_dl) if _dl else Path.home() / "Downloads"
HYPERLINK_REL = ("http://schemas.openxmlformats.org/"
                 "officeDocument/2006/relationships/hyperlink")
_IMG_SKIP = ("icon", "logo", "sprite", "avatar", "pixel", "placeholder", ".svg",
             "static.ancestry", "/css/", "gravatar")
# record links worth opening
_REC_PATTERNS = ("/discoveryui-content/view/", "/imageviewer/",
                 "/family-tree/person/", "recordpid", "/cgi-bin/sse.dll")


# ── Utils ─────────────────────────────────────────────────────────────────── #

def _sim(a: str, b: str) -> float:
    a = re.sub(r"\s+", " ", (a or "").strip().lower())
    b = re.sub(r"\s+", " ", (b or "").strip().lower())
    if not a or not b:
        return 0.0
    wa, wb = set(a.split()), set(b.split())
    return max(len(wa & wb) / max(len(wa), 1),
               difflib.SequenceMatcher(None, a, b).ratio()) * 100


def safe_fn(s: str, n: int = 100) -> str:
    return re.sub(r"\s+", " ",
                  re.sub(r'[\\/*?:"<>|]', "_", (s or "").strip()))[:n].strip() \
        or "document"


def _abs(href: str) -> str:
    return ANC_BASE + href if href.startswith("/") else href


def _is_record(href: str) -> bool:
    h = (href or "").lower()
    return any(p in h for p in _REC_PATTERNS)


def _name_part(s: str) -> str:
    """«Alexander W» → «Alexander+W» (spaces → +, as in Ancestry's name= param)."""
    return "+".join((s or "").split())


# ── Search-URL building ────────────────────────────────────────────────────── #

def _build_search_url(fn, ln, place, year, exact, adv) -> str:
    """Ancestry /search/ URL. name=First+Middle_Last, birth=YEAR, residence=…,
    spouse/father/mother=…, with `<field>_x=1` for the EXACT flags. Spaces are a
    literal «+», so the query is assembled by hand (urlencode would escape it)."""
    exact = exact or {}
    adv   = adv or {}
    parts = []

    name = ""
    if fn or ln:
        name = f"{_name_part(fn)}_{_name_part(ln)}".strip("_")
    if name:
        parts.append(("name", name))
        if exact.get("name") or exact.get("surname"):
            parts.append(("name_x", "1"))   # Ancestry name is one unit (given+surname)
    if year:
        y = re.sub(r"\D", "", str(year))[:4]
        if y:
            parts.append(("birth", y))
            if exact.get("year"):
                parts.append(("birth_x", "1"))
    if place:
        parts.append(("residence", _name_part(place)))
        if exact.get("place"):
            parts.append(("residence_x", "1"))
    for rel in ("spouse", "father", "mother"):
        rf, rl = adv.get(f"{rel}_first", ""), adv.get(f"{rel}_last", "")
        if rf or rl:
            parts.append((rel, f"{_name_part(rf)}_{_name_part(rl)}".strip("_")))
            if adv.get(f"{rel}_exact"):
                parts.append((f"{rel}_x", "1"))
    if adv.get("keyword"):
        parts.append(("keyword", _name_part(adv["keyword"])))
    # query string with literal '+' for spaces, '_' kept intact
    q = "&".join(f"{k}={quote(v, safe='+_')}" for k, v in parts)
    return f"{SEARCH_URL}?{q}" if q else SEARCH_URL


def _apply_exact_to_results(url: str, exact: dict, adv: dict) -> str:
    """Add the `_x=1` flags / spouse-parent params to whatever Ancestry already
    put in the results URL after the form search, then return the new URL."""
    exact = exact or {}; adv = adv or {}
    pr    = urlparse(url)
    pairs = parse_qsl(pr.query, keep_blank_values=True)
    keys  = {k for k, _ in pairs}

    def _add(k, v):
        if k not in keys:
            pairs.append((k, v)); keys.add(k)

    if ("name" in keys) and (exact.get("name") or exact.get("surname")):
        _add("name_x", "1")
    if ("birth" in keys) and exact.get("year"):
        _add("birth_x", "1")
    if exact.get("place") and ("residence" in keys or "anyplace" in keys):
        _add("residence_x", "1")
    for rel in ("spouse", "father", "mother"):
        rf, rl = adv.get(f"{rel}_first", ""), adv.get(f"{rel}_last", "")
        if rf or rl:
            _add(rel, f"{_name_part(rf)}_{_name_part(rl)}".strip("_"))
            if adv.get(f"{rel}_exact"):
                _add(f"{rel}_x", "1")
    return pr._replace(query=urlencode(pairs, safe="+_")).geturl()


# ── Field typing ──────────────────────────────────────────────────────────── #

async def _type_field(page, sel: str, val: str, label: str, log) -> bool:
    if not val:
        return True
    try:
        el = page.locator(sel).first
        if not await el.count():
            log(f"  !! поле {label} не найдено ({sel})")
            return False
        await el.scroll_into_view_if_needed(timeout=4000)
        await el.click(timeout=3000)
        await asyncio.sleep(0.1)
        await page.keyboard.press("Control+a")
        await page.keyboard.press("Delete")
        await page.keyboard.type(val, delay=35)
        await asyncio.sleep(0.2)
        log(f"  OK  {label} = {val!r}")
        return True
    except Exception as e:
        log(f"  !! {label}: {e}")
        return False


# ── 1. SEARCH (home form) ─────────────────────────────────────────────────── #

async def _search(page, fn, ln, place, year, exact, adv, log):
    log(f"  Открываю {HOME_URL}")
    await page.goto(HOME_URL, wait_until="domcontentloaded", timeout=40000)
    await asyncio.sleep(3)

    # wait for the home search form
    for sel in ["#firstName", 'input[name="firstName"]']:
        try:
            await page.locator(sel).first.wait_for(state="visible", timeout=10000)
            break
        except Exception:
            continue

    filled = await _type_field(page, '#firstName, input[name="firstName"]',
                               fn, "First Name", log)
    filled |= await _type_field(page, '#lastName, input[name="lastName"]',
                                ln, "Last Name", log)
    if place:
        await _type_field(page, '#place, input[name="place"]',
                          place, "Place", log)
    if year:
        await _type_field(page, '#birthYear, input[name="birthYear"]',
                          year, "Birth Year", log)

    clicked = False
    for sel in ['button.ancBtn[type="submit"]', 'button[type="submit"]',
                'button:has-text("Search")']:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=6000)
                log(f"  Search нажат ({sel})")
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        # last resort: build the search URL directly
        log("  !! кнопка Search не найдена — иду по URL")
        await page.goto(_build_search_url(fn, ln, place, year, exact, adv),
                        wait_until="domcontentloaded", timeout=40000)

    try:
        await page.wait_for_url(lambda u: "/search" in u, timeout=20000)
    except Exception:
        pass
    await asyncio.sleep(5)

    # narrow the result set BEFORE scraping: exact flags + spouse/parents
    if "/search" in page.url:
        u2 = _apply_exact_to_results(page.url, exact, adv)
        if u2 != page.url:
            log(f"  → уточняю выдачу (exact/родня): {[k for k,v in (exact or {}).items() if v]}")
            try:
                await page.goto(u2, wait_until="domcontentloaded", timeout=40000)
                await asyncio.sleep(5)
            except Exception as e:
                log(f"  !! не вышло уточнить URL: {type(e).__name__}")
    log(f"  Результаты: {page.url}")


# ── 2. LOGIN ──────────────────────────────────────────────────────────────── #

async def _login(page, email: str, password: str, log) -> bool:
    """Fill #username + #password, submit. Call ONCE per session."""
    log(f"  Страница логина: {page.url[:80]}")
    try:
        await page.locator('#username, input[name="username"]').first.wait_for(
            state="visible", timeout=20000)
    except Exception:
        log("  !! поле username не появилось")
        return False
    await asyncio.sleep(2)

    async def _fill(sel, val, name):
        for attempt in range(3):
            try:
                el = page.locator(sel).first
                await el.fill(val)
                await asyncio.sleep(0.5)
                if (await el.input_value()).strip():
                    log(f"  {name} заполнен")
                    return True
            except Exception as e:
                log(f"  !! fill {name}: {e}")
            await asyncio.sleep(1)
        return False

    if not await _fill('#username, input[name="username"]', email, "username"):
        return False
    if not await _fill('#password, input[name="password"]', password, "password"):
        return False

    for sel in ['#signInBtn', 'button[type="submit"]',
                'button:has-text("Sign in")', 'button:has-text("Sign In")']:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=5000)
                log(f"  Sign in нажат ({sel})")
                break
        except Exception:
            continue

    try:
        await page.wait_for_url(
            lambda u: "ancestry.com" in u and "signin" not in u and "login" not in u,
            timeout=30000)
    except Exception:
        await asyncio.sleep(6)
    ok = "signin" not in page.url and "login" not in page.url
    log(f"  Логин {'OK ✓' if ok else 'ПРОВАЛИЛСЯ ✗'}  URL: {page.url[:80]}")
    return ok


async def _sign_in_if_needed(page, email, password, logged_in_ref, log) -> bool:
    """Sign in ONCE up front (or skip if the persistent profile already has the
    session). Sets logged_in_ref[0]=True only on a real login."""
    log("  Проверяю авторизацию...")
    # already logged in? a logged-in Ancestry nav has an account avatar, no LOG IN
    try:
        has_login = await page.evaluate(
            """() => {
                const t = document.body.innerText || '';
                const link = [...document.querySelectorAll('a,button')].some(
                    e => /log\\s*in|sign\\s*in/i.test((e.textContent||'').trim())
                         && (e.textContent||'').trim().length < 12);
                return link;
            }""")
    except Exception:
        has_login = True
    if not has_login:
        log("  Уже авторизованы — вход не нужен")
        return True
    if not email or not password:
        log("  !! Нет логина/пароля — пробую без входа (только дерево/превью)")
        return False

    log("  Не авторизованы — вхожу ОДИН раз до открытия записей...")
    results_url = page.url
    try:
        await page.goto(SIGNIN_URL, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2)
        ok = await _login(page, email, password, log)
        if ok:
            logged_in_ref[0] = True
            if "/search" not in page.url and results_url:
                await page.goto(results_url, wait_until="domcontentloaded",
                                timeout=30000)
                await asyncio.sleep(3)
        return ok
    except Exception as e:
        log(f"  !! Sign in: {e}")
        return False


# ── 3. COLLECT RESULTS ────────────────────────────────────────────────────── #

_COLLECT_JS = r"""() => {
    const seen = new Set(), out = [];
    const pats = ["/discoveryui-content/view/", "/imageviewer/",
                  "/family-tree/person/", "recordpid", "/cgi-bin/sse.dll"];
    const isRec = h => h && pats.some(p => h.toLowerCase().includes(p));
    for (const a of document.querySelectorAll('a[href]')) {
        const href = a.href || '';
        if (!isRec(href)) continue;
        if (seen.has(href)) continue;
        // climb to the row container for the full text
        let row = a, hops = 0;
        while (row && hops < 6 &&
               (row.innerText || '').trim().length < 30) { row = row.parentElement; hops++; }
        const text = ((row && row.innerText) || a.innerText || '').trim();
        const name = (a.innerText || '').trim() || text.split('\n')[0] || '';
        if (!name) continue;
        seen.add(href);
        out.push({url: href, name: name, text: text.slice(0, 1500)});
    }
    return out;
}"""


async def _collect(page, qname: str, log) -> list:
    await asyncio.sleep(2)
    try:
        raw = await page.evaluate(_COLLECT_JS)
    except Exception as e:
        log(f"  !! сбор результатов: {type(e).__name__}")
        raw = []
    out = []
    for r in raw:
        if not _is_record(r["url"]):
            continue
        r["score"] = round(_sim(qname, r["name"]), 1)
        out.append(r)
        log(f"    {r['score']:5.1f}%  {r['name'][:60]}")
    log(f"  Кандидатов на странице: {len(out)}")
    return out


async def _collect_all(page, qname, base_url, log) -> list:
    """Walk up to MAX_PAGES result pages (Ancestry paginates with &page=N)."""
    all_rows, seen = [], set()
    for pg in range(1, MAX_PAGES + 1):
        if pg > 1:
            sep = "&" if "?" in base_url else "?"
            nu = re.sub(r"([?&])page=\d+", r"\1", base_url).rstrip("?&")
            nu = f"{nu}{sep}page={pg}"
            try:
                await page.goto(nu, wait_until="domcontentloaded", timeout=40000)
                await asyncio.sleep(4)
            except Exception:
                break
        rows = await _collect(page, qname, log)
        new = [r for r in rows if r["url"] not in seen]
        for r in new:
            seen.add(r["url"])
        all_rows.extend(new)
        if not new:
            break
        if len([r for r in all_rows if r["score"] >= MIN_MATCH]) >= MAX_SCRAPE:
            break
    return all_rows


async def _diag(page, log):
    """Dump page state when 0 results — so the real DOM can be inspected from the
    log (live site is anti-bot / login-walled for me)."""
    try:
        d = await page.evaluate(r"""() => ({
            url: location.href,
            title: document.title,
            anchors: document.querySelectorAll('a[href]').length,
            recAnchors: [...document.querySelectorAll('a[href]')]
                .filter(a => /discoveryui-content|imageviewer|family-tree\/person/i
                    .test(a.href)).length,
            hasLogin: /log\s*in|sign\s*in/i.test(document.body.innerText||''),
            sample: [...document.querySelectorAll('a[href]')]
                .map(a=>a.href).filter(h=>/ancestry\.com/.test(h)).slice(0,12),
        })""")
        log(f"  ДИАГНОСТИКА: {json.dumps(d, ensure_ascii=False)[:1200]}")
    except Exception as e:
        log(f"  ДИАГНОСТИКА не удалась: {e}")


# ── 4. ADVANCED (More options) — best-effort open, mostly URL-driven ───────── #

async def _open_more_options(page, log):
    for sel in ['[data-testid="moreOptionsBtnTestId"]',
                '[data-testid="btnEditForm"]']:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=4000)
                await asyncio.sleep(1.5)
                log(f"  Расширенный поиск открыт ({sel})")
                return True
        except Exception:
            continue
    return False


# ── 5. IMAGES ─────────────────────────────────────────────────────────────── #

async def _best_img(page) -> str:
    """Largest document image; scroll first (lazy) and keep a viewer fallback."""
    try:
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await asyncio.sleep(0.6)
        await page.evaluate("window.scrollTo(0, 0)")
        await asyncio.sleep(0.3)
    except Exception:
        pass
    best, area, fallback = "", 0, ""
    for el in await page.query_selector_all("img"):
        try:
            src = (await el.evaluate("e => e.currentSrc || e.src || ''") or "").strip()
            if not src.startswith("http"):
                continue
            if any(b.lower() in src.lower() for b in _IMG_SKIP):
                continue
            w = int(await el.evaluate("e => e.naturalWidth")  or 0)
            h = int(await el.evaluate("e => e.naturalHeight") or 0)
            if w * h > area:
                area, best = w * h, src
            if not fallback and any(k in src.lower() for k in
                                    ("imageviewer", "/media", "dms", "mediasvc",
                                     "interactive")):
                fallback = src
        except Exception:
            continue
    return best or fallback


async def _fetch_bytes(ctx, src: str) -> bytes | None:
    if not src or not src.startswith("http"):
        return None
    pg = await ctx.new_page()
    try:
        r = await pg.goto(src, timeout=20000)
        if r and r.ok:
            body = await r.body()
            if len(body) > 4000:
                return body
    except Exception:
        pass
    finally:
        try: await pg.close()
        except Exception: pass
    return None


async def _download_image(ctx, page, dest_dir: Path, title: str, log) -> str | None:
    """Best-effort: the Ancestry viewer's own download (wrapped in expect_download,
    like FS), else save the largest image on the page."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / (safe_fn(title) + ".jpg")
    try:
        async with page.expect_download(timeout=25000) as dl_info:
            done = False
            for sel in ['button[aria-label*="Download" i]',
                        '[data-testid*="download" i]',
                        'button[title*="Download" i]',
                        'a[aria-label*="Download" i]']:
                try:
                    el = page.locator(sel).first
                    if await el.count() and await el.is_visible():
                        await el.click(timeout=4000)
                        await asyncio.sleep(1)
                        done = True
                        # a confirm step may follow (JPG / Save)
                        for c in ['button:has-text("Download")',
                                  'button:has-text("Save")',
                                  '[data-testid*="confirm" i]']:
                            try:
                                cb = page.locator(c).first
                                if await cb.count() and await cb.is_visible():
                                    await cb.click(timeout=3000)
                                    break
                            except Exception:
                                continue
                        break
                except Exception:
                    continue
            if not done:
                raise RuntimeError("кнопка Download не найдена")
        dl = await dl_info.value
        await dl.save_as(str(dest))
        log(f"    🖼 документ сохранён: {dest.name} ({dest.stat().st_size//1024}KB)")
        return str(dest)
    except Exception as exc:
        log(f"    (download через вьюер не вышло: {type(exc).__name__}) — беру картинку")

    src = await _best_img(page)
    if src:
        body = await _fetch_bytes(ctx, src)
        if body:
            dest.write_bytes(body)
            log(f"    🖼 изображение сохранено: {dest.name} ({len(body)//1024}KB)")
            return str(dest)
    log("    (изображение документа не найдено)")
    return None


# ── 6. SCRAPE ONE RECORD ──────────────────────────────────────────────────── #

_FIELDS_JS = r"""() => {
    const out = [];
    const push = (k, v) => {
        k = (k||'').replace(/\s+/g,' ').trim().replace(/:$/,'');
        v = (v||'').replace(/\s+/g,' ').trim();
        if (k && v && k.length < 60 && v.length < 600) out.push([k, v]);
    };
    // dl/dt/dd
    document.querySelectorAll('dl').forEach(dl => {
        const dts = dl.querySelectorAll('dt'), dds = dl.querySelectorAll('dd');
        for (let i=0; i<Math.min(dts.length, dds.length); i++)
            push(dts[i].innerText, dds[i].innerText);
    });
    // tables: label | value
    document.querySelectorAll('table tr').forEach(tr => {
        const c = tr.querySelectorAll('td,th');
        if (c.length >= 2) push(c[0].innerText, c[1].innerText);
    });
    // Ancestry fact rows (two-cell flex rows with a label + value)
    document.querySelectorAll('[class*="tableRow"],[class*="recordField"],[class*="fact"]')
        .forEach(r => {
            const kids = r.children;
            if (kids && kids.length >= 2)
                push(kids[0].innerText, kids[1].innerText);
        });
    return out;
}"""


async def _scrape_page(ctx, page, url, name_hint, images_root,
                       logged_in_ref, email, password, log) -> dict:
    rec = {"url": url, "title": name_hint, "name": name_hint,
           "table_data": {}, "images": [], "thumb_bytes": None,
           "collection": "", "score": 0}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=40000)
        await asyncio.sleep(3)
    except Exception as e:
        log(f"    !! не открылась запись ({type(e).__name__})")
        return rec

    # login redirect fallback (early sign-in should already cover it)
    if ("signin" in page.url or "login" in page.url):
        if not logged_in_ref[0] and email and password:
            log("  → запись требует входа, вхожу...")
            if await _login(page, email, password, log):
                logged_in_ref[0] = True
                await page.goto(url, wait_until="domcontentloaded", timeout=40000)
                await asyncio.sleep(3)
        else:
            log("    !! запись за логином — пропускаю")
            return rec

    # title
    for sel in ["h1", '[class*="title" i]', "h2"]:
        try:
            t = (await page.locator(sel).first.inner_text(timeout=2500) or "").strip()
            if 2 < len(t) < 200:
                rec["title"] = t
                break
        except Exception:
            pass

    # fields
    try:
        pairs = await page.evaluate(_FIELDS_JS)
    except Exception:
        pairs = []
    td = {}
    for k, v in pairs:
        if k not in td:
            td[k] = v
    rec["table_data"] = td

    label = rec["title"] or name_hint
    img_dir = images_root / safe_fn(label)

    src = await _best_img(page)
    if src:
        rec["thumb_bytes"] = await _fetch_bytes(ctx, src)

    jp = await _download_image(ctx, page, img_dir, label, log)
    if jp:
        rec["images"] = [jp]
    elif rec["thumb_bytes"]:
        img_dir.mkdir(parents=True, exist_ok=True)
        fp = img_dir / (safe_fn(label) + "_preview.jpg")
        fp.write_bytes(rec["thumb_bytes"])
        rec["images"] = [str(fp)]
    return rec


# ── Word ──────────────────────────────────────────────────────────────────── #

def _add_link(para, text, url):
    rid = para.part.relate_to(url, HYPERLINK_REL, is_external=True)
    hl  = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), rid)
    run = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    c   = OxmlElement("w:color"); c.set(qn("w:val"), "0563C1")
    u   = OxmlElement("w:u");     u.set(qn("w:val"), "single")
    rPr.append(c); rPr.append(u); run.append(rPr)
    t   = OxmlElement("w:t"); t.text = text or url
    t.set(qn("xml:space"), "preserve")
    run.append(t); hl.append(run); para._p.append(hl)


def _docx_add_record(doc, i, rec):
    title = rec.get("title") or rec.get("name", "—")
    doc.add_heading(f"{i}. {title}", level=2)
    if rec.get("url"):
        pp = doc.add_paragraph()
        pp.add_run(f"Источник ({SITE_NAME}): ").bold = True
        _add_link(pp, "Открыть запись", rec["url"])
    if rec.get("score"):
        p = doc.add_paragraph()
        p.add_run("Совпадение: ").bold = True
        p.add_run(f"{rec.get('score','?')}%")

    rows = []
    if rec.get("collection"):
        rows.append(("Коллекция", rec["collection"]))
    for f, v in rec.get("table_data", {}).items():
        rows.append((str(f), str(v)))
    if rows:
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        hdr[0].text = "Поле"; hdr[1].text = "Значение"
        for cell in hdr:
            for run in cell.paragraphs[0].runs:
                run.bold = True
        for f, v in rows:
            r = tbl.add_row().cells
            r[0].text = f; set_cell_lines(r[1], v)
    doc.add_paragraph("")

    imgs = rec.get("images", []); tb = rec.get("thumb_bytes")
    if imgs and Path(imgs[0]).exists():
        doc.add_paragraph("Изображение документа:").runs[0].bold = True
        try:
            doc.add_picture(imgs[0], width=Inches(4))
        except Exception:
            doc.add_paragraph(f"  [{Path(imgs[0]).name}]")
    elif tb:
        doc.add_paragraph("Превью документа:").runs[0].bold = True
        try:
            doc.add_picture(io.BytesIO(tb), width=Inches(4))
        except Exception:
            doc.add_paragraph("  [не удалось вставить]")
    doc.add_paragraph("")


def write_docx(path: Path, records: list, qlines: list, append: bool = False):
    if not _DOCX_OK:
        raise RuntimeError("python-docx не установлен")
    existing = append and Path(path).exists()
    if existing:
        doc = Document(str(path))
        doc.add_page_break()
        sep = doc.add_heading(
            f"➕ Добавлено ещё {len(records)} (совпадение ≥{MIN_MATCH}%) — "
            f"{time.strftime('%Y-%m-%d %H:%M')}", level=1)
        sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Параметры:")
        for ln in qlines:
            doc.add_paragraph(ln, style="List Bullet")
        doc.add_paragraph("")
    else:
        doc = Document()
        s = doc.sections[0]
        s.page_width  = Mm(297); s.page_height = Mm(210)
        s.left_margin = s.right_margin = Mm(15)
        s.top_margin  = s.bottom_margin = Mm(15)
        h = doc.add_heading("Ancestry — Результаты", 0)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Параметры:")
        for ln in qlines:
            doc.add_paragraph(ln, style="List Bullet")
        doc.add_paragraph(f"Найдено: {len(records)} (совпадение ≥{MIN_MATCH}%)")
        doc.add_paragraph("")
    for i, rec in enumerate(records, 1):
        _docx_add_record(doc, i, rec)
    doc.save(path)


# ── Excel ─────────────────────────────────────────────────────────────────── #

def write_xlsx(path: Path, records: list, qlines: list, append: bool = False):
    if not _OPENPYXL_OK:
        raise RuntimeError("openpyxl не установлен")
    HF = PatternFill("solid", fgColor="6B8E23")
    HN = Font(bold=True, color="FFFFFF", size=11)
    LINKF = Font(color="0563C1", underline="single")
    TS = Side(style="thin", color="C0D0A0")
    T  = Border(left=TS, right=TS, top=TS, bottom=TS)

    aff = []
    for rec in records:
        for k in rec.get("table_data", {}):
            if k not in aff:
                aff.append(k)
    base_cols = ["#", "База", "Имя", "Совп. %", "Коллекция", "URL"]

    existing = append and Path(path).exists()
    if existing:
        wb = load_workbook(str(path)); ws = wb.active
        header = [ws.cell(row=1, column=c).value
                  for c in range(1, ws.max_column + 1)]
        for name in base_cols + aff:
            if name not in header:
                header.append(name)
                c = ws.cell(row=1, column=len(header), value=name)
                c.font = HN; c.fill = HF; c.border = T
                c.alignment = Alignment(horizontal="center",
                                        vertical="center", wrap_text=True)
        start_row = ws.max_row + 1
        start_num = ws.max_row - 1
    else:
        wb = Workbook(); ws = wb.active; ws.title = "Ancestry"
        header = base_cols + aff
        for ci, cn in enumerate(header, 1):
            c = ws.cell(row=1, column=ci, value=cn)
            c.font = HN; c.fill = HF; c.border = T
            c.alignment = Alignment(horizontal="center",
                                    vertical="center", wrap_text=True)
        start_row = 2
        start_num = 0
    col_idx = {name: i + 1 for i, name in enumerate(header)}

    for n, rec in enumerate(records):
        ri = start_row + n
        td = rec.get("table_data", {})
        row = {"#": start_num + n + 1, "База": SITE_NAME,
               "Имя": rec.get("title", rec.get("name", "")),
               "Совп. %": rec.get("score", ""),
               "Коллекция": rec.get("collection", ""),
               "URL": rec.get("url", "")}
        for f in aff:
            row[f] = td.get(f, "")
        for name, val in row.items():
            ci = col_idx.get(name)
            if not ci:
                continue
            c = ws.cell(row=ri, column=ci)
            if name == "URL" and val:
                c.value = "Открыть"; c.hyperlink = val; c.font = LINKF
            else:
                c.value = val
            c.border = T
            c.alignment = Alignment(wrap_text=True, vertical="top")

    for ci in range(1, len(header) + 1):
        ltr = get_column_letter(ci)
        mw  = max(len(str(header[ci-1] or "")),
                  *(len(str(ws.cell(row=r, column=ci).value or "").split("\n")[0])
                    for r in range(2, ws.max_row + 1)), 8)
        ws.column_dimensions[ltr].width = min(mw + 4, 60)
    wb.save(path)


# ── Main entry point ──────────────────────────────────────────────────────── #

async def run_scraper(
    *,
    first_names:   str       = "",
    last_names:    str       = "",
    place_lived:   str       = "",
    birth_year:    str       = "",
    advanced:      dict|None = None,
    exact:         dict|None = None,
    output_format: str       = "both",
    output_folder            = Path("."),
    email:         str|None  = None,
    password:      str|None  = None,
    log                      = print,
    progress                 = None,
    cancel_event             = None,
    ask_file_conflict        = None,
) -> dict:

    def _prog(pct, txt):
        log(txt)
        if progress: progress(pct, txt)

    def _done():
        return bool(cancel_event and cancel_event.is_set())

    adv   = advanced or {}
    exact = exact or {}
    want_docx = output_format in ("docx", "both")
    want_xlsx = output_format in ("xlsx", "both")
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    images_root = output_folder / "images"

    qname  = " ".join(p for p in (first_names, last_names) if p)
    qlines = [ln for ln in [
        f"First Names: {first_names}", f"Last Names: {last_names}",
        f"Place: {place_lived}", f"Birth Year: {birth_year}",
    ] if not ln.endswith(": ")]
    for rel in ("spouse", "father", "mother"):
        nm = " ".join(p for p in (adv.get(f"{rel}_first", ""),
                                  adv.get(f"{rel}_last", "")) if p)
        if nm:
            qlines.append(f"{rel.capitalize()}: {nm}")
    summary   = {"ok": False}
    file_base = safe_fn(f"ancestry_{qname}") if qname else "ancestry_results"
    logged_in_ref = [False]

    _prog(0, "Запускаю браузер...")
    async with async_playwright() as pw:
        # PERSISTENT profile → the Ancestry login is kept between runs.
        for _lk in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            try: (ANC_PROFILE_DIR / _lk).unlink()
            except Exception: pass
        ANC_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        ctx = await pw.chromium.launch_persistent_context(
            str(ANC_PROFILE_DIR),
            headless=False,
            no_viewport=True,
            accept_downloads=True,
            args=["--start-maximized",
                  "--disable-blink-features=AutomationControlled"],
        )
        browser = ctx
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for _p in list(ctx.pages)[1:]:
            try: await _p.close()
            except Exception: pass

        try:
            # 1. SEARCH
            _prog(5, "Поиск...")
            await _search(page, first_names, last_names, place_lived,
                          birth_year, exact, adv, log)
            if _done(): return summary

            # 2. SIGN IN ONCE (before opening records)
            _prog(15, "Sign in...")
            await _sign_in_if_needed(page, email or "", password or "",
                                     logged_in_ref, log)
            if _done(): return summary

            # 3. COLLECT (logged-in results)
            _prog(25, "Сбор результатов...")
            base_url = page.url
            raw = await _collect_all(page, qname, base_url, log)
            qualified = [r for r in raw if r["score"] >= MIN_MATCH][:MAX_SCRAPE]
            log(f"  Подходящих (≥{MIN_MATCH}%): {len(qualified)}")

            if not qualified:
                await _diag(page, log)
                _prog(100, f"Нет записей ≥{MIN_MATCH}%.")
                summary.update({"ok": True, "n_records": 0,
                                "message": f"Нет записей ≥{MIN_MATCH}%."})
                return summary

            # 4. SCRAPE each record on ONE page
            records = []
            for i, r in enumerate(qualified, 1):
                if _done(): break
                _prog(25 + int(65 * i / len(qualified)),
                      f"[{i}/{len(qualified)}] {r['name'][:60]}...")
                det = await _scrape_page(ctx, page, r["url"], r["name"],
                                         images_root, logged_in_ref,
                                         email or "", password or "", log)
                det["score"] = r["score"]
                records.append(det)
                log(f"  ✓  {det.get('title','')[:70]}  ({r['score']}%)")

            # 5. SAVE (overwrite / append / skip)
            _prog(94, "Сохранение файлов...")
            docx_p = output_folder / f"{file_base}.docx"
            xlsx_p = output_folder / f"{file_base}.xlsx"
            existing_names = [p.name for p, want in
                              ((docx_p, want_docx), (xlsx_p, want_xlsx))
                              if want and records and p.exists()]
            decision = "overwrite"
            if existing_names and ask_file_conflict:
                try:
                    decision = (ask_file_conflict(existing_names)
                                or "overwrite").lower()
                except Exception as _e:
                    log(f"  !! диалог конфликта файлов: {_e}")
                log(f"  → Файл(ы) уже существуют {existing_names} → выбор: {decision}")
            append = (decision == "append")

            sd = sx = False
            if decision == "skip":
                log("  → Сохранение пропущено по выбору пользователя.")
            else:
                if want_docx and records:
                    write_docx(docx_p, records, qlines, append=append)
                    sd = True
                    log(f"  Word: {docx_p}")
                if want_xlsx and records:
                    write_xlsx(xlsx_p, records, qlines, append=append)
                    sx = True
                    log(f"  Excel: {xlsx_p}")
            _prog(100, f"Готово — {len(records)} записей.")
            summary.update({
                "ok":            True,
                "docx_count":    1 if sd else 0,
                "xlsx_path":     str(xlsx_p) if sx else None,
                "output_folder": str(output_folder),
                "n_records":     len(records),
            })
        except Exception as exc:
            summary.update({"error": "exception",
                            "message": f"{type(exc).__name__}: {exc}"})
            log(f"  !! {exc}")
        finally:
            try: await ctx.close()
            except Exception: pass
            try: await browser.close()
            except Exception: pass
    return summary


if __name__ == "__main__":
    asyncio.run(run_scraper(
        first_names="Alexander W", last_names="Sanders", birth_year="1897",
        output_folder=Path.home() / "Downloads" / "Ancestry_results"))
