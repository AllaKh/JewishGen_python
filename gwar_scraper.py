#!/usr/bin/env python3
"""
gwar_scraper.py
===============
Scraper for «Памяти героев Великой войны» (https://gwar.mil.ru/heroes/) — WWI
(1914–1918) participant records (same ministry/structure family as pamyat).

Flow:
  1. Open /heroes/, fill the search fields (all by language-independent id/name),
     set the «Разделы» source checkboxes, submit «Найти».
  2. Collect the result documents (links to /heroes/chelovek…). When a name was
     given, keep only FUZZY-matching people (surname ~1–2 letters, first name /
     initials, patronymic stem) and, if a birth year is given, drop other years.
  3. Open each matched record, copy the left-panel fields to Word, and save the
     scanned document image(s) — every page (Страница X из N) via #btnSaveImage,
     falling back to the displayed image.

The site is Russian-only; the GUI is English. No login.

NEVER crash on a broken/empty page: every record and every step is wrapped in
try/except, urlparse().netloc (not .host), bad links are skipped.
"""

import asyncio
import difflib
import io
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlparse, urlencode, parse_qsl, urlsplit, urlunsplit

if getattr(sys, "frozen", False):
    bd = Path(sys.executable).resolve().parent / "ms-playwright"
    if bd.exists():
        import os
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(bd)

try:
    from playwright.async_api import async_playwright
except ImportError:
    sys.stderr.write("pip install playwright && playwright install chromium\n")
    sys.exit(1)

try:
    from docx import Document
    from docx.shared import Mm, Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

try:
    from PIL import Image
    _PIL_OK = True
except ImportError:
    _PIL_OK = False

HYPERLINK_REL = ("http://schemas.openxmlformats.org/"
                 "officeDocument/2006/relationships/hyperlink")

# «Разделы» (sources) — English label → checkbox id on the page.
SECTIONS = {
    "Awards":        "award_tag",
    "Losses":        "dead_tag",
    "Personal data": "frc_tag",
    "Commanders":    "commander_tag",
    "Notable people": "prs_tag",
}


# ── Helpers ───────────────────────────────────────────────────────────────── #
def safe_fn(s: str, n: int = 80) -> str:
    s = re.sub(r'[\\/*?:"<>|]', "_", (s or "").strip())
    return re.sub(r"\s+", "_", s)[:n].strip("_") or "record"


def _host(u: str) -> str:
    u = u or ""
    if not u or u.startswith(("chrome-error", "about:", "data:")):
        return ""
    try:
        return urlparse(u).netloc or ""
    except Exception:
        return ""


def _to_png(data: bytes):
    if not data:
        return None
    if (data[:3] == b"\xff\xd8\xff" or data[:8] == b"\x89PNG\r\n\x1a\n"
            or data[:4] == b"GIF8" or data[:2] == b"BM"):
        return data
    if _PIL_OK:
        try:
            im = Image.open(io.BytesIO(data)).convert("RGB")
            out = io.BytesIO(); im.save(out, format="PNG")
            return out.getvalue()
        except Exception:
            return None
    return None


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


# ── FIO fuzzy matching (same rules as pamyat) ─────────────────────────────── #
def _sim(a: str, b: str) -> float:
    a, b = (a or "").lower().strip(), (b or "").lower().strip()
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def _is_initial(token: str, full: str) -> bool:
    t = (token or "").strip(". ").lower()
    f = (full or "").lower()
    return len(t) == 1 and bool(f) and f.startswith(t)


def _parse_fio(name: str):
    name = (name or "").strip()
    name = re.sub(r"([А-ЯЁA-Za-zа-яё])\.([А-ЯЁA-Z])", r"\1. \2", name)
    parts = [p for p in re.split(r"\s+", name) if p]
    return (parts[0] if parts else "",
            parts[1] if len(parts) > 1 else "",
            parts[2] if len(parts) > 2 else "")


def _first_match(r: str, w: str) -> bool:
    r = (r or "").strip(". ").lower(); w = (w or "").strip(". ").lower()
    if not r or not w:
        return True
    if _is_initial(r, w) or _is_initial(w, r):
        return True
    if r[:2] == w[:2]:                       # Герш / Гирш
        return True
    return _sim(r, w) >= 0.5


def _patr_match(r: str, w: str) -> bool:
    r = (r or "").strip(". ").lower(); w = (w or "").strip(". ").lower()
    if not r or not w:
        return True
    short, long = (r, w) if len(r) <= len(w) else (w, r)
    if short and long.startswith(short):
        return True
    if r[:3] == w[:3]:
        return True
    return _sim(r, w) >= 0.78


def _person_matches(result_name, want_last, want_first, want_middle) -> bool:
    if not want_last:
        return True
    r_last, r_first, r_middle = _parse_fio(result_name)
    if _sim(r_last, want_last) < 0.7:
        return False
    if want_first and r_first and not _first_match(r_first, want_first):
        return False
    if want_middle and r_middle and not _patr_match(r_middle, want_middle):
        return False
    return True


# ── Detail-page field labels (gwar left info panel) ───────────────────────── #
_GWAR_LABELS = sorted([
    "Дата рождения", "Место рождения", "Должность/Звание", "Должность / Звание",
    "Должность/ Звание", "Должность",
    "Воинское звание", "Воинская часть", "Место службы", "Лагерь военнопленных",
    "Лагерь", "Дата водворения в лагерь", "Дата пленения", "Место пленения",
    "Дата начала события", "Дата окончания события", "Дата события", "Событие",
    "Место события", "Тип документа", "Картотека", "Архив", "Название фонда",
    "Фонд", "Опись", "Шкаф", "Дело", "Ящик", "Награда", "Дата документа",
    "Номер документа", "Источник информации", "Губерния", "Уезд", "Волость",
    "Населенный пункт",
], key=len, reverse=True)

_GWAR_LABEL_RE = re.compile(
    r"(" + "|".join(re.escape(_l) for _l in _GWAR_LABELS) + r")\s*:?\s*")

_CHROME_MARKERS = ("О проекте", "Вопросы и ответы", "Как искать", "Обратная связь",
                   "Правовая информация", "Пользовательское соглашение",
                   "Министерство обороны", "© Министерство", "Перейти к просмотру",
                   "Инструкция по поиску", "Видео-инструкция", "Главная страница")


def _clean_block(txt: str) -> str:
    if not txt:
        return ""
    cut = len(txt)
    for m in _CHROME_MARKERS:
        i = txt.find(m)
        if 0 <= i < cut:
            cut = i
    return txt[:cut].strip(" ;,—-·•\t\n")


def _parse_fields(txt: str):
    """Split the info-panel text by the KNOWN labels → ordered [(label, value)]."""
    txt = re.sub(r"[ \t ]+", " ", (txt or ""))
    txt = re.sub(r"\s*\n\s*", " ", txt).strip()
    marks = list(_GWAR_LABEL_RE.finditer(txt))
    out, seen = [], set()
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(txt)
        val = _clean_block(txt[m.end():end].strip(" ;,"))
        k = m.group(1).strip()
        if val and len(val) < 400 and (k, val) not in seen:
            seen.add((k, val)); out.append((k, val))
    return out


# ── Image download through the browser (CDN blocks plain requests) ────────── #
async def _img_bytes_via_goto(context, url, referer) -> bytes:
    if not url or not url.startswith("http"):
        return b""
    p = await context.new_page()
    try:
        try:
            await p.set_extra_http_headers({"referer": referer})
        except Exception:
            pass
        resp = await p.goto(url, timeout=20000, wait_until="commit")
        if resp and resp.ok:
            body = await resp.body()
            if body and len(body) > 2000:
                return body
    except Exception:
        pass
    finally:
        try:
            await p.close()
        except Exception:
            pass
    return b""


# ── Search ────────────────────────────────────────────────────────────────── #
BASE = "https://gwar.mil.ru"


async def _fill_field(page, fid: str, value: str, log):
    if not value:
        return
    try:
        el = page.locator(f"#{fid}").first
        if await el.count():
            await el.scroll_into_view_if_needed(timeout=2000)
            await el.click(timeout=3000)
            await page.keyboard.press("Control+a")
            await page.keyboard.press("Delete")
            await page.keyboard.type(str(value), delay=20)
            log(f"  ✓ {fid} = {value!r}")
    except Exception as e:
        log(f"  !! поле {fid}: {type(e).__name__}")


async def _do_search(page, params, log) -> bool:
    log(f"  → Открываю поиск: {BASE}/heroes/")
    await page.goto(f"{BASE}/heroes/", wait_until="domcontentloaded", timeout=40000)
    await asyncio.sleep(2)

    # Expand «Дополнительные параметры поиска» if collapsed.
    try:
        more = page.get_by_text("Больше параметров", exact=False).first
        if await more.count() and await more.is_visible():
            await more.click(timeout=2000)
            await asyncio.sleep(0.5)
    except Exception:
        pass

    # Text fields (id == name on the page).
    await _fill_field(page, "last_name",   params.get("last_name", ""), log)
    await _fill_field(page, "first_name",  params.get("first_name", ""), log)
    await _fill_field(page, "middle_name", params.get("middle_name", ""), log)
    # «Дата рождения» has a date+month mask — only fill a REAL date (with dots);
    # a bare year (1889) would be mangled by the mask, so we keep it for the
    # result-side year filter instead.
    _bd = params.get("birth_date", "")
    if "." in _bd:
        await _fill_field(page, "birth_date", _bd, log)
    elif _bd:
        log("  → год рождения учту при фильтре результатов (поле даты не трогаю)")
    await _fill_field(page, "birth_place_gubernia", params.get("gubernia", ""), log)
    await _fill_field(page, "birth_place_uezd",     params.get("uezd", ""), log)
    await _fill_field(page, "birth_place_volost",   params.get("volost", ""), log)
    await _fill_field(page, "birth_place",          params.get("settlement", ""), log)
    await _fill_field(page, "rank",                 params.get("rank", ""), log)
    await _fill_field(page, "military_unit_name",   params.get("unit", ""), log)
    await _fill_field(page, "event_place",          params.get("event_place", ""), log)
    await _fill_field(page, "fund",      params.get("fund", ""), log)
    await _fill_field(page, "inventory", params.get("inventory", ""), log)
    await _fill_field(page, "file",      params.get("file", ""), log)

    # «Разделы» checkboxes — set each to the requested state (default: all on).
    sections = params.get("sections") or {}
    for label, cid in SECTIONS.items():
        want = sections.get(label, True)
        try:
            cb = page.locator(f"#{cid}").first
            if await cb.count():
                is_on = await cb.is_checked()
                if is_on != want:
                    lbl = page.locator(f'label[for="{cid}"]').first
                    target = lbl if await lbl.count() else cb
                    await target.click(timeout=2000)
        except Exception:
            pass

    # Submit «Найти».
    clicked = False
    for sel in ('input.button-search-big[value="Найти"]',
                'input.button-search-big', 'input[type="submit"][value="Найти"]',
                'input[type="submit"]'):
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=5000)
                clicked = True
                log(f"  ✓ Отправил поиск ({sel})")
                break
        except Exception:
            continue
    if not clicked:
        await page.keyboard.press("Enter")
        log("  → Поиск: Enter")

    for _ in range(30):
        await asyncio.sleep(1)
        try:
            ok = await page.evaluate(
                "() => document.querySelectorAll('a[href*=\"/heroes/chelovek\"]')"
                ".length > 0 || /найдено|ничего не/i.test(document.body.innerText)")
            if ok:
                return True
        except Exception:
            pass
    return True


def _with_page(url: str, n: int) -> str:
    """Return url with the `page` query parameter set to n."""
    sp = urlsplit(url)
    q = [(k, v) for k, v in parse_qsl(sp.query, keep_blank_values=True)
         if k != "page"]
    q.append(("page", str(n)))
    return urlunsplit((sp.scheme, sp.netloc, sp.path, urlencode(q, doseq=True),
                       sp.fragment))


_COLLECT_JS = r"""() => {
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const out = [], seen = new Set();
    document.querySelectorAll('a[href*="/heroes/chelovek"]').forEach(a => {
        const href = a.href || '';
        const name = norm(a.textContent);
        if (!href || !name || seen.has(href)) return;
        seen.add(href);
        // snippet = the smallest ancestor that adds context beyond the name
        let box = a;
        for (let up = 0; up < 5 && box.parentElement; up++) {
            box = box.parentElement;
            if (norm(box.textContent).length > name.length + 15) break;
        }
        out.push({href, name, snippet: norm(box ? box.textContent : '').slice(0, 400)});
    });
    return out;
}"""


async def _collect_results(page, params, log, max_pages, max_records) -> list:
    want = (params.get("last_name", ""), params.get("first_name", ""),
            params.get("middle_name", ""))
    ym = re.search(r"(18|19|20)\d{2}", params.get("birth_date", "") or "")
    want_year = ym.group(0) if ym else ""
    out, seen = [], set()
    results_url = page.url
    page_no = 1
    while page_no <= max_pages and len(out) < max_records:
        try:
            rows = await page.evaluate(_COLLECT_JS)
        except Exception:
            rows = []
        new = 0
        for r in rows:
            href = r.get("href", "")
            if not href or href in seen:
                continue
            seen.add(href)
            name = r.get("name", "")
            if not _person_matches(name, *want):
                continue
            if want_year:
                yrs = re.findall(r"(?:18|19|20)\d{2}", r.get("snippet", ""))
                if yrs and want_year not in yrs:
                    continue
            out.append({"name": name, "href": href, "snippet": r.get("snippet", "")})
            new += 1
            if len(out) >= max_records:
                break
        log(f"  → Страница {page_no}: ссылок {len(rows)}, подходящих +{new}, "
            f"всего {len(out)}")
        if len(out) >= max_records:
            break
        # next page via the page= query param; stop if it doesn't change
        first_before = rows[0]["href"] if rows else ""
        try:
            await page.goto(_with_page(results_url, page_no + 1),
                            wait_until="domcontentloaded", timeout=30000)
            await asyncio.sleep(1.5)
            changed = False
            for _ in range(12):
                cur = await page.evaluate(
                    "() => { const a = document.querySelector("
                    "'a[href*=\"/heroes/chelovek\"]'); return a ? a.href : ''; }")
                if cur and cur != first_before:
                    changed = True; break
                await asyncio.sleep(1)
            if not changed:
                break
        except Exception:
            break
        page_no += 1
    return out


# ── Record (detail) page ──────────────────────────────────────────────────── #
_DETAIL_JS = r"""(labels) => {
    const norm = s => (s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
    const h1 = document.querySelector('h1');
    const name = h1 ? norm(h1.textContent) : '';
    // card type = first non-empty short text right after the h1
    let typ = '';
    if (h1) {
        let n = h1.nextElementSibling;
        while (n && !norm(n.textContent)) n = n.nextElementSibling;
        if (n) typ = norm(n.textContent).slice(0, 90);
    }
    return {name, typ, body: norm(document.body.innerText)};
}"""


async def _grab_doc_scans(page, images_dir, base, rec, log):
    """Save the scanned document — every page (Страница X из N). Per page: click
    #btnSaveImage (catch the download); if that fails, download the largest
    displayed image via the browser. Move pages with td.to-right-arrow."""
    images_dir.mkdir(parents=True, exist_ok=True)
    # how many pages does the document have?
    pages = 1
    try:
        t = await page.evaluate(
            "() => { const m=(document.body.innerText||'').match(/из\\s*(\\d+)/);"
            " return m ? m[1] : '1'; }")
        pages = max(1, min(int(t), 20))
    except Exception:
        pages = 1

    saved = 0
    for pi in range(pages):
        await asyncio.sleep(1.2)
        ok = False
        # 1) save button (icon-save-dropdown) → catch download
        try:
            btn = page.locator("#btnSaveImage").first
            if await btn.count():
                async with page.expect_download(timeout=15000) as dl:
                    await btn.click(timeout=4000)
                    await asyncio.sleep(0.6)
                    for sel in ('.icon-save-dropdown a', '.icon-save-dropdown li',
                                'a:has-text("Скачать")', 'a[href*="download" i]'):
                        it = page.locator(sel).first
                        try:
                            if await it.count() and await it.is_visible():
                                await it.click(timeout=2000); break
                        except Exception:
                            continue
                d = await dl.value
                suf = Path(d.suggested_filename or "scan.jpg").suffix or ".jpg"
                dest = images_dir / f"{base}_лист{pi + 1}{suf}"
                await d.save_as(str(dest))
                rec["scans"].append(str(dest)); saved += 1; ok = True
                log(f"      🖼 документ сохранён: {dest.name}")
        except Exception:
            ok = False
        # 2) fallback: largest displayed image → download via browser
        if not ok:
            try:
                src = await page.evaluate(r"""() => {
                    let best = '', area = 0;
                    for (const im of document.querySelectorAll('img[src]')) {
                        const s = (im.src || '').toLowerCase();
                        if (!s.startsWith('http')) continue;
                        if (/\.svg|sprite|logo|icon|placeholder|avatar/.test(s)) continue;
                        const a = (im.naturalWidth||0) * (im.naturalHeight||0);
                        if (a > area && a > 60000) { area = a; best = im.src; }
                    }
                    return best;
                }""")
                if src:
                    body = await _img_bytes_via_goto(page.context, src, page.url)
                    if body:
                        dest = images_dir / f"{base}_лист{pi + 1}.jpg"
                        dest.write_bytes(body)
                        rec["scans"].append(str(dest)); saved += 1; ok = True
                        log(f"      🖼 документ сохранён (img): {dest.name}")
            except Exception:
                pass
        if not ok:
            log(f"      (лист {pi + 1}: скан не сохранился)")
        # next page within the document
        if pi < pages - 1:
            try:
                nx = page.locator("td.to-right-arrow, .to-right-arrow").first
                if await nx.count():
                    await nx.click(timeout=3000)
                else:
                    break
            except Exception:
                break
    return saved > 0


async def _extract_record(page, url, images_dir, base_dir, log, result_name=""):
    rec = {"name": "", "type": "", "url": url, "fields": [], "scans": []}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2.0)
        info = await page.evaluate(_DETAIL_JS, _GWAR_LABELS)
        rec["name"] = info.get("name") or result_name
        rec["type"] = info.get("typ") or ""
        rec["fields"] = _parse_fields(info.get("body") or "")
        base = safe_fn(rec.get("name") or result_name or "record")
        log(f"      полей: {len(rec['fields'])}, тип: {rec['type'] or '—'}")
        await _grab_doc_scans(page, images_dir, base, rec, log)
    except Exception as e:
        log(f"      !! страница записи ({type(e).__name__})")
    rec["fields"] = [(k, _clean_block(v)) for k, v in rec.get("fields", [])
                     if _clean_block(v)]
    return rec


# ── Word output ───────────────────────────────────────────────────────────── #
def _kv_table(doc, pairs):
    if not pairs:
        return
    tbl = doc.add_table(rows=0, cols=2); tbl.style = "Table Grid"
    for k, v in pairs:
        r = tbl.add_row().cells
        r[0].text = str(k); r[1].text = str(v)
        for run in r[0].paragraphs[0].runs:
            run.bold = True


def _docx_add_record(doc, i, rec):
    name = rec.get("name") or f"Запись {i}"
    title = name + (f" — {rec['type']}" if rec.get("type") else "")
    hp = doc.add_paragraph(); hr = hp.add_run(f"{i}. {title}")
    hr.bold = True; hr.font.size = Pt(13)

    if rec.get("fields"):
        _kv_table(doc, rec["fields"])

    for img in rec.get("scans", []):
        try:
            png = _to_png(Path(img).read_bytes())
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(4.5))
        except Exception:
            pass

    if rec.get("url"):
        clean = re.sub(r"\?.*$", "", rec["url"]) or rec["url"]
        p = doc.add_paragraph()
        _add_hyperlink(p, "Источник", clean)
    doc.add_paragraph("")


def write_docx(path, records, qlines, append=False):
    if not _DOCX_OK:
        raise RuntimeError("python-docx не установлен")
    existing = append and Path(path).exists()
    if existing:
        doc = Document(str(path)); doc.add_page_break()
        ap = doc.add_paragraph(); ar = ap.add_run(f"➕ Добавлено {len(records)}")
        ar.bold = True; ar.font.size = Pt(13)
    else:
        doc = Document()
        s = doc.sections[0]
        s.page_width = Mm(210); s.page_height = Mm(297)
        s.left_margin = s.right_margin = Mm(18)
        s.top_margin = s.bottom_margin = Mm(18)
        ht = doc.add_paragraph()
        htr = ht.add_run("Памяти героев Великой войны — результаты поиска")
        htr.bold = True; htr.font.size = Pt(14)
        ht.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for ln in qlines:
            doc.add_paragraph(ln)
        doc.add_paragraph(f"Найдено записей: {len(records)}")
        doc.add_paragraph("")
    for i, rec in enumerate(records, 1):
        _docx_add_record(doc, i, rec)
    doc.save(str(path))


# ── Main entry point ──────────────────────────────────────────────────────── #
async def run_scraper(*,
    last_name="", first_name="", middle_name="",
    birth_date="", gubernia="", uezd="", volost="", settlement="",
    rank="", unit="", event_place="",
    fund="", inventory="", file="",
    sections=None,
    output_folder=Path("."),
    log=print,
    progress=None,
    cancel_event=None,
    ask_file_conflict=None,
    max_pages=20,
    max_records=80,
) -> dict:

    def _prog(pct, txt):
        log(txt)
        if progress:
            progress(pct, txt)

    def _done():
        return bool(cancel_event and cancel_event.is_set())

    params = dict(
        last_name=last_name.strip(), first_name=first_name.strip(),
        middle_name=middle_name.strip(), birth_date=birth_date.strip(),
        gubernia=gubernia.strip(), uezd=uezd.strip(), volost=volost.strip(),
        settlement=settlement.strip(), rank=rank.strip(), unit=unit.strip(),
        event_place=event_place.strip(), fund=fund.strip(),
        inventory=inventory.strip(), file=file.strip(),
        sections=sections or {})

    output_folder = Path(output_folder); output_folder.mkdir(parents=True, exist_ok=True)
    qkey = " ".join(p for p in (last_name, first_name, middle_name) if p) or "gwar"
    images_dir = output_folder / "images" / (safe_fn(qkey) or "gwar")
    summary = {"ok": False}

    if not any(params.get(k) for k in
               ("last_name", "first_name", "middle_name", "rank", "unit",
                "settlement", "gubernia", "uezd")):
        _prog(100, "Пустой запрос.")
        return summary

    qlines = [f"Запрос: {qkey}", "Сайт: gwar.mil.ru (1914–1918)"]

    _prog(0, "Запускаю браузер…")
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"])
        ctx = await browser.new_context(no_viewport=True, accept_downloads=True)
        page = await ctx.new_page()
        try:
            _prog(5, "Поиск…")
            if not await _do_search(page, params, log):
                summary["message"] = "Не удалось выполнить поиск."
                return summary
            if _done():
                return summary

            _prog(15, "Сбор результатов…")
            recs_meta = await _collect_results(page, params, log, max_pages, max_records)
            log(f"  Подходящих записей: {len(recs_meta)}")
            if not recs_meta:
                _prog(100, "Ничего подходящего не найдено.")
                summary.update({"ok": True, "n_records": 0})
                return summary

            records = []
            n = len(recs_meta)
            for i, rm in enumerate(recs_meta, 1):
                if _done():
                    break
                _prog(20 + int(70 * i / n), f"[{i}/{n}] {rm['name'][:50]}…")
                log(f"  [{i}/{n}] {rm['name']}")
                try:
                    dp = await ctx.new_page()
                    try:
                        rec = await _extract_record(dp, rm["href"], images_dir,
                                                    output_folder, log,
                                                    result_name=rm["name"])
                    finally:
                        await dp.close()
                    if not rec.get("name"):
                        rec["name"] = rm["name"]
                    records.append(rec)
                except Exception as _e:
                    log(f"      !! пропускаю запись ({type(_e).__name__})")
                    records.append({"name": rm["name"], "type": "", "url": rm["href"],
                                    "fields": [], "scans": []})
                await asyncio.sleep(0.3)

            _prog(92, "Сохранение…")
            base = safe_fn(f"gwar_{qkey}") or "gwar_results"
            docx_p = output_folder / f"{base}.docx"
            decision = "overwrite"
            if docx_p.exists() and ask_file_conflict:
                try:
                    decision = (ask_file_conflict([docx_p.name]) or "overwrite").lower()
                except Exception:
                    decision = "overwrite"
                log(f"  → Файл существует → {decision}")
            if decision != "skip" and records:
                try:
                    write_docx(docx_p, records, qlines, append=(decision == "append"))
                    log(f"  → Word: {docx_p.name}")
                except PermissionError:
                    alt = output_folder / f"{base}_{time.strftime('%H%M%S')}.docx"
                    write_docx(alt, records, qlines, append=False)
                    docx_p = alt
                    log(f"  !! файл занят (открыт в Word) → сохранил как {alt.name}")

            _prog(100, f"Готово — {len(records)} запис(ей).")
            summary.update({"ok": True, "n_records": len(records),
                            "output_folder": str(output_folder)})
        except Exception as exc:
            summary["message"] = f"{type(exc).__name__}: {exc}"
            log(f"  !! Ошибка: {exc}")
        finally:
            try:
                await browser.close()
            except Exception:
                pass
    return summary
