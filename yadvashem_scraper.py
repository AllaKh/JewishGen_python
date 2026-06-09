#!/usr/bin/env python3
"""
yadvashem_scraper.py
====================
Scraper for the Yad Vashem «Central Database of Shoah Victims' Names»
(https://collections.yadvashem.org/ru/names/search-results-names).

Open database, no login. The search is performed by URL parameters
(gn_<field> for the value, gnt_<field> for the match type). The scraper builds
the search URL, collects the matching victim records (paginating), opens each
record and copies its fields to Word, saving any document image (Page of
Testimony scan) it finds.

NOTE: the exact URL parameter names and the result/detail DOM are being verified
against the live site (the tooling can't reach the domain). PARAM_MAP and the
result/detail selectors are best-effort and easy to correct.

NEVER crash on a broken/empty page: every record is wrapped in try/except.
"""

import asyncio
import io
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlencode, urlparse, urljoin

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

BASE = "https://collections.yadvashem.org"
SEARCH_PATH = "/ru/names/search-results-names"

# GUI «search type» label → URL value (gnt_*). Confirmed: yvSynonym.
SEARCH_TYPES = {
    "YV synonyms": "yvSynonym",
    "Literal":     "literal",
    "Phonetic":    "phonetic",
}
# GUI «birth/death year precision» → value.
YEAR_PREC = {
    "Exact": "0",
    "± 2":   "2",
    "± 5":   "5",
}

# GUI field key → URL parameter (gn_*). The match type goes to gnt_<same suffix>.
# CONFIRMED from the address bar: last_name→gn_last_name, gnt_last_name; place→gn_place.
PARAM_MAP = {
    "last_name":      "gn_last_name",
    "first_name":     "gn_first_name",
    "maiden_name":    "gn_maiden_name",
    "birth_place":    "gn_birth_place",
    "place_before":   "gn_place_before_war",
    "place_during":   "gn_place_during_war",
    "death_place":    "gn_death_place",
    "father_name":    "gn_father_name",
    "mother_name":    "gn_mother_name",
    "mother_maiden":  "gn_mother_maiden_name",
    "spouse_name":    "gn_spouse_name",
    "spouse_maiden":  "gn_spouse_maiden_name",
    "submitter_first": "gn_submitter_first_name",
    "submitter_last":  "gn_submitter_last_name",
    "birth_year":     "gn_birth_year",
    "death_year":     "gn_death_year",
}
_YEAR_TYPE_PARAM = {"birth_year": "gn_birth_year_type",
                    "death_year": "gn_death_year_type"}


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


# ── Build the search URL from the GUI fields ──────────────────────────────── #
def _build_search_url(fields: dict, place_mode: str, global_text: str) -> str:
    q = [("page", "1")]
    for key, param in PARAM_MAP.items():
        val = (fields.get(key) or "").strip()
        if not val:
            continue
        # «any place» mode: collapse the specific place fields into gn_place
        if place_mode == "anyplace" and key in (
                "birth_place", "place_before", "place_during", "death_place"):
            continue
        q.append((param, val))
        # match type → gnt_<suffix> (for name fields) / year precision
        t_label = fields.get(key + "_type")
        if key in _YEAR_TYPE_PARAM and t_label:
            q.append((_YEAR_TYPE_PARAM[key], YEAR_PREC.get(t_label, "0")))
        elif t_label:
            q.append(("gnt_" + param[3:], SEARCH_TYPES.get(t_label, "yvSynonym")))
    if place_mode == "anyplace":
        anyplace = next((fields.get(k) for k in
                         ("birth_place", "place_before", "place_during", "death_place")
                         if (fields.get(k) or "").strip()), "")
        if anyplace:
            q.append(("gn_place", anyplace.strip()))
    if (global_text or "").strip():
        q.append(("gn_freetext", global_text.strip()))
    return f"{BASE}{SEARCH_PATH}?{urlencode(q, doseq=True)}"


# ── Results ───────────────────────────────────────────────────────────────── #
# Generic: a record detail link points under /ru/names/<id> (not …search…). We
# collect those + the surrounding row text (name / year / place / fate / source).
_COLLECT_JS = r"""() => {
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const out = [], seen = new Set();
    const links = document.querySelectorAll(
        'a[href*="/names/"]:not([href*="search"]):not([href*="advanced"])');
    links.forEach(a => {
        const href = a.href || '';
        if (!/\/names\/\d/.test(href) || seen.has(href)) return;
        seen.add(href);
        let box = a;
        for (let up = 0; up < 6 && box.parentElement; up++) {
            box = box.parentElement;
            if (norm(box.innerText).length > 30) break;
        }
        out.push({href, name: norm(a.innerText),
                  snippet: norm(box ? box.innerText : '').slice(0, 400)});
    });
    return out;
}"""


async def _collect_results(page, log, max_pages, max_records) -> list:
    out, seen = [], set()
    base_url = page.url
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
            out.append({"name": r.get("name", ""), "href": href,
                        "snippet": r.get("snippet", "")})
            new += 1
            if len(out) >= max_records:
                break
        log(f"  → Страница {page_no}: ссылок {len(rows)}, новых {new}, всего {len(out)}")
        if len(out) >= max_records or new == 0:
            break
        # pagination by the page= query param
        try:
            nxt = re.sub(r"([?&]page=)\d+", lambda m: m.group(1) + str(page_no + 1),
                         base_url)
            if "page=" not in nxt:
                sep = "&" if "?" in nxt else "?"
                nxt = f"{nxt}{sep}page={page_no + 1}"
            first_before = rows[0]["href"] if rows else ""
            await page.goto(nxt, wait_until="domcontentloaded", timeout=30000)
            await asyncio.sleep(2)
            changed = False
            for _ in range(12):
                cur = await page.evaluate(
                    "() => { const a=document.querySelector("
                    "'a[href*=\"/names/\"]:not([href*=\"search\"])'); return a?a.href:''; }")
                if cur and cur != first_before:
                    changed = True; break
                await asyncio.sleep(1)
            if not changed:
                break
        except Exception:
            break
        page_no += 1
    return out


# ── Record detail ─────────────────────────────────────────────────────────── #
_DETAIL_JS = r"""() => {
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const h1 = document.querySelector('h1, h2, [class*="title" i]');
    const name = h1 ? norm(h1.textContent) : '';
    // field rows: label/value pairs anywhere on the page
    const pairs = [];
    document.querySelectorAll('tr, li, [class*="field" i], [class*="detail" i], dl > *')
        .forEach(e => {
            const t = norm(e.innerText);
            const m = t.match(/^(.{2,40}?)\s*[:\-–]\s+(.{1,200})$/);
            if (m) pairs.push([m[1].trim(), m[2].trim()]);
        });
    // biggest content image (Page of Testimony scan / photo)
    let img = '', area = 0;
    document.querySelectorAll('img[src]').forEach(im => {
        const s = (im.src || '').toLowerCase();
        if (/\.svg|sprite|logo|icon|placeholder|avatar/.test(s)) return;
        const a = (im.naturalWidth || 0) * (im.naturalHeight || 0);
        if (a > area && a > 40000) { area = a; img = im.src; }
    });
    return {name, pairs: pairs.slice(0, 60), img};
}"""


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


async def _extract_record(page, url, images_dir, log, result_name=""):
    rec = {"name": "", "url": url, "fields": [], "images": []}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2)
        info = await page.evaluate(_DETAIL_JS)
        rec["name"] = info.get("name") or result_name
        # de-dupe pairs, keep order
        seen, fields = set(), []
        for k, v in info.get("pairs", []):
            kk = (k, v)
            if kk not in seen and len(k) < 45:
                seen.add(kk); fields.append((k, v))
        rec["fields"] = fields
        base = safe_fn(rec.get("name") or result_name or "record")
        log(f"      полей: {len(rec['fields'])}")
        img = info.get("img")
        if img:
            body = await _img_bytes_via_goto(page.context, img, page.url)
            if body:
                images_dir.mkdir(parents=True, exist_ok=True)
                dest = images_dir / f"{base}.jpg"
                dest.write_bytes(body)
                rec["images"].append(str(dest))
                log(f"      🖼 изображение: {dest.name}")
    except Exception as e:
        log(f"      !! страница записи ({type(e).__name__})")
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
    name = rec.get("name") or f"Record {i}"
    hp = doc.add_paragraph(); hr = hp.add_run(f"{i}. {name}")
    hr.bold = True; hr.font.size = Pt(13)
    for img in rec.get("images", []):
        try:
            png = _to_png(Path(img).read_bytes())
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(2.6))
        except Exception:
            pass
    if rec.get("fields"):
        _kv_table(doc, rec["fields"])
    if rec.get("url"):
        p = doc.add_paragraph(); _add_hyperlink(p, "Source / Источник", rec["url"])
    doc.add_paragraph("")


def write_docx(path, records, qlines, append=False):
    if not _DOCX_OK:
        raise RuntimeError("python-docx не установлен")
    if append and Path(path).exists():
        doc = Document(str(path)); doc.add_page_break()
        ap = doc.add_paragraph(); ar = ap.add_run(f"➕ Added {len(records)}")
        ar.bold = True; ar.font.size = Pt(13)
    else:
        doc = Document()
        s = doc.sections[0]
        s.page_width = Mm(210); s.page_height = Mm(297)
        s.left_margin = s.right_margin = Mm(18)
        s.top_margin = s.bottom_margin = Mm(18)
        ht = doc.add_paragraph()
        htr = ht.add_run("Yad Vashem — Shoah Victims' Names — search results")
        htr.bold = True; htr.font.size = Pt(14)
        ht.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for ln in qlines:
            doc.add_paragraph(ln)
        doc.add_paragraph(f"Records: {len(records)}")
        doc.add_paragraph("")
    for i, rec in enumerate(records, 1):
        _docx_add_record(doc, i, rec)
    doc.save(str(path))


# ── Main entry point ──────────────────────────────────────────────────────── #
async def run_scraper(*,
    fields=None, place_mode="byfield", global_text="",
    output_folder=Path("."),
    log=print, progress=None, cancel_event=None, ask_file_conflict=None,
    max_pages=10, max_records=60,
) -> dict:

    def _prog(pct, txt):
        log(txt)
        if progress:
            progress(pct, txt)

    def _done():
        return bool(cancel_event and cancel_event.is_set())

    fields = fields or {}
    output_folder = Path(output_folder); output_folder.mkdir(parents=True, exist_ok=True)
    qkey = " ".join(v for k, v in fields.items()
                    if not k.endswith("_type") and v) or global_text or "yadvashem"
    images_dir = output_folder / "images" / (safe_fn(qkey) or "yadvashem")
    summary = {"ok": False}

    if not any(v for k, v in fields.items() if not k.endswith("_type") and v) \
            and not global_text:
        _prog(100, "Пустой запрос.")
        return summary

    url = _build_search_url(fields, place_mode, global_text)
    qlines = [f"Запрос: {qkey}", f"URL: {url}"]

    _prog(0, "Запускаю браузер…")
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"])
        ctx = await browser.new_context(no_viewport=True, accept_downloads=True,
                                        ignore_https_errors=True)
        page = await ctx.new_page()
        try:
            _prog(5, "Поиск…")
            log(f"  → Открываю поиск: {url}")
            await page.goto(url, wait_until="domcontentloaded", timeout=45000)
            await asyncio.sleep(4)            # SPA renders results via XHR
            if _done():
                return summary

            _prog(15, "Сбор результатов…")
            recs_meta = await _collect_results(page, log, max_pages, max_records)
            log(f"  Записей: {len(recs_meta)}")
            if not recs_meta:
                _prog(100, "Ничего не найдено (или селекторы результата надо поправить).")
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
                        rec = await _extract_record(dp, rm["href"], images_dir, log,
                                                    result_name=rm["name"])
                    finally:
                        await dp.close()
                    if not rec.get("name"):
                        rec["name"] = rm["name"]
                    records.append(rec)
                except Exception as _e:
                    log(f"      !! пропускаю запись ({type(_e).__name__})")
                    records.append({"name": rm["name"], "url": rm["href"],
                                    "fields": [], "images": []})
                await asyncio.sleep(0.3)

            _prog(92, "Сохранение…")
            base = safe_fn(f"yadvashem_{qkey}") or "yadvashem_results"
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
                    log(f"  !! файл занят → сохранил как {alt.name}")

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
