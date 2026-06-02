#!/usr/bin/env python3
"""
skarb_scraper.py
================
Scraper for АИС Скарб (https://archiveskarb.by/db/)

Search by one or more surnames (OR logic — joined with " или ").
Filter results by keyword(s).
For each matching record: scrape all fields, download photos if available.
Save to Word (.docx) and/or Excel (.xlsx).

No login required.
"""

import asyncio
import io
import re
import sys
from pathlib import Path
from urllib.parse import urljoin, urlencode, quote

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
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# ── Constants ─────────────────────────────────────────────────────────────── #
BASE_URL    = "https://archiveskarb.by/db/"
LOGO_SRCS   = ("/bitrix/templates/", "/local/templates/")  # skip these images
HYPERLINK_REL = ("http://schemas.openxmlformats.org/"
                 "officeDocument/2006/relationships/hyperlink")


# ── Helpers ───────────────────────────────────────────────────────────────── #

def safe_fn(s: str, n: int = 80) -> str:
    s = re.sub(r'[\\/*?:"<>|]', "_", (s or "").strip())
    return re.sub(r"\s+", "_", s)[:n].strip("_") or "document"


def _matches(text: str, keywords: list, mode: str) -> bool:
    """Return True if text satisfies keyword filter."""
    if not keywords:
        return True
    text_l = text.lower()
    hits = [k.lower() in text_l for k in keywords]
    return all(hits) if mode == "AND" else any(hits)


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


# ── Page scraping ─────────────────────────────────────────────────────────── #

async def _collect_rows(page) -> list:
    """
    Parse the results table on the current page.
    Returns list of dicts: {name, birth, category, url}
    """
    rows = []
    try:
        trs = await page.query_selector_all("table tr")
        for tr in trs:
            tds = await tr.query_selector_all("td")
            if len(tds) < 2:
                continue
            # First cell: name link
            a = await tds[0].query_selector("a[href]")
            if not a:
                continue
            name = (await a.text_content() or "").strip()
            href = (await a.get_attribute("href") or "").strip()
            if not name or not href:
                continue
            url = urljoin(BASE_URL, href)
            birth    = (await tds[1].text_content() or "").strip() if len(tds) > 1 else ""
            category = (await tds[2].text_content() or "").strip() if len(tds) > 2 else ""
            rows.append({"name": name, "birth": birth,
                         "category": category, "url": url})
    except Exception:
        pass
    return rows


async def _get_next_url(page) -> str | None:
    """Find the 'След.' (Next) pagination link."""
    try:
        # Try text "След." first
        for sel in ['a:has-text("След.")', 'a:has-text("Next")', 'a[title*="След"]']:
            el = page.locator(sel).first
            if await el.count():
                href = await el.get_attribute("href") or ""
                if href:
                    return urljoin(BASE_URL, href)
    except Exception:
        pass
    return None


async def _scrape_record(page, url: str, images_dir: Path, log) -> dict:
    """
    Open a record page and scrape all fields.
    If "Наличие фото: да" — download the photo.
    """
    rec = {"url": url, "fields": {}, "image_path": None}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(1.5)

        # ── Extract all label: value pairs ──────────────────────── #
        content = await page.content()

        # Strategy 1: look for dl dt/dd
        dts = await page.query_selector_all("dl dt, .field-label, td.label")
        dds = await page.query_selector_all("dl dd, .field-value, td.value")
        for dt, dd in zip(dts, dds):
            k = (await dt.text_content() or "").strip().rstrip(":")
            v = (await dd.text_content() or "").strip()
            if k and v:
                rec["fields"][k] = v

        # Strategy 2: table rows with 2 cells (label | value)
        if not rec["fields"]:
            for tr in await page.query_selector_all("table tr"):
                tds = await tr.query_selector_all("td")
                if len(tds) == 2:
                    k = (await tds[0].text_content() or "").strip().rstrip(":")
                    v = (await tds[1].text_content() or "").strip()
                    if k and v and len(k) < 80:
                        rec["fields"][k] = v

        # Strategy 3: text pattern "Label: Value" from paragraphs
        if not rec["fields"]:
            for el in await page.query_selector_all("p, li, div.item"):
                txt = (await el.text_content() or "").strip()
                m = re.match(r"^([^:]{1,50}):\s*(.+)$", txt, re.DOTALL)
                if m:
                    rec["fields"][m.group(1).strip()] = m.group(2).strip()

        log(f"      Полей: {len(rec['fields'])}")

        # ── Photo ────────────────────────────────────────────────── #
        has_photo = False
        for k, v in rec["fields"].items():
            if "фото" in k.lower() and "да" in v.lower():
                has_photo = True
                break

        if has_photo:
            log("      Наличие фото: да — ищу изображение...")
            # Find the photo img tag (skip logos/templates)
            for img in await page.query_selector_all("img[src]"):
                src = (await img.get_attribute("src") or "").strip()
                if not src or not src.startswith("/"):
                    if not src.startswith("http"):
                        continue
                if any(skip in src for skip in LOGO_SRCS):
                    continue
                # This looks like a content image
                abs_src = urljoin(BASE_URL, src)
                # Download it
                img_path = await _download_image(page, abs_src, images_dir,
                                                 url, log)
                if img_path:
                    rec["image_path"] = img_path
                    break
        else:
            # Even without "Наличие фото: да", check for content images
            for img in await page.query_selector_all("img[src]"):
                src = (await img.get_attribute("src") or "").strip()
                if not src:
                    continue
                if any(skip in src for skip in LOGO_SRCS):
                    continue
                abs_src = urljoin(BASE_URL, src)
                img_path = await _download_image(page, abs_src, images_dir,
                                                 url, log)
                if img_path:
                    rec["image_path"] = img_path
                    break

    except Exception as e:
        log(f"      !! record error: {e}")
    return rec


async def _download_image(page, src: str, images_dir: Path,
                          record_url: str, log) -> str | None:
    """Download an image by opening its URL in a new page."""
    try:
        images_dir.mkdir(parents=True, exist_ok=True)
        # Get filename from URL
        fname = src.rstrip("/").split("/")[-1]
        if not fname or "." not in fname:
            fname = "photo.jpg"
        dest = images_dir / safe_fn(fname, 60)

        # Open new page for image
        img_page = await page.context.new_page()
        try:
            resp = await img_page.goto(src, timeout=20000)
            if resp and resp.ok:
                body = await resp.body()
                if len(body) > 1000:
                    dest.write_bytes(body)
                    log(f"      Фото: {dest.name} ({len(body)//1024}KB)")
                    return str(dest)
        finally:
            await img_page.close()
    except Exception as e:
        log(f"      !! image download: {e}")
    return None


async def _get_thumbnail(page, record_url: str, log) -> bytes | None:
    """Get a small preview of the photo for Word document."""
    try:
        rec_page = await page.context.new_page()
        try:
            await rec_page.goto(record_url, wait_until="domcontentloaded", timeout=15000)
            for img in await rec_page.query_selector_all("img[src]"):
                src = (await img.get_attribute("src") or "").strip()
                if not src or any(skip in src for skip in LOGO_SRCS):
                    continue
                abs_src = urljoin(BASE_URL, src)
                resp_p = await rec_page.context.new_page()
                try:
                    resp = await resp_p.goto(abs_src, timeout=10000)
                    if resp and resp.ok:
                        body = await resp.body()
                        if len(body) > 500:
                            return body
                finally:
                    await resp_p.close()
        finally:
            await rec_page.close()
    except Exception:
        pass
    return None


# ── Word output ───────────────────────────────────────────────────────────── #

def write_docx(path: Path, records: list, query_info: dict):
    if not _DOCX_OK:
        raise RuntimeError("python-docx не установлен")
    doc = Document()
    s = doc.sections[0]
    s.page_width  = Mm(210); s.page_height = Mm(297)
    s.left_margin = s.right_margin = Mm(20)
    s.top_margin  = s.bottom_margin = Mm(20)

    # Title
    h = doc.add_heading("АИС Скарб — Результаты поиска", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Query info
    for k, v in query_info.items():
        p = doc.add_paragraph()
        p.add_run(f"{k}: ").bold = True
        p.add_run(str(v))
    doc.add_paragraph(f"Найдено записей: {len(records)}")
    doc.add_paragraph("")

    for i, rec in enumerate(records, 1):
        name = rec.get("fields", {}).get(
            next((k for k in rec.get("fields", {}) if "фио" in k.lower() or "имя" in k.lower()), ""),
            "") or f"Запись {i}"
        doc.add_heading(f"{i}. {name}", level=2)

        # URL
        if rec.get("url"):
            pp = doc.add_paragraph()
            pp.add_run("Источник: ").bold = True
            _add_hyperlink(pp, rec["url"], rec["url"])

        # Fields table
        fields = rec.get("fields", {})
        if fields:
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            hdr[0].text = "Поле"; hdr[1].text = "Значение"
            for cell in hdr:
                for run in cell.paragraphs[0].runs:
                    run.bold = True
            for k, v in fields.items():
                r = tbl.add_row().cells
                r[0].text = str(k); r[1].text = str(v)

        # Image
        img_path = rec.get("image_path")
        thumb    = rec.get("thumb_bytes")
        if img_path and Path(img_path).exists():
            doc.add_paragraph("Изображение:").runs[0].bold = True
            try:
                doc.add_picture(img_path, width=Inches(3.5))
            except Exception:
                doc.add_paragraph(f"  [{Path(img_path).name}]")
        elif thumb:
            doc.add_paragraph("Превью:").runs[0].bold = True
            try:
                doc.add_picture(io.BytesIO(thumb), width=Inches(3.5))
            except Exception:
                pass

        doc.add_paragraph("")

    doc.save(str(path))


# ── Excel output ──────────────────────────────────────────────────────────── #

def write_xlsx(path: Path, records: list, query_info: dict):
    if not _OPENPYXL_OK:
        raise RuntimeError("openpyxl не установлен")
    wb = Workbook(); ws = wb.active; ws.title = "Скарб"
    HF = PatternFill("solid", fgColor="2A4A7F")
    HN = Font(bold=True, color="FFFFFF", size=11)
    TS = Side(style="thin", color="B0B8C8")
    T  = Border(left=TS, right=TS, top=TS, bottom=TS)

    # Collect all field names
    all_fields: list = []
    for rec in records:
        for k in rec.get("fields", {}):
            if k not in all_fields:
                all_fields.append(k)

    cols = ["#", "Имя", "Дата рождения", "Категория", "Фото", "URL"] + all_fields
    for ci, cn in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=cn)
        c.font = HN; c.fill = HF; c.border = T
        c.alignment = Alignment(horizontal="center", vertical="center",
                                 wrap_text=True)

    for ri, rec in enumerate(records, 2):
        fields = rec.get("fields", {})
        # Try to guess name from fields
        name = ""
        for k in fields:
            if any(x in k.lower() for x in ("фио", "имя", "name")):
                name = fields[k]; break
        birth = fields.get("Дата рождения", fields.get("Год рождения", ""))
        category = fields.get("Категория", fields.get("Тип документа", ""))
        has_img = "да" if rec.get("image_path") else "нет"
        vals = [ri-1, name, birth, category, has_img, rec.get("url", "")
                ] + [fields.get(f, "") for f in all_fields]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = T
            c.alignment = Alignment(wrap_text=True, vertical="top")

    for ci in range(1, len(cols)+1):
        ltr = get_column_letter(ci)
        mw = max(len(str(cols[ci-1])),
                 *(len(str(ws.cell(row=r, column=ci).value or ""))
                   for r in range(2, ws.max_row+1)), 8)
        ws.column_dimensions[ltr].width = min(mw + 4, 60)
    wb.save(str(path))


# ── Main scraper ──────────────────────────────────────────────────────────── #

async def run_scraper(
    *,
    surnames:      list        = None,   # list of surnames to search
    keywords:      list        = None,   # filter keywords
    keyword_mode:  str         = "OR",
    output_format: str         = "both",
    output_folder              = Path("."),
    log                        = print,
    progress                   = None,
    cancel_event               = None,
) -> dict:

    def _prog(pct, txt):
        log(txt)
        if progress: progress(pct, txt)

    def _done():
        return bool(cancel_event and cancel_event.is_set())

    surnames      = [s.strip() for s in (surnames or []) if s.strip()]
    keywords      = [k.strip() for k in (keywords or []) if k.strip()]
    want_docx     = output_format in ("docx", "both")
    want_xlsx     = output_format in ("xlsx", "both")
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    images_dir    = output_folder / "images"
    summary       = {"ok": False}

    if not surnames:
        _prog(100, "Не указаны фамилии для поиска.")
        return summary

    # Build search query: "Фамилия1 или Фамилия2 или ..."
    query = " или ".join(surnames)
    q_prefix = safe_fn("_".join(surnames[:3]), 50)

    query_info = {
        "Фамилии":  query,
        "Ключевые слова": " " + keyword_mode + " ".join(keywords) if keywords else "(нет)",
    }

    _prog(0, "Запускаю браузер...")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=["--start-maximized",
                  "--disable-blink-features=AutomationControlled"],
        )
        ctx  = await browser.new_context(no_viewport=True, accept_downloads=True)
        page = await ctx.new_page()

        try:
            # ── 1. Search ──────────────────────────────────────────── #
            # Build URL: GET parameter ФИО
            search_url = BASE_URL + "?" + urlencode({"ФИО": query})
            _prog(5, f"Поиск: {query}")
            await page.goto(search_url, wait_until="domcontentloaded", timeout=30000)
            await asyncio.sleep(2)

            # ── 2. Collect all pages ───────────────────────────────── #
            _prog(10, "Сбор результатов по всем страницам...")
            all_rows: list = []
            page_num = 1

            while not _done():
                rows = await _collect_rows(page)
                log(f"  Страница {page_num}: {len(rows)} строк")

                for r in rows:
                    row_text = f"{r['name']} {r['birth']} {r['category']}"
                    if _matches(row_text, keywords, keyword_mode):
                        all_rows.append(r)

                next_url = await _get_next_url(page)
                if not next_url:
                    break
                page_num += 1
                await page.goto(next_url, wait_until="domcontentloaded", timeout=30000)
                await asyncio.sleep(1.5)

            log(f"  Итого совпадений: {len(all_rows)}")
            if not all_rows:
                _prog(100, "Ничего не найдено по заданным критериям.")
                summary.update({"ok": True, "n_records": 0})
                return summary

            # ── 3. Scrape individual records ───────────────────────── #
            records: list = []
            n = len(all_rows)

            for i, row in enumerate(all_rows, 1):
                if _done(): break
                _prog(20 + int(70 * i / n),
                      f"[{i}/{n}] {row['name'][:60]}...")
                log(f"  [{i}/{n}] {row['name']}")

                rec = await _scrape_record(page, row["url"], images_dir, log)
                # Merge search-table data into fields if not already there
                if "Дата рождения" not in rec["fields"] and row["birth"]:
                    rec["fields"]["Дата рождения"] = row["birth"]
                if "Категория" not in rec["fields"] and row["category"]:
                    rec["fields"]["Категория"] = row["category"]
                # Also try to get thumbnail for Word
                if rec.get("image_path") and Path(rec["image_path"]).exists():
                    try:
                        rec["thumb_bytes"] = Path(rec["image_path"]).read_bytes()
                    except Exception:
                        pass
                records.append(rec)

            # ── 4. Save ────────────────────────────────────────────── #
            _prog(92, "Сохранение файлов...")
            if want_docx and records:
                p = output_folder / f"{q_prefix}_skarb.docx"
                write_docx(p, records, query_info)
                log(f"  Word: {p}")
            if want_xlsx and records:
                p = output_folder / f"{q_prefix}_skarb.xlsx"
                write_xlsx(p, records, query_info)
                log(f"  Excel: {p}")

            _prog(100, f"Готово — {len(records)} записей.")
            summary.update({"ok": True, "n_records": len(records),
                            "output_folder": str(output_folder)})

        except Exception as exc:
            summary["message"] = f"{type(exc).__name__}: {exc}"
            log(f"  !! {exc}")
        finally:
            try: await ctx.close()
            except Exception: pass
            try: await browser.close()
            except Exception: pass

    return summary
