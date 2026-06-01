#!/usr/bin/env python3
"""
myheritage_scraper.py
---------------------
MyHeritage search scraper — driven by Playwright.

Flow
----
1. Opens https://www.myheritage.com and logs in with the provided credentials.
2. Navigates to the search form and fills in first/patronymic name and surname.
3. Applies the selected record-type filter (All Records / Historical Records /
   Family Trees).
4. Scrapes every result that has a match score ≥ 80 %.
5. For each qualifying result: opens the detail page, copies full name,
   category, and all data from the detail table.
6. Saves matched records to .docx and/or .xlsx in the output folder.

Match-score logic
-----------------
MyHeritage shows a relevance percentage next to each result.  We only keep
results where that score is ≥ MIN_MATCH_PCT (80 by default).
If MyHeritage does not expose a numeric score we fall back to string-similarity
(difflib.SequenceMatcher) between the query name and the result name.
"""

import asyncio
import difflib
import os
import re
import sys
from pathlib import Path

# ── Playwright ────────────────────────────────────────────────────────────── #
if getattr(sys, "frozen", False):
    exe_dir = Path(sys.executable).resolve().parent
    browser_dir = exe_dir / "ms-playwright"
    if browser_dir.exists():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browser_dir)

try:
    from playwright.async_api import async_playwright, TimeoutError as PWTimeout
except ImportError:
    sys.stderr.write(
        "Playwright is not installed. Run:\n"
        "    pip install playwright\n"
        "    playwright install chromium\n"
    )
    sys.exit(1)

# ── python-docx ───────────────────────────────────────────────────────────── #
try:
    from docx import Document
    from docx.shared import Pt, Mm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

# ── openpyxl ─────────────────────────────────────────────────────────────── #
try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False


# ═══════════════════════════════════════════════════════════════════  CONST  ═ #

SEARCH_URL     = "https://www.myheritage.com/research"
LOGIN_URL      = "https://www.myheritage.com/login"
HOME_URL       = "https://www.myheritage.com"
MIN_MATCH_PCT  = 80          # keep results with ≥ this match score

FILTER_OPTIONS = ["All Records", "Historical Records", "Family Trees"]

PROFILE_DIR    = Path(__file__).resolve().parent / ".mh_profile"

HYPERLINK_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)


# ═══════════════════════════════════════════════════════════════  HELPERS  ═══ #

def _name_similarity(query: str, candidate: str) -> float:
    """Return 0-100 string similarity between two names."""
    a = query.strip().lower()
    b = candidate.strip().lower()
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio() * 100


def safe_filename(s: str, max_len: int = 80) -> str:
    s = re.sub(r'[\\/*?:"<>|]', "_", s)
    s = re.sub(r"\s+", "_", s.strip())
    return s[:max_len] or "result"


# ── docx helpers ────────────────────────────────────────────────────────────  #

def _add_hyperlink(paragraph, text, url):
    part  = paragraph.part
    r_id  = part.relate_to(url, HYPERLINK_REL, is_external=True)
    hl    = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), r_id)
    run   = OxmlElement("w:r")
    rPr   = OxmlElement("w:rPr")
    color = OxmlElement("w:color"); color.set(qn("w:val"), "0563C1")
    ul    = OxmlElement("w:u");     ul.set(qn("w:val"), "single")
    rPr.append(color); rPr.append(ul)
    run.append(rPr)
    t = OxmlElement("w:t"); t.text = text or url
    t.set(qn("xml:space"), "preserve")
    run.append(t); hl.append(run); paragraph._p.append(hl)


def write_docx(path: Path, records: list[dict], query_lines: list[str]):
    """Write matched records to a .docx file."""
    if not _DOCX_OK:
        raise RuntimeError("python-docx not installed")

    doc = Document()
    # Page size A4 landscape
    sec = doc.sections[0]
    sec.page_width  = Mm(297); sec.page_height = Mm(210)
    sec.left_margin = sec.right_margin = Mm(18)
    sec.top_margin  = sec.bottom_margin = Mm(15)

    # Title
    h = doc.add_heading("MyHeritage Search Results", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Query summary
    doc.add_paragraph("Search parameters:")
    for ln in query_lines:
        doc.add_paragraph(ln, style="List Bullet")

    doc.add_paragraph(f"Records saved: {len(records)}  (match ≥ {MIN_MATCH_PCT} %)")
    doc.add_paragraph("")

    for idx, rec in enumerate(records, 1):
        doc.add_heading(f"{idx}. {rec.get('full_name', '—')}", level=2)

        # Category + score
        p = doc.add_paragraph()
        p.add_run("Category: ").bold = True
        p.add_run(rec.get("category", "—"))
        p2 = doc.add_paragraph()
        p2.add_run("Match score: ").bold = True
        p2.add_run(f"{rec.get('score', '?')} %")

        # Source URL
        if rec.get("url"):
            p3 = doc.add_paragraph()
            p3.add_run("Source: ").bold = True
            _add_hyperlink(p3, rec["url"], rec["url"])

        # Detail table
        table_data = rec.get("table_data", {})
        if table_data:
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            hdr[0].text = "Field"; hdr[1].text = "Value"
            for cell in hdr:
                for run in cell.paragraphs[0].runs:
                    run.bold = True
            for field, value in table_data.items():
                row = tbl.add_row().cells
                row[0].text = str(field); row[1].text = str(value)

        doc.add_paragraph("")

    doc.save(path)


def write_xlsx(path: Path, records: list[dict], query_lines: list[str]):
    """Write matched records to an .xlsx file."""
    if not _OPENPYXL_OK:
        raise RuntimeError("openpyxl not installed")

    wb = Workbook()
    ws = wb.active
    ws.title = "MyHeritage Results"

    # Header styling
    HDR_FILL  = PatternFill("solid", fgColor="2A4A7F")
    HDR_FONT  = Font(bold=True, color="FFFFFF", size=11)
    THIN_SIDE = Side(style="thin", color="B0B8C8")
    THIN      = Border(left=THIN_SIDE, right=THIN_SIDE,
                       top=THIN_SIDE, bottom=THIN_SIDE)

    # Collect all unique field names from table_data
    all_fields: list[str] = []
    for rec in records:
        for k in rec.get("table_data", {}):
            if k not in all_fields:
                all_fields.append(k)

    columns = ["#", "Full Name", "Category", "Match %", "URL"] + all_fields

    # Write column headers
    for col_i, col_name in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_i, value=col_name)
        cell.font   = HDR_FONT
        cell.fill   = HDR_FILL
        cell.border = THIN
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)

    # Write data rows
    for row_i, rec in enumerate(records, 2):
        td = rec.get("table_data", {})
        row_vals = (
            [row_i - 1,
             rec.get("full_name", ""),
             rec.get("category", ""),
             rec.get("score", ""),
             rec.get("url", "")]
            + [td.get(f, "") for f in all_fields]
        )
        for col_i, val in enumerate(row_vals, 1):
            cell = ws.cell(row=row_i, column=col_i, value=val)
            cell.border = THIN
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Auto-width (cap at 60)
    for col_i, col_name in enumerate(columns, 1):
        letter = get_column_letter(col_i)
        max_w  = max(len(str(col_name)),
                     *(len(str(ws.cell(row=r, column=col_i).value or ""))
                       for r in range(2, ws.max_row + 1)),
                     8)
        ws.column_dimensions[letter].width = min(max_w + 4, 60)

    wb.save(path)


# ═══════════════════════════════════════════════  SCRAPER CORE  ══════════════ #

async def _login(page, email: str, password: str, log) -> bool:
    """Try to log in.  Returns True on success."""
    log("  → Navigating to login page…")
    try:
        await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=30_000)
    except Exception as exc:
        log(f"  !! Could not load login page: {exc}")
        return False

    # Accept cookies if banner appears
    try:
        await page.click("text=Accept", timeout=4_000)
    except Exception:
        pass

    # Fill e-mail
    try:
        await page.fill('input[type="email"], input[name="email"], #email',
                        email, timeout=10_000)
    except Exception:
        try:
            await page.get_by_placeholder("Email").fill(email)
        except Exception as exc:
            log(f"  !! Could not find email field: {exc}")
            return False

    # Fill password
    try:
        await page.fill('input[type="password"], input[name="password"], #password',
                        password, timeout=10_000)
    except Exception:
        try:
            await page.get_by_placeholder("Password").fill(password)
        except Exception as exc:
            log(f"  !! Could not find password field: {exc}")
            return False

    # Click login button
    try:
        await page.click('button[type="submit"], button:has-text("Log in")',
                         timeout=10_000)
    except Exception:
        try:
            await page.keyboard.press("Enter")
        except Exception:
            pass

    try:
        await page.wait_for_load_state("domcontentloaded", timeout=20_000)
    except Exception:
        pass

    # Check if login succeeded (URL changes away from /login)
    if "login" not in page.url.lower():
        log("  ✓ Logged in.")
        return True

    log("  !! Login may have failed — check credentials or CAPTCHA.")
    return False


async def _search(page, first_name: str, surname: str,
                  record_filter: str, log) -> bool:
    """Navigate to search results page.  Returns True if results found."""
    log("  → Opening MyHeritage Research…")
    try:
        await page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=30_000)
    except Exception as exc:
        log(f"  !! Could not open research page: {exc}")
        return False

    await asyncio.sleep(1.5)

    # Fill first name (first + patronymic combined in one field)
    try:
        fn_input = page.locator(
            'input[placeholder*="first"], input[name*="first"], '
            'input[id*="first"], input[placeholder*="given"]'
        ).first
        await fn_input.fill(first_name, timeout=8_000)
    except Exception:
        log("  !! Could not find first-name field — trying generic approach.")
        try:
            inputs = await page.query_selector_all('input[type="text"]')
            if inputs:
                await inputs[0].fill(first_name)
        except Exception as exc:
            log(f"  !! {exc}")

    # Fill surname
    try:
        sn_input = page.locator(
            'input[placeholder*="last"], input[name*="last"], '
            'input[id*="last"], input[placeholder*="surname"]'
        ).first
        await sn_input.fill(surname, timeout=8_000)
    except Exception:
        log("  !! Could not find surname field — trying generic approach.")
        try:
            inputs = await page.query_selector_all('input[type="text"]')
            if len(inputs) > 1:
                await inputs[1].fill(surname)
        except Exception as exc:
            log(f"  !! {exc}")

    # Submit search
    try:
        await page.click(
            'button[type="submit"], button:has-text("Search"), '
            'input[type="submit"]',
            timeout=8_000,
        )
    except Exception:
        await page.keyboard.press("Enter")

    try:
        await page.wait_for_load_state("networkidle", timeout=25_000)
    except Exception:
        await asyncio.sleep(3)

    # Apply filter
    if record_filter != "All Records":
        log(f"  → Applying filter: {record_filter}")
        try:
            # MyHeritage typically has tab/filter buttons with the filter name
            await page.click(
                f'text="{record_filter}", button:has-text("{record_filter}")',
                timeout=6_000,
            )
            await asyncio.sleep(2)
        except Exception:
            # Try sidebar checkboxes / radio buttons
            try:
                await page.get_by_text(record_filter).first.click(timeout=4_000)
                await asyncio.sleep(2)
            except Exception:
                log(f"  !! Could not apply filter '{record_filter}' — continuing.")

    return True


async def _extract_score(result_el) -> float:
    """Try to extract the match % from a result card element."""
    try:
        # Look for a percentage text like "95%", "Match: 87%", etc.
        text = await result_el.text_content()
        m = re.search(r"(\d{1,3})\s*%", text or "")
        if m:
            return float(m.group(1))
    except Exception:
        pass
    return -1.0  # sentinel → caller will use name similarity


async def _scrape_detail(page, url: str, log) -> dict:
    """Open a result detail page and extract full name, category, table data."""
    data: dict = {"url": url, "full_name": "", "category": "", "table_data": {}}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=30_000)
        await asyncio.sleep(1)
    except Exception as exc:
        log(f"    !! Could not open detail page: {exc}")
        return data

    # Full name — try common heading selectors
    for sel in [
        "h1.person-name", "h1", ".record-title", ".full-name",
        ".name", '[class*="name"]',
    ]:
        try:
            el = page.locator(sel).first
            name = (await el.text_content(timeout=3_000) or "").strip()
            if name:
                data["full_name"] = name
                break
        except Exception:
            continue

    # Category — usually visible near the top of the detail page
    for sel in [
        ".record-type", ".category", ".collection-name",
        '[class*="category"]', '[class*="collection"]',
    ]:
        try:
            el = page.locator(sel).first
            cat = (await el.text_content(timeout=3_000) or "").strip()
            if cat:
                data["category"] = cat
                break
        except Exception:
            continue

    # Detail table — scrape all key-value pairs visible on the page
    table_data: dict = {}
    # Strategy A: <table> with label / value cells
    try:
        rows = await page.query_selector_all("table tr")
        for row in rows:
            cells = await row.query_selector_all("td, th")
            if len(cells) >= 2:
                key   = (await cells[0].text_content() or "").strip().rstrip(":")
                value = (await cells[1].text_content() or "").strip()
                if key and value:
                    table_data[key] = value
    except Exception:
        pass

    # Strategy B: definition lists (<dl>/<dt>/<dd>)
    if not table_data:
        try:
            dts = await page.query_selector_all("dt")
            dds = await page.query_selector_all("dd")
            for dt, dd in zip(dts, dds):
                key   = (await dt.text_content() or "").strip().rstrip(":")
                value = (await dd.text_content() or "").strip()
                if key and value:
                    table_data[key] = value
        except Exception:
            pass

    # Strategy C: label/value divs (common in modern SPAs)
    if not table_data:
        try:
            labels = await page.query_selector_all('[class*="label"], [class*="field-name"]')
            values = await page.query_selector_all('[class*="value"], [class*="field-value"]')
            for lbl, val in zip(labels, values):
                key   = (await lbl.text_content() or "").strip().rstrip(":")
                value = (await val.text_content() or "").strip()
                if key and value:
                    table_data[key] = value
        except Exception:
            pass

    data["table_data"] = table_data
    return data


async def run_scraper(
    *,
    first_name: str,
    surname: str,
    record_filter: str = "All Records",
    output_format: str = "both",
    output_folder: Path,
    email: str | None = None,
    password: str | None = None,
    log=print,
    progress=None,
    cancel_event=None,
) -> dict:
    """
    Main entry point called by the GUI worker thread.

    Parameters
    ----------
    first_name      : given name + patronymic (combined)
    surname         : family name
    record_filter   : "All Records" | "Historical Records" | "Family Trees"
    output_format   : "docx" | "xlsx" | "both"
    output_folder   : Path — directory to save results
    email / password: MyHeritage credentials
    log             : callable(str) for status messages
    progress        : callable(int, str) for progress bar (0-100, text)
    cancel_event    : threading.Event or None
    """

    def _progress(pct: int, text: str):
        log(text)
        if progress:
            progress(pct, text)

    def _cancelled() -> bool:
        return bool(cancel_event and cancel_event.is_set())

    want_docx = output_format in ("docx", "both")
    want_xlsx = output_format in ("xlsx", "both")
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    query_name  = " ".join(p for p in (first_name, surname) if p)
    query_lines = [
        f"First/patronymic name: {first_name or '—'}",
        f"Surname: {surname or '—'}",
        f"Filter: {record_filter}",
        f"Minimum match: {MIN_MATCH_PCT} %",
    ]

    summary_result: dict = {"ok": False}

    _progress(0, "Launching browser…")

    async with async_playwright() as pw:
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        context = await pw.chromium.launch_persistent_context(
            str(PROFILE_DIR),
            headless=False,
            args=["--start-maximized"],
            no_viewport=True,
        )
        page = context.pages[0] if context.pages else await context.new_page()

        try:
            # ── Login ─────────────────────────────────────────────────────── #
            if email and password:
                _progress(5, "Logging in…")
                await _login(page, email, password, log)
            else:
                log("  No credentials — proceeding as guest.")

            if _cancelled():
                return summary_result

            # ── Search ────────────────────────────────────────────────────── #
            _progress(15, "Running search…")
            ok = await _search(page, first_name, surname, record_filter, log)
            if not ok:
                summary_result["error"]   = "search_failed"
                summary_result["message"] = "Could not reach MyHeritage search results page."
                return summary_result

            if _cancelled():
                return summary_result

            _progress(30, "Collecting result links…")

            # Collect result cards / links
            await asyncio.sleep(2)
            result_links: list[dict] = []   # {url, score, name_text}

            # MyHeritage result items  (selectors may need tweaking on layout changes)
            card_selectors = [
                "a.results_result_link",
                'a[class*="result"]',
                ".result-item a",
                ".search-result a",
                'li[class*="result"] a',
            ]
            cards_found = False
            for sel in card_selectors:
                try:
                    els = await page.query_selector_all(sel)
                    if els:
                        cards_found = True
                        for el in els:
                            href  = await el.get_attribute("href") or ""
                            text  = (await el.text_content() or "").strip()
                            score = await _extract_score(el)
                            if href and href.startswith("http"):
                                result_links.append({
                                    "url":       href,
                                    "name_text": text,
                                    "score":     score,
                                })
                        break
                except Exception:
                    continue

            if not cards_found or not result_links:
                # Fallback: grab all links on the page that look like record URLs
                try:
                    all_links = await page.query_selector_all("a[href]")
                    for el in all_links:
                        href = await el.get_attribute("href") or ""
                        text = (await el.text_content() or "").strip()
                        if (
                            "myheritage.com" in href
                            and any(kw in href for kw in
                                    ["/record/", "/person/", "/family-site/", "/research/"])
                            and href not in [r["url"] for r in result_links]
                        ):
                            result_links.append({
                                "url":       href,
                                "name_text": text,
                                "score":     -1.0,
                            })
                except Exception as exc:
                    log(f"  !! Could not collect fallback links: {exc}")

            log(f"  Found {len(result_links)} candidate result(s).")

            # ── Filter by score / name similarity ─────────────────────────── #
            qualified: list[dict] = []
            for r in result_links:
                score = r["score"]
                if score < 0:
                    # No numeric score — use name similarity
                    score = _name_similarity(query_name, r["name_text"])
                if score >= MIN_MATCH_PCT:
                    r["score"] = round(score, 1)
                    qualified.append(r)

            log(f"  Qualified (≥ {MIN_MATCH_PCT} % match): {len(qualified)}")

            if not qualified:
                _progress(100, "No results above match threshold.")
                summary_result.update({
                    "ok":        True,
                    "docx_count": 0,
                    "xlsx_path": None,
                    "message":   f"No records found with match ≥ {MIN_MATCH_PCT} %.",
                })
                return summary_result

            # ── Open each detail page ─────────────────────────────────────── #
            records: list[dict] = []
            n = len(qualified)
            for i, r in enumerate(qualified, 1):
                if _cancelled():
                    log("Cancelled.")
                    break
                pct = 35 + int(50 * (i / n))
                _progress(pct, f"[{i}/{n}] Reading detail page…")
                detail_page = await context.new_page()
                try:
                    detail = await _scrape_detail(detail_page, r["url"], log)
                    detail["score"] = r["score"]
                    if not detail.get("full_name"):
                        detail["full_name"] = r["name_text"]
                    records.append(detail)
                    log(f"    ✓ {detail['full_name']} — {detail['score']} %")
                finally:
                    await detail_page.close()
                await asyncio.sleep(0.8)  # polite pause

            # ── Save output ───────────────────────────────────────────────── #
            _progress(88, "Saving output…")
            base   = safe_filename(f"myheritage_{query_name}") or "myheritage_results"
            docx_p = output_folder / f"{base}.docx"
            xlsx_p = output_folder / f"{base}.xlsx"

            saved_docx, saved_xlsx = False, False

            if want_docx and records:
                write_docx(docx_p, records, query_lines)
                saved_docx = True
                log(f"  → Saved Word: {docx_p.name}")

            if want_xlsx and records:
                write_xlsx(xlsx_p, records, query_lines)
                saved_xlsx = True
                log(f"  → Saved Excel: {xlsx_p.name}")

            _progress(100, f"Done — {len(records)} record(s) saved.")

            summary_result.update({
                "ok":          True,
                "docx_count":  1 if saved_docx else 0,
                "xlsx_path":   str(xlsx_p) if saved_xlsx else None,
                "output_folder": str(output_folder),
                "n_records":   len(records),
            })

        except Exception as exc:
            summary_result["error"]   = "exception"
            summary_result["message"] = f"{type(exc).__name__}: {exc}"
            log(f"  !! Unhandled exception: {exc}")

        finally:
            try:
                await context.close()
            except Exception:
                pass

    return summary_result
