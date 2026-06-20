"""hryc_scraper.py — search hryc.by (archives, newspapers and books).

Runs a VISIBLE (headed) Chromium via Playwright so the user can watch it work
(login, search, paging) — like the other scrapers. A persistent profile
(.hryc_profile) keeps the login between runs.

Login is required for results; the site gates the session behind an ASP.NET
cookie-consent banner («Принять») which we click first. The search is the
server-rendered GET form; we navigate the full results URL (R.Q + all 164
R.S[i].Chk/.Id source toggles + R.Page + the option flags) and walk every page.

For now results are plain lists → Word + Excel. Opening/saving the actual
documents needs a paid account and will be added later.
"""
import sys, re, json, time, asyncio
import html as _html
from pathlib import Path
from urllib import parse as _up

try:
    from playwright.async_api import async_playwright
    _PW_OK = True
except ImportError:
    _PW_OK = False

try:
    from docx import Document
    from docx.shared import Mm, Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

BASE_URL    = "https://hryc.by"
LOGIN_URL   = BASE_URL + "/Identity/Account/Login"
SEARCH_URL  = BASE_URL + "/search"
SITE_NAME   = "hryc.by"
SRC_FILE    = Path(__file__).resolve().parent / "config" / "hryc_sources.json"
PROFILE_DIR = Path(__file__).resolve().parent / ".hryc_profile"

HYPERLINK_REL = ("http://schemas.openxmlformats.org/"
                 "officeDocument/2006/relationships/hyperlink")


# OCR snippets carry stray control characters that python-docx / openpyxl reject
# ("All strings must be XML compatible … no … control characters"). Strip everything
# illegal for XML — _CTRL_KEEP preserves our \x02/\x03 bold sentinels during parsing;
# _CTRL_ALL removes everything (used at write time, after sentinels are consumed).
_CTRL_KEEP = re.compile(r'[\x00\x01\x04-\x08\x0b\x0c\x0e-\x1f]')
_CTRL_ALL  = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _safe(s) -> str:
    return _CTRL_ALL.sub('', s or '')


# ── helpers ─────────────────────────────────────────────────────────────────── #
def safe_fn(s: str, n: int = 100) -> str:
    return re.sub(r"\s+", " ",
                  re.sub(r'[\\/*?:"<>|]', "_", (s or "").strip()))[:n].strip() or "hryc"


def _strip_tags(html: str) -> str:
    txt = re.sub(r"<[^>]+>", " ", html or "")
    return re.sub(r"\s+", " ", _html.unescape(txt)).strip()


def load_sources() -> list:
    """Flat list of every searchable source: {idx, id, label, en}, idx-sorted."""
    tree = json.loads(SRC_FILE.read_text(encoding="utf-8"))
    flat = []

    def walk(ns):
        for n in ns:
            if "idx" in n:
                flat.append(n)
            else:
                walk(n.get("children", []))

    walk(tree)
    return sorted(flat, key=lambda x: x["idx"])


def build_search_url(query: str, selected_ids, page: int = 1, *,
                     no_stemming: bool = False, no_fuzziness: bool = False,
                     show_experts: bool = False, fund: str = "", inventory: str = "",
                     record: str = "", doc_dates: str = "", added_since: str = "") -> str:
    """Replicate the site form (verified against the user's live URLs): R.Q + every
    source's R.S[i].Chk / .Id (Chk=true only for ticked sources; indices must be
    contiguous, so ALL 164 are sent in order), plus R.Page, R.SearchInRecords, the three
    option flags (Без стемминга / Без ошибок / Эксперты) and the archive-reference /
    date fields (R.Fund / R.Inventory / R.Record / R.DocDateRange / R.UpdateStartDate)."""
    sel = set(selected_ids or [])
    params = [("R.Q", query)]
    for s in load_sources():
        params.append((f"R.S[{s['idx']}].Chk", "true" if s["id"] in sel else "false"))
        params.append((f"R.S[{s['idx']}].Id", s["id"]))
    params += [
        ("R.Page", str(page)),
        ("R.SearchInRecords", "True"),
        ("R.NoStemming",  "true" if no_stemming  else "false"),
        ("R.NoFuzziness", "true" if no_fuzziness else "false"),
        ("R.ShowExperts", "true" if show_experts else "false"),
    ]
    # Mirror the real form submission EXACTLY (verified against the user's live URL):
    # these five fields are ALWAYS present, even when empty. R.UpdateStartDate is
    # data-val-required, so it must be sent — empty = the form's «0001-01-01» sentinel
    # (= no «added since» filter). An empty R.DocDateRange means no document-date filter.
    params += [
        ("R.Fund", fund),
        ("R.Inventory", inventory),
        ("R.Record", record),
        ("R.DocDateRange", doc_dates),
        ("R.UpdateStartDate", added_since or "0001-01-01"),
    ]
    return SEARCH_URL + "?" + _up.urlencode(params)


def parse_results(html: str, log) -> dict:
    """Each result on a hryc results page is a text-snippet block
    `<div style="margin: 1em"> …matched OCR text with <b>highlights</b>… <span>Для
    открытия документа …</span></div>` (≈20 per page; «Total: N» up the top, capped at
    1000; the document itself opens only on the paid «Эксперт» tariff). We keep the
    snippet text as the list row. Returns {rows, total}."""
    rows, seen = [], set()

    m = re.search(r'Total:.{0,60}?(\d[\d\s ,]*)', html, re.S)
    total = int(re.sub(r'\D', '', m.group(1))) if m else 0

    for bm in re.finditer(r'<div style="margin: ?1em">(.*?)</div>', html, re.S):
        inner = bm.group(1)
        # document link (paid account): <a href="/document?id=…">Title</a>
        doc_url, doc_title = "", ""
        dm = re.search(r'<a[^>]+href="(/document\?id=[^"]+)"[^>]*>(.*?)</a>', inner, re.S)
        if dm:
            doc_url = BASE_URL + _html.unescape(dm.group(1))
            doc_title = _strip_tags(dm.group(2))
        # drop the «pay to open» note and the document-link anchor from the snippet text
        inner = re.sub(r'<span[^>]*>\s*(Для открытия|To open|Каб адкрыць).*$', '',
                       inner, flags=re.S)
        inner = re.sub(r'<a[^>]+href="/document[^"]*"[^>]*>.*?</a>', '', inner, flags=re.S)
        inner = re.sub(r'<br\s*/?>', ' ', inner)
        # keep the site's <b>…</b> match highlights as sentinels (\x02…\x03) so Word
        # can render the searched word in bold; everything else → plain text
        inner = re.sub(r'<\s*b\s*>', '\x02', inner, flags=re.I)
        inner = re.sub(r'<\s*/\s*b\s*>', '\x03', inner, flags=re.I)
        inner = re.sub(r'<[^>]+>', ' ', inner)
        inner = re.sub(r'[ \t\r\n]+', ' ', _html.unescape(inner)).strip()
        inner = _CTRL_KEEP.sub('', inner)                           # drop XML-illegal ctrl chars
        text = inner.replace('\x02', '').replace('\x03', '')        # plain (Excel/dedup)
        if not text or len(text) < 4:
            continue
        key = text + "|" + doc_url
        if key in seen:
            continue
        seen.add(key)
        rows.append({"source": doc_title, "text": text, "rich": inner, "url": doc_url})

    return {"rows": rows, "total": total}


# ── Word ────────────────────────────────────────────────────────────────────── #
def _add_link(para, text, url):
    part = para.part
    r_id = part.relate_to(url, HYPERLINK_REL, is_external=True)
    hl = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), r_id)
    run = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color"); c.set(qn("w:val"), "0563C1"); rpr.append(c)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rpr.append(u)
    run.append(rpr)
    t = OxmlElement("w:t"); t.text = text or url; run.append(t)
    hl.append(run); para._p.append(hl)


# landscape A4 with 0.7" L/R margins → ~10.3" usable; keep the table inside it so the
# right margin survives (sum ≈ 9.8"). #, Source, File, Link narrow — Record wide.
# python-docx ignores table-level widths → set on EVERY cell per row.
_COLW = [Inches(0.4), Inches(1.4), Inches(5.6), Inches(1.7), Inches(0.7)]


def _set_widths(cells):
    for c, w in zip(cells, _COLW):
        c.width = w


def _add_rich(cell, rich):
    """Render the snippet into the cell, bolding the searched word(s) — the site marked
    them with <b>…</b>, kept here as \\x02…\\x03 sentinels."""
    para = cell.paragraphs[0]
    plain = lambda s: _safe(s.replace("\x02", "").replace("\x03", ""))
    i = 0
    for m in re.finditer(r'\x02(.*?)\x03', rich):
        if m.start() > i:
            para.add_run(plain(rich[i:m.start()])).font.size = Pt(8)
        para.add_run(_safe(m.group(1))).bold = True
        para.runs[-1].font.size = Pt(8)
        i = m.end()
    if i < len(rich):
        para.add_run(plain(rich[i:])).font.size = Pt(8)


def _docx_table(doc, rows):
    tbl = doc.add_table(rows=1, cols=5)
    tbl.style = "Light Grid Accent 1"
    tbl.autofit = False
    tbl.allow_autofit = False
    hdr = tbl.rows[0].cells
    for c, h in zip(hdr, ("#", "Источник", "Запись", "Файл", "Ссылка")):
        c.text = ""
        c.paragraphs[0].add_run(h).bold = True
    _set_widths(hdr)
    for i, r in enumerate(rows, 1):
        cells = tbl.add_row().cells
        cells[0].text = str(i)
        cells[1].text = r.get("source", "") or ""
        _add_rich(cells[2], r.get("rich") or r.get("text", "") or "")
        # saved document scan(s) — full path(s), one per line
        files = r.get("files") or []
        cells[3].paragraphs[0].add_run("\n".join(files)).font.size = Pt(8)
        if r.get("url"):
            _add_link(cells[4].paragraphs[0], "Открыть", r["url"])
        _set_widths(cells)


def write_docx(path: Path, rows: list, qlines: list, append: bool = False):
    if not _DOCX_OK:
        return
    if append and path.exists():
        doc = Document(str(path))
        doc.add_page_break()
        doc.add_paragraph().add_run(f"➕ Добавлено ещё {len(rows)} записей").bold = True
    else:
        doc = Document()
        for s in doc.sections:
            s.orientation = 1
            s.page_width, s.page_height = Mm(297), Mm(210)
            s.left_margin = s.right_margin = Inches(0.7)   # keep real left/right margins
        h = doc.add_heading("hryc.by — результаты поиска", level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for q in qlines:
            doc.add_paragraph(q)
        doc.add_paragraph(f"Найдено записей: {len(rows)}")
    _docx_table(doc, rows)
    doc.save(str(path))


# ── Excel ─────────────────────────────────────────────────────────────────── #
def write_xlsx(path: Path, rows: list, qlines: list, append: bool = False):
    if not _OPENPYXL_OK:
        return
    cols = ["#", "База", "Источник", "Запись", "Файл", "URL"]
    if append and path.exists():
        wb = load_workbook(str(path)); ws = wb.active
        n0 = ws.max_row - 1
    else:
        wb = Workbook(); ws = wb.active; ws.title = "hryc.by"
        ws.append(cols)
        for c in range(1, len(cols) + 1):
            ws.cell(row=1, column=c).font = Font(bold=True)
            ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor="DCE6F1")
        n0 = 0
    for i, r in enumerate(rows, 1):
        ws.append([n0 + i, SITE_NAME, _safe(r.get("source", "")),
                   _safe(r.get("text", "")), _safe("\n".join(r.get("files") or [])),
                   _safe(r.get("url", ""))])
        if r.get("url"):
            cell = ws.cell(row=ws.max_row, column=6)
            cell.hyperlink = r["url"]; cell.value = "Открыть"
            cell.font = Font(color="0563C1", underline="single")
    # #, Database, Source, File, URL narrow — Record wide
    for i, wd in enumerate([5, 10, 10, 90, 40, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = wd
    # wrap the long snippet + file columns
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=5):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    wb.save(str(path))


# ── browser: cookies / login ────────────────────────────────────────────────── #
async def _accept_cookies(page, log):
    """Click the «Принять» cookie-consent banner (#accept-cookies) so the session
    sticks — «сначала прими куки, потом войти»."""
    for sel in ('#accept-cookies', 'button:has-text("Принять")',
                'button:has-text("Accept")'):
        try:
            btn = page.locator(sel).first
            if await btn.count() and await btn.is_visible():
                await btn.click(timeout=3000)
                log("  ✓ Куки приняты")
                await asyncio.sleep(0.5)
                return
        except Exception:
            pass


async def _is_logged_in(page) -> bool:
    try:
        return await page.evaluate(
            """() => !!document.querySelector('a[href*="Account/Logout"], a[href*="Account/Manage"]')
                     || !document.querySelector('a[href="/Identity/Account/Login"]')""")
    except Exception:
        return False


async def _login(page, email, password, log) -> bool:
    await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=30000)
    await _accept_cookies(page, log)
    try:
        await page.fill("#Input_Email", email, timeout=8000)
        await page.fill("#Input_Password", password, timeout=8000)
        try:
            await page.check("#Input_RememberMe", timeout=2000)
        except Exception:
            pass
        await page.click('button[type="submit"]:has-text("Войти"), #login-submit, '
                         'form#account button[type="submit"]', timeout=8000)
    except Exception as e:
        log(f"  !! не удалось заполнить форму логина: {e}")
        return False
    try:
        await page.wait_for_load_state("domcontentloaded", timeout=10000)
    except Exception:
        pass
    ok = await _is_logged_in(page)
    log("  ✓ Вход выполнен" if ok else
        "  !! вход не подтверждён — проверь email/пароль (или войди вручную в окне)")
    return ok


def _clear_singleton_locks():
    for n in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
        try:
            (PROFILE_DIR / n).unlink()
        except Exception:
            pass


def _looks_image(b: bytes) -> bool:
    return (b[:3] == b"\xff\xd8\xff" or b[:8] == b"\x89PNG\r\n\x1a\n"
            or b[:6] in (b"GIF87a", b"GIF89a")
            or (b[:4] == b"RIFF" and b[8:12] == b"WEBP"))


async def _open_document(page, url, images_dir, base, dump, log):
    """Open a document page (/document?id=…) and save its scanned page image(s).
    The exact /document markup wasn't available offline, so this is best-effort + the
    FIRST document's HTML is dumped (results/hryc_document_sample.html) to refine later.
    Returns (saved_file_paths, page_text)."""
    saved, text = [], ""
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=40000)
    except Exception as e:
        log(f"    !! документ не открылся: {e}")
        return saved, text
    try:
        html = await page.content()
    except Exception:
        html = ""
    if dump is not False:
        try:
            (Path(dump) / "hryc_document_sample.html").write_text(html, encoding="utf-8")
        except Exception:
            pass
    text = _strip_tags(re.sub(r'<(script|style)[^>]*>.*?</\1>', ' ', html, flags=re.S))[:6000]
    # largest images on the page = the scanned document page(s)
    try:
        srcs = await page.evaluate(r"""() => {
            const out = [];
            for (const im of document.querySelectorAll('img')) {
                const s = im.currentSrc || im.src || '';
                const a = (im.naturalWidth || 0) * (im.naturalHeight || 0);
                if (s && a > 40000 && !/logo|icon|avatar|sprite|\.svg/i.test(s)) out.push([s, a]);
            }
            return out.sort((x, y) => y[1] - x[1]).map(z => z[0]).slice(0, 25);
        }""")
    except Exception:
        srcs = []
    for i, src in enumerate(srcs, 1):
        try:
            resp = await page.request.get(src, timeout=30000)
            b = await resp.body()
            if len(b) > 3000 and _looks_image(b):
                images_dir.mkdir(parents=True, exist_ok=True)
                ext = ".png" if b[:4] == b"\x89PNG" else (".webp" if b[8:12] == b"WEBP" else ".jpg")
                fp = images_dir / f"{base}_{i}{ext}"
                fp.write_bytes(b)
                saved.append(str(fp.resolve()))
        except Exception:
            continue
    return saved, text


# ── orchestrator ────────────────────────────────────────────────────────────── #
async def run_scraper(
    *,
    email:         str       = "",
    password:      str       = "",
    query:         str       = "",
    sources:       list|None = None,
    no_stemming:   bool      = False,
    no_fuzziness:  bool      = False,
    show_experts:  bool      = False,
    fund:          str       = "",
    inventory:     str       = "",
    record:        str       = "",
    doc_dates:     str       = "",
    added_since:   str       = "",
    open_documents: bool     = True,    # open & save the linked documents (paid account)
    max_docs:      int       = 60,
    max_pages:     int       = 50,
    output_format: str       = "both",
    output_folder            = Path("."),
    log                      = print,
    progress                 = None,
    cancel_event             = None,
    ask_file_conflict        = None,
) -> dict:

    def _prog(pct, txt):
        log(txt)
        if progress:
            progress(pct, txt)

    def _cancelled():
        return bool(cancel_event and cancel_event.is_set())

    want_docx = output_format in ("docx", "both")
    want_xlsx = output_format in ("xlsx", "both")
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    qlines = [f"Запрос: {query}", f"Выбрано источников: {len(sources or [])}"]
    summary = {"ok": False}

    if not _PW_OK:
        summary.update({"error": "playwright",
                        "message": "Playwright is not installed."})
        return summary

    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    _clear_singleton_locks()

    _prog(2, "Запускаю браузер…")
    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            str(PROFILE_DIR),
            headless=False, accept_downloads=True, no_viewport=True,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"])
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for extra in list(ctx.pages):
            if extra is not page:
                try: await extra.close()
                except Exception: pass

        try:
            # ── 1. cookies + login ───────────────────────────────────── #
            await page.goto(BASE_URL + "/", wait_until="domcontentloaded", timeout=30000)
            await _accept_cookies(page, log)
            if email and password and not await _is_logged_in(page):
                _prog(8, "Вход на hryc.by…")
                if not await _login(page, email, password, log):
                    # leave the window open a bit so the user can log in manually
                    for _ in range(60):
                        if _cancelled() or await _is_logged_in(page):
                            break
                        await asyncio.sleep(1)
            logged = await _is_logged_in(page)
            log("  → Статус: залогинен" if logged else
                "  → Статус: гость (большинство источников будут недоступны)")

            # ── 2. SEARCH (walk every page) ──────────────────────────── #
            # Results are server-rendered in the page (≈20 snippet blocks per page),
            # so we only need domcontentloaded — NO networkidle/sleep (that 10s wait
            # per page was the hang the user complained about).
            _prog(20, "Поиск…")
            rows, seen, total, pages = [], set(), 0, max_pages
            for pno in range(1, max_pages + 1):
                if _cancelled():
                    break
                url = build_search_url(query, sources or [], pno,
                                       no_stemming=no_stemming, no_fuzziness=no_fuzziness,
                                       show_experts=show_experts, fund=fund,
                                       inventory=inventory, record=record,
                                       doc_dates=doc_dates, added_since=added_since)
                await page.goto(url, wait_until="domcontentloaded", timeout=30000)
                html = await page.content()
                if pno == 1:
                    try:
                        (output_folder / "hryc_last_results.html").write_text(
                            html, encoding="utf-8")
                    except Exception:
                        pass
                res = parse_results(html, log)
                if pno == 1:
                    total = res.get("total", 0)
                    if total:
                        pages = min(max_pages, (total + 19) // 20)  # ≈20 per page
                new = 0
                for r in res["rows"]:
                    if r["text"] in seen:
                        continue
                    seen.add(r["text"]); rows.append(r); new += 1
                _prog(20 + min(60, pno * 60 // max(pages, 1)),
                      f"  стр.{pno}/{pages}: +{new} (всего {len(rows)})")
                if new == 0 or pno >= pages:
                    break

            _prog(85, f"Найдено (Total на сайте: {total}): собрано {len(rows)}")

            if not rows:
                summary.update({"ok": True, "n_records": 0, "total": total,
                                "message": "Ничего не найдено (или нужен платный доступ)."})
                _prog(100, "Ничего не найдено.")
                return summary

            # ── 2b. Open & save the linked documents (paid account) ─────── #
            doc_rows = [r for r in rows if r.get("url")]
            n_docs = 0
            if open_documents and doc_rows:
                images_dir = output_folder / "images" / safe_fn(query)
                todo = doc_rows[:max_docs]
                _prog(85, f"Открываю документы: {len(todo)} из {len(doc_rows)}…")
                for i, r in enumerate(todo, 1):
                    if _cancelled():
                        break
                    files, dtext = await _open_document(
                        page, r["url"], images_dir,
                        safe_fn(r.get("source") or f"{query}_{i}"),
                        dump=(output_folder if i == 1 else False), log=log)
                    if files:
                        r["files"] = files; n_docs += 1
                    if dtext and len(dtext) > len(r.get("text", "")):
                        r["text"] = dtext        # richer text from the opened document
                    _prog(85 + min(8, i * 8 // max(len(todo), 1)),
                          f"  документ {i}/{len(todo)}: {len(files)} скан(ов)")
                log(f"  → Открыто документов со сканами: {n_docs}")

            # ── 3. SAVE ──────────────────────────────────────────────── #
            _prog(88, "Сохранение файлов…")
            base   = safe_fn(f"hryc_{query}") or "hryc_results"
            docx_p = output_folder / f"{base}.docx"
            xlsx_p = output_folder / f"{base}.xlsx"
            existing = [p.name for p, want in ((docx_p, want_docx), (xlsx_p, want_xlsx))
                        if want and p.exists()]
            decision = "overwrite"
            if existing and ask_file_conflict:
                try:
                    decision = (ask_file_conflict(existing) or "overwrite").lower()
                except Exception:
                    decision = "overwrite"
                log(f"  → Файл(ы) существуют {existing} → {decision}")
            append = (decision == "append")

            sd = sx = False
            if decision == "skip":
                log("  → Сохранение пропущено.")
            else:
                if want_docx:
                    write_docx(docx_p, rows, qlines, append=append); sd = True
                    log(f"  Word: {docx_p}")
                if want_xlsx:
                    write_xlsx(xlsx_p, rows, qlines, append=append); sx = True
                    log(f"  Excel: {xlsx_p}")

            _prog(100, f"Готово — {len(rows)} записей.")
            summary.update({"ok": True, "n_records": len(rows), "total": total,
                            "documents": n_docs,
                            "docx_count": 1 if sd else 0,
                            "xlsx_path": str(xlsx_p) if sx else None,
                            "output_folder": str(output_folder)})
        except Exception as exc:
            summary.update({"error": "exception",
                            "message": f"{type(exc).__name__}: {exc}"})
            log(f"  !! {exc}")
        finally:
            try:
                await ctx.close()
            except Exception:
                pass
    return summary


if __name__ == "__main__":
    asyncio.run(run_scraper(query=sys.argv[1] if len(sys.argv) > 1 else "Иванов",
                            output_folder=Path("results")))
