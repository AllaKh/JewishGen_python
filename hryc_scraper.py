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
    from docx.shared import Mm, Pt
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
                     show_experts: bool = False) -> str:
    """Replicate the site form (verified against the user's live URLs): R.Q + every
    source's R.S[i].Chk / .Id (Chk=true only for ticked sources; indices must be
    contiguous, so ALL 164 are sent in order), plus R.Page, R.SearchInRecords and the
    three option flags (Без стемминга / Без ошибок / Эксперты)."""
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
    return SEARCH_URL + "?" + _up.urlencode(params)


def parse_results(html: str, log) -> dict:
    """Best-effort extraction of the result list. The logged-in result markup is still
    being confirmed, so this stays generic AND the caller dumps the raw HTML — that dump
    pins down the real structure.

    Returns {rows, denied}: rows = [{source, text, url}], denied = source names that need
    a higher (paid) access level."""
    rows, denied, seen = [], [], set()

    for m in re.finditer(
            r'data-on="([^"]+)"[^>]*class="[^"]*search-source-check-denied', html):
        denied.append(_html.unescape(m.group(1)))
    for m in re.finditer(
            r'class="[^"]*search-source-check-denied[^"]*"[^>]*data-on="([^"]+)"', html):
        denied.append(_html.unescape(m.group(1)))

    SKIP = re.compile(r'^/(Identity|lib|css|js|Home|Help|search|zakroma|favicon)', re.I)
    cur_src = ""
    for tok in re.finditer(
            r'<span class="search-source-name">(.*?)</span>'
            r'|<a\s+[^>]*href="([^"#]+)"[^>]*>(.*?)</a>', html, re.S):
        if tok.group(1) is not None:
            cur_src = _strip_tags(tok.group(1))
            continue
        href, inner = tok.group(2), tok.group(3)
        if not href or SKIP.match(href):
            continue
        if href.startswith(("http://", "https://")) and "hryc.by" not in href:
            continue
        text = _strip_tags(inner)
        if not text or len(text) < 2:
            continue
        url = href if href.startswith("http") else BASE_URL + href
        key = (cur_src, text, url)
        if key in seen:
            continue
        seen.add(key)
        rows.append({"source": cur_src, "text": text, "url": url})

    return {"rows": rows, "denied": sorted(set(denied))}


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


def _docx_table(doc, rows):
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Light Grid Accent 1"
    for c, h in zip(tbl.rows[0].cells, ("#", "Источник", "Запись", "Ссылка")):
        c.text = ""
        c.paragraphs[0].add_run(h).bold = True
    for i, r in enumerate(rows, 1):
        cells = tbl.add_row().cells
        cells[0].text = str(i)
        cells[1].text = r.get("source", "") or "—"
        cells[2].text = r.get("text", "") or ""
        if r.get("url"):
            _add_link(cells[3].paragraphs[0], "Открыть", r["url"])


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
    cols = ["#", "База", "Источник", "Запись", "URL"]
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
        ws.append([n0 + i, SITE_NAME, r.get("source", ""), r.get("text", ""), r.get("url", "")])
        if r.get("url"):
            cell = ws.cell(row=ws.max_row, column=5)
            cell.hyperlink = r["url"]; cell.value = "Открыть"
            cell.font = Font(color="0563C1", underline="single")
    for i, wd in enumerate([6, 12, 26, 70, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = wd
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
        await page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        await asyncio.sleep(2)
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
            _prog(20, "Поиск…")
            rows, denied, seen = [], [], set()
            for pno in range(1, max_pages + 1):
                if _cancelled():
                    break
                url = build_search_url(query, sources or [], pno,
                                       no_stemming=no_stemming, no_fuzziness=no_fuzziness,
                                       show_experts=show_experts)
                await page.goto(url, wait_until="domcontentloaded", timeout=40000)
                try:
                    await page.wait_for_load_state("networkidle", timeout=10000)
                except Exception:
                    await asyncio.sleep(1.5)
                html = await page.content()
                if pno == 1:
                    try:
                        (output_folder / "hryc_last_results.html").write_text(
                            html, encoding="utf-8")
                    except Exception:
                        pass
                res = parse_results(html, log)
                if pno == 1:
                    denied = res["denied"]
                new = 0
                for r in res["rows"]:
                    k = (r["source"], r["text"], r["url"])
                    if k in seen:
                        continue
                    seen.add(k); rows.append(r); new += 1
                _prog(20 + min(55, pno * 4), f"  стр.{pno}: +{new} (всего {len(rows)})")
                if new == 0:
                    break
                await asyncio.sleep(0.4)

            if denied:
                log(f"  ⚠ Источников без доступа (нужен платный аккаунт): {len(denied)}")
            _prog(80, f"Найдено записей: {len(rows)}")

            if not rows:
                summary.update({"ok": True, "n_records": 0, "denied": len(denied),
                                "message": "Ничего не найдено (или нет доступа)."})
                _prog(100, "Ничего не найдено.")
                return summary

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
            summary.update({"ok": True, "n_records": len(rows), "denied": len(denied),
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
