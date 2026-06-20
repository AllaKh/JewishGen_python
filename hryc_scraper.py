"""hryc_scraper.py — search hryc.by (archives, newspapers and books).

Login is required (ASP.NET Core Identity). Search itself is a server-rendered GET
form, so this scraper needs NO browser — urllib + a cookie jar, like
pogroms_scraper / wikisource_scraper.

The site exposes 164 searchable "sources" (archives, gazettes, books) grouped in a
tree (config/hryc_sources.json). The GUI lets the user tick any of them; we replicate
the form's R.S[i].Chk / R.S[i].Id parameters.

For now results are plain lists → Word + Excel. Opening and saving the actual
documents needs a paid account and will be added later (the user will get one).
"""
import sys, re, json, time
import html as _html
from pathlib import Path
from urllib import request as _rq, parse as _up
from http.cookiejar import CookieJar, Cookie

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

BASE_URL   = "https://hryc.by"
LOGIN_URL  = BASE_URL + "/Identity/Account/Login"
SEARCH_URL = BASE_URL + "/search"
SITE_NAME  = "hryc.by"
SRC_FILE   = Path(__file__).resolve().parent / "config" / "hryc_sources.json"

HYPERLINK_REL = ("http://schemas.openxmlformats.org/"
                 "officeDocument/2006/relationships/hyperlink")

_HDRS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/141.0.0.0 Safari/537.36"),
    "Accept": ("text/html,application/xhtml+xml,application/xml;q=0.9,"
               "image/avif,image/webp,*/*;q=0.8"),
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
}


# ── helpers ─────────────────────────────────────────────────────────────────── #
def safe_fn(s: str, n: int = 100) -> str:
    return re.sub(r"\s+", " ",
                  re.sub(r'[\\/*?:"<>|]', "_", (s or "").strip()))[:n].strip() or "hryc"


def _strip_tags(html: str) -> str:
    txt = re.sub(r"<[^>]+>", " ", html or "")
    return re.sub(r"\s+", " ", _html.unescape(txt)).strip()


def _opener():
    """A urllib opener with its own cookie jar (keeps the login session). The site
    gates the session behind ASP.NET cookie-consent (`.AspNet.Consent=yes`, set by the
    «Принять» banner) — so we pre-seed that cookie («сначала прими куки, потом войти»),
    otherwise login is not remembered."""
    jar = CookieJar()
    jar.set_cookie(Cookie(
        version=0, name=".AspNet.Consent", value="yes",
        port=None, port_specified=False, domain=".hryc.by", domain_specified=True,
        domain_initial_dot=True, path="/", path_specified=True, secure=False,
        expires=None, discard=False, comment=None, comment_url=None,
        rest={"SameSite": "Lax"}, rfc2109=False))
    return _rq.build_opener(_rq.HTTPCookieProcessor(jar))


def _get(opener, url: str, timeout: int = 45) -> str:
    req = _rq.Request(url, headers=dict(_HDRS))
    return opener.open(req, timeout=timeout).read().decode("utf-8", "replace")


def _post(opener, url: str, data: dict, timeout: int = 45):
    body = _up.urlencode(data).encode()
    hdrs = dict(_HDRS)
    hdrs["Content-Type"] = "application/x-www-form-urlencoded"
    hdrs["Origin"] = BASE_URL
    hdrs["Referer"] = LOGIN_URL
    req = _rq.Request(url, data=body, headers=hdrs)
    return opener.open(req, timeout=timeout)


# ── source tree (config/hryc_sources.json) ─────────────────────────────────── #
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


# ── login ───────────────────────────────────────────────────────────────────── #
def login(opener, email: str, password: str, log) -> bool:
    """Accept cookies → GET the login page (antiforgery token + cookie) → POST creds.
    Success is detected by the POST redirecting AWAY from the login page (backup: a
    logout link on the home page)."""
    _get(opener, BASE_URL + "/")               # collect session/consent cookies first
    page = _get(opener, LOGIN_URL)
    m = re.search(r'name="__RequestVerificationToken"[^>]*value="([^"]+)"', page)
    token = m.group(1) if m else ""
    if not token:
        log("  !! антифорж-токен не найден на странице логина")
    data = {
        "Input.Email":    email,
        "Input.Password": password,
        "Input.RememberMe": "true",
        "__RequestVerificationToken": token,
    }
    try:
        resp = _post(opener, LOGIN_URL + "?returnUrl=%2F", data)
        final = (resp.geturl() or "").lower()
        body  = resp.read().decode("utf-8", "replace")
    except Exception as e:
        log(f"  !! POST логина не прошёл: {e}")
        return False
    # a failed login re-renders the login form (we stay on /Identity/Account/Login)
    if "identity/account/login" in final or 'name="Input.Password"' in body:
        if re.search(r'(field-validation-error|validation-summary-errors)[^>]*>\s*[^<\s]',
                     body):
            log("  !! сайт вернул ошибку входа (неверный email/пароль?)")
        else:
            log("  !! вход не подтверждён (остались на странице логина)")
        return False
    log("  ✓ Вход выполнен")
    return True


# ── search ──────────────────────────────────────────────────────────────────── #
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
    """Best-effort extraction of the result list. The logged-in results markup wasn't
    available offline, so this is deliberately generic AND the caller dumps the raw
    HTML — if rows come back empty/wrong, that dump tells us the real structure.

    Returns {rows, denied}: rows = [{source, text, url}], denied = [source names that
    need a higher (paid) access level]."""
    rows, denied, seen = [], [], set()

    # sources the account may not open (access-gated) — surfaced to the user
    for m in re.finditer(r'data-on="([^"]+)"[^>]*class="[^"]*search-source-check-denied',
                         html):
        denied.append(_html.unescape(m.group(1)))
    for m in re.finditer(r'class="[^"]*search-source-check-denied[^"]*"[^>]*data-on="([^"]+)"',
                         html):
        denied.append(_html.unescape(m.group(1)))

    # result anchors: links into a record / document / page, not chrome
    SKIP = re.compile(r'^/(Identity|lib|css|js|Home|Help|search|zakroma|favicon)', re.I)
    cur_src = ""
    # walk the document in order so we can attribute each link to the source heading
    for tok in re.finditer(
            r'<span class="search-source-name">(.*?)</span>'
            r'|<a\s+[^>]*href="([^"#]+)"[^>]*>(.*?)</a>', html, re.S):
        if tok.group(1) is not None:
            cur_src = _strip_tags(tok.group(1))
            continue
        href, inner = tok.group(2), tok.group(3)
        if not href or SKIP.match(href) or href.startswith(("http://", "https://")) \
                and "hryc.by" not in href:
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
    hdr = tbl.rows[0].cells
    for c, h in zip(hdr, ("#", "Источник", "Запись", "Ссылка")):
        c.text = ""
        run = c.paragraphs[0].add_run(h); run.bold = True
    for i, r in enumerate(rows, 1):
        cells = tbl.add_row().cells
        cells[0].text = str(i)
        cells[1].text = r.get("source", "") or "—"
        cells[2].text = r.get("text", "") or ""
        if r.get("url"):
            _add_link(cells[3].paragraphs[0], "Открыть", r["url"])
        else:
            cells[3].text = ""


def write_docx(path: Path, rows: list, qlines: list, append: bool = False):
    if not _DOCX_OK:
        return
    if append and path.exists():
        doc = Document(str(path))
        doc.add_page_break()
        doc.add_paragraph().add_run(
            f"➕ Добавлено ещё {len(rows)} записей").bold = True
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
        start = ws.max_row
        n0 = start - 1
    else:
        wb = Workbook(); ws = wb.active; ws.title = "hryc.by"
        ws.append(cols)
        for c in range(1, len(cols) + 1):
            ws.cell(row=1, column=c).font = Font(bold=True)
            ws.cell(row=1, column=c).fill = PatternFill(
                "solid", fgColor="DCE6F1")
        n0 = 0
    for i, r in enumerate(rows, 1):
        ws.append([n0 + i, SITE_NAME, r.get("source", ""),
                   r.get("text", ""), r.get("url", "")])
        if r.get("url"):
            cell = ws.cell(row=ws.max_row, column=5)
            cell.hyperlink = r["url"]; cell.value = "Открыть"
            cell.font = Font(color="0563C1", underline="single")
    widths = [6, 12, 26, 70, 14]
    for i, wd in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = wd
    ws.freeze_panes = "A2"
    wb.save(str(path))


# ── orchestrator ────────────────────────────────────────────────────────────── #
def run_scraper(
    *,
    email:         str       = "",
    password:      str       = "",
    query:         str       = "",
    sources:       list|None = None,    # list of source IDs to search
    no_stemming:   bool      = False,   # «Без стемминга»
    no_fuzziness:  bool      = False,   # «Без ошибок»
    show_experts:  bool      = False,   # «Эксперты»
    max_pages:     int       = 25,
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

    want_docx = output_format in ("docx", "both")
    want_xlsx = output_format in ("xlsx", "both")
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    qlines = [f"Запрос: {query}", f"Выбрано источников: {len(sources or [])}"]
    summary = {"ok": False}

    try:
        opener = _opener()

        # ── 1. LOGIN ─────────────────────────────────────────────────── #
        if email and password:
            _prog(5, "Вход на hryc.by…")
            if not login(opener, email, password, log):
                summary.update({"error": "login_failed",
                                "message": "Login failed — check email/password."})
                _prog(100, "Не удалось войти.")
                return summary
        else:
            log("  Без логина — большинство источников будут недоступны.")

        if cancel_event and cancel_event.is_set():
            return summary

        # ── 2. SEARCH (walk every page) ──────────────────────────────── #
        _prog(25, "Поиск…")
        rows, denied, seen = [], [], set()
        for page in range(1, max_pages + 1):
            if cancel_event and cancel_event.is_set():
                break
            url = build_search_url(query, sources or [], page,
                                   no_stemming=no_stemming, no_fuzziness=no_fuzziness,
                                   show_experts=show_experts)
            html = _get(opener, url)
            if page == 1:
                # raw dump so the (offline-unknown) result markup can be refined
                try:
                    (output_folder / "hryc_last_results.html").write_text(
                        html, encoding="utf-8")
                except Exception:
                    pass
            res = parse_results(html, log)
            if page == 1:
                denied = res["denied"]
            new = 0
            for r in res["rows"]:
                k = (r["source"], r["text"], r["url"])
                if k in seen:
                    continue
                seen.add(k); rows.append(r); new += 1
            _prog(25 + min(40, page * 5),
                  f"  стр.{page}: +{new} (всего {len(rows)})")
            if new == 0:                       # no new records → last page
                break
            time.sleep(0.4)
        if denied:
            log(f"  ⚠ Источников без доступа (нужен платный аккаунт): {len(denied)}")
        _prog(65, f"Найдено записей: {len(rows)}")

        if not rows:
            summary.update({"ok": True, "n_records": 0, "denied": len(denied),
                            "message": "Ничего не найдено (или нет доступа к источникам)."})
            _prog(100, "Ничего не найдено.")
            return summary

        # ── 3. SAVE ──────────────────────────────────────────────────── #
        _prog(85, "Сохранение файлов…")
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
    return summary


if __name__ == "__main__":
    run_scraper(query=sys.argv[1] if len(sys.argv) > 1 else "Шендерович",
                output_folder=Path("results"))
