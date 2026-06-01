#!/usr/bin/env python3
"""
familysearch_scraper.py
========================
АЛГОРИТМ (точно по требованию):

1.  Домашняя страница → First Names, Last Names, Place, Birth Year → SEARCH.
2.  Ждать 5 секунд результаты.
3.  Кликнуть таб Historical Records [data-testid="hr-tab"].
    URL меняется на ?tab=records — ждать ИМЕННО ЭТО, не networkidle.
4.  Кликнуть ПЕРВЫЙ результат → если редирект на логин → залогиниться ОДИН РАЗ:
      page.fill('#userName', email)
      page.fill('#password', password)
      page.click('#login')
5.  Все записи открываются на ОДНОЙ главной странице (не в новых вкладках) —
    сессия не теряется, логин не повторяется.
6.  На каждой странице записи: скопировать текст + превьюшку → Word.
7.  Кликнуть картинку → вьюер → download → JPG Only →
    [data-testid="full-text-confirm-download"] → сохранить JPG.
8.  Вернуться на результаты (go_back).
9.  Если НЕТ advanced: оставшиеся результаты ≥80%.
10. Если ЕСТЬ advanced: reload (обязательно!) → Advanced Search →
    [data-testid="advanced-search-form-button"] → все результаты ≥80%.
11. Имена файлов: {FirstNames} {LastNames}.docx / .xlsx
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

# ── Конфиг ────────────────────────────────────────────────────────────────── #
_HERE     = Path(__file__).resolve().parent
try:
    _CFG = json.loads((_HERE / "config" / "familysearch.json").read_text("utf-8"))
except Exception:
    _CFG = {}

HOME_URL      = _CFG.get("home_url", "https://www.familysearch.org/en/global")
FS_BASE       = "https://www.familysearch.org"
MIN_MATCH     = int(_CFG.get("min_match", 80))
BAD_PATHS     = ["/records/images", "/search/linker", "/linker",
                 "/en/tree/", "/tree/person/", "/tree/",
                 "/catalog", "/wiki", "/books", "/films"]
_dl           = _CFG.get("downloads_dir", "")
DOWNLOADS_DIR = Path(_dl) if _dl else Path.home() / "Downloads"
HYPERLINK_REL = ("http://schemas.openxmlformats.org/"
                 "officeDocument/2006/relationships/hyperlink")
_IMG_SKIP     = ("icon", "logo", "sprite", "avatar", "pixel", "placeholder",
                 ".svg", "fscdn.org")


# ── Утилиты ───────────────────────────────────────────────────────────────── #

def _sim(a: str, b: str) -> float:
    a = re.sub(r"\s+", " ", a.strip().lower())
    b = re.sub(r"\s+", " ", b.strip().lower())
    if not a or not b:
        return 0.0
    wa, wb = set(a.split()), set(b.split())
    return max(len(wa & wb) / max(len(wa), 1),
               difflib.SequenceMatcher(None, a, b).ratio()) * 100


def safe_fn(s: str, n: int = 100) -> str:
    return re.sub(r"\s+", " ",
                  re.sub(r'[\\/*?:"<>|]', "_", s.strip()))[:n].strip() or "document"


def _is_record(href: str) -> bool:
    if not href:
        return False
    for bad in BAD_PATHS:
        if bad in href:
            return False
    return "/ark:" in href


def _abs(href: str) -> str:
    return FS_BASE + href if href.startswith("/") else href


# ── Ввод в поле поиска ────────────────────────────────────────────────────── #

async def _type_field(page, sel: str, val: str, label: str, log) -> bool:
    """Кликнуть, очистить Ctrl+A/Del, печатать посимвольно."""
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
        await page.keyboard.type(val, delay=40)
        await asyncio.sleep(0.2)
        got = (await el.input_value(timeout=2000)).strip()
        if got:
            log(f"  OK  {label} = {got!r}")
            return True
        log(f"  !! {label}: поле осталось пустым")
    except Exception as e:
        log(f"  !! {label}: {e}")
    return False


# ── 1. ПОИСК ──────────────────────────────────────────────────────────────── #

async def _search(page, fn, ln, place, year, log):
    log(f"  Открываю {HOME_URL}")
    await page.goto(HOME_URL, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(3)

    # Ждём форму
    for sel in ['input[placeholder*="First and Middle" i]',
                'input[id*="givenName" i]']:
        try:
            await page.locator(sel).first.wait_for(state="visible", timeout=10000)
            break
        except Exception:
            continue

    await _type_field(page,
        'input[placeholder*="First and Middle" i], input[id*="givenName" i]',
        fn, "First Names", log)
    await _type_field(page,
        'input[placeholder*="Last or Maiden" i], input[id*="surname" i]',
        ln, "Last Names", log)
    if place:
        await _type_field(page,
            'input[placeholder*="City, County, State" i]',
            place, "Place Lived", log)
    if year:
        await _type_field(page,
            'input[placeholder*="Birth Year" i], input[id*="birthYear" i]',
            year, "Birth Year", log)

    for sel in ['button:has-text("SEARCH")', 'button:has-text("Search")',
                'button[type="submit"]']:
        try:
            el = page.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=6000)
                log("  SEARCH нажат")
                break
        except Exception:
            continue

    try:
        await page.wait_for_url(lambda u: "discovery/results" in u, timeout=20000)
    except Exception:
        pass
    await asyncio.sleep(5)  # обязательно 5 секунд
    log(f"  Результаты: {page.url}")


# ── 2. ЛОГИН ЧЕРЕЗ NAV-КНОПКУ (до HR tab) ────────────────────────────────── #

async def _sign_in_if_needed(page, email: str, password: str, log) -> bool:
    """
    Смотрим есть ли [data-testid='no-loggedin-sign-in-button'].
    Если есть — кликаем, логинимся, возвращаемся на results.
    Если нет — уже залогинены.
    Без логина HR tab показывает 0 строк!
    """
    log("  Проверяю авторизацию...")
    results_url = page.url

    try:
        btn = page.locator('[data-testid="no-loggedin-sign-in-button"]').first
        await btn.wait_for(state="visible", timeout=6000)
        log("  Не авторизован — кликаю Sign In...")
        await btn.click(timeout=5000)
        try:
            await page.wait_for_url(lambda u: "login" in u, timeout=10000)
        except Exception:
            pass
        await asyncio.sleep(2)
        ok = await _login(page, email, password, log)
        if ok and "discovery/results" not in page.url:
            log("  Возвращаюсь на страницу результатов...")
            await page.goto(results_url, wait_until="domcontentloaded", timeout=20000)
            await asyncio.sleep(3)
        return ok
    except Exception:
        log("  Уже авторизован (кнопка Sign In не найдена)")
        return True


# ── 3. ТАБ HISTORICAL RECORDS ─────────────────────────────────────────────── #

async def _click_hr(page, log):
    """
    Кликнуть [data-testid='hr-tab'].
    URL меняется на ?tab=records.
    После этого ждать появления строк таблицы (tbody tr) — до 10 секунд.
    """
    clicked = False
    try:
        el = page.locator('[data-testid="hr-tab"]').first
        if await el.count() and await el.is_visible():
            await el.click(timeout=5000)
            clicked = True
    except Exception:
        pass

    if not clicked:
        try:
            el = page.get_by_role("tab", name=re.compile(r"Historical Records", re.I)).first
            if await el.count():
                await el.click(timeout=5000)
                clicked = True
        except Exception:
            pass

    if not clicked:
        log("  (HR tab не найден)")
        return

    # Ждём URL с tab=records
    try:
        await page.wait_for_url(lambda u: "tab=records" in u, timeout=8000)
    except Exception:
        await asyncio.sleep(2)

    # Ждём СТРОКИ ТАБЛИЦЫ (таблица грузится через XHR после смены URL)
    try:
        await page.wait_for_selector("tbody tr", timeout=10000)
    except Exception:
        await asyncio.sleep(5)  # если не дождались — всё равно идём дальше

    log(f"  HR tab → {page.url}")


# ── 3. ЛОГИН — ОДИН РАЗ ──────────────────────────────────────────────────── #

async def _login(page, username: str, password: str, log) -> bool:
    """
    Заполнить #userName и #password через page.fill().
    Нажать #login. Ждать редиректа обратно на familysearch.org.
    Вызывать ТОЛЬКО ОДИН РАЗ за сессию.
    """
    log(f"  Страница логина: {page.url[:80]}")

    # Ждём поле #userName
    try:
        await page.locator("#userName").wait_for(state="visible", timeout=20000)
    except Exception:
        log("  !! #userName не появился")
        return False

    await asyncio.sleep(1)  # дать форме полностью отрисоваться

    # page.fill() — официальный метод Playwright для React-инпутов
    try:
        await page.fill("#userName", username)
        log(f"  #userName заполнен")
    except Exception as e:
        log(f"  !! fill #userName: {e}")
        return False

    await asyncio.sleep(0.3)

    try:
        await page.fill("#password", password)
        log(f"  #password заполнен")
    except Exception as e:
        log(f"  !! fill #password: {e}")
        return False

    await asyncio.sleep(0.3)

    # Проверить что поля НЕ пусты
    u_val = await page.locator("#userName").input_value()
    p_val = await page.locator("#password").input_value()
    log(f"  Проверка: user={'OK' if u_val else '!ПУСТО'}, "
        f"pass={'OK' if p_val else '!ПУСТО'}")

    if not u_val.strip() or not p_val.strip():
        log("  !! Поля пустые — логин провалится")
        return False

    # Нажать ТОЛЬКО #login — ничего другого
    try:
        await page.click("#login", timeout=5000)
        log("  #login нажат")
    except Exception as e:
        log(f"  !! #login: {e}")
        return False

    # Ждём редирект
    try:
        await page.wait_for_url(
            lambda u: "familysearch.org" in u and "login" not in u,
            timeout=30000)
    except Exception:
        await asyncio.sleep(6)

    ok = "familysearch.org" in page.url and "login" not in page.url
    log(f"  Логин {'OK ✓' if ok else 'ПРОВАЛИЛСЯ ✗'}  URL: {page.url[:80]}")
    return ok


# ── 4. СОБРАТЬ СТРОКИ ТАБЛИЦЫ ────────────────────────────────────────────── #

async def _collect(page, qname: str, log) -> list:
    """
    Структура таблицы (подтверждена):
      cells[0] = ссылка "More"  href=/ark:... ← URL записи, НЕ имя!
      cells[1] = картинка
      cells[2] = <strong>Имя</strong> + коллекция 2-й строкой
      cells[3] = Events
      cells[4] = Relationships
    """
    await asyncio.sleep(2)
    results, seen = [], set()
    rows = await page.query_selector_all("tbody tr")
    log(f"  Строк: {len(rows)}")

    for row in rows:
        try:
            cells = await row.query_selector_all("td")
            if not cells:
                continue
            # URL из ark-ссылки в cells[0]
            url = ""
            for a in await row.query_selector_all("a[href]"):
                h = (await a.get_attribute("href") or "").strip()
                if _is_record(h):
                    url = _abs(h)
                    break
            if not url or url in seen:
                continue
            seen.add(url)

            # Имя из cells[2] <strong>
            idx  = 2 if len(cells) > 3 else max(0, len(cells) - 3)
            name = ""
            coll = ""
            if len(cells) > idx:
                nc = cells[idx]
                try:
                    b = await nc.query_selector("strong, b")
                    if b:
                        name = (await b.text_content() or "").strip()
                except Exception:
                    pass
                if not name:
                    lines = [ln.strip()
                             for ln in (await nc.text_content() or "").splitlines()
                             if ln.strip()]
                    name = lines[0] if lines else ""
                    coll = lines[1] if len(lines) > 1 else ""
                else:
                    try:
                        lines = [ln.strip()
                                 for ln in (await nc.text_content() or "").splitlines()
                                 if ln.strip()]
                        coll = lines[1] if len(lines) > 1 else ""
                    except Exception:
                        pass

            if not name:
                continue
            evts = (await cells[idx+1].text_content() or "").strip() if len(cells)>idx+1 else ""
            rels = (await cells[idx+2].text_content() or "").strip() if len(cells)>idx+2 else ""
            score = round(_sim(qname, name), 1)
            results.append({"url": url, "name": name, "coll": coll,
                            "evts": evts, "rels": rels, "score": score})
            log(f"    {score:5.1f}%  {name}")
        except Exception:
            continue

    log(f"  Кандидатов: {len(results)}")
    return results


# ── 5. УСТАНОВИТЬ 60 НА СТРАНИЦУ ─────────────────────────────────────────── #

async def _set_60(page, log):
    for sel in ['select[aria-label*="result" i]',
                'select[name*="result" i]',
                'select[id*="result" i]']:
        try:
            el = page.locator(sel).last
            if not await el.count():
                continue
            opts = await el.evaluate("e => Array.from(e.options).map(o=>o.value)")
            if "60" in opts:
                await el.select_option(value="60")
                await asyncio.sleep(1)
                log("  60 результатов на странице")
                return
        except Exception:
            continue


# ── 6. РАСШИРЕННЫЙ ПОИСК ─────────────────────────────────────────────────── #

async def _advanced(page, adv: dict, log):
    log("  Открываю Advanced Search...")
    try:
        btn = page.locator('[data-testid="advanced-search-form-button"]').first
        await btn.wait_for(state="visible", timeout=5000)
        await btn.click(timeout=5000)
        await asyncio.sleep(1.5)
        log("  Модальное окно открыто")
    except Exception as e:
        log(f"  !! Advanced Search: {e}")
        return

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
    opened: set = set()
    for key, (tab, sel) in event_map.items():
        val = adv.get(key, "")
        if not val:
            continue
        if tab not in opened:
            try:
                t = page.get_by_text(re.compile(rf"^{re.escape(tab)}$", re.I)).first
                if await t.count():
                    await t.click(timeout=3000)
                    await asyncio.sleep(0.7)
                    opened.add(tab)
            except Exception:
                pass
        await _type_field(page, sel, val, key, log)

    fam_map = {
        "spouse": ("SPOUSE",       "Spouse's First Names",       "Spouse's Last Names"),
        "father": ("FATHER",       "Father's First Names",       "Father's Last Names"),
        "mother": ("MOTHER",       "Mother's First Names",       "Mother's Last Names"),
        "other":  ("OTHER PERSON", "Other Person's First Names", "Other Person's Last Names"),
    }
    for key, (tab, fp, lp) in fam_map.items():
        fv = adv.get(f"{key}_first", "")
        lv = adv.get(f"{key}_last",  "")
        if not fv and not lv:
            continue
        try:
            t = page.get_by_text(re.compile(rf"^{re.escape(tab)}$", re.I)).first
            if await t.count():
                await t.click(timeout=3000)
                await asyncio.sleep(0.7)
        except Exception:
            pass
        if fv:
            await _type_field(page, f'input[placeholder="{fp}"]', fv, f"{key}_first", log)
        if lv:
            await _type_field(page, f'input[placeholder="{lp}"]', lv, f"{key}_last", log)

    if adv.get("country"):
        try:
            t = page.get_by_text(re.compile(r"^LOCATION$", re.I)).first
            if await t.count():
                await t.click(timeout=3000)
                await asyncio.sleep(0.5)
        except Exception:
            pass
        await _type_field(page, 'input[placeholder="Country or Location"]',
                         adv["country"], "country", log)
    if adv.get("state"):
        await _type_field(page, 'input[placeholder="State or Province"]',
                         adv["state"], "state", log)
    if adv.get("keywords"):
        await _type_field(page, 'input[placeholder*="keyword" i]',
                         adv["keywords"], "keywords", log)

    for sel in ['button:has-text("SEARCH")', 'button:has-text("Search")']:
        try:
            el = page.locator(sel).last
            if await el.count() and await el.is_visible():
                await el.click(timeout=5000)
                break
        except Exception:
            continue

    try:
        await page.wait_for_url(lambda u: "tab=records" in u, timeout=15000)
    except Exception:
        await asyncio.sleep(3)
    log("  Advanced Search отправлен")


# ── 7. СКАЧАТЬ БАЙТЫ КАРТИНКИ ─────────────────────────────────────────────── #

async def _fetch_bytes(ctx, src: str) -> bytes | None:
    if not src or not src.startswith("http"):
        return None
    pg = await ctx.new_page()
    try:
        r = await pg.goto(src, timeout=15000)
        if r and r.ok:
            body = await r.body()
            if len(body) > 5000:
                return body
    except Exception:
        pass
    finally:
        await pg.close()
    return None


# ── 8. НАЙТИ ЛУЧШУЮ КАРТИНКУ НА СТРАНИЦЕ ─────────────────────────────────── #

async def _best_img(page) -> str:
    best, area = "", 0
    for el in await page.query_selector_all("img[src]"):
        try:
            src = (await el.get_attribute("src") or "").strip()
            if not src.startswith("http"):
                continue
            if any(b.lower() in src.lower() for b in _IMG_SKIP):
                continue
            w = int(await el.evaluate("e => e.naturalWidth")  or 0)
            h = int(await el.evaluate("e => e.naturalHeight") or 0)
            if w * h > area:
                area = w * h
                best = src
        except Exception:
            continue
    return best


# ── 9. СКАЧАТЬ ПОЛНОФОРМАТНУЮ JPG ЧЕРЕЗ ВЬЮЕР ────────────────────────────── #

async def _download_jpg(ctx, page, dest_dir: Path, title: str, log) -> str | None:
    """
    Кликнуть картинку → вьюер → download → JPG Only →
    [data-testid="full-text-confirm-download"] → сохранить файл.
    """
    dest_dir.mkdir(parents=True, exist_ok=True)
    fname  = safe_fn(title) + ".jpg"
    dest   = dest_dir / fname
    before = set(DOWNLOADS_DIR.glob("*.jpg")) | set(DOWNLOADS_DIR.glob("*.jpeg"))
    tabs_before = set(ctx.pages)

    # Найти и кликнуть лучшую картинку
    img_src = await _best_img(page)
    clicked = False
    if img_src:
        try:
            el = page.locator(f'img[src="{img_src}"]').first
            if await el.count() and await el.is_visible():
                await el.click(timeout=5000)
                await asyncio.sleep(3)
                clicked = True
                log("    Картинка кликнута")
        except Exception:
            pass

    if not clicked:
        for sel in ['img[src*="dz/v1"]', 'img[src*="apiv2"]',
                    'img[alt="Thumbnail"]', 'main img']:
            try:
                el = page.locator(sel).first
                if not await el.count():
                    continue
                src = await el.get_attribute("src") or ""
                if any(b.lower() in src.lower() for b in _IMG_SKIP):
                    continue
                if await el.is_visible():
                    await el.click(timeout=5000)
                    await asyncio.sleep(3)
                    img_src = src
                    clicked = True
                    log(f"    Картинка кликнута ({sel})")
                    break
            except Exception:
                continue

    if not clicked:
        log("    Картинка на странице не найдена")
        return None

    # Вьюер может открыться в новой вкладке
    viewer = page
    await asyncio.sleep(1)
    new_tabs = set(ctx.pages) - tabs_before
    if new_tabs:
        viewer = list(new_tabs)[0]
        try:
            await viewer.wait_for_load_state("domcontentloaded", timeout=10000)
        except Exception:
            pass
        await asyncio.sleep(2)
        log("    Вьюер в новой вкладке")

    # Весь блок download обёрнут в expect_download чтобы не пропустить событие
    downloaded = None
    try:
        async with viewer.expect_download(timeout=45000) as dl_info:
            # Кнопка download (стрелка вниз)
            dl_ok = False
            for sel in [
                'button[aria-label*="Download" i]',
                'button[title*="Download" i]',
                '[data-testid*="download" i]:not([data-testid="full-text-confirm-download"])',
                '[class*="toolbar"] button:nth-last-child(3)',
                '[class*="toolbar"] button:nth-last-child(2)',
                '[class*="tools"] button:nth-last-child(2)',
            ]:
                try:
                    el = viewer.locator(sel).first
                    if await el.count() and await el.is_visible():
                        await el.click(timeout=4000)
                        await asyncio.sleep(1.5)
                        dl_ok = True
                        log(f"    Download нажат ({sel})")
                        break
                except Exception:
                    continue

            if not dl_ok:
                raise RuntimeError("кнопка download не найдена")

            # JPG Only
            for lbl in ("JPG Only", "JPG only"):
                try:
                    el = viewer.get_by_text(lbl, exact=True).first
                    if await el.count():
                        await el.click(timeout=3000)
                        await asyncio.sleep(0.5)
                        log("    JPG Only выбран")
                        break
                except Exception:
                    continue

            # Confirm download
            for sel in ['[data-testid="full-text-confirm-download"]',
                        'button:has-text("Download")',
                        'button:has-text("DOWNLOAD")']:
                try:
                    btn = viewer.locator(sel).first
                    if await btn.count() and await btn.is_visible():
                        await btn.click(timeout=5000)
                        log(f"    Download нажат ({sel})")
                        break
                except Exception:
                    continue

        dl = await dl_info.value
        await dl.save_as(str(dest))
        downloaded = str(dest)
        log(f"    Сохранено: {fname} ({dest.stat().st_size//1024}KB)")
    except Exception as exc:
        log(f"    expect_download: {exc} → жду в Downloads...")
        for _ in range(15):
            await asyncio.sleep(1)
            after = set(DOWNLOADS_DIR.glob("*.jpg")) | set(DOWNLOADS_DIR.glob("*.jpeg"))
            nw = after - before
            if nw:
                src_f = max(nw, key=lambda p: p.stat().st_mtime)
                shutil.move(str(src_f), str(dest))
                downloaded = str(dest)
                log(f"    Перемещено: {fname}")
                break

    if not downloaded and not img_src:
        # Кнопка не найдена, нет картинки
        if viewer is not page:
            try: await viewer.close()
            except Exception: pass
        return None

    # Закрыть лишние вкладки
    await asyncio.sleep(0.5)
    for pg in list(ctx.pages):
        if pg not in (page, viewer) and "familysearch" not in pg.url:
            try: await pg.close()
            except Exception: pass
    if viewer is not page:
        try: await viewer.close()
        except Exception: pass

    if downloaded:
        return downloaded
    # Fallback: сохранить превьюшку
    if img_src:
        body = await _fetch_bytes(ctx, img_src)
        if body:
            fp = dest_dir / (safe_fn(title) + "_preview.jpg")
            fp.write_bytes(body)
            log(f"    Fallback превьюшка: {fp.name}")
            return str(fp)
    return None


# ── 10. СКРАПИНГ СТРАНИЦЫ ЗАПИСИ ─────────────────────────────────────────── #

async def _scrape_page(ctx, page, url: str, name_hint: str,
                       images_root: Path, logged_in_ref: list,
                       email: str, password: str, log) -> dict:
    """
    Навигировать main page на url.
    Если редирект на логин — войти (только если logged_in_ref[0] == False).
    Скрапить данные, скачать картинку.
    НЕ закрывает page — вызывающий делает go_back().
    """
    rec = {"url": url, "title": name_hint, "name": name_hint,
           "table_data": {}, "images": [], "thumb_bytes": None,
           "events": "", "relationships": "", "collection": ""}

    for bad in BAD_PATHS:
        if bad in url:
            return rec

    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(3)

    # Если редирект на логин — залогиниться ОДИН РАЗ
    if "login" in page.url:
        if logged_in_ref[0]:
            log(f"  !! Повторный логин-редирект (уже был вход). URL: {page.url[:60]}")
            return rec
        if not email or not password:
            log("  !! Нет credentials для логина")
            return rec
        log("  → Форма логина, входим...")
        ok = await _login(page, email, password, log)
        if not ok:
            log("  !! Вход провалился")
            return rec
        logged_in_ref[0] = True
        # После логина state= редиректит на нужную запись,
        # но если нет — навигируем вручную
        if url.split("?")[0] not in page.url:
            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            await asyncio.sleep(3)

    # Заголовок
    for sel in ["h1", '[class*="title" i]', "h2"]:
        try:
            t = (await page.locator(sel).first.text_content(timeout=3000) or "").strip()
            if 3 < len(t) < 300:
                rec["title"] = t
                break
        except Exception:
            pass

    # Данные: dl/dt/dd
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

    # Метка для имени файла
    img_label = name_hint
    if td:
        parts = [name_hint]
        for k in ("Event Type", "Type"):
            if td.get(k):
                parts.append(td[k]); break
        for k in ("Event Date", "Date", "Birth Date", "Death Date", "Marriage Date"):
            if td.get(k):
                parts.append(td[k]); break
        if len(parts) > 1:
            img_label = " — ".join(parts)

    img_dir = images_root / safe_fn(img_label)

    # Превьюшка (маленькая) для Word
    img_src = await _best_img(page)
    if img_src:
        rec["thumb_bytes"] = await _fetch_bytes(ctx, img_src)
        if rec["thumb_bytes"]:
            log(f"    Превьюшка: {len(rec['thumb_bytes'])//1024}KB")
        else:
            log("    Превьюшка: не получена")
    else:
        log("    На странице нет картинки документа")

    # Полноформатная JPG
    if img_src:
        jp = await _download_jpg(ctx, page, img_dir, img_label, log)
        rec["images"] = [jp] if jp else []
    else:
        rec["images"] = []

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


def write_docx(path: Path, records: list, qlines: list):
    if not _DOCX_OK:
        raise RuntimeError("python-docx не установлен")
    doc = Document()
    s = doc.sections[0]
    s.page_width  = Mm(297); s.page_height = Mm(210)
    s.left_margin = s.right_margin = Mm(15)
    s.top_margin  = s.bottom_margin = Mm(15)

    h = doc.add_heading("FamilySearch — Результаты", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Параметры:")
    for ln in qlines:
        doc.add_paragraph(ln, style="List Bullet")
    doc.add_paragraph(f"Найдено: {len(records)} (совпадение ≥{MIN_MATCH}%)")
    doc.add_paragraph("")

    for i, rec in enumerate(records, 1):
        title = rec.get("title") or rec.get("name", "—")
        doc.add_heading(f"{i}. {title}", level=2)

        if rec.get("url"):
            pp = doc.add_paragraph()
            pp.add_run("Источник: ").bold = True
            _add_link(pp, rec["url"], rec["url"])

        p = doc.add_paragraph()
        p.add_run("Совпадение: ").bold = True
        p.add_run(f"{rec.get('score','?')}%")

        rows_data = []
        if rec.get("collection"):
            rows_data.append(("Коллекция", rec["collection"]))
        if rec.get("events"):
            rows_data.append(("События", rec["events"]))
        if rec.get("relationships"):
            rows_data.append(("Родственники", rec["relationships"]))
        for f, v in rec.get("table_data", {}).items():
            rows_data.append((str(f), str(v)))

        if rows_data:
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            hdr[0].text = "Поле"; hdr[1].text = "Значение"
            for cell in hdr:
                for run in cell.paragraphs[0].runs:
                    run.bold = True
            for f, v in rows_data:
                r = tbl.add_row().cells
                r[0].text = f; r[1].text = v

        doc.add_paragraph("")

        imgs = rec.get("images", [])
        tb   = rec.get("thumb_bytes")
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
        else:
            doc.add_paragraph("  [изображение недоступно]")
        doc.add_paragraph("")
    doc.save(path)


# ── Excel ─────────────────────────────────────────────────────────────────── #

def write_xlsx(path: Path, records: list, qlines: list):
    if not _OPENPYXL_OK:
        raise RuntimeError("openpyxl не установлен")
    wb = Workbook(); ws = wb.active; ws.title = "FamilySearch"
    HF = PatternFill("solid", fgColor="006B6B")
    HN = Font(bold=True, color="FFFFFF", size=11)
    TS = Side(style="thin", color="B0C8C8")
    T  = Border(left=TS, right=TS, top=TS, bottom=TS)

    aff: list = []
    for rec in records:
        for k in rec.get("table_data", {}):
            if k not in aff: aff.append(k)

    cols = ["#", "Имя", "Совп. %", "Коллекция", "События",
            "Родственники", "Файл JPG", "URL"] + aff
    for ci, cn in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=cn)
        c.font = HN; c.fill = HF; c.border = T
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for ri, rec in enumerate(records, 2):
        td   = rec.get("table_data", {})
        imgs = "\n".join(Path(p).name for p in rec.get("images", []))
        vals = [ri-1, rec.get("title", rec.get("name","")),
                rec.get("score",""), rec.get("collection",""),
                rec.get("events",""), rec.get("relationships",""),
                imgs, rec.get("url","")] + [td.get(f,"") for f in aff]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = T
            c.alignment = Alignment(wrap_text=True, vertical="top")

    for ci in range(1, len(cols)+1):
        ltr = get_column_letter(ci)
        mw  = max(len(str(cols[ci-1])),
                  *(len(str(ws.cell(row=r, column=ci).value or "").split("\n")[0])
                    for r in range(2, ws.max_row+1)), 8)
        ws.column_dimensions[ltr].width = min(mw+4, 60)
    wb.save(path)


# ── ГЛАВНАЯ ФУНКЦИЯ ───────────────────────────────────────────────────────── #

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
        if progress: progress(pct, txt)

    def _done():
        return bool(cancel_event and cancel_event.is_set())

    adv       = advanced or {}
    has_adv   = any(v for v in adv.values() if v and v not in (False, "Unspecified"))
    want_docx = output_format in ("docx", "both")
    want_xlsx = output_format in ("xlsx", "both")
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    images_root = output_folder / "images"

    qname     = " ".join(p for p in (first_names, last_names) if p)
    qlines    = [ln for ln in [
        f"First Names: {first_names}", f"Last Names: {last_names}",
        f"Place Lived: {place_lived}", f"Birth Year: {birth_year}",
    ] if not ln.endswith(": ")]
    summary   = {"ok": False}
    file_base = safe_fn(qname) if qname else "familysearch_results"

    # logged_in_ref[0] == True после первого успешного входа
    # Передаём как список чтобы _scrape_page мог изменить флаг
    logged_in_ref = [False]

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
            # ── 1. ПОИСК ──────────────────────────────────────────────── #
            _prog(5, "Поиск...")
            await _search(page, first_names, last_names,
                          place_lived, birth_year, log)
            if _done(): return summary

            # ── 2. ТАБ HISTORICAL RECORDS ─────────────────────────────── #
            # HR tab работает без логина (ограниченные результаты).
            # Логин произойдёт автоматически при открытии первой записи.
            _prog(12, "Historical Records tab...")
            await _click_hr(page, log)
            if _done(): return summary

            # ── 4. 60 НА СТРАНИЦУ + СБОР РЕЗУЛЬТАТОВ ─────────────────── #
            _prog(20, "Сбор результатов...")
            await _set_60(page, log)
            raw       = await _collect(page, qname, log)
            qualified = [r for r in raw if r["score"] >= MIN_MATCH]
            log(f"  Подходящих (≥{MIN_MATCH}%): {len(qualified)}")

            if not qualified:
                _prog(100, f"Нет записей с совпадением ≥{MIN_MATCH}%.")
                summary.update({"ok": True, "n_records": 0,
                                "message": f"Нет записей ≥{MIN_MATCH}%."})
                return summary

            # ── 4-5. СКРАПИНГ ЗАПИСЕЙ на ОДНОЙ главной странице ──────── #
            # Все записи открываются на page — сессия гарантированно одна.
            results_url = page.url
            records: list = []

            def _all_to_scrape():
                return qualified  # будет переопределено после advanced

            to_scrape = list(qualified)

            for i, r in enumerate(to_scrape, 1):
                if _done(): break
                _prog(20 + int(55 * i / len(to_scrape)),
                      f"[{i}/{len(to_scrape)}] {r['name'][:60]}...")

                det = await _scrape_page(ctx, page, r["url"], r["name"],
                                         images_root, logged_in_ref,
                                         email or "", password or "", log)
                det["score"]         = r["score"]
                det["collection"]    = r.get("coll", "")
                det["events"]        = r.get("evts", "")
                det["relationships"] = r.get("rels", "")
                records.append(det)
                log(f"  ✓  {det['title'][:70]}  ({r['score']}%)")

                # Вернуться на results_url напрямую (go_back ненадёжен после login redirect chain)
                await page.goto(results_url,
                                wait_until="domcontentloaded", timeout=20000)
                try:
                    await page.wait_for_selector("tbody tr", timeout=8000)
                except Exception:
                    await asyncio.sleep(3)

                # После первой записи: если есть advanced search
                if i == 1 and has_adv:
                    _prog(78, "Advanced Search: обязательный reload...")
                    await page.reload(wait_until="domcontentloaded", timeout=20000)
                    await asyncio.sleep(3)
                    await _click_hr(page, log)
                    await asyncio.sleep(1)
                    _prog(82, "Advanced Search...")
                    await _advanced(page, adv, log)
                    await asyncio.sleep(2)
                    await _set_60(page, log)
                    _prog(85, "Сбор результатов после Advanced Search...")
                    raw_adv   = await _collect(page, qname, log)
                    qualified = [r for r in raw_adv if r["score"] >= MIN_MATCH]
                    log(f"  После Advanced: {len(qualified)}")
                    results_url = page.url
                    # Перезапустить цикл по всем результатам
                    to_scrape = list(qualified)
                    records   = []
                    break  # выйти из первого цикла, войти во второй

            # Если был advanced — второй проход по всем результатам
            if has_adv and to_scrape:
                for i, r in enumerate(to_scrape, 1):
                    if _done(): break
                    _prog(85 + int(10 * i / len(to_scrape)),
                          f"[{i}/{len(to_scrape)}] {r['name'][:60]}...")
                    det = await _scrape_page(ctx, page, r["url"], r["name"],
                                             images_root, logged_in_ref,
                                             email or "", password or "", log)
                    det["score"]         = r["score"]
                    det["collection"]    = r.get("coll", "")
                    det["events"]        = r.get("evts", "")
                    det["relationships"] = r.get("rels", "")
                    records.append(det)
                    log(f"  ✓  {det['title'][:70]}  ({r['score']}%)")
                    await page.goto(results_url,
                                    wait_until="domcontentloaded", timeout=20000)
                    try:
                        await page.wait_for_selector("tbody tr", timeout=8000)
                    except Exception:
                        await asyncio.sleep(3)


            # ── 6. СОХРАНЕНИЕ ──────────────────────────────────────── #
            _prog(96, "Сохранение файлов...")
            docx_p = output_folder / f"{file_base}.docx"
            xlsx_p = output_folder / f"{file_base}.xlsx"
            sd = sx = False
            if want_docx and records:
                write_docx(docx_p, records, qlines)
                sd = True
                log(f"  Word: {docx_p}")
            if want_xlsx and records:
                write_xlsx(xlsx_p, records, qlines)
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
