r"""
hryc_split.py — pull ONE gazette's ALREADY-DOWNLOADED scans OUT of the mixed «hryc.by_<year>»
folders into their own folder, WITHOUT re-downloading anything.

Use it when a gazette got downloaded into the shared D:\Archives\hryc.by_<year> folders (mixed
with another gazette). It re-runs that gazette's SEARCH on hryc (the same ~500 substrings) but
downloads NOTHING — it only reads which document ids (base HAIDs) belong to the source — then
MOVES the matching scan files from <src>/hryc.by_<year>/ into <dest>/hryc.by_<year>/. What stays
behind in the mixed folders is the OTHER gazette.

    python hryc_split.py --source "губернские киевские" --from 1854 --to 1917 ^
                         --dest "D:\Archives\Киев" --show

--show is recommended (solve captchas; a captcha skipped in headless = some files left behind).
It's safe to run again — already-moved files are simply not there anymore.
"""

import argparse
import asyncio
import shutil
import sys
from pathlib import Path

import hryc_scraper as H
import hryc_pages as P
import hryc_pages_resume as RZ
from hryc_pages import (_aid, _base_haid, _goto, _resolve_sources, _source_paths,
                        _creds, _instance_profile, _log, DEFAULT_OUT)

try:
    from playwright.async_api import async_playwright
    _PW_OK = True
except Exception:
    _PW_OK = False


async def _source_haids(page, sources, year, max_pages, zero_stop, ns, nf, se):
    """Every base HAID the source returns for this year — SEARCH ONLY, no download."""
    bases, zero = set(), 0
    for combo in RZ.COMBOS:
        before = len(bases)
        seen = set()
        for pno in range(1, max_pages + 1):
            u = H.build_search_url(combo, sources, pno, doc_dates=str(year),
                                   no_stemming=ns, no_fuzziness=nf, show_experts=se)
            if not await _goto(page, u):
                break
            res = H.parse_results(await page.content(), _log)
            new = 0
            for r in res["rows"]:
                uu = r.get("url")
                if uu and uu not in seen:
                    seen.add(uu)
                    b = _base_haid(_aid(uu))
                    if b:
                        bases.add(b)
                    new += 1
            if new == 0:
                break
        zero = zero + 1 if len(bases) == before else 0
        if zero >= zero_stop:
            break
    return bases


def _move_year(bases, year, src_root, dest_root, log):
    src = src_root / f"hryc.by_{year}"
    if not src.exists() or not bases:
        return 0
    dst = dest_root / f"hryc.by_{year}"
    suffix = f"_{year}"
    moved = 0
    for f in list(src.iterdir()):
        if not f.is_file():
            continue
        stem = f.stem
        if not stem.endswith(suffix):
            continue
        aid = stem[:-len(suffix)]
        if aid.startswith("HAID") and _base_haid(aid) in bases:
            dst.mkdir(parents=True, exist_ok=True)
            try:
                shutil.move(str(f), str(dst / f.name))
                moved += 1
            except Exception as e:
                log(f"    !! не перенести {f.name}: {type(e).__name__}")
    return moved


async def run(source_ids, years, src_root, dest_root, max_pages, zero_stop, instance,
              headless, ns, nf, se):
    if not _PW_OK:
        _log("Playwright не установлен."); return
    src_root, dest_root = Path(src_root), Path(dest_root)
    P._HEADLESS = headless
    profile = _instance_profile(instance, False)
    _log(f"Разделение источника:  {src_root}\\hryc.by_<год>  →  {dest_root}\\hryc.by_<год>")
    _log(f"Годы: {years[0]}…{years[-1]}, профиль {profile.name}, "
         f"браузер: {'headless' if headless else 'видимый (--show)'}  — сканы НЕ качаются, только переносятся")

    def _clear():
        for n in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            try: (profile / n).unlink()
            except Exception: pass

    async with async_playwright() as pw:
        ctx = None
        for _ in range(3):
            _clear()
            for ch in ("chrome", None):
                try:
                    kw = dict(headless=headless, accept_downloads=True, no_viewport=True,
                              args=["--disable-blink-features=AutomationControlled",
                                    "--window-size=1680,1050" if headless else "--start-maximized"])
                    if ch:
                        kw["channel"] = ch
                    ctx = await pw.chromium.launch_persistent_context(str(profile), **kw)
                    break
                except Exception as e:
                    _log(f"  !! {ch or 'chromium'}: {type(e).__name__}")
            if ctx:
                break
            await asyncio.sleep(3)
        if ctx is None:
            _log("  !! Не удалось открыть браузер."); return
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for ex in list(ctx.pages):
            if ex is not page:
                try: await ex.close()
                except Exception: pass
        total = 0
        try:
            await _goto(page, H.BASE_URL + "/")
            await H._accept_cookies(page, _log)
            if not await H._is_logged_in(page):
                em, pwd = _creds()
                if em and pwd:
                    await H._login(page, em, pwd, _log)
                if not await H._is_logged_in(page) and not headless:
                    _log("  → Войди в открытом окне (жду до 3 минут)…")
                    for _ in range(180):
                        if await H._is_logged_in(page):
                            break
                        await asyncio.sleep(1)
            if not await H._is_logged_in(page):
                _log("  !! НЕ залогинен — без поиска не определить документы источника. "
                     "Запусти с --show и войди."); return
            _log("  → Статус: залогинен")
            for year in years:
                bases = await _source_haids(page, source_ids, year, max_pages, zero_stop, ns, nf, se)
                moved = _move_year(bases, str(year), src_root, dest_root, _log)
                total += moved
                _log(f"[{year}] документов источника: {len(bases)}, перенесено файлов: {moved}")
            _log(f"\nГотово. Перенесено всего: {total} файлов  →  {dest_root}")
            _log("В hryc.by_<год> остался только другой источник (Могилёв).")
        finally:
            try: await ctx.close()
            except Exception: pass


def main():
    ap = argparse.ArgumentParser(
        description="Вытащить уже скачанные сканы ОДНОГО источника из смешанных hryc.by_<год> в "
                    "отдельную папку, БЕЗ повторной закачки (поиск только определяет, какие "
                    "документы переносить).")
    ap.add_argument("--source", action="append", default=[], required=True, metavar="PHRASE",
                    help="источник (фраза пути), напр. \"губернские киевские\"")
    ap.add_argument("--from", dest="y_from", type=int, required=True, help="первый год")
    ap.add_argument("--to", dest="y_to", type=int, required=True, help="последний год")
    ap.add_argument("--src-root", default=DEFAULT_OUT,
                    help="где смешанные hryc.by_<год> (по умолч. D:\\Archives)")
    ap.add_argument("--dest", required=True, help="куда переносить, напр. D:\\Archives\\Киев")
    ap.add_argument("--max-pages", type=int, default=25)
    ap.add_argument("--zero-stop", type=int, default=10)
    ap.add_argument("--instance", type=int, default=0, metavar="N",
                    help="отдельный профиль (чтобы не конфликтовать с запущенным загрузчиком)")
    ap.add_argument("--show", action="store_true", help="показать окно (для входа и капчи)")
    ap.add_argument("--no-stemming", action="store_true")
    ap.add_argument("--no-fuzziness", action="store_true")
    ap.add_argument("--no-experts", action="store_true")
    a = ap.parse_args()
    ids, chosen = _resolve_sources(a.source)
    if not ids:
        ap.error(f"источник не найден: {a.source}. Запусти: python hryc_pages.py --list-sources")
    print("Источник:", chosen)
    step = -1 if a.y_from >= a.y_to else 1
    years = list(range(a.y_from, a.y_to + step, step))
    asyncio.run(run(ids, years, a.src_root, a.dest, a.max_pages, a.zero_stop, a.instance,
                    not a.show, a.no_stemming, a.no_fuzziness, not a.no_experts))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано.")
