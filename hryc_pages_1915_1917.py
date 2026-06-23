"""
hryc_pages_1915_1917.py — the SAME automatic bulk downloader as hryc_pages_resume.py, but for
the WWI years 1915–1917.

hryc_pages_resume.py walks 1914 → 1838 (downwards) and therefore does NOT cover 1915-1917. Run
THIS script alongside it to grab those three years. Everything else is identical — the same
common-substring combos, skip-by-document-id, headless, login, --instance, --show — it only
changes the default year range to 1915-1917 and hands off to hryc_pages_resume.

    python hryc_pages_1915_1917.py --show         # FIRST run: log in / solve captcha in the window
    python hryc_pages_1915_1917.py                # 1915, 1916, 1917 — headless, logs only
    python hryc_pages_1915_1917.py --instance 3   # run in parallel (its own Chrome profile)

All hryc_pages_resume.py flags work (--source, --zero-stop, --max-pages, --instance, --show,
--no-stemming/-fuzziness/-experts). You can still override the years with --from / --to.
"""

import hryc_pages_resume as RZ

YEAR_FROM, YEAR_TO = 1915, 1917


if __name__ == "__main__":
    try:
        RZ.main(default_from=YEAR_FROM, default_to=YEAR_TO)
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
