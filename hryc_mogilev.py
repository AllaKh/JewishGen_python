"""
hryc_mogilev.py — AUTO bulk downloader for «Могилевские губернские ведомости» (hryc.by).

  • source = Газеты / Губернские ведомости / Могилевские  (hcid1_4_2)
  • years 1838 → 1917 (override with --from / --to)
  • writes into the EXISTING «hryc.by_<год>» folders, exactly like before — it does NOT create
    a separate «Могилевские губернские ведомости <год>» folder. (No --gazette → the old
    per-document folder naming «hryc.by_<год>» is kept; a year you've already downloaded is
    reused, nothing duplicated.)
  • ~500 query variants per year; saves every NEW document (skips what's already on disk);
    after 10 queries in a row with 0 new scans → next year.

    python hryc_mogilev.py --show        # FIRST run: log in / solve captcha in the window
    python hryc_mogilev.py               # 1838…1917, headless, logs only
    python hryc_mogilev.py --instance 1  # run in parallel (its own Chrome profile)

All hryc_pages_resume.py flags work (--zero-stop, --max-pages, --instance, --show, …); you can
override the years with --from / --to or the source with --source.
"""

import hryc_pages_resume as RZ

SOURCE       = "губернские могилевские"   # → hcid1_4_2 (Газеты / Губернские ведомости / Могилевские)
YEAR_FROM, YEAR_TO = 1838, 1917


if __name__ == "__main__":
    try:
        # NO default_gazette → keeps the existing «hryc.by_<год>» folders (как и раньше).
        RZ.main(default_from=YEAR_FROM, default_to=YEAR_TO, default_source=SOURCE)
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
