"""
hryc_mogilev.py — AUTO bulk downloader for «Могилевские губернские ведомости» (hryc.by).

This COMBINES the two old Mogilev scripts (the 1914→1838 run + the 1915-1917 wrapper) into ONE,
preset for the Mogilev provincial gazette:
  • source = Газеты / Губернские ведомости / Могилевские  (hcid1_4_2)
  • years 1917 → 1838 (the full span; override with --from / --to)
  • each year → ONE folder «Могилевские губернские ведомости <год>»
  • ~500 query variants per year; saves every NEW document (skips what's already on disk);
    after 10 queries in a row with 0 new scans → next year.

    python hryc_mogilev.py --show        # FIRST run: log in / solve captcha in the window
    python hryc_mogilev.py               # 1917…1838, headless, logs only
    python hryc_mogilev.py --instance 1  # run in parallel (its own Chrome profile)

All hryc_pages_resume.py flags work (--zero-stop, --max-pages, --instance, --show, …); you can
override the years with --from / --to or the source with --source.
"""

import hryc_pages_resume as RZ

GAZETTE      = "Могилевские губернские ведомости"
SOURCE       = "губернские могилевские"   # → hcid1_4_2 (Газеты / Губернские ведомости / Могилевские)
YEAR_FROM, YEAR_TO = 1917, 1838


if __name__ == "__main__":
    try:
        RZ.main(default_from=YEAR_FROM, default_to=YEAR_TO,
                default_gazette=GAZETTE, default_source=SOURCE)
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
