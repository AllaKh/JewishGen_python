"""
hryc_minsk.py — AUTO bulk downloader for «Минские губернские ведомости» (hryc.by).

Same engine as hryc_pages_resume.py, preset for the Minsk provincial gazette:
  • source = Газеты / Губернские ведомости / Минские  (hcid1_4_11)
  • years 1837–1917 (override with --from / --to)
  • each year → ONE folder «Минские губернские ведомости <год>»
  • ~500 query variants per year; saves every NEW document (skips what's already on disk);
    after 10 queries in a row with 0 new scans → next year.

    python hryc_minsk.py --show        # FIRST run: log in / solve captcha in the window
    python hryc_minsk.py               # 1837…1917, headless, logs only
    python hryc_minsk.py --instance 2  # run in parallel (its own Chrome profile)

All hryc_pages_resume.py flags work (--zero-stop, --max-pages, --instance, --show, …); you can
override the years with --from / --to or the source with --source.
"""

import hryc_pages_resume as RZ

GAZETTE      = "Минские губернские ведомости"
SOURCE       = "губернские минские"     # → hcid1_4_11 (Газеты / Губернские ведомости / Минские)
YEAR_FROM, YEAR_TO = 1837, 1917


if __name__ == "__main__":
    try:
        RZ.main(default_from=YEAR_FROM, default_to=YEAR_TO,
                default_gazette=GAZETTE, default_source=SOURCE)
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
