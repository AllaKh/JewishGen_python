"""
hryc_grodno.py — AUTO bulk downloader for «Гродненские губернские ведомости» (hryc.by).

Same engine as hryc_pages_resume.py, preset for the Grodno provincial gazette:
  • source = Газеты / Губернские ведомости / Гродненские  (hcid1_4_9)
  • years 1838–1839 (override with --from / --to)
  • each year → ONE folder «Гродненские губернские ведомости <год>»
  • ~500 query variants per year; saves every NEW document (skips what's already on disk);
    after 10 queries in a row with 0 new scans → next year.

    python hryc_grodno.py --show        # FIRST run: log in / solve captcha in the window
    python hryc_grodno.py               # 1838, 1839, headless, logs only
    python hryc_grodno.py --instance 3  # run in parallel (its own Chrome profile)

All hryc_pages_resume.py flags work (--zero-stop, --max-pages, --instance, --show, …); you can
override the years with --from / --to or the source with --source.
"""

import hryc_pages_resume as RZ

GAZETTE      = "Гродненские губернские ведомости"
SOURCE       = "губернские гродненские"   # → hcid1_4_9 (Газеты / Губернские ведомости / Гродненские)
YEAR_FROM, YEAR_TO = 1838, 1839


if __name__ == "__main__":
    try:
        RZ.main(default_from=YEAR_FROM, default_to=YEAR_TO,
                default_gazette=GAZETTE, default_source=SOURCE)
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
