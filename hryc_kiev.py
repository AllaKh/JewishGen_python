r"""
hryc_kiev.py — AUTO bulk downloader for «Киевские губернские ведомости» (hryc.by).

  • source = Газеты / Губернские ведомости / Киевские  (hcid1_4_8)
  • years 1854 → 1917 (override with --from / --to)
  • writes into ITS OWN folder «D:\Archives\Киев\hryc.by_<год>» — SEPARATE from Mogilev (whose
    legacy lives in «D:\Archives\hryc.by_<год>»). Kiev and Mogilev are different gazettes, and
    their scan filenames are content hashes that don't say which gazette they came from, so they
    must NOT share a folder. (Change with --out if you want a different location.)
  • ~500 query variants per year; saves every NEW document (skips what's already on disk);
    after 10 queries in a row with 0 new scans → next year.

    python hryc_kiev.py --show        # FIRST run: log in / solve captcha in the window
    python hryc_kiev.py               # 1854…1917, headless, logs only
    python hryc_kiev.py --instance 4  # run in parallel (its own Chrome profile)

All hryc_pages_resume.py flags work (--zero-stop, --max-pages, --instance, --show, …); you can
override the years with --from / --to or the source with --source.
"""

import hryc_pages_resume as RZ

SOURCE       = "губернские киевские"   # → hcid1_4_8 (Газеты / Губернские ведомости / Киевские)
YEAR_FROM, YEAR_TO = 1854, 1917


if __name__ == "__main__":
    try:
        # SEPARATE folder so Kiev never mixes with Mogilev's «D:\Archives\hryc.by_<год>».
        RZ.main(default_from=YEAR_FROM, default_to=YEAR_TO, default_source=SOURCE,
                default_out=r"D:\Archives\Киев")
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
