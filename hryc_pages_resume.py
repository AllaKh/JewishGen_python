"""
hryc_pages_resume.py — same as hryc_pages.py, but SKIPS documents already saved on disk.

Use this to re-run the SAME year with DIFFERENT queries without re-downloading. Any document
whose scans are already in the output folder is skipped WHOLE — without opening it.

KEY DIFFERENCE from the old (reverted) version: the skip is keyed on the DOCUMENT ID
(base HAID, read from the result URL), NOT on the folder title. Several different documents
can share a title — keying on the title made it skip whole documents by mistake (that was the
bug). Keying on the unique HAID skips exactly the documents already downloaded, and a folder
you DELETE is re-fetched automatically (the on-disk scan filenames are the source of truth).

Run SEVERAL copies at once with --instance (each needs a different N) — see hryc_pages.py.

Examples:
    python hryc_pages_resume.py "Левин" 1900
    python hryc_pages_resume.py "*а*" 1900 --instance 2
    python hryc_pages_resume.py "Шендерович" 1859 --source "губернские могилевские"

CAPTCHA: same as hryc_pages.py — the script NEVER solves it; it waits for you to click it.
"""

import argparse
import asyncio

import hryc_pages as P


def main():
    ap = argparse.ArgumentParser(
        description="hryc.by: save scans for a word + year, SKIPPING documents already on disk "
                    "(skip keyed on the document id, not the folder title).")
    ap.add_argument("query", nargs="?", default="", help="search word (e.g. Шендерович or *а*)")
    ap.add_argument("year", nargs="?", default="", help="year (Doc dates), e.g. 1859")
    ap.add_argument("--query", dest="query_opt", default="")
    ap.add_argument("--year", dest="year_opt", default="")
    ap.add_argument("--source", action="append", default=[], metavar="PHRASE",
                    help="limit to sources whose full path contains ALL these words; repeat for "
                         "several; omit = all.")
    ap.add_argument("--list-sources", action="store_true",
                    help="print every source path (optionally filtered by --source) and exit")
    ap.add_argument("--out", default=P.DEFAULT_OUT)
    ap.add_argument("--max-pages", type=int, default=25, help="result pages to scan")
    ap.add_argument("--max-folders", type=int, default=0, help="stop after N new folders (0 = all)")
    ap.add_argument("--shared", action="store_true",
                    help="reuse the app's .hryc_profile (close the app first)")
    ap.add_argument("--instance", type=int, default=0, metavar="N",
                    help="run SEVERAL copies at once: give each a different N (1, 2, 3…). Each "
                         "gets its own Chrome profile copied from the master (already logged in). "
                         "0 = master profile (log in here first). Practical max ~3-4.")
    ap.add_argument("--show", action="store_true",
                    help="show the browser window (default = headless, logs only). Use it for "
                         "the FIRST login and for solving a CAPTCHA.")
    ap.add_argument("--no-stemming",  action="store_true", help="«Без стемминга» — exact word form only")
    ap.add_argument("--no-fuzziness", action="store_true", help="«Без ошибок» — no typo tolerance")
    ap.add_argument("--no-experts",   action="store_true", help="«Эксперты» OFF (default ON → more results)")
    a = ap.parse_args()

    source_ids, chosen = P._resolve_sources(a.source)
    if a.list_sources:
        for sid, path in P._source_paths():
            if not a.source or sid in source_ids:
                print(f"{sid:24} {path}")
        return
    if a.source and not source_ids:
        ap.error(f"no source matches {a.source!r}. Run --list-sources to see them.")
    if a.source:
        print("Источники:"); [print("  •", p) for p in chosen]

    query = a.query_opt or a.query
    year = a.year_opt or a.year
    if not query or not year:
        ap.error('give a search word and a year, e.g.:  python hryc_pages_resume.py "Левин" 1900')
    # skip_saved=True is the ONLY difference from hryc_pages.py — same battle-tested run().
    asyncio.run(P.run(query, year, a.out, a.max_pages, a.max_folders, a.shared, source_ids,
                      a.no_stemming, a.no_fuzziness, not a.no_experts,
                      instance=a.instance, skip_saved=True, headless=not a.show))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано — скачанное сохранено.")
