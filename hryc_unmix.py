r"""
hryc_unmix.py — remove ONE gazette's scans from the mixed «hryc.by_<year>» folders, using a
CLEAN per-gazette folder as the reference. Pure local file operation — no browser, no network.

Background: scan files are named by a content hash (HAID) that does NOT say which gazette they
came from, so two gazettes downloaded into the same «hryc.by_<year>» can't be told apart by name.
Once you've RE-DOWNLOADED one gazette into its own clean folder (e.g. Kiev → D:\Archives\Киев via
hryc_kiev.py), this tool finds every file in the mixed root (D:\Archives\hryc.by_<year>) whose
document id (base HAID) is also present in the clean folder, and removes it from the mixed root —
leaving the OTHER gazette (Mogilev) clean there.

SAFE BY DEFAULT — a dry run that only reports counts. Then:
  --apply            MOVE the matches into a quarantine folder (reversible) and report it
  --apply --delete   actually DELETE the matches

    python hryc_unmix.py --ref "D:\\Archives\\Киев" --mixed "D:\\Archives"            # dry run
    python hryc_unmix.py --ref "D:\\Archives\\Киев" --mixed "D:\\Archives" --apply    # move to quarantine
"""

import argparse
import re
import shutil
import sys
from pathlib import Path

_YEAR_DIR = re.compile(r"hryc\.by_(\d{3,4})$")


def _base_haid(stem: str, year: str):
    """base HAID of a scan file «<aid>_<year>» (aid = «HAID…_<page>»); None if it doesn't fit."""
    suf = f"_{year}"
    if not stem.endswith(suf):
        return None
    aid = stem[:-len(suf)]
    if not aid.startswith("HAID"):
        return None
    return re.sub(r"_\d+$", "", aid)        # drop the trailing «_<page>»


def main():
    ap = argparse.ArgumentParser(
        description="Убрать сканы одной газеты из смешанных hryc.by_<год>, опираясь на её чистую "
                    "папку-эталон. По умолчанию только показывает (dry-run).")
    ap.add_argument("--ref", required=True,
                    help="чистая папка газеты-эталона, напр. D:\\Archives\\Киев")
    ap.add_argument("--mixed", required=True,
                    help="смешанный корень, напр. D:\\Archives")
    ap.add_argument("--apply", action="store_true",
                    help="реально применить (иначе только показать, что совпало)")
    ap.add_argument("--delete", action="store_true",
                    help="с --apply: УДАЛИТЬ совпавшие файлы вместо переноса в карантин")
    ap.add_argument("--trash", default="",
                    help="папка карантина (по умолчанию <mixed>\\_unmixed_<имя эталона>)")
    a = ap.parse_args()

    ref, mixed = Path(a.ref), Path(a.mixed)
    if not ref.exists():
        ap.error(f"нет папки-эталона: {ref}")
    if not mixed.exists():
        ap.error(f"нет смешанной папки: {mixed}")
    trash = Path(a.trash) if a.trash else mixed / f"_unmixed_{ref.name}"

    total_match = total_done = 0
    for refdir in sorted(ref.glob("hryc.by_*")):
        m = _YEAR_DIR.search(refdir.name)
        if not m or not refdir.is_dir():
            continue
        year = m.group(1)
        mixdir = mixed / refdir.name
        if not mixdir.exists() or mixdir.resolve() == refdir.resolve():
            continue                                       # nothing mixed for this year / same dir

        # EXACT filenames present in the CLEAN reference folder (ПОЛНОЕ совпадение имени — не по
        # части HAID): only a byte-for-byte same-named scan counts, so we never touch the other
        # gazette's files by mistake.
        refnames = set(f.name for f in refdir.iterdir() if f.is_file())
        if not refnames:
            continue

        matches = [f for f in mixdir.iterdir()
                   if f.is_file() and f.name in refnames]
        if not matches:
            continue
        total_match += len(matches)
        print(f"[{year}] эталон: {len(refnames)} файлов  →  ПОЛНОЕ совпадение в смешанной: {len(matches)}"
              + ("" if a.apply else "   (dry-run)"))

        if a.apply:
            dst = trash / refdir.name
            for f in matches:
                try:
                    if a.delete:
                        f.unlink()
                    else:
                        dst.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(f), str(dst / f.name))
                    total_done += 1
                except Exception as e:
                    print(f"   !! {f.name}: {type(e).__name__}")

    print()
    if not a.apply:
        print(f"DRY-RUN: совпало {total_match} файлов в смешанных папках. "
              f"Запусти с --apply, чтобы {'удалить' if a.delete else 'перенести их в карантин'}.")
    elif a.delete:
        print(f"Готово. УДАЛЕНО {total_done} файлов из {mixed} (остался только Могилёв).")
    else:
        print(f"Готово. Перенесено в карантин: {total_done} файлов → {trash}\n"
              f"Проверь, что в hryc.by_<год> остался только Могилёв, потом карантин можно удалить.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано.")
