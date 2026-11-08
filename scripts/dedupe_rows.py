#!/usr/bin/env python3
"""Merge same-name/same-PLZ duplicate rows in the Hausverwaltungen tab (see add_sources.merge_same_name_rows).

The same merge runs automatically at the start of add_sources.py and at the end of light_enrich.py;
this CLI applies it to an existing workbook. Re-run score.py afterwards.

Usage: python3 scripts/dedupe_rows.py --in IN.xlsx --out OUT.xlsx
"""
import argparse
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from add_sources import merge_same_name_rows  # noqa: E402
from common import DATA  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    a = ap.parse_args()
    if a.out.resolve() in (a.inp.resolve(), (DATA / "heidi_merged_v3.xlsx").resolve()):
        sys.exit(f"Refusing to write {a.out}")
    wb = openpyxl.load_workbook(a.inp)
    merged = merge_same_name_rows(wb["Hausverwaltungen"])
    wb.save(a.out)
    print(f"zusammengeführt: {len(merged)} | Geschrieben: {a.out.name}")


if __name__ == "__main__":
    main()
