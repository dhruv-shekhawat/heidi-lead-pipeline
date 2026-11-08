#!/usr/bin/env python3
"""Merge hand-updated contact columns from a v10 workbook into the scored workbook.

For every Hausverwaltungen row of --contacts (default data/heidi_contacts_v10.xlsx) the matching
row in --base (default data/heidi_scored_v9.xlsx) is found by full identity (normalised Firma + PLZ +
Website domain + Name alt), then Website domain, then Firma + PLZ or Name alt + PLZ – a name key is
only used when it points to exactly one row (the list contains a few same-name/same-PLZ duplicates);
ambiguous rows are reported, not guessed. For each of E-Mail, E-Mail-Status, LinkedIn,
Ansprechpartner the v10 value is copied when the v10 cell is non-empty (this deliberately
overwrites: v10 is the hand-checked source). Empty v10 cells never clear a value.
Other columns, other tabs and the scoring are untouched (contact fields carry no points).
Unmatched v10 rows are reported, not added.

Usage: python3 scripts/merge_contacts.py [--base …v9.xlsx] [--contacts …v10.xlsx] [--out …v11.xlsx]
Then:  python3 scripts/export_sheet.py --in data/heidi_scored_v11.xlsx --out data/heidi_final.xlsx
"""
import argparse
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import DATA, norm_domain, norm_name, plz_of  # noqa: E402

HV = "Hausverwaltungen"
COLS = ["E-Mail", "E-Mail-Status", "LinkedIn", "Ansprechpartner"]


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def keys(get):
    """Most specific first: full identity, domain, then name keys."""
    plz = plz_of(get("Bezirk/PLZ"))
    out = [f"id:{norm_name(get('Firma'))}|{plz}|{norm_domain(get('Website'))}|{norm_name(get('Name alt'))}"]
    if norm_domain(get("Website")):
        out.append(f"d:{norm_domain(get('Website'))}")
    out.append(f"n:{norm_name(get('Firma'))}|{plz}")
    if get("Name alt"):
        out.append(f"n:{norm_name(get('Name alt'))}|{plz}")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", type=Path, default=DATA / "heidi_scored_v9.xlsx")
    ap.add_argument("--contacts", type=Path, default=DATA / "heidi_contacts_v10.xlsx")
    ap.add_argument("--out", type=Path, default=DATA / "heidi_scored_v11.xlsx")
    a = ap.parse_args()
    if a.out.resolve() in (a.base.resolve(), a.contacts.resolve(), (DATA / "heidi_merged_v3.xlsx").resolve()):
        sys.exit(f"Refusing to write {a.out}")
    wb = openpyxl.load_workbook(a.base)
    ws = wb[HV]
    hm = header_map(ws)
    for c in COLS:
        if c not in hm:
            sys.exit(f"Spalte {c} fehlt in {a.base.name}")
    idx = {}
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, hm["Firma"]).value:
            for k in keys(lambda col, r=r: ws.cell(r, hm[col]).value if col in hm else None):
                idx.setdefault(k, set()).add(r)

    src = openpyxl.load_workbook(a.contacts, read_only=True)[HV]
    shdr = [c.value for c in next(src.iter_rows(max_row=1))]
    missing = [c for c in COLS if c not in shdr]
    if missing:
        print(f"Hinweis: Spalten fehlen in {a.contacts.name}: {missing}")
    changed = {c: 0 for c in COLS}
    rows_changed, unmatched, n = set(), [], 0
    for vals in src.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(shdr, vals))
        if not rec.get("Firma"):
            continue
        n += 1
        row = None
        for k in keys(rec.get):
            if k in idx:
                if len(idx[k]) == 1:
                    row = next(iter(idx[k]))
                    break
                if k.startswith("n:"):
                    continue  # ambiguous name key: try the next key, never guess
        if not row:
            if any(rec.get(c) not in (None, "") for c in COLS):
                unmatched.append(rec["Firma"])
            continue
        for c in COLS:
            v = rec.get(c)
            if v not in (None, "") and ws.cell(row, hm[c]).value != v:
                ws.cell(row, hm[c]).value = v
                changed[c] += 1
                rows_changed.add(row)
    wb.save(a.out)
    print(f"{a.contacts.name}: {n} Zeilen gelesen | Zeilen mit Änderung: {len(rows_changed)}")
    print("Geänderte Zellen je Spalte:", changed)
    print(f"Nicht zugeordnet (mit Kontaktdaten): {len(unmatched)} {unmatched[:20]}")
    print(f"Geschrieben: {a.out}")


if __name__ == "__main__":
    main()
