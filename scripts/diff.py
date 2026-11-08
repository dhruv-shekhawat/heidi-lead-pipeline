#!/usr/bin/env python3
"""Compare two workbooks and write a Markdown report.

Usage: python3 scripts/diff.py OLD.xlsx NEW.xlsx [--out data/diff_report.md]

Rows are identified per tab by normalised Firma + PLZ (from Bezirk/PLZ); a NEW row whose key is
not found is retried via its Name alt (catches rows renamed by light_enrich). Reports: new rows,
removed rows, renames, Prio changes and changes in the TRACK columns (Einheiten, Verwaltungsart,
Kundenportal, contact columns).
If OLD has no Prio column, Prio is summarised instead of listed row by row.
"""
import argparse
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import DATA, norm_name, plz_of  # noqa: E402

TRACK = ["Einheiten", "Einheiten Berlin", "Verwaltungsart", "Kundenportal", "Ansprechpartner", "Funktion", "E-Mail",
         "E-Mail-Status", "Allgemeine E-Mail"]


def rows_of(ws):
    hdr = [c.value for c in ws[1]]
    out = []
    for vals in ws.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(hdr, vals))
        if rec.get("Firma"):
            out.append(rec)
    return hdr, out


def key(rec, field="Firma"):
    return f"{norm_name(rec.get(field))}|{plz_of(rec.get('Bezirk/PLZ'))}"


def md_row(cells):
    return "| " + " | ".join(str(c if c is not None else "").replace("|", "/").replace("\n", " ") for c in cells) + " |"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("old", type=Path)
    ap.add_argument("new", type=Path)
    ap.add_argument("--out", type=Path, default=DATA / "diff_report.md")
    a = ap.parse_args()
    wa, wb = openpyxl.load_workbook(a.old, read_only=True), openpyxl.load_workbook(a.new, read_only=True)
    L = [f"# Diff: `{a.old.name}` → `{a.new.name}`", "",
         f"Erstellt {datetime.now():%Y-%m-%d %H:%M}. Zeilen-Identität: normalisierte Firma + PLZ "
         f"(Fallback Name alt).", ""]
    only_old = [t for t in wa.sheetnames if t not in wb.sheetnames]
    only_new = [t for t in wb.sheetnames if t not in wa.sheetnames]
    if only_old or only_new:
        L += [f"Tabs nur alt: {only_old or '–'} · Tabs nur neu: {only_new or '–'}", ""]
    summary = []
    for tab in [t for t in wa.sheetnames if t in wb.sheetnames]:
        ha, ra = rows_of(wa[tab])
        hb, rb = rows_of(wb[tab])
        if "Firma" not in ha or "Firma" not in hb:
            continue
        old_by = defaultdict(list)
        for rec in ra:
            old_by[key(rec)].append(rec)
        pairs, new_rows = [], []
        for rec in rb:
            k = key(rec)
            if not old_by.get(k) and rec.get("Name alt"):
                k = key(rec, "Name alt")
            if old_by.get(k):
                # same key several times (duplicates): prefer the old row with the same website
                cand = old_by[k]
                i = next((j for j, o in enumerate(cand) if o.get("Website") and o.get("Website") == rec.get("Website")), 0)
                pairs.append((cand.pop(i), rec))
            else:
                new_rows.append(rec)
        removed = [r for lst in old_by.values() for r in lst]
        L += [f"## {tab}", "", f"Alt {len(ra)} Zeilen · neu {len(rb)} Zeilen · zugeordnet {len(pairs)} · "
              f"**neu {len(new_rows)}** · **entfernt {len(removed)}**", ""]
        summary.append((tab, len(new_rows), len(removed)))
        if new_rows:
            L += [f"### Neue Zeilen ({len(new_rows)})", "",
                  md_row(["Firma", "PLZ", "Quelle", "Recherchetiefe", "Prio"]), md_row(["---"] * 5)]
            L += [md_row([r.get("Firma"), plz_of(r.get("Bezirk/PLZ")), r.get("Quelle"), r.get("Recherchetiefe"),
                          r.get("Prio")]) for r in new_rows] + [""]
        if removed:
            L += [f"### Entfernte Zeilen ({len(removed)})", "", md_row(["Firma", "PLZ", "Quelle"]), md_row(["---"] * 3)]
            L += [md_row([r.get("Firma"), plz_of(r.get("Bezirk/PLZ")), r.get("Quelle")]) for r in removed] + [""]
        renamed = [(o, n) for o, n in pairs if o.get("Firma") != n.get("Firma")]
        if renamed:
            L += [f"### Umbenannt ({len(renamed)})", "", md_row(["alt", "neu"]), md_row(["---"] * 2)]
            L += [md_row([o.get("Firma"), n.get("Firma")]) for o, n in renamed] + [""]
        if "Prio" in hb:
            if "Prio" not in ha:
                c = Counter(n.get("Prio") for _, n in pairs)
                L += ["### Prio", "", f"Alte Datei hat keine Prio-Spalte – {len(pairs)} bestehende Zeilen neu "
                      f"priorisiert: " + ", ".join(f"{k}: {v}" for k, v in c.most_common()), ""]
            else:
                ch = [(o, n) for o, n in pairs if o.get("Prio") != n.get("Prio")]
                L += [f"### Prio-Änderungen ({len(ch)})", ""]
                if ch:
                    L += [md_row(["Firma", "alt", "neu", "Prio-Grund"]), md_row(["---"] * 4)]
                    L += [md_row([n.get("Firma"), o.get("Prio"), n.get("Prio"), n.get("Prio-Grund")]) for o, n in ch]
                L += [""]
        for f in TRACK:
            if f not in ha and f not in hb:
                continue
            ch = [(o, n) for o, n in pairs if (o.get(f) or None) != (n.get(f) or None)]
            L += [f"### {f}-Änderungen ({len(ch)})", ""]
            if ch:
                L += [md_row(["Firma", "alt", "neu"]), md_row(["---"] * 3)]
                L += [md_row([n.get("Firma"), o.get(f), n.get(f)]) for o, n in ch]
            L += [""]
    a.out.write_text("\n".join(L), encoding="utf-8")
    for tab, n, r in summary:
        print(f"{tab}: neu {n}, entfernt {r}")
    print(f"Geschrieben: {a.out}")


if __name__ == "__main__":
    main()
