#!/usr/bin/env python3
"""Merge data/directory_clean.csv into the enriched workbook (default v2) -> data/heidi_merged_run.xlsx.

Never writes data/heidi_merged_v3.xlsx (frozen deliverable; it was built by this script plus
a one-off HVIM/hv-im merge, which the compact name normalisation now reproduces).

Usage:
  python3 scripts/merge.py [--in|--base data/heidi_enriched_v2.xlsx] [--out data/heidi_merged_run.xlsx]
                           [--plz-from data/heidi_enriched_v1.xlsx] [--csv data/directory_clean.csv]

Rules:
* Directory rows go into the "Hausverwaltungen" tab only; existing rows are not modified.
* Columns Recherchetiefe, Straße, Telefon, Makler auch are added to Hausverwaltungen only.
  Existing rows: "A (Website geprüft)". Wohnungsunternehmen stays unchanged.
* Dedupe (either-or): a directory row is a duplicate if its normalised domain OR
  normalised name+PLZ matches a row on Hausverwaltungen, Wohnungsunternehmen or Log
  (Log = excluded/renamed/uncertain firms, so they are not re-added). Rows without a PLZ
  (all Log rows) are matched on name alone, unless the name is only generic words.
* Name key = common.norm_name: lowercase, umlauts folded, Rechtsform tokens and branch
  suffixes (NL/Niederlassung/trailing 'Berlin') removed, then spaces/hyphens/punctuation
  removed, so 'HVIM …' == 'hv-im …'.
* Log rows with Bereich "ungeklärt" and no website are moved to Hausverwaltungen.
  Their PLZ comes from the v1 workbook (Log has no PLZ column). If one matches a
  directory row by name+PLZ, the two are merged into one row with Recherchetiefe
  "B (Verzeichnis)" (directory fills Website/Straße/Telefon); otherwise "C (nur Name)".
  The moved rows are deleted from Log; all other Log rows stay as they are.
"""
import argparse
import csv
import difflib
import re
import sys
from copy import copy
from datetime import date
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import norm_domain, norm_name, norm_tokens  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DEF_IN = ROOT / "data" / "heidi_enriched_v2.xlsx"
DEF_OUT = ROOT / "data" / "heidi_merged_run.xlsx"
# frozen deliverable, never written by this script
PROTECTED = [ROOT / "data" / "heidi_merged_v3.xlsx"]
DEF_PLZ = ROOT / "data" / "heidi_enriched_v1.xlsx"
DEF_CSV = ROOT / "data" / "directory_clean.csv"

HV_TAB, WU_TAB, LOG_TAB = "Hausverwaltungen", "Wohnungsunternehmen", "Log"
A_LABEL, B_LABEL, C_LABEL = "A (Website geprüft)", "B (Verzeichnis)", "C (nur Name)"
NEW_COLS = ["Recherchetiefe", "Straße", "Telefon", "Makler auch"]
PLZ_RE = re.compile(r"\b1\d{4}\b")
GENERIC = {"hausverwaltung", "hausverwaltungen", "verwaltung", "immobilien",
           "immobilienverwaltung", "hv", "grundstuecksverwaltung", "wohnungsverwaltung",
           "immobilienmanagement", "property", "management", "service", "services"}


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def add_columns(ws):
    hm = header_map(ws)
    last = ws.max_column
    tmpl = ws.cell(row=1, column=last)
    for name in NEW_COLS:
        if name in hm:
            continue
        last += 1
        cell = ws.cell(row=1, column=last, value=name)
        cell.font, cell.fill = copy(tmpl.font), copy(tmpl.fill)
        cell.alignment, cell.border = copy(tmpl.alignment), copy(tmpl.border)
        ws.column_dimensions[cell.column_letter].width = 22
    return header_map(ws)


def plz_of(v):
    m = PLZ_RE.search(str(v or ""))
    return m.group(0) if m else ""


def name_only_ok(name):
    toks = norm_tokens(name)
    return bool(toks) and not all(t in GENERIC for t in toks)


class Keys:
    """Dedupe keys with the tab each key came from (for reporting)."""

    def __init__(self):
        self.dom, self.np, self.name = {}, {}, {}

    def add(self, origin, names, website, plz):
        d = norm_domain(website)
        if d:
            self.dom.setdefault(d, origin)
        for n in names:
            if not n:
                continue
            nn = norm_name(n)
            if plz:
                self.np.setdefault(f"{nn}|{plz}", origin)
            elif name_only_ok(n):
                self.name.setdefault(nn, (origin, n))

    def match(self, row):
        d = norm_domain(row["Website"])
        nn = norm_name(row["Firma"])
        if d and d in self.dom:
            return "Domain", self.dom[d], None
        if f"{nn}|{row['PLZ']}" in self.np:
            return "Name+PLZ", self.np[f"{nn}|{row['PLZ']}"], None
        if nn in self.name:
            origin, orig_name = self.name[nn]
            return "Name allein", origin, orig_name
        return None, None, None


def copy_row_style(ws, src_row, dst_row):
    for col in range(1, ws.max_column + 1):
        src = ws.cell(src_row, col)
        if src.has_style:
            ws.cell(dst_row, col)._style = copy(src._style)


def append(ws, hm, style_row, values):
    r = ws.max_row + 1
    copy_row_style(ws, style_row, r)
    for k, v in values.items():
        if v not in (None, ""):
            ws.cell(r, hm[k], v)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", "--base", dest="inp", type=Path, default=DEF_IN)
    ap.add_argument("--out", type=Path, default=DEF_OUT)
    ap.add_argument("--plz-from", type=Path, default=DEF_PLZ)
    ap.add_argument("--csv", type=Path, default=DEF_CSV)
    a = ap.parse_args()
    if a.out.resolve() in [a.inp.resolve(), a.plz_from.resolve()] + [p.resolve() for p in PROTECTED]:
        sys.exit(f"Refusing to write {a.out}: input or protected workbook")
    today = date.today().isoformat()

    wb = openpyxl.load_workbook(a.inp)
    hv, wu, log = wb[HV_TAB], wb[WU_TAB], wb[LOG_TAB]
    keys = Keys()

    # --- existing Hausverwaltungen rows: tier A + keys
    hm = add_columns(hv)
    style_row = hv.max_row
    for r in range(2, hv.max_row + 1):
        if not hv.cell(r, hm["Firma"]).value:
            continue
        hv.cell(r, hm["Recherchetiefe"], A_LABEL)
        keys.add(HV_TAB, [hv.cell(r, hm["Firma"]).value, hv.cell(r, hm["Name alt"]).value],
                 hv.cell(r, hm["Website"]).value, plz_of(hv.cell(r, hm["Bezirk/PLZ"]).value))

    # --- Wohnungsunternehmen: keys only, tab untouched
    wm = header_map(wu)
    for r in range(2, wu.max_row + 1):
        if wu.cell(r, wm["Firma"]).value:
            keys.add(WU_TAB, [wu.cell(r, wm["Firma"]).value, wu.cell(r, wm["Name alt"]).value],
                     wu.cell(r, wm["Website"]).value, plz_of(wu.cell(r, wm["Bezirk/PLZ"]).value))

    # --- Log: split into rows to move (ungeklärt, no website) and the rest (keys only)
    lm = header_map(log)
    to_move = []  # (row_idx, dict)
    for r in range(2, log.max_row + 1):
        rec = {k: log.cell(r, c).value for k, c in lm.items()}
        if not rec.get("Firma"):
            continue
        if rec.get("Bereich") == "ungeklärt" and not rec.get("Website"):
            to_move.append((r, rec))
        else:
            keys.add(f"{LOG_TAB} ({rec.get('Bereich')})", [rec["Firma"], rec.get("Name alt")],
                     rec.get("Website"), "")

    # --- PLZ (+ Quelle) for moved rows from v1, matched on normalised name
    v1 = {}
    wb1 = openpyxl.load_workbook(a.plz_from, read_only=True)
    for ws1 in wb1.worksheets:
        h1 = {c.value: i for i, c in enumerate(next(ws1.iter_rows(max_row=1)))}
        for row in ws1.iter_rows(min_row=2, values_only=True):
            if row[h1["Firma"]]:
                v1.setdefault(norm_name(row[h1["Firma"]]), []).append(
                    {"Bezirk/PLZ": row[h1["Bezirk/PLZ"]], "Quelle": row[h1["Quelle"]]})

    with open(a.csv, encoding="utf-8-sig") as f:
        directory = list(csv.DictReader(f))
    dir_by_np = {}
    for i, d in enumerate(directory):
        dir_by_np.setdefault(f"{norm_name(d['Firma'])}|{d['PLZ']}", i)

    consumed, moved, missing_plz = set(), [], []
    for _, rec in to_move:
        hits = v1.get(norm_name(rec["Firma"]), [])
        plzs = {plz_of(h["Bezirk/PLZ"]) for h in hits} - {""}
        bez = hits[0]["Bezirk/PLZ"] if len(plzs) == 1 else None
        plz = plzs.pop() if len(plzs) == 1 else ""
        if not plz:
            missing_plz.append(rec["Firma"])
        di = None
        for n in (rec["Firma"], rec.get("Name alt")):
            if n and plz and f"{norm_name(n)}|{plz}" in dir_by_np:
                di = dir_by_np[f"{norm_name(n)}|{plz}"]
                break
        d = directory[di] if di is not None else None
        if di is not None:
            consumed.add(di)
        quelle = hits[0]["Quelle"] if hits else None
        hinweis = rec.get("Grund / Hinweis") or ""
        if d and d["Firma"] != rec["Firma"]:  # keep both spellings (e.g. HVIM / hv-im)
            hinweis += f" | Verzeichnisname: '{d['Firma']}' ({d['Quelle']})"
        if d and d["Kategorie"]:
            hinweis += f" | Verzeichnis-Kategorie: {d['Kategorie']}"
        moved.append({
            "Firma": rec["Firma"],
            "Name alt": rec.get("Name alt"),
            "Website": norm_domain(d["Website"]) if d else None,
            "Bezirk/PLZ": bez,
            "Quelle": "; ".join(x for x in (quelle, d["Quelle"] if d else None) if x),
            "Typ": rec.get("Typ"),
            "Hinweis": f"aus Log (ungeklärt) verschoben: {hinweis}".strip(),
            "Quellen": "; ".join(x for x in (rec.get("Quelle"), d["Quell-URL"] if d else None) if x),
            "Stand": today,
            "Recherchetiefe": B_LABEL if d else C_LABEL,
            "Straße": d["Straße"] if d else None,
            "Telefon": d["Telefon"] if d else None,
            "Makler auch": (d.get("Makler auch") or None) if d else None,
            "_plz": plz,
            "_dir": d,
        })

    # moved rows go in first, and become dedupe keys for the directory pass
    for m in moved:
        append(hv, hm, style_row, {k: v for k, v in m.items() if not k.startswith("_")})
        keys.add("Log → Hausverwaltungen (verschoben)", [m["Firma"], m["Name alt"]],
                 m["Website"], m["_plz"])

    # --- directory rows
    dup = {"Domain": {}, "Name+PLZ": {}, "Name allein": {}}
    name_only_hits, appended = [], 0
    for i, d in enumerate(directory):
        if i in consumed:
            continue
        how, origin, orig_name = keys.match(d)
        if how:
            dup[how][origin] = dup[how].get(origin, 0) + 1
            if how == "Name allein":
                name_only_hits.append((d["Firma"], d["PLZ"], orig_name, origin))
            continue
        dom = norm_domain(d["Website"])
        append(hv, hm, style_row, {
            "Firma": d["Firma"],
            "Website": dom or None,
            "Bezirk/PLZ": d["PLZ"],
            "Quelle": d["Quelle"],
            "Typ": "Hausverwaltung",
            "Hinweis": f"Verzeichnis-Kategorie: {d['Kategorie']}" if d["Kategorie"] else None,
            "Quellen": d["Quell-URL"],
            "Stand": today,
            "Recherchetiefe": B_LABEL,
            "Straße": d["Straße"],
            "Telefon": d["Telefon"],
            "Makler auch": d.get("Makler auch"),
        })
        keys.add("Verzeichnis (neu)", [d["Firma"]], d["Website"], d["PLZ"])
        appended += 1

    # --- remove moved rows from Log (bottom-up so indices stay valid)
    for r, _ in sorted(to_move, reverse=True):
        log.delete_rows(r)

    wb.save(a.out)

    # ------------------------------------------------------------- report
    def rel(p):
        try:
            return p.resolve().relative_to(ROOT)
        except ValueError:
            return p

    print(f"Eingabe: {rel(a.inp)}  |  Verzeichnis: {rel(a.csv)} ({len(directory)} Zeilen)  |  "
          f"PLZ für Log-Zeilen aus: {rel(a.plz_from)}")
    print("\n=== Log 'ungeklärt' → Hausverwaltungen ===")
    nb = sum(m["Recherchetiefe"] == B_LABEL for m in moved)
    print(f"verschoben: {len(moved)}  (B mit Verzeichnis-Treffer Name+PLZ: {nb}, C nur Name: {len(moved) - nb})")
    for m in moved:
        if m["_dir"]:
            print(f"   B: {m['Firma']} ({m['_plz']}) <= {m['_dir']['Quelle']}: {m['_dir']['Firma']}")
    if missing_plz:
        print(f"   ohne eindeutige PLZ aus v1 (nur C möglich): {missing_plz}")
    near = []
    for m in moved:
        if m["_dir"]:
            continue
        for d in directory:
            if d["PLZ"] == m["_plz"] and difflib.SequenceMatcher(
                    None, norm_name(m["Firma"]), norm_name(d["Firma"])).ratio() >= 0.85:
                near.append((m["Firma"], d["Firma"], d["PLZ"]))
    print(f"   C-Zeilen mit ähnlichem Verzeichnis-Namen, gleiche PLZ (nicht übernommen): "
          f"{near if near else 'keine'}")

    print("\n=== Duplikate (Verzeichniszeilen nicht angehängt) ===")
    for how, by in dup.items():
        detail = ", ".join(f"{o}: {n}" for o, n in sorted(by.items())) or "-"
        print(f"   via {how:12s} {sum(by.values()):4d}   ({detail})")
    for f_, plz, orig, origin in name_only_hits:
        print(f"      Name allein: '{f_}' ({plz}) = '{orig}' [{origin}]")
    print(f"   als B-Zusammenführung mit Log-Zeile verbraucht: {len(consumed)}")
    print(f"   neu angehängt (B):  {appended}")

    print("\n=== Zeilen je Tab ===")
    wb2 = openpyxl.load_workbook(a.out)
    for w in wb2.worksheets:
        h2 = header_map(w)
        key = "Firma" if "Firma" in h2 else next(iter(h2))
        n = sum(1 for r in range(2, w.max_row + 1) if w.cell(r, h2[key]).value)
        print(f"   {w.title:20s} {n}")
    print("=== Zeilen je Recherchetiefe (nur Tab Hausverwaltungen) ===")
    w = wb2[HV_TAB]
    h2 = header_map(w)
    tiers = {}
    for r in range(2, w.max_row + 1):
        if w.cell(r, h2["Firma"]).value:
            v = w.cell(r, h2["Recherchetiefe"]).value
            tiers[v] = tiers.get(v, 0) + 1
    for k, v in tiers.items():
        print(f"   {k:22s} {v}")
    print(f"\nGeschrieben: {rel(a.out)}")


if __name__ == "__main__":
    main()
