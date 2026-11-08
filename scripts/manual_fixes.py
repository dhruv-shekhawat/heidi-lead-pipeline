#!/usr/bin/env python3
"""Apply manual, case-by-case decisions from scripts/manual_fixes.json (rows found by Website domain).

* einzelobjekt: Einheiten / Einheiten-Quelle are moved to the Log ("verworfen: Einzelobjekt"),
  Recherchetiefe goes back to "B+ (automatisch)".
* zu_wohnungsunternehmen: the row moves from Hausverwaltungen to Wohnungsunternehmen (Segment as
  configured, Einheiten + Einheiten-Quelle kept, Einheiten (Art) set); fields the Wohnungsunternehmen
  tab has no column for (Ansprechpartner, Telefon, Straße, …) are kept in its Hinweis.
Idempotent: a firm already fixed is skipped. Run before score.py.

Usage: python3 scripts/manual_fixes.py --in IN.xlsx --out OUT.xlsx [--config scripts/manual_fixes.json]
"""
import argparse
import json
import sys
from copy import copy
from datetime import date
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import DATA, norm_domain  # noqa: E402

HV, WU, LOG = "Hausverwaltungen", "Wohnungsunternehmen", "Log"
EXTRA_TO_HINWEIS = ["Ansprechpartner", "Funktion", "E-Mail", "Telefon", "Straße", "Recherchetiefe"]


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def find_row(ws, hm, domain):
    for r in range(2, ws.max_row + 1):
        if norm_domain(ws.cell(r, hm["Website"]).value) == norm_domain(domain):
            return r
    return None


def append_row(ws, values, hm):
    style = ws.max_row
    r = ws.max_row + 1
    for c in range(1, ws.max_column + 1):
        if ws.cell(style, c).has_style:
            ws.cell(r, c)._style = copy(ws.cell(style, c)._style)
    for k, v in values.items():
        if k in hm and v not in (None, ""):
            ws.cell(r, hm[k]).value = v
    return r


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--config", type=Path, default=Path(__file__).resolve().parent / "manual_fixes.json")
    a = ap.parse_args()
    if a.out.resolve() in (a.inp.resolve(), (DATA / "heidi_merged_v3.xlsx").resolve()):
        sys.exit(f"Refusing to write {a.out}")
    cfg = json.loads(a.config.read_text(encoding="utf-8"))
    wb = openpyxl.load_workbook(a.inp)
    hv, wu, log = wb[HV], wb[WU], wb[LOG]
    hm, wm, lm = header_map(hv), header_map(wu), header_map(log)
    today = date.today().isoformat()

    for fx in cfg.get("einzelobjekt", []):
        r = find_row(hv, hm, fx["website"])
        if not r or hv.cell(r, hm["Einheiten"]).value in (None, ""):
            print(f"einzelobjekt {fx['website']}: übersprungen (nicht gefunden oder keine Einheiten)")
            continue
        firma, einh = hv.cell(r, hm["Firma"]).value, hv.cell(r, hm["Einheiten"]).value
        quelle = hv.cell(r, hm["Einheiten-Quelle"]).value
        append_row(log, {"Bereich": "Einheiten verworfen", "Firma": firma, "Website": hv.cell(r, hm["Website"]).value,
                         "Typ": hv.cell(r, hm["Typ"]).value,
                         "Grund / Hinweis": f"verworfen: Einzelobjekt – Einheiten {einh}; {fx.get('grund', '')}; "
                                            f"Quelle: {quelle}",
                         "Quelle": str(quelle or "").split(" – ")[0]}, lm)
        hv.cell(r, hm["Einheiten"]).value = None
        hv.cell(r, hm["Einheiten-Quelle"]).value = None
        hv.cell(r, hm["Recherchetiefe"]).value = "B+ (automatisch)"
        print(f"einzelobjekt: {firma} – Einheiten {einh} → Log, Recherchetiefe B+ (automatisch)")

    for fx in cfg.get("zu_wohnungsunternehmen", []):
        r = find_row(hv, hm, fx["website"])
        if not r:
            print(f"zu_wohnungsunternehmen {fx['website']}: übersprungen (nicht in Hausverwaltungen)")
            continue
        rec = {k: hv.cell(r, c).value for k, c in hm.items()}
        extra = "; ".join(f"{k}: {rec[k]}" for k in EXTRA_TO_HINWEIS if rec.get(k) not in (None, ""))
        hinweis = " | ".join(x for x in (rec.get("Hinweis"), f"{fx.get('grund', 'Bestandshalter')} – aus "
                                         f"Hausverwaltungen verschoben {today}", extra) if x)
        vals = {k: rec.get(k) for k in wm if k in rec}
        vals.update({"Segment": fx.get("segment", "sonstige"), "Typ": "Wohnungsunternehmen",
                     "Einheiten (Art)": rec.get("Einheiten (Art)") or fx.get("einheiten_art"),
                     "Hinweis": hinweis, "Stand": today})
        append_row(wu, vals, wm)
        hv.delete_rows(r)
        print(f"zu_wohnungsunternehmen: {rec['Firma']} (Einheiten {rec.get('Einheiten')}) → {WU}, "
              f"Segment {vals['Segment']}")

    wb.save(a.out)
    print(f"Geschrieben: {a.out.name}")


if __name__ == "__main__":
    main()
