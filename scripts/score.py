#!/usr/bin/env python3
"""Score Hausverwaltungen rows and assign a Prio.

Usage: python3 scripts/score.py --in data/heidi_contacts_v5.xlsx --out data/heidi_scored_v6.xlsx
       [--set prio_a=5 --note "Grund"]   (changes a Schwelle in the Gewichte tab of OUT, with a dated note)

* Weights/thresholds live in the tab "Gewichte" (Schlüssel | Beschreibung | Schwelle | Punkte).
  The tab is created with defaults only if missing; existing values are read, so edits survive.
* Adds/recomputes Punkte, Prio, Prio-Grund on the Hausverwaltungen tab (derived columns, so
  they are recalculated on every run). Other tabs pass through unchanged.
* Einheiten Berlin overrides Einheiten. "unbekannt" counts as not filled.
* Prio: Enterprise if Einheiten >= enterprise threshold or Hinweis contains "bundesweit";
  else A >= prio_a, B >= prio_b, C below that if at least one scored field is filled,
  otherwise "nicht bewertbar". Anzahl Quellen does not count as "filled" (always present).
"""
import argparse
import re
import sys
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import DATA  # noqa: E402

HV = "Hausverwaltungen"
GW = "Gewichte"
DEFAULTS = [  # key, description, threshold, points
    ("einheiten_hoch", "Einheiten ≥ Schwelle UND URL in Einheiten-Quelle", 300, 3),
    ("einheiten_mittel", "Einheiten ≥ Schwelle (bis unter einheiten_hoch)", 100, 1),
    ("bewertungen_hoch", "Anzahl Bewertungen ≥ Schwelle", 200, 2),
    ("bewertungen_mittel", "Anzahl Bewertungen ≥ Schwelle (bis unter bewertungen_hoch)", 50, 1),
    ("verwaltungsart_beides", "Verwaltungsart beides", None, 2),
    ("verwaltungsart_einzeln", "Verwaltungsart WEG oder Miet", None, 1),
    ("kundenportal", "Kundenportal ja", None, 1),
    ("software", "Software genannt", None, 1),
    ("verband", "Mitgliedschaft VDIV/IVD oder VDIV-Logo ja", None, 1),
    ("quellen", "Anzahl Quellen ≥ Schwelle", 2, 1),
    ("fernablesung", "Fernablesung erwähnt ja", None, -1),
    ("gruppe", "Gruppe ausgefüllt", None, -1),
    ("website_fehler", "Website-Status Fehler", None, -2),
    ("prio_enterprise", "Enterprise ab Einheiten", 5000, None),
    ("enterprise_bundesweit", "'bundesweit' im Hinweis macht Enterprise nur ab Einheiten ≥ Schwelle "
     "(seit 2026-11-08; vorher ohne Einheiten-Bedingung, betraf eine Zeile ohne Einheiten-Angabe)", 5000, None),
    ("prio_a", "Prio A ab Punkten (Standard 5 seit 2026-11-08, vorher 6: A hatte bei 6 nur 10 Zeilen)", 5, None),
    ("prio_b", "Prio B ab Punkten (darunter C)", 3, None),
]


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def weights(wb):
    if GW not in wb.sheetnames:
        ws = wb.create_sheet(GW)
        ws.append(["Schlüssel", "Beschreibung", "Schwelle", "Punkte"])
        for row in DEFAULTS:
            ws.append(list(row))
        for c in ws[1]:
            c.font = Font(bold=True)
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 58
        created = True
    else:
        created = False
    ws = wb[GW]
    w = {}
    for key, _, thr, pts in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        if key:
            w[key] = (thr, pts)
    for row in DEFAULTS:  # rules added after the tab was created: append with defaults
        if row[0] not in w:
            ws.append(list(row))
            w[row[0]] = (row[2], row[3])
    return w, created


def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"\d{1,3}(?:\.\d{3})+|\d+", str(v))
    return float(m.group(0).replace(".", "")) if m else None


def fmt(n):
    return f"{int(n):,}".replace(",", ".")


def filled(v):
    return v is not None and str(v).strip() != "" and str(v).strip().lower() != "unbekannt"


def score_row(g, w):
    pts, parts = 0, []

    def add(p, label):
        nonlocal pts
        pts += p
        parts.append(f"{label} ({p:+d})")

    einh = num(g("Einheiten Berlin")) if filled(g("Einheiten Berlin")) else num(g("Einheiten"))
    src = "Einheiten Berlin" if filled(g("Einheiten Berlin")) else "Einheiten"
    has_url = "http" in str(g("Einheiten-Quelle") or "")
    if einh is not None:
        if einh >= w["einheiten_hoch"][0]:
            if has_url:
                add(w["einheiten_hoch"][1], f"{src} {fmt(einh)}")
            else:
                parts.append(f"{src} {fmt(einh)} ohne Quelle (+0)")
        elif einh >= w["einheiten_mittel"][0]:
            add(w["einheiten_mittel"][1], f"{src} {fmt(einh)}")
    bew = num(g("Anzahl Bewertungen"))
    if bew is not None:
        if bew >= w["bewertungen_hoch"][0]:
            add(w["bewertungen_hoch"][1], f"{fmt(bew)} Bewertungen")
        elif bew >= w["bewertungen_mittel"][0]:
            add(w["bewertungen_mittel"][1], f"{fmt(bew)} Bewertungen")
    va = str(g("Verwaltungsart") or "").strip().lower()
    if va.startswith("beides"):
        add(w["verwaltungsart_beides"][1], "beides")
    elif va.startswith("weg"):
        add(w["verwaltungsart_einzeln"][1], "WEG")
    elif va.startswith("miet"):
        add(w["verwaltungsart_einzeln"][1], "Miet")
    if str(g("Kundenportal") or "").lower().startswith("ja"):
        add(w["kundenportal"][1], "Portal")
    sw = str(g("Software") or "").strip()
    if filled(sw) and sw.lower() not in ("nein", "keine"):
        add(w["software"][1], f"Software {sw.replace('other: ', '')}")
    mg = str(g("Mitgliedschaft") or "")
    if re.search(r"VDIV|IVD", mg) or str(g("VDIV-Logo") or "").lower().startswith("ja"):
        add(w["verband"][1], "VDIV" if ("VDIV" in mg or str(g("VDIV-Logo") or "").lower().startswith("ja")) else "IVD")
    nq = num(g("Anzahl Quellen"))
    if nq is not None and nq >= w["quellen"][0]:
        add(w["quellen"][1], f"{int(nq)} Quellen")
    if str(g("Fernablesung erwähnt") or "").lower().startswith("ja"):
        add(w["fernablesung"][1], "Fernablesung")
    if filled(g("Gruppe")):
        add(w["gruppe"][1], "Gruppe")
    if str(g("Website-Status") or "").startswith("Fehler"):
        add(w["website_fehler"][1], "Website-Fehler")

    any_filled = any(filled(g(k)) for k in (
        "Einheiten", "Einheiten Berlin", "Anzahl Bewertungen", "Verwaltungsart", "Kundenportal", "Software",
        "Mitgliedschaft", "VDIV-Logo", "Fernablesung erwähnt", "Gruppe", "Website-Status"))
    reason = ", ".join(parts)
    bundesweit = "bundesweit" in str(g("Hinweis") or "").lower() and einh is not None \
        and einh >= w["enterprise_bundesweit"][0]
    if (einh is not None and einh >= w["prio_enterprise"][0]) or bundesweit:
        why = f"Einheiten {fmt(einh)} ≥ {fmt(w['prio_enterprise'][0])}" if einh and einh >= w["prio_enterprise"][0] \
            else "Hinweis 'bundesweit'"
        prio = "Enterprise"
        text = f"{reason + ' = ' if reason else ''}{pts} Punkte; {why} → Enterprise"
    else:
        if pts >= w["prio_a"][0]:
            prio = "A"
        elif pts >= w["prio_b"][0]:
            prio = "B"
        elif any_filled:
            prio = "C"
        else:
            prio = "nicht bewertbar"
        text = f"{reason + ' = ' if reason else ''}{pts} → {prio}" if any_filled or reason \
            else "keine bewertbaren Felder → nicht bewertbar"
    return pts, prio, text, einh


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, default=DATA / "heidi_contacts_v5.xlsx")
    ap.add_argument("--out", type=Path, default=DATA / "heidi_scored_v6.xlsx")
    ap.add_argument("--set", dest="sets", nargs="*", default=[], help="key=Schwelle, e.g. prio_a=5")
    ap.add_argument("--note", default="")
    a = ap.parse_args()
    if a.out.resolve() in (a.inp.resolve(), (DATA / "heidi_merged_v3.xlsx").resolve()):
        sys.exit(f"Refusing to write {a.out}")
    wb = openpyxl.load_workbook(a.inp)
    w, created = weights(wb)
    if a.sets:
        from datetime import date
        gw = wb[GW]
        for kv in a.sets:
            key, val = kv.split("=", 1)
            for row in gw.iter_rows(min_row=2):
                if row[0].value == key:
                    old = row[2].value
                    row[2].value = float(val) if "." in val else int(val)
                    row[1].value = f"{row[1].value} – geändert {old} → {val} am {date.today().isoformat()}" \
                                   f"{': ' + a.note if a.note else ''}"
                    print(f"Gewichte: {key} Schwelle {old} → {val}")
        w, _ = weights(wb)
    ws = wb[HV]
    hm = header_map(ws)
    tmpl = ws.cell(1, ws.max_column)
    for n in ("Punkte", "Prio", "Prio-Grund"):
        if n not in hm:
            c = ws.cell(1, ws.max_column + 1, n)
            c.font, c.fill, c.alignment = copy(tmpl.font), copy(tmpl.fill), copy(tmpl.alignment)
            ws.column_dimensions[c.column_letter].width = 60 if n == "Prio-Grund" else 12
            hm[n] = c.column
    counts, listing, no_url = {}, [], 0
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, hm["Firma"]).value:
            continue
        g = lambda k, r=r: ws.cell(r, hm[k]).value if k in hm else None  # noqa: E731
        pts, prio, text, einh = score_row(g, w)
        ws.cell(r, hm["Punkte"]).value = pts
        ws.cell(r, hm["Prio"]).value = prio
        ws.cell(r, hm["Prio-Grund"]).value = text
        counts[prio] = counts.get(prio, 0) + 1
        no_url += "ohne Quelle" in text
        if prio in ("A", "Enterprise"):
            listing.append((prio, g("Firma"), fmt(einh) if einh is not None else "", g("Anzahl Bewertungen") or "",
                            g("Verwaltungsart") or "", g("Mitgliedschaft") or "", g("Ansprechpartner") or "", pts))
    wb.save(a.out)
    print(f"Gewichte: {'Tab neu angelegt (Standardwerte)' if created else 'aus bestehendem Tab gelesen'}")
    order = ["Enterprise", "A", "B", "C", "nicht bewertbar"]
    print("Zeilen je Prio:", {k: counts.get(k, 0) for k in order})
    print(f"Einheiten ≥ {fmt(w['einheiten_hoch'][0])} ohne URL (0 Punkte): {no_url}")
    for p in ("Enterprise", "A"):
        rows = sorted([x for x in listing if x[0] == p], key=lambda x: -x[7])
        print(f"\n=== {p} ({len(rows)}) ===")
        print("Firma | Einheiten | Anzahl Bewertungen | Verwaltungsart | Mitgliedschaft | Ansprechpartner | Punkte")
        for x in rows:
            print(" | ".join(str(v) for v in x[1:]))
    print(f"\nGeschrieben: {a.out.name}")


if __name__ == "__main__":
    main()
