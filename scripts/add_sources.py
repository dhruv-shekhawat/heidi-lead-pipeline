#!/usr/bin/env python3
"""Add Gelbe Seiten detail data and JACASA data to the merged workbook.

Inputs : data/heidi_merged_v3.xlsx (read only), data/directory_detail.csv,
         data/jacasa.csv (optional; missing file = no JACASA rows)
Output : data/heidi_merged_v3c.xlsx

* Only the Hausverwaltungen tab is changed; other tabs pass through unchanged.
* Matching per source row: normalised domain, else normalised name + PLZ
  (Firma or Name alt; PLZ parsed from Bezirk/PLZ).
* Matched rows: fill EMPTY Website / Telefon / Straße only; JACASA columns
  Mitgliedschaft, Mitarbeiter, Aktiv seit, Google-Bewertung, Anzahl Bewertungen,
  JACASA-Spezialisierung, JACASA-URL; Verwaltungsart (if empty/unbekannt) and
  VDIV-Logo (if empty) derived from JACASA.
* Unmatched JACASA rows -> new rows (Typ Hausverwaltung, Recherchetiefe "B (Verzeichnis)",
  Quelle "JACASA"), unless they match a Log or Wohnungsunternehmen row.
* Anzahl Quellen = number of distinct entries in Quelle (split on ; and ,).

Usage: python3 scripts/add_sources.py [--in …] [--detail …] [--jacasa …] [--out …]
"""
import argparse
import csv
import re
import sys
from copy import copy
from datetime import date
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import DATA, ROOT, is_company_domain, norm_domain, norm_name, plz_of  # noqa: E402

HV, WU, LOG = "Hausverwaltungen", "Wohnungsunternehmen", "Log"
JACASA_COLS = ["Mitgliedschaft", "Mitarbeiter", "Aktiv seit", "Google-Bewertung",
               "Anzahl Bewertungen", "JACASA-Spezialisierung", "JACASA-URL"]
NEW_COLS = JACASA_COLS + ["Anzahl Quellen"]
PROTECTED = [DATA / "heidi_merged_v3.xlsx"]


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def ensure_cols(ws, names):
    hm = header_map(ws)
    tmpl = ws.cell(1, ws.max_column)
    for n in names:
        if n not in hm:
            c = ws.cell(1, ws.max_column + 1, n)
            c.font, c.fill, c.alignment, c.border = copy(tmpl.font), copy(tmpl.fill), copy(tmpl.alignment), copy(tmpl.border)
            ws.column_dimensions[c.column_letter].width = 20
            hm[n] = c.column
    return hm


def empty(v):
    return v is None or str(v).strip() == ""


def read_csv(p):
    if not Path(p).exists():
        return None
    with open(p, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


GENERIC_WORDS = {"hausverwaltung", "hausverwaltungen", "immobilien", "immobilienverwaltung", "verwaltung",
                 "verwaltungs", "berlin", "gmbh", "management", "immobilienmanagement", "grundstuecksverwaltung",
                 "gesellschaft", "service", "services", "property", "haus", "real", "estate", "wohnungsverwaltung",
                 "objektverwaltung", "hausverwalter", "group", "gruppe", "holding", "partner", "consulting"}


def words4(name):
    """4+ letter words, lowercase, umlauts folded, Rechtsform tokens dropped."""
    s = (name or "").lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        s = s.replace(a, b)
    return set(re.findall(r"[a-z]{4,}", s)) - {"gmbh"}


COMMON = set(GENERIC_WORDS)  # extended at runtime with words occurring in > 3 firm names


def subset_ok(w, w2):
    short, long_ = (w, w2) if len(w) <= len(w2) else (w2, w)
    return bool(short) and bool(short - COMMON) and short <= long_


class Index:
    def __init__(self):
        self.dom, self.np, self.name, self.by_plz, self.no_plz = {}, {}, {}, {}, []

    def add(self, row, names, website, plz):
        d = norm_domain(website)
        if is_company_domain(d):
            self.dom.setdefault(d, row)
        for n in names:
            if n:
                if plz:
                    self.np.setdefault(f"{norm_name(n)}|{plz}", row)
                    self.by_plz.setdefault(plz, []).append((words4(n), row))
                else:
                    self.no_plz.append((norm_name(n), words4(n), row))
                self.name.setdefault(norm_name(n), row)

    def match(self, website, name, plz, name_only=False, subset=False):
        d = norm_domain(website)
        if is_company_domain(d) and d in self.dom:
            return "Domain", self.dom[d]
        if plz and f"{norm_name(name)}|{plz}" in self.np:
            return "Name+PLZ", self.np[f"{norm_name(name)}|{plz}"]
        if subset and plz:
            # every 4+ letter word of the shorter name occurs in the longer one, same PLZ
            # (shorter name must contain at least one non-generic word)
            w = words4(name)
            hits = [row for w2, row in self.by_plz.get(plz, []) if subset_ok(w, w2)]
            if hits:
                return "Wörter+PLZ", hits[0]
            # existing rows WITHOUT a PLZ: same name or same distinctive words (merge.py rule)
            nn = norm_name(name)
            for n2, w2, row in self.no_plz:
                if nn == n2 or subset_ok(w, w2):
                    return "Name (Bestand ohne PLZ)", row
        if name_only and norm_name(name) in self.name:
            return "Name", self.name[norm_name(name)]
        return None, None


def verwaltungsart_from(spez):
    s = (spez or "")
    miet = bool(re.search(r"\bMiet\b|Mietobjekt", s))
    weg = bool(re.search(r"\bWEG\b|Wohnungseigentum", s))
    return "beides (JACASA)" if miet and weg else "Miet (JACASA)" if miet else "WEG (JACASA)" if weg else None


BERLIN_RX = re.compile(r"Berlin|Mitte|Charlottenburg|Wilmersdorf|Friedrichshain|Kreuzberg|Pankow|Prenzlauer|"
                       r"Weißensee|Spandau|Steglitz|Zehlendorf|Tempelhof|Schöneberg|Neukölln|Treptow|Köpenick|"
                       r"Marzahn|Hellersdorf|Lichtenberg|Reinickendorf|Moabit|Wedding|Tegel|Lichterfelde|"
                       r"Dahlem|Grunewald|Friedenau|Schmargendorf|Britz|Rudow|Adlershof|Hohenschönhausen", re.I)


SKIP_ON_MERGE = {"Punkte", "Prio", "Prio-Grund", "Anzahl Quellen", "Stand", "Batch", "Hinweis"}
LIST_FIELDS = {"Quelle", "Quellen"}


def merge_same_name_rows(ws, log=print):
    """Merge Hausverwaltungen rows with the same normalised Firma + PLZ when they have at most one
    distinct website (typically: light_enrich renamed a directory row to the name of an existing row).
    Target = the row with the website. Empty target cells are filled; Quelle/Quellen are unioned;
    any other differing value goes into the target's Hinweis. The other rows are deleted."""
    hm = {c.value: c.column for c in ws[1] if c.value}
    groups = {}
    for r in range(2, ws.max_row + 1):
        f = ws.cell(r, hm["Firma"]).value
        if f:
            groups.setdefault(f"{norm_name(f)}|{plz_of(ws.cell(r, hm['Bezirk/PLZ']).value)}", []).append(r)
    to_delete, merged = [], []
    for key, rows in groups.items():
        if len(rows) < 2:
            continue
        doms = {norm_domain(ws.cell(r, hm["Website"]).value) for r in rows} - {""}
        if len(doms) > 1:
            log(f"   gleiche Firma+PLZ, verschiedene Websites – nicht zusammengeführt: {key} {sorted(doms)}")
            continue
        target = next((r for r in rows if ws.cell(r, hm["Website"]).value), rows[0])
        for o in (r for r in rows if r != target):
            notes = []
            for col, ci in hm.items():
                if col in SKIP_ON_MERGE:
                    continue
                tv, ov = ws.cell(target, ci).value, ws.cell(o, ci).value
                if ov in (None, "") or tv == ov or str(ov).strip().lower() == "unbekannt":
                    continue
                if tv in (None, ""):
                    ws.cell(target, ci).value = ov
                elif col in LIST_FIELDS:
                    parts = [x.strip() for x in re.split(r";", f"{tv};{ov}") if x.strip()]
                    ws.cell(target, ci).value = "; ".join(dict.fromkeys(parts))
                else:
                    notes.append(f"{col}: {ov}")
            oh = ws.cell(o, hm["Hinweis"]).value
            if oh and oh != ws.cell(target, hm["Hinweis"]).value:
                notes.append(f"Hinweis: {oh}")
            note = "Dublette zusammengeführt (Zeile ohne Website)" + (": " + "; ".join(notes) if notes else "")
            h = ws.cell(target, hm["Hinweis"]).value
            ws.cell(target, hm["Hinweis"]).value = f"{h} | {note}" if h else note
            to_delete.append(o)
            merged.append((ws.cell(target, hm["Firma"]).value, key))
        if "Anzahl Quellen" in hm:  # derived: recompute after Quelle was unioned
            q = ws.cell(target, hm["Quelle"]).value
            n = n_sources(q)
            if any(ws.cell(target, hm[k]).value not in (None, "") for k in JACASA_COLS if k in hm) \
                    and "jacasa" not in str(q or "").lower():
                n += 1
            ws.cell(target, hm["Anzahl Quellen"]).value = n
    for r in sorted(to_delete, reverse=True):
        ws.delete_rows(r)
    for f, k in merged:
        log(f"   Dublette zusammengeführt: {f} ({k.split('|')[1]})")
    return merged


def n_sources(quelle):
    return len({x.strip().lower() for x in re.split(r"[;,]", str(quelle or "")) if x.strip()})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, default=DATA / "heidi_merged_v3.xlsx")
    ap.add_argument("--detail", type=Path, default=DATA / "directory_detail.csv")
    ap.add_argument("--jacasa", type=Path, default=DATA / "jacasa.csv")
    ap.add_argument("--out", type=Path, default=DATA / "heidi_merged_v3c.xlsx")
    a = ap.parse_args()
    if a.out.resolve() in [a.inp.resolve()] + [p.resolve() for p in PROTECTED]:
        sys.exit(f"Refusing to write {a.out}")

    wb = openpyxl.load_workbook(a.inp)
    ws = wb[HV]
    print("Dubletten-Prüfung (gleiche Firma + PLZ):")
    merge_same_name_rows(ws)
    hm = ensure_cols(ws, NEW_COLS)
    style_row = ws.max_row
    cell = lambda r, k: ws.cell(r, hm[k])  # noqa: E731

    idx = Index()
    for r in range(2, ws.max_row + 1):
        if cell(r, "Firma").value:
            idx.add(r, [cell(r, "Firma").value, cell(r, "Name alt").value], cell(r, "Website").value,
                    plz_of(cell(r, "Bezirk/PLZ").value))
    other = Index()  # Log + Wohnungsunternehmen: only used to block new JACASA rows
    for tab in (LOG, WU):
        w = wb[tab]
        h = header_map(w)
        for r in range(2, w.max_row + 1):
            if tab == LOG and str(w.cell(r, h["Bereich"]).value or "").startswith("Einheiten"):
                continue  # rejected Einheiten candidates of existing firms, not excluded firms
            if w.cell(r, h["Firma"]).value:
                other.add(f"{tab}:{r}", [w.cell(r, h["Firma"]).value, w.cell(r, h["Name alt"]).value],
                          w.cell(r, h["Website"]).value,
                          plz_of(w.cell(r, h["Bezirk/PLZ"]).value) if "Bezirk/PLZ" in h else "")

    fills = {k: 0 for k in ["Website", "Telefon", "Straße", "Verwaltungsart", "VDIV-Logo"] + JACASA_COLS}

    def fill(r, k, v):
        if not empty(v) and empty(cell(r, k).value):
            cell(r, k).value = v
            fills[k] += 1
            return True
        return False

    stats = {}
    hidden_dups = []

    # ---------------------------------------------------- Gelbe Seiten details
    detail = read_csv(a.detail)
    if detail is None:
        print(f"WARNUNG: {a.detail} fehlt – keine Detaildaten")
        detail = []
    det_stats = {"Domain": 0, "Name+PLZ": 0, None: 0}
    for d in detail:
        if not d.get("Abruf", "").startswith("ok"):
            continue
        how, r = idx.match(d["Website"], d["Firma"], d["PLZ"])
        det_stats[how] += 1
        if not r:
            continue
        if how == "Domain" and norm_name(d["Firma"]) != norm_name(cell(r, "Firma").value):
            # GS listing whose website belongs to another row: duplicate hidden by name
            own = idx.np.get(f"{norm_name(d['Firma'])}|{d['PLZ']}")
            if own and own != r:
                hidden_dups.append((ws.cell(own, hm["Firma"]).value, d["PLZ"], cell(r, "Firma").value,
                                    norm_domain(d["Website"]), cell(r, "Recherchetiefe").value))
        dom = norm_domain(d["Website"])
        if is_company_domain(dom) and fill(r, "Website", dom):
            idx.dom.setdefault(dom, r)
        fill(r, "Telefon", d["Telefon"])
        fill(r, "Straße", d["Straße"])
    stats["Gelbe Seiten Details"] = det_stats

    # ---------------------------------------------------------------- JACASA
    jac = read_csv(a.jacasa)
    jac_stats = {"Domain": 0, "Name+PLZ": 0, "Wörter+PLZ": 0, "Name (Bestand ohne PLZ)": 0, "neu": 0,
                 "Log/WU-Treffer": 0}
    if jac:
        from collections import Counter
        freq = Counter(w for nm in [cell(r, "Firma").value for r in range(2, ws.max_row + 1)] +
                       [x.get("Firma") for x in jac] for w in words4(nm))
        COMMON.update(w for w, c in freq.items() if c > 3)
    existing_rows = [r for r in range(2, ws.max_row + 1) if cell(r, "Firma").value]
    possible_dups, exact_other_plz = [], []
    jac_stats["Name exakt (andere PLZ)"] = 0
    today = date.today().isoformat()
    if jac is None:
        print(f"Hinweis: {a.jacasa} fehlt – JACASA-Spalten bleiben leer")
        jac = []
    jac_rows, subset_hits, outside = set(), [], []
    g = lambda j, k: (j.get(k) if j.get(k) is not None else "")  # noqa: E731  tolerate missing columns
    for j in jac:
        firma, plz = g(j, "Firma"), str(g(j, "PLZ")).strip()
        how, r = idx.match(g(j, "Website"), firma, plz, subset=True)
        if not r:
            # same firm under another PLZ: merge only if the normalised name is IDENTICAL
            nn = norm_name(firma)
            for r0 in existing_rows:
                if nn and nn in (norm_name(cell(r0, "Firma").value), norm_name(cell(r0, "Name alt").value)):
                    p0 = plz_of(cell(r0, "Bezirk/PLZ").value)
                    note = f"JACASA-Eintrag mit PLZ {plz} (Bestand PLZ {p0 or '–'}) zusammengeführt"
                    h = cell(r0, "Hinweis").value
                    cell(r0, "Hinweis").value = f"{h} | {note}" if h else note
                    how, r = "Name exakt (andere PLZ)", r0
                    exact_other_plz.append((firma, plz, cell(r0, "Firma").value, p0))
                    break
        if not r:
            how2, _ = other.match(g(j, "Website"), firma, plz, name_only=True)
            if how2:
                jac_stats["Log/WU-Treffer"] += 1
                continue
            r = ws.max_row + 1
            for col in range(1, ws.max_column + 1):
                if ws.cell(style_row, col).has_style:
                    ws.cell(r, col)._style = copy(ws.cell(style_row, col)._style)
            for k, v in {"Firma": firma, "Bezirk/PLZ": plz, "Quelle": "JACASA",
                         "Typ": "Hausverwaltung", "Recherchetiefe": "B (Verzeichnis)",
                         "Quellen": g(j, "Profil-URL"), "Stand": today}.items():
                cell(r, k).value = v or None
            idx.add(r, [firma], g(j, "Website"), plz)
            jac_stats["neu"] += 1
            w = words4(firma)
            for r0 in existing_rows:  # same firm under another PLZ? flag, don't merge (rule: PLZ equal)
                if subset_ok(w, words4(cell(r0, "Firma").value) | words4(cell(r0, "Name alt").value)):
                    p0 = plz_of(cell(r0, "Bezirk/PLZ").value)
                    note = f"mögliches Duplikat von '{cell(r0, 'Firma').value}' (PLZ {p0 or '–'})"
                    cell(r, "Hinweis").value = note
                    possible_dups.append((firma, plz, cell(r0, "Firma").value, p0))
                    break
        else:
            jac_stats[how] += 1
            if how in ("Wörter+PLZ", "Name (Bestand ohne PLZ)"):
                subset_hits.append((how, firma, cell(r, "Firma").value, plz))
        jac_rows.add(r)
        dom = norm_domain(g(j, "Website"))
        if is_company_domain(dom):
            fill(r, "Website", dom)
        fill(r, "Telefon", g(j, "Telefon"))
        fill(r, "Straße", g(j, "Straße"))
        for k, src, typ in (("Mitgliedschaft", "Mitgliedschaft", str), ("Mitarbeiter", "Mitarbeiter", int),
                            ("Aktiv seit", "Aktiv seit", str), ("Google-Bewertung", "Bewertung", float),
                            ("Anzahl Bewertungen", "Anzahl Bewertungen", int),
                            ("JACASA-Spezialisierung", "Spezialisierung", str), ("JACASA-URL", "Profil-URL", str)):
            v = g(j, src)
            if v not in ("", None) and typ is not str:
                try:
                    v = typ(float(str(v).replace(",", ".")))
                except ValueError:
                    pass
            fill(r, k, v)
        va = cell(r, "Verwaltungsart").value
        if empty(va) or str(va).strip().lower() == "unbekannt":
            v = verwaltungsart_from(g(j, "Spezialisierung"))
            if v:
                cell(r, "Verwaltungsart").value = v
                fills["Verwaltungsart"] += 1
        if "VDIV" in g(j, "Mitgliedschaft"):
            fill(r, "VDIV-Logo", "ja (JACASA)")
        taetig = g(j, "Tätig in")
        if taetig and not BERLIN_RX.search(taetig):
            note = f"Tätig in: {taetig}"
            h = cell(r, "Hinweis").value
            if note not in str(h or ""):
                cell(r, "Hinweis").value = f"{h} | {note}" if h else note
                outside.append(firma)
    stats["JACASA"] = jac_stats

    # ---------------------------------------------------------- Anzahl Quellen
    # derived: distinct entries in Quelle (+1 for JACASA when the row carries JACASA data and
    # "JACASA" is not already in Quelle); Quelle itself is never modified
    for r in range(2, ws.max_row + 1):
        if not cell(r, "Firma").value:
            continue
        q = cell(r, "Quelle").value
        n = n_sources(q)
        has_jac = r in jac_rows or any(not empty(cell(r, k).value) for k in JACASA_COLS)
        if has_jac and "jacasa" not in str(q or "").lower():
            n += 1
        cell(r, "Anzahl Quellen").value = n

    wb.save(a.out)

    # ------------------------------------------------------------------ report
    print(f"Eingabe {a.inp.name} → Ausgabe {a.out.relative_to(ROOT) if a.out.is_absolute() else a.out}")
    d = stats["Gelbe Seiten Details"]
    print(f"Gelbe-Seiten-Details ({sum(d.values())} ok): Treffer via Domain {d['Domain']}, "
          f"via Name+PLZ {d['Name+PLZ']}, ohne Treffer {d[None]}")
    j = stats["JACASA"]
    print(f"JACASA ({len(jac)} Zeilen): via Domain {j['Domain']}, via Name+PLZ {j['Name+PLZ']}, "
          f"via Wörter+PLZ {j['Wörter+PLZ']}, via Name (Bestand ohne PLZ) {j['Name (Bestand ohne PLZ)']} → "
          f"Name exakt/andere PLZ {j['Name exakt (andere PLZ)']} → zugeordnet "
          f"{j['Domain'] + j['Name+PLZ'] + j['Wörter+PLZ'] + j['Name (Bestand ohne PLZ)'] + j['Name exakt (andere PLZ)']}; "
          f"neue Zeilen {j['neu']}; wegen Log/WU-Treffer nicht angelegt {j['Log/WU-Treffer']}")
    for x in subset_hits:
        print(f"   {x[0]}: '{x[1]}' ≙ '{x[2]}' ({x[3]})")
    for x in exact_other_plz:
        print(f"   Name exakt, andere PLZ → zusammengeführt: '{x[0]}' ({x[1]}) → '{x[2]}' ({x[3] or '–'})")
    print(f"   neue Zeilen mit Hinweis 'mögliches Duplikat' (andere PLZ, nicht zusammengeführt): {len(possible_dups)}")
    for x in possible_dups:
        print(f"      '{x[0]}' ({x[1]}) ~ '{x[2]}' ({x[3] or '–'})")
    print(f"   'Tätig in' ohne Berlin-Ort → Hinweis: {len(outside)} {outside}")
    print("Befüllt (nur leere Zellen):", {k: v for k, v in fills.items() if v})
    print(f"Versteckte Duplikate (GS-Eintrag mit Website einer anderen Zeile, nicht gelöscht): {len(hidden_dups)}")
    for x in hidden_dups:
        print(f"   '{x[0]}' ({x[1]}) ≙ '{x[2]}' [{x[4]}] über {x[3]}")
    tiers, nq = {}, {}
    for r in range(2, ws.max_row + 1):
        if cell(r, "Firma").value:
            t = cell(r, "Recherchetiefe").value
            tiers[t] = tiers.get(t, 0) + 1
            q = cell(r, "Anzahl Quellen").value
            nq[q] = nq.get(q, 0) + 1
    print("Zeilen je Recherchetiefe:", tiers)
    print("Anzahl Quellen (aus Spalte Quelle):", dict(sorted(nq.items())))


if __name__ == "__main__":
    main()
