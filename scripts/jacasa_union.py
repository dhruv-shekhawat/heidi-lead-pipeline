#!/usr/bin/env python3
"""Find all JACASA CSV exports, normalise both layouts and union them -> data/jacasa_all.csv.

Searched: data/*jacasa*.csv and ~/Downloads/jacasa*.csv (jacasa.csv, jacasa_1.csv, jacasa (1).csv …);
data/jacasa_all.csv itself is ignored. Identical files (same bytes) are read once.
Layout 1: Firma, Straße, PLZ, Bewertung, Anzahl Bewertungen, Mitgliedschaft, Mitarbeiter, Jahre aktiv, Tätig in
Layout 2: Firma, Bewertung, Anzahl_Bewertungen, Strasse, PLZ, Ort, Mitgliedschaft, Taetig_in,
          Spezialisiert_auf, URL, Gefunden_in_Sortierung
Normalisation: headers unified; "Mitglied im Verband der Immobilienverwalter (VDIV)" -> VDIV,
"Mitglied im Immobilienverband Deutschland (IVD)" -> IVD; Spezialisiert_auf Mietobjekte -> Miet,
Gewerbeobjekte -> Gewerbe; comma decimals -> numbers. Union key = normalised Firma + PLZ; later
files only fill empty fields of an existing row.

Usage: python3 scripts/jacasa_union.py [--extra PATH …]
"""
import argparse
import csv
import hashlib
import io
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import DATA, ROOT, norm_name  # noqa: E402

OUT = DATA / "jacasa_all.csv"
FIELDS = ["Firma", "Straße", "PLZ", "Ort", "Bewertung", "Anzahl Bewertungen", "Mitgliedschaft", "Mitarbeiter",
          "Aktiv seit", "Tätig in", "Spezialisierung", "Profil-URL", "Gefunden_in_Sortierung", "Quelldatei"]
HEADER_MAP = {
    "firma": "Firma", "straße": "Straße", "strasse": "Straße", "plz": "PLZ", "ort": "Ort",
    "bewertung": "Bewertung", "anzahl bewertungen": "Anzahl Bewertungen", "anzahl_bewertungen": "Anzahl Bewertungen",
    "mitgliedschaft": "Mitgliedschaft", "mitarbeiter": "Mitarbeiter", "jahre aktiv": "Aktiv seit",
    "jahre_aktiv": "Aktiv seit", "aktiv seit": "Aktiv seit", "tätig in": "Tätig in", "taetig_in": "Tätig in",
    "tätig_in": "Tätig in", "spezialisiert_auf": "Spezialisierung", "spezialisierung": "Spezialisierung",
    "url": "Profil-URL", "profil-url": "Profil-URL", "gefunden_in_sortierung": "Gefunden_in_Sortierung",
}
MEMBER = [(re.compile(r"Verband der Immobilienverwalter|\bVDIV\b", re.I), "VDIV"),
          (re.compile(r"Immobilienverband Deutschland|\bIVD\b", re.I), "IVD"),
          (re.compile(r"\bBVI\b|Bundesfachverband der Immobilienverwalter", re.I), "BVI")]
SPEZ = {"mietobjekte": "Miet", "gewerbeobjekte": "Gewerbe"}


def find_files(extra):
    cands = sorted(DATA.glob("*jacasa*.csv")) + sorted((Path.home() / "Downloads").glob("jacasa*.csv")) + \
        [Path(p) for p in extra]
    out, seen = [], set()
    for p in cands:
        if p.name == OUT.name or not p.is_file():
            continue
        h = hashlib.sha1(p.read_bytes()).hexdigest()
        if h in seen:
            print(f"  (identische Kopie übersprungen: {p})")
            continue
        seen.add(h)
        out.append(p)
    return out


def read_any(p):
    raw = p.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(txt[:4000], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.DictReader(io.StringIO(txt), dialect=dialect))
    return rows, list(rows[0].keys()) if rows else []


def to_num(v, as_int=False):
    v = (v or "").strip()
    if not v:
        return ""
    v = v.replace(".", "").replace(",", ".") if re.fullmatch(r"\d{1,3}(\.\d{3})*(,\d+)?|\d+,\d+", v) else v
    try:
        f = float(v)
    except ValueError:
        return v
    return int(f) if as_int or f.is_integer() and as_int else f


def normalise(rec, src):
    out = {k: "" for k in FIELDS}
    for k, v in rec.items():
        key = HEADER_MAP.get((k or "").strip().lower())
        if key:
            out[key] = (v or "").strip()
    m = [tag for rx, tag in MEMBER if rx.search(out["Mitgliedschaft"])]
    out["Mitgliedschaft"] = ", ".join(m) if m else out["Mitgliedschaft"]
    if out["Spezialisierung"]:
        parts = [p.strip() for p in out["Spezialisierung"].split(",") if p.strip()]
        out["Spezialisierung"] = ", ".join(SPEZ.get(p.lower(), p) for p in parts)
    out["Bewertung"] = to_num(out["Bewertung"])
    out["Anzahl Bewertungen"] = to_num(out["Anzahl Bewertungen"], as_int=True)
    out["Mitarbeiter"] = to_num(out["Mitarbeiter"], as_int=True)
    if re.fullmatch(r"\d+", str(out["Aktiv seit"])):
        out["Aktiv seit"] = f"{out['Aktiv seit']} Jahre"
    out["Quelldatei"] = src
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--extra", nargs="*", default=[])
    a = ap.parse_args()
    files = find_files(a.extra)
    print("Gefundene JACASA-Dateien:")
    union, order, total = {}, [], 0
    for p in files:
        rows, hdr = read_any(p)
        layout = "Layout 2" if "Spezialisiert_auf" in hdr or "Taetig_in" in hdr else \
            "Layout 1" if "Jahre aktiv" in hdr or "Tätig in" in hdr else "unbekannt"
        print(f"  {p} | {len(rows)} Zeilen | {layout}\n     Header: {hdr}")
        total += len(rows)
        for rec in rows:
            n = normalise(rec, p.name)
            if not n["Firma"]:
                continue
            k = f"{norm_name(n['Firma'])}|{n['PLZ']}"
            if k in union:
                for f, v in n.items():
                    if v not in ("", None) and union[k][f] in ("", None):
                        union[k][f] = v
                if p.name not in union[k]["Quelldatei"]:
                    union[k]["Quelldatei"] += f"; {p.name}"
            else:
                union[k] = n
                order.append(k)
    with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(union[k] for k in order)
    both = sum(";" in union[k]["Quelldatei"] for k in order)
    print(f"Zeilen gesamt {total} → eindeutig (Firma+PLZ) {len(order)} (in beiden Dateien: {both})")
    print(f"  mit Mitgliedschaft {sum(bool(union[k]['Mitgliedschaft']) for k in order)}, "
          f"mit Spezialisierung {sum(bool(union[k]['Spezialisierung']) for k in order)}, "
          f"mit Bewertung {sum(union[k]['Bewertung'] != '' for k in order)}")
    print(f"Geschrieben: {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
