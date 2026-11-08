#!/usr/bin/env python3
"""Export the scored workbook as a clean deliverable: data/heidi_final.xlsx.

Tabs (in this order): "Leads Berlin" (Hausverwaltungen, fixed column set, sorted Enterprise, A, B,
C, nicht bewertbar, then Punkte desc), "Wohnungsunternehmen", "Gewichte", "Log", "Methodik".
Formatting: bold + frozen header, autofilter, autofit widths, dropdowns on Prio and
Verwaltungsart, conditional fill on Prio. Values are copied, nothing is recomputed.
Bezirk is only filled where Bezirk/PLZ contains a district name (no PLZ→Bezirk guessing).

Usage: python3 scripts/export_sheet.py --in data/heidi_scored_v9.xlsx --out data/heidi_final.xlsx
"""
import argparse
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import DATA, plz_of  # noqa: E402

LEAD_COLS = ["Prio", "Punkte", "Firma", "Website", "PLZ", "Bezirk", "Verwaltungsart", "Einheiten",
             "Einheiten-Quelle", "Software", "Kundenportal", "Fernablesung erwähnt", "Messdienst",
             "Mitgliedschaft", "VDIV-Logo", "Google-Bewertung", "Anzahl Bewertungen", "Gruppe", "Makler auch",
             "Recherchetiefe", "Anzahl Quellen", "Prio-Grund", "Ansprechpartner", "Funktion", "E-Mail",
             "E-Mail-Status", "Allgemeine E-Mail", "LinkedIn", "Telefon", "Straße", "Hinweis", "Quellen", "Stand"]
AGENT_COLS = ["Agent-Quote", "Agent-Abweichung"]  # exported only when the workbook has them (run_all --agent)
PRIO_ORDER = ["Enterprise", "A", "B", "C", "nicht bewertbar"]
PRIO_FILL = {"Enterprise": "C9B6E4", "A": "A9D18E", "B": "FFE699", "C": "F8CBAD", "nicht bewertbar": "D9D9D9"}
LONG = {"Prio-Grund", "Hinweis", "Quellen", "Einheiten-Quelle", "Grund / Hinweis", "Beschreibung", "Inhalt"}
BOLD = Font(bold=True)


def rows_of(ws):
    hdr = [c.value for c in next(ws.iter_rows(max_row=1))]
    return hdr, [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True) if any(v not in (None, "") for v in r)]


def filled(v):
    return v is not None and str(v).strip() != "" and str(v).strip().lower() != "unbekannt"


def bezirk(v):
    s = str(v or "")
    if " / " in s:
        return s.split(" / ")[0].strip()
    return s.strip() if re.search(r"[A-Za-zÄÖÜäöü]", s) and not re.search(r"\d{5}", s) else None


def format_sheet(ws, n_cols):
    for c in ws[1]:
        c.font = BOLD
        c.alignment = Alignment(vertical="top", wrap_text=False)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{ws.max_row}"
    for i in range(1, n_cols + 1):
        col = get_column_letter(i)
        head = str(ws.cell(1, i).value or "")
        lens = [len(head)] + [len(str(ws.cell(r, i).value)) for r in range(2, min(ws.max_row, 800) + 1)
                              if ws.cell(r, i).value is not None]
        cap = 70 if head in LONG else 45
        ws.column_dimensions[col].width = max(8, min(max(lens) + 2, cap))


def copy_tab(src_ws, dst):
    hdr, rows = rows_of(src_ws)
    dst.append(hdr)
    for r in rows:
        dst.append([r.get(h) for h in hdr])
    format_sheet(dst, len(hdr))
    return len(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, default=DATA / "heidi_scored_v9.xlsx")
    ap.add_argument("--out", type=Path, default=DATA / "heidi_final.xlsx")
    a = ap.parse_args()
    src = openpyxl.load_workbook(a.inp, read_only=True)
    hv_hdr, hv = rows_of(src["Hausverwaltungen"])
    cols = LEAD_COLS + [c for c in AGENT_COLS if c in hv_hdr]

    leads = []
    for r in hv:
        rec = {c: r.get(c) for c in cols}
        rec["PLZ"] = plz_of(r.get("Bezirk/PLZ")) or None
        rec["Bezirk"] = bezirk(r.get("Bezirk/PLZ"))
        leads.append(rec)
    leads.sort(key=lambda x: (PRIO_ORDER.index(x["Prio"]) if x["Prio"] in PRIO_ORDER else 9,
                              -(x["Punkte"] or 0), str(x["Firma"]).lower()))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads Berlin"
    ws.append(cols)
    for rec in leads:
        ws.append([rec[c] for c in cols])
    n = len(leads)
    cidx = {c: i + 1 for i, c in enumerate(cols)}
    for r in range(2, n + 2):
        ws.cell(r, cidx["PLZ"]).number_format = "@"
        ws.cell(r, cidx["Einheiten"]).number_format = "#,##0"
        ws.cell(r, cidx["Google-Bewertung"]).number_format = "0.0"
    format_sheet(ws, len(cols))
    last = n + 1
    pcol, vcol = get_column_letter(cidx["Prio"]), get_column_letter(cidx["Verwaltungsart"])
    dv = DataValidation(type="list", formula1='"' + ",".join(PRIO_ORDER) + '"', allow_blank=True)
    dv.add(f"{pcol}2:{pcol}{max(last, 2) + 500}")
    va_vals = ["beides", "WEG", "Miet", "Gewerbe", "unbekannt"]
    va_vals += sorted({str(r["Verwaltungsart"]) for r in leads if r["Verwaltungsart"]} - set(va_vals))
    dv2 = DataValidation(type="list", formula1='"' + ",".join(va_vals)[:250] + '"', allow_blank=True)
    dv2.add(f"{vcol}2:{vcol}{max(last, 2) + 500}")
    ws.add_data_validation(dv)
    ws.add_data_validation(dv2)
    for p, color in PRIO_FILL.items():
        ws.conditional_formatting.add(f"{pcol}2:{pcol}{max(last, 2) + 500}",
                                      CellIsRule(operator="equal", formula=[f'"{p}"'],
                                                 fill=PatternFill("solid", start_color=color, end_color=color)))

    counts = {}
    for name in ("Wohnungsunternehmen", "Gewichte", "Log"):
        counts[name] = copy_tab(src[name], wb.create_sheet(name))

    # ------------------------------------------------------------------ Methodik
    m = wb.create_sheet("Methodik")
    m.column_dimensions["A"].width = 34
    m.column_dimensions["B"].width = 110
    m.column_dimensions["C"].width = 12

    def section(title):
        m.append([])
        m.append([title])
        m.cell(m.max_row, 1).font = Font(bold=True, size=12)

    def line(a_, b="", c=None):
        m.append([a_, b, c])
        m.cell(m.max_row, 2).alignment = Alignment(wrap_text=True, vertical="top")

    m.append([f"Methodik – Hausverwaltungen Berlin (Stand {date.today().isoformat()}, Quelle: {a.inp.name})"])
    m.cell(1, 1).font = Font(bold=True, size=14)
    line("Datenherkunft", "Öffentliche Impressen und Verzeichnisse (Google Places, Apollo, 11880, Gelbe Seiten, "
                          "JACASA); B2B-Kontaktdaten. Bei E-Mail-Ansprache UWG §7 beachten (Werbung per E-Mail "
                          "auch im B2B-Bereich nur mit Einwilligung bzw. mutmaßlicher Einwilligung).")

    section("Quellen (Zeilen im Tab Leads Berlin, Mehrfachnennung möglich)")
    src_count = Counter()
    for r in hv:
        for q in {x.strip() for x in re.split(r"[;,]", str(r.get("Quelle") or "")) if x.strip()}:
            src_count[q] += 1
    jac = sum(1 for r in hv if filled(r.get("Anzahl Bewertungen")) or filled(r.get("JACASA-URL"))
              or filled(r.get("Google-Bewertung")))
    for q, c in src_count.most_common():
        line(q, "", c)
    line("davon mit JACASA-Daten", "inkl. Treffer auf bestehende Zeilen (Quelle-Spalte bleibt unverändert)", jac)
    line("Leads gesamt", "", n)
    for name, c in counts.items():
        line(f"Tab {name}", "", c)

    section("Abgleich / Dubletten")
    for a_, b in [
        ("Namensnormalisierung", "Kleinbuchstaben, Umlaute ausgeschrieben, Rechtsformen (GmbH, mbH, UG, KG, AG, "
                                 "e.K., GbR, OHG, & Co.) und Zusätze wie 'NL Berlin' entfernt, danach Leerzeichen, "
                                 "Bindestriche und Satzzeichen entfernt ('ABC HV' = 'abc-hv')."),
        ("Verzeichnis → Liste", "Dublette, wenn Domain ODER normalisierter Name + PLZ übereinstimmt; Zeilen ohne PLZ "
                                "nur über den Namen (nicht bei rein generischen Namen)."),
        ("JACASA → Liste", "Name + PLZ; sonst alle Wörter ab 4 Buchstaben des kürzeren Namens im längeren, gleiche "
                           "PLZ (mindestens ein unterscheidendes Wort); sonst Name bei Zeilen ohne PLZ; sonst exakt "
                           "gleicher Name bei anderer PLZ (beide PLZ im Hinweis). Lockere Treffer bei anderer PLZ "
                           "bleiben eigene Zeilen mit Hinweis 'mögliches Duplikat'."),
        ("Ausschlüsse", "Firmen im Log (ausgeschlossen, umbenannt, unsicher) werden nicht erneut angelegt."),
        ("Grundsätze", "Keine Werte erfunden (leer = nicht gefunden); nur leere Zellen befüllt; robots.txt "
                       "beachtet; bei 403/429 abgebrochen, keine Umgehung."),
    ]:
        line(a_, b)

    section("Recherchetiefe")
    for a_, b in [
        ("A (Website geprüft)", "Website manuell bzw. per Agent gelesen, Felder aus der Website belegt."),
        ("A (Zahl geprüft)", "Wie B+, zusätzlich Einheiten-Zahl per Regel freigegeben: Zahl direkt vor "
                             "Einheiten/Wohneinheiten/Wohnungen und im selben Satz verwalten/betreuen/Verwaltung/"
                             "Bestand; Satz + URL in Einheiten-Quelle."),
        ("B+ (automatisch)", "Verzeichniseintrag, Startseite + Impressum automatisch ausgewertet (Stichwörter); "
                             "nicht manuell geprüft."),
        ("B+ (Agent)", "Wie B+, zusätzlich Startseite + Impressum vom Extraktions-Agenten (claude-haiku-5-5) "
                       "gelesen: jeder Wert mit wörtlichem Zitat + URL (Spalte Agent-Quote), von einem zweiten "
                       "Aufruf gegen den Seitentext geprüft; verworfene Werte im Log (Bereich 'Agent verworfen'). "
                       "Fehlerquoten je Feld: doc/eval.md."),
        ("B (Verzeichnis)", "Nur Verzeichnisdaten (11880, Gelbe Seiten, JACASA); Website nicht oder nicht "
                            "erfolgreich geprüft."),
        ("C (nur Name)", "Nur Name und PLZ bekannt, keine Website gefunden."),
    ]:
        line(a_, b)

    section("Scoring (Werte aus Tab Gewichte)")
    g_hdr, g_rows = rows_of(src["Gewichte"])
    gw = {r["Schlüssel"]: r for r in g_rows}

    def pts(k):
        v = gw.get(k, {}).get("Punkte")
        return f"{int(v):+d}" if v is not None else "?"

    def thr(k):
        v = gw.get(k, {}).get("Schwelle")
        return f"{int(v):,}".replace(",", ".") if v is not None else "?"

    for a_, b in [
        ("Einheiten", f"ab {thr('einheiten_hoch')} mit Quell-URL {pts('einheiten_hoch')}; ab "
                      f"{thr('einheiten_mittel')} {pts('einheiten_mittel')}. 'Einheiten Berlin' geht vor 'Einheiten'."),
        ("Bewertungen", f"ab {thr('bewertungen_hoch')} {pts('bewertungen_hoch')}; ab {thr('bewertungen_mittel')} "
                        f"{pts('bewertungen_mittel')}."),
        ("Verwaltungsart", f"beides {pts('verwaltungsart_beides')}; nur WEG oder nur Miet "
                           f"{pts('verwaltungsart_einzeln')}."),
        ("Digitalisierung", f"Kundenportal ja {pts('kundenportal')}; Verwaltersoftware genannt {pts('software')}."),
        ("Verband", f"Mitglied VDIV/IVD oder VDIV-Logo {pts('verband')}."),
        ("Quellenlage", f"in mindestens {thr('quellen')} Quellen {pts('quellen')}."),
        ("Abzüge", f"Fernablesung erwähnt {pts('fernablesung')} (bereits versorgt); Gruppe/Konzern "
                   f"{pts('gruppe')}; Website nicht erreichbar {pts('website_fehler')}."),
        ("Prio", f"Enterprise ab {thr('prio_enterprise')} Einheiten ('bundesweit' im Hinweis nur ab "
                 f"{thr('enterprise_bundesweit')} Einheiten); sonst A ab {thr('prio_a')} Punkten, B ab "
                 f"{thr('prio_b')}, darunter C; 'nicht bewertbar', wenn kein bewertbares Feld befüllt ist."),
        ("Änderung", str(gw.get("prio_a", {}).get("Beschreibung") or "")),
    ]:
        line(a_, b)

    section("Feldvollständigkeit Leads Berlin (befüllt, 'unbekannt' zählt als leer)")
    completeness = []
    for c in cols:
        k = sum(filled(r[c]) for r in leads)
        pct = round(100 * k / n, 1) if n else 0
        completeness.append((c, k, pct))
        line(c, f"{k} von {n}", pct / 100)
        m.cell(m.max_row, 3).number_format = "0.0%"

    section("Bekannte Lücken")
    gaps = [
        ("11880", "Nur Seite 1 (50 von 1.987 Treffern): ab Seite 2 HTTP 403, nicht umgangen."),
        ("JACASA", "jacasa.de blockiert automatischen Abruf (Cloudflare 403); Daten aus bereitgestellten "
                   "CSV-Exporten (243 Firmen). Keine Website/Telefon in JACASA-Daten; 'Google-Bewertung' = "
                   "Bewertung laut JACASA-Profil."),
        ("Websites", f"{sum(1 for r in hv if str(r.get('Website-Status') or '').startswith('Fehler'))} geprüfte "
                     f"Websites nicht erreichbar (Verbindung, 404, 403/429, robots.txt)."),
        ("Einheiten", f"Nur {sum(filled(r['Einheiten']) for r in leads)} Leads mit Einheiten-Zahl; automatische "
                      f"Kandidaten nur nach Regel freigegeben, Rest im Log."),
        ("Kontakte", f"Ansprechpartner {sum(filled(r['Ansprechpartner']) for r in leads)}, persönliche E-Mail "
                     f"{sum(filled(r['E-Mail']) for r in leads)}; LinkedIn nicht erhoben; Impressum-Namen sind "
                     f"Vertretungsberechtigte, nicht zwingend der passende Ansprechpartner."),
        ("Automatische Felder", "Verwaltungsart, Software, Kundenportal, Messdienst, Fernablesung, VDIV-Logo bei B+ "
                                "aus Stichwörtern (Startseite + Impressum); 'unbekannt' = nichts gefunden, nicht 'nein'."),
        ("Bezirk", "Nur wo die Quelle einen Bezirk nennt; keine Ableitung aus der PLZ."),
        ("Agent", "Nur Zeilen mit Startseite + Impressum im Cache. A-Zeilen: nur leere Felder ergänzt, abweichende "
                  "Agent-Werte in 'Agent-Abweichung'. Hinweis 'Agent: kein Verwalter, prüfen' = Verdacht, nicht "
                  "ausgeschlossen. Kein temperature=0 möglich (Modell lehnt es ab); reproduzierbar über den "
                  "Antwort-Cache."),
        ("Dubletten", "Lockere Namensähnlichkeit bei anderer PLZ nicht zusammengeführt (Hinweis "
                      "'mögliches Duplikat')."),
    ]
    for a_, b in gaps:
        line(a_, b)

    wb._sheets = [wb["Leads Berlin"], wb["Wohnungsunternehmen"], wb["Gewichte"], wb["Log"], wb["Methodik"]]
    wb.save(a.out)

    pc = Counter(r["Prio"] for r in leads)
    print(f"Tabs: {wb.sheetnames}")
    print(f"Leads Berlin: {n} Zeilen | Prio " + ", ".join(f"{p} {pc.get(p, 0)}" for p in PRIO_ORDER))
    print("\nFeldvollständigkeit Leads Berlin ('unbekannt' = leer):")
    for c, k, pct in completeness:
        print(f"  {c:22s} {pct:5.1f}%  ({k}/{n})")
    print(f"\nGeschrieben: {a.out}")


if __name__ == "__main__":
    main()
