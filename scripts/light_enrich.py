#!/usr/bin/env python3
"""Light automatic website check for directory rows (Recherchetiefe B or C with a Website).

Input : data/heidi_merged_v3c.xlsx   Output: data/heidi_light_v3d.xlsx, data/review_einheiten.csv
Per row: fetch homepage + Impressum (1 s delay per host, 10 s timeout, robots.txt, cache).
* Website-Status: "ok" or "Fehler (<reason>)". A 403/429 only fails that row's host.
* Fills EMPTY cells only ("unbekannt" counts as empty): Verwaltungsart, Software,
  Kundenportal, Messdienst, Fernablesung erwähnt, VDIV-Logo.
* Firma: first Impressum line with a Rechtsform; replaces Firma (old name -> Name alt) only if
  it shares a 4+ letter word with the current Firma (generic words like Hausverwaltung,
  Immobilien, Berlin ignored) and Name alt is empty; otherwise noted in Hinweis.
* Einheiten candidates -> review CSV only, never into the sheet.
* Recherchetiefe -> "B+ (automatisch)" where Website-Status is ok.

Usage: python3 scripts/light_enrich.py [--in …] [--out …] [--review …] [--workers 8]
"""
import argparse
import csv
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from copy import copy
from pathlib import Path

import openpyxl
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402
import websites as site  # noqa: E402
HV = "Hausverwaltungen"
NEW_COLS = ["Website-Status"]
FIELDS = ["Verwaltungsart", "Software", "Kundenportal", "Messdienst", "Fernablesung erwähnt", "VDIV-Logo"]

LEGAL_LINE = re.compile(r"\b(GmbH|UG|KG|AG|GbR|OHG|oHG)\b|\be\.\s?K\.")
GENERIC = {"hausverwaltung", "hausverwaltungen", "immobilien", "immobilienverwaltung", "verwaltung",
           "verwaltungs", "berlin", "gmbh", "management", "immobilienmanagement", "grundstücksverwaltung",
           "gesellschaft", "service", "services", "property", "haus", "real", "estate", "wohnungsverwaltung",
           "objektverwaltung", "hausverwalter", "group", "gruppe", "holding", "partner"}
SKIP_LINE = re.compile(r"Handelsregister|Registergericht|Amtsgericht|\bAG (Charlottenburg|Berlin)|"
                       r"Versicherung|Haftpflicht|Aufsichtsbehörde|Fachbereich|Bezirksamt|Hosting|Provider|"
                       r"Internet GmbH|Agentur|Webdesign|Gestaltung|Umsetzung|Realisierung|Bildnachweis", re.I)
LEGAL_WORDS = {"gmbh", "mbh"}
WEG_RX = re.compile(r"WEG-Verwaltung|Wohnungseigentum|Eigentümergemeinschaft|Sondereigentum|"
                    r"Eigentümerversammlung", re.I)
WEG_UPPER = re.compile(r"\bWEG\b")
MIET_RX = re.compile(r"Mietverwaltung|Mietshaus|Mietobjekt|Mieterbetreuung|Vermietung", re.I)
SOFTWARE = ["Casavi", "immocloud", "DOMUS", "PowerHaus", "Wodis", "Aareon", "iX-Haus", "immoware24",
            "idwell", "etg24", "facilioo", "Hausperfekt", "wowiport"]
PORTAL_RX = re.compile(r"portal|login|kundenbereich|mieterportal|eigentuemerportal|eigentümerportal", re.I)
PORTAL_DOM = ("casavi", "etg24", "idwell")
MESS = ["ista", "Techem", "Brunata", "Minol", "KALO"]
HEIDI_CTX = re.compile(r"Messdienst|Heizkosten|Ablesung|Zähler|Verbrauch|Submetering|Abrechnung", re.I)
FUNK_RX = re.compile(r"\bFunk\b|\bFunk-|Funkzähler|Funkablesung|Fernablesung|fernablesbar", re.I)
EINH_KW = re.compile(r"Verwaltungseinheiten|Wohneinheiten|Einheiten|Wohnungen|Objekte")
NUM_RX = re.compile(r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+|\d+)(?![\d]|[.,]\d)")
NUM_BAD_AFTER = re.compile(r"^\s*(%|€|T€|TEUR|Euro|EUR|m²|qm|Mio|Mrd|Jahre|Jahren|Prozent)", re.I)
STREET_BEFORE = re.compile(r"(str\.|straße|strasse|allee|platz|damm|ufer|ring|chaussee|weg)\s*$", re.I)
PORTAL_PROVIDERS = ("casavi", "etg24", "idwell", "hausperfekt", "immoware24", "facilioo", "immocloud", "wowiport")


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def empty(v):
    return v is None or str(v).strip() == "" or str(v).strip().lower() == "unbekannt"


def words4(s):
    return {w for w in re.findall(r"[a-zäöüß]{4,}", (s or "").lower())} - GENERIC


def impressum_firma(lines):
    for ln in lines:
        if len(ln) <= 160 and LEGAL_LINE.search(ln):
            # "Impressum | X GmbH | Berlin" -> keep the segment with the Rechtsform
            segs = [s.strip() for s in re.split(r"\s+[|–—·]\s+|\s+-\s+", ln) if s.strip()]
            ln = next((s for s in segs if LEGAL_LINE.search(s)), ln)
            if re.search(r"impressum", ln, re.I):
                continue
            cand = re.sub(r"^(Firma|Firmenname|Anbieter|Betreiber|Herausgeber|Unternehmen|"
                          r"Name|Angaben gemäß.*?)\s*:\s*", "", ln, flags=re.I).strip(" ,;")
            cand = re.sub(r"^(Rechtliche\s+)?(Informationen|Angaben|Hinweise)\b.*?\b(zur|zum|der|über die|für die)\s+",
                          "", cand).strip(" .,;:")
            if SKIP_LINE.search(ln) or not (words4(cand) - LEGAL_WORDS):
                continue  # register/insurance/host lines, or nothing but "Hausverwaltung GmbH"
            if 3 <= len(cand) <= 100:
                return cand
    return None


def einheiten_candidates(text, plz_set):
    out = []
    for m in EINH_KW.finditer(text):
        a, b = max(0, m.start() - 80), min(len(text), m.end() + 80)
        win = text[a:b]
        for n in NUM_RX.finditer(win):
            if NUM_BAD_AFTER.match(win[n.end():]):
                continue
            raw = n.group(1)
            val = int(raw.replace(".", ""))
            if STREET_BEFORE.search(win[:n.start()]):
                continue  # house number
            # year exclusion only for plain 4-digit numbers ("2.000 Einheiten" is not a year)
            if not (20 <= val <= 200000) or ("." not in raw and 1900 <= val <= 2030) or str(val) in plz_set:
                continue
            out.append((val, re.sub(r"\s+", " ", win).strip()))
    seen, uniq = set(), []
    for v, p in out:
        if v not in seen:
            seen.add(v)
            uniq.append((v, p))
    return uniq


def analyse(row):
    """Network + parsing for one row (runs in a worker thread). Returns dict of findings."""
    res = {"r": row["r"], "status": None, "found": {}, "einheiten": [], "firma_cand": None,
           "heidi": [], "imp_url": None}
    home, err = site.homepage(row["Website"])
    if not home:
        res["status"] = f"Fehler ({err})"
        return res
    res["status"] = "ok"
    imp, imp_url, _ = site.impressum(home)
    res["imp_url"] = imp_url if imp else None
    pages = [home] + ([imp] if imp else [])
    html = "\n".join(p.get("text", "") for p in pages)
    lines = [ln for p in pages for ln in common.html_lines(p.get("text", ""))]
    text = "\n".join(lines)
    all_links = [l for p in pages for l in site.links(p)]

    if imp:
        res["firma_cand"] = impressum_firma(common.html_lines(imp["text"]))
    weg = bool(WEG_RX.search(text) or WEG_UPPER.search(text))
    miet = bool(MIET_RX.search(text))
    res["found"]["Verwaltungsart"] = "beides" if weg and miet else "WEG" if weg else "Miet" if miet else "unbekannt"

    sw = []
    for s in SOFTWARE:
        if s == "DOMUS":
            if re.search(r"\bDOMUS\b", html) and "DOMUS" not in (row["Firma"] or "").upper():
                sw.append(s)
        elif re.search(rf"\b{re.escape(s)}\b", html, re.I):
            sw.append(s)
    if sw:
        res["found"]["Software"] = ", ".join(sw)

    own = common.norm_domain(row["Website"])
    for t, u in all_links:
        ul, ud = u.lower(), common.norm_domain(u)
        if "wp-login" in ul or "wp-admin" in ul:
            continue
        provider = any(d in ud for d in PORTAL_DOM + PORTAL_PROVIDERS)
        on_site = ud == own or ud.endswith("." + own) or site.same_site(u, home)
        if provider or (on_site and (PORTAL_RX.search(t) or PORTAL_RX.search(ul))):
            res["found"]["Kundenportal"] = "ja"
            break

    mess = [m for m in MESS if re.search(rf"\b{m}\b", text, re.I)]
    hrefs = " ".join(u for _, u in all_links)
    for m in re.finditer(r"\bHeidi\b", text):
        ctx = text[max(0, m.start() - 60): m.end() + 60]
        if HEIDI_CTX.search(ctx):
            res["heidi"].append(ctx.replace("\n", " "))
    if re.search(r"heidi", " ".join(common.norm_domain(u) for _, u in all_links)):
        res["heidi"].append("Link: " + next(u for _, u in all_links if "heidi" in common.norm_domain(u)))
    if res["heidi"]:
        mess.append("Heidi")
    if mess:
        res["found"]["Messdienst"] = ", ".join(mess)
    if FUNK_RX.search(text):
        res["found"]["Fernablesung erwähnt"] = "ja"
    if "VDIV" in html:
        res["found"]["VDIV-Logo"] = "ja"

    plz_set = {row["PLZ"]} | set(re.findall(r"\b(\d{5})\s+[A-ZÄÖÜ]", text))
    for p in pages:
        t = " ".join(common.html_lines(p.get("text", "")))
        for v, phrase in einheiten_candidates(t, plz_set):
            if v not in {x[0] for x in res["einheiten"]}:
                res["einheiten"].append((v, phrase, site.page_url(p)))
    _ = hrefs
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, default=common.DATA / "heidi_merged_v3c.xlsx")
    ap.add_argument("--out", type=Path, default=common.DATA / "heidi_light_v3d.xlsx")
    ap.add_argument("--review", type=Path, default=common.DATA / "review_einheiten.csv")
    ap.add_argument("--workers", type=int, default=8)
    a = ap.parse_args()
    if a.out.resolve() in (a.inp.resolve(), (common.DATA / "heidi_merged_v3.xlsx").resolve()):
        sys.exit(f"Refusing to write {a.out}")

    wb = openpyxl.load_workbook(a.inp)
    ws = wb[HV]
    hm = header_map(ws)
    tmpl = ws.cell(1, ws.max_column)
    for n in NEW_COLS:
        if n not in hm:
            c = ws.cell(1, ws.max_column + 1, n)
            c.font, c.fill, c.alignment = copy(tmpl.font), copy(tmpl.fill), copy(tmpl.alignment)
            ws.column_dimensions[c.column_letter].width = 22
            hm[n] = c.column
    cell = lambda r, k: ws.cell(r, hm[k])  # noqa: E731

    rows = []
    for r in range(2, ws.max_row + 1):
        tier = str(cell(r, "Recherchetiefe").value or "")
        web = cell(r, "Website").value
        if cell(r, "Firma").value and web and tier[:1] in ("B", "C") and not tier.startswith("B+"):
            rows.append({"r": r, "Website": web, "Firma": cell(r, "Firma").value,
                         "PLZ": common.plz_of(cell(r, "Bezirk/PLZ").value)})
    print(f"Zeilen B/C mit Website: {len(rows)} – starte ({a.workers} parallel, je Host 1 s Abstand)…", flush=True)
    with ThreadPoolExecutor(max_workers=a.workers) as ex:
        results = list(ex.map(analyse, rows))

    fills = {k: 0 for k in FIELDS + ["Firma (umbenannt)", "Hinweis (Impressum-Name ungeprüft)"]}
    review, errors, heidi_hits = [], {}, []
    for res in results:
        r = res["r"]
        cell(r, "Website-Status").value = res["status"]
        if res["status"] != "ok":
            key = re.sub(r" on .*| bei .*", "", res["status"])
            errors[key] = errors.get(key, 0) + 1
            continue
        for k, v in res["found"].items():
            if empty(cell(r, k).value) and v:
                if k == "Verwaltungsart" and v == "unbekannt" and cell(r, k).value == "unbekannt":
                    continue
                cell(r, k).value = v
                fills[k] += 1
        cand, old = res["firma_cand"], cell(r, "Firma").value
        if cand and common.norm_name(cand) != common.norm_name(old):
            if words4(cand) & words4(old) and not cell(r, "Name alt").value:
                cell(r, "Name alt").value = old
                cell(r, "Firma").value = cand
                fills["Firma (umbenannt)"] += 1
            else:
                h = cell(r, "Hinweis").value
                note = f"Impressum-Name ungeprüft: {cand}"
                cell(r, "Hinweis").value = f"{h} | {note}" if h else note
                fills["Hinweis (Impressum-Name ungeprüft)"] += 1
        for v, phrase, url in res["einheiten"]:
            review.append({"Firma": cell(r, "Firma").value, "Website": cell(r, "Website").value,
                           "Einheiten_auto": v, "Phrase": phrase, "URL": url, "Freigabe": ""})
        if res["heidi"]:
            heidi_hits.append((cell(r, "Firma").value, res["heidi"]))
        cell(r, "Recherchetiefe").value = "B+ (automatisch)"

    from add_sources import merge_same_name_rows  # renames can produce same-name duplicates
    print("Dubletten nach Umbenennung:")
    merge_same_name_rows(ws)
    wb.save(a.out)
    with open(a.review, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["Firma", "Website", "Einheiten_auto", "Phrase", "URL", "Freigabe"])
        w.writeheader()
        w.writerows(review)

    ok = sum(r["status"] == "ok" for r in results)
    print(f"versucht: {len(results)} | ok: {ok} | Fehler: {len(results) - ok}")
    print("  Fehlergründe:", dict(sorted(errors.items(), key=lambda x: -x[1])))
    print(f"  Impressum gefunden: {sum(bool(r['imp_url']) for r in results if r['status'] == 'ok')} von {ok}")
    print("Befüllt (nur leere/unbekannt-Zellen):", fills)
    vals = {}
    for res in results:
        if res["status"] == "ok":
            v = res["found"].get("Verwaltungsart")
            vals[v] = vals.get(v, 0) + 1
    print("  Verwaltungsart-Befund (ok-Zeilen):", vals)
    print(f"Zeilen in {a.review.name}: {len(review)} (aus {len({x['Firma'] for x in review})} Firmen)")
    print(f"Heidi-Treffer (Messdienst-Kontext): {len(heidi_hits)}")
    for f_, ctx in heidi_hits:
        print(f"   {f_}: {ctx[:2]}")
    print(f"Geschrieben: {a.out.name}, {a.review.name}")


if __name__ == "__main__":
    main()
