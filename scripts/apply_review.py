#!/usr/bin/env python3
"""Apply data/review_einheiten.csv to a scored workbook.

Usage:
  python3 scripts/apply_review.py --in data/heidi_scored_v6.xlsx --out OUT.xlsx [--auto]

--auto  sets Freigabe automatically (and writes it back into the review CSV):
        ja  = the row's own number is followed within 3 words by Einheiten / Wohneinheiten /
              Verwaltungseinheiten / Wohnungen AND the same sentence contains verwalten /
              verwaltet(e…) / betreuen / betreut(e…) / Verwaltung(en) / Bestand (whole words, so
              "Hausverwaltung" or "Verwaltungseinheiten" do not count)
        nein (unsicher) = everything else
Without --auto the Freigabe column filled by hand is used (ja / nein…; empty = skipped).

ja   -> Hausverwaltungen: Einheiten (only if empty; several ja numbers per firm -> largest),
        Einheiten-Quelle = "URL – „sentence“", Recherchetiefe "A (Zahl geprüft)"
nein -> one Log row per candidate (Bereich "Einheiten verworfen") with the phrase
Run score.py on OUT afterwards.
"""
import argparse
import csv
import re
import sys
from copy import copy
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import DATA  # noqa: E402

KW = re.compile(r"^(Einheiten|Wohneinheiten|Verwaltungseinheiten|Wohnungen)$", re.I)
VERB = re.compile(r"\b(verwalten|verwaltet\w*|verwaltete\w*|betreuen|betreut\w*|Verwaltung(?:en)?|"
                  r"Bestand)\b", re.I)
SENT_SPLIT = re.compile(r"(?<=[.!?])\s+(?=[A-ZÄÖÜ„\"“])")


def fmt_variants(n):
    n = int(n)
    return {str(n), f"{n:,}".replace(",", ".")}


def judge(number, phrase):
    """Return (ok, sentence) for the auto rule."""
    variants = fmt_variants(number)
    for sent in SENT_SPLIT.split(phrase):
        toks = sent.split()
        for i, t in enumerate(toks):
            core = t.strip("(),;:~≈+").rstrip(".")
            if core not in variants:
                continue
            nxt = [x.strip(".,;:()!?") for x in toks[i + 1:i + 4]]
            if any(KW.match(x) for x in nxt) and VERB.search(sent):
                return True, sent.strip()
    # sentence for the record even when rejected
    for sent in SENT_SPLIT.split(phrase):
        if any(v in sent for v in variants):
            return False, sent.strip()
    return False, phrase


NOT_MANAGED = re.compile(r"\b(Einheiten|Wohneinheiten|Wohnungen)\s+(verkauft|vermittelt|gebaut|errichtet|"
                         r"modernisiert|saniert|entwickelt)", re.I)


def de(n):
    return f"{int(n):,}".replace(",", ".")


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, default=DATA / "heidi_scored_v6.xlsx")
    ap.add_argument("--review", type=Path, default=DATA / "review_einheiten.csv")
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--auto", action="store_true")
    a = ap.parse_args()
    if a.out.resolve() in (a.inp.resolve(), (DATA / "heidi_merged_v3.xlsx").resolve()):
        sys.exit(f"Refusing to write {a.out}")
    with open(a.review, encoding="utf-8-sig") as f:
        rev = list(csv.DictReader(f))
    for r in rev:
        ok, sent = judge(r["Einheiten_auto"], r["Phrase"])
        r["_sent"] = sent
        if a.auto:
            r["Freigabe"] = "ja" if ok else "nein (unsicher)"
    if a.auto:
        with open(a.review, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["Firma", "Website", "Einheiten_auto", "Phrase", "URL", "Freigabe"],
                               extrasaction="ignore")
            w.writeheader()
            w.writerows(rev)

    wb = openpyxl.load_workbook(a.inp)
    ws, log = wb["Hausverwaltungen"], wb["Log"]
    hm, lm = header_map(ws), header_map(log)
    by_web = {}
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, hm["Website"]).value:
            by_web.setdefault(str(ws.cell(r, hm["Website"]).value).strip().lower(), r)

    ja = [r for r in rev if str(r["Freigabe"]).strip().lower().startswith("ja")]
    nein = [r for r in rev if str(r["Freigabe"]).strip().lower().startswith("nein")]
    # several ja numbers for one firm: largest number whose own keyword is not followed by
    # "verkauft/vermittelt/gebaut/…" (e.g. "11.806 Wohnungen verkauft" is not a managed portfolio)
    best = {}
    for r in ja:
        k = r["Website"].strip().lower()
        own = re.search(rf"(?<![\d.]){re.escape(de(r['Einheiten_auto']))}(?![\d])\S*\s+(\S+\s+){{0,2}}?\S*", r["_sent"])
        r["_not_managed"] = bool(own and NOT_MANAGED.search(own.group(0) + " " + r["_sent"][own.end():own.end() + 20]))
        cand = best.get(k)
        if cand is None or (cand["_not_managed"] and not r["_not_managed"]) or \
                (r["_not_managed"] == cand["_not_managed"] and int(r["Einheiten_auto"]) > int(cand["Einheiten_auto"])):
            best[k] = r
    applied, skipped = [], []
    for k, r in best.items():
        row = by_web.get(k)
        if not row:
            skipped.append((r["Firma"], "Zeile nicht gefunden"))
            continue
        if ws.cell(row, hm["Einheiten"]).value not in (None, ""):
            skipped.append((r["Firma"], f"Einheiten schon belegt ({ws.cell(row, hm['Einheiten']).value})"))
            continue
        ws.cell(row, hm["Einheiten"]).value = int(r["Einheiten_auto"])
        if not ws.cell(row, hm["Einheiten-Quelle"]).value:
            ws.cell(row, hm["Einheiten-Quelle"]).value = f"{r['URL']} – „{r['_sent'][:250]}“"
        ws.cell(row, hm["Recherchetiefe"]).value = "A (Zahl geprüft)"
        applied.append(r)
    style = log.max_row
    for r in nein:
        row = by_web.get(r["Website"].strip().lower())
        n = log.max_row + 1
        for c in range(1, log.max_column + 1):
            if log.cell(style, c).has_style:
                log.cell(n, c)._style = copy(log.cell(style, c)._style)
        vals = {"Bereich": "Einheiten verworfen", "Firma": r["Firma"], "Website": r["Website"],
                "Typ": ws.cell(row, hm["Typ"]).value if row else None,
                "Grund / Hinweis": f"Einheiten_auto {de(r['Einheiten_auto'])} {r['Freigabe']}: {r['Phrase']}",
                "Quelle": r["URL"]}
        for k2, v in vals.items():
            if k2 in lm and v:
                log.cell(n, lm[k2]).value = v
    wb.save(a.out)
    print(f"Review-Zeilen: {len(rev)} | ja: {len(ja)} | nein: {len(nein)} | leer: {len(rev) - len(ja) - len(nein)}")
    print(f"angewendet (Einheiten gesetzt, A (Zahl geprüft)): {len(applied)} Firmen | übersprungen: {skipped or 0}")
    print(f"Log-Zeilen neu (Einheiten verworfen): {len(nein)}")
    print("\n=== ja-Zeilen ===")
    for r in ja:
        mark = "✓ übernommen" if r in applied else "– nicht übernommen (andere Zahl derselben Firma gewählt)"
        print(f"- {r['Firma']} | {de(r['Einheiten_auto'])} | „{r['_sent'][:200]}“ {mark}")
    print(f"\nGeschrieben: {a.out.name}{', ' + a.review.name + ' (Freigabe gesetzt)' if a.auto else ''}")


if __name__ == "__main__":
    main()
