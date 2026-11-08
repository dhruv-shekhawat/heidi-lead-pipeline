#!/usr/bin/env python3
"""Fetch Gelbe Seiten detail pages for every Gelbe Seiten row in data/directory_clean.csv.

Extracts Website, Telefon, E-Mail and the full category list ("Branchen") and writes
data/directory_detail.csv (identity columns Firma/Straße/PLZ/Quell-URL carried over for matching).
11880 rows are skipped. 2 s delay, robots.txt respected, pages cached in data/cache/.
A 403/429 stops the whole step (single host); rows fetched so far are still written.

Usage: python3 scripts/fetch_details.py
"""
import csv
import json
import sys
from pathlib import Path

from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402

IN_CSV = common.DATA / "directory_clean.csv"
OUT_CSV = common.DATA / "directory_detail.csv"
FIELDS = ["Firma", "Straße", "PLZ", "Quelle", "Quell-URL", "Website", "Telefon", "E-Mail",
          "Kategorien", "Abruf"]


def parse(html):
    soup = BeautifulSoup(html, "lxml")
    ld = {}
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            d = json.loads(s.string or "")
        except ValueError:
            continue
        if isinstance(d, dict) and d.get("address"):
            ld = d
            break
    web = ""
    a = soup.select_one(".contains-icon-big-homepage a[href]")
    if a and a["href"].startswith("http"):
        web = a["href"]
    if not web:
        same = ld.get("sameAs")
        for u in (same if isinstance(same, list) else [same] if same else []):
            if common.is_company_domain(common.norm_domain(u)):
                web = u
                break
    if web and not common.is_company_domain(common.norm_domain(web)):
        web = ""
    mail = ""
    el = soup.select_one("#email_versenden[data-link]")
    if el and el["data-link"].lower().startswith("mailto:"):
        mail = el["data-link"][7:].split("?")[0].strip()
    tel = ld.get("telephone") or ""
    if not tel:
        t = soup.select_one(".telefonnummer")
        tel = t.get_text(" ", strip=True) if t else ""
    cats = soup.select_one(".mod-Branchen__content")
    return {"Website": web, "Telefon": tel, "E-Mail": mail,
            "Kategorien": cats.get_text(" ", strip=True) if cats else ""}


def main():
    with open(IN_CSV, encoding="utf-8-sig") as f:
        rows = [r for r in csv.DictReader(f) if r["Quelle"] == "Gelbe Seiten"]
    out, fetched, cached, errors, stopped = [], 0, 0, 0, None
    for i, r in enumerate(rows, 1):
        base = {k: r[k] for k in ("Firma", "Straße", "PLZ", "Quelle", "Quell-URL")}
        try:
            e = common.fetch(r["Quell-URL"], delay=2.0, timeout=20)
        except common.Blocked as ex:
            stopped = f"{ex} bei {r['Quell-URL']}"
            break
        except common.FetchError as ex:
            errors += 1
            out.append({**base, "Abruf": f"Fehler ({ex})"})
            continue
        fetched += 1
        cached += bool(e.get("cached"))
        out.append({**base, **parse(e["text"]), "Abruf": "ok"})
        if i % 50 == 0:
            print(f"  {i}/{len(rows)} …", flush=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(out)
    print(f"Gelbe-Seiten-Zeilen: {len(rows)} | abgerufen: {fetched} (davon aus Cache {cached}) | "
          f"Fehler: {errors}")
    print(f"  Website gefunden: {sum(bool(r.get('Website')) for r in out)} | Telefon: "
          f"{sum(bool(r.get('Telefon')) for r in out)} | E-Mail: {sum(bool(r.get('E-Mail')) for r in out)} | "
          f"Kategorien: {sum(bool(r.get('Kategorien')) for r in out)}")
    if stopped:
        print(f"GESTOPPT: {stopped}")
    print(f"Geschrieben: {OUT_CSV.relative_to(common.ROOT)}")
    return 2 if stopped else 0


if __name__ == "__main__":
    sys.exit(main())
