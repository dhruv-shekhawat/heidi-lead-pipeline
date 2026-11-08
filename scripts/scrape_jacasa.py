#!/usr/bin/env python3
"""Scrape JACASA Hausverwaltung profiles in Berlin -> data/jacasa.csv.

STATUS 2026-11-08: jacasa.de answers 403 (Cloudflare block page) for the Berlin list,
so this script stops there (global rule: 403/429 -> stop the step). The extraction
below is written against generic markup (JSON-LD + labelled fields) and has NOT been
validated on real profile pages. Check the output before trusting it if the site opens.

Flow:
  1. GET https://www.jacasa.de/hausverwaltung/be/berlin (cached, robots.txt respected)
  2. Find the "Weitere anzeigen" load-more endpoint in the page source and page through it
     to collect all profile URLs.
  3. Only if no endpoint can be found (NOT after a 403): sitemap, keep URLs with "berlin".
  4. Fetch each profile (2 s delay), keep PLZ 10000-14199, extract fields.

Usage: python3 scripts/scrape_jacasa.py        (exit code 2 if blocked)
"""
import csv
import json
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402

BASE = "https://www.jacasa.de"
LIST_URL = f"{BASE}/hausverwaltung/be/berlin"
SITEMAP = f"{BASE}/sitemaps/sitemap-index.xml"
OUT = common.DATA / "jacasa.csv"
DELAY = 2.0
FIELDS = ["Firma", "Straße", "PLZ", "Website", "Telefon", "Spezialisierung", "Mitgliedschaft",
          "Mitarbeiter", "Aktiv seit", "Bewertung", "Anzahl Bewertungen", "Tätig in",
          "Beschreibung", "Profil-URL"]
PROFILE_RE = re.compile(r"^/hausverwaltung/[^/?#]+/?$|^/profil/|/hausverwaltung/be/berlin/[^/?#]+")


def get(url, **kw):
    return common.fetch(url, delay=DELAY, timeout=20, **kw)


def profile_links(html, base):
    soup = BeautifulSoup(html, "lxml")
    out = []
    for a in soup.select("a[href]"):
        href = a["href"].split("#")[0]
        path = href.replace(BASE, "")
        if path.startswith("/") and PROFILE_RE.search(path) and path.rstrip("/") != "/hausverwaltung/be/berlin":
            out.append(urljoin(base, href))
    return list(dict.fromkeys(out))


def find_loadmore(html):
    """Return (url_template, start) for the load-more call, or None."""
    soup = BeautifulSoup(html, "lxml")
    for el in soup.find_all(attrs={"data-url": True}) + soup.find_all(attrs={"data-href": True}) \
            + soup.find_all(attrs={"data-next": True}) + soup.find_all("a", href=True):
        txt = el.get_text(" ", strip=True).lower()
        url = el.get("data-url") or el.get("data-next") or el.get("data-href") or el.get("href")
        if url and ("weitere anzeigen" in txt or "load" in url.lower() or "page=" in url):
            return urljoin(LIST_URL, url)
    m = re.search(r"""['"]([^'"]*(?:load-?more|ajax|api)[^'"]*berlin[^'"]*)['"]""", html, re.I)
    return urljoin(LIST_URL, m.group(1)) if m else None


def collect_profiles():
    entry = get(LIST_URL)
    urls = profile_links(entry["text"], LIST_URL)
    nxt = find_loadmore(entry["text"])
    print(f"  Liste: {len(urls)} Profile auf Seite 1, Load-more: {nxt or 'nicht gefunden'}")
    if nxt:
        page = 2
        while nxt and page < 100:
            e = get(nxt)
            body = e["text"]
            if "json" in e.get("content_type", ""):
                data = json.loads(body)
                body = data.get("html") or data.get("content") or json.dumps(data)
            new = [u for u in profile_links(body, LIST_URL) if u not in urls]
            print(f"  Load-more Seite {page}: {len(new)} neu")
            if not new:
                break
            urls += new
            nxt = find_loadmore(body) or re.sub(r"page=\d+", f"page={page + 1}", nxt)
            page += 1
        return urls, "load-more"
    # fallback only when no endpoint was found
    sm = get(SITEMAP)
    locs = re.findall(r"<loc>([^<]+)</loc>", sm["text"])
    subs = [u for u in locs if u.endswith(".xml")]
    for s in subs:
        locs += re.findall(r"<loc>([^<]+)</loc>", get(s)["text"])
    urls = [u for u in locs if "berlin" in u.lower() and not u.endswith(".xml")
            and PROFILE_RE.search(u.replace(BASE, ""))]
    print(f"  Sitemap-Fallback: {len(urls)} URLs mit 'berlin' im Slug")
    return list(dict.fromkeys(urls)), "sitemap"


def labelled(lines, *labels):
    for i, ln in enumerate(lines):
        for lab in labels:
            if ln.lower().startswith(lab.lower()):
                rest = ln[len(lab):].strip(" :")
                return rest or (lines[i + 1] if i + 1 < len(lines) else "")
    return ""


def parse_profile(html, url):
    soup = BeautifulSoup(html, "lxml")
    ld = {}
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            d = json.loads(s.string or "")
        except ValueError:
            continue
        for item in (d if isinstance(d, list) else [d]):
            if isinstance(item, dict) and item.get("address"):
                ld = item
                break
    lines = common.html_lines(html)
    text = "\n".join(lines)
    addr = ld.get("address") or {}
    agg = ld.get("aggregateRating") or {}
    h1 = soup.find("h1")
    web = ""
    for a in soup.select("a[href^='http']"):
        d = common.norm_domain(a["href"])
        if d and "jacasa" not in d and common.is_company_domain(d) and \
                re.search(r"website|webseite|homepage|zur website", a.get_text(" ", strip=True), re.I):
            web = a["href"]
            break
    spez = [x for x in ("Mietobjekte", "WEG", "Sondereigentum", "Gewerbe") if re.search(rf"\b{x}\b", text)]
    mitgl = [x for x in ("VDIV", "IVD", "BVI") if re.search(rf"\b{x}\b", text)]
    desc = ld.get("description") or labelled(lines, "Über uns", "Beschreibung")
    return {
        "Firma": ld.get("name") or (h1.get_text(" ", strip=True) if h1 else ""),
        "Straße": addr.get("streetAddress", ""),
        "PLZ": addr.get("postalCode", "") or common.plz_of(text[:2000]),
        "Website": web,
        "Telefon": ld.get("telephone", ""),
        "Spezialisierung": ", ".join(spez),
        "Mitgliedschaft": ", ".join(mitgl),
        "Mitarbeiter": labelled(lines, "Mitarbeiter"),
        "Aktiv seit": labelled(lines, "Aktiv seit", "Gegründet", "Jahre am Markt"),
        "Bewertung": str(agg.get("ratingValue", "")),
        "Anzahl Bewertungen": str(agg.get("reviewCount") or agg.get("ratingCount") or ""),
        "Tätig in": labelled(lines, "Tätig in"),
        "Beschreibung": (desc or "")[:300],
        "Profil-URL": url,
    }


def main():
    try:
        urls, how = collect_profiles()
    except common.Blocked as e:
        print(f"GESTOPPT: {e} – jacasa.de blockiert (403/429). Keine Ausgabe geschrieben.")
        return 2
    except common.FetchError as e:
        print(f"FEHLER Liste: {e}")
        return 1
    rows, fetched, errors = [], 0, 0
    for u in urls:
        try:
            e = get(u)
        except common.Blocked as ex:
            print(f"GESTOPPT bei Profil {u}: {ex}")
            break
        except common.FetchError:
            errors += 1
            continue
        fetched += 1
        r = parse_profile(e["text"], u)
        if r["PLZ"].isdigit() and 10000 <= int(r["PLZ"]) <= 14199:
            rows.append(r)
    with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)
    print(f"Quelle der Profil-URLs: {how} ({len(urls)})")
    print(f"Profile abgerufen: {fetched} (Fehler {errors}) | PLZ 10000-14199: {len(rows)}")
    print(f"  mit Website: {sum(bool(r['Website']) for r in rows)} | mit Mitgliedschaft: "
          f"{sum(bool(r['Mitgliedschaft']) for r in rows)} | mit Spezialisierung: "
          f"{sum(bool(r['Spezialisierung']) for r in rows)}")
    print(f"Geschrieben: {OUT.relative_to(common.ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
