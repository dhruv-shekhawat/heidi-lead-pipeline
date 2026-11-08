#!/usr/bin/env python3
"""Scrape "Hausverwaltung" listings in Berlin from 11880 and Gelbe Seiten.

Outputs:
  data/directory_raw.csv    every parsed listing, unfiltered
  data/directory_clean.csv  Berlin PLZ only, non-Hausverwaltung dropped, deduped

Usage:
  python3 scripts/scrape_directories.py               # both sources live
  python3 scripts/scrape_directories.py --skip-11880  # Gelbe Seiten live, 11880 rows
                                                      # reused from existing raw CSV
  python3 scripts/scrape_directories.py --from-raw    # no requests, rebuild clean CSV

Rules:
  * requests + browser User-Agent, robots.txt respected, every page cached in data/cache/
  * 11880: 1-2 s random delay; Gelbe Seiten: fixed 2 s delay
  * Gelbe Seiten: first 50 hits from the HTML page, the rest via POST /ajaxsuche
    (the "Mehr Anzeigen" endpoint; use approved by the user 2026-11-07)
  * 403 / 429 / any other non-200 stops that source; nothing is retried
"""
import base64
import csv
import json
import random
import re
import sys
import time
import unicodedata
from pathlib import Path
from urllib.parse import urljoin, urlparse

from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
RAW_CSV = DATA / "directory_raw.csv"
CLEAN_CSV = DATA / "directory_clean.csv"

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")
RAW_FIELDS = ["Firma", "Straße", "PLZ", "Ort", "Telefon", "Website",
              "Kategorie", "Quelle", "Quell-URL"]
CLEAN_FIELDS = RAW_FIELDS + ["Makler auch"]
PLZ_MIN, PLZ_MAX = 10000, 14199
MAX_PAGES = 200  # hard safety cap

# a category counts as "Hausverwaltung" if it matches this
# (Gelbe Seiten names the trade "Haus- und Grundstücksverwaltungen")
HV_CAT = re.compile(r"hausverwalt|haus- und grundstücksverwalt", re.I)
MAKLER_CAT = re.compile(r"makler", re.I)

URL_11880 = "https://www.11880.com/suche/hausverwaltung/berlin"
URL_GS = "https://www.gelbeseiten.de/branchen/hausverwaltung/berlin"
URL_GS_AJAX = "https://www.gelbeseiten.de/ajaxsuche"
GS_DELAY = 2.0

notes = []  # human-readable stop reasons / limitations, printed at the end


class SourceBlocked(Exception):
    pass


class _Resp:
    def __init__(self, e):
        self.text, self.url, self.status_code = e.get("text", ""), e.get("final_url") or e["url"], e["status"]

    def json(self):
        return json.loads(self.text)


def request(method, url, delay=None, **kw):
    """Cached, robots-aware fetch (scripts/common.py); 403/429/non-200 -> SourceBlocked."""
    try:
        e = common.fetch(url, method=method, data=kw.get("data"), headers=kw.get("headers"),
                         delay=delay if delay is not None else random.uniform(1.0, 2.0), timeout=30)
    except (common.Blocked, common.FetchError) as ex:
        raise SourceBlocked(f"{ex} on {method} {url}")
    return _Resp(e)


def clean(s):
    s = re.sub(r"\s+", " ", s or "")
    return re.sub(r"\s+,", ",", s).strip(" ,")


# --------------------------------------------------------------------- 11880
def parse_11880(html, page_url):
    soup = BeautifulSoup(html, "lxml")
    ld_items = []
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(s.string or "")
        except ValueError:
            continue
        if data.get("@type") == "SearchResultsPage":
            ld_items = data.get("mainEntity", {}).get("itemListElement", [])
            break

    # Kategorie + Website live only in the HTML list items; join on detail URL
    by_url = {}
    for li in soup.select("li.search-result-entry-item"):
        a = li.select_one("a.entry-detail-link")
        if not a:
            continue
        url = urljoin(page_url, a["href"])
        trades = li.select_one(".trades-list")
        web = ""
        for btn in li.select('a.customerButton[href^="http"]'):
            if "11880." not in urlparse(btn["href"]).netloc:
                web = btn["href"]
                break
        by_url[url] = {"Kategorie": clean(trades.get_text(" ")) if trades else "",
                       "Website": web}

    rows = []
    for it in ld_items:
        item = it.get("item", {})
        addr = item.get("address", {}) or {}
        url = item.get("url", "")
        extra = by_url.get(url, {})
        rows.append({
            "Firma": clean(item.get("name")),
            "Straße": clean(addr.get("streetAddress")),
            "PLZ": clean(addr.get("postalCode")),
            "Ort": clean(addr.get("addressLocality")),
            "Telefon": clean(item.get("telephone")),
            "Website": extra.get("Website", ""),
            "Kategorie": extra.get("Kategorie", ""),
            "Quelle": "11880",
            "Quell-URL": url,
        })
    return rows


def total_11880(html):
    m = re.search(r"(\d[\d.]*)\s*Treffer", BeautifulSoup(html, "lxml").get_text(" "))
    return int(m.group(1).replace(".", "")) if m else None


def scrape_11880():
    rows, seen = [], set()
    total = None
    for page in range(1, MAX_PAGES + 1):
        url = URL_11880 if page == 1 else f"{URL_11880}?page={page}"
        try:
            r = request("GET", url)
        except SourceBlocked as e:
            notes.append(f"11880: gestoppt auf Seite {page}: {e}")
            break
        if page == 1:
            total = total_11880(r.text)
        page_rows = parse_11880(r.text, r.url)
        new = [x for x in page_rows if x["Quell-URL"] not in seen]
        print(f"  11880 Seite {page}: {len(page_rows)} Einträge, {len(new)} neu")
        if not new:  # empty page, or page beyond the end repeating results
            break
        seen.update(x["Quell-URL"] for x in new)
        rows.extend(new)
    if total is not None:
        notes.append(f"11880: Seite meldet {total} Treffer, erfasst {len(rows)}")
    return rows


# --------------------------------------------------------------- Gelbe Seiten
def b64(s):
    try:
        return base64.b64decode(s).decode("utf-8", "replace")
    except Exception:
        return ""


def parse_gs(html):
    """Parse article.mod-Treffer blocks (used for page 1 and /ajaxsuche HTML)."""
    soup = BeautifulSoup(html, "lxml")
    rows = []
    for art in soup.select("article.mod-Treffer"):
        name = art.select_one("h2.mod-Treffer__name")
        cat = art.select_one(".mod-Treffer--besteBranche")
        link = art.select_one('a[href*="/gsbiz/"]')
        street = plz = ort = ""
        addr = art.select_one(".mod-AdresseKompakt__adress-text")
        if addr:
            ort_el = addr.select_one(".mod-AdresseKompakt__adress__ort")
            ort_txt = clean(ort_el.get_text(" ")) if ort_el else ""
            for sp in addr.find_all("span"):
                sp.extract()
            street = clean(addr.get_text(" "))
            street = clean(re.sub(r"\([^)]*\)\s*$", "", street))  # drop "(Ortsteil)"
            m = re.match(r"(\d{5})\s+(.*)", ort_txt)
            if m:
                plz, ort = m.group(1), m.group(2)
            else:
                ort = ort_txt
        tel = art.select_one(".mod-TelefonnummerKompakt__phoneNumber")
        web_el = art.select_one("[data-webseitelink]")
        rows.append({
            "Firma": clean(name.get_text(" ")) if name else "",
            "Straße": street,
            "PLZ": plz,
            "Ort": ort,
            "Telefon": clean(tel.get_text(" ")) if tel else "",
            "Website": b64(web_el["data-webseitelink"]) if web_el else "",
            "Kategorie": clean(cat.get_text(" ")) if cat else "",
            "Quelle": "Gelbe Seiten",
            "Quell-URL": link["href"] if link else "",
        })
    return rows, soup


def scrape_gs():
    rows, seen = [], set()

    def add(batch):
        new = [x for x in batch if x["Quell-URL"] not in seen]
        seen.update(x["Quell-URL"] for x in new)
        rows.extend(new)
        return new

    try:
        r = request("GET", URL_GS, delay=GS_DELAY)
    except SourceBlocked as e:
        notes.append(f"Gelbe Seiten: gestoppt: {e}")
        return rows
    batch, soup = parse_gs(r.text)
    add(batch)
    form = soup.select_one("form#mod-LoadMore")
    total_el = soup.select_one("#loadMoreGesamtzahl")
    total = int(total_el.get_text(strip=True)) if total_el else len(rows)
    print(f"  Gelbe Seiten Seite 1: {len(batch)} Einträge (gemeldet gesamt: {total})")
    if not form:
        notes.append("Gelbe Seiten: kein 'Mehr Anzeigen'-Formular gefunden")
        return rows

    params = {i["name"]: i.get("value", "")
              for i in form.select("input[name]")}
    position = int(params.get("position") or len(rows) + 1)
    step = int(params.get("anzahl") or 10)
    headers = {"X-Requested-With": "XMLHttpRequest", "Referer": URL_GS,
               "Origin": "https://www.gelbeseiten.de",
               "Accept": "application/json, text/javascript, */*; q=0.01"}
    for _ in range(MAX_PAGES):
        if position > total:
            break
        params.update(position=str(position), anzahl=str(step))
        try:
            r = request("POST", URL_GS_AJAX, delay=GS_DELAY, data=params, headers=headers)
            data = r.json()
        except SourceBlocked as e:
            notes.append(f"Gelbe Seiten: gestoppt bei Position {position}: {e}")
            break
        except ValueError:
            notes.append(f"Gelbe Seiten: keine JSON-Antwort bei Position {position} – gestoppt")
            break
        batch, _ = parse_gs(data.get("html", ""))
        new = add(batch)
        print(f"  Gelbe Seiten ab Position {position}: {len(batch)} Einträge, {len(new)} neu")
        if not batch or not new:
            break
        position += int(data.get("anzahlTreffer") or len(batch))
    if len(rows) < total:
        notes.append(f"Gelbe Seiten: gemeldet {total}, erfasst {len(rows)}")
    return rows


# ------------------------------------------------------------ normalisation
# shared with merge.py & co.: see scripts/common.py
from common import norm_domain, norm_name, norm_tokens  # noqa: E402,F401


def plz_ok(plz):
    return bool(re.fullmatch(r"\d{5}", plz or "")) and PLZ_MIN <= int(plz) <= PLZ_MAX


def categories(kategorie):
    # "Hausverwaltung, Immobilienmakler & Anlageberatung" -> 3 categories.
    # Split on "&" only when spaced, so "Haus- und Grundstücks…" stays whole.
    return [c.strip() for c in re.split(r",|\s&\s", kategorie or "") if c.strip()]


def dedupe(rows):
    """Row is a duplicate if its domain OR its name+PLZ key was already seen."""
    seen_dom, seen_np, out = set(), set(), []
    for r in rows:
        d = norm_domain(r["Website"])
        np_ = f"{norm_name(r['Firma'])}|{r['PLZ']}"
        if (d and d in seen_dom) or np_ in seen_np:
            continue
        if d:
            seen_dom.add(d)
        seen_np.add(np_)
        out.append(r)
    return out


def read_csv(path):
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def write_csv(path, rows, fields):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def main():
    DATA.mkdir(exist_ok=True)
    if "--from-raw" in sys.argv:  # rebuild clean CSV offline, no requests
        raw = read_csv(RAW_CSV)
        r11 = [r for r in raw if r["Quelle"] == "11880"]
        rgs = [r for r in raw if r["Quelle"] == "Gelbe Seiten"]
        print(f"Offline aus {RAW_CSV.relative_to(ROOT)} ({len(raw)} Zeilen)")
    else:
        if "--skip-11880" in sys.argv:
            r11 = [r for r in read_csv(RAW_CSV) if r["Quelle"] == "11880"]
            print(f"11880: nicht abgerufen, {len(r11)} Zeilen aus bestehender "
                  f"{RAW_CSV.relative_to(ROOT)} übernommen")
        else:
            print("11880 …")
            r11 = scrape_11880()
        print("Gelbe Seiten …")
        rgs = scrape_gs()
        raw = r11 + rgs
        write_csv(RAW_CSV, raw, RAW_FIELDS)

    print("\n=== Zählung je Quelle ===")
    filtered = []
    for src, rows in (("11880", r11), ("Gelbe Seiten", rgs)):
        in_plz = [r for r in rows if plz_ok(r["PLZ"])]
        kept, dropped_cats = [], {}
        for r in in_plz:
            cats = categories(r["Kategorie"])
            if any(HV_CAT.search(c) for c in cats):
                r["Makler auch"] = "ja" if any(MAKLER_CAT.search(c) for c in cats) else ""
                kept.append(r)
            else:
                k = r["Kategorie"] or "(leer)"
                dropped_cats[k] = dropped_cats.get(k, 0) + 1
        filtered.extend(kept)
        print(f"{src:13s} abgerufen {len(rows):4d} | PLZ 10000-14199 {len(in_plz):4d} | "
              f"behalten (Hausverwaltung in Kategorien) {len(kept):4d} | "
              f"davon Makler auch {sum(r['Makler auch'] == 'ja' for r in kept)}")
        if dropped_cats:
            print(f"   verworfen nach Kategorie: {dropped_cats}")

    clean_rows = dedupe(filtered)
    write_csv(CLEAN_CSV, clean_rows, CLEAN_FIELDS)
    print(f"{'Gesamt':13s} vor Dedupe {len(filtered):4d} | nach Dedupe {len(clean_rows):4d}")
    for src in ("11880", "Gelbe Seiten"):
        print(f"   davon {src}: {sum(r['Quelle'] == src for r in clean_rows)}")
    print(f"   mit Website: {sum(bool(r['Website']) for r in clean_rows)}")

    print(f"\nGeschrieben: {CLEAN_CSV.relative_to(ROOT)}"
          + ("" if "--from-raw" in sys.argv else f", {RAW_CSV.relative_to(ROOT)}"))
    if notes:
        print("\n=== Hinweise ===")
        for n in notes:
            print(" -", n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
