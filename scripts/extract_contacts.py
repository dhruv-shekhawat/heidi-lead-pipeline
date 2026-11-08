#!/usr/bin/env python3
"""Extract a contact person and a personal e-mail from Impressum / Kontakt pages.

Input : data/heidi_light_v3d.xlsx  (Hausverwaltungen rows with Recherchetiefe A… or B+… and a Website)
Output: data/heidi_contacts_v5.xlsx
Columns created if missing: Ansprechpartner, Funktion, E-Mail, E-Mail-Status, LinkedIn, Kontakt-Quelle.
* Pages come from the cache (data/cache/) or are fetched (1 s delay per host, robots.txt).
* Name: first title keyword in the Impressum (Geschäftsführer/in, Geschäftsführung, Inhaber/in,
  vertreten durch, Vertretungsberechtigt…) followed by up to 4 capitalised words on the same
  line, stopping at a comma; occurrences that yield no person name (company names, register
  data) are skipped and the next occurrence is used.
* Soft hyphens / zero-width characters are removed at fetch (common.fetch); academic degrees
  (Dipl.-Ing., Dipl.-Kfm., …) before the name are skipped; "Geschäftsführende Gesellschafter" counts
  as a title (Funktion "Geschäftsführender Gesellschafter").
* E-Mail: first address on Impressum or Kontakt page on the company's own domain whose local
  part does not start with info/kontakt/mail/office/post/buero/verwaltung/service/hausverwaltung.
* LinkedIn is created empty (no extraction rule defined).
Only empty cells are filled.

Usage: python3 scripts/extract_contacts.py [--in …] [--out …] [--workers 8]
"""
import argparse
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
NEW_COLS = ["Ansprechpartner", "Funktion", "E-Mail", "E-Mail-Status", "LinkedIn", "Kontakt-Quelle"]
KW = re.compile(r"(Geschäftsführende[rn]?\s+Gesellschafter(?:in)?|Geschäftsführerin|Geschäftsführer|"
                r"Geschäftsführung|Inhaberin|Inhaber|vertreten durch|Vertretungsberechtigt\w*)", re.I)
# academic degrees before a name ("Dipl.-Ing., Dipl.-Kfm. Max Muster") are skipped, not taken as the name
DEGREE = r"Dipl\.\s?-?\s?[A-ZÄÖÜa-zäöü]+\.?(?:-[A-ZÄÖÜa-zäöü]+\.?)*(?:\s?\(FH\))?|M\.\s?A\.|B\.\s?A\.|MBA|LL\.\s?M\.|M\.\s?Sc\.|B\.\s?Sc\."
TITLE_SKIP = re.compile(r"^[\s:()\-–,]*(?:(?:die|den|der|des|durch|alleinige[nr]?|einzige[nr]?|"
                        r"Geschäftsführer(?:in)?|Geschäftsführung|Geschäftsführende[rn]?|GF|"
                        r"Gesellschafter(?:in)?|Inhaber(?:in)?|" + DEGREE + r"|"
                        r"\(?in\)?|\(r\)|Herrn?|Frau)[\s:()\-–,]*)*", re.I)
WORD = r"(?:[A-ZÄÖÜ][\wäöüß'.\-]*|von|van|de|zu|der|den)"
NAME_RX = re.compile(rf"{WORD}(?:[ \t]+{WORD}){{0,3}}")
STOP = {"Amtsgericht", "Registergericht", "Handelsregister", "HRB", "HRA", "Sitz", "Tel", "Tel.", "Telefon",
        "E-Mail", "Email", "Fax", "Umsatzsteuer", "USt", "USt-IdNr", "USt-IdNr.", "Steuernummer",
        "Registernummer", "Kontakt", "Anschrift", "Adresse", "Verantwortlich", "Inhaltlich", "Impressum",
        "Postanschrift", "Website", "Internet", "Mail", "Die", "Der", "Das", "Unsere", "Unser"}
COMPANY = re.compile(r"\b(GmbH|mbH|UG|KG|AG|GbR|OHG|oHG|e\.?\s?K\.?|Verwaltung|Hausverwaltung|"
                     r"Immobilien|Gesellschaft|Holding)\b")
GENERIC_MAIL = ("info", "kontakt", "mail", "office", "post", "buero", "büro", "verwaltung", "service",
                "hausverwaltung")
# Extension beyond the plan's list (logged in PROGRESS): role mailboxes that are not a person
ROLE_MAIL = ("datenschutz", "webmaster", "bewerbung", "buchhaltung", "support", "team", "hallo", "hello",
             "zentrale", "miete", "mieter", "weg", "hv", "verkauf", "immobilien", "immo", "neukunden",
             "liegenschaften", "zukunft", "empfang", "anfrage", "anfragen", "jobs", "karriere", "presse",
             "noreply", "no-reply", "admin", "sekretariat", "vermietung", "abrechnung", "technik", "rechnung",
             "kunden", "eigentuemer", "objekt", "notfall", "dsb", "privacy", "contact", "welcome", "mieterservice")
MAIL_RX = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def find_person(lines):
    text = "\n".join(lines)
    for m in KW.finditer(text):
        rest = text[m.end():]
        rest = re.sub(r"^\w*", "", rest)           # suffix: Vertretungsberechtigt-er
        skip = TITLE_SKIP.match(rest)
        rest = rest[skip.end():] if skip else rest
        line = rest.split("\n", 1)[0].split(",", 1)[0]
        nm = NAME_RX.match(line)
        if not nm:
            continue
        words = []
        for w in nm.group(0).split():
            if w.rstrip(":.") in STOP or w in STOP:
                break
            words.append(w)
        while words and words[-1].lower() in ("von", "van", "de", "zu", "der", "den"):
            words.pop()
        name = " ".join(words).strip(" .:-")
        if len(words) < 2 or COMPANY.search(name) or any(re.fullmatch(r"[A-ZÄÖÜ]{3,}", w) for w in words) \
                or re.search(r"Gesellschafter|Geschäftsf|Haus\b", name):
            continue
        title = m.group(1)
        if title.lower().startswith(("vertretungsberechtigt", "vertreten")):
            t2 = re.search(r"Geschäftsführende[rn]?\s+Gesellschafter(?:in)?|Geschäftsführer(?:in)?|Inhaber(?:in)?|"
                           r"Gesellschafter(?:in)?", text[m.end():m.end() + 60])
            title = t2.group(0) if t2 else title
        if re.match(r"Geschäftsführende", title, re.I):  # "Geschäftsführende Gesellschafter: A und B"
            title = "Geschäftsführende Gesellschafterin" if title.endswith("in") else "Geschäftsführender Gesellschafter"
        return name, title[0].upper() + title[1:]
    return None, None


def decode_cf(hexstr):
    try:
        key = int(hexstr[:2], 16)
        return "".join(chr(int(hexstr[i:i + 2], 16) ^ key) for i in range(2, len(hexstr), 2))
    except ValueError:
        return ""


def emails(entry):
    html = entry.get("text", "")
    soup = BeautifulSoup(html, "lxml")
    found = []
    for a in soup.select("a[href^='mailto:'], a[href^='MAILTO:']"):
        found.append(a["href"][7:].split("?")[0])
    for el in soup.select("[data-cfemail]"):
        found.append(decode_cf(el["data-cfemail"]))
    found += MAIL_RX.findall(" ".join(common.html_lines(html)))
    out = []
    for e in found:
        e = e.strip().strip(".").lower()
        if MAIL_RX.fullmatch(e) and e not in out:
            out.append(e)
    return out


PERSON_LOCAL = re.compile(r"^[a-z]{1,20}(?:\.[a-z][a-z-]{0,20}){1,2}$")  # first.last, f.last, c.d.last


def fold(s):
    s = (s or "").lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss"), ("é", "e"), ("è", "e")):
        s = s.replace(a, b)
    return s


def personal_email(addresses, domain, person=None):
    """Plan rule (own domain, not info/kontakt/…) PLUS a positive person test: local part looks
    like first.last / f.last, or contains a 4+ letter part of the contact person's name."""
    d = common.norm_domain(domain)
    name_parts = {w for w in re.findall(r"[a-z]{4,}", fold(person))}
    for e in addresses:
        local, _, host = e.partition("@")
        host = common.norm_domain(host)
        if not (host == d or host.endswith("." + d) or d.endswith("." + host)):
            continue
        if local.startswith(GENERIC_MAIL) or local.split(".")[0].split("-")[0] in ROLE_MAIL:
            continue
        if len(local) >= 3 and local.replace("-", "") in d.replace("-", ""):
            continue  # company-name mailbox, e.g. musterhv@musterhv-gmbh.de
        if PERSON_LOCAL.match(local) or any(n in local for n in name_parts):
            return e
    return None


def analyse(row):
    res = {"r": row["r"], "name": None, "title": None, "imp_url": None, "mail": None, "mail_url": None,
           "err": None}
    home, err = site.homepage(row["Website"])
    if not home:
        res["err"] = err
        return res
    imp, imp_url, _ = site.impressum(home)
    kon, kon_url, _ = site.kontakt(home)
    if imp:
        res["imp_url"] = site.page_url(imp)
        res["name"], res["title"] = find_person(common.html_lines(imp["text"]))
    for page in (imp, kon):
        if page:
            m = personal_email(emails(page), row["Website"], res["name"] or row.get("Ansprechpartner"))
            if m:
                res["mail"], res["mail_url"] = m, site.page_url(page)
                break
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, default=common.DATA / "heidi_light_v3d.xlsx")
    ap.add_argument("--out", type=Path, default=common.DATA / "heidi_contacts_v5.xlsx")
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
            ws.column_dimensions[c.column_letter].width = 24
            hm[n] = c.column
    cell = lambda r, k: ws.cell(r, hm[k])  # noqa: E731
    empty = lambda v: v is None or str(v).strip() == ""  # noqa: E731

    rows = []
    for r in range(2, ws.max_row + 1):
        tier = str(cell(r, "Recherchetiefe").value or "")
        if cell(r, "Firma").value and cell(r, "Website").value and (tier.startswith("A") or tier.startswith("B+")):
            rows.append({"r": r, "Website": cell(r, "Website").value,
                         "Ansprechpartner": cell(r, "Ansprechpartner").value})
    print(f"Zeilen A/B+ mit Website: {len(rows)} – starte…", flush=True)
    with ThreadPoolExecutor(max_workers=a.workers) as ex:
        results = list(ex.map(analyse, rows))

    n_name = n_mail = n_err = n_imp = 0
    for res in results:
        r = res["r"]
        if res["err"]:
            n_err += 1
            continue
        n_imp += bool(res["imp_url"])
        if res["name"] and empty(cell(r, "Ansprechpartner").value):
            cell(r, "Ansprechpartner").value = res["name"]
            if empty(cell(r, "Funktion").value):
                cell(r, "Funktion").value = res["title"]
            if empty(cell(r, "Kontakt-Quelle").value):
                cell(r, "Kontakt-Quelle").value = res["imp_url"]
            n_name += 1
        if res["mail"] and empty(cell(r, "E-Mail").value):
            cell(r, "E-Mail").value = res["mail"]
            cell(r, "E-Mail-Status").value = "Website"
            if empty(cell(r, "Kontakt-Quelle").value):
                cell(r, "Kontakt-Quelle").value = res["mail_url"]
            n_mail += 1
    wb.save(a.out)
    tot_name = sum(1 for r in range(2, ws.max_row + 1) if not empty(cell(r, "Ansprechpartner").value))
    tot_mail = sum(1 for r in range(2, ws.max_row + 1) if not empty(cell(r, "E-Mail").value))
    print(f"versucht: {len(results)} | Website nicht erreichbar: {n_err} | Impressum gefunden: {n_imp}")
    print(f"Zeilen mit Name (neu befüllt): {n_name} | mit persönlicher E-Mail (neu): {n_mail}")
    print(f"Gesamt im Tab: Ansprechpartner {tot_name}, E-Mail {tot_mail}")
    print(f"Geschrieben: {a.out.name}")


if __name__ == "__main__":
    main()
