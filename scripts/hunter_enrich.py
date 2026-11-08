#!/usr/bin/env python3
"""Hunter.io enrichment for the top leads: email pattern, listed people, verified e-mail.

Input : data/heidi_contacts_v10.xlsx (Hausverwaltungen rows with Prio Enterprise / A / B and a Website)
Output: data/heidi_contacts_v10b.xlsx
Key   : HUNTER_API_KEY from the environment or from .env in the project root (never logged, never cached).

Budget (Hunter free plan = 50 credits/month: Domain Search 1 credit, Email Verifier 0.5 credit):
  rows ordered Enterprise, A, B, then Punkte desc; at most --max-domains (25) distinct domains are
  searched and at most --max-verifications (50) addresses verified. Domains already in the Hunter
  cache are reused even below the cut (no credits), so a changed ranking never drops earlier results. The free /v2/account call is used
  first to cap both by what is actually left this month.
Per domain (Domain Search, limit 10):
  * new columns Hunter-Pattern and Hunter-Personen ("First Last (Position) email NN%; …")
  * if a listed person matches the Ansprechpartner → E-Mail, E-Mail-Status "Hunter NN%"
  * else if a pattern exists and an Ansprechpartner is known → address built from the pattern,
    E-Mail-Status "Muster (Hunter-Pattern)"
  * only EMPTY E-Mail cells are written (an address already found on the website is kept); a row whose
    status is "verifiziert" is never touched
Then Email Verifier on every address written here: valid → "verifiziert", invalid → "ungültig",
accept_all → "unsicher (accept-all)", anything else → "unsicher (<status>)". An invalid address is
removed from E-Mail and the status "ungültig (Hunter)" is kept; such rows are not retried.
HEIDI_OFFLINE=1 (run_all --dry-run): cache only, no API calls, no credits.
Every Hunter response is cached in data/cache/hunter_<sha1>.json (key stripped), so reruns cost nothing.
HTTP 401/403/429 or a quota error stops the API calls (global rule); results so far are saved.

Usage: python3 scripts/hunter_enrich.py [--in …v10.xlsx] [--out …v10b.xlsx] [--max-domains 25]
                                       [--max-verifications 50] [--plan]
  --plan  list the domains that would be searched, no API calls (except the free quota check)
"""
import argparse
import hashlib
import json
import os
import re
import sys
import time
import unicodedata
from copy import copy
from pathlib import Path

import openpyxl
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import CACHE, DATA, ROOT, norm_domain  # noqa: E402

API = "https://api.hunter.io/v2"
HV = "Hausverwaltungen"
PRIOS = ["Enterprise", "A", "B"]
NEW_COLS = ["Hunter-Pattern", "Hunter-Personen"]
TITLES = {"dr", "prof", "dipl", "ing", "kfm", "kffr", "rain", "ra", "herr", "frau", "mba", "gf", "msc", "bsc",
          "dipl-kfm", "dipl-ing", "dipl-kffr", "mag"}
PARTICLES = {"von", "van", "de", "zu", "der", "den", "vom"}


class Stop(Exception):
    pass


class Offline(Exception):
    pass


# ------------------------------------------------------------------ API + cache
def api_key():
    k = os.environ.get("HUNTER_API_KEY")
    env = ROOT / ".env"
    if not k and env.exists():
        for line in env.read_text().splitlines():
            if line.strip().startswith("HUNTER_API_KEY="):
                k = line.split("=", 1)[1].strip().strip('"').strip("'")
    if not k:
        sys.exit("HUNTER_API_KEY fehlt (Umgebung oder .env)")
    return k


QUOTA_HEADERS = {}


def is_cached(endpoint, params):
    ck = hashlib.sha1(f"{endpoint}?{json.dumps(params, sort_keys=True)}".encode()).hexdigest()
    return (CACHE / f"hunter_{ck}.json").exists()


def call(endpoint, params, key, use_cache=True):
    """GET Hunter endpoint; cache by endpoint+params WITHOUT the key. Returns (json, from_cache)."""
    ck = hashlib.sha1(f"{endpoint}?{json.dumps(params, sort_keys=True)}".encode()).hexdigest()
    path = CACHE / f"hunter_{ck}.json"
    if use_cache and path.exists():
        return json.loads(path.read_text(encoding="utf-8"))["body"], True
    if OFFLINE:
        raise Offline(endpoint)
    for attempt in range(4):
        r = requests.get(f"{API}/{endpoint}", params={**params, "api_key": key}, timeout=30)
        for h, v in r.headers.items():
            if any(x in h.lower() for x in ("ratelimit", "quota", "credit", "remaining", "usage")):
                QUOTA_HEADERS[h] = v
        if r.status_code == 202 and endpoint == "email-verifier":  # verification still running
            time.sleep(3 + 3 * attempt)
            continue
        break
    try:
        body = r.json()
    except ValueError:
        body = {"errors": [{"details": r.text[:200]}]}
    if r.status_code in (401, 403, 429) or (r.status_code >= 400 and "usage" in json.dumps(body).lower()):
        raise Stop(f"Hunter {endpoint}: HTTP {r.status_code} {json.dumps(body.get('errors'))[:200]}")
    if r.status_code == 202:
        return {"data": {"status": "pending"}}, False
    if r.status_code == 200 and use_cache:
        CACHE.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({"endpoint": endpoint, "params": params, "body": body}, ensure_ascii=False),
                        encoding="utf-8")
    if r.status_code != 200:
        return {"errors": body.get("errors"), "data": {}}, False
    return body, False


def account(key):
    if OFFLINE:
        return {"credits": 50.0, "searches": 25, "verifications": 50}, "offline"
    body, _ = call("account", {}, key, use_cache=False)
    req = body.get("data", {}).get("requests", {})
    return {k: req.get(k, {}).get("available", 0) - req.get(k, {}).get("used", 0)
            for k in ("credits", "searches", "verifications")}, body.get("data", {}).get("reset_date")


# ------------------------------------------------------------------ names
def fold(s):
    s = (s or "").strip().lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        s = s.replace(a, b)
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()


def split_name(full):
    """'Dipl.-Kfm. Hans Peter Schulz' -> ('hans', 'schulz'); particles kept with the surname."""
    toks = [t.strip(".,") for t in re.split(r"\s+", full or "") if t.strip(".,")]
    toks = [t for t in toks if fold(t).replace(".", "") not in TITLES and not t.endswith(".")]
    if len(toks) < 2:
        return None, None
    first = fold(toks[0])
    i = len(toks) - 1
    while i > 1 and fold(toks[i - 1]) in PARTICLES:
        i -= 1
    last = fold("".join(toks[i:]))
    return re.sub(r"[^a-z-]", "", first), re.sub(r"[^a-z-]", "", last)


def variants(s):
    """'Özcoban' -> {'oezcoban', 'ozcoban'} (Hunter may drop the umlaut instead of folding it)."""
    plain = unicodedata.normalize("NFKD", (s or "").strip().lower()).encode("ascii", "ignore").decode()
    return {re.sub(r"[^a-z-]", "", fold(s)), re.sub(r"[^a-z-]", "", plain.replace("ß", "ss"))} - {""}


def same_person(person, ansprechpartner):
    toks = [t for t in re.split(r"\s+", ansprechpartner or "") if t]
    f0, l0 = split_name(ansprechpartner)
    if not l0:
        return False
    lasts = set().union(*(variants(t) for t in toks[-1:])) | {l0}
    firsts = set().union(*(variants(t) for t in toks if split_name(ansprechpartner)[0] in variants(t) or
                           fold(t) == f0)) | {f0}
    pls, pfs = variants(person.get("last_name")), variants(person.get("first_name"))
    if not pls or not any(pl == la or la.endswith(pl) or pl.endswith(la) for pl in pls for la in lasts):
        return False
    return not pfs or any(pf == fi or (len(pf) <= 2 and pf[0] == fi[0]) for pf in pfs for fi in firsts)


def build(pattern, ansprechpartner, domain):
    first, last = split_name(ansprechpartner)
    if not first or not last or not pattern:
        return None
    local = (pattern.replace("{first}", first).replace("{last}", last)
             .replace("{f}", first[0]).replace("{l}", last[0]))
    if "{" in local or not re.fullmatch(r"[a-z0-9._-]+", local):
        return None
    return f"{local}@{domain}"


def status_from(v):
    s = (v.get("data") or {}).get("status")
    return {"valid": "verifiziert", "invalid": "ungültig (Hunter)", "accept_all": "unsicher (accept-all)"}.get(
        s, f"unsicher ({s or 'unbekannt'})")


OFFLINE = os.environ.get("HEIDI_OFFLINE") == "1"  # cache only, never call the API (run_all --dry-run)


# ------------------------------------------------------------------ main
def header_map(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, default=DATA / "heidi_contacts_v10.xlsx")
    ap.add_argument("--out", type=Path, default=DATA / "heidi_contacts_v10b.xlsx")
    ap.add_argument("--max-domains", type=int, default=25)
    ap.add_argument("--max-verifications", type=int, default=50)
    ap.add_argument("--plan", action="store_true")
    a = ap.parse_args()
    if a.out.resolve() in (a.inp.resolve(), (DATA / "heidi_merged_v3.xlsx").resolve()):
        sys.exit(f"Refusing to write {a.out}")
    if not a.inp.exists():
        sys.exit(f"Eingabe fehlt: {a.inp}")
    key = "" if OFFLINE else api_key()

    wb = openpyxl.load_workbook(a.inp)
    ws = wb[HV]
    hm = header_map(ws)
    tmpl = ws.cell(1, ws.max_column)
    for n in NEW_COLS + ["E-Mail", "E-Mail-Status", "Kontakt-Quelle"]:
        if n not in hm:
            c = ws.cell(1, ws.max_column + 1, n)
            c.font, c.fill, c.alignment = copy(tmpl.font), copy(tmpl.fill), copy(tmpl.alignment)
            hm[n] = c.column
    cell = lambda r, k: ws.cell(r, hm[k])  # noqa: E731

    rows = []
    for r in range(2, ws.max_row + 1):
        prio, web = cell(r, "Prio").value, norm_domain(cell(r, "Website").value)
        if prio in PRIOS and web:
            rows.append((PRIOS.index(prio), -(cell(r, "Punkte").value or 0), r, web, prio))
    rows.sort()
    targets, seen = [], set()
    for _, _, r, dom, prio in rows:
        if dom not in seen:
            seen.add(dom)
            targets.append((r, dom, prio))

    left, reset = account(key)
    # cached domains below the cut stay in (no credits): a reshuffled ranking never drops earlier results
    sel = targets[:a.max_domains] + [t for t in targets[a.max_domains:]
                                     if is_cached("domain-search", {"domain": t[1], "limit": 10})]
    n_cached = sum(is_cached("domain-search", {"domain": d, "limit": 10}) for _, d, _ in sel)
    budget = min(int(left["searches"]), int(left["credits"]))  # only uncached searches cost credits
    print(f"Hunter-Kontingent vor dem Lauf: {left} (Reset {reset}) → {len(sel)} von {len(targets)} Kandidaten "
          f"(max {a.max_domains}), davon {n_cached} aus Cache, {len(sel) - n_cached} neu (Budget {budget})")
    if a.plan:
        for r, dom, prio in targets[:a.max_domains]:
            print(f"  {prio:10s} {cell(r, 'Punkte').value:>3}  {dom:35s} {cell(r, 'Firma').value} | "
                  f"Ansprechpartner: {cell(r, 'Ansprechpartner').value or '–'} | E-Mail: {cell(r, 'E-Mail').value or '–'}")
        return

    stats = {p: {"Domains": 0, "Pattern": 0, "E-Mails geschrieben": 0, "verifiziert": 0, "ungültig": 0,
                 "unsicher": 0} for p in PRIOS}
    written, stopped = [], None
    try:
        for r, dom, prio in sel:
            if not is_cached("domain-search", {"domain": dom, "limit": 10}):
                if budget <= 0:
                    print(f"  Kontingent erschöpft – {dom} nicht gesucht")
                    continue
                budget -= 1
            try:
                body, cached = call("domain-search", {"domain": dom, "limit": 10}, key)
            except Offline:
                print(f"  offline: {dom} nicht im Cache – übersprungen")
                continue
            d = body.get("data") or {}
            stats[prio]["Domains"] += 1
            pattern = d.get("pattern")
            people = [e for e in d.get("emails") or [] if e.get("first_name") or e.get("last_name")]
            if pattern:
                stats[prio]["Pattern"] += 1
                if not cell(r, "Hunter-Pattern").value:
                    cell(r, "Hunter-Pattern").value = pattern
            if people and not cell(r, "Hunter-Personen").value:
                cell(r, "Hunter-Personen").value = "; ".join(
                    f"{p.get('first_name') or ''} {p.get('last_name') or ''} ({p.get('position') or '–'}) "
                    f"{p.get('value')} {p.get('confidence')}%".strip() for p in people)
            status = str(cell(r, "E-Mail-Status").value or "")
            if status.startswith("verifiziert") or status.startswith("ungültig") or cell(r, "E-Mail").value:
                continue  # never overwrite a verified (or any existing) address; don't retry an invalid one
            ap_name = cell(r, "Ansprechpartner").value
            if not ap_name:
                continue
            match = next((p for p in people if p.get("value") and same_person(p, ap_name)), None)
            if match:
                addr, st = match["value"].lower(), f"Hunter {match.get('confidence')}%"
            else:
                addr, st = build(pattern, ap_name, dom), "Muster (Hunter-Pattern)"
            if addr:
                cell(r, "E-Mail").value = addr
                cell(r, "E-Mail-Status").value = st
                if not cell(r, "Kontakt-Quelle").value:
                    cell(r, "Kontakt-Quelle").value = f"Hunter Domain Search ({dom})"
                stats[prio]["E-Mails geschrieben"] += 1
                written.append((r, addr, prio))
        left, _ = account(key)
        vbudget = min(int(left["verifications"]), int(left["credits"] * 2))
        n_ver = min(a.max_verifications, len(written))
        for r, addr, prio in written[:n_ver]:
            if not is_cached("email-verifier", {"email": addr}):
                if vbudget <= 0:
                    print(f"  Kontingent erschöpft – {addr} nicht verifiziert")
                    continue
                vbudget -= 1
            try:
                v, _ = call("email-verifier", {"email": addr}, key)
            except Offline:
                print(f"  offline: Verifizierung {addr} nicht im Cache")
                continue
            st = status_from(v)
            cell(r, "E-Mail-Status").value = st
            if st.startswith("ungültig"):
                cell(r, "E-Mail").value = None  # keep the check visible in the status, drop the address
            stats[prio][st.split(" ")[0]] += 1
        if len(written) > n_ver:
            print(f"Hinweis: {len(written) - n_ver} geschriebene Adressen nicht verifiziert (Kontingent)")
    except Stop as e:
        stopped = str(e)
    except requests.RequestException as e:
        stopped = f"Netzwerkfehler: {type(e).__name__}"

    wb.save(a.out)
    print("\nPrio       | Domains | Pattern | E-Mails geschrieben | verifiziert | ungültig | unsicher")
    for p in PRIOS:
        s = stats[p]
        print(f"{p:10s} | {s['Domains']:7d} | {s['Pattern']:7d} | {s['E-Mails geschrieben']:19d} | "
              f"{s['verifiziert']:11d} | {s['ungültig']:8d} | {s['unsicher']:8d}")
    if stopped:
        print(f"\nGESTOPPT: {stopped}")
    print(f"\nKontingent-Header aus den Antworten: {QUOTA_HEADERS or 'keine gesendet'}")
    try:
        left, reset = account(key)
        print(f"Restkontingent laut /v2/account: {left} (Reset {reset})" if not OFFLINE
              else "Restkontingent: offline (nur Cache, keine Credits verbraucht)")
    except (Stop, requests.RequestException) as e:
        print(f"Restkontingent nicht abrufbar: {e}")
    print(f"Geschrieben: {a.out}")


if __name__ == "__main__":
    main()
