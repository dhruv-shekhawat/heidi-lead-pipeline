#!/usr/bin/env python3
"""Extraction agent: Claude reads the cached homepage + Impressum of every Hausverwaltungen row and
returns the fields as JSON with an exact quote and URL per value; a second call verifies each value.

Model claude-haiku-5-5, adaptive thinking at effort "low", JSON via output_config.format (structured
outputs). No temperature: claude-haiku-5-5 rejects non-default sampling parameters with a 400 – runs
are reproducible through the response cache instead (data/cache/llm_<sha1>.json, key = model +
prompt text + input; editing a prompt therefore invalidates its cache).

Rows: every Hausverwaltungen row with a Website whose homepage AND Impressum are in the page cache
(pages are read cache-only, never fetched here). Includes the A rows light_enrich skips.
Per row:
1. extract (scripts/prompts/extract_<v>.md): Firma, Verwaltungsart, Einheiten, Software,
   Kundenportal, Messdienst, Fernablesung, VDIV, Ansprechpartner, Funktion, Typ_Verwalter; every
   non-null value carries "quote" + "url"; null when not stated.
2. quote check (code): the quote must occur in the text of the named page (whitespace/quote marks
   normalised) – otherwise "verworfen: Zitat nicht wörtlich im Seitentext", no verifier call.
3. verify (scripts/prompts/verify_<v>.md): value + quote + page text -> "bestätigt"/"verworfen" +
   one-line reason. Verworfen -> value dropped, Log row (Bereich "Agent verworfen").
Writing (confirmed values only):
* A rows (Recherchetiefe A…, hand-checked or rule-approved): only empty cells ("unbekannt" counts
  as empty) are filled; a differing existing value is kept and recorded in "Agent-Abweichung".
* other rows: the agent value replaces the regex value (old value -> "Agent-Abweichung"), except for the
  RULES_FIRST fields (Software, chosen on the holdout), where an existing rule value is kept; a null
  agent value keeps the regex value. Firma: replaced (old name -> Name alt) only if Name alt is
  empty and the names share a distinctive 4+ letter word (light_enrich's rule), otherwise only noted. Einheiten -> Einheiten-Quelle "URL – „quote“". Recherchetiefe -> "B+ (Agent)".
* Typ_Verwalter "nein" -> Hinweis "Agent: kein Verwalter, prüfen (<reason>)"; never auto-excluded,
  and no other field is written for that row (the website may belong to someone else).
* Einheiten >= the Enterprise threshold (5,000) written by the agent: Hinweis "Agent-Zahl, prüfen
  (Einheiten N)"; the number stays in the cell, score.py treats it as empty until it is verified
  (a manual_fixes Einheiten entry removes the marker).
* "Agent-Quote": "Feld: „quote“ (url)" for every confirmed value.
Same-name/same-PLZ duplicates created by a rename are merged afterwards (add_sources.merge_same_name_rows).
manual_fixes.py runs after this step, so hand-checked corrections always win.
The regex values stay available: the workbook passed in (and data/heidi_scored_v14.xlsx, built
without --agent) is the regex output; --report writes regex vs agent per row and field.

Key: ANTHROPIC_API_KEY from the environment or .env (git-ignored). HEIDI_OFFLINE=1 (run_all
--dry-run): cached responses only, no API calls.

Usage: python3 scripts/llm_enrich.py --in IN.xlsx --out OUT.xlsx [--prompt v2] [--report CSV]
                                     [--only DOMAIN …] [--workers 8]
"""
import argparse
import csv
import hashlib
import json
import os
import re
import sys
import threading
from concurrent.futures import ThreadPoolExecutor
from copy import copy
from pathlib import Path

OFFLINE_LLM = os.environ.get("HEIDI_OFFLINE") == "1"
os.environ["HEIDI_OFFLINE"] = "1"  # pages: cache only, this step never fetches websites

import openpyxl  # noqa: E402
from bs4 import BeautifulSoup  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402
import websites as site  # noqa: E402
from light_enrich import words4  # noqa: E402 – same rename guard as light_enrich
import score  # noqa: E402

# agent Einheiten >= this (Gewichte prio_enterprise default) are marked AGENT_EINHEITEN_MARK and not scored
ENTERPRISE_MIN = next(t for k, _, t, _ in score.DEFAULTS if k == "prio_enterprise")

HV, LOG = "Hausverwaltungen", "Log"
MODEL = "claude-haiku-5-5"
PRICE_IN, PRICE_OUT = 0.10, 0.50  # $ per MTok, prompts <= 100K tokens (pages are capped well below)
PAGE_CAP = 60000  # characters per page text sent to the model (truncation is reported)
PROMPTS = Path(__file__).resolve().parent / "prompts"
NEW_COLS = ["Agent-Abweichung", "Agent-Quote"]
# input options per prompt version (v1 is kept as recorded so its cached responses stay valid)
VERSION_OPTS = {"v1": {"links": False, "reason_first": False}}
DEFAULT_OPTS = {"links": True, "reason_first": True}  # v2+: "Links" list per page, verifier reasons first
MAX_LINKS = 120
# agent field -> workbook column
COLMAP = {"Firma": "Firma", "Verwaltungsart": "Verwaltungsart", "Einheiten": "Einheiten", "Software": "Software",
          "Kundenportal": "Kundenportal", "Messdienst": "Messdienst", "Fernablesung": "Fernablesung erwähnt",
          "VDIV": "VDIV-Logo", "Ansprechpartner": "Ansprechpartner", "Funktion": "Funktion"}
FIELDS = list(COLMAP) + ["Typ_Verwalter"]
# Source per field, decided on the 25-row holdout by the rule pre-registered in doc/eval.md (2026-11-09):
# the agent wins a field only with >= 2 fewer errors than the rules. Rules won Software (3 vs 2 errors) and
# Typ_Verwalter (0 vs 0). For a RULES_FIRST field the agent only fills an empty cell; an existing rule value is
# kept and the agent value goes to Agent-Abweichung. Typ_Verwalter has no cell: exclusion stays with the rules
# (lodging rule), the agent's "nein" stays a review flag. Unscored fields (Funktion, Messdienst, Fernablesung,
# VDIV) keep the agent-first behaviour.
RULES_FIRST = {"Software"}


def _claim(value_schema, extra=None):
    props = {"value": value_schema, "quote": {"type": "string"}, "url": {"type": "string"}}
    props.update(extra or {})
    return {"anyOf": [{"type": "null"}, {"type": "object", "properties": props, "required": list(props),
                                          "additionalProperties": False}]}


S = {"type": "string"}
EXTRACT_SCHEMA = {"type": "object", "additionalProperties": False, "required": FIELDS, "properties": {
    "Firma": _claim(S), "Verwaltungsart": _claim({"type": "string", "enum": ["WEG", "Miet", "beides"]}),
    "Einheiten": _claim({"type": "integer"}), "Software": _claim(S),
    "Kundenportal": _claim({"type": "string", "enum": ["ja"]}), "Messdienst": _claim(S),
    "Fernablesung": _claim({"type": "string", "enum": ["ja"]}), "VDIV": _claim({"type": "string", "enum": ["ja"]}),
    "Ansprechpartner": _claim(S), "Funktion": _claim(S),
    "Typ_Verwalter": _claim({"type": "string", "enum": ["ja", "nein"]}, {"reason": S})}}
def verify_schema(reason_first):
    """v1: verdict before reason. v2+: reason first – the model writes its reasoning before it commits
    to a verdict (v1 often wrote "daher bestätigt" and then returned "verworfen")."""
    props = {"field": {"type": "string", "enum": FIELDS}}
    order = (("reason", S), ("verdict", {"type": "string", "enum": ["bestätigt", "verworfen"]}))
    for k, v in (order if reason_first else order[::-1]):
        props[k] = v
    return {"type": "object", "additionalProperties": False, "required": ["checks"], "properties": {
        "checks": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                              "required": list(props), "properties": props}}}}

_client = None
_lock = threading.Lock()
COST = {"calls": 0, "cached": 0, "in": 0, "out": 0, "in_new": 0, "out_new": 0}


class Offline(Exception):
    pass


def api_key():
    if os.environ.get("ANTHROPIC_API_KEY"):
        return os.environ["ANTHROPIC_API_KEY"]
    env = common.ROOT / ".env"
    if env.exists():
        for ln in env.read_text().splitlines():
            if ln.startswith("ANTHROPIC_API_KEY="):
                return ln.split("=", 1)[1].strip().strip("'\"")
    sys.exit("ANTHROPIC_API_KEY fehlt (Umgebung oder .env)")


def client():
    global _client
    if _client is None:
        import anthropic
        _client = anthropic.Anthropic(api_key=api_key(), max_retries=4)
    return _client


def llm(kind, system, user, schema):
    """One structured-output call, cached on disk. Returns parsed JSON."""
    key = hashlib.sha1(json.dumps([MODEL, kind, system, user, schema], ensure_ascii=False,
                                  sort_keys=True).encode()).hexdigest()
    path = common.CACHE / f"llm_{key}.json"
    if path.exists():
        e = json.loads(path.read_text(encoding="utf-8"))
        with _lock:
            COST["calls"] += 1
            COST["cached"] += 1
            COST["in"] += e["usage"]["input_tokens"]
            COST["out"] += e["usage"]["output_tokens"]
        return json.loads(e["text"])
    if OFFLINE_LLM:
        raise Offline("nicht im LLM-Cache")
    r = client().messages.create(model=MODEL, max_tokens=8000, system=system,
                                 output_config={"effort": "low", "format": {"type": "json_schema", "schema": schema}},
                                 messages=[{"role": "user", "content": user}])
    if r.stop_reason != "end_turn":
        raise RuntimeError(f"stop_reason {r.stop_reason}")
    text = next(b.text for b in r.content if b.type == "text")
    json.loads(text)  # must parse before it is cached
    u = {"input_tokens": r.usage.input_tokens, "output_tokens": r.usage.output_tokens}
    path.write_text(json.dumps({"model": MODEL, "kind": kind, "text": text, "usage": u}, ensure_ascii=False),
                    encoding="utf-8")
    with _lock:
        COST["calls"] += 1
        for k, uk in (("in", "input_tokens"), ("out", "output_tokens")):
            COST[k] += u[uk]
            COST[k + "_new"] += u[uk]
    return json.loads(text)


def cost(i, o):
    return i / 1e6 * PRICE_IN + o / 1e6 * PRICE_OUT


def norm_q(s):
    s = (s or "").replace("\xa0", " ")
    s = re.sub(r"[„“”\"»«‚‘’']", "'", s)
    s = re.sub(r"[–—‐-]", "-", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def pages_of(website):
    home, err = site.homepage(website)
    if not home:
        return None, f"Startseite nicht im Cache ({err})"
    imp, _, why = site.impressum(home)
    if not imp:
        return None, f"Impressum nicht im Cache ({why})"
    out = []
    for p in (home, imp):
        t = BeautifulSoup(p.get("text", ""), "lxml").title
        title = re.sub(r"\s+", " ", t.get_text(" ")).strip() if t else ""
        text = common.page_text(p.get("text", ""))
        links = []
        for lt, u in site.links(p):
            if u.lower().startswith(("http://", "https://")):
                label = re.sub(r"\s+", " ", lt).strip()
                ln = f"Link: {label} -> {u}"
                if ln not in links:
                    links.append(ln)
        out.append({"url": site.page_url(p), "title": title, "text": text[:PAGE_CAP],
                    "truncated": len(text) > PAGE_CAP, "links": links[:MAX_LINKS]})
    return out, None


def page_block(pages, links=False):
    return "\n\n".join(f"=== URL: {p['url']}\nSeitentitel: {p['title']}\nSeitentext"
                       f"{' (gekürzt)' if p['truncated'] else ''}:\n{p['text']}"
                       + ("\nLinks:\n" + "\n".join(p["links"]) if links and p["links"] else "") for p in pages)


def run_row(row, version):
    res = {"row": row, "claims": {}, "checks": {}, "err": None, "pages": None}
    pages, err = pages_of(row["Website"])
    if not pages:
        res["err"] = err
        return res
    res["pages"] = pages
    opts = VERSION_OPTS.get(version, DEFAULT_OPTS)
    sys_x = (PROMPTS / f"extract_{version}.md").read_text(encoding="utf-8")
    sys_v = (PROMPTS / f"verify_{version}.md").read_text(encoding="utf-8")
    try:
        data = llm("extract", sys_x, page_block(pages, opts["links"]) + "\n\nGib die Felder als JSON zurück.", EXTRACT_SCHEMA)
        claims = {f: v for f, v in data.items() if v and v.get("value") not in (None, "")}
        res["claims"] = claims
        texts = {p["url"]: norm_q(p["title"] + "\n" + p["text"] + ("\n" + "\n".join(p["links"]) if opts["links"] else ""))
                 for p in pages}
        to_verify = {}
        for f, c in claims.items():
            body = texts.get(c.get("url")) or " ".join(texts.values())
            parts = [q for q in (c.get("quote") or "").split("||") if q.strip()]
            if not parts or any(norm_q(q) not in body for q in parts):
                res["checks"][f] = ("verworfen", "Zitat nicht wörtlich im Seitentext")
            else:
                to_verify[f] = c
        if to_verify:
            listing = json.dumps([{"field": f, "value": c["value"], "quote": c["quote"], "url": c["url"]}
                                  for f, c in to_verify.items()], ensure_ascii=False, indent=1)
            v = llm("verify", sys_v, page_block(pages, opts["links"]) + "\n\nZu prüfende Angaben:\n" + listing,
                    verify_schema(opts["reason_first"]))
            for chk in v.get("checks", []):
                if chk["field"] in to_verify:
                    res["checks"][chk["field"]] = (chk["verdict"], chk["reason"])
            for f in to_verify:  # a claim the verifier skipped is not confirmed
                res["checks"].setdefault(f, ("verworfen", "Prüfer hat die Angabe nicht bewertet"))
    except Offline as e:
        res["err"] = f"offline: {e}"
    except Exception as e:  # noqa: BLE001 – API/parse errors fail the row, not the run
        import anthropic
        if isinstance(e, (anthropic.AuthenticationError, anthropic.PermissionDeniedError)):
            raise
        res["err"] = f"{type(e).__name__}: {str(e)[:120]}"
    return res


def empty(v):
    return v is None or str(v).strip() == "" or str(v).strip().lower() == "unbekannt"


def surname(name):
    toks = [t for t in re.findall(r"[a-zäöüß]+", (name or "").lower())
            if t not in {"dipl", "ing", "kfm", "kff", "dr", "prof", "herr", "frau", "betriebsw", "fh"}]
    return toks[-1] if toks else ""


def same(field, a, b):
    if empty(a) or empty(b):
        return empty(a) and empty(b)
    if field == "Firma":
        return common.norm_name(a) == common.norm_name(b)
    if field == "Einheiten":
        try:
            return int(str(a).replace(".", "")) == int(str(b).replace(".", ""))
        except ValueError:
            return False
    if field in ("Software", "Messdienst"):
        sp = lambda s: {x.strip().lower() for x in re.split(r"[,;/]| und ", str(s)) if x.strip()}  # noqa: E731
        return sp(a) == sp(b)
    if field == "Ansprechpartner":
        return surname(a) == surname(b)
    return str(a).strip().lower() == str(b).strip().lower()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--prompt", default="v2", help="prompt version in scripts/prompts/ (extract_<v>.md, verify_<v>.md)")
    ap.add_argument("--report", type=Path, default=None, help="CSV: regex/table value vs agent value per field")
    ap.add_argument("--only", nargs="+", default=None, metavar="DOMAIN")
    ap.add_argument("--workers", type=int, default=8)
    a = ap.parse_args()
    if a.out.resolve() in (a.inp.resolve(), (common.DATA / "heidi_merged_v3.xlsx").resolve()):
        sys.exit(f"Refusing to write {a.out}")
    wb = openpyxl.load_workbook(a.inp)
    ws, log = wb[HV], wb[LOG]
    hm = {c.value: c.column for c in ws[1] if c.value}
    lm = {c.value: c.column for c in log[1] if c.value}
    tmpl = ws.cell(1, ws.max_column)
    for n in NEW_COLS:
        if n not in hm:
            c = ws.cell(1, ws.max_column + 1, n)
            c.font, c.fill, c.alignment = copy(tmpl.font), copy(tmpl.fill), copy(tmpl.alignment)
            hm[n] = c.column
    cell = lambda r, k: ws.cell(r, hm[k])  # noqa: E731
    only = {common.norm_domain(d) for d in a.only} if a.only else None
    rows = [{"r": r, "Website": cell(r, "Website").value, "Firma": cell(r, "Firma").value,
             "tier": str(cell(r, "Recherchetiefe").value or "")}
            for r in range(2, ws.max_row + 1)
            if cell(r, "Firma").value and cell(r, "Website").value
            and (only is None or common.norm_domain(cell(r, "Website").value) in only)]
    print(f"Agent ({MODEL}, Prompt {a.prompt}{', nur Cache' if OFFLINE_LLM else ''}): {len(rows)} Zeilen mit Website",
          flush=True)
    with ThreadPoolExecutor(max_workers=a.workers) as ex:
        results = list(ex.map(lambda row: run_row(row, a.prompt), rows))

    stats = {"verarbeitet": 0, "ohne Seiten": 0, "Fehler": 0, "bestätigt": 0, "verworfen": 0, "geschrieben": 0,
             "Abweichung (A behalten)": 0, "Regex ersetzt": 0, "kein Verwalter": 0, "Agent-Zahl, prüfen": 0,
             "Abweichung (Regel behalten)": 0}
    report, flagged, errors, marked = [], [], {}, []
    log_rows = []
    for res in results:
        row, r = res["row"], res["row"]["r"]
        if res["err"]:
            k = "ohne Seiten" if "nicht im Cache" in res["err"] and "LLM" not in res["err"] else "Fehler"
            stats[k] += 1
            key = re.sub(r"\(.*", "", res["err"]).strip()
            errors[key] = errors.get(key, 0) + 1
            continue
        stats["verarbeitet"] += 1
        is_a = row["tier"].startswith("A")
        tv = res["claims"].get("Typ_Verwalter")
        # agent says the pages show another business: the site may not be this firm's – flag only, write nothing
        not_hv = bool(tv and tv.get("value") == "nein" and res["checks"].get("Typ_Verwalter", ("",))[0] == "bestätigt")
        quotes, notes = [], []
        for f in FIELDS:
            c = res["claims"].get(f)
            col = COLMAP.get(f)
            before = cell(r, col).value if col else None
            verdict, reason = res["checks"].get(f, (None, None)) if c else (None, None)
            action = ""
            if c and verdict == "verworfen":
                stats["verworfen"] += 1
                log_rows.append({"Bereich": "Agent verworfen", "Firma": row["Firma"], "Website": row["Website"],
                                 "Typ": cell(r, "Typ").value, "Quelle": c.get("url"),
                                 "Grund / Hinweis": f"{f} = {c['value']}: {reason} (Zitat: „{c.get('quote')}“)"})
                action = "verworfen"
            elif c and verdict == "bestätigt":
                stats["bestätigt"] += 1
                val = c["value"]
                quotes.append(f"{f}: „{c['quote']}“ ({c['url']})")
                if f == "Typ_Verwalter":
                    if val == "nein":
                        stats["kein Verwalter"] += 1
                        flagged.append((row["Firma"], row["Website"], c.get("reason")))
                        h = cell(r, "Hinweis").value
                        note = f"Agent: kein Verwalter, prüfen ({c.get('reason')})"
                        cell(r, "Hinweis").value = f"{h} | {note}" if h else note
                        action = "Hinweis"
                elif f == "Funktion":
                    pass  # written together with Ansprechpartner below
                elif not_hv:
                    action = "nicht geschrieben (kein Verwalter)"
                elif same(f, before, val):
                    action = "gleich"
                elif f == "Firma" and not is_a and (cell(r, "Name alt").value or not (words4(val) & words4(before))):
                    # Name alt taken, or no shared distinctive word (Impressum of another company): keep Firma
                    notes.append(f"Firma laut Impressum (Agent), prüfen: {val}")
                    action = "notiert"
                elif empty(before) or (not is_a and f not in RULES_FIRST):
                    if not empty(before):
                        notes.append(f"{col}: Regex {before} → Agent {val}")
                        stats["Regex ersetzt"] += 1
                        if f == "Firma":
                            cell(r, "Name alt").value = before
                    cell(r, col).value = val
                    if f == "Einheiten":
                        cell(r, "Einheiten-Quelle").value = f"{c['url']} – „{c['quote']}“"
                        n_val = score.num(val)
                        if n_val is not None and n_val >= ENTERPRISE_MIN:
                            h = cell(r, "Hinweis").value or ""
                            if common.AGENT_EINHEITEN_MARK not in h:
                                mark = f"{common.AGENT_EINHEITEN_MARK} (Einheiten {score.fmt(n_val)})"
                                cell(r, "Hinweis").value = f"{h} | {mark}" if h else mark
                            stats["Agent-Zahl, prüfen"] += 1
                            marked.append((row["Firma"], row["Website"], score.fmt(n_val)))
                    if f == "Ansprechpartner":
                        fc = res["claims"].get("Funktion")
                        ok = fc and res["checks"].get("Funktion", ("",))[0] == "bestätigt"
                        cell(r, "Funktion").value = fc["value"] if ok else None
                        if "Kontakt-Quelle" in hm:
                            cell(r, "Kontakt-Quelle").value = c["url"]
                    stats["geschrieben"] += 1
                    action = "geschrieben"
                else:
                    notes.append(f"{col}: Tabelle {before} | Agent {val}")
                    stats["Abweichung (A behalten)" if is_a else "Abweichung (Regel behalten)"] += 1
                    action = "Abweichung"
            report.append({"Firma": row["Firma"], "Website": row["Website"], "Recherchetiefe": row["tier"],
                           "Feld": f, "Tabelle/Regex": before, "Agent": (c or {}).get("value"),
                           "Urteil": verdict or "", "Grund": reason or (c or {}).get("reason") or "",
                           "Zitat": (c or {}).get("quote") or "", "URL": (c or {}).get("url") or "",
                           "Aktion": action, "Prompt": a.prompt})
        # Funktion alone (Ansprechpartner unchanged and Funktion empty): fill it
        fc = res["claims"].get("Funktion")
        if fc and not not_hv and res["checks"].get("Funktion", ("",))[0] == "bestätigt" and empty(cell(r, "Funktion").value) \
                and same("Ansprechpartner", cell(r, "Ansprechpartner").value,
                         (res["claims"].get("Ansprechpartner") or {}).get("value")):
            cell(r, "Funktion").value = fc["value"]
        cell(r, "Agent-Quote").value = " | ".join(quotes) or None
        cell(r, "Agent-Abweichung").value = " | ".join(notes) or None
        if not is_a and not not_hv:
            cell(r, "Recherchetiefe").value = "B+ (Agent)"
    for lr in log_rows:
        nr = log.max_row + 1
        for c_ in range(1, log.max_column + 1):
            if log.cell(nr - 1, c_).has_style:
                log.cell(nr, c_)._style = copy(log.cell(nr - 1, c_)._style)
        for k, v in lr.items():
            if k in lm:
                log.cell(nr, lm[k]).value = v
    from add_sources import merge_same_name_rows  # Impressum renames can produce same-name duplicates
    print("Dubletten nach Umbenennung:")
    merge_same_name_rows(ws)
    wb.save(a.out)
    if a.report:
        with open(a.report, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(report[0]) if report else ["Firma"])
            w.writeheader()
            w.writerows(report)
    print("Ergebnis:", stats)
    if errors:
        print("  nicht verarbeitet:", errors)
    print(f"Einheiten ≥ {score.fmt(ENTERPRISE_MIN)} vom Agenten (Hinweis '{common.AGENT_EINHEITEN_MARK}', "
          f"zählen nicht für Enterprise): {len(marked)}")
    for f_, w_, n_ in marked:
        print(f"   {f_} ({w_}): {n_}")
    print(f"Als kein Verwalter markiert (Hinweis, nicht ausgeschlossen): {len(flagged)}")
    for f_, w_, why in flagged:
        print(f"   {f_} ({w_}): {why}")
    print(f"API: {COST['calls']} Antworten, davon {COST['cached']} aus Cache | Tokens gesamt in {COST['in']:,} / "
          f"out {COST['out']:,} = ${cost(COST['in'], COST['out']):.4f} | in diesem Lauf neu bezahlt: "
          f"${cost(COST['in_new'], COST['out_new']):.4f}")
    print(f"Geschrieben: {a.out}" + (f", {a.report}" if a.report else ""))


if __name__ == "__main__":
    main()
