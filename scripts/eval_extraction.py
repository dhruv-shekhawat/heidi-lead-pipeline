#!/usr/bin/env python3
"""Error rate per field on a hand-checked sample: regex vs workbooks vs agent prompts vs ensemble.

Truth: any sample file in the verification_sample format (a Markdown table with a Website column and
one column per field; extra columns are ignored). Default doc/verification_sample.md (30 rows, built
from doc/verification_report.md); the holdout is doc/holdout_blind.md once filled in (both untracked:
firm and person names). Cells "nicht prüfbar" or empty are skipped.
Systems:
* Regex     – light_enrich / extract_contacts / apply_review rules run on the same cached pages
              (Firma = Impressum line with Rechtsform, Einheiten = largest rule-approved candidate,
              Typ_Verwalter = nein if the lodging exclusion rule fires).
* Tabelle   – a workbook (rows moved to the Log as "ausgeschlossen" count as Typ_Verwalter nein) (default data/heidi_scored_v12.xlsx = hand research + regex, the state the
              sample was checked against; its error count must equal the report's bold cells).
* Agent vN  – llm_enrich.run_row with prompt version N (cached responses; only "bestätigt" values).
* Ensemble  – agent (--ensemble version, default v2) for Firma, Verwaltungsart, Einheiten, regex rules for the
              other fields; where the agent has no confirmed value the regex value is kept (as in the pipeline).
Comparison: Firma by normalised name ("und" = "&"); Verwaltungsart exact (empty = unbekannt); Einheiten as
integer ("keine Angabe" = correct only if empty; same for Firma "nicht angegeben"); Ansprechpartner by surname, any listed person counts; Software as a set of names (empty =
keine); Kundenportal ja/nein (empty = nein); Typ_Verwalter ja/nein (empty = ja). A row without
output (no cached Impressum, API error) counts as empty.
Output: aggregate table (stdout and --md, no firm names) + per-row errors in --details
(default data/eval_details.md, untracked).

Usage: python3 scripts/eval_extraction.py [--sample doc/verification_sample.md] --prompts v1 v2
                                          [--table v12=data/heidi_scored_v12.xlsx …] (--table alone: none)
                                          [--ensemble v2] [--md doc/eval_tables.md] [--details data/eval_details.md]
Holdout: python3 scripts/eval_extraction.py --sample doc/holdout_blind.md --prompts v2 --table
"""
import argparse
import os
import re
import sys
from pathlib import Path

LLM_OFFLINE = os.environ.get("HEIDI_OFFLINE") == "1"  # caller's choice applies to LLM calls only
os.environ["HEIDI_OFFLINE"] = "1"  # pages from cache only
sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402

import common  # noqa: E402
import websites as site  # noqa: E402

FIELDS = ["Firma", "Verwaltungsart", "Einheiten", "Ansprechpartner", "Software", "Kundenportal", "Typ_Verwalter"]
AGENT_FIELDS = ("Firma", "Verwaltungsart", "Einheiten")  # ensemble: agent here, regex rules for the rest
NO_FIGURE = ("keine angabe", "keine", "nicht angegeben")


def parse_sample(path):
    lines = [ln for ln in Path(path).read_text(encoding="utf-8").splitlines() if ln.startswith("|")]
    hdr = [c.strip() for c in lines[0].strip("|").split("|")]
    return [dict(zip(hdr, [c.strip() for c in ln.strip("|").split("|")])) for ln in lines[2:]]


def surname(n):
    toks = [t for t in re.findall(r"[a-zäöüß]+", (n or "").lower())
            if t not in {"dipl", "ing", "kfm", "kff", "dr", "prof", "herr", "frau", "betriebsw", "fh"}]
    return toks[-1] if toks else ""


def softset(s):
    s = re.sub(r"(?i)\bother:\s*", "", "" if s is None else str(s))  # table/report notation "other: idwell"
    return {x.strip().lower() for x in re.split(r"[,;/]| und ", s) if x.strip() and x.strip().lower() not in
            ("keine", "unbekannt", "none")}


def correct(field, truth, pred):
    if field == "Firma":
        if truth.strip().lower() in NO_FIGURE:  # Impressum names no firm: only "no value" is correct
            return pred is None or str(pred).strip() == ""
        if truth == "Ferienhaus Berlin":
            return "ferienhausberlin" in common.norm_name(pred)
        und = lambda x: re.sub(r"\s+und\s+", " & ", str(x))  # noqa: E731 – "A und B" == "A & B"
        return bool(pred) and common.norm_name(und(pred)) == common.norm_name(und(truth))
    if field == "Verwaltungsart":
        p = str(pred or "unbekannt").strip() or "unbekannt"
        return p.lower() == truth.lower()
    if field == "Einheiten":
        if truth.strip().lower() in NO_FIGURE:  # site states no figure: any number is an error
            return pred is None or str(pred).strip() == ""
        try:
            return int(str(pred).replace(".", "")) == int(truth)
        except (TypeError, ValueError):
            return False
    if field == "Ansprechpartner":
        names = [x for x in re.split(r",| und ", truth) if x.strip()]
        return bool(pred) and surname(pred) in {surname(x) for x in names}
    if field == "Software":
        return softset(pred) == softset(truth)
    if field == "Kundenportal":
        return ("ja" if str(pred or "").lower().startswith("ja") else "nein") == truth
    if field == "Typ_Verwalter":
        return ("nein" if str(pred or "").lower() == "nein" else "ja") == truth
    raise KeyError(field)


def regex_values(dom):
    import extract_contacts as ec
    import light_enrich as le
    from apply_review import NOT_MANAGED, judge
    res = le.analyse({"r": 0, "Website": dom, "Firma": "", "PLZ": ""})
    if res["status"] != "ok":
        return {}
    home, _ = site.homepage(dom)
    imp, _, _ = site.impressum(home)
    name = ec.find_person(common.html_lines(imp["text"]))[0] if imp else None
    ok = []
    for v, phrase, _ in res["einheiten"]:
        good, sent = judge(v, phrase)
        if good and not NOT_MANAGED.search(sent):
            ok.append(v)
    return {"Firma": res["firma_cand"], "Verwaltungsart": res["found"].get("Verwaltungsart"),
            "Einheiten": max(ok) if ok else None, "Ansprechpartner": name,
            "Software": res["found"].get("Software"), "Kundenportal": res["found"].get("Kundenportal"),
            "Typ_Verwalter": "nein" if res["exclude"] else "ja"}


def table_values(path):
    out = {}
    for tab in ("Hausverwaltungen",):
        ws = openpyxl.load_workbook(path, read_only=True)[tab]
        h = [c.value for c in next(ws.iter_rows(max_row=1))]
        for v in ws.iter_rows(min_row=2, values_only=True):
            r = dict(zip(h, v))
            if r.get("Website"):
                d = common.norm_domain(r["Website"])
                out[d] = {f: r.get(f) for f in FIELDS[:-1]}
                out[d]["Typ_Verwalter"] = "nein" if r.get("Typ") == "ausschließen" or \
                    "kein Verwalter" in str(r.get("Hinweis") or "") else "ja"
    # rows the pipeline excluded live in the Log (Bereich "ausgeschlossen"): Typ_Verwalter = nein
    ws = openpyxl.load_workbook(path, read_only=True)["Log"]
    h = [c.value for c in next(ws.iter_rows(max_row=1))]
    for v in ws.iter_rows(min_row=2, values_only=True):
        r = dict(zip(h, v))
        d = common.norm_domain(r.get("Website"))
        if d and d not in out and r.get("Bereich") == "ausgeschlossen":
            out[d] = {"Firma": r.get("Firma"), "Typ_Verwalter": "nein"}
    return out


def agent_values(dom, version):
    import llm_enrich as le
    le.OFFLINE_LLM = LLM_OFFLINE
    res = le.run_row({"r": 0, "Website": dom, "Firma": "", "tier": "B"}, version)
    if res["err"]:
        return {}, res["err"]
    vals = {f: c["value"] for f, c in res["claims"].items() if res["checks"].get(f, ("",))[0] == "bestätigt"}
    return vals, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sample", default="doc/verification_sample.md")
    ap.add_argument("--prompts", nargs="*", default=["v1", "v2"])
    ap.add_argument("--table", nargs="*", default=["v12=data/heidi_scored_v12.xlsx"])
    ap.add_argument("--ensemble", default="v2", help="agent version for the Ensemble column ('' = none)")
    ap.add_argument("--md", type=Path, default=None)
    ap.add_argument("--details", type=Path, default=common.DATA / "eval_details.md")
    a = ap.parse_args()
    sample = parse_sample(a.sample)
    systems = {"Regex": {}}
    notes = {}
    for s in sample:
        systems["Regex"][s["Website"]] = regex_values(s["Website"])
    for t in a.table:
        name, path = t.split("=", 1)
        tv = table_values(path)
        systems[f"Tabelle {name}"] = {s["Website"]: tv.get(s["Website"], {}) for s in sample}
    for v in a.prompts:
        systems[f"Agent {v}"] = {}
        for s in sample:
            vals, err = agent_values(s["Website"], v)
            systems[f"Agent {v}"][s["Website"]] = vals
            if err:
                notes.setdefault(f"Agent {v}", []).append(f"{s['Website']}: {err}")
    if a.ensemble:
        if a.ensemble not in a.prompts:
            sys.exit(f"--ensemble {a.ensemble} needs --prompts … {a.ensemble}")
        ag = systems[f"Agent {a.ensemble}"]
        systems["Ensemble"] = {}
        for s in sample:
            rx, av = systems["Regex"][s["Website"]], ag[s["Website"]]
            systems["Ensemble"][s["Website"]] = {
                f: (av.get(f) if f in AGENT_FIELDS and av.get(f) not in (None, "") else rx.get(f)) for f in FIELDS}
    import llm_enrich
    cost = llm_enrich.cost(llm_enrich.COST["in"], llm_enrich.COST["out"])

    names = list(systems)
    agg = {n: {f: [0, 0] for f in FIELDS} for n in names}
    details = ["# Eval details (untracked: firm names)", ""]
    for n in names:
        details += [f"## {n}", "", "| Website | Feld | wahr | geliefert |", "|---|---|---|---|"]
        for s in sample:
            vals = systems[n][s["Website"]]
            for f in FIELDS:
                truth = s.get(f, "nicht prüfbar")
                if truth in ("nicht prüfbar", ""):
                    continue
                agg[n][f][1] += 1
                if not correct(f, truth, vals.get(f)):
                    agg[n][f][0] += 1
                    details.append(f"| {s['Website']} | {f} | {truth} | {vals.get(f)} |")
        details.append("")
    L = ["| Feld | geprüft | " + " | ".join(names) + " |", "|---|---|" + "---|" * len(names)]
    for f in FIELDS:
        n_chk = agg[names[0]][f][1]
        L.append(f"| {f} | {n_chk} | " + " | ".join(
            f"{agg[n][f][0]} ({agg[n][f][0] / n_chk:.0%})" if n_chk else "–" for n in names) + " |")
    tot = {n: [sum(agg[n][f][0] for f in FIELDS), sum(agg[n][f][1] for f in FIELDS)] for n in names}
    L.append(f"| **gesamt** | {tot[names[0]][1]} | " + " | ".join(
        f"**{tot[n][0]} ({tot[n][0] / tot[n][1]:.1%})**" if tot[n][1] else "–" for n in names) + " |")
    print(f"Stichprobe: {a.sample} ({len(sample)} Zeilen)")
    print("Fehler je Feld (Fehler / geprüfte Zellen):")
    print("\n".join(L))
    berlin = [s.get("Einheiten nur Berlin", "").strip().lower() for s in sample if "Einheiten nur Berlin" in s]
    if any(berlin):  # information only: no system distinguishes Berlin-only figures
        print("Einheiten nur Berlin (nicht bewertet): " + ", ".join(
            f"{k} {berlin.count(k)}" for k in sorted(set(berlin)) if k))
    for n, ns in notes.items():
        print(f"{n}: ohne Ausgabe {len(ns)} Zeilen ({'; '.join(sorted({re.sub(r'.*: ', '', x) for x in ns}))})")
    print(f"LLM-Kosten der Eval-Antworten (Cache-Summe): ${cost:.4f}")
    a.details.write_text("\n".join(details) + "\n", encoding="utf-8")
    if a.md:
        a.md.write_text("\n".join(L) + "\n", encoding="utf-8")
    print(f"Details: {a.details}")


if __name__ == "__main__":
    main()
