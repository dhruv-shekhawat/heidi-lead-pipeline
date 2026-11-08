#!/usr/bin/env python3
"""Select the holdout sample for the extraction eval and write a blind checking sheet.

Rows (none of them among the 30 already verified in doc/verification_sample.md, matched by domain):
1. every row whose Einheiten figure >= 5,000 comes from the agent (Hinweis "Agent-Zahl, prüfen" in the
   current workbook; these were Enterprise in v15),
2. every row that is Prio A in the current workbook but below A in the regex-only workbook (v14), i.e.
   reached A only through agent values (matched by domain, because the agent renamed some firms),
3. random "B+ (Agent)" rows up to --n (pool sorted by domain, random.Random(seed).sample).
The blind sheet is ordered by a shuffled number so the group is not visible; it lists only Firma and
Website plus empty columns for the fields to determine. Group, seed and source workbooks go to
data/holdout_keys.json. Each row is checked to have cached pages and cached agent responses, so the
later eval scores exactly the agent output the workbook contains, without new API calls.

Outputs (both untracked: firm names): doc/holdout_blind.md, data/holdout_keys.json.
Usage: python3 scripts/holdout_select.py [--current data/heidi_scored_v16.xlsx] [--regex data/heidi_scored_v14.xlsx]
                                         [--n 25] [--seed 42]
"""
import argparse
import json
import os
import random
import sys
from pathlib import Path

os.environ["HEIDI_OFFLINE"] = "1"  # pages from cache only
sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl  # noqa: E402

import common  # noqa: E402
from eval_extraction import parse_sample  # noqa: E402

BLIND_COLS = ["Nr", "Firma (Liste)", "Website", "Firma", "Verwaltungsart", "Einheiten", "Einheiten nur Berlin",
              "Ansprechpartner", "Software", "Kundenportal", "Typ_Verwalter"]
HEADER = """# Holdout sample (25 rows, blind)

**Not versioned** (firm names; see .gitignore). Selection: `scripts/holdout_select.py`, keys in
`data/holdout_keys.json`. Pre-registered decision rule: `doc/eval.md`. Scored with
`python3 scripts/eval_extraction.py --sample doc/holdout_blind.md --prompts v2 --table`.

Check each firm on its own website (all pages, not only homepage and Impressum) without looking at the
workbook. Fill every empty cell; the columns use the conventions of `doc/verification_sample.md`:

- `Firma`: legal name with Rechtsform as written in the Impressum.
- `Verwaltungsart`: `WEG`, `Miet`, `beides` or `unbekannt` (site says neither).
- `Einheiten`: integer the firm says it manages ("rund 2.000" = 2000); `keine Angabe` if no figure is stated.
- `Einheiten nur Berlin`: `ja` if the figure counts Berlin units only, `nein` if it includes other cities
  or is nationwide, `unklar` if the site doesn't say; empty when Einheiten is `keine Angabe`. Information
  only, not scored.
- `Ansprechpartner`: first person named as Geschäftsführer/in, Inhaber/in or vertretungsberechtigt in the
  Impressum; several equally valid names comma-separated.
- `Software`: property-management or portal software named on the site (comma-separated), `keine` if none.
- `Kundenportal`: `ja` / `nein` (online portal or app for owners or tenants).
- `Typ_Verwalter`: `ja` if the firm manages property for others, `nein` if the site shows another business.
- `nicht prüfbar`: cell could not be checked (site down, no Impressum); it is skipped in scoring.

"""


def load(path):
    ws = openpyxl.load_workbook(path, read_only=True)["Hausverwaltungen"]
    rows = list(ws.iter_rows(values_only=True))
    return [dict(zip(rows[0], r)) for r in rows[1:] if r and r[0]]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--current", type=Path, default=common.DATA / "heidi_scored_v16.xlsx")
    ap.add_argument("--regex", type=Path, default=common.DATA / "heidi_scored_v14.xlsx")
    ap.add_argument("--agent-ref", type=Path, default=common.DATA / "heidi_scored_v15.xlsx",
                    help="only for the printed cross-check of group 2")
    ap.add_argument("--verified", default="doc/verification_sample.md")
    ap.add_argument("--n", type=int, default=25)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--prompt", default="v2")
    ap.add_argument("--out", type=Path, default=Path("doc/holdout_blind.md"))
    ap.add_argument("--keys", type=Path, default=common.DATA / "holdout_keys.json")
    a = ap.parse_args()

    verified = {common.norm_domain(s["Website"]) for s in parse_sample(a.verified)}
    cur = [r for r in load(a.current) if r.get("Website")]
    dom = lambda r: common.norm_domain(r["Website"])  # noqa: E731
    prio_regex = {dom(r): r.get("Prio") for r in load(a.regex) if r.get("Website")}
    below_a = ("B", "C", "nicht bewertbar")

    g1 = [r for r in cur if common.AGENT_EINHEITEN_MARK in str(r.get("Hinweis") or "")]
    g2 = [r for r in cur if r.get("Prio") == "A" and prio_regex.get(dom(r)) in below_a]
    ref = {dom(r) for r in load(a.agent_ref) if r.get("Website") and r.get("Prio") == "A"
           and prio_regex.get(dom(r)) in below_a}
    print(f"Gruppe 1 (Agent-Einheiten ≥ 5.000): {len(g1)}")
    print(f"Gruppe 2 (A nur durch Agent-Werte, {a.current.name} vs {a.regex.name}): {len(g2)}; "
          f"gleiche Rechnung mit {a.agent_ref.name}: {len(ref)}"
          + ("" if ref == {dom(r) for r in g2} else
             f" (abweichend: {sorted(ref ^ {dom(r) for r in g2})})"))
    overlap = [dom(r) for r in g1 + g2 if dom(r) in verified]
    if overlap:
        print(f"bereits geprüft, ausgelassen: {overlap}")

    chosen, seen = [], set()
    for grp, rows in (("agent_enterprise", g1), ("a_via_agent", g2)):
        for r in sorted(rows, key=dom):
            if dom(r) not in verified and dom(r) not in seen:
                chosen.append((grp, r))
                seen.add(dom(r))
    pool = sorted((r for r in cur if r.get("Recherchetiefe") == "B+ (Agent)"
                   and dom(r) not in verified and dom(r) not in seen), key=dom)
    rng = random.Random(a.seed)
    need = max(0, a.n - len(chosen))
    if need > len(pool):
        sys.exit(f"Pool zu klein: {len(pool)} < {need}")
    chosen += [("random_b_plus_agent", r) for r in rng.sample(pool, need)]
    rng.shuffle(chosen)  # blind order: the group is not visible in the sheet

    import llm_enrich
    llm_enrich.OFFLINE_LLM = True
    keys, lines = [], ["| " + " | ".join(BLIND_COLS) + " |", "|" + "---|" * len(BLIND_COLS)]
    missing = []
    for i, (grp, r) in enumerate(chosen, 1):
        nr = f"H{i:02d}"
        res = llm_enrich.run_row({"r": 0, "Website": dom(r), "Firma": "", "tier": "B"}, a.prompt)
        if res["err"]:
            missing.append(f"{nr} {dom(r)}: {res['err']}")
        keys.append({"Nr": nr, "Firma": r["Firma"], "Website": dom(r), "Gruppe": grp, "Prio": r.get("Prio"),
                     "Prio regex-only": prio_regex.get(dom(r)), "Recherchetiefe": r.get("Recherchetiefe")})
        lines.append(f"| {nr} | {str(r['Firma']).replace('|', '/')} | {dom(r)} |" + " |" * (len(BLIND_COLS) - 3))
    a.out.write_text(HEADER + "\n".join(lines) + "\n", encoding="utf-8")
    a.keys.write_text(json.dumps({"seed": a.seed, "n": a.n, "prompt": a.prompt, "current": a.current.name,
                                  "regex_only": a.regex.name, "excluded_verified": len(verified),
                                  "pool_b_plus_agent": len(pool), "rows": keys},
                                 ensure_ascii=False, indent=1), encoding="utf-8")
    cnt = {}
    for k in keys:
        cnt[k["Gruppe"]] = cnt.get(k["Gruppe"], 0) + 1
    print(f"Holdout {len(keys)} Zeilen: {cnt}; Zufallspool B+ (Agent) {len(pool)}, Seed {a.seed}")
    print("ohne gecachte Agent-Antwort: " + ("; ".join(missing) if missing else "keine"))
    print(f"Geschrieben: {a.out}, {a.keys}")


if __name__ == "__main__":
    main()
