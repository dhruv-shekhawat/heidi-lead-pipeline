#!/usr/bin/env python3
"""Run the whole pipeline in order; every step can be skipped; the cache is reused.

  scrape_directories → scrape_jacasa → jacasa_union → fetch_details → merge → add_sources →
  light_enrich → extract_contacts → apply_review (--auto) → manual_fixes → score →
  [merge_contacts, optional: --contacts data/heidi_contacts_v10.xlsx] →
  [hunter_enrich, optional: --hunter; uses Hunter.io credits, never in --dry-run] → diff → export

Workbooks and the review CSV of a run go to --outdir (default data/run/), so the step
deliverables (v3, v3c, v3d, v5, v6) and the hand-edited data/review_einheiten.csv are never
overwritten. The final scored workbook is copied to data/heidi_latest.xlsx.
A failing step is reported and the chain continues with the last good workbook.

Usage:
  python3 scripts/run_all.py [--base data/heidi_enriched_v2.xlsx] [--skip STEP …] [--dry-run]
  --dry-run  = skip scrape_directories, scrape_jacasa, fetch_details and run light_enrich /
               extract_contacts from the cache only (HEIDI_OFFLINE=1, no network access)
11880 is never fetched (scrape_directories always runs with --skip-11880).
"""
import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
S = ROOT / "scripts"
DATA = ROOT / "data"
STEPS = ["scrape_directories", "scrape_jacasa", "jacasa_union", "fetch_details", "merge", "add_sources",
         "light_enrich", "extract_contacts", "apply_review", "manual_fixes", "score", "merge_contacts",
         "hunter_enrich", "diff", "export"]
FETCH_STEPS = ["scrape_directories", "scrape_jacasa", "fetch_details"]
WEB_STEPS = ["light_enrich", "extract_contacts"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", type=Path, default=DATA / "heidi_enriched_v2.xlsx")
    ap.add_argument("--outdir", type=Path, default=DATA / "run")
    ap.add_argument("--skip", nargs="*", default=[], choices=STEPS)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--contacts", type=Path, default=None,
                    help="optional hand-updated contacts workbook (e.g. data/heidi_contacts_v10.xlsx); "
                         "without it merge_contacts is skipped")
    ap.add_argument("--hunter", action="store_true",
                    help="run hunter_enrich (Hunter.io, max 25 domain searches + 50 verifications; "
                         "key from .env; skipped in --dry-run)")
    ap.add_argument("--diff-old", type=Path, default=None,
                    help="default: previous data/heidi_latest.xlsx, else data/heidi_merged_v3.xlsx")
    a = ap.parse_args()
    out = a.outdir
    out.mkdir(parents=True, exist_ok=True)
    skip = set(a.skip) | (set(FETCH_STEPS) if a.dry_run else set()) | (set() if a.contacts else {"merge_contacts"}) \
        | (set() if a.hunter and not a.dry_run else {"hunter_enrich"})
    latest = DATA / "heidi_latest.xlsx"
    diff_old = a.diff_old or (latest if latest.exists() else DATA / "heidi_merged_v3.xlsx")
    if diff_old == latest and latest.exists():
        prev = out / "previous_latest.xlsx"
        shutil.copy2(latest, prev)
        diff_old = prev

    paths = {"merge": out / "merged.xlsx", "add_sources": out / "merged_sources.xlsx",
             "light_enrich": out / "light.xlsx", "extract_contacts": out / "contacts.xlsx",
             "apply_review": out / "reviewed.xlsx", "manual_fixes": out / "fixed.xlsx",
             "merge_contacts": out / "scored_contacts.xlsx", "hunter_enrich": out / "hunter.xlsx",
             "score": out / "scored.xlsx"}
    current = None  # workbook flowing through the chain
    py = sys.executable
    status = []
    for step in STEPS:
        if step in skip:
            status.append((step, "übersprungen"))
            print(f"\n### {step}: übersprungen")
            continue
        env = dict(os.environ)
        if a.dry_run and step in WEB_STEPS:
            env["HEIDI_OFFLINE"] = "1"
        if step == "scrape_directories":
            cmd = [py, S / "scrape_directories.py", "--skip-11880"]
        elif step == "scrape_jacasa":
            cmd = [py, S / "scrape_jacasa.py"]
        elif step == "jacasa_union":
            cmd = [py, S / "jacasa_union.py"]
        elif step == "fetch_details":
            cmd = [py, S / "fetch_details.py"]
        elif step == "merge":
            cmd = [py, S / "merge.py", "--in", a.base, "--out", paths["merge"]]
        elif step == "add_sources":
            cmd = [py, S / "add_sources.py", "--in", current or paths["merge"], "--out", paths[step],
                   "--jacasa", DATA / "jacasa_all.csv"]
        elif step == "light_enrich":
            cmd = [py, S / "light_enrich.py", "--in", current, "--out", paths[step],
                   "--review", out / "review_einheiten.csv"]
        elif step == "extract_contacts":
            cmd = [py, S / "extract_contacts.py", "--in", current, "--out", paths[step]]
        elif step == "apply_review":
            cmd = [py, S / "apply_review.py", "--auto", "--in", current, "--out", paths[step],
                   "--review", out / "review_einheiten.csv"]
        elif step == "manual_fixes":
            cmd = [py, S / "manual_fixes.py", "--in", current, "--out", paths[step]]
        elif step == "score":
            cmd = [py, S / "score.py", "--in", current, "--out", paths[step]]
        elif step == "merge_contacts":
            cmd = [py, S / "merge_contacts.py", "--base", current, "--contacts", a.contacts, "--out", paths[step]]
        elif step == "hunter_enrich":
            cmd = [py, S / "hunter_enrich.py", "--in", current, "--out", paths[step]]
        elif step == "diff":
            cmd = [py, S / "diff.py", diff_old, current, "--out", out / "diff_report.md"]
        else:  # export
            cmd = [py, S / "export_sheet.py", "--in", current, "--out", out / "heidi_final.xlsx"]
        if None in cmd:
            status.append((step, "nicht ausgeführt (keine Eingabe)"))
            continue
        print(f"\n### {step}{' (offline, nur Cache)' if env.get('HEIDI_OFFLINE') else ''}", flush=True)
        t0 = time.time()
        rc = subprocess.run([str(c) for c in cmd], env=env, cwd=ROOT).returncode
        dt = time.time() - t0
        if rc == 0 and step in paths and paths[step].exists():
            current = paths[step]
        status.append((step, f"ok ({dt:.0f}s)" if rc == 0 else f"Fehler/Stopp rc={rc} ({dt:.0f}s) – weiter"))

    print("\n=== run_all Zusammenfassung ===")
    for s, st in status:
        print(f"  {s:20s} {st}")
    if current and current in (paths["score"], paths["merge_contacts"], paths["hunter_enrich"]):
        shutil.copy2(current, latest)
        print(f"Endergebnis kopiert: {current.relative_to(ROOT)} → {latest.relative_to(ROOT)}")
    else:
        print("score lief nicht erfolgreich – data/heidi_latest.xlsx nicht aktualisiert")


if __name__ == "__main__":
    main()
