# Heidi Systems case: Berlin Hausverwaltungen lead pipeline

Collects Berlin property-management firms (Hausverwaltungen), enriches them from their own websites and scores each lead into a priority (Enterprise / A / B / C). The deliverable is `data/heidi_final.xlsx`. Its first tab, "Top Leads", holds the Enterprise, A and B rows; "Alle Verwalter" holds all of them.

## Architecture

1. **Collect:** directory listings (Gelbe Seiten, 11880 page 1, JACASA CSV exports) are merged into a researched base list (Google Places, Apollo) and deduplicated by domain or name + PLZ.
2. **Fetch:** every HTTP request goes through one cached fetcher. It respects robots.txt, stops on 403/429 and never bypasses a block. Reruns read the cache.
3. **Rules:** keyword and regex parsers read each homepage and Impressum for Firma, Verwaltungsart, Einheiten, Software, Kundenportal, contact person and exclusion signals.
4. **Agent (optional):** `claude-haiku-5-5` extracts the same fields as JSON, with an exact quote and URL per value. A second call verifies each value; rejected values are logged.
5. **Hand checks:** `scripts/manual_fixes.json` applies verified corrections last, so they survive every rerun. Hand-checked A rows are never overwritten by the agent.
6. **Score:** points from editable weights (tab "Gewichte") give the Prio. An Einheiten figure of 5,000 or more that only the agent found does not count until it is checked.
7. **Contacts:** optional hand-researched contacts and Hunter.io e-mail patterns, both cached.
8. **Export:** formatted workbook with a Methodik tab, a Log of everything excluded or rejected, and a diff against the previous run.

Accuracy of the rules vs the agent, per field, on hand-checked rows, plus the pre-registered holdout: [doc/eval.md](doc/eval.md).
Scripts, inputs, outputs and rules: [scripts/README.md](scripts/README.md).

## How to run

```
python3 scripts/run_all.py --dry-run --contacts data/heidi_contacts_v10.xlsx --hunter --agent
```

- **`--dry-run`:** reads the page, Hunter and agent caches only; no network access.
- **Without `--dry-run`:** fetches what isn't cached, with Hunter and Anthropic keys from `.env` (git-ignored).
- **`--agent` off:** the pipeline runs on the rules alone.
- **`--skip STEP …`:** skips steps.

The run's workbooks go to `data/run/`, the scored workbook to `data/heidi_latest.xlsx` and the deliverable to `data/heidi_final.xlsx`.

## Cost per run

| Run | Cost |
|---|---|
| Dry run (cache only) | $0, about 1 minute |
| Agent, first run | ≈ $0.43 for 143 websites (283 calls). New rows cost about $0.003 each; cached rows are free. |
| Hunter.io | Free plan, at most 25 domain searches and 50 verifications a month. Only uncached domains use credits. |
| Directories and websites | Free |

`data/` is excluded from the repository (`.gitignore`) because it contains business contact data.
