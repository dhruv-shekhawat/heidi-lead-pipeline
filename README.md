# Heidi Systems case: Berlin Hausverwaltungen lead pipeline

Collects Berlin property-management firms (Hausverwaltungen) from public directories, merges them with a researched base list, enriches them from their websites' Impressum pages, and scores each lead into a priority (Enterprise / A / B / C).
Sources: 11880, Gelbe Seiten, JACASA (CSV exports), Google Places and Apollo (base list), plus each firm's own website and Impressum; robots.txt is respected and blocked sites (403/429) are not bypassed.
Run: `python3 scripts/run_all.py` (`--dry-run` reuses the page cache without network access; `--skip STEP …` skips steps).
Output lands in `data/`: the run's workbooks in `data/run/`, the latest scored workbook in `data/heidi_latest.xlsx`, the formatted deliverable in `data/heidi_final.xlsx`.
Pipeline, scripts, inputs, outputs and rules: [scripts/README.md](scripts/README.md).
`data/` is excluded from the repository (`.gitignore`) because it contains business contact data.
