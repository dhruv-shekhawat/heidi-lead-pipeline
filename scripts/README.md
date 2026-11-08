# Scripts

**Pipeline.** Directory listings (11880, Gelbe Seiten; JACASA from supplied CSV exports via `jacasa_union.py`) are collected into CSVs, merged into the researched workbook (`heidi_enriched_v2.xlsx`) as new rows (`merge.py`), enriched from directory detail pages (`add_sources.py`), checked automatically on their own websites (`light_enrich.py`), mined for a contact person and personal e-mail (`extract_contacts.py`), Einheiten candidates are approved by rule (`apply_review.py --auto`), case-by-case decisions from `manual_fixes.json` are applied (`manual_fixes.py`), everything is scored into a Prio (`score.py`) and exported as the deliverable `heidi_final.xlsx` (`export_sheet.py`); `diff.py` shows what changed between two workbooks and `run_all.py` runs everything in order. Every HTTP request goes through `common.fetch`: normal browser User-Agent, robots.txt checked per host, every page cached in `data/cache/<sha1 of URL>.json` (POST: URL + sorted body) and reused on reruns, 403/429 → stop (a whole step for single-site scrapers, only that row/host for the per-company website steps). Values are never invented (blank if not found), only empty cells are filled, `data/heidi_merged_v3.xlsx` is never written, every step writes a new file and tabs a step does not mention pass through unchanged. Name matching uses `common.norm_name` (lowercase, umlauts folded, Rechtsform tokens and branch suffixes such as "NL Berlin" removed, then spaces/hyphens/punctuation removed, so "HVIM" = "hv-im"); PLZ comes from `Bezirk/PLZ` via `common.plz_of`.

Helper modules (not run directly): `common.py` (cache, robots, fetch, normalisation), `websites.py` (homepage URL variants, Impressum/Kontakt link discovery).

## scrape_directories.py
- **Purpose:** "Hausverwaltung" listings in Berlin from 11880 (`?page=N`) and Gelbe Seiten (page 1 + `POST /ajaxsuche`); keeps PLZ 10000–14199 and entries with a Hausverwaltung category (`Makler auch` = ja if Makler is also listed); dedupes on domain OR name+PLZ.
- **Input:** live sites (cached); `--skip-11880` reuses 11880 rows from the existing raw CSV; `--from-raw` makes no requests.
- **Output:** `data/directory_raw.csv`, `data/directory_clean.csv`.
- **Run:** `python3 scripts/scrape_directories.py --skip-11880`

## scrape_jacasa.py
- **Purpose:** JACASA Berlin profiles (load-more endpoint, sitemap fallback only if no endpoint is found), PLZ 10000–14199, fields Firma … Profil-URL. **jacasa.de currently returns a Cloudflare 403 → the script stops with exit code 2 and writes nothing; its extraction is untested on real profiles.**
- **Input:** jacasa.de (cached, 2 s delay).
- **Output:** `data/jacasa.csv`.
- **Run:** `python3 scripts/scrape_jacasa.py`

## jacasa_union.py
- **Purpose:** finds every JACASA CSV (`data/*jacasa*.csv`, `~/Downloads/jacasa*.csv`), detects layout 1 (Firma, Straße, PLZ, Bewertung, Anzahl Bewertungen, Mitgliedschaft, Mitarbeiter, Jahre aktiv, Tätig in) or layout 2 (… Anzahl_Bewertungen, Strasse, Ort, Taetig_in, Spezialisiert_auf, URL, Gefunden_in_Sortierung), unifies headers, maps the VDIV/IVD membership phrases to `VDIV`/`IVD` and Mietobjekte/Gewerbeobjekte to `Miet`/`Gewerbe`, converts comma decimals, unions on normalised Firma + PLZ (later files only fill empty fields).
- **Input:** the CSV exports (no network).
- **Output:** `data/jacasa_all.csv` (Profil-URL only where layout 2 had one; no Website/Telefon columns exist).
- **Run:** `python3 scripts/jacasa_union.py`

## fetch_details.py
- **Purpose:** Gelbe Seiten detail page for every Gelbe Seiten row (11880 rows skipped): Website, Telefon, E-Mail, full category list.
- **Input:** `data/directory_clean.csv` (Quell-URL), gelbeseiten.de (cached, 2 s delay).
- **Output:** `data/directory_detail.csv`.
- **Run:** `python3 scripts/fetch_details.py`

## merge.py
- **Purpose:** merges directory rows into the Hausverwaltungen tab (dup if domain OR name+PLZ matches any tab; name alone for rows without PLZ unless generic); moves Log "ungeklärt" rows without website to Hausverwaltungen as B (directory match) or C (name only); adds Recherchetiefe, Straße, Telefon, Makler auch.
- **Input:** `--in/--base data/heidi_enriched_v2.xlsx`, `data/directory_clean.csv`, `data/heidi_enriched_v1.xlsx` (PLZ for the moved Log rows).
- **Output:** `data/heidi_merged_run.xlsx` by default (refuses to write v3 or an input).
- **Run:** `python3 scripts/merge.py --base data/heidi_enriched_v2.xlsx --out data/run/merged.xlsx`

## add_sources.py
- **Purpose:** matches Gelbe Seiten detail rows and JACASA rows (domain, else name+PLZ) to Hausverwaltungen rows; fills empty Website/Telefon/Straße, JACASA columns (Mitgliedschaft, Mitarbeiter, Aktiv seit, Google-Bewertung, Anzahl Bewertungen, JACASA-Spezialisierung, JACASA-URL), Verwaltungsart/VDIV-Logo from JACASA; unmatched JACASA rows become B rows unless they match Log/Wohnungsunternehmen; adds Anzahl Quellen (distinct entries in Quelle); reports hidden duplicates (directory row whose website belongs to another row).
- **JACASA matching (no Website column):** normalised name + PLZ; else every 4+ letter word of the shorter name occurs in the longer one with the same PLZ (shorter name needs a distinctive word, i.e. not generic and in ≤ 3 firm names); else name alone against existing rows without PLZ. If the normalised name is identical to an existing firm under another PLZ, the JACASA row is merged into it (Hinweis "JACASA-Eintrag mit PLZ X (Bestand PLZ Y) zusammengeführt"); looser similarity under another PLZ stays a separate row with Hinweis "mögliches Duplikat von …". "Tätig in" without a Berlin location → Hinweis "Tätig in: …". Missing CSV columns are tolerated. Anzahl Quellen is recomputed: distinct entries in Quelle, +1 for JACASA when the row carries JACASA data (Quelle itself is not modified). Log rows "Einheiten verworfen" don't block new rows.
- **Input:** `--in` workbook (default `data/heidi_merged_v3.xlsx`), `data/directory_detail.csv`, `--jacasa data/jacasa_all.csv` (optional).
- **Output:** `--out` (default `data/heidi_merged_v3c.xlsx`).
- **Run:** `python3 scripts/add_sources.py --in data/heidi_scored_v7.xlsx --jacasa data/jacasa_all.csv --out data/heidi_v8_sources.xlsx`

## light_enrich.py
- **Purpose:** for Recherchetiefe B/C rows with a Website: homepage + Impressum (1 s/host, 10 s timeout); Website-Status; fills empty (or "unbekannt") Verwaltungsart, Software, Kundenportal, Messdienst, Fernablesung erwähnt, VDIV-Logo; Impressum Rechtsform name (page `<head>`, register, insurance and web-agency lines ignored) replaces Firma only when it shares a non-generic 4+ letter word and Name alt is empty (old name → Name alt), else noted in Hinweis; Kundenportal counts only links on the company's own site or to portal providers (casavi, etg24, idwell, hausperfekt, …); Heidi as Messdienst only in metering context; Einheiten candidates skip house numbers, money/area values and plain years ("2.000" is kept); Einheiten candidates go to a review CSV only; ok rows become "B+ (automatisch)".
- **Input:** `data/heidi_merged_v3c.xlsx`.
- **Output:** `data/heidi_light_v3d.xlsx`, `data/review_einheiten.csv` (fill column Freigabe with ja/nein).
- **Run:** `python3 scripts/light_enrich.py` (`HEIDI_OFFLINE=1` = cache only)

## extract_contacts.py
- **Purpose:** for A/B+ rows with a Website: first person named after Geschäftsführer/in, Geschäftsführung, Inhaber/in, vertreten durch, Vertretungsberechtigt in the Impressum → Ansprechpartner/Funktion/Kontakt-Quelle; first own-domain e-mail on Impressum or Kontakt page that is not info/kontakt/… **and** not a role or company-name mailbox **and** looks personal (`first.last`, `f.last`, or contains the contact's name) → E-Mail ("Website"); LinkedIn created empty.
- **Input:** `data/heidi_light_v3d.xlsx` (pages from cache, fetched if missing).
- **Output:** `data/heidi_contacts_v5.xlsx`.
- **Run:** `python3 scripts/extract_contacts.py`

## apply_review.py
- **Purpose:** applies `review_einheiten.csv`. `--auto` sets Freigabe itself: **ja** only if the row's own number is followed within 3 words by Einheiten/Wohneinheiten/Verwaltungseinheiten/Wohnungen and the same sentence contains verwalten/verwaltet(e)/betreuen/betreut/Verwaltung/Bestand (whole words); otherwise **nein (unsicher)**. ja → Einheiten (only if empty; several ja numbers per firm → the largest whose keyword is not followed by verkauft/vermittelt/gebaut/…), Einheiten-Quelle = URL + sentence, Recherchetiefe "A (Zahl geprüft)"; nein → Log row "Einheiten verworfen" with the phrase. Without `--auto` the hand-filled Freigabe is used.
- **Input:** a scored workbook + `data/review_einheiten.csv`.
- **Output:** `--out` workbook (then run score.py); with `--auto` the Freigabe column of the CSV is written.
- **Run:** `python3 scripts/apply_review.py --auto --in data/heidi_scored_v6.xlsx --out data/heidi_v7_applied.xlsx && python3 scripts/score.py --in data/heidi_v7_applied.xlsx --out data/heidi_scored_v7.xlsx`

## manual_fixes.py
- **Purpose:** applies case-by-case decisions from `scripts/manual_fixes.json` (rows found by Website domain): `einzelobjekt` moves Einheiten + Einheiten-Quelle to the Log ("verworfen: Einzelobjekt") and resets Recherchetiefe to "B+ (automatisch)" (currently Vogel und Partner 39, conwimo 165); `zu_wohnungsunternehmen` moves a row to the Wohnungsunternehmen tab with Segment, keeps Einheiten + phrase, puts Ansprechpartner/Telefon/Straße into Hinweis (currently UG Philipp Immobilien, Bestandshalter, 1.200). Idempotent.
- **Input:** a workbook after apply_review, `scripts/manual_fixes.json`.
- **Output:** `--out` workbook (then score.py).
- **Run:** `python3 scripts/manual_fixes.py --in data/heidi_v9_contacts.xlsx --out data/heidi_v9_fixed.xlsx`

## score.py
- **Purpose:** Punkte, Prio (Enterprise/A/B/C/nicht bewertbar) and a German Prio-Grund per Hausverwaltungen row; weights and thresholds read from the editable tab "Gewichte" (created with defaults if missing).
- **Input:** `--in data/heidi_contacts_v5.xlsx`.
- **Output:** `--out data/heidi_scored_v6.xlsx`.
- **Rule change 2026-11-08 (v9):** "bundesweit" in Hinweis makes a row Enterprise only with Einheiten ≥ 5.000 (Gewichte key `enterprise_bundesweit`); this only affected JURAG (now scored by points → C). Ralph, talyo, allod are Enterprise via their Einheiten (10.000 / 8.000 / 20.000). New default rows are appended to an existing Gewichte tab.
- **Threshold change 2026-11-08:** Prio A threshold lowered **6 → 5** (at 6 only 10 rows were A, below the target of 25; at 5: 17). v8 uses 5, the Gewichte tab records the change, and 5 is now the default for new Gewichte tabs. `--set key=value --note …` changes a Schwelle and writes a dated note into Gewichte.
- **Run:** `python3 scripts/score.py --in data/heidi_v8_contacts.xlsx --out data/heidi_scored_v8.xlsx [--set prio_a=5 --note "…"]`

## export_sheet.py
- **Purpose:** builds the deliverable: tabs "Leads Berlin" (Hausverwaltungen, 32 fixed columns, PLZ + Bezirk split from Bezirk/PLZ – Bezirk only where the source names one –, sorted Enterprise, A, B, C, nicht bewertbar, then Punkte desc), "Wohnungsunternehmen", "Gewichte", "Log", "Methodik" (sources with row counts, dedupe rules, Recherchetiefe definitions, scoring rules in plain German from Gewichte, field completeness, known gaps, data origin + UWG §7 note). Bold frozen header, autofilter, autofit, dropdowns on Prio and Verwaltungsart, conditional fill on Prio.
- **Input:** a scored workbook.
- **Output:** `data/heidi_final.xlsx`; prints the completeness table.
- **Run:** `python3 scripts/export_sheet.py --in data/heidi_scored_v9.xlsx --out data/heidi_final.xlsx`

## merge_contacts.py
- **Purpose:** final contact merge: copies E-Mail, E-Mail-Status, LinkedIn, Ansprechpartner from a hand-updated workbook into the scored workbook wherever the source cell is non-empty (overwrites on purpose; empty cells never clear). Rows matched by normalised Firma + PLZ, then Name alt + PLZ, then Website domain; unmatched rows are reported, not added. Scoring is not rerun (contact fields carry no points).
- **Input:** `--base data/heidi_scored_v9.xlsx`, `--contacts data/heidi_contacts_v10.xlsx` (same layout as v9).
- **Output:** `data/heidi_scored_v11.xlsx`; then `export_sheet.py` → `data/heidi_final.xlsx`.
- **Run:** `python3 scripts/merge_contacts.py && python3 scripts/export_sheet.py --in data/heidi_scored_v11.xlsx --out data/heidi_final.xlsx` · in run_all: `--contacts data/heidi_contacts_v10.xlsx` (optional step, skipped without the flag)

## diff.py
- **Purpose:** compares two workbooks per tab (row identity = normalised Firma + PLZ, fallback Name alt): new, removed, renamed rows, Prio changes, Einheiten / Ansprechpartner changes.
- **Input:** two workbook paths.
- **Output:** `data/diff_report.md` (or `--out`).
- **Run:** `python3 scripts/diff.py data/heidi_merged_v3.xlsx data/heidi_scored_v6.xlsx`

## run_all.py
- **Purpose:** runs scrape_directories → scrape_jacasa → jacasa_union → fetch_details → merge → add_sources → light_enrich → extract_contacts → apply_review (--auto) → manual_fixes → score → [merge_contacts, only with `--contacts`] → diff → export, each skippable (`--skip STEP …`), continues past failures, reuses the cache, copies the final scored workbook to `data/heidi_latest.xlsx`; `--dry-run` skips the three scrapers and runs the website steps from cache only.
- **Input:** `--base data/heidi_enriched_v2.xlsx` + the CSVs in `data/`.
- **Output:** `data/run/*` (merged, merged_sources, light, contacts, reviewed, fixed, scored, review_einheiten.csv, diff_report.md, heidi_final.xlsx), `data/heidi_latest.xlsx`. A dry-run reproduces `data/heidi_scored_v9.xlsx` exactly (checked with diff.py).
- **Run:** `python3 scripts/run_all.py --dry-run`
