# Scripts

**Pipeline.** Directory listings (11880, Gelbe Seiten; JACASA from supplied CSV exports via `jacasa_union.py`) are collected into CSVs, merged into the researched workbook (`heidi_enriched_v2.xlsx`) as new rows (`merge.py`), enriched from directory detail pages (`add_sources.py`), checked automatically on their own websites (`light_enrich.py`), mined for a contact person and personal e-mail (`extract_contacts.py`), Einheiten candidates are approved by rule (`apply_review.py --auto`), case-by-case decisions from `manual_fixes.json` are applied (`manual_fixes.py`), everything is scored into a Prio (`score.py`) and exported as the deliverable `heidi_final.xlsx` (`export_sheet.py`); `diff.py` shows what changed between two workbooks and `run_all.py` runs everything in order. Every HTTP request goes through `common.fetch`: normal browser User-Agent, robots.txt checked per host, every page cached in `data/cache/<sha1 of URL>.json` (POST: URL + sorted body) and reused on reruns (timeouts/HTTP 5xx are cached as transient: retried online, replayed offline, so `--dry-run` reproduces the last online run exactly), 403/429 → stop (a whole step for single-site scrapers, only that row/host for the per-company website steps). `fetch` returns page text with soft hyphens and zero-width characters (raw or as entities) removed, so "Miet[soft hyphen]verwaltung" reads as "Mietverwaltung" in every parser; the cache files keep the original bytes. Values are never invented (blank if not found), only empty cells are filled, `data/heidi_merged_v3.xlsx` is never written, every step writes a new file and tabs a step does not mention pass through unchanged. Name matching uses `common.norm_name` (lowercase, umlauts folded, Rechtsform tokens and branch suffixes such as "NL Berlin" removed, then spaces/hyphens/punctuation removed, so "ABC HV" = "abc-hv"); PLZ comes from `Bezirk/PLZ` via `common.plz_of`.

Helper modules (not run directly): `common.py` (cache, robots, fetch, text cleaning, `html_lines` = visible lines for the line parsers, `page_text` = full page text for keyword rules, normalisation), `websites.py` (homepage URL variants, Impressum/Kontakt link discovery).

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

## dedupe_rows.py
- **Purpose:** applies the same-name/same-PLZ merge (`add_sources.merge_same_name_rows`) to an existing workbook; re-run score.py afterwards.
- **Input / Output:** `--in` / `--out` workbook.
- **Run:** `python3 scripts/dedupe_rows.py --in data/heidi_scored_v9.xlsx --out data/tmp.xlsx`

## add_sources.py
- **Purpose:** matches Gelbe Seiten detail rows and JACASA rows (domain, else name+PLZ) to Hausverwaltungen rows; fills empty Website/Telefon/Straße, JACASA columns (Mitgliedschaft, Mitarbeiter, Aktiv seit, Google-Bewertung, Anzahl Bewertungen, JACASA-Spezialisierung, JACASA-URL), Verwaltungsart/VDIV-Logo from JACASA; unmatched JACASA rows become B rows unless they match Log/Wohnungsunternehmen; adds Anzahl Quellen (distinct entries in Quelle); reports hidden duplicates (directory row whose website belongs to another row).
- **JACASA matching (no Website column):** normalised name + PLZ; else every 4+ letter word of the shorter name occurs in the longer one with the same PLZ (shorter name needs a distinctive word, i.e. not generic and in ≤ 3 firm names); else name alone against existing rows without PLZ. If the normalised name is identical to an existing firm under another PLZ, the JACASA row is merged into it (Hinweis "JACASA-Eintrag mit PLZ X (Bestand PLZ Y) zusammengeführt"); looser similarity under another PLZ stays a separate row with Hinweis "mögliches Duplikat von …". "Tätig in" without a Berlin location → Hinweis "Tätig in: …". Missing CSV columns are tolerated. Anzahl Quellen is recomputed: distinct entries in Quelle, +1 for JACASA when the row carries JACASA data (Quelle itself is not modified). Log rows "Einheiten verworfen" don't block new rows.
- **Same-name duplicates:** at the start, rows with the same normalised Firma + PLZ and at most one distinct website are merged into the row with the website (empty cells filled, Quelle/Quellen unioned, other differing values – e.g. a second Bezirk/PLZ spelling or Recherchetiefe – noted in Hinweis, Anzahl Quellen recomputed). light_enrich runs the same merge after its renames, which is where such duplicates arise.
- **Input:** `--in` workbook (default `data/heidi_merged_v3.xlsx`), `data/directory_detail.csv`, `--jacasa data/jacasa_all.csv` (optional).
- **Output:** `--out` (default `data/heidi_merged_v3c.xlsx`).
- **Run:** `python3 scripts/add_sources.py --in data/heidi_scored_v7.xlsx --jacasa data/jacasa_all.csv --out data/heidi_v8_sources.xlsx`

## light_enrich.py
- **Purpose:** for Recherchetiefe B/C rows with a Website: homepage + Impressum (1 s/host, 10 s timeout); Website-Status; fills empty (or "unbekannt") Verwaltungsart, Software, Kundenportal, Messdienst, Fernablesung erwähnt, VDIV-Logo; Impressum Rechtsform name (page `<head>`, register, insurance and web-agency lines ignored) replaces Firma only when it shares a non-generic 4+ letter word and Name alt is empty (old name → Name alt), else noted in Hinweis; Kundenportal counts only links on the company's own site or to portal providers (casavi, etg24, idwell, hausperfekt, …); Heidi as Messdienst only in metering context; Einheiten candidates skip house numbers, money/area values and plain years ("2.000" is kept); Einheiten candidates go to a review CSV only; ok rows become "B+ (automatisch)". Keyword rules (Verwaltungsart, Kundenportal, Messdienst, Fernablesung, Einheiten) read the full page text: content hidden in `<template>`/`<noscript>` (slider slides, lazy sections), JSON-LD descriptions, alt/title/aria-label and plain-text `data-*` attributes (counter targets). Kundenportal is also "ja" for links to an App Store / Google Play listing, links to an app./portal./kunden. subdomain of the firm's own domain, or a named "Mieter-App". **Exclusion rule:** if the homepage `<title>` or the Impressum (the homepage text when no Impressum page exists) names Ferienwohnung(en), Ferienhaus, Ferienvermietung, Gästehaus, Pension or Hotel(s) – Hotel only in title/Impressum, not in homepage body text, where hotels appear as references – and neither names Hausverwaltung/Immobilienverwaltung, Typ becomes "ausschließen" and the row moves to the Log (Bereich "ausgeschlossen", reason + trigger word). "Apartments" alone only adds Hinweis "prüfen: Apartments" (could be serviced/business living). Only rows whose homepage loaded are judged; A rows are never processed by this step.
- **Input:** `data/heidi_merged_v3c.xlsx`.
- **Output:** `data/heidi_light_v3d.xlsx`, `data/review_einheiten.csv` (fill column Freigabe with ja/nein).
- **Run:** `python3 scripts/light_enrich.py` (`HEIDI_OFFLINE=1` = cache only)

## extract_contacts.py
- **Purpose:** for A/B+ rows with a Website: first person named after Geschäftsführer/in, Geschäftsführung, Inhaber/in, vertreten durch, Vertretungsberechtigt in the Impressum → Ansprechpartner/Funktion/Kontakt-Quelle; first own-domain e-mail on Impressum or Kontakt page that is not info/kontakt/… **and** not a role or company-name mailbox **and** looks personal (`first.last`, `f.last`, or contains the contact's name) → E-Mail ("Website"); LinkedIn created empty. Academic degrees before the name (Dipl.-Ing., Dipl.-Kfm., …) are skipped; "Geschäftsführende Gesellschafter" is a title of its own (Funktion "Geschäftsführender Gesellschafter").
- **Input:** `data/heidi_light_v3d.xlsx` (pages from cache, fetched if missing).
- **Output:** `data/heidi_contacts_v5.xlsx`.
- **Run:** `python3 scripts/extract_contacts.py`

## llm_enrich.py
- **Purpose:** extraction agent. For every Hausverwaltungen row whose homepage **and** Impressum are in the page cache (A rows included), `claude-haiku-5-5` returns Firma, Verwaltungsart, Einheiten, Software, Kundenportal, Messdienst, Fernablesung, VDIV, Ansprechpartner, Funktion and Typ_Verwalter (ja/nein + reason) as JSON (structured outputs); every value carries an exact `quote` and the page `url`, null when not stated. Code checks that each quote occurs in the page text; a second call (the verifier) judges value + quote against the page → "bestätigt"/"verworfen" + one-line reason; rejected values are dropped and logged (Log, Bereich "Agent verworfen"). Writing: A rows only get empty cells filled, differing values go to `Agent-Abweichung`; other rows take the confirmed agent value over the regex value (old value in `Agent-Abweichung`, Recherchetiefe "B+ (Agent)") – except Software, where the rules won on the holdout (`RULES_FIRST`, decided by the pre-registered rule in `doc/eval.md`): there an existing rule value is kept and the agent only fills empty cells; `Agent-Quote` holds the quotes; Typ_Verwalter "nein" → Hinweis "Agent: kein Verwalter, prüfen (…)", never auto-excluded. A written Einheiten figure ≥ 5.000 → Hinweis "Agent-Zahl, prüfen (Einheiten N)" (not scored until verified, see score.py). Runs before manual_fixes, so hand-checked corrections win.
- **Model settings:** adaptive thinking at effort `low`, JSON schema via `output_config.format`; no `temperature` (claude-haiku-5-5 rejects non-default sampling parameters) – reruns are reproducible through the response cache `data/cache/llm_<sha1>.json` (key = model + prompt text + input). Prompts are versioned in `scripts/prompts/` (`extract_vN.md`, `verify_vN.md`); v2 additionally gives the model each page's link list and makes the verifier write its reason before the verdict. Pages are capped at 60,000 characters each (marked "gekürzt"), keeping prompts in the ≤ 100K-token price band ($0.10 / $0.50 per MTok).
- **Input:** a workbook after apply_review; key `ANTHROPIC_API_KEY` from the environment or `.env` (git-ignored).
- **Output:** `--out` workbook, `--report` CSV with regex/table value vs agent value, verdict, reason, quote and URL per row and field (the regex output itself stays in the input workbook and in `data/heidi_scored_v14.xlsx`, built without the agent). Prints counts, the rows flagged "kein Verwalter" and the API cost.
- **Run:** `python3 scripts/llm_enrich.py --in data/run/reviewed.xlsx --out data/run/agent.xlsx --prompt v2 --report data/run/agent_extraction.csv` · in run_all: `--agent` (with `--dry-run`: cached responses only).

## eval_extraction.py
- **Purpose:** error rate per field on any hand-checked sample in the same Markdown format (`--sample`, default the 30-row sample (`doc/verification_sample.md`, built from `doc/verification_report.md`; both untracked because they hold firm and person names) for the regex rules, one or more workbooks (default v12 = hand research + regex; `--table` alone = none), each agent prompt version and the Ensemble (agent `--ensemble v2` for Firma, Verwaltungsart, Einheiten, regex rules for the rest; regex value kept where the agent has none). Field comparisons are documented in the script; "nicht prüfbar" cells are skipped.
- **Output:** aggregate table (stdout, `--md`); per-row errors in `data/eval_details.md` (untracked). Results and the v1 → v2 prompt changes: `doc/eval.md`.
- **Run:** `python3 scripts/eval_extraction.py --prompts v1 v2 --table v12=data/heidi_scored_v12.xlsx v15=data/heidi_scored_v15.xlsx` · holdout (once filled in): `python3 scripts/eval_extraction.py --sample doc/holdout_blind.md --prompts v2 --table`

## holdout_normalize.py
- **Purpose:** turns the checker's free-text holdout truth (`doc/holdout_truth.md`) into the eval format by fixed rules (remarks in brackets dropped, "beides + Gewerbe" → beides, "nicht angegeben" → keine Angabe / keine / unbekannt, Einheiten = number before the quote, Berlin-only from "nicht nur Berlin"/"unklar"); Website from `data/holdout_keys.json`.
- **Output:** `doc/holdout_sample.md` (untracked). **Run:** `python3 scripts/holdout_normalize.py`, then `python3 scripts/eval_extraction.py --sample doc/holdout_sample.md --prompts v2 --table`

## holdout_select.py
- **Purpose:** selects the 25-row holdout for the extraction eval, none among the 30 verified rows: every row with an unverified agent Einheiten ≥ 5.000, every row that is A only through agent values (A in the current workbook, below A in the regex-only v14, matched by domain), then random "B+ (Agent)" rows (pool sorted by domain, seed 42). Checks that every row has cached pages and agent responses. The decision rule for the result is pre-registered in `doc/eval.md`.
- **Output:** `doc/holdout_blind.md` (Firma + Website + empty field columns, shuffled order) and `data/holdout_keys.json` (group per row, seed, source workbooks); both untracked (firm names).
- **Run:** `python3 scripts/holdout_select.py --current data/heidi_scored_v16.xlsx --regex data/heidi_scored_v14.xlsx`

## apply_review.py
- **Purpose:** applies `review_einheiten.csv`. `--auto` sets Freigabe itself: **ja** only if the row's own number is followed within 3 words by Einheiten/Wohneinheiten/Verwaltungseinheiten/Wohnungen and the same sentence contains verwalten/verwaltet(e)/betreuen/betreut/Verwaltung/Bestand (whole words); otherwise **nein (unsicher)**. ja → Einheiten (only if empty; several ja numbers per firm → the largest whose keyword is not followed by verkauft/vermittelt/gebaut/…), Einheiten-Quelle = URL + sentence, Recherchetiefe "A (Zahl geprüft)"; nein → Log row "Einheiten verworfen" with the phrase. Without `--auto` the hand-filled Freigabe is used.
- **Input:** a scored workbook + `data/review_einheiten.csv`.
- **Output:** `--out` workbook (then run score.py); with `--auto` the Freigabe column of the CSV is written.
- **Run:** `python3 scripts/apply_review.py --auto --in data/heidi_scored_v6.xlsx --out data/heidi_v7_applied.xlsx && python3 scripts/score.py --in data/heidi_v7_applied.xlsx --out data/heidi_scored_v7.xlsx`

## manual_fixes.py
- **Purpose:** applies case-by-case decisions from `scripts/manual_fixes.json`. The tracked file holds only generic entries with an ID; which firm an ID means is stored in the untracked `data/manual_fixes_keys.json` (firm key = normalised name + PLZ, website as fallback, plus the detailed reason), so no firm names are versioned. `einzelobjekt` moves Einheiten + Einheiten-Quelle to the Log ("verworfen: Einzelobjekt") and resets Recherchetiefe to "B+ (automatisch)" (currently two single-property references); `zu_wohnungsunternehmen` moves a row to the Wohnungsunternehmen tab with Segment, keeps Einheiten + phrase and puts Ansprechpartner/Telefon/Straße into Hinweis (currently one owner-managed stock); `felder` writes values checked by hand (currently ten rows from the 2026-11-08 sample: Verwaltungsart, Kundenportal, Einheiten + Einheiten-Quelle, Ansprechpartner/Funktion, one Firma) – these overwrite on purpose, the old value is noted in Hinweis, Einheiten get Einheiten-Quelle "URL – „Satz“" and "A (Zahl geprüft)", a renamed Firma keeps the old name in Name alt; the values themselves live only in the keys file; `ausschliessen` moves a row to the Log as "ausgeschlossen" (currently two lodging businesses), using the same helper as light_enrich's exclusion rule. Without the keys file all entries are skipped. Idempotent.
- **Input:** a workbook after apply_review, `scripts/manual_fixes.json`, `data/manual_fixes_keys.json` (untracked).
- **Output:** `--out` workbook (then score.py).
- **Run:** `python3 scripts/manual_fixes.py --in data/heidi_v9_contacts.xlsx --out data/heidi_v9_fixed.xlsx`

## score.py
- **Purpose:** Punkte, Prio (Enterprise/A/B/C/nicht bewertbar) and a German Prio-Grund per Hausverwaltungen row; weights and thresholds read from the editable tab "Gewichte" (created with defaults if missing).
- **Input:** `--in data/heidi_contacts_v5.xlsx`.
- **Output:** `--out data/heidi_scored_v6.xlsx`.
- **Rule change 2026-11-08 (v9):** "bundesweit" in Hinweis makes a row Enterprise only with Einheiten ≥ 5.000 (Gewichte key `enterprise_bundesweit`); this affected one row without an Einheiten figure (now scored by points). Rows with ≥ 5.000 Einheiten stay Enterprise regardless. New default rows are appended to an existing Gewichte tab.
- **Threshold change 2026-11-08:** Prio A threshold lowered **6 → 5** (at 6 only 10 rows were A, below the target of 25; at 5: 17). v8 uses 5, the Gewichte tab records the change, and 5 is now the default for new Gewichte tabs. `--set key=value --note …` changes a Schwelle and writes a dated note into Gewichte.
- **Rule 2026-11-08 (v16):** an Einheiten figure ≥ 5.000 that only the extraction agent found (Hinweis "Agent-Zahl, prüfen", set by llm_enrich) is scored as empty – no Einheiten points, no Enterprise – until it is verified; Prio-Grund shows "Einheiten N Agent-Zahl, nicht gewertet (+0)" and the number stays in the cell. A `manual_fixes.json` entry that sets Einheiten removes the marker. In v16 this moved 4 rows from Enterprise to B or C.
- **Run:** `python3 scripts/score.py --in data/heidi_v8_contacts.xlsx --out data/heidi_scored_v8.xlsx [--set prio_a=5 --note "…"]`

## export_sheet.py
- **Purpose:** builds the deliverable: tabs "Top Leads" (Prio Enterprise, A and B only; 16 sales columns: Prio, Firma, Website, PLZ, Verwaltungsart, Einheiten, Software, Kundenportal, Mitgliedschaft, Ansprechpartner, Funktion, E-Mail, E-Mail-Status, LinkedIn, Telefon, Prio-Grund; same sort and formatting as the full tab), "Alle Verwalter" (Hausverwaltungen, 33 fixed columns incl. Allgemeine E-Mail next to E-Mail, plus Agent-Quote and Agent-Abweichung as the last two columns when the workbook comes from `run_all --agent`, PLZ + Bezirk split from Bezirk/PLZ – Bezirk only where the source names one –, sorted Enterprise, A, B, C, nicht bewertbar, then Punkte desc), "Wohnungsunternehmen", "Gewichte", "Log", "Methodik" (sources with row counts, dedupe rules, Recherchetiefe definitions, scoring rules in plain German from Gewichte, field completeness, known gaps, data origin + UWG §7 note). Bold frozen header, autofilter, autofit, dropdowns on Prio and Verwaltungsart, conditional fill on Prio (both lead tabs).
- **Input:** a scored workbook.
- **Output:** `data/heidi_final.xlsx`; prints the completeness table.
- **Run:** `python3 scripts/export_sheet.py --in data/heidi_scored_v9.xlsx --out data/heidi_final.xlsx`

## merge_contacts.py
- **Purpose:** final contact merge: copies E-Mail, E-Mail-Status, LinkedIn, Ansprechpartner and Allgemeine E-Mail (column created if missing) from a hand-updated workbook into the scored workbook wherever the source cell is non-empty (overwrites on purpose; empty cells never clear). Rows matched by full identity (Firma + PLZ + Website + Name alt), then Website domain, then Firma + PLZ / Name alt + PLZ only if that name key is unique (the list has a few same-name/same-PLZ duplicates); ambiguous or unmatched rows are reported, not guessed or added. If the copied Ansprechpartner is a different person (other surname, degrees ignored), Funktion is cleared because it described the replaced person. Scoring is not rerun (contact fields carry no points).
- **Input:** `--base data/heidi_scored_v9.xlsx`, `--contacts data/heidi_contacts_v10.xlsx` (same layout as v9).
- **Output:** `data/heidi_scored_v11.xlsx`; then `export_sheet.py` → `data/heidi_final.xlsx`.
- **Run:** `python3 scripts/merge_contacts.py && python3 scripts/export_sheet.py --in data/heidi_scored_v11.xlsx --out data/heidi_final.xlsx` · in run_all: `--contacts data/heidi_contacts_v10.xlsx` (optional step, skipped without the flag)

## hunter_enrich.py
- **Purpose:** Hunter.io for the top leads. Rows with Prio Enterprise / A / B and a Website, ordered by Prio then Punkte, at most 25 distinct domains (free plan: 50 credits/month = 25 Domain Searches at 1 credit + 50 verifications at 0.5; capped by the remaining quota from the free `/v2/account` call). Domain Search records `Hunter-Pattern` and `Hunter-Personen`; if a listed person matches the Ansprechpartner (umlaut variants and particles handled) → E-Mail with status "Hunter NN%", else the Ansprechpartner's address is built from the pattern → "Muster (Hunter-Pattern)". Only empty E-Mail cells are written (addresses found on the website and anything "verifiziert" are kept). Email Verifier on every written address (≤ 50): "verifiziert" / "ungültig (Hunter)" (the invalid address is removed from E-Mail, the status stays so the check is visible; such rows are not retried) / "unsicher (accept-all)" / "unsicher (<status>)". Only uncached calls count against the credit budget; domains already in the Hunter cache are reused even when a changed ranking pushes them below the 25 cut (no credits), so earlier results never drop out; with `HEIDI_OFFLINE=1` (run_all `--dry-run`) only cached responses are used. Responses cached in `data/cache/hunter_*.json` without the key, so reruns cost no credits; 401/403/429 or a quota error stops the API calls. Prints a table per Prio and the remaining quota (response headers if Hunter sends any, plus `/v2/account`).
- **Key:** `HUNTER_API_KEY` in `.env` (project root, git-ignored) or the environment.
- **Input:** `data/heidi_contacts_v10.xlsx`.
- **Output:** `data/heidi_contacts_v10b.xlsx`.
- **Run:** `python3 scripts/hunter_enrich.py` · `--only DOMAIN …` looks up exactly these domains, even below the 25 cut (once cached they stay selected in later runs) · `--plan` lists the 25 domains without spending credits · in run_all: `--hunter` (with `--dry-run`: cache only, no credits).
- **Note:** pattern-built addresses are guesses until verified; for e-mail outreach observe UWG §7 and GDPR.

## diff.py
- **Purpose:** compares two workbooks per tab (row identity = normalised Firma + PLZ, fallback Name alt): new, removed, renamed rows, Prio changes and changes in Einheiten, Einheiten Berlin, Verwaltungsart, Kundenportal, Ansprechpartner, Funktion, E-Mail, E-Mail-Status, Allgemeine E-Mail.
- **Input:** two workbook paths.
- **Output:** `data/diff_report.md` (or `--out`).
- **Run:** `python3 scripts/diff.py data/heidi_merged_v3.xlsx data/heidi_scored_v6.xlsx`

## run_all.py
- **Purpose:** runs scrape_directories → scrape_jacasa → jacasa_union → fetch_details → merge → add_sources → light_enrich → extract_contacts → apply_review (--auto) → [llm_enrich, only with `--agent`; cache only in `--dry-run`] → manual_fixes → score → [merge_contacts, only with `--contacts`] → [hunter_enrich, only with `--hunter`; cache only in `--dry-run`] → diff → export, each skippable (`--skip STEP …`), continues past failures, reuses the cache, copies the final scored workbook to `data/heidi_latest.xlsx`; `--dry-run` skips the three scrapers and runs the website steps from cache only.
- **Input:** `--base data/heidi_enriched_v2.xlsx` + the CSVs in `data/`.
- **Output:** `data/run/*` (merged, merged_sources, light, contacts, reviewed, fixed, scored, review_einheiten.csv, diff_report.md, heidi_final.xlsx), `data/heidi_latest.xlsx`. A dry-run reproduces `data/heidi_scored_v9.xlsx` exactly (checked with diff.py).
- **Run:** `python3 scripts/run_all.py --dry-run` · full reproduction of v18 (`data/heidi_scored_v18.xlsx` and `data/heidi_final.xlsx`; v17 before the Hunter lookups of 2026-11-09): `python3 scripts/run_all.py --dry-run --contacts data/heidi_contacts_v10.xlsx --hunter --agent` (checked 2026-11-09: 0 differing cells in the scored workbook's four tabs and the export's six tabs; without `--agent` it reproduces the regex-only v14)
