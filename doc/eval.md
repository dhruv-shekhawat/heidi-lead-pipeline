# Extraction eval: regex vs agent (2026-11-08)

**Sample:** the 30 rows hand-checked on 2026-11-08 (`doc/verification_sample.md`, built from `doc/verification_report.md`; both untracked because they contain firm and person names). Seven fields were compared: Firma, Verwaltungsart, Einheiten, Ansprechpartner, Software, Kundenportal, Typ_Verwalter. That gives 162 cells; cells marked "nicht prüfbar" are left out. Run it with `python3 scripts/eval_extraction.py --prompts v1 v2 --table v12=… v15=…`. Per-row errors are written to `data/eval_details.md`, which is also untracked.

**Systems:**

- **Regex:** the light_enrich, extract_contacts and apply_review rules, run on the same cached pages.
- **Tabelle v12:** the workbook the sample was checked against, i.e. hand research for the A rows plus regex for the rest. Its 16 errors are the report's 14 bold cells plus Typ_Verwalter for the two lodging rows, which the report records only as an instruction. This matches the hand check exactly and confirms the comparison rules.
- **Agent v1 / v2:** `claude-haiku-5-5`, extractor plus verifier. Only verified values count.
- **Tabelle v15:** the final workbook. It contains the agent output, and the manual corrections from `manual_fixes` are applied on top. So its score is not an independent measurement.

**How errors are counted:** a field with no value counts as an error wherever the truth has a value. Three sample rows have no cached Impressum page (HTTP 404 or no Impressum at all). The agent gets no input for those rows, so all of their checked fields count against it.

## Error rate per field (errors / checked cells)

| Feld | geprüft | Regex | Tabelle v12 | Tabelle v15 | Agent v1 | Agent v2 |
|---|---|---|---|---|---|---|
| Firma | 30 | 10 (33%) | 3 (10%) | 2 (7%) | 8 (27%) | 5 (17%) |
| Verwaltungsart | 29 | 7 (24%) | 4 (14%) | 0 (0%) | 17 (59%) | 5 (17%) |
| Einheiten | 16 | 9 (56%) | 2 (12%) | 0 (0%) | 10 (62%) | 4 (25%) |
| Ansprechpartner | 26 | 3 (12%) | 2 (8%) | 0 (0%) | 6 (23%) | 4 (15%) |
| Software | 13 | 3 (23%) | 0 (0%) | 0 (0%) | 8 (62%) | 5 (38%) |
| Kundenportal | 18 | 1 (6%) | 3 (17%) | 0 (0%) | 5 (28%) | 2 (11%) |
| Typ_Verwalter | 30 | 0 (0%) | 2 (7%) | 0 (0%) | 1 (3%) | 1 (3%) |
| **gesamt** | 162 | **33 (20.4%)** | **16 (9.9%)** | **2 (1.2%)** | **55 (34.0%)** | **26 (16.0%)** |

**Reading the table:**

- **Agent v2 vs regex.** On the same pages, agent v2 makes fewer errors than the regex rules (16.0% vs 20.4%). It is better on Firma, Verwaltungsart and Einheiten. It is worse on Kundenportal, Software, Ansprechpartner and Typ_Verwalter. Regex reads every link and logo, and it does better there.
- **v12 table vs agent.** The v12 table, i.e. hand research plus regex, is still better than the agent alone (9.9%).
- **v15 is not an independent result.** The final table, with agent and manual fixes, shows 1.2%. That number is mostly the manual corrections from this same sample, so don't read it as independent accuracy.

## What changed from prompt v1 to v2 (one round, based on the v1 errors)

Most v1 errors were empty values: the agent either didn't extract a value or the verifier rejected it. v2 changes six things:

1. **Two-sided Verwaltungsart.** A single quote rarely names both WEG and Miet, so the v1 verifier rejected "beides" in 12 of 23 cases.
   - v2 lets a value cite two passages joined with `||`.
   - Both prompts now list the signal words for each side, e.g. Eigentümergemeinschaften = WEG; Miethäuser and Sondereigentumsverwaltung = Miet.
   - Menu items and service headings count as offered services.
2. **Figures.**
   - "rund", "ca.", "über" and "X+" still mean X.
   - Figures blocks ("in Zahlen") with the label Einheiten or Wohnungen count as managed units.
   - The quote joins the number with its label, e.g. "8000 Wohnungen". v1 quoted bare numbers, and the verifier rejected them.
3. **Links.** Software and portals often appear only as links, e.g. a `mycasavi.com` login.
   - v2 passes each page's link list (link text → target URL) to the model.
   - Link lines are valid quotes.
4. **Contact person.** The quote must contain the name. v1 quoted only the line "Geschäftsführer" when the name stood on the next line. Academic degrees are dropped from the name.
5. **Verifier leniency.** Differences in letter case or spacing ("Gmbh" vs "GmbH") no longer lead to a rejection.
6. **Reason before verdict.** In v1 the JSON schema asked for the verdict first, and the verifier often wrote "daher bestätigt" in the reason but returned "verworfen". v2 orders the schema reason → verdict and tells the model that the verdict must match the reason.

**v2 was scored twice on this sample.**

- **First scoring: changes 1–5.** Prompt changes plus the link list, with the verifier still deciding before reasoning: 35 errors (21.6%).
- **Second scoring: change 6 added.** Reason before verdict was added after that result: 26 errors (16.0%), the figure in the table.

So 16.0% reflects two looks at the test set, not one.

v1 is kept exactly as recorded: the link list and the field order are switched per version in `llm_enrich.VERSION_OPTS`. That keeps its cached responses valid and the comparison reproducible.

## Caveats

- **Tuned on the test set.** v2 was improved on the errors of these same 30 rows, so its score is optimistic. An independent sample is needed to measure it honestly.
- **Missing input.** Remaining agent misses are mostly rows without usable input: a site that renders only via JavaScript, no Impressum page, or facts that sit on subpages the agent doesn't read (e.g. the Einheiten figure on an "Über uns" page).
- **Disputed contact person.** For one row the truth names a different partner than the first person in the Impressum. Every system that follows the "first person in the Impressum" rule fails there.
- **Haiku 5.5 rejects `temperature`.** It returns a 400 for any non-default value, so the agent can't run at temperature 0. Repeat runs stay reproducible through the response cache instead.

## Cost

- **Total spend:** $0.54 across 364 paid calls (3.9 M input / 0.31 M output tokens). Each call was paid once, and reruns are served from `data/cache/llm_*.json` at no cost. Of that:
  - full run with prompt v2: 193 rows with a website, of which 143 had a cached homepage plus Impressum. 283 responses, $0.43 ($0.34 new in that run)
  - eval runs with v1 and v2 on the sample: ≈ $0.16
- **Prices:** claude-haiku-5-5 at $0.10 / $0.50 per million input / output tokens, for prompts up to 100K tokens.
