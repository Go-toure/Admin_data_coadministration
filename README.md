# Admin data update (co-administration, population corrections, data dictionary)

One script, `add_coadmin_columns.py`, does the whole update of `AFRO_admin_data.xlsx` (nOPV2 + bOPV campaigns,
following the feedback from Valentina Bigoni, WHO AFRO): the co-administration columns, the population
corrections, the data dictionary, the audit workbook and the feedback workbook for the countries.

```
python add_coadmin_columns.py INPUT.xlsx [--out-dir DIR]
```

Everything is written under `DIR` (default `coadmin_output`):

| File | What it is |
|---|---|
| `upload_to_sharepoint/AFRO_admin_data.xlsx` | the admin data, `Sheet1` only |
| `upload_to_sharepoint/AFRO_admin_data_dictionary.xlsx` | the data dictionary |
| `review/AFRO_admin_data_coadmin_audit.xlsx` | every merge, relabel and correction, for checking |
| `review/Feedback_suspicious_coverage.xlsx` | to send to the countries / focal points |

Needs Python 3, `pandas`, `openpyxl` and `scikit-learn` (without it the population model step is skipped).
Switches to leave a step out: `--no-fix-poppolio`, `--no-model-poppolio`, `--no-dictionary`, `--no-audit`,
`--no-feedback`.

After the run: open the admin data in Excel and save once (this stores the `check_dp` values), upload the
admin data and the dictionary to SharePoint, send the feedback workbook, then run the SharePoint backup.

Re-running on its own output is safe: merged records stay merged, and only populations that are still wrong
are corrected (a few more district medians may change). The script recognises a file that was already processed
(it has the co-administration columns) and then skips the population model by itself, because the model would
use its own earlier estimates as history. `--force-model-poppolio` overrides this. On such a file the feedback workbook lists only what is still wrong or
new; the estimates of the first run are not remembered, so keep the workbook from that run.

## Repair of copy and drag-fill errors

Before anything else the script looks for the errors made by dragging or pasting a value down in Excel and fixes
them. Two references are used: the **campaign lookup** (primary) and the SharePoint backups.

**The lookup** (`lookup.xlsx`: Round Number, OBR Name, Vaccines, Country, Round Start Date, Round Status) lists the
planned and implemented rounds. The file lives in this repository, next to the script (`lookup.xlsx`), and is refreshed by the lookup fetch pipeline
(`refresh_preparedness_lookup.py` of the IM workflow copies it here after every successful refresh). Another file can be
given with `--lookup FILE`; `--no-lookup` switches it off.

What is detected:

- **Response of another country**: a campaign name that the lookup lists for a different country, or a name starting
  with the code of a different country in the backups (e.g. `SSD-...` on DRC, Botswana or Zimbabwe rows). The lookup
  finds copies that are already in every backup, where the code is shared by two countries and the backups cannot
  tell the owner (for example the Malawi name on 303 Angola rows, the Namibia name on Cameroon and Niger).
- **`nOPV3` ... `nOPV15`** in `Vaccine_type`: a drag-fill of `nOPV2`.

How a wrong Response is restored, in this order:

1. **The lookup**: the one round of the same country and vaccine that started on the row's SIA_date gives the
   Response and the round number (shifted by the way the data numbers the rounds of that Response, as the lookup
   starts some campaigns at Round 0). When the districts of the data point to another campaign, nothing is changed
   and the rows are listed as `UNRESOLVED` with both answers.
2. **The same row in the newest earlier backup** that holds a valid value.
3. **The other rows of the same country, date and round** (majority).
4. **The footprint of the campaign in the data** (`--no-infer` switches this off): a block of rows (5 or more with
   the same wrong Response, date and round) takes the Response of the same country and vaccine that covers the same
   districts (at least 80% both ways, accents ignored); a single row takes the one campaign round of the same
   country, vaccine and province, within 90 days, that covers at least 90% of the province's districts and lacks
   this district (its Response, SIA_date and Round_Add all come from that round, because the date and the round were
   copied from the other country together with the Response). Only a unique answer is used.

For `Vaccine_type` the backups come first, then what the Response name says (bOPV / nOPV), otherwise `nOPV2`.

Every repair (row, value found, value restored, source, confidence, and whether the lookup confirms it) is in the
audit sheet `Repairs`. `Source` starts with `lookup:`, `backup`, `other rows` or `INFERRED:`. What is still unclear
is not guessed: it is `UNRESOLVED`. Renames that stay within the same country are kept.

### Lookup cross-check of the whole file

The audit workbook also has two sheets built from the lookup, after the repairs:

- `Lookup_check`: one line per Country / Response / SIA_date / Round_Add of the data with its status: OK, Response not
  in the lookup, Response listed for another country, Date differs from the lookup by N days, Date not in the lookup for
  this Response, Round differs from the lookup (with the lookup's round and start date). Nothing is changed from this
  sheet: it is the list to take to the countries (for example a round that the data has twice and the lookup once).
- `Lookup_no_data`: rounds of the lookup (planned or implemented) with no records in the data.

### Rounds that only exist as a copy

If the data has two rounds of the same Response on the same date, the lookup has only one of them for that date, and
the extra round is not a round of that Response in the lookup, the script checks whether the extra round copies
the real one (the same `TotalNbrVaccPolio` in at least 80% of the districts they share). If so, the lookup is
trusted and the extra round is removed from the upload file (example: `CHD-2024-11-1_nOPV` has `Rnd1` and `Rnd2` on
2024-11-22 with the same totals; the lookup has Round 1 only, so the 158 `Rnd2` rows go). The removed rows are
listed in the audit sheet `Removed_copies`, and the backups still hold them. `--keep-copied-rounds` switches this
off. A round that the lookup does have on another date is never removed (then the date is wrong, see `Lookup_check`).

## Quality checks (merged from the admin portal)

All the checks of the portal (`admin_quality_checks.R`) run on the finished file, so one script both updates and checks
the admin data. They need `geo_reference.csv` (country, province, district; kept in this repository next to the script,
rebuilt from the portal's `input/geo_reference.rds`; another file with `--geo FILE`). Severity: **Error** = fix before the
record is used, **Warning** = needs a human look, **Info** = good practice.

| Check | Severity |
|---|---|
| Missing key field (Country, Province, District, Vaccine_type, Response, Round_Add, SIA_date) | Error |
| Country name not in the geography reference / alternative spelling (CAMEROUN, TANZANIA ...) | Error / Info |
| Province of another country (stray copy) / used by many records of the country (other naming system) / not in the reference | Error / Warning / Info |
| SIA_date outside 2015 - next year; more than 4 months ahead | Error / Warning |
| Report date before the SIA, in the future, or more than a year after | Warning / Warning / Info |
| Negative numbers; vaccinated but no target; target larger than total population | Error / Error / Warning |
| Coverage above the ceiling (150%); coverage above 120%; reported coverage differs from vaccinated / target | Warning / Info / Warning or Info |
| Target population 3 times above or below the district's other rounds (possible extra / missing digit) | Warning |
| Co-administration: combined record, combined and separate for the same district and date, identical figures for both vaccines | Info / Warning / Info |
| Co-administration: `TotalNbrVaccPolio` is not the larger of `Tot_Vacc_nOPV2` and `Total_Vacc_bOPV` (new) | Error |
| Vaccine type spelling | Info |
| Duplicate record | Warning |

The report is `review/AFRO_admin_data_quality_report.xlsx` (Overview, By check, By country, Issues, Suggested fixes, Fixes
applied). Fixes made by the script: a province that belongs to **another country** and is rare in its own country (fewer than
5 records) is replaced with the province of that district in the geography reference (or the one its other records use);
names used by many records of a country (Rwanda `EAST`, Malawi `NORTH`) are only reported. Everything else is reported,
not changed. Thresholds: `--cv-ceiling` (default 150) and `--pop-factor` (default 3); `--no-quality` and `--no-fix-province`
switch the checks / the province fix off; `--fail-on-error` exits with code 2 when records with an Error remain.

`--check-only` runs only the checks, on a file as it is (nothing is updated), like the portal's daily
`run_admin_quality_check.R`; `--country Nigeria` limits it to one country.

## Duplicates, text dates and the repository (merged from the admin portal)

- **Duplicate key**: Country | Province | District | Admin_3 | SIA_date | Round_Add | Vaccine_type | Response, cleaned as in
  the portal (the portal's key without Admin_3; sub-districts of one district are not duplicates). Records with an empty key
  field are not treated as duplicates (they are reported as missing key fields).
- **Exact duplicates** (same key and every other value the same) are removed by the update, the first copy is kept
  (`--no-dedupe` switches this off). They are listed in `Removed_duplicates` of the quality report.
- **Same key, different values** are never removed by default: they are listed in `Duplicate_groups` (group, record number,
  the columns that differ, the most complete record marked `keep`). `--resolve-duplicates complete` keeps the most complete
  record of each group, `--resolve-duplicates first` keeps the first one (what the portal does).
- **Text dates**: a SIA_date written as text ("2020-09-18 00:00:00 UTC", "18-sept-20", "18/09/2020") is turned into a real
  date (logged in `Fixes applied`).
- **Repository** (`--repository AFRO_admin_data_repository.xlsx`): records whose key already exists in the portal's repository
  are listed in `Already_in_repository` (`--skip-existing` leaves them out of the output, like the portal's "Skip duplicates"),
  and the repository rows of the same countries give context to the target population check; only the records of the file
  are reported.

## Filling empty key fields and empty coverage

Two gaps are filled by default, and every fill is logged:

- **Empty `Response`, from the lookup** (`--no-fill-response` switches this off). Records with a blank Response are
  grouped by country, SIA_date and round. For each group the lookup must give **one** campaign round of the same country
  and a compatible vaccine that starts on that date (confidence `high`), or within 21 days of it with the same round number
  (confidence `medium`). When the lookup name carries a province (for example `Kwara Response`), only the rows of that province
  are filled. A group is left empty if the lookup gives zero or several answers. The spelling is the one the data already
  uses for that Response. Fills are in the audit sheet `Repairs` (`Source` starts with `lookup:`). **Confirm the `medium` fills
  with the countries**: the date in the data differs from the lookup by a few days.
- **Empty `CVPolio`, from the target** (`--no-fill-cvpolio` switches this off): where the vaccinated total is known and the target
  is known, `CVPolio` = `TotalNbrVaccPolio` / `PopPolio` (population = target). Rows with a target but no vaccinated total stay
  empty. The fills are in `Fixes applied` of the quality report (column `CVPolio`).

What the lookup cannot fix stays an Error ("Missing key field") for the countries: Nigeria records from 2022 to January 2023
(before the lookup starts), Ethiopia records with a blank District (no other identifier in the record or in any backup), and the
few records with a blank Province that has no match.

## Filling an empty District from a recovered file

`--district-source FILE` fills empty `District` values from a recovered or original file (a country template with
`Prov_Name` / `Dist_Name`, or an older admin file; xlsx or csv, first sheet). Records with an empty District are grouped by
country, Response, round and province and matched one to one with the records of the file in the same group:

1. same `PopPolio` and same vaccinated total (confidence high);
2. same `PopPolio` and a vaccinated total within 10%, only if unique (confidence medium);
3. same vaccinated total and a different `PopPolio`, only if unique (high from 1,000 vaccinated, otherwise medium);
4. one record and one district of the file left unmatched in the same province (medium).

Each record of the file is used once; what does not match stays empty. The district written is the file's `Dist_Name` (the woreda,
which is what the data uses as District for Ethiopia; `Zone_Name` is the level above), in capitals, and in the spelling the data
already uses for that district when there is one (the same name once WORHO / TOWN / CITY are ignored, or at least 90% alike).
Every fill is in the audit sheet `Repairs` (`Source` starts with `recovered file:`). Example: the Ethiopia file `AFRO-ETH FIVE REGIONS R1`
filled 279 of the 282 records of `ETH-2024-08-nOPV2` round 1.

## Publishing to SharePoint (`publish_admin_data.py`)

A separate script puts the latest corrected file on SharePoint, in place of the one that is there. Double-click
`publish_admin_data.bat` or run `python publish_admin_data.py` (`--dry-run` runs the checks only).

1. It takes the highest `AFRO_admin_data_vNN.xlsx` in `coadmin_output/upload_to_sharepoint` (or the scope
   `upload_to_sharepoint` folder), or the file given with `--file`.
2. It checks the file: `Sheet1` only, the key and co-administration columns, about 38 countries (not a country file), and
   that `check_dp` holds values (the file was opened and saved once in Excel).
3. It downloads the live file and compares: no column lost, not more than 10% fewer records, not the same file. If the live file was modified **after** the file to publish was saved (someone else edited it), it stops (`--force` after you checked what changed), because publishing would overwrite their work.
4. It backs up the live file with `SharePoint_version_backup` and checks the backup holds exactly that file (SHA-256).
   If not, it stops before touching SharePoint.
5. It uploads the file as `AFRO_admin_data.xlsx` **into the existing SharePoint file** (same file, same link). SharePoint
   refuses the upload if someone changed the file meanwhile or has it open. The upload is checked by downloading it back
   (SharePoint rewrites the document properties of an Office file, so the bytes can differ slightly; the sheet data must match).
6. Storage: SharePoint keeps every overwrite as a full extra copy, so the script deletes the version history of that file
   (the local backup keeps the history) and tries to empty the deleted copies from the site recycle bin. If the recycle bin
   cannot be emptied with the app's permissions, it says so; then empty it in Site contents > Recycle bin > Second-stage
   recycle bin, or ask IT. `--keep-versions` and `--no-purge-bin` switch these off.
7. It runs the backup again, so the file just published is the newest local version.

Credentials come from the environment or a `secrets.env` (`--secrets FILE`; otherwise the backup repository's and the IM
workflow's `config/secrets.env` are tried). The dictionary file is not part of this script.

## Which files it accepts

- The full regional admin data, raw or already processed (it recognises a processed file, see above).
- An older or a country template: it only needs the columns `Country`, `Province`, `District`, `SIA_date`
  (`SIA-date` is renamed), `Response`, `Vaccine_type`, `TotalNbrVaccPolio`, `PopPolio`, `Round_Add` and `Year`,
  plus `CVPolio` for the population corrections. If one is missing it stops before writing anything and lists
  them. Extra columns are kept; columns the dictionary does not know are marked "No description yet".
- The sheet may have any name (the first sheet is used; the output sheet is `Sheet1`).
- A header-only file stops with a message; with little history the population model is skipped by itself.
- A file with fewer than 25 countries (a country file) is processed, but the script warns twice that the output
  is not the regional file and must not be uploaded to SharePoint as `AFRO_admin_data.xlsx`.

## What it does

Co-administered campaigns are reported in different ways: separate nOPV2 and bOPV records, one combined
"nOPV2 & bOPV" record, or both in different campaigns. A campaign is co-administered when its `Response`
contains both "nOPV" and "bOPV", or when a district has an nOPV2 record and a bOPV record for the same
`Response`, year and round. The second case covers countries that did not put both vaccines in the Response
name (DRC and South Sudan, about 1,090 districts); their `Response` is left as reported.

New columns, placed right before `Response`:

`Tot_Vacc_nOPV2 | CV_nOPV2 | Total_Vacc_bOPV | CV_bOPV | Vaccine_type | TotalNbrVaccPolio | CVPolio`

`Vaccine_type` and `TotalNbrVaccPolio` move here from their old position. `CVPolio` (as stored in the source,
never recalculated) follows `TotalNbrVaccPolio` so the vaccinated total and its coverage sit together; the
dashboards read it by name. A vaccine that was not used has a total of 0 (not a blank) and no coverage.
`check_dp` is rewritten as a live formula (Country, Province, District, SIA_date, TotalNbrVaccPolio, Response)
on the new column letters. The file is written with Sheet1 only; open it in Excel and save once so the
`check_dp` values are stored before it is uploaded.

| Situation | Result |
|---|---|
| Not co-administered | `Vaccine_type` unchanged. `TotalNbrVaccPolio` goes in the column of the vaccine used. |
| Combined record | `Tot_Vacc_nOPV2` = `Total_Vacc_bOPV` = `TotalNbrVaccPolio`. |
| Separate nOPV2 and bOPV records (same Response, round, district) | Merged into one record, `Vaccine_type` = "nOPV2 & bOPV". `TotalNbrVaccPolio` = the **maximum** of the two (same children, not a sum). |
| nOPV2-only record in a co-administered Response where other districts report separate records | Kept as single vaccine, listed in the audit for review. |

`CV_nOPV2` and `CV_bOPV` = vaccinated / `PopPolio`. If the input has a "Children Reached" sheet, its formulas are updated to
the new column positions. The audit workbook lists every merged, relabelled and flagged record.

Vaccine_type clean-up: `Nopv2`, `nOPV3` ... `nOPV14` (drag-fill error) become `nOPV2`; `bopv` becomes `bOPV`.

## Population corrections (default; `--no-fix-poppolio` to skip)

All logged in the audit sheet `PopPolio_corrected`:

1. Vaccinated / `PopPolio` above 1.5 while `CVPolio` is plausible (up to 1.5), or above 5 while `CVPolio` is less
   than half of it: `PopPolio` = vaccinated / `CVPolio`, rounded to a whole number.
2. Vaccinated / `PopPolio` below 0.5 while `CVPolio` is plausible (0.5 to 1.5): `PopPolio` = vaccinated /
   `CVPolio`, rounded to a whole number. In these records `PopPolio` is usually the total population, or
   a province figure repeated in every district, and not the target population. (Where both coverages are
   already between 0.5 and 1.5 `PopPolio` is left as reported.)
3. Decimal populations are rounded to whole numbers.
4. Coverage still above 1.5: `PopPolio` = median of the same district's other polio records in the same
   year, only if that gives a coverage between 0.5 and 1.5.

Whatever is still above 1.5 is listed in the audit sheet `CV_still_above_1.5` for the focal points.

## Population model (default; `--no-model-poppolio` to skip)

For the records still above 1.5, `PopPolio` is estimated with a gradient boosting model trained on the
district's own history: the other campaigns of the same district (more than 0.3 years apart) with a
credible coverage give the nearest, median, previous and next population, the vaccinated history, the
median coverage, the date and the country. The model is validated with 5-fold cross-validation grouped by
district; the results are in the audit sheet `Model_backtest` (on the 26 Oct 2026 data: median error 10%,
70% of estimates within 20%). An estimate is applied only when it gives a coverage between 0.5 and 1.5.
Districts without at least 2 credible campaigns are left unchanged. These are estimates, not reported
data, and each one is listed in `PopPolio_corrected` with its reason.

## Feedback workbook (`Feedback_suspicious_coverage.xlsx`)

A workbook to send to the countries / focal points, with one sheet per question and empty yellow
columns for their answer:

| Sheet | Content |
|---|---|
| `By_country` | counts per country for each sheet below |
| `To_correct` | records still above 1.5 coverage (more children vaccinated than the target population) |
| `Estimates_to_confirm` | populations that were estimated (district median or model) and need confirmation |
| `Combined_to_validate` | combined "nOPV2 & bOPV" records where nOPV2 and bOPV were given the same total by default; the country gives the real nOPV2 and bOPV totals |

## Data dictionary

`AFRO_admin_data_dictionary.xlsx` is built at the end of the run from the finished data, so the completeness,
ranges and examples always match the file. The descriptions are written in the script (`COLUMNS`). If a column is
added to the admin data, the script stops at that point and asks for its description, so the dictionary never
falls behind the data.
