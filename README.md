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
planned and implemented rounds. It is found automatically (`--lookup FILE`, the `ADMIN_LOOKUP` variable, `lookup.xlsx`
or `data/lookup.xlsx` next to the input or the script, or `../data`); `--no-lookup` switches it off.

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
