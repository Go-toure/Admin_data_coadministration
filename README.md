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
