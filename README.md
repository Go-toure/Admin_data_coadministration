# Admin data co-administration

Adds the co-administration columns to `AFRO_admin_data.xlsx` (nOPV2 + bOPV campaigns), following the
feedback from Valentina Bigoni (WHO AFRO). Run it every time a new version of the admin data is built.

```
python add_coadmin_columns.py INPUT.xlsx OUTPUT.xlsx --audit AUDIT.xlsx [--fix-poppolio [--model-poppolio]] [--feedback FEEDBACK.xlsx]
```

Needs Python 3, `pandas` and `openpyxl` (`scikit-learn` for `--model-poppolio`). Re-running it on its own output (without `--fix-poppolio`) gives the same file.

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

`CV_nOPV2` and `CV_bOPV` = vaccinated / `PopPolio`. The "Children Reached" sheet formulas are updated to
the new column positions. The audit workbook lists every merged, relabelled and flagged record.

Vaccine_type clean-up: `Nopv2`, `nOPV3` ... `nOPV14` (drag-fill error) become `nOPV2`; `bopv` becomes `bOPV`.

## `--fix-poppolio` (optional)

Run it once, on the raw file: the corrections are logged and a second pass could adjust a few more populations.

Population fixes, all logged in the audit sheet `PopPolio_corrected`:

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

## `--model-poppolio` (optional, needs `--fix-poppolio`)

For the records still above 1.5, `PopPolio` is estimated with a gradient boosting model trained on the
district's own history: the other campaigns of the same district (more than 0.3 years apart) with a
credible coverage give the nearest, median, previous and next population, the vaccinated history, the
median coverage, the date and the country. The model is validated with 5-fold cross-validation grouped by
district; the results are in the audit sheet `Model_backtest` (on the 26 Oct 2026 data: median error 10%,
70% of estimates within 20%). An estimate is applied only when it gives a coverage between 0.5 and 1.5.
Districts without at least 2 credible campaigns are left unchanged. These are estimates, not reported
data, and each one is listed in `PopPolio_corrected` with its reason.

## `--feedback FILE.xlsx` (optional)

Writes a workbook to send to the countries / focal points, with one sheet per question and empty yellow
columns for their answer:

| Sheet | Content |
|---|---|
| `By_country` | counts per country for each sheet below |
| `To_correct` | records still above 1.5 coverage (more children vaccinated than the target population) |
| `Estimates_to_confirm` | populations that were estimated (district median or model) and need confirmation |
| `Combined_to_validate` | combined "nOPV2 & bOPV" records where nOPV2 and bOPV were given the same total by default; the country gives the real nOPV2 and bOPV totals |

## Data dictionary

`build_data_dictionary.py` builds `AFRO_admin_data_dictionary.xlsx`, to upload to SharePoint next to the admin
data. Completeness, ranges and examples are computed from the data file, the descriptions are in the script.

```
python build_data_dictionary.py AFRO_admin_data.xlsx AFRO_admin_data_dictionary.xlsx
```

If a column is added to the admin data, the script stops and asks for its description.
