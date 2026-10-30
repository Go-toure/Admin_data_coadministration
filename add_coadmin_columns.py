#!/usr/bin/env python3
"""
add_coadmin_columns.py  --  one script for the AFRO admin data update
=====================================================================

Takes AFRO_admin_data.xlsx and produces, in one run: the admin data with the co-administration columns and the
population corrections, the data dictionary, the audit workbook and the feedback workbook for the countries.

Why
---
Some co-administered SIAs (nOPV2 + bOPV) are reported in three different ways
(feedback from Valentina Bigoni):

  * separate vaccine-specific records  - one nOPV2 record + one bOPV record
  * combined records                   - one record, Vaccine_type "nOPV2 & bOPV"
  * mixed practices                    - both, in different campaigns / rounds

A campaign is co-administered when its Response name contains both "nOPV" and
"bOPV", or when a district has an nOPV2 record and a bOPV record for the same
Response, year and round (some countries, e.g. DRC and South Sudan, did not put
both vaccines in the Response name).  This script makes every record carry the same structure, so
co-administration can be reported correctly:

    ... Year | Round_Add | Tot_Vacc_nOPV2 | CV_nOPV2 | Total_Vacc_bOPV | CV_bOPV |
        Vaccine_type | TotalNbrVaccPolio | CVPolio | Response | ...

Rules
-----
  * Not co-administered      Vaccine_type unchanged. The vaccinated total
                             (TotalNbrVaccPolio) goes in the column of the
                             vaccine used (nOPV2 -> Tot_Vacc_nOPV2,
                             bOPV -> Total_Vacc_bOPV); the other stays empty.
  * Combined record          (Vaccine_type already holds nOPV and bOPV)
                             Tot_Vacc_nOPV2 = Total_Vacc_bOPV = TotalNbrVaccPolio
  * Separate records         an nOPV2 record and a bOPV record of the same
                             Response / Round / district are merged into ONE
                             record: Tot_Vacc_nOPV2 and Total_Vacc_bOPV come
                             from the two records, TotalNbrVaccPolio = the
                             MAXIMUM of the two (not the sum: it is the same
                             children), Vaccine_type = "nOPV2 & bOPV".
  * Single record in a       (no separate vaccine records anywhere in that
    co-administered Response  Response/Round) treated as a combined record.
  * CV_nOPV2 / CV_bOPV       = Tot_Vacc / PopPolio (target population).

  An nOPV2-only record in a Response/Round where other districts DO report
  separate nOPV2 and bOPV records is kept as a single-vaccine record (no bOPV
  record was reported there) and listed in the audit file for review.

  Vaccine_type clean-up: "Nopv2" -> "nOPV2"; "nOPV3".."nOPV14" (drag-fill
  error) -> "nOPV2"; "bopv" -> "bOPV".

The script is safe to run again on its own output (the corrections are then applied only to what is still wrong).

Usage
-----
  python add_coadmin_columns.py INPUT.xlsx [--out-dir DIR]

Everything is written under DIR (default: coadmin_output):

  DIR/upload_to_sharepoint/AFRO_admin_data.xlsx             the admin data (Sheet1 only)
  DIR/upload_to_sharepoint/AFRO_admin_data_dictionary.xlsx  the data dictionary
  DIR/review/AFRO_admin_data_coadmin_audit.xlsx             every change, for checking
  DIR/review/Feedback_suspicious_coverage.xlsx              to send to the countries / focal points

By default it also corrects PopPolio (CVPolio, the district's other rounds, then a model on the district's
history; needs scikit-learn). Switches to leave a step out: --no-fix-poppolio, --no-model-poppolio,
--no-dictionary, --no-audit, --no-feedback.

After the run: open the admin data in Excel and save once (stores the check_dp values), upload both files
to SharePoint, send the feedback workbook, then run the SharePoint backup.
"""
import argparse
import re
import sys
from copy import copy

import numpy as np
import openpyxl
import pandas as pd
import os
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils import get_column_letter as L

SHEET = "Sheet1"
SUMMARY_SHEET = "Children Reached"
BLOCK = ["Tot_Vacc_nOPV2", "CV_nOPV2", "Total_Vacc_bOPV", "CV_bOPV",
         "Vaccine_type", "TotalNbrVaccPolio"]
BEFORE = "Response"
KEY = ["Country", "Province", "District", "Admin_3", "Year", "Round_Add", "Response"]
COMBINED = "nOPV2 & bOPV"
CV_HIGH = 1.5
CV_EXTREME = 5.0   # coverage above this is not credible even when CVPolio is also above CV_HIGH
CV_LOW = 0.5       # coverage below this is not credible when the stored CVPolio is plausible

A_NONE = "Not co-administered"
A_COMBINED = "Combined record (as reported)"
A_MERGED = "Separate vaccine records merged into one"
A_RESP = "Co-administered by Response (single record)"
A_KEPT = "Single-vaccine record kept (no bOPV record in a co-administered Response)"


def num(s):
    return pd.to_numeric(s, errors="coerce")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src", help="the admin data workbook (AFRO_admin_data.xlsx)")
    ap.add_argument("--out-dir", default="coadmin_output", help="folder for all the outputs (default: coadmin_output)")
    ap.add_argument("--no-fix-poppolio", action="store_true",
                    help="do not correct PopPolio (coverage above %.1f or below %.1f with a plausible CVPolio, "
                         "district median, decimals)" % (CV_HIGH, CV_LOW))
    ap.add_argument("--no-model-poppolio", action="store_true",
                    help="do not estimate the PopPolio that is still above %.1f with the district-history model" % CV_HIGH)
    ap.add_argument("--force-model-poppolio", action="store_true",
                    help="run the population model even when the input was already processed "
                         "(not recommended: the model would use its own earlier estimates as history)")
    ap.add_argument("--backups", help="folder with the SharePoint backups (AFRO_admin_data__*.xlsx), the reference for "
                                      "the repair of copy errors (found automatically next to this repository)")
    ap.add_argument("--no-repair", action="store_true", help="do not look for copy / drag-fill errors")
    ap.add_argument("--no-infer", action="store_true",
                    help="do not infer the campaign of a copied row from the rest of the file when the backups "
                         "do not hold the right value (the row is then only listed as UNRESOLVED)")
    ap.add_argument("--lookup", help="campaign lookup workbook (Round Number, OBR Name, Vaccines, Country, Round Start "
                                     "Date, Round Status) used to cross-check the repairs and the Response / date / round "
                                     "of every campaign (default: lookup.xlsx next to this script, kept up to date by the "
                                     "lookup fetch pipeline)")
    ap.add_argument("--no-lookup", action="store_true", help="do not use the campaign lookup")
    ap.add_argument("--keep-copied-rounds", action="store_true",
                    help="do not remove the rounds that the lookup does not have and that copy another round of the "
                         "same Response and date (they are removed by default and listed in the audit)")
    ap.add_argument("--align-response-names", nargs="*", metavar="COUNTRY", default=None,
                    help="rename a Response to the exact name of the lookup when it is clearly the same campaign (same country, "
                         "start date within 7 days, the lookup name starts with the Response). Done by default for all countries; "
                         "give countries (e.g. --align-response-names MALAWI) to restrict it to them")
    ap.add_argument("--no-align-response-names", action="store_true",
                    help="keep the Response names as reported even when the lookup spells them differently")
    ap.add_argument("--district-source", metavar="FILE",
                    help="a recovered or original file (e.g. a country template with Prov_Name / Dist_Name, or an older admin file) "
                         "used to fill empty District values: records of the same Response and round are matched on province "
                         "and the target / vaccinated figures, and only a unique answer is used")
    ap.add_argument("--geo", help="geography reference (country, province, district) for the province / country check "
                                  "(default: geo_reference.csv next to this script)")
    ap.add_argument("--no-quality", action="store_true", help="do not run the quality checks (merged from the admin portal)")
    ap.add_argument("--no-fix-province", action="store_true",
                    help="do not correct a province that belongs to another country (it is then only reported)")
    ap.add_argument("--check-only", action="store_true",
                    help="only run the quality checks on the file as it is (no update, no changes) and write the report")
    ap.add_argument("--country", help="with --check-only: limit the checks to this country")
    ap.add_argument("--cv-ceiling", type=float, default=150.0,
                    help="quality checks: coverage above this %% is flagged (default 150)")
    ap.add_argument("--pop-factor", type=float, default=3.0,
                    help="quality checks: target population this many times above / below the district median is flagged (default 3)")
    ap.add_argument("--fail-on-error", action="store_true",
                    help="exit with code 2 when the quality checks still find records with an Error")
    ap.add_argument("--no-fill-response", action="store_true",
                    help="do not fill an empty Response from the campaign lookup (country, start date, round)")
    ap.add_argument("--no-fill-cvpolio", action="store_true",
                    help="do not fill an empty CVPolio with TotalNbrVaccPolio / PopPolio")
    ap.add_argument("--no-dedupe", action="store_true",
                    help="do not remove exact duplicates (same key and every other value the same); they are removed by default")
    ap.add_argument("--resolve-duplicates", choices=["none", "complete", "first"], default="none",
                    help="records with the same key but different values: none = only list them in the quality report "
                         "(default), complete = keep the most complete record, first = keep the first one (the portal's rule)")
    ap.add_argument("--repository", help="the portal repository (AFRO_admin_data_repository.xlsx): records whose key is already "
                                         "in it are listed (Already_in_repository) and its rows give context to the target population check")
    ap.add_argument("--skip-existing", action="store_true",
                    help="with --repository: leave out of the output the records that already exist in the repository")
    ap.add_argument("--no-dictionary", action="store_true", help="do not build the data dictionary")
    ap.add_argument("--no-audit", action="store_true", help="do not write the audit workbook")
    ap.add_argument("--no-feedback", action="store_true", help="do not write the feedback workbook")
    a = ap.parse_args()

    up_dir = os.path.join(a.out_dir, "upload_to_sharepoint")
    rev_dir = os.path.join(a.out_dir, "review")
    os.makedirs(up_dir, exist_ok=True)
    os.makedirs(rev_dir, exist_ok=True)
    a.dst = os.path.join(up_dir, "AFRO_admin_data.xlsx")
    a.dictionary = None if a.no_dictionary else os.path.join(up_dir, "AFRO_admin_data_dictionary.xlsx")
    a.audit = None if a.no_audit else os.path.join(rev_dir, "AFRO_admin_data_coadmin_audit.xlsx")
    a.feedback = None if a.no_feedback else os.path.join(rev_dir, "Feedback_suspicious_coverage.xlsx")
    a.quality = None if a.no_quality else os.path.join(rev_dir, "AFRO_admin_data_quality_report.xlsx")
    if a.check_only:
        run_check_only(a, rev_dir)
        return
    a.fix_poppolio = not a.no_fix_poppolio
    a.model_poppolio = a.fix_poppolio and not a.no_model_poppolio
    if a.model_poppolio:
        try:
            import sklearn  # noqa: F401
        except ImportError:
            print("scikit-learn is not installed: the population model step is skipped "
                  "(pip install scikit-learn to include it)", flush=True)
            a.model_poppolio = False

    print("loading", a.src, flush=True)
    wb = openpyxl.load_workbook(a.src)
    sheet_name = SHEET if SHEET in wb.sheetnames else wb.sheetnames[0]
    if sheet_name != SHEET:
        print("No sheet called %s: using the first sheet (%s). The output sheet is named %s." % (SHEET, sheet_name, SHEET), flush=True)
        if SUMMARY_SHEET in wb.sheetnames:
            del wb[SUMMARY_SHEET]  # its formulas point to the original sheet name
    ws = wb[sheet_name]
    header = []
    for c in ws[1]:
        header.append(c.value)
    while header and header[-1] is None:
        header.pop()
    header = ["SIA_date" if h == "SIA-date" and "SIA_date" not in header else h for h in header]  # same fix as the portal
    ncol = len(header)
    old_letter = {h: get_column_letter(i + 1) for i, h in enumerate(header)}
    missing_req = [need for need in ["Country", "Province", "District", "SIA_date", "Response", "Vaccine_type",
                                     "TotalNbrVaccPolio", "PopPolio", "Round_Add", "Year"] if need not in header]
    if missing_req:
        sys.exit("The input is not an admin data file: these columns are missing: " + ", ".join(missing_req))
    # a file that already has the co-administration columns was processed before: its populations include the
    # earlier model estimates, so the model is not run again (it would use them as history)
    already_processed = all(b in header for b in ("Tot_Vacc_nOPV2", "Total_Vacc_bOPV"))
    if already_processed and a.model_poppolio and not a.force_model_poppolio:
        print("The input already has the co-administration columns, so it was processed before: "
              "the population model is skipped (use --force-model-poppolio to run it anyway). "
              "The feedback workbook then lists only what is still wrong or new; the estimates of the earlier "
              "run are not remembered, keep the workbook from that run.", flush=True)
        a.model_poppolio = False
    if a.fix_poppolio and "CVPolio" not in header:
        sys.exit("The input has no CVPolio column, which the population corrections need. "
                 "Use a file that has it, or run with --no-fix-poppolio.")

    cells = list(ws.iter_rows(min_row=2, max_col=ncol))
    cells = [r for r in cells if any(c.value is not None for c in r)]
    n = len(cells)
    df = pd.DataFrame([[c.value for c in r] for r in cells], columns=header, dtype=object)
    df["_row"] = np.arange(2, n + 2)
    print("rows:", n, "cols:", ncol, flush=True)
    if n == 0:
        sys.exit("The file has no data rows (only the header): nothing to update.")
    n_countries = int(df["Country"].nunique())
    partial = n_countries < 25   # the regional file has about 38 countries
    if partial:
        print("WARNING: the file has only %d countries. It is not the full regional admin data: use the output "
              "for checking only, do NOT upload it to SharePoint as AFRO_admin_data.xlsx." % n_countries, flush=True)

    cvp_all = num(df["CVPolio"]) if "CVPolio" in df.columns else pd.Series(np.nan, index=df.index)

    # campaign lookup (planned and implemented rounds): second reference for the repairs and the cross-check
    lk = None
    if not a.no_lookup:
        lk_path = find_lookup(a.lookup, a.src)
        if lk_path:
            try:
                lk = load_lookup(lk_path)
                print("lookup: %s (%d rounds, %d countries)" % (lk_path, len(lk), lk["_ck"].nunique()), flush=True)
            except Exception as e:
                print("lookup not used (%s): %s" % (lk_path, e), flush=True)
        else:
            print("lookup: not found (give it with --lookup): campaigns are not cross-checked", flush=True)

    # copy / drag-fill errors (Response from another country, nOPV3..nOPV15): restored from the backups
    repairs = []
    repl = {}
    if not a.no_repair:
        repairs = repair_copy_errors(df, find_backups(a.backups, a.src), a.src, os.path.join(a.out_dir, "_backup_cache"),
                                     infer=not a.no_infer, lk=lk)
        for r in repairs:
            if r["Restored"] and r["Column"] in ("Response", "Round_Add", "Year", "SIAMTH"):
                repl[(int(r["Excel row"]) - 2, r["Column"])] = r["Restored"]
            elif r["Restored"] and r["Column"] == "SIA_date":
                repl[(int(r["Excel row"]) - 2, "SIA_date")] = pd.Timestamp(r["Restored"]).to_pydatetime()

    # geography reference (country / province / district of the region): a province that belongs to ANOTHER country is
    # a copy error too; it is corrected with the reference before anything else is built on it
    geo = None
    prov_log = []
    if not a.no_quality:
        gp = find_geo(a.geo, a.src)
        if gp:
            try:
                geo = load_geo(gp)
                print("geography reference: %s (%d districts)" % (gp, len(geo)), flush=True)
            except Exception as e:
                print("geography reference not used (%s): %s" % (gp, e), flush=True)
        else:
            print("geography reference: not found (geo_reference.csv): the province / country check is skipped", flush=True)
    if geo is not None and not a.no_fix_province:
        _, corr0, _ = qc_geo(qc_prepare(df), geo)
        if corr0 is not None and len(corr0):
            for r in corr0[corr0["certain"]].itertuples():
                i = int(r.idx)
                df.at[i, "Province"] = r.suggested_province
                repl[(i, "Province")] = r.suggested_province
                prov_log.append({"Excel row": int(df.at[i, "_row"]), "Country": r.Country, "District": r.District,
                                 "Column": "Province", "Found": r.current_province, "Restored": r.suggested_province,
                                 "Source": r.source})
            if prov_log:
                print("province of another country corrected: %d records" % len(prov_log), flush=True)

    # an empty Response is filled from the campaign lookup (country + start date + round)
    if lk is not None and not a.no_fill_response:
        fills = fill_blank_response(df, lk)
        for i, inf in fills.items():
            val = inf["set"]["Response"]
            df.at[i, "Response"] = val
            repl[(i, "Response")] = val
            repairs.append({"Excel row": int(df.at[i, "_row"]), "Country": df.at[i, "Country"], "Province": df.at[i, "Province"],
                            "District": df.at[i, "District"], "SIA_date": _dkey(df.at[i, "SIA_date"]), "Column": "Response",
                            "Found": "", "Rule": "Response is empty", "Restored": val, "Source": "lookup: " + inf["why"],
                            "Confidence": inf["confidence"], "Check": ""})
        if fills:
            print("empty Response filled from the lookup: %d records (%d exact, %d within a few days)" %
                  (len(fills), sum(1 for v in fills.values() if v["confidence"] == "high"),
                   sum(1 for v in fills.values() if v["confidence"] != "high")), flush=True)

    # a Response that is not a name of the lookup but clearly is the same campaign gets the lookup's name
    if lk is not None and not a.no_align_response_names:
        afills = align_response_names(df, lk, a.align_response_names)
        for i, inf in afills.items():
            val = inf["set"]["Response"]
            old_ = df.at[i, "Response"]
            df.at[i, "Response"] = val
            repl[(i, "Response")] = val
            repairs.append({"Excel row": int(df.at[i, "_row"]), "Country": df.at[i, "Country"], "Province": df.at[i, "Province"],
                            "District": df.at[i, "District"], "SIA_date": _dkey(df.at[i, "SIA_date"]), "Column": "Response",
                            "Found": old_, "Rule": "Response is not a name of the lookup", "Restored": val,
                            "Source": "lookup: " + inf["why"], "Confidence": inf["confidence"], "Check": ""})
        if afills:
            print("Response renamed to the lookup's name: %d records (%s)" %
                  (len(afills), ", ".join(sorted({v["set"]["Response"] for v in afills.values()}))), flush=True)

    # an empty District is filled from a recovered file (only when its figures match the record exactly or nearly)
    if getattr(a, "district_source", None):
        try:
            dfills = fill_blank_district(df, a.district_source)
        except Exception as e:
            dfills = {}
            print("district source not used (%s): %s" % (a.district_source, e), flush=True)
        for i, inf in dfills.items():
            val = inf["set"]["District"]
            df.at[i, "District"] = val
            repl[(i, "District")] = val
            repairs.append({"Excel row": int(df.at[i, "_row"]), "Country": df.at[i, "Country"], "Province": df.at[i, "Province"],
                            "District": val, "SIA_date": _dkey(df.at[i, "SIA_date"]), "Column": "District",
                            "Found": "", "Rule": "District is empty", "Restored": val,
                            "Source": "recovered file: " + inf["why"], "Confidence": inf["confidence"], "Check": ""})
        if dfills:
            print("empty District filled from %s: %d records (%d with an exact vaccinated total, %d less certain)" %
                  (os.path.basename(a.district_source), len(dfills),
                   sum(1 for v in dfills.values() if v["confidence"] == "high"),
                   sum(1 for v in dfills.values() if v["confidence"] != "high")), flush=True)

    # Addis Ababa: repeated sub-city names become "<sub-city> WOREDA NN" (rows the recovered file can identify)
    if getattr(a, "district_source", None):
        try:
            rfills = rename_addis_districts(df, a.district_source)
        except Exception as e:
            rfills = {}
            print("Addis Ababa rename not done (%s)" % e, flush=True)
        for i, inf in rfills.items():
            val = inf["set"]["District"]
            df.at[i, "District"] = val
            repl[(i, "District")] = val
            repairs.append({"Excel row": int(df.at[i, "_row"]), "Country": df.at[i, "Country"], "Province": df.at[i, "Province"],
                            "District": val, "SIA_date": _dkey(df.at[i, "SIA_date"]), "Column": "District",
                            "Found": inf["old"], "Rule": "Addis Ababa: sub-city name repeated, woreda missing", "Restored": val,
                            "Source": "recovered file: " + inf["why"], "Confidence": inf["confidence"], "Check": ""})
        if rfills:
            print("Addis Ababa districts renamed to sub-city + woreda from %s: %d records (%d high confidence, %d medium)" %
                  (os.path.basename(a.district_source), len(rfills),
                   sum(1 for v in rfills.values() if v["confidence"] == "high"),
                   sum(1 for v in rfills.values() if v["confidence"] != "high")), flush=True)

    # SIA_date written as text ("18-sept-20", "2020-09-18 00:00:00 UTC" ...) is turned into a real date
    date_log = []
    for i in df.index[df["SIA_date"].map(lambda x: isinstance(x, str))]:
        dv = parse_sia_text(df.at[i, "SIA_date"])
        if pd.isna(dv):
            continue
        old_ = df.at[i, "SIA_date"]
        df.at[i, "SIA_date"] = dv.to_pydatetime()
        repl[(i, "SIA_date")] = dv.to_pydatetime()
        date_log.append({"Excel row": int(df.at[i, "_row"]), "Country": df.at[i, "Country"], "District": df.at[i, "District"],
                         "Column": "SIA_date", "Found": old_, "Restored": dv.strftime("%Y-%m-%d"), "Source": "date written as text"})
    if date_log:
        print("SIA_date written as text converted: %d records" % len(date_log), flush=True)
    prov_log = prov_log + date_log

    # duplicates (the portal's duplicate check and fix) and records that already exist in the repository
    dup_drop = pd.Series(False, index=df.index)
    dup_log, dup_groups = pd.DataFrame(), pd.DataFrame()
    if not a.no_dedupe:
        dup_drop, dup_log, dup_groups = find_duplicates(df, a.resolve_duplicates)
        n_ex = len(dup_log)
        n_conf = int(dup_groups["duplicate_group"].nunique()) if len(dup_groups) else 0
        print("duplicates: %d exact copies removed; %d groups of records with the same key but different values (%s)" %
              (n_ex, n_conf, "removed as asked" if a.resolve_duplicates != "none" else "listed in the quality report, not removed"),
              flush=True)
    exist_drop = pd.Series(False, index=df.index)
    repo_ctx = None
    existing_rows = pd.DataFrame()
    if a.repository:
        try:
            ex_mask, repo_ctx = existing_in_repository(df, a.repository)
            existing_rows = df.loc[ex_mask, ["_row", "Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response"]] \
                .rename(columns={"_row": "Excel_row"})
            print("repository: %d records of the file already exist in it%s" %
                  (int(ex_mask.sum()), " (left out of the output)" if a.skip_existing else ""), flush=True)
            if a.skip_existing:
                exist_drop = ex_mask
        except Exception as e:
            print("repository not used (%s): %s" % (a.repository, e), flush=True)

    # rounds that only exist as a copy of another round of the same Response and date (lookup has no such round)
    copy_drop = pd.Series(False, index=df.index)
    copy_log = pd.DataFrame()
    if lk is not None and not a.keep_copied_rounds:
        copy_drop, copy_log = find_copied_rounds(df, lk)
        if copy_drop.any():
            print("copied rounds removed: %d rows (%s); listed in sheet Removed_copies of the audit workbook" %
                  (int(copy_drop.sum()), ", ".join(sorted(set(copy_log["Response"].astype(str) + " " + copy_log["Round_Add"].astype(str))))),
                  flush=True)

    pre_drop = copy_drop | dup_drop | exist_drop
    pre_label = pd.Series("", index=df.index)
    pre_label[copy_drop] = "(copied round removed)"
    pre_label[dup_drop] = "(duplicate removed)"
    pre_label[exist_drop] = "(already in the repository)"

    # ------------------------------------------------------------------ logic
    vt_orig = df["Vaccine_type"].astype(object)
    vt = vt_orig.map(lambda x: x.strip() if isinstance(x, str) else x)
    low = vt.map(lambda x: x.lower() if isinstance(x, str) else "")
    fix = low.str.fullmatch(r"nopv\d+") | (low == "bopv")
    vt_norm = vt.copy()
    vt_norm[low.str.fullmatch(r"nopv\d+")] = "nOPV2"
    vt_norm[low == "bopv"] = "bOPV"
    vt_changed = (vt_norm.astype(str) != vt_orig.astype(str))

    resp = df["Response"].map(lambda x: x.lower() if isinstance(x, str) else "")
    resp_both = resp.str.contains("nopv") & resp.str.contains("bopv")
    low_n = vt_norm.map(lambda x: x.lower() if isinstance(x, str) else "")
    vt_both = low_n.str.contains("nopv") & low_n.str.contains("bopv")
    is_n = vt_norm == "nOPV2"
    is_b = vt_norm == "bOPV"

    total = num(df["TotalNbrVaccPolio"])
    pop = num(df["PopPolio"])
    have_block = all(b in header for b in ["Tot_Vacc_nOPV2", "Total_Vacc_bOPV"])
    ex_n = num(df["Tot_Vacc_nOPV2"]) if have_block else pd.Series(np.nan, index=df.index)
    ex_b = num(df["Total_Vacc_bOPV"]) if have_block else pd.Series(np.nan, index=df.index)

    key_str = pd.Series(["|".join(str(v) for v in t) for t in df[KEY].itertuples(index=False)], index=df.index)
    cand = (is_n | is_b) & resp_both & ~vt_both & ~pre_drop
    # a district that has an nOPV2 record AND a bOPV record for the same Response, year and round is
    # co-administered even when the Response name does not mention both vaccines (DRC, South Sudan)
    cand_any = (is_n | is_b) & ~vt_both & ~pre_drop
    df["_key"] = key_str
    occ = pd.Series(0, index=df.index)
    occ[cand_any] = df[cand_any].groupby(["_key", vt_norm[cand_any]]).cumcount()
    pk = key_str + "#" + occ.astype(str)

    n_set = set(pk[cand_any & is_n]); b_set = set(pk[cand_any & is_b])
    paired_keys = n_set & b_set
    paired = cand_any & pk.isin(paired_keys)
    # response+round groups that show the "separate records" practice
    grp = pd.Series(["|".join(str(v) for v in t) for t in df[["Response", "Year", "Round_Add"]].itertuples(index=False)], index=df.index)
    # (on a re-run the pairs are already merged: a combined record with two different
    #  vaccine totals is the trace of an earlier merge)
    prev_merged = vt_both & ex_n.notna() & ex_b.notna() & (ex_n != ex_b)
    sep_groups = set(grp[paired]) | set(grp[prev_merged])
    unpaired = cand & ~paired
    kept_single = unpaired & grp.isin(sep_groups)
    resp_single = unpaired & ~grp.isin(sep_groups)

    approach = pd.Series(A_NONE, index=df.index, dtype=object)
    approach[vt_both] = A_COMBINED
    approach[paired] = A_MERGED
    approach[resp_single] = A_RESP
    approach[kept_single] = A_KEPT

    # ------------------------------------------------ per-row output values
    out_tn = pd.Series(np.nan, index=df.index)
    out_tb = pd.Series(np.nan, index=df.index)
    out_total = total.copy()
    out_vt = vt_norm.copy()
    drop = pd.Series(False, index=df.index)
    drop |= pre_drop

    single = approach.isin([A_NONE, A_KEPT])
    out_tn[single & is_n] = total[single & is_n]
    out_tb[single & is_b] = total[single & is_b]

    comb = approach.isin([A_COMBINED, A_RESP])
    out_tn[comb] = ex_n[comb].where(ex_n[comb].notna(), total[comb])
    out_tb[comb] = ex_b[comb].where(ex_b[comb].notna(), total[comb])
    out_total[comb] = pd.concat([out_tn[comb], out_tb[comb]], axis=1).max(axis=1)
    # a combined record with no vaccinated total yet (target only) keeps its total empty, not 0
    nodata = comb & total.isna() & (out_tn.fillna(0) == 0) & (out_tb.fillna(0) == 0)
    out_total[nodata] = np.nan
    out_vt[approach == A_RESP] = COMBINED

    merged_log = []
    pdf = df[paired].copy()
    pdf["_pk"] = pk[paired]
    for pkey, g in pdf.groupby("_pk", sort=False):
        gn = g[is_n[g.index]]
        gb = g[is_b[g.index]]
        i_n, i_b = gn.index[0], gb.index[0]
        t_n, t_b = total[i_n], total[i_b]
        base = i_n if (np.nan_to_num(t_n) >= np.nan_to_num(t_b)) else i_b
        other = i_b if base == i_n else i_n
        out_tn[base] = t_n
        out_tb[base] = t_b
        out_total[base] = np.nanmax([t_n, t_b]) if not (np.isnan(t_n) and np.isnan(t_b)) else np.nan
        out_vt[base] = COMBINED
        drop[other] = True
        merged_log.append((df.at[base, "_row"], df.at[other, "_row"]))

    # optional correction of clearly wrong target populations (see --fix-poppolio)
    pop_fixed = pd.Series(False, index=df.index)
    pop_reason = pd.Series("", index=df.index, dtype=object)
    pop_new = pop.copy()
    if a.fix_poppolio:
        cvp = cvp_all
        with np.errstate(divide="ignore", invalid="ignore"):
            cv_now = out_total / pop.where(pop > 0)
            implied = (out_total / cvp.where(cvp > 0)).round()
        cond1 = (cv_now > CV_HIGH) & (cvp > 0) & (cvp <= CV_HIGH)
        # extreme coverage while CVPolio is less than half of it: the stored CVPolio used a much better population
        cond2 = (cv_now > CV_EXTREME) & (cvp > 0) & (cvp < cv_now / 2)
        pop_fixed = (cond1 | cond2) & implied.notna() & (implied > 0) & ~drop
        pop_new[pop_fixed] = implied[pop_fixed]
        pop_reason[pop_fixed] = "Population implied by CVPolio (vaccinated / CVPolio)"
        # coverage far too LOW (below CV_LOW) while the stored CVPolio is plausible (CV_LOW - CV_HIGH): the
        # stored PopPolio is too large (often the total population, or a province figure repeated in every
        # district), so it is replaced by vaccinated / CVPolio as well
        polio_early = out_vt.map(lambda x: isinstance(x, str) and "opv" in x.lower())
        cond3 = ((cv_now > 0) & (cv_now < CV_LOW) & (cvp >= CV_LOW) & (cvp <= CV_HIGH)
                 & polio_early & implied.notna() & (implied > 0) & ~drop & ~pop_fixed)
        pop_new[cond3] = implied[cond3]
        pop_reason[cond3] = "Population implied by CVPolio (coverage was below %.1f)" % CV_LOW
        pop_fixed = pop_fixed | cond3
        # populations are counts of children: whole numbers only
        frac = pop_new.notna() & ((pop_new % 1) != 0)
        pop_new[frac] = np.floor(pop_new[frac] + 0.5)
        pop_reason[frac & ~pop_fixed] = "Decimal population rounded to a whole number"
        pop_fixed = pop_fixed | frac
        # coverage still above CV_HIGH: more children vaccinated than the target population, so the
        # population is too low. Use the median PopPolio of the SAME district's other polio records in
        # the same year, only when that gives a plausible coverage (0.5 - CV_HIGH).
        low = out_vt.map(lambda x: x.lower() if isinstance(x, str) else "")
        polio_row = low.str.contains("opv")
        with np.errstate(divide="ignore", invalid="ignore"):
            cv2 = out_total / pop_new.where(pop_new > 0)
        todo = (cv2 > CV_HIGH) & polio_row & ~drop & (out_total > 0)
        dkey = pd.Series(["|".join(str(v) for v in t) for t in df[["Country", "Province", "District", "Year"]].itertuples(index=False)], index=df.index)
        pop_base = pop_new.copy()  # snapshot: medians never depend on the order of the corrections
        groups = {}
        for i in df.index[polio_row & ~drop & (pop_base > 0)]:
            groups.setdefault(dkey.iat[i], []).append(i)
        n_med = 0
        for i in df.index[todo]:
            others = [pop_base.iat[j] for j in groups.get(dkey.iat[i], []) if j != i]
            if not others:
                continue
            med = float(np.floor(np.median(others) + 0.5))
            if med > 0 and 0.5 <= out_total.iat[i] / med <= CV_HIGH:
                pop_new.iat[i] = med
                pop_reason.iat[i] = "District median of the same year (coverage was above %.1f)" % CV_HIGH
                pop_fixed.iat[i] = True
                n_med += 1
        print("PopPolio corrected on", int(pop_fixed.sum()), "records (decimals:", int(frac.sum()),
              "| district median:", n_med, ")", flush=True)
    model_info = None
    if a.model_poppolio:
        if not a.fix_poppolio:
            sys.exit("--model-poppolio needs --fix-poppolio")
        model_info = model_poppolio(df, out_total, out_vt, pop_new, pop_reason, pop_fixed, drop)
    pop_before = pop
    pop = pop_new
    # an empty CVPolio is filled with vaccinated / target population (TotalNbrVaccPolio / PopPolio); a record
    # without a vaccinated total (target only) keeps an empty coverage
    cvp_fill = pd.Series(False, index=df.index)
    cvp_val = pd.Series(np.nan, index=df.index)
    if "CVPolio" in df.columns and not a.no_fill_cvpolio:
        with np.errstate(divide="ignore", invalid="ignore"):
            cvp_val = (out_total / pop.where(pop > 0)).round(6)
        cvp_fill = cvp_all.isna() & cvp_val.notna() & ~drop
        if cvp_fill.any():
            print("empty CVPolio filled with TotalNbrVaccPolio / PopPolio: %d records (%d target-only records stay empty)" %
                  (int(cvp_fill.sum()), int((cvp_all.isna() & ~cvp_fill & ~drop).sum())), flush=True)
            for i in df.index[cvp_fill]:
                prov_log.append({"Excel row": int(df.at[i, "_row"]), "Country": df.at[i, "Country"], "District": df.at[i, "District"],
                                 "Column": "CVPolio", "Found": "", "Restored": float(cvp_val[i]),
                                 "Source": "TotalNbrVaccPolio / PopPolio (CVPolio was empty)"})

    with np.errstate(divide="ignore", invalid="ignore"):
        pop_ok = pop.where(pop > 0)
        out_cvn = out_tn / pop_ok
        out_cvb = out_tb / pop_ok
    # no vaccinated total yet (target only): no coverage either
    no_total = out_total.isna() & (out_tn.fillna(0) == 0) & (out_tb.fillna(0) == 0)
    out_cvn[no_total] = np.nan
    out_cvb[no_total] = np.nan

    keep = ~drop
    print("pairs merged:", int(drop.sum() - pre_drop.sum()), "| copied rounds removed:", int(copy_drop.sum()),
          "| duplicates removed:", int(dup_drop.sum()), "-> rows out:", int(keep.sum()), flush=True)

    # -------------------------------------------------------- new sheet build
    base_cols = [h for h in header if h not in BLOCK]
    new_header = []
    for h in base_cols:
        if h == BEFORE:
            new_header.extend(BLOCK)
        new_header.append(h)
    if BEFORE not in base_cols:
        sys.exit("Response column not found")
    # CVPolio (as stored) sits right after TotalNbrVaccPolio, so the vaccinated total and its coverage are together
    if "CVPolio" in new_header:
        new_header.remove("CVPolio")
        new_header.insert(new_header.index("TotalNbrVaccPolio") + 1, "CVPolio")
    new_letter = {h: get_column_letter(i + 1) for i, h in enumerate(new_header)}
    ncol_new = len(new_header)

    ws_new = wb.create_sheet("_new")
    tot_idx = header.index("TotalNbrVaccPolio")
    vt_idx = header.index("Vaccine_type")
    style_total = ws_new  # placeholder to keep linters quiet
    # header
    for j, h in enumerate(new_header, start=1):
        if h in header:
            src = ws.cell(row=1, column=header.index(h) + 1)
        else:
            src = ws.cell(row=1, column=tot_idx + 1)
        c = ws_new.cell(row=1, column=j, value=h)
        c._style = copy(src._style)
    # widths
    for h in new_header:
        j = new_header.index(h) + 1
        if h in header:
            w = ws.column_dimensions[old_letter[h]].width
        else:
            w = 16
        if w:
            ws_new.column_dimensions[get_column_letter(j)].width = w
    ws_new.freeze_panes = ws.freeze_panes

    # a vaccine that was not used has an empty total (not 0)
    vals = {
        "Tot_Vacc_nOPV2": out_tn, "CV_nOPV2": out_cvn,
        "Total_Vacc_bOPV": out_tb, "CV_bOPV": out_cvb,
        "Vaccine_type": out_vt, "TotalNbrVaccPolio": out_total,
    }
    total_style_src = tot_idx  # column index in original row tuples
    r_out = 2
    idxs = list(df.index[keep])
    pos_of = {name: new_header.index(name) for name in new_header}
    for i in idxs:
        row_cells = cells[i]
        for j, h in enumerate(new_header):
            if h in vals:
                v = vals[h].iat[i] if hasattr(vals[h], "iat") else None
                if isinstance(v, (float, np.floating)):
                    if np.isnan(v):
                        v = None
                    elif float(v).is_integer() and not h.startswith("CV_"):
                        v = int(v)
                if v is None:
                    continue
                c = ws_new.cell(row=r_out, column=j + 1, value=v)
                if h in header:
                    c._style = copy(row_cells[header.index(h)]._style)
                else:
                    c._style = copy(row_cells[total_style_src]._style)
                    if h.startswith("CV_"):
                        c.number_format = "0.00"
            else:
                oc = row_cells[header.index(h)]
                if h == "check_dp":
                    # duplicate-check key kept as a live formula on the new column letters:
                    # Country & Province & District & SIA_date & TotalNbrVaccPolio & Response
                    parts = [new_letter[x] + str(r_out) for x in
                             ("Country", "Province", "District", "SIA_date", "TotalNbrVaccPolio", "Response")]
                    c = ws_new.cell(row=r_out, column=j + 1, value="=" + "&".join(parts))
                    c._style = copy(oc._style)
                    continue
                if h == "CVPolio" and bool(cvp_fill.iat[i]):
                    c = ws_new.cell(row=r_out, column=j + 1, value=float(cvp_val.iat[i]))
                    c._style = copy(oc._style)
                    if c.number_format == "General":
                        c.number_format = "0.00"
                    continue
                if h == "PopPolio" and bool(pop_fixed.iat[i]):
                    c = ws_new.cell(row=r_out, column=j + 1, value=int(pop.iat[i]))
                    c._style = copy(oc._style)
                    continue
                val = repl.get((i, h), oc.value)
                if val is None:
                    continue
                c = ws_new.cell(row=r_out, column=j + 1, value=val)
                c._style = copy(oc._style)
        r_out += 1
    last_row = r_out - 1
    ws_new.auto_filter.ref = "A1:%s%d" % (get_column_letter(ncol_new), last_row)

    # --------------------------------------------- Children Reached formulas
    if SUMMARY_SHEET in wb.sheetnames:
        s2 = wb[SUMMARY_SHEET]
        pat = re.compile(r"Sheet1!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)")
        rev_old = {v: k for k, v in old_letter.items()}

        def repl(m):
            c1, c2 = m.group(1), m.group(3)
            n1 = new_letter[rev_old[c1]] if c1 in rev_old else c1
            n2 = new_letter[rev_old[c2]] if c2 in rev_old else c2
            return "Sheet1!$%s$%s:$%s$%d" % (n1, m.group(2), n2, last_row)

        for row in s2.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = pat.sub(repl, c.value)
                elif isinstance(c.value, str) and c.value.startswith("Method:") and "co-administered" not in c.value:
                    c.value = (c.value + " Co-administered records (Vaccine_type 'nOPV2 & bOPV') carry "
                               "TotalNbrVaccPolio = the maximum of Tot_Vacc_nOPV2 and Total_Vacc_bOPV, "
                               "so each child reached is counted once.")
                elif c.value == 42822 or (c.column == 2 and c.row == 4 and isinstance(c.value, (int, float))):
                    c.value = last_row - 1

    # swap sheets
    wb.remove(ws)
    ws_new.title = SHEET
    wb.move_sheet(ws_new, offset=-(len(wb.sheetnames) - 1))
    wb.active = 0
    print("saving", a.dst, flush=True)
    wb.save(a.dst)

    # -------------------------------------------------------------- feedback
    if a.feedback:
        write_feedback(a.feedback, df, keep, out_vt, out_total, pop, pop_before, pop_fixed, pop_reason,
                       out_cvn, out_cvb, cvp_all, approach, out_tn, out_tb)

    # ----------------------------------------------------------------- audit
    lk_report = lookup_crosscheck(df, lk) if lk is not None else None
    if a.audit:
        write_audit(a.audit, df, approach, vt_orig, out_vt, vt_changed, total, out_total, out_tn, out_tb,
                    drop, merged_log, keep, paired, pop_fixed, pop_before, pop, cvp_all, pop_reason, out_cvn, out_cvb, model_info, repairs, lk_report, copy_log, pre_label)

    # ------------------------------------------------------------- self-check
    chk = pd.read_excel(a.dst, sheet_name=SHEET)
    assert len(chk) == int(keep.sum()), "row count"
    assert list(chk.columns[:0]) == []
    assert chk.columns.tolist() == new_header, "column order"
    pos = chk.columns.get_loc(BEFORE)
    if "CVPolio" in chk.columns:
        assert chk.columns[pos - 1] == "CVPolio" and chk.columns[pos - 7:pos - 1].tolist() == BLOCK, \
            "block and CVPolio must sit right before Response"
    else:
        assert chk.columns[pos - 6:pos].tolist() == BLOCK, "block must sit right before Response"
    co = chk["Vaccine_type"].astype(str).str.lower().str.contains("nopv") & chk["Vaccine_type"].astype(str).str.lower().str.contains("bopv")
    mx = chk[["Tot_Vacc_nOPV2", "Total_Vacc_bOPV"]].max(axis=1)
    assert (chk.loc[co, "TotalNbrVaccPolio"].fillna(0) == mx[co].fillna(0)).all(), "co-admin total must be the max"
    print("self-check OK: %d rows, co-administered records: %d" % (len(chk), int(co.sum())))
    print(approach.value_counts().to_string())

    # ------------------------------------------------------------ data dictionary
    if a.dictionary:
        build_dictionary(chk, a.dictionary)

    # ------------------------------------------------------------ quality checks (merged from the portal)
    qres = None
    if not a.no_quality:
        # read again without turning the text "NA" into an empty cell (a province called NA is not an empty province)
        qdf = pd.read_excel(a.dst, sheet_name=SHEET, keep_default_na=False, na_values=[""])
        ctx_ = None
        if repo_ctx is not None:
            ctx_ = repo_ctx.rename(columns={"SIA-date": "SIA_date"})
        qres = qc_run(qdf, geo, qc_config(cv_ceiling=a.cv_ceiling / 100.0, pop_factor=a.pop_factor), context=ctx_)
        qres["fixes"] = pd.DataFrame(prov_log)
        qres["dup_log"], qres["dup_groups"], qres["existing"] = dup_log, dup_groups, existing_rows
        qc_print(qres)
        qc_write(qres, a.quality)

    print("\nDone. Files written under", a.out_dir)
    print("  upload_to_sharepoint/ : AFRO_admin_data.xlsx" + (" + the data dictionary" if a.dictionary else ""))
    print("  review/               : audit" + (" + feedback workbook" if a.feedback else ""))
    unresolved = [r for r in repairs if not r["Restored"]]
    inferred = [r for r in repairs if r["Restored"] and str(r.get("Source", "")).startswith("INFERRED")]
    if inferred:
        print("  %d values were inferred from the rest of the file (backups had no right value): sheet Repairs, "
              "Source = INFERRED. Confirm them with the countries." % len(inferred))
    if unresolved:
        print("  %d copy errors could not be fixed: see sheet Repairs in the audit workbook "
              "(Source = UNRESOLVED), to correct with the country." % len(unresolved))
    if qres is not None:
        print("  review/               : quality report (%d records with an Error, %d with a Warning)" % (qres["n_err"], qres["n_warn"]))
    if partial:
        print("WARNING: only %d countries in this file: do NOT upload the output as AFRO_admin_data.xlsx." % n_countries)
    else:
        print("Next: open AFRO_admin_data.xlsx in Excel and save once, then upload it to SharePoint.")
    if a.fail_on_error and qres is not None and qres["n_err"]:
        sys.exit(2)


# =====================================================================================================
# Quality checks (merged from the admin portal: admin_quality_checks.R)
# =====================================================================================================
# The same checks as the portal, run on the finished file (and, for the province / country check, also before
# the update so that the province of another country can be corrected). Severity:
#   Error    must be fixed before the record is used (wrong geography, missing key field, impossible date or number)
#   Warning  needs a human look: may be a real problem or a genuine exception
#   Info     good practice (combined co-administration record, spelling ...)
# Besides the portal's checks there is one for the new co-administration columns (the combined total must be the
# larger of the two vaccines). Only the province of another country is corrected by the script, with the
# geography reference; everything else is reported in AFRO_admin_data_quality_report.xlsx.
QC_LABELS = {
    "missing_key": "Missing key field", "country_name": "Country name not standard",
    "geo_province": "Province does not belong to the country",
    "geo_unknown": "Province not found in the geography reference", "sia_date": "SIA date missing or implausible",
    "report_date": "Report date before the SIA start / in the future", "negative_number": "Negative number",
    "target_missing": "Target population missing or zero", "target_gt_total": "Target population larger than total population",
    "coverage_ceiling": "Coverage above the ceiling", "coverage_high": "Coverage high (verify)",
    "coverage_mismatch": "Reported coverage differs from vaccinated / target",
    "target_jump": "Target population very different from other rounds",
    "coadmin_mixed": "Co-administration reported both combined and separate",
    "coadmin_combined": "Combined co-administration record", "coadmin_identical": "Identical figures for both vaccines",
    "coadmin_total": "Combined total is not the larger of the two vaccines",
    "vaccine_spelling": "Vaccine type spelling", "duplicate_key": "Duplicate record",
}
QC_SEV = ["Error", "Warning", "Info"]
QC_OWN_USE_MIN = 5   # a province name used by this many records of the country itself is not a stray copy
QC_COUNTRY_ALIAS = {
    "CAMEROUN": "CAMEROON", "REPUBLIQUE CENTRAFRIQUE": "CENTRAL AFRICAN REPUBLIC",
    "REPUBLIQUE CENTRAFRICAINE": "CENTRAL AFRICAN REPUBLIC", "TANZANIA": "UNITED REPUBLIC OF TANZANIA",
    "TCHAD": "CHAD", "THE GAMBIA": "GAMBIA", "SWAZILAND": "ESWATINI", "CABO VERDE": "CABO VERDE",
    "CAPE VERDE": "CABO VERDE", "COTE DIVOIRE": "COTE D IVOIRE", "IVORY COAST": "COTE D IVOIRE",
    "DRC": "DEMOCRATIC REPUBLIC OF THE CONGO", "CONGO DEMOCRATIC REPUBLIC": "DEMOCRATIC REPUBLIC OF THE CONGO",
    "REPUBLIC OF CONGO": "CONGO", "REPUBLIC OF THE CONGO": "CONGO", "GUINEA BISSAU": "GUINEA BISSAU",
    "GUINEE BISSAU": "GUINEA BISSAU", "GUINEE": "GUINEA", "GUINEE EQUATORIALE": "EQUATORIAL GUINEA",
}
QC_VACCINE_CANON = ["nOPV2", "bOPV", "mOPV2", "nOPV2 & bOPV", "nOPV2_bOPV_IPV", "nOPV1", "nOPV3", "IPV", "fIPV",
                    "MR", "Measles", "YF", "OPV"]


def qc_config(cv_ceiling=1.5, cv_high=1.2, cv_recalc_tol=0.02, pop_factor=3.0, pop_min_rounds=3,
              report_late_days=365, min_sia_year=2015):
    return dict(cv_ceiling=cv_ceiling, cv_high=cv_high, cv_recalc_tol=cv_recalc_tol, pop_factor=pop_factor,
                pop_min_rounds=pop_min_rounds, report_late_days=report_late_days, min_sia_year=min_sia_year)


_QN_CACHE = {}


def _qn(x):
    """Name for matching: no accents, capitals, letters and digits separated by single spaces."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    k = x if isinstance(x, str) else str(x)
    r = _QN_CACHE.get(k)
    if r is None:
        import unicodedata
        t = unicodedata.normalize("NFKD", k).encode("ascii", "ignore").decode().upper()
        r = re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", t)).strip()
        _QN_CACHE[k] = r
    return r


def _qcountry(x):
    k = _qn(x)
    return QC_COUNTRY_ALIAS.get(k, k)


def _qblank(s):
    return s.isna() | (s.astype(str).str.strip() == "")


def _qdate(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_datetime(s, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(s, errors="coerce")


# ---------------------------------------------------------------------------------------------------------
# Duplicates and text dates (merged from the portal: create_unique_id_dt, remove_duplicates_from_repo,
# find_all_duplicates, validate_admin_data, parse_sia_date)
# ---------------------------------------------------------------------------------------------------------
PORTAL_KEY_COLS = ["Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response"]
FR_MONTHS = [(r"janv\.?", "jan"), (r"f[eé]vr\.?|f[eé]v\.?", "feb"), (r"mars", "mar"), (r"avr\.?", "apr"), (r"mai", "may"),
             (r"juin", "jun"), (r"juil\.?", "jul"), (r"ao[uû]t", "aug"), (r"sept\.?", "sep"), (r"oct\.?", "oct"),
             (r"nov\.?", "nov"), (r"d[eé]c\.?", "dec")]


def parse_sia_text(x):
    """A SIA date written as text ("2020-09-18", "2020-09-18 00:00:00 UTC", "18-sept-20", "18/09/2020") -> Timestamp or NaT."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return pd.NaT
    if hasattr(x, "year"):
        return pd.Timestamp(x)
    s = str(x).strip().lower()
    if s in ("", "na", "nan"):
        return pd.NaT
    s = re.sub(r"\s+utc$", "", s)
    for pat, rep in FR_MONTHS:
        s = re.sub(pat, rep, s)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return pd.Timestamp(pd.to_datetime(s, format=fmt))
        except Exception:
            continue
    return pd.NaT


def _pk_clean(s):
    v = s.where(s.notna(), "NA").astype(str).str.strip().replace("", "NA").str.upper()
    v = v.str.replace(r"\s+", "_", regex=True).str.replace(r"[^A-Za-z0-9_\-]", "", regex=True)
    return v.replace("", "EMPTY")


def portal_key(df, with_admin3=False):
    """The portal's duplicate key (KEY_ID): Country | Province | District | SIA_date | Round_Add | Vaccine_type | Response,
    each cleaned the same way. with_admin3 adds Admin_3 (sub-districts of the same district are not duplicates)."""
    cols = list(PORTAL_KEY_COLS)
    if with_admin3 and "Admin_3" in df.columns:
        cols.insert(3, "Admin_3")
    parts = []
    for c in cols:
        if c not in df.columns:
            continue
        if c == "SIA_date":
            dv = df[c].map(parse_sia_text)
            parts.append(dv.dt.strftime("%Y-%m-%d").fillna("NA"))
        else:
            parts.append(_pk_clean(df[c]))
    return parts[0].str.cat(parts[1:], sep="|") if parts else pd.Series("", index=df.index)


def find_duplicates(df, resolve="none"):
    """Duplicate check and fix of the portal.
    - exact duplicates (same key AND every other value the same): the extra copies are removed (the first is kept);
    - conflicting duplicates (same key, other values differ): listed with a suggestion (the most complete record, then
      the first one) and removed only with resolve='complete' or 'first';
    - records with an empty key field are not treated as duplicates (they are reported as missing key fields).
    Returns (mask of rows to remove, log of the removed exact copies, table of the duplicate groups)."""
    key = portal_key(df, with_admin3=True)
    need = [c for c in ("Country", "Province", "District", "Round_Add", "Vaccine_type", "Response") if c in df.columns]
    blank = pd.Series(False, index=df.index)
    for c in need:
        blank |= df[c].isna() | (df[c].astype(str).str.strip() == "")
    blank |= df["SIA_date"].map(parse_sia_text).isna()
    cols = []
    for c in df.columns:
        if str(c).startswith("_") or c in ("check_dp", "prov_candidates"):
            continue
        first = df[c].dropna()
        if len(first) and isinstance(first.iloc[0], str) and first.iloc[0].startswith("="):
            continue
        cols.append(c)
    sig = pd.util.hash_pandas_object(df[cols].astype(str), index=False)
    d = pd.DataFrame({"key": key, "sig": sig, "blank": blank}, index=df.index)
    d["n"] = d["key"].map(d["key"].value_counts())
    cand = (d["n"] > 1) & ~d["blank"]
    drop = pd.Series(False, index=df.index)
    log_rows, grp_rows = [], []
    if not cand.any():
        return drop, pd.DataFrame(), pd.DataFrame()
    sub = d[cand]
    extra = sub.groupby(["key", "sig"]).cumcount() > 0
    first_of = sub[~extra].drop_duplicates(["key", "sig"]).set_index(["key", "sig"]).index
    first_row = {ks: i for i, ks in zip(sub[~extra].index, zip(sub[~extra]["key"], sub[~extra]["sig"]))}
    for i in sub.index[extra]:
        drop[i] = True
        log_rows.append({"Excel_row": int(df.at[i, "_row"]), "Kept_row": int(df.at[first_row[(d.at[i, "key"], d.at[i, "sig"])], "_row"]),
                         "Country": df.at[i, "Country"], "Province": df.at[i, "Province"], "District": df.at[i, "District"],
                         "SIA_date": parse_sia_text(df.at[i, "SIA_date"]).strftime("%Y-%m-%d"),
                         "Round_Add": df.at[i, "Round_Add"], "Vaccine_type": df.at[i, "Vaccine_type"],
                         "Response": df.at[i, "Response"], "Reason": "exact duplicate: every value is the same as in row %d"
                                                                  % int(df.at[first_row[(d.at[i, "key"], d.at[i, "sig"])], "_row"])})
    rest = sub[~extra]
    rest = rest[rest["key"].map(rest["key"].value_counts()) > 1]
    filled = df[cols].notna() & (df[cols].astype(str) != "")
    score = filled.sum(axis=1)
    gid = 0
    for key_, g in rest.groupby("key", sort=False):
        gid += 1
        idx = list(g.index)
        keeper = max(idx, key=lambda i: (score[i], -idx.index(i))) if resolve != "first" else idx[0]
        suggested = max(idx, key=lambda i: (score[i], -idx.index(i)))
        differ = [c for c in cols if df.loc[idx, c].astype(str).nunique() > 1]
        for n_, i in enumerate(idx, 1):
            applied = resolve in ("complete", "first") and i != keeper
            if applied:
                drop[i] = True
            grp_rows.append({"duplicate_group": gid, "record_num": n_, "Excel_row": int(df.at[i, "_row"]),
                             "Country": df.at[i, "Country"], "Province": df.at[i, "Province"], "District": df.at[i, "District"],
                             "SIA_date": parse_sia_text(df.at[i, "SIA_date"]).strftime("%Y-%m-%d"),
                             "Round_Add": df.at[i, "Round_Add"], "Vaccine_type": df.at[i, "Vaccine_type"],
                             "Response": df.at[i, "Response"], "TotalNbrVaccPolio": df.at[i, "TotalNbrVaccPolio"],
                             "PopPolio": df.at[i, "PopPolio"], "Filled_cells": int(score[i]),
                             "Suggested": "keep" if i == suggested else "remove", "Applied": "removed" if applied else "",
                             "Columns_that_differ": ", ".join(differ[:12]) + (" ..." if len(differ) > 12 else "")})
    return drop, pd.DataFrame(log_rows), pd.DataFrame(grp_rows)


def existing_in_repository(df, repo_path):
    """Records of the file whose portal key is already in the repository (the portal's 'records already exist' check)."""
    xl = pd.ExcelFile(repo_path)
    rp = xl.parse(SHEET if SHEET in xl.sheet_names else xl.sheet_names[0], keep_default_na=False, na_values=[""])
    rp = rp.rename(columns={"SIA-date": "SIA_date"}) if "SIA_date" not in rp.columns else rp
    rk = set(portal_key(rp))
    nk = portal_key(df)
    return nk.isin(rk), rp


def _dist_core(z):
    z = str(z).upper()
    z = re.sub(r"\b(WORHO|WOREDA|WEREDA|TOWN|CITY|ADMIN|SPECIAL|URBAN|RURAL|REFUGEE|CAMP|ZONE)\b", " ", z)
    return re.sub(r"[^A-Z0-9]", "", z)


def _align_spelling(nm, existing):
    """Use the spelling the data already has for this district: the same name once WORHO / TOWN / CITY ... are
    ignored, or a name that is at least 90% alike (LAY ARMACHEHO -> LAY ARMACHO). `existing` = names used in the
    province by other records. Only a unique answer is used."""
    ex = [e for e in existing if e]
    if nm in ex:
        return nm, ""
    k = _dist_core(nm)
    same = sorted({e for e in ex if k and _dist_core(e) == k})
    if len(same) == 1:
        return same[0], same[0]
    if not same and k:
        import difflib
        close = sorted({e for e in ex if difflib.SequenceMatcher(None, k, _dist_core(e)).ratio() >= 0.90})
        if len(close) == 1:
            return close[0], close[0]
    return nm, ""


def fill_blank_district(df, src_path):
    """Fill an empty District from a recovered file.

    The source (xlsx or csv, first sheet) may use the template names (Ctry_Name, Prov_Name, Dist_Name, PopPolio,
    TotalNbrVacPolio, SIA Code, Round Number) or the admin names (Country, Province, District, TotalNbrVaccPolio,
    Response, Round_Add). Records with an empty District are grouped by country, Response, round and province, and
    matched one to one with the source records of the same group:
      1. the same PopPolio and the same vaccinated total (confidence high);
      2. the same PopPolio and a vaccinated total within 10% (the total was edited later), only a unique answer;
      3. the same vaccinated total and a different PopPolio, only a unique answer (confidence high from 1,000 vaccinated);
      4. one record and one district of the file left unmatched in the same province (confidence medium).
    Each source record is used once. Records that do not match stay empty. The district is written in capitals,
    like the rest of the file. Returns {row index: {set, why, confidence, check}}."""
    out = {}
    blank = df["District"].isna() | (df["District"].astype(str).str.strip() == "")
    if not blank.any():
        return out
    if str(src_path).lower().endswith((".csv", ".txt")):
        src = pd.read_csv(src_path, dtype=str, keep_default_na=False, na_values=[""])
    else:
        src = pd.read_excel(src_path, dtype=str, keep_default_na=False, na_values=[""])

    def pick(*names):
        for n in names:
            if n in src.columns:
                return n
        return None
    cc, pc, dc = pick("Country", "Ctry_Name"), pick("Province", "Prov_Name"), pick("District", "Dist_Name")
    pp, tc = pick("PopPolio"), pick("TotalNbrVaccPolio", "TotalNbrVacPolio")
    rc, nc = pick("Response", "SIA Code"), pick("Round_Add", "Round Number")
    if not (pc and dc and pp and tc and rc):
        raise ValueError("the file needs Province/Prov_Name, District/Dist_Name, PopPolio, the vaccinated total and Response/SIA Code")
    S = pd.DataFrame({"country": src[cc].map(_lkc) if cc else "", "prov": src[pc].map(_pk_prov), "dist": src[dc],
                      "pop": pd.to_numeric(src[pp], errors="coerce"), "tot": pd.to_numeric(src[tc], errors="coerce"),
                      "resp": src[rc].map(_rk), "rnd": src[nc].map(_s) if nc else ""})
    S = S[S["dist"].notna() & S["resp"].astype(bool)]
    pop = pd.to_numeric(df["PopPolio"], errors="coerce")
    tot = pd.to_numeric(df["TotalNbrVaccPolio"], errors="coerce")
    ck, rk_, rd, pk = df["Country"].map(_lkc), df["Response"].map(_rk), df["Round_Add"].map(_s), df["Province"].map(_pk_prov)
    named = df[~blank]
    used_names = {}
    for (c_, p_), grp in named.groupby([ck[~blank], pk[~blank]]):
        used_names[(c_, p_)] = sorted(set(grp["District"].astype(str).str.strip()))
    for (c, r_, n_, p_), g in df[blank].groupby([ck[blank], rk_[blank], rd[blank], pk[blank]], dropna=False):
        cand = S[(S["resp"] == r_) & (S["prov"] == p_) & ((S["country"] == c) | (S["country"] == ""))]
        if nc:
            cand = cand[(cand["rnd"] == n_) | (cand["rnd"] == "")]
        if cand.empty:
            continue
        free_c, free_g = set(cand.index), set(g.index)
        # stage 1: exact on PopPolio and total, only when the pair is unique on both sides
        for stage in (1, 2, 3):
            keyfn = {1: lambda p, t: (p, t), 2: lambda p, t: (p,), 3: lambda p, t: (t,)}[stage]
            for i in sorted(free_g):
                if pd.isna(pop[i]) and stage != 3:
                    continue
                if stage == 3 and pd.isna(tot[i]):
                    continue
                if stage == 1:
                    m = [j for j in free_c if cand.at[j, "pop"] == pop[i] and cand.at[j, "tot"] == tot[i]]
                    twin = [k for k in free_g if pop[k] == pop[i] and tot[k] == tot[i]]
                elif stage == 2:
                    m = [j for j in free_c if cand.at[j, "pop"] == pop[i] and
                         (pd.notna(tot[i]) and pd.notna(cand.at[j, "tot"]) and
                          abs(cand.at[j, "tot"] - tot[i]) <= 0.10 * max(tot[i], 1))]
                    twin = [k for k in free_g if pop[k] == pop[i]]
                else:
                    m = [j for j in free_c if cand.at[j, "tot"] == tot[i]]
                    twin = [k for k in free_g if tot[k] == tot[i]]
                if len(m) != 1 or len(twin) != 1:
                    continue
                j = m[0]
                nm = re.sub(r"\s+", " ", str(cand.at[j, "dist"]).strip()).upper()
                why = {1: "same province, PopPolio and vaccinated total",
                       2: "same province and PopPolio; vaccinated total %d here, %d in the file" % (tot[i], cand.at[j, "tot"]),
                       3: "same province and vaccinated total; PopPolio %s here, %s in the file" %
                          (int(pop[i]) if pd.notna(pop[i]) else "empty", int(cand.at[j, "pop"]) if pd.notna(cand.at[j, "pop"]) else "empty")}[stage]
                conf = "high" if (stage == 1 or (stage == 3 and tot[i] >= 1000)) else "medium"
                nm, al = _align_spelling(nm, used_names.get((c, p_), []))
                if al:
                    why += "; spelling as already used in the data (" + str(cand.at[j, "dist"]).strip() + " in the file)"
                out[i] = {"set": {"District": nm}, "why": why, "confidence": conf, "check": ""}
                free_c.discard(j)
                free_g.discard(i)
        # stage 3b: several districts of the file have this vaccinated total, but the target implied by the record's own
        # coverage (vaccinated / CVPolio) is within 2% of the target of only one of them
        cvn = pd.to_numeric(df["CVPolio"], errors="coerce") if "CVPolio" in df.columns else None
        for i in sorted(free_g):
            if cvn is None or pd.isna(tot[i]) or pd.isna(cvn[i]) or cvn[i] <= 0:
                continue
            imp = tot[i] / cvn[i]
            m = [j for j in free_c if cand.at[j, "tot"] == tot[i] and pd.notna(cand.at[j, "pop"]) and
                 abs(cand.at[j, "pop"] - imp) <= 0.02 * imp]
            twin = [k for k in free_g if tot[k] == tot[i]]
            if len(m) != 1 or len(twin) != 1:
                continue
            j = m[0]
            nm = re.sub(r"\s+", " ", str(cand.at[j, "dist"]).strip()).upper()
            nm, al = _align_spelling(nm, used_names.get((c, p_), []))
            why = ("same province and vaccinated total; PopPolio is %s here, but the target implied by the coverage "
                   "(%d / CVPolio = %d) is the file's target of this district (%d)" %
                   ("" if pd.isna(pop[i]) else int(pop[i]), tot[i], round(imp), cand.at[j, "pop"]))
            if al:
                why += "; spelling as already used in the data (" + str(cand.at[j, "dist"]).strip() + " in the file)"
            out[i] = {"set": {"District": nm}, "why": why, "confidence": "high" if tot[i] >= 1000 else "medium", "check": ""}
            free_c.discard(j)
            free_g.discard(i)
        # stage 4: one record and one source district left in the group: they belong together
        if len(free_g) == 1 and len(free_c) == 1:
            i, j = next(iter(free_g)), next(iter(free_c))
            nm = re.sub(r"\s+", " ", str(cand.at[j, "dist"]).strip()).upper()
            nm, _al = _align_spelling(nm, used_names.get((c, p_), []))
            out[i] = {"set": {"District": nm}, "confidence": "medium", "check": "",
                      "why": "the only record and the only district of the file left in this province (all others matched)"}
            free_c.discard(j)
            free_g.discard(i)
        # stage 5: records still unmatched and districts of the file still unused: a district is given to a record when
        # they share the PopPolio or the vaccinated total and the pairing is unique on both sides (confidence medium)
        for i in sorted(free_g):
            m = [j for j in free_c if (pd.notna(pop[i]) and cand.at[j, "pop"] == pop[i]) or
                 (pd.notna(tot[i]) and cand.at[j, "tot"] == tot[i])]
            if len(m) != 1:
                continue
            j = m[0]
            twin = [k for k in free_g if (pd.notna(pop[k]) and cand.at[j, "pop"] == pop[k]) or
                    (pd.notna(tot[k]) and cand.at[j, "tot"] == tot[k])]
            if len(twin) != 1:
                continue
            nm = re.sub(r"\s+", " ", str(cand.at[j, "dist"]).strip()).upper()
            nm, _al = _align_spelling(nm, used_names.get((c, p_), []))
            same = "PopPolio" if cand.at[j, "pop"] == pop[i] else "vaccinated total"
            out[i] = {"set": {"District": nm}, "confidence": "medium", "check": "",
                      "why": "same province and %s; the other figure was edited later (PopPolio %s / vaccinated %s here, %s / %s in the file); "
                             "the only district of the file left" % (same,
                             "" if pd.isna(pop[i]) else int(pop[i]), "" if pd.isna(tot[i]) else int(tot[i]),
                             "" if pd.isna(cand.at[j, "pop"]) else int(cand.at[j, "pop"]),
                             "" if pd.isna(cand.at[j, "tot"]) else int(cand.at[j, "tot"]))}
            free_c.discard(j)
            free_g.discard(i)
    return out


def rename_addis_districts(df, src_path):
    """Addis Ababa: the data holds the sub-city name on every row (NIFAS SILK LAFTO x13), the recovered file lists the
    woredas of each sub-city. Records that look like duplicates are renamed "<SUB-CITY> WOREDA NN".

    Within a campaign, round and sub-city, every record is matched to the file by PopPolio + vaccinated total, or by the
    vaccinated total alone, only when the answer is unique on both sides. A sub-city is renamed only when ALL its records
    are matched to different woredas (all or nothing), and the new names must not already be used. Sub-cities or rounds
    the file does not cover are left unchanged. Returns {row index: {set, why, confidence, check, old}}."""
    import difflib
    out = {}
    if str(src_path).lower().endswith((".csv", ".txt")):
        src = pd.read_csv(src_path, dtype=str, keep_default_na=False, na_values=[""])
    else:
        src = pd.read_excel(src_path, dtype=str, keep_default_na=False, na_values=[""])
    need = ["Prov_Name", "Zone_Name", "Dist_Name", "PopPolio", "SIA Code"]
    tc = "TotalNbrVacPolio" if "TotalNbrVacPolio" in src.columns else ("TotalNbrVaccPolio" if "TotalNbrVaccPolio" in src.columns else None)
    if tc is None or any(n not in src.columns for n in need):
        return out
    cc = "Ctry_Name" if "Ctry_Name" in src.columns else None
    nc = "Round Number" if "Round Number" in src.columns else None

    def sub_key(x):
        k = re.sub(r"\b(SUB\s*CITY|SUBCITY|SC)\b", " ", str(x).upper())
        return re.sub(r"[^A-Z]", "", k)

    def woreda_no(x):
        m = re.findall(r"\d+", str(x))
        return int(m[-1]) if m else None
    S = pd.DataFrame({"country": src[cc].map(_lkc) if cc else "", "prov": src["Prov_Name"].map(_pk_prov),
                      "sub": src["Zone_Name"].map(sub_key), "no": src["Dist_Name"].map(woreda_no),
                      "pop": pd.to_numeric(src["PopPolio"], errors="coerce"), "tot": pd.to_numeric(src[tc], errors="coerce"),
                      "resp": src["SIA Code"].map(_rk), "rnd": src[nc].map(_s) if nc else ""})
    S = S[(S["prov"] == "ADDIS ABABA") & S["no"].notna() & S["sub"].astype(bool)]
    if S.empty:
        return out
    pop = pd.to_numeric(df["PopPolio"], errors="coerce")
    tot = pd.to_numeric(df["TotalNbrVaccPolio"], errors="coerce")
    ck, rk_, rd, pk = df["Country"].map(_lkc), df["Response"].map(_rk), df["Round_Add"].map(_s), df["Province"].map(_pk_prov)
    dist = df["District"].astype(str).str.strip()
    is_a = (pk == "ADDIS ABABA") & df["District"].notna() & (dist != "")

    def base_of(d):
        b = re.sub(r"[\s\-_]*(WOREDA|WEREDA|WORHO)?[\s\-_]*\d+$", "", str(d).upper().strip())
        return re.sub(r"\s+", " ", b).strip()
    base = dist.map(base_of)
    used_all = set(zip(ck[is_a], rk_[is_a], rd[is_a], dist[is_a].str.upper()))
    for (c, r_, n_), grp in df[is_a].groupby([ck[is_a], rk_[is_a], rd[is_a]]):
        cand = S[(S["resp"] == r_) & ((S["country"] == c) | (S["country"] == ""))]
        if nc:
            cand = cand[(cand["rnd"] == n_) | (cand["rnd"] == "")]
        if cand.empty:
            continue
        for b, gb in grp.groupby(base[grp.index]):
            sk = re.sub(r"[^A-Z]", "", b)
            keys = sorted(cand["sub"].unique())
            hit = [k for k in keys if k == sk or difflib.SequenceMatcher(None, k, sk).ratio() >= 0.88]
            if len(hit) != 1:
                continue
            cs = cand[cand["sub"] == hit[0]]
            free_c, got = set(cs.index), {}
            for i in gb.index:
                if pd.isna(tot[i]):
                    continue
                m = [j for j in free_c if cs.at[j, "tot"] == tot[i] and pd.notna(pop[i]) and cs.at[j, "pop"] == pop[i]]
                twin = [k for k in gb.index if pop[k] == pop[i] and tot[k] == tot[i]]
                how = "PopPolio and vaccinated total"
                if len(m) != 1 or len(twin) != 1:
                    m = [j for j in free_c if cs.at[j, "tot"] == tot[i]]
                    twin = [k for k in gb.index if tot[k] == tot[i]]
                    how = "vaccinated total"
                if len(m) == 1 and len(twin) == 1:
                    got[i] = (m[0], how)
            # all or nothing: every record matched, to different woredas
            if len(got) != len(gb) or len({j for j, _h in got.values()}) != len(got):
                continue
            news = {i: "%s WOREDA %02d" % (b, int(cs.at[j, "no"])) for i, (j, _h) in got.items()}
            if len(set(news.values())) != len(news):
                continue
            # a new name must not collide with a name used by another record of the same campaign and round
            if any((c, r_, nm) in used_all and (dist[i].upper() != nm) for i, nm in news.items()):
                continue
            for i, nm in news.items():
                j, how = got[i]
                if nm == dist[i].upper():
                    continue
                out[i] = {"set": {"District": nm}, "confidence": "high" if (how.startswith("PopPolio") or tot[i] >= 1000) else "medium",
                          "check": "", "old": df.at[i, "District"],
                          "why": "woreda %d of %s in the recovered file (same %s)" % (int(cs.at[j, "no"]), b, how)}
    return out


def _pk_prov(x):
    """Province key for matching a recovered file: capitals, no accents, no trailing CITY, a few spellings merged."""
    k = _qn(x)
    k = re.sub(r"\s+CITY$", "", k)
    return {"OROMIYA": "OROMIA", "GAMBELA": "GAMBELLA", "GAMBELLA": "GAMBELLA"}.get(k, k)


def align_response_names(df, lk, countries=None):
    """Rename a Response to the exact name the lookup uses for the same campaign.

    A (country, Response, SIA_date) group is renamed when the Response is not a name of the lookup, and the lookup has,
    for the same country, a round starting within 7 days of the date whose name starts with the Response (or is at least
    90% alike) - only when that is one single lookup name. Example: MLW-2026-02-1_nOPV -> MLW-2026-02-1_nOPV-NID.
    Only for the countries given. Returns {row index: {set, why, confidence, check}}."""
    import difflib
    out = {}
    names = set(lk["_rk"])
    ck = df["Country"].map(_lkc)
    rk_ = df["Response"].map(_rk)
    dk = df["SIA_date"].map(_dkey)
    ok = df["Response"].notna() & (rk_ != "") & ~rk_.isin(names) & (dk != "")
    if countries:
        ok &= ck.isin({_lkc(x) for x in countries})
    if not ok.any():
        return out
    for (c, r_, d_), g in df[ok].groupby([ck[ok], rk_[ok], dk[ok]]):
        e = lk[lk["_ck"] == c]
        if e.empty:
            continue
        try:
            days = (pd.to_datetime(e["_d"], errors="coerce") - pd.Timestamp(d_)).abs().dt.days
        except Exception:
            continue
        e = e[days <= 7]
        if e.empty:
            continue
        e = e[e["_rk"].map(lambda x: x.startswith(r_) or difflib.SequenceMatcher(None, r_, x).ratio() >= 0.90).astype(bool)]
        if e.empty:
            continue
        nm = sorted(set(e["OBR Name"].astype(str)))
        if len(nm) != 1:
            continue
        for i in g.index:
            out[i] = {"set": {"Response": nm[0]}, "confidence": "high", "check": "",
                      "why": "name used by the lookup for this round (%s, round starting %s)" %
                             (nm[0], str(e.iloc[0]["Round Start Date"])[:10])}
    return out


def fill_blank_response(df, lk):
    """Fill an empty Response from the campaign lookup (the portal reports it as a missing key field).

    For every group of records with the same country, SIA_date and round and an empty Response:
      - exact: the one lookup round of that country and vaccine that started on that date;
      - near: when there is none, the one round of that country and vaccine whose start date is within 21 days and
        whose round number is the group's round number (the rounds of the data are often dated a few days off);
    Only a unique answer is used. When the lookup name contains the name of a province of the country ("Kwara
    Response"), only the records of that province are filled. The spelling already used in the data is kept.
    Returns {row index: {set, why, confidence, check}}."""
    out = {}
    resp_blank = df["Response"].isna() | (df["Response"].astype(str).str.strip() == "")
    if not resp_blank.any():
        return out
    ck = df["Country"].map(_lkc)
    dts = df["SIA_date"].map(parse_sia_text)
    rd = df["Round_Add"].map(_s)
    spell = (df.loc[~resp_blank, "Response"].groupby(df.loc[~resp_blank, "Response"].map(_rk)).agg(lambda x: x.value_counts().index[0]).to_dict())
    pkey = df["Province"].map(_qn)
    b = df[resp_blank]
    for (c, d_, r_), g in b.groupby([ck[b.index], dts[b.index].dt.strftime("%Y-%m-%d"), rd[b.index]], dropna=False):
        if not isinstance(d_, str):
            continue
        lc = lk[lk["_ck"] == c]
        if lc.empty:
            continue
        vs = set().union(*[_vset(v) for v in g["Vaccine_type"]])
        if not vs:
            continue
        lc = lc.loc[[bool(v & vs) or not v for v in lc["_vs"]]]
        e = lc[lc["_d"] == d_]
        near = None
        if e.empty:
            days = (pd.to_datetime(lc["_d"], errors="coerce") - pd.Timestamp(d_)).abs().dt.days
            rn = _rnum(r_)
            e = lc[(days <= 21) & (lc["_rn"] == rn)] if rn is not None else lc.iloc[0:0]
            if len(e):
                near = int(days[e.index].min())
        rks = list(dict.fromkeys(e["_rk"]))
        if len(rks) != 1:
            continue
        x = e[e["_rk"] == rks[0]].iloc[0]
        nm = _qn(x["OBR Name"])
        hit = {p for p in set(pkey[g.index]) if p and re.search(r"\b" + re.escape(p) + r"\b", nm)}
        rows = [i for i in g.index if (not hit or pkey[i] in hit)]
        name = spell.get(rks[0]) or x["OBR Name"]
        why = "lookup: %s %s starts %s" % (x["OBR Name"], x["Round Number"], x["Round Start Date"])
        if near is not None:
            why += " (%d day(s) from the record's SIA_date, same round number)" % near
        if hit:
            why += "; only the records of " + ", ".join(sorted(hit))
        for i in rows:
            out[i] = {"set": {"Response": name}, "why": why, "confidence": "high" if near is None else "medium", "check": ""}
    return out


def find_geo(user_path, src):
    cands = [user_path] if user_path else []
    here = os.path.dirname(os.path.abspath(__file__))
    cands.append(os.path.join(here, "geo_reference.csv"))
    if os.environ.get("ADMIN_GEO"):
        cands.append(os.environ["ADMIN_GEO"])
    for base in (os.path.dirname(os.path.abspath(src)), os.getcwd(), here):
        for rel in ("geo_reference.csv", os.path.join("input", "geo_reference.csv"),
                    os.path.join("..", "data", "geo_reference.csv")):
            cands.append(os.path.join(base, rel))
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def load_geo(path):
    g = pd.read_excel(path) if path.lower().endswith((".xlsx", ".xls")) else pd.read_csv(path, encoding="utf-8")
    g.columns = [str(c).strip().lower() for c in g.columns]
    miss = [c for c in ("country", "province", "district") if c not in g.columns]
    if miss:
        raise ValueError("columns missing: " + ", ".join(miss))
    g = g[["country", "province", "district"]].dropna().drop_duplicates().copy()
    g["c_key"] = g["country"].map(_qcountry)
    g["p_key"] = g["province"].map(_qn)
    g["d_key"] = g["district"].map(_qn)
    return g.reset_index(drop=True)


def qc_prepare(df):
    def col(name):
        return df[name] if name in df.columns else pd.Series(np.nan, index=df.index)
    w = pd.DataFrame({"idx": np.arange(len(df))}, index=df.index)
    for k, name in (("Country", "Country"), ("Province", "Province"), ("District", "District"),
                    ("Round", "Round_Add"), ("Vaccine", "Vaccine_type"), ("Response", "Response")):
        w[k] = col(name).map(lambda x: x if isinstance(x, str) else ("" if pd.isna(x) else str(x)))
    w["sia_date"] = _qdate(col("SIA_date"))
    w["report_date"] = _qdate(col("Date_Report"))
    for k, name in (("tot_pop", "TotPop"), ("target", "PopPolio"), ("vacc", "TotalNbrVaccPolio"), ("cv_rep", "CVPolio"),
                    ("tn", "Tot_Vacc_nOPV2"), ("tb", "Total_Vacc_bOPV")):
        w[k] = pd.to_numeric(col(name), errors="coerce")
    w["age_grp"] = col("AgeGPolio").map(lambda x: "" if pd.isna(x) else str(x))
    w["c_key"] = w["Country"].map(_qcountry)
    w["p_key"] = w["Province"].map(_qn)
    w["d_key"] = w["District"].map(_qn)
    kn = w["Vaccine"].map(_qn)
    w["fam"] = np.where(kn.str.contains("NOPV2") & kn.str.contains("BOPV"), "combined",
                        np.where(kn.str.contains("NOPV2"), "nOPV2", np.where(kn.str.contains("BOPV"), "bOPV", "other")))
    w["key_id"] = (w["c_key"] + "|" + w["p_key"] + "|" + w["d_key"] + "|" + w["sia_date"].dt.strftime("%Y-%m-%d").fillna("")
                   + "|" + w["Round"].map(_qn) + "|" + kn + "|" + w["Response"].map(_qn))
    return w


def _qi(idx, check, sev, message, current=None, suggested=None):
    idx = np.asarray(idx, dtype=int)
    if len(idx) == 0:
        return None
    n = len(idx)

    def rep(v):
        if v is None:
            return [""] * n
        if isinstance(v, (list, tuple, np.ndarray, pd.Series)):
            return ["" if (x is None or (isinstance(x, float) and np.isnan(x))) else str(x) for x in v]
        return [str(v)] * n
    return pd.DataFrame({"idx": idx, "check": check, "severity": sev, "message": rep(message),
                         "current": rep(current), "suggested": rep(suggested)})


def _qfmt(x):
    return "{:,.0f}".format(x)


def qc_geo(w, geo):
    """Province / country consistency. Returns (issues, corrections, effective province)."""
    prov_eff = w["Province"].copy()
    if geo is None:
        return [], None, prov_eff
    cp = geo[["c_key", "p_key"]].drop_duplicates()
    okset = set(zip(cp["c_key"], cp["p_key"]))
    prov_countries = cp.groupby("p_key")["c_key"].agg(lambda x: sorted(set(x))).to_dict()
    ok = np.array([(c, p) in okset for c, p in zip(w["c_key"], w["p_key"])], dtype=bool)
    known = w["c_key"].isin(set(geo["c_key"])).values
    test = np.where(known & ~_qblank(w["Province"]).values & ~ok)[0]
    if len(test) == 0:
        return [], None, prov_eff
    cd = geo.drop_duplicates(["c_key", "d_key", "province"])
    grp = cd.groupby(["c_key", "d_key"])["province"].agg(["nunique", "first"])
    uniq = grp[grp["nunique"] == 1]["first"].to_dict()
    sugg = [uniq.get((w["c_key"].iat[i], w["d_key"].iat[i])) for i in test]
    src = ["geography reference (district)" if s else None for s in sugg]
    good = w[ok & ~_qblank(w["District"]).values]
    ev = {}
    if len(good):
        cnt = good.groupby(["c_key", "d_key", "Province"]).size().reset_index(name="n")
        for (c, d), g in cnt.groupby(["c_key", "d_key"]):
            tot, top = g["n"].sum(), g["n"].max()
            if top / tot >= 0.8:
                ev[(c, d)] = g.sort_values("n", ascending=False)["Province"].iloc[0]
    for k, i in enumerate(test):
        if sugg[k] is None:
            s = ev.get((w["c_key"].iat[i], w["d_key"].iat[i]))
            if s:
                sugg[k], src[k] = s, "other records of the same district"
    other = [", ".join(c for c in prov_countries.get(w["p_key"].iat[i], []) if c) for i in test]
    cross = np.array([bool(o) for o in other])
    # a name that the country itself uses in many records is a naming variant (Rwanda EAST / NORTH, Malawi NORTH),
    # not a stray copy from another country: reported as a Warning, never corrected by the script
    use = w.groupby(["c_key", "p_key"]).size().to_dict()
    n_use = np.array([use.get((w["c_key"].iat[i], w["p_key"].iat[i]), 0) for i in test])
    stray = cross & (n_use < QC_OWN_USE_MIN)
    variant = cross & ~stray
    issues = []
    for mask, sev, fmt in ((stray, "Error", "Province '%s' belongs to %s, not to %s"),
                           (variant, "Warning", "Province '%s' is not an Admin1 of %s in the reference but is used by many records of %s: "
                                                "another naming system or spelling? (reference says: %s)")):
        ci = test[mask]
        if sev == "Error":
            msgs = [fmt % (w["Province"].iat[i], o, w["Country"].iat[i]) for i, o in zip(ci, np.array(other)[mask])]
        else:
            msgs = [fmt % (w["Province"].iat[i], w["Country"].iat[i], w["Country"].iat[i], o)
                    for i, o in zip(ci, np.array(other)[mask])]
        issues.append(_qi(ci, "geo_province", sev, msgs, current=w["Province"].iloc[ci].values,
                          suggested=[sugg[k] for k in np.where(mask)[0]]))
    ui = test[~cross]
    issues.append(_qi(ui, "geo_unknown", "Info",
                      ["Province '%s' is not an Admin1 of %s in the geography reference (spelling, or a new Admin1?)"
                       % (w["Province"].iat[i], w["Country"].iat[i]) for i in ui],
                      current=w["Province"].iloc[ui].values, suggested=[sugg[k] for k in np.where(~cross)[0]]))
    corr = pd.DataFrame({"idx": test, "Country": w["Country"].iloc[test].values, "District": w["District"].iloc[test].values,
                         "current_province": w["Province"].iloc[test].values, "suggested_province": sugg,
                         "source": src, "certain": stray})
    corr = corr[corr["suggested_province"].notna()].copy()
    if len(corr):
        prov_eff = prov_eff.copy()
        prov_eff.iloc[corr["idx"].values] = corr["suggested_province"].values
    return [x for x in issues if x is not None], corr, prov_eff


def qc_run(df, geo=None, cfg=None, today=None, context=None):
    """Run every check on a data frame. Returns {issues, corrections, by_check, by_country, n_rows, n_flagged}."""
    cfg = cfg or qc_config()
    today = pd.Timestamp(today or pd.Timestamp.today().normalize())
    n_ctx = 0
    n_new = len(df)
    if context is not None and len(context):
        # rows of the repository (same countries) only give context to the target population check; they are not reported
        ctx = context[context["Country"].map(_qcountry).isin(set(df["Country"].map(_qcountry)))]
        n_ctx = len(ctx)
        df = pd.concat([ctx, df], ignore_index=True, sort=False)
    w = qc_prepare(df)
    parts = []

    def add(x):
        if x is None:
            return
        if isinstance(x, list):
            parts.extend(p for p in x if p is not None)
        else:
            parts.append(x)

    # --- missing key fields
    for nm, label in (("Country", "Country"), ("Province", "Province"), ("District", "District"),
                      ("Vaccine", "Vaccine_type"), ("Response", "Response"), ("Round", "Round_Add")):
        add(_qi(np.where(_qblank(w[nm]).values)[0], "missing_key", "Error", label + " is empty"))
    add(_qi(np.where(w["sia_date"].isna().values)[0], "sia_date", "Error", "SIA_date is empty or could not be read"))
    # --- country name
    if geo is not None:
        ref = set(geo["c_key"])
        std = geo.drop_duplicates("c_key").set_index("c_key")["country"].to_dict()
        i = np.where((~_qblank(w["Country"]) & ~w["c_key"].isin(ref)).values)[0]
        add(_qi(i, "country_name", "Error", "Country name is not in the geography reference", current=w["Country"].iloc[i].values))
        raw = w["Country"].map(_qn)
        v = np.where((~_qblank(w["Country"]) & raw.isin(set(QC_COUNTRY_ALIAS)) & (raw != w["c_key"]) & w["c_key"].isin(ref)).values)[0]
        standard = [std[w["c_key"].iat[k]] for k in v]
        add(_qi(v, "country_name", "Info",
                ["Country spelled '%s'; the standard name is '%s'" % (w["Country"].iat[k], s) for k, s in zip(v, standard)],
                current=w["Country"].iloc[v].values, suggested=standard))
    gi, corr, prov_eff = qc_geo(w, geo)
    add(gi)
    # --- dates
    yr = w["sia_date"].dt.year
    i = np.where((w["sia_date"].notna() & ((yr < cfg["min_sia_year"]) | (yr > today.year + 1))).values)[0]
    add(_qi(i, "sia_date", "Error", ["SIA_date year %d is outside the plausible range" % yr.iat[k] for k in i],
            current=w["sia_date"].dt.strftime("%Y-%m-%d").iloc[i].values))
    i = np.where((w["sia_date"].notna() & (w["sia_date"] > today + pd.Timedelta(days=120))).values)[0]
    add(_qi(i, "sia_date", "Warning", "SIA_date is more than 4 months in the future",
            current=w["sia_date"].dt.strftime("%Y-%m-%d").iloc[i].values))
    d = (w["report_date"] - w["sia_date"]).dt.days
    i = np.where((d < 0).fillna(False).values)[0]
    add(_qi(i, "report_date", "Warning", ["Report date is %d day(s) before the SIA start" % -d.iat[k] for k in i],
            current=w["report_date"].dt.strftime("%Y-%m-%d").iloc[i].values,
            suggested=w["sia_date"].dt.strftime("%Y-%m-%d").iloc[i].values))
    i = np.where((w["report_date"].notna() & (w["report_date"] > today)).values)[0]
    add(_qi(i, "report_date", "Warning", "Report date is in the future",
            current=w["report_date"].dt.strftime("%Y-%m-%d").iloc[i].values))
    i = np.where((d > cfg["report_late_days"]).fillna(False).values)[0]
    add(_qi(i, "report_date", "Info", ["Report date is %d days after the SIA start" % d.iat[k] for k in i],
            current=w["report_date"].dt.strftime("%Y-%m-%d").iloc[i].values))
    # --- numbers
    t, vv, tp = w["target"], w["vacc"], w["tot_pop"]
    neg = np.where(((tp < 0) | (t < 0) | (vv < 0)).fillna(False).values)[0]
    add(_qi(neg, "negative_number", "Error", "Population or number vaccinated is negative"))
    i = np.where((vv.notna() & (vv > 0) & (t.isna() | (t <= 0))).values)[0]
    add(_qi(i, "target_missing", "Error", "Vaccinated children reported but the target population is missing or zero",
            current=t.iloc[i].values))
    i = np.where(((tp > 0) & t.notna() & (t > tp * 1.0001)).fillna(False).values)[0]
    add(_qi(i, "target_gt_total", "Warning",
            ["Target population (%s) is larger than the total population (%s)" % (_qfmt(t.iat[k]), _qfmt(tp.iat[k])) for k in i]))
    with np.errstate(divide="ignore", invalid="ignore"):
        rc = vv / t.where(t > 0)
    okc = rc.notna() & np.isfinite(rc)
    i = np.where((okc & (rc > cfg["cv_ceiling"])).values)[0]
    add(_qi(i, "coverage_ceiling", "Warning",
            ["Vaccinated / target = %d%% (above the %d%% ceiling): invalid for coverage analysis until verified"
             % (round(100 * rc.iat[k]), round(100 * cfg["cv_ceiling"])) for k in i],
            current=["%.1f%%" % (100 * rc.iat[k]) for k in i]))
    i = np.where((okc & (rc > cfg["cv_high"]) & (rc <= cfg["cv_ceiling"])).values)[0]
    add(_qi(i, "coverage_high", "Info", ["Vaccinated / target = %d%%: please verify target and vaccinated" % round(100 * rc.iat[k]) for k in i],
            current=["%.1f%%" % (100 * rc.iat[k]) for k in i]))
    rep_ = w["cv_rep"]
    i = np.where((okc & rep_.notna() & ((rc - rep_).abs() > cfg["cv_recalc_tol"])).values)[0]
    sev = np.where(np.abs(rc.iloc[i].values - rep_.iloc[i].values) > 0.10, "Warning", "Info")
    msg = ["Reported coverage %.1f%% but vaccinated / target = %.1f%%" % (100 * rep_.iat[k], 100 * rc.iat[k]) for k in i]
    cur = ["%.1f%%" % (100 * rep_.iat[k]) for k in i]
    sug = ["%.1f%%" % (100 * rc.iat[k]) for k in i]
    for s_ in ("Warning", "Info"):
        m_ = sev == s_
        add(_qi(i[m_], "coverage_mismatch", s_, [x for x, f in zip(msg, m_) if f],
                current=[x for x, f in zip(cur, m_) if f], suggested=[x for x, f in zip(sug, m_) if f]))
    # --- target population compared with the same district's other campaign dates
    dj = w[(w["target"] > 0) & w["sia_date"].notna()].copy()
    if len(dj):
        dj["grp"] = dj["c_key"] + "|" + prov_eff.loc[dj.index].map(_qn) + "|" + dj["d_key"] + "|" + dj["age_grp"].map(_qn)
        per = dj.groupby(["grp", "sia_date"])["target"].median().reset_index(name="tp")
        per["n_dates"] = per.groupby("grp")["grp"].transform("size")
        per = per[per["n_dates"] >= cfg["pop_min_rounds"]].copy()
        if len(per):
            per = per.sort_values(["grp", "sia_date"]).reset_index(drop=True)
            per["med_others"] = np.nan
            for _, g in per.groupby("grp", sort=False):
                a_ = g["tp"].values
                per.loc[g.index, "med_others"] = [np.median(np.delete(a_, k)) for k in range(len(a_))]
            per["ratio"] = per["tp"] / per["med_others"]
            fl = per[(per["ratio"] >= cfg["pop_factor"]) | (per["ratio"] <= 1 / cfg["pop_factor"])].copy()
            if len(fl):
                def digit(r):
                    for k in (1, 2, 3):
                        if abs(r / 10 ** k - 1) < 0.25:
                            return "possible extra digit"
                        if abs(r * 10 ** k - 1) < 0.25:
                            return "possible missing digit"
                    return ""
                fl["note"] = fl["ratio"].map(digit)
                hit = dj.merge(fl[["grp", "sia_date", "med_others", "ratio", "note"]], on=["grp", "sia_date"], how="inner")
                msgs = ["Target %s is %s the district median of its other campaign dates (%s)%s" % (
                    _qfmt(r.target), ("%.1fx above" % r.ratio) if r.ratio >= 1 else ("%.1fx below" % (1 / r.ratio)),
                    _qfmt(round(r.med_others)), (" - " + r.note) if r.note else "") for r in hit.itertuples()]
                add(_qi(hit["idx"].values, "target_jump", "Warning", msgs, current=[_qfmt(x) for x in hit["target"]],
                        suggested=[_qfmt(round(x)) for x in hit["med_others"]]))
    # --- co-administration
    comb = np.where((w["fam"] == "combined").values)[0]
    add(_qi(comb, "coadmin_combined", "Info",
            "Combined nOPV2 & bOPV record: the total is the larger of the two vaccines; confirm the figures apply to both",
            current=w["Vaccine"].iloc[comb].values))
    gk = w["c_key"] + "|" + w["d_key"] + "|" + w["sia_date"].dt.strftime("%Y-%m-%d").fillna("")
    has_c = (w["fam"] == "combined").groupby(gk).transform("any")
    has_s = w["fam"].isin(["nOPV2", "bOPV"]).groupby(gk).transform("any")
    i = np.where((has_c & has_s & w["fam"].isin(["combined", "nOPV2", "bOPV"])).values)[0]
    add(_qi(i, "coadmin_mixed", "Warning",
            "Same district and start date reported as a combined record and as separate vaccine records (risk of double counting)",
            current=w["Vaccine"].iloc[i].values))
    sep = w[w["fam"].isin(["nOPV2", "bOPV"]) & w["vacc"].notna()].copy()
    if len(sep):
        sep["gk"] = gk.loc[sep.index]
        g = sep.groupby("gk").agg(nf=("fam", "nunique"), nv=("vacc", "nunique"), nt=("target", "nunique"))
        gg = set(g[(g["nf"] == 2) & (g["nv"] == 1) & (g["nt"] == 1)].index)
        i = sep["idx"].values[sep["gk"].isin(gg).values]
        add(_qi(i, "coadmin_identical", "Info",
                "nOPV2 and bOPV records of the same district and date have identical target and vaccinated: please confirm",
                current=w["vacc"].iloc[i].values))
    both = (w["fam"] == "combined") & w["tn"].notna() & w["tb"].notna()
    i = np.where((both & (w["tn"] == w["tb"]) & (w["tn"] > 0)).values)[0]
    add(_qi(i, "coadmin_identical", "Info",
            "Combined record with the same figure for both vaccines (taken as both when only one total was reported): please confirm",
            current=w["tn"].iloc[i].values))
    mx = w[["tn", "tb"]].max(axis=1)
    i = np.where((both & ((w["vacc"].fillna(0) - mx.fillna(0)).abs() > 0.5)).values)[0]
    add(_qi(i, "coadmin_total", "Error", "TotalNbrVaccPolio is not the larger of Tot_Vacc_nOPV2 and Total_Vacc_bOPV",
            current=w["vacc"].iloc[i].values, suggested=mx.iloc[i].values))
    # --- vaccine spelling
    vtrim = w["Vaccine"].str.strip()
    canon_map = {c.lower(): c for c in QC_VACCINE_CANON}
    canon = vtrim.str.lower().map(canon_map)
    i = np.where((~_qblank(vtrim) & canon.notna() & (vtrim != canon)).values)[0]
    add(_qi(i, "vaccine_spelling", "Info", ["Vaccine type '%s' is spelled differently from '%s'" % (vtrim.iat[k], canon.iat[k]) for k in i],
            current=vtrim.iloc[i].values, suggested=canon.iloc[i].values))
    # --- duplicates
    dup = w["key_id"].duplicated(keep=False)
    i = np.where(dup.values)[0]
    cnt = w["key_id"].map(w["key_id"].value_counts())
    add(_qi(i, "duplicate_key", "Warning",
            ["Same country, province, district, SIA date, round, vaccine and response appears %d times" % cnt.iat[k] for k in i]))

    iss = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=["idx", "check", "severity", "message", "current", "suggested"])
    if n_ctx:
        iss = iss[(iss["idx"] >= n_ctx) & (iss["check"] != "duplicate_key")].copy()
        iss["idx"] = iss["idx"] - n_ctx
        w = w.iloc[n_ctx:].copy()
        w["idx"] = w["idx"] - n_ctx
        if corr is not None and len(corr):
            corr = corr[corr["idx"] >= n_ctx].copy()
            corr["idx"] = corr["idx"] - n_ctx
    base = w[["idx", "Country", "Province", "District", "sia_date", "Round", "Vaccine", "Response"]].copy()
    base["SIA_date"] = base.pop("sia_date").dt.strftime("%Y-%m-%d")
    iss = iss.merge(base, on="idx", how="left")
    iss["check_label"] = iss["check"].map(QC_LABELS)
    iss["_o"] = iss["severity"].map({s: k for k, s in enumerate(QC_SEV)})
    iss = iss.sort_values(["_o", "Country", "check", "District", "SIA_date"]).drop(columns="_o")
    iss["Excel_row"] = iss["idx"] + 2
    iss = iss[["severity", "check", "check_label", "Country", "Province", "District", "SIA_date", "Round", "Vaccine",
               "Response", "message", "current", "suggested", "Excel_row", "idx"]].reset_index(drop=True)
    if corr is not None and len(corr):
        corr = corr.assign(Excel_row=corr["idx"] + 2)
    # --- summaries
    if len(iss):
        by_check = (iss.groupby(["check_label", "severity"]).agg(Records=("idx", "nunique"), Countries=("Country", "nunique"))
                    .reset_index().rename(columns={"check_label": "Check", "severity": "Severity"}))
        by_check["_o"] = by_check["Severity"].map({s: k for k, s in enumerate(QC_SEV)})
        by_check = by_check.sort_values(["_o", "Records"], ascending=[True, False]).drop(columns="_o")
        tot = w.groupby("Country").size().rename("Records")
        s = iss.groupby("Country").apply(lambda g: pd.Series({
            "Errors": g.loc[g.severity == "Error", "idx"].nunique(), "Warnings": g.loc[g.severity == "Warning", "idx"].nunique(),
            "Infos": g.loc[g.severity == "Info", "idx"].nunique(), "Records with an issue": g["idx"].nunique()}), include_groups=False)
        bc = pd.concat([tot, s], axis=1).fillna(0).reset_index().rename(columns={"index": "Country"})
        ew = iss[iss.severity.isin(["Error", "Warning"])].groupby("Country")["idx"].nunique()
        bc["% with an error or warning"] = (100 * bc["Country"].map(ew).fillna(0) / bc["Records"].clip(lower=1)).round(1)
        by_country = bc.sort_values(["Errors", "Warnings"], ascending=False)
    else:
        by_check = pd.DataFrame(columns=["Check", "Severity", "Records", "Countries"])
        by_country = w.groupby("Country").size().rename("Records").reset_index()
    return {"issues": iss, "corrections": corr, "by_check": by_check, "by_country": by_country, "n_rows": n_new,
            "n_flagged": int(iss["idx"].nunique()), "config": cfg, "prov_eff": prov_eff,
            "n_err": int(iss.loc[iss.severity == "Error", "idx"].nunique()),
            "n_warn": int(iss.loc[iss.severity == "Warning", "idx"].nunique()),
            "n_info": int(iss.loc[iss.severity == "Info", "idx"].nunique())}


def qc_write(res, path, country=None):
    cfg = res["config"]
    ov = pd.DataFrame({"Item": ["Records checked", "Records with at least one issue", "Errors (records)", "Warnings (records)",
                                "Info (records)", "Coverage ceiling", "Target population flag threshold (times higher or lower than usual)",
                                "Run at"],
                       "Value": [res["n_rows"], res["n_flagged"], res["n_err"], res["n_warn"], res["n_info"],
                                 "%d%%" % round(100 * cfg["cv_ceiling"]), "%gx" % cfg["pop_factor"],
                                 pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")]})
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        ov.to_excel(xw, sheet_name="Overview", index=False)
        res["by_check"].to_excel(xw, sheet_name="By check", index=False)
        res["by_country"].to_excel(xw, sheet_name="By country", index=False)
        res["issues"].drop(columns="idx").to_excel(xw, sheet_name="Issues", index=False)
        if res["corrections"] is not None and len(res["corrections"]):
            res["corrections"].drop(columns="idx").to_excel(xw, sheet_name="Suggested fixes", index=False)
        if res.get("fixes") is not None and len(res["fixes"]):
            res["fixes"].to_excel(xw, sheet_name="Fixes applied", index=False)
        for key_, sheet in (("dup_log", "Removed_duplicates"), ("dup_groups", "Duplicate_groups"),
                            ("existing", "Already_in_repository")):
            if res.get(key_) is not None and len(res[key_]):
                res[key_].to_excel(xw, sheet_name=sheet, index=False)
    print("quality report written:", path, flush=True)


def qc_print(res):
    print("quality checks: %d records | Errors: %d | Warnings: %d | Info: %d records (a record can have several)" %
          (res["n_rows"], res["n_err"], res["n_warn"], res["n_info"]), flush=True)
    bc = res["by_check"]
    for r in bc[bc["Severity"] == "Error"].itertuples():
        print("   Error   %-55s %6d records" % (r.Check, r.Records), flush=True)
    for r in bc[bc["Severity"] == "Warning"].itertuples():
        print("   Warning %-55s %6d records" % (r.Check, r.Records), flush=True)


def run_check_only(a, rev_dir):
    """--check-only: run the quality checks on a file as it is (like run_admin_quality_check.R of the portal)."""
    print("loading", a.src, flush=True)
    xl = pd.ExcelFile(a.src)
    x = xl.parse(SHEET if SHEET in xl.sheet_names else xl.sheet_names[0], keep_default_na=False, na_values=[""])
    x = x.rename(columns={"SIA-date": "SIA_date"}) if "SIA_date" not in x.columns else x
    if a.country:
        x = x[x["Country"].map(_qcountry) == _qcountry(a.country)]
    orig = x.index.values
    x = x.reset_index(drop=True)
    print(len(x), "records", flush=True)
    if not len(x):
        sys.exit("no records to check")
    geo = None
    gp = find_geo(a.geo, a.src)
    if gp:
        geo = load_geo(gp)
        print("geography reference: %s (%d districts)" % (gp, len(geo)), flush=True)
    else:
        print("geo_reference.csv not found: the province / country check is skipped", flush=True)
    x["_row"] = orig + 2
    dup_drop, dup_log, dup_groups = find_duplicates(x, "none")
    print("duplicates: %d exact copies (would be removed by the update); %d groups with the same key but different values" %
          (len(dup_log), int(dup_groups["duplicate_group"].nunique()) if len(dup_groups) else 0), flush=True)
    ctx_, existing_rows = None, pd.DataFrame()
    if a.repository:
        try:
            ex_mask, rp = existing_in_repository(x, a.repository)
            existing_rows = x.loc[ex_mask, ["_row", "Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response"]] \
                .rename(columns={"_row": "Excel_row"})
            ctx_ = rp.rename(columns={"SIA-date": "SIA_date"})
            print("repository: %d records of the file already exist in it" % int(ex_mask.sum()), flush=True)
        except Exception as e:
            print("repository not used (%s): %s" % (a.repository, e), flush=True)
    res = qc_run(x.drop(columns="_row"), geo, qc_config(cv_ceiling=a.cv_ceiling / 100.0, pop_factor=a.pop_factor), context=ctx_)
    res["dup_log"], res["dup_groups"], res["existing"] = dup_log, dup_groups, existing_rows
    res["issues"]["Excel_row"] = orig[res["issues"]["idx"].values] + 2
    if res["corrections"] is not None and len(res["corrections"]):
        res["corrections"]["Excel_row"] = orig[res["corrections"]["idx"].values] + 2
    qc_print(res)
    qc_write(res, os.path.join(rev_dir, "AFRO_admin_data_quality_report.xlsx"))
    if a.fail_on_error and res["n_err"]:
        sys.exit(2)


def model_poppolio(df, out_total, out_vt, pop_new, pop_reason, pop_fixed, drop):
    """Estimate PopPolio for records whose coverage is still above CV_HIGH.

    Gradient boosting on the district's own history: for every record the other campaigns of the same
    district (more than 0.3 years apart) with a credible coverage (0.5 - CV_HIGH) give the features
    (nearest / median / previous / next population, vaccinated history, median coverage, time, country).
    The model is validated with 5-fold cross-validation grouped by district (a district is never in
    training and test at the same time). A prediction is applied only when it gives a coverage between
    0.5 and CV_HIGH. Records without enough history are left unchanged.
    """
    try:
        from sklearn.ensemble import HistGradientBoostingRegressor
        from sklearn.model_selection import GroupKFold
    except ImportError:
        sys.exit("--model-poppolio needs scikit-learn (pip install scikit-learn)")

    low = out_vt.map(lambda x: x.lower() if isinstance(x, str) else "")
    polio = low.str.contains("opv") & ~drop
    t = pd.to_datetime(df["SIA_date"], errors="coerce")
    tt = t.dt.year + t.dt.dayofyear / 365.25
    dk = pd.Series(["|".join(str(v) for v in r) for r in df[["Country", "Province", "District"]].itertuples(index=False)],
                   index=df.index)
    with np.errstate(divide="ignore", invalid="ignore"):
        cv = out_total / pop_new.where(pop_new > 0)
    median_fixed = pop_reason.str.startswith("District median")
    cred = polio & cv.between(0.5, CV_HIGH) & tt.notna() & ~median_fixed
    high = polio & (cv > CV_HIGH) & tt.notna() & ~pop_fixed
    country = df["Country"].astype("category").cat.codes

    hist = {}
    for k, g in df.index[cred].to_series().groupby(dk[cred]):
        hist[k] = np.column_stack([g.values, tt[g.values], pop_new[g.values], out_total[g.values]]).astype(float)

    rows, idx = [], []
    for i in df.index[cred | high]:
        g = hist.get(dk[i])
        if g is None:
            continue
        o = g[(g[:, 0] != i) & (np.abs(g[:, 1] - tt[i]) > 0.3)]
        if len(o) < 2:
            continue
        near = o[np.argsort(np.abs(o[:, 1] - tt[i]))[:2]]
        prior, post = o[o[:, 1] < tt[i]], o[o[:, 1] >= tt[i]]
        rows.append(dict(
            n_hist=len(o), near_pop=np.log(near[:, 2].mean()), med_pop=np.log(np.median(o[:, 2])),
            last_pop=np.log(prior[np.argmax(prior[:, 1]), 2]) if len(prior) else np.nan,
            next_pop=np.log(post[np.argmin(post[:, 1]), 2]) if len(post) else np.nan,
            near_vacc=np.log(near[:, 3].mean()), med_vacc=np.log(np.median(o[:, 3])),
            med_cv=np.median(o[:, 3] / o[:, 2]), dt_near=np.abs(o[:, 1] - tt[i]).min(), t=tt[i], country=country[i]))
        idx.append(i)
    if len(rows) < 50:
        print("population model skipped: not enough district history in this file", flush=True)
        return None
    X = pd.DataFrame(rows, index=idx)
    feat = list(X.columns)
    y = np.log(pop_new[idx].where(pop_new[idx] > 0))
    is_cred = cred[idx].values
    tr, hi = X[is_cred], X[~is_cred & high[idx].values]

    def make():
        return HistGradientBoostingRegressor(max_iter=300, learning_rate=0.05, max_leaf_nodes=15,
                                             random_state=0, loss="absolute_error")

    oof = pd.Series(np.nan, index=tr.index)
    groups = dk[tr.index]
    if len(tr) < 50 or groups.nunique() < 5 or len(hi) == 0:
        print("population model skipped: not enough district history in this file", flush=True)
        return None
    for a_, b_ in GroupKFold(5).split(tr, groups=groups):
        oof.iloc[b_] = make().fit(tr.iloc[a_], y[tr.index].iloc[a_]).predict(tr.iloc[b_])
    err = (np.exp(oof) - np.exp(y[tr.index])).abs() / np.exp(y[tr.index])
    model = make().fit(tr, y[tr.index])
    pred = pd.Series(np.exp(model.predict(hi)), index=hi.index).round()
    new_cv = out_total[hi.index] / pred
    ok = new_cv.between(0.5, CV_HIGH)
    for i in pred.index[ok]:
        pop_new.at[i] = pred[i]
        pop_reason.at[i] = "Model estimate from the district's history (typical error about %d%%)" % round(100 * err.median())
        pop_fixed.at[i] = True
    info = {
        "Method": "Gradient boosting on the district's own history, validated by 5-fold CV grouped by district",
        "Records used for training (credible coverage)": int(len(tr)),
        "Records above %.1f with enough history" % CV_HIGH: int(len(hi)),
        "Estimates applied (coverage 0.5 - %.1f after correction)" % CV_HIGH: int(ok.sum()),
        "Left unchanged (prediction not plausible)": int((~ok).sum()),
        "Left unchanged (district without enough history)": int(high.sum() - len(hi)),
        "Backtest: median absolute error": "%.1f%%" % (100 * err.median()),
        "Backtest: mean absolute error": "%.1f%%" % (100 * err.mean()),
        "Backtest: share within +/-20%": "%.0f%%" % (100 * (err <= 0.2).mean()),
        "Backtest: 80th percentile of the error": "%.1f%%" % (100 * err.quantile(0.8)),
        "Note": "Estimates, not reported data. Each corrected record is listed in PopPolio_corrected.",
    }
    print("PopPolio model: %d estimates applied, %d records left (backtest median error %.0f%%)" % (
        int(ok.sum()), int(high.sum() - ok.sum()), 100 * err.median()), flush=True)
    return info


def write_feedback(path, df, keep, out_vt, out_total, pop, pop_before, pop_fixed, pop_reason, cv_n, cv_b, cvp,
                   approach=None, out_tn=None, out_tb=None):
    """Workbook for the focal points: records that still look wrong, populations that were estimated, and the
    combined records whose nOPV2 and bOPV totals were assumed equal."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as L

    mx = pd.concat([cv_n, cv_b], axis=1).max(axis=1)
    ident = ["Country", "Province", "District", "Year", "Round_Add", "Response"]

    def base(mask):
        d = df.loc[mask, ident].copy()
        d.insert(0, "Row_in_original_file", df.loc[mask, "_row"])
        d["Vaccine_type"] = out_vt[mask]
        d["Children_vaccinated (TotalNbrVaccPolio)"] = out_total[mask]
        return d

    sus = base((mx > CV_HIGH) & keep)
    sus["Target_population (PopPolio)"] = pop[sus.index]
    sus["Coverage (vaccinated / population)"] = mx[sus.index].round(2)
    sus["CVPolio_in_data"] = cvp[sus.index].round(2)
    sus["Level"] = pd.cut(mx[sus.index], [CV_HIGH, 2, 3, 5, np.inf],
                          labels=["1.5 - 2", "2 - 3", "3 - 5", "above 5"]).astype(str)
    sus["Why it looks wrong"] = ("More children vaccinated than the target population: the population is probably "
                                 "too low, or the vaccinated count includes other areas or rounds.")
    sus["Correct target population"] = None
    sus["Correct children vaccinated"] = None
    sus["Comment"] = None
    sus = sus.sort_values(["Country", "Coverage (vaccinated / population)"], ascending=[True, False])

    est_mask = pop_fixed & keep & (pop_reason.str.startswith("Model") | pop_reason.str.startswith("District median"))
    est = base(est_mask)
    est["Population_in_source"] = pop_before[est.index]
    est["Estimated_population"] = pop[est.index]
    est["Coverage_in_source"] = (out_total[est.index] / pop_before[est.index].where(pop_before[est.index] > 0)).round(2)
    est["Coverage_with_estimate"] = (out_total[est.index] / pop[est.index].where(pop[est.index] > 0)).round(2)
    est["How it was estimated"] = pop_reason[est.index]
    est["Confirmed population"] = None
    est["Comment"] = None
    est = est.sort_values(["Country", "Province", "District", "Year"])

    comb = None
    if approach is not None:
        cm = approach.isin([A_COMBINED, A_RESP]) & keep & (out_tn == out_tb)
        comb = base(cm)
        comb["Tot_Vacc_nOPV2 (assumed = total)"] = out_tn[comb.index]
        comb["Total_Vacc_bOPV (assumed = total)"] = out_tb[comb.index]
        comb["How it was reported"] = approach[comb.index]
        comb["Correct Tot_Vacc_nOPV2"] = None
        comb["Correct Total_Vacc_bOPV"] = None
        comb["Comment"] = None
        comb = comb.sort_values(["Country", "Year", "Round_Add", "Province", "District"])

    cnt = sus.groupby("Country").agg(Records=("Level", "size"),
                                     Above_5=("Level", lambda x: int((x == "above 5").sum())),
                                     Highest_coverage=("Coverage (vaccinated / population)", "max"))
    cnt["Estimated_populations_to_confirm"] = est.groupby("Country").size()
    if comb is not None:
        cnt["Combined_records_to_validate"] = comb.groupby("Country").size()
    cnt = cnt.fillna(0).reset_index()
    for c_ in cnt.columns:
        if c_ not in ("Country", "Highest_coverage"):
            cnt[c_] = cnt[c_].astype(int)
    cnt = cnt.sort_values("Records", ascending=False)
    how = pd.DataFrame({"Read me": [
        "Sheet To_correct: records where more children were vaccinated than the target population (coverage above %.1f)." % CV_HIGH,
        "Please give the correct target population (and the correct number vaccinated if that is what is wrong).",
        "Sheet Estimates_to_confirm: populations that were not reported correctly and were estimated (from the stored "
        "CVPolio, the district's other rounds, or a model on the district's history). Please confirm or correct.",
        "Sheet Combined_to_validate: records reported as one combined nOPV2 & bOPV record. By default nOPV2 and bOPV "
        "were given the same number of children (the reported total). If the two vaccines reached different "
        "numbers of children, please give the correct nOPV2 and bOPV totals.",
        "Row_in_original_file is the row number in AFRO_admin_data.xlsx before the co-administration merge.",
    ]})
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        cnt.to_excel(xw, sheet_name="By_country", index=False)
        sus.to_excel(xw, sheet_name="To_correct", index=False)
        est.to_excel(xw, sheet_name="Estimates_to_confirm", index=False)
        if comb is not None:
            comb.to_excel(xw, sheet_name="Combined_to_validate", index=False)
        how.to_excel(xw, sheet_name="Read_me", index=False)
        for name, ws in xw.sheets.items():
            ws.freeze_panes = "A2" if name != "To_correct" else "D2"
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="005C97")
                c.alignment = Alignment(wrap_text=True, vertical="center")
            for j, col in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 200)), start=1):
                w = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[L(j)].width = min(max(w + 2, 10), 45)
            if name in ("To_correct", "Estimates_to_confirm", "By_country", "Combined_to_validate"):
                ws.auto_filter.ref = ws.dimensions
        for name in ("To_correct", "Estimates_to_confirm", "Combined_to_validate"):
            if name not in xw.sheets:
                continue
            ws = xw.sheets[name]
            hdr = [c.value for c in ws[1]]
            for j, h in enumerate(hdr, start=1):
                if h and ("population" in str(h).lower() or "vaccinated" in str(h).lower()) and "Coverage" not in str(h):
                    for r in range(2, ws.max_row + 1):
                        ws.cell(r, j).number_format = "#,##0"
                if h in ("Correct target population", "Correct children vaccinated", "Confirmed population", "Comment",
                         "Correct Tot_Vacc_nOPV2", "Correct Total_Vacc_bOPV"):
                    for r in range(2, ws.max_row + 1):
                        ws.cell(r, j).fill = PatternFill("solid", fgColor="FFF8E1")
    print("feedback file written:", path, "| to correct:", len(sus), "| estimates to confirm:", len(est),
          "| combined to validate:", 0 if comb is None else len(comb), flush=True)


def write_audit(path, df, approach, vt_orig, out_vt, vt_changed, total, out_total, out_tn, out_tb,
                drop, merged_log, keep, paired, pop_fixed=None, pop_before=None, pop_after=None, cvp=None,
                pop_reason=None, cv_n=None, cv_b=None, model_info=None, repairs=None, lk_report=None, copy_log=None, pre_label=None):
    cols = ["_row", "Country", "Province", "District", "Admin_3", "Year", "Round_Add", "Response", "SIA_date"]
    a = df[cols].copy()
    a.rename(columns={"_row": "Original_row"}, inplace=True)
    a["Approach"] = approach
    a["Vaccine_type_before"] = vt_orig
    a["Vaccine_type_after"] = out_vt.where(keep, "(merged into the other record)")
    if pre_label is not None and (pre_label != "").any():
        m_ = (pre_label != "").values
        a.loc[m_, "Vaccine_type_after"] = pre_label[m_].values
    a["TotalNbrVaccPolio_before"] = total
    a["TotalNbrVaccPolio_after"] = out_total.where(keep)
    a["Tot_Vacc_nOPV2"] = out_tn.where(keep)
    a["Total_Vacc_bOPV"] = out_tb.where(keep)
    a["Record_removed"] = drop
    summ = (a.groupby(["Country", "Approach"]).size().rename("Records").reset_index()
            .sort_values(["Country", "Approach"]))
    changed = a[(a["Approach"] != A_NONE) | vt_changed]
    review = a[a["Approach"].isin([A_KEPT, A_RESP])]
    vtfix = a[vt_changed][["Original_row", "Country", "Province", "District", "Round_Add", "Response",
                           "Vaccine_type_before", "Vaccine_type_after"]]
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        summ.to_excel(xw, sheet_name="Summary", index=False)
        if repairs:
            pd.DataFrame(repairs).to_excel(xw, sheet_name="Repairs", index=False)
        if lk_report is not None:
            lk_report[0].to_excel(xw, sheet_name="Lookup_check", index=False)
            lk_report[1].to_excel(xw, sheet_name="Lookup_no_data", index=False)
        if copy_log is not None and len(copy_log):
            copy_log.to_excel(xw, sheet_name="Removed_copies", index=False)
        review.to_excel(xw, sheet_name="To_review", index=False)
        vtfix.to_excel(xw, sheet_name="Vaccine_type_fixes", index=False)
        changed.to_excel(xw, sheet_name="All_co-admin_records", index=False)
        if pop_fixed is not None and pop_fixed.any():
            pf = df.loc[pop_fixed, ["_row", "Country", "Province", "District", "Year", "Round_Add", "Response"]].copy()
            pf.rename(columns={"_row": "Original_row"}, inplace=True)
            pf["Vaccine_type"] = out_vt[pop_fixed]
            pf["TotalNbrVaccPolio"] = out_total[pop_fixed]
            pf["CVPolio"] = cvp[pop_fixed]
            pf["PopPolio_before"] = pop_before[pop_fixed]
            pf["PopPolio_after"] = pop_after[pop_fixed].astype("int64")
            pf["Reason"] = pop_reason[pop_fixed]
            pf.sort_values(["Country", "Province", "District"]).to_excel(xw, sheet_name="PopPolio_corrected", index=False)
        if model_info is not None:
            pd.DataFrame({"Measure": list(model_info.keys()), "Value": list(model_info.values())}).to_excel(
                xw, sheet_name="Model_backtest", index=False)
        if cv_n is not None:
            mx = pd.concat([cv_n, cv_b], axis=1).max(axis=1)
            rem = df.loc[(mx > CV_HIGH) & keep, ["_row", "Country", "Province", "District", "Year", "Round_Add", "Response"]].copy()
            rem.rename(columns={"_row": "Original_row"}, inplace=True)
            rem["Vaccine_type"] = out_vt[rem.index]
            rem["TotalNbrVaccPolio"] = out_total[rem.index]
            rem["PopPolio"] = pop_after[rem.index]
            rem["CV_nOPV2"] = cv_n[rem.index]
            rem["CV_bOPV"] = cv_b[rem.index]
            rem["CVPolio"] = cvp[rem.index]
            rem.sort_values(["Country", "Province", "District", "Year"]).to_excel(
                xw, sheet_name="CV_still_above_1.5", index=False)
    print("audit written:", path)


# =====================================================================================================
# Repair of copy / drag-fill errors (uses the SharePoint backups as the reference)
# =====================================================================================================
REF_COLS = ["Country", "Province", "District", "Admin_3", "SIA_date", "Round_Add", "Response", "Vaccine_type",
            "TotalNbrVaccPolio", "PopPolio"]
COUNTRY_ALIAS = {"CAMEROUN": "CAMEROON", "RÉPUBLIQUE CENTRAFRICAINE": "CENTRAL AFRICAN REPUBLIC",
                 "RÉPUBLIQUE CENTRAFRIQUE": "CENTRAL AFRICAN REPUBLIC", "CAR": "CENTRAL AFRICAN REPUBLIC"}
KNOWN_BAD_NOPV = re.compile(r"^nopv(\d+)$", re.I)


def _ckey(c):
    c = str(c).strip().upper()
    return COUNTRY_ALIAS.get(c, c)


def _prefix(r):
    m = re.match(r"^([A-Za-z]+)", r) if isinstance(r, str) else None
    return m.group(1).upper() if m else None


def _dkey(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    return x.strftime("%Y-%m-%d") if hasattr(x, "strftime") else str(x)[:10]


def _s(x):
    """Text of a cell for matching; an empty cell, NaN and the text 'NA' (as pandas reads it) are all ''."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    t = str(x).strip()
    return "" if t.upper() in ("NA", "N/A", "NAN", "NONE", "NULL") else t


def _numkey(x):
    v = pd.to_numeric(x, errors="coerce")
    return "" if pd.isna(v) else str(round(float(v), 2))


def _prep(d):
    """Add the helper columns used to recognise a row in another version of the file."""
    d = d.copy()
    for c in REF_COLS:
        if c not in d.columns:
            d[c] = np.nan
    d["_ck"] = [_ckey(c) for c in d["Country"]]
    d["_pre"] = [_prefix(r) for r in d["Response"]]
    parts = [(str(a), _s(b), _s(c), _s(e), _dkey(f), _s(g), _numkey(h), _numkey(i)) for a, b, c, e, f, g, h, i in
             zip(d["_ck"], d["Province"], d["District"], d["Admin_3"], d["SIA_date"], d["Round_Add"],
                 d["TotalNbrVaccPolio"], d["PopPolio"])]
    d["_kA"] = ["|".join(p) for p in parts]
    d["_kB"] = ["|".join(p[:6]) for p in parts]
    d["_kA"] = d["_kA"] + "#" + d.groupby("_kA").cumcount().astype(str)
    d["_kB"] = d["_kB"] + "#" + d.groupby("_kB").cumcount().astype(str)
    return d


def find_backups(user_dir, src):
    import glob
    cands = []
    if user_dir:
        cands.append(user_dir)
    if os.environ.get("ADMIN_BACKUP_DIR"):
        cands.append(os.environ["ADMIN_BACKUP_DIR"])
    here = os.path.dirname(os.path.abspath(__file__))
    for base in (here, os.getcwd(), os.path.dirname(os.path.abspath(src))):
        cands.append(os.path.join(base, "..", "SharePoint_version_backup", "snapshots", "admin"))
        cands.append(base)
    for d in cands:
        files = sorted(glob.glob(os.path.join(d, "AFRO_admin_data__*.xlsx")))
        if files:
            return files
    return []


def _md5(path):
    import hashlib
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_backup(path, cache_dir):
    cache = os.path.join(cache_dir, os.path.basename(path) + ".pkl")
    if os.path.exists(cache) and os.path.getmtime(cache) >= os.path.getmtime(path):
        return pd.read_pickle(cache)
    d = pd.read_excel(path, sheet_name=0, usecols=lambda c: c in REF_COLS)
    os.makedirs(cache_dir, exist_ok=True)
    d.to_pickle(cache)
    return d


def _owners(frames):
    """Response prefix (country code at the start of the name) -> the country that owns it, learned from the
    versions of the file: a prefix is owned by a country that holds 95% of its rows in at least half of the
    versions where the prefix is used (a copy error in one version cannot change the owner)."""
    present, votes = {}, {}
    for f in frames:
        t = f.dropna(subset=["_pre"]).groupby(["_pre", "_ck"]).size().unstack(fill_value=0)
        tot = t.sum(axis=1)
        for p in t.index:
            if tot[p] >= 20:
                present[p] = present.get(p, 0) + 1
                c = t.loc[p].idxmax()
                if t.loc[p, c] / tot[p] >= 0.95:
                    votes.setdefault(p, {}).setdefault(c, 0)
                    votes[p][c] += 1
    owners = {}
    for p, cn in votes.items():
        if len(cn) == 1:
            c, v = next(iter(cn.items()))
            if v * 2 >= present[p]:
                owners[p] = c
    return owners


def _nk(x):
    """District / province name for matching: no accents, capitals, letters and digits only."""
    import unicodedata
    t = unicodedata.normalize("NFKD", _s(x)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]", "", t)


# ---------------------------------------------------------------- campaign lookup (planned and implemented rounds)
LOOKUP_COUNTRY_ALIAS = {"UNITEDREPUBLICOFTANZANIA": "TANZANIA", "CAMEROUN": "CAMEROON",
                        "REPUBLIQUECENTRAFRIQUE": "CENTRALAFRICANREPUBLIC",
                        "REPUBLIQUECENTRAFRICAINE": "CENTRALAFRICANREPUBLIC", "CAR": "CENTRALAFRICANREPUBLIC"}


def _lkc(c):
    t = _nk(c)
    return LOOKUP_COUNTRY_ALIAS.get(t, t)


def _rk(r):
    """Response / OBR name for matching: lower case, letters and digits only (nOPV-sNID == nOPV-Snid)."""
    return re.sub(r"[^a-z0-9]", "", str(r).lower()) if isinstance(r, str) else ""


def _vset(v):
    return {m.lower() for m in re.findall(r"[nbm]OPV\d?", str(v), re.I)}


def _rnum(x):
    m = re.search(r"(\d+)", str(x))
    return int(m.group(1)) if m else None


def find_lookup(user_path, src):
    """The lookup is kept in the repository, next to this script (lookup.xlsx); the fetch pipeline refreshes it."""
    cands = [user_path] if user_path else []
    here = os.path.dirname(os.path.abspath(__file__))
    cands.append(os.path.join(here, "lookup.xlsx"))
    if os.environ.get("ADMIN_LOOKUP"):
        cands.append(os.environ["ADMIN_LOOKUP"])
    for base in (os.path.dirname(os.path.abspath(src)), os.getcwd(), here):
        for rel in ("lookup.xlsx", os.path.join("data", "lookup.xlsx"), os.path.join("..", "data", "lookup.xlsx")):
            cands.append(os.path.join(base, rel))
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def load_lookup(path):
    L = pd.read_excel(path)
    need = ["Round Number", "OBR Name", "Country", "Round Start Date"]
    miss = [c for c in need if c not in L.columns]
    if miss:
        raise ValueError("columns missing: " + ", ".join(miss))
    L = L[L["OBR Name"].notna()].copy()
    L["_ck"] = L["Country"].map(_lkc)
    L["_rk"] = L["OBR Name"].map(_rk)
    L["_d"] = pd.to_datetime(L["Round Start Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    L["_rn"] = L["Round Number"].map(_rnum)
    L["_vs"] = L["Vaccines"].map(_vset) if "Vaccines" in L.columns else [set() for _ in range(len(L))]
    return L.reset_index(drop=True)


def _round_offsets(pairs, lk):
    """Per Response: how the round numbers of the data differ from the lookup's (the lookup starts some
    campaigns at Round 0, the data at Rnd1). pairs: DataFrame with _rk, _d, rn (round number in the data)."""
    m = pairs.merge(lk[["_rk", "_d", "_rn"]].drop_duplicates(["_rk", "_d"]), on=["_rk", "_d"], how="inner")
    m = m[m["rn"].notna() & m["_rn"].notna()]
    if m.empty:
        return {}
    m["diff"] = m["rn"] - m["_rn"]
    return {k: int(g["diff"].mode().iloc[0]) for k, g in m.groupby("_rk")}


def lookup_crosscheck(df, lk):
    """Compare every Country / Response / SIA_date / Round_Add of the data with the lookup.
    Returns (one row per campaign round of the data with its status, lookup rounds that have no data)."""
    g = pd.DataFrame({"Country": df["Country"], "Response": df["Response"],
                      "_d": df["SIA_date"].map(_dkey), "Round_Add": df["Round_Add"].map(_s),
                      "Vaccine_type": df["Vaccine_type"].map(lambda x: x.strip() if isinstance(x, str) else "")})
    g["_rk"] = g["Response"].map(_rk)
    g["_ck"] = g["Country"].map(_lkc)
    g["rn"] = g["Round_Add"].map(_rnum)
    agg = (g.groupby(["Country", "Response", "_d", "Round_Add"], dropna=False)
           .agg(Records=("Country", "size"), Vaccine_type=("Vaccine_type", lambda x: ", ".join(sorted(set(x) - {""}))),
                _rk=("_rk", "first"), _ck=("_ck", "first"), rn=("rn", "first")).reset_index())
    off = _round_offsets(agg[["_rk", "_d", "rn"]].drop_duplicates(), lk)
    rows = []
    for r in agg.to_dict("records"):
        st, lc, lr, ld, days, lv = "OK", "", "", "", "", ""
        rk_, d_, rn_ = r["_rk"], r["_d"], r["rn"]
        e = lk[lk["_rk"] == rk_] if rk_ else lk.iloc[0:0]
        if not rk_:
            st = "No Response name"
        elif e.empty:
            st = "Response not in the lookup"
        else:
            lc = " / ".join(sorted(set(e["Country"].astype(str))))
            if r["_ck"] not in set(e["_ck"]):
                st = "Response listed for another country in the lookup"
            ex = e[e["_d"] == d_]
            if len(ex):
                x = ex.iloc[0]
                lr, ld, lv = x["Round Number"], x["Round Start Date"], x.get("Vaccines", "")
                if st == "OK" and rn_ is not None and not pd.isna(rn_) and not pd.isna(x["_rn"]) \
                        and rn_ != x["_rn"] + off.get(rk_, 0):
                    st = "Round differs from the lookup"
            else:
                try:
                    diff = (pd.to_datetime(e["_d"], errors="coerce") - pd.Timestamp(d_)).abs().dt.days
                    x = e.loc[diff.idxmin()]
                    days = int(diff.min())
                    lr, ld, lv = x["Round Number"], x["Round Start Date"], x.get("Vaccines", "")
                    if st == "OK":
                        st = ("Date differs from the lookup by %d day(s)" % days) if days <= 7 else \
                             "Date not in the lookup for this Response"
                except Exception:
                    if st == "OK":
                        st = "Date not readable"
        rows.append({"Country": r["Country"], "Response": r["Response"], "SIA_date": d_, "Round_Add": r["Round_Add"],
                     "Records": r["Records"], "Vaccine_type": r["Vaccine_type"], "Status": st, "Lookup_country": lc,
                     "Lookup_round": lr, "Lookup_start_date": str(ld)[:10], "Days_difference": days,
                     "Lookup_vaccines": lv})
    rep = pd.DataFrame(rows)
    rep["_o"] = (rep["Status"] != "OK").astype(int)
    rep = rep.sort_values(["_o", "Status", "Country", "Response", "SIA_date"], ascending=[False, True, True, True, True]) \
             .drop(columns="_o")
    ok = int(rep["Status"].eq("OK").sum())
    print("lookup: %d of %d campaign rounds of the data agree with the lookup; the others are in sheet Lookup_check" %
          (ok, len(rep)), flush=True)
    # lookup rounds with no records in the data (a start date within 7 days counts as present)
    have = {}
    for r in agg.to_dict("records"):
        have.setdefault(r["_rk"], []).append(r["_d"])
    miss = []
    for _, x in lk.iterrows():
        ds = have.get(x["_rk"], [])
        try:
            near = any(abs((pd.Timestamp(d) - pd.Timestamp(x["_d"])).days) <= 7 for d in ds if d)
        except Exception:
            near = False
        if not near:
            miss.append({"Country": x["Country"], "OBR Name": x["OBR Name"], "Round Number": x["Round Number"],
                         "Vaccines": x.get("Vaccines", ""), "Round Start Date": str(x["Round Start Date"])[:10],
                         "Round Status": x.get("Round Status", ""),
                         "Response_in_data": "yes (other dates)" if ds else "no"})
    nodata = pd.DataFrame(miss, columns=["Country", "OBR Name", "Round Number", "Vaccines", "Round Start Date",
                                         "Round Status", "Response_in_data"])
    return rep, nodata


def find_copied_rounds(df, lk):
    """Rounds that exist in the data but not in the lookup and only copy another round of the same Response and date.

    Example: CHD-2024-11-1_nOPV has Rnd1 and Rnd2 on the same date with the same totals, the lookup has Round 1
    only. The lookup is trusted: the extra round is removed. All conditions must hold:
      - the Response and the date are in the lookup for exactly one round (the real one);
      - the extra round is not a round of that Response in the lookup (otherwise it is a real round with a wrong
        date, which is only reported in Lookup_check);
      - the extra round has the same TotalNbrVaccPolio as the real round in at least 80% of the districts they share.
    Returns (mask of the rows to remove, log of them)."""
    g = pd.DataFrame({"_ck": df["Country"].map(_lkc), "_rk": df["Response"].map(_rk), "_d": df["SIA_date"].map(_dkey),
                      "rd": df["Round_Add"].map(_s)}, index=df.index)
    g["rn"] = g["rd"].map(_rnum)
    g = g[(g["_rk"] != "") & (g["_d"] != "") & g["rn"].notna()]
    nr = g.groupby(["_rk", "_d"])["rn"].transform("nunique")
    off = _round_offsets(g.loc[nr == 1, ["_rk", "_d", "rn"]].drop_duplicates(), lk)
    valid = {k: {int(x) + off.get(k, 0) for x in grp["_rn"].dropna()} for k, grp in lk.groupby("_rk")}
    mask = pd.Series(False, index=df.index)
    rows = []
    tot = pd.to_numeric(df["TotalNbrVaccPolio"], errors="coerce")
    key = (df["Province"].map(_nk) + "|" + df["District"].map(_nk) + "|" + df["Admin_3"].map(_nk))
    for (ck, rk_, d_), grp in g[nr > 1].groupby(["_ck", "_rk", "_d"]):
        e = lk[(lk["_rk"] == rk_) & (lk["_d"] == d_)]
        if len(e) != 1 or pd.isna(e.iloc[0]["_rn"]):
            continue
        real = int(e.iloc[0]["_rn"]) + off.get(rk_, 0)
        if real not in set(grp["rn"].astype(int)):
            continue
        base = grp[grp["rn"] == real]
        for extra_rn in sorted(set(grp["rn"].astype(int)) - {real}):
            if extra_rn in valid.get(rk_, set()):
                continue   # a real round of this Response: the date is wrong, not the round
            ex = grp[grp["rn"] == extra_rn]
            b_t = pd.Series(tot[base.index].values, index=key[base.index].values)
            e_t = pd.Series(tot[ex.index].values, index=key[ex.index].values)
            b_t, e_t = b_t[~b_t.index.duplicated()], e_t[~e_t.index.duplicated()]
            common = b_t.index.intersection(e_t.index)
            if len(common) < 5:
                continue
            same = float((b_t[common] == e_t[common]).mean())
            if same < 0.8:
                continue
            mask[ex.index] = True
            for i in ex.index:
                rows.append({"Original_row": int(df.at[i, "_row"]), "Country": df.at[i, "Country"],
                             "Province": df.at[i, "Province"], "District": df.at[i, "District"],
                             "SIA_date": d_, "Round_Add": df.at[i, "Round_Add"], "Response": df.at[i, "Response"],
                             "TotalNbrVaccPolio": df.at[i, "TotalNbrVaccPolio"],
                             "Reason": "the lookup has %s (%s) only; Rnd%d has the same total as Rnd%d in %.0f%% of the %d "
                                       "districts they share: a copy" % (e.iloc[0]["OBR Name"], e.iloc[0]["Round Number"],
                                                                       extra_rn, real, same * 100, len(common))})
    return mask, pd.DataFrame(rows)


def infer_campaigns(cur, pending, bad_resp, bad_vt, lk=None, footprint=True):
    """Right campaign for rows whose Response comes from another country and that no backup can restore.

    Two kinds of evidence, used together:
      the lookup (planned and implemented rounds: Response, round number, start date, vaccine, country): the one
        round of the same country and vaccine that started on the row's SIA_date gives the Response and the round;
      the footprint of the campaign in the data itself:
        block (5 rows or more with the same wrong Response, date and round): the Response of the same country
          whose districts are the same districts (at least 80% both ways) and the same vaccine;
        single row (or a few rows): the one campaign round of the same country, vaccine and province, within 90
          days, that covers at least 90% of the province's districts and lacks this district. The row is that
          district's missing record: its Response, SIA_date and Round_Add all come from that round, because the
          date and the round were copied from the other country together with the Response.
    When both exist and disagree nothing is changed. The round number comes from the lookup, shifted by the
    way the data numbers the rounds of that Response. Only a unique answer is used.
    Returns {row index: {set, why, confidence, check}} (or {unresolved: reason})."""
    out = {}
    if not pending:
        return out
    good = cur[~bad_resp & ~bad_vt].copy()
    good = good[good["Response"].map(lambda r: isinstance(r, str))]
    good["_d"] = good["SIA_date"].map(_dkey)
    good["_nd"] = good["District"].map(_nk)
    good["_np"] = good["Province"].map(_nk)
    good["_vt"] = good["Vaccine_type"].map(lambda x: x.strip() if isinstance(x, str) else "")
    good["_rd"] = good["Round_Add"].map(_s)
    good["_rk"] = good["Response"].map(_rk)
    spell = good.groupby("_rk")["Response"].agg(lambda x: x.value_counts().index[0]).to_dict()
    off = {}
    if lk is not None:
        pr = good[["_rk", "_d"]].copy()
        pr["rn"] = good["_rd"].map(_rnum)
        off = _round_offsets(pr.drop_duplicates(), lk)

    def lk_info(rk, d):
        """(label of the round in the data's numbering, lookup text) for a Response and a start date."""
        if lk is None:
            return None, ""
        e = lk[(lk["_rk"] == rk) & (lk["_d"] == d)]
        if len(e) != 1:
            return None, ""
        x = e.iloc[0]
        lab = ("Rnd%d" % (x["_rn"] + off.get(rk, 0))) if not pd.isna(x["_rn"]) else None
        return lab, "lookup: %s %s starts %s" % (x["OBR Name"], x["Round Number"], d)

    def lk_by_date(ck, d, vts):
        if lk is None or not d:
            return None
        e = lk[(lk["_ck"] == _lkc(ck)) & (lk["_d"] == d)]
        vs = set().union(*[_vset(v) for v in vts]) if vts else set()
        if vs:
            keep_ = [bool(v & vs) or not v for v in e["_vs"]]
            e = e.loc[keep_]
        rks = list(dict.fromkeys(e["_rk"]))
        return (rks[0], e[e["_rk"] == rks[0]].iloc[0]["OBR Name"]) if len(rks) == 1 else None

    groups = {}
    for i in pending:
        groups.setdefault((cur.at[i, "_ck"], cur.at[i, "Response"], _dkey(cur.at[i, "SIA_date"]),
                           _s(cur.at[i, "Round_Add"])), []).append(i)
    for (ck, found, dk, rd), idx in groups.items():
        gi = good[good["_ck"] == ck]
        vts = {x.strip() if isinstance(x, str) else "" for x in cur.loc[idx, "Vaccine_type"]}
        ldate = lk_by_date(ck, dk, vts)            # lookup: the round of this country that starts on this date
        if len(idx) >= 5 and not gi.empty:
            G = {_nk(x) for x in cur.loc[idx, "District"]}
            scored = []
            for resp, grp in gi.groupby("Response"):
                if not (set(grp["_vt"]) & vts):
                    continue
                C = set(grp["_nd"])
                inter = len(G & C)
                scored.append((min(inter / len(G), inter / max(len(C), 1)), resp, grp))
            scored.sort(key=lambda t: -t[0])
            fp = scored[0] if footprint and scored and scored[0][0] >= 0.8 and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.1) else None
            fp_rk = _rk(fp[1]) if fp else None
            if ldate and fp_rk and ldate[0] != fp_rk:
                for i in idx:
                    out[i] = {"unresolved": "the lookup says %s (round starting %s), the districts say %s: check with the country"
                                            % (ldate[1], dk, fp[1])}
                continue
            rk_ = ldate[0] if ldate else fp_rk
            if rk_ is None:
                continue
            resp = spell.get(rk_) or (ldate[1] if ldate else fp[1])
            lab, ltxt = lk_info(rk_, dk)
            st = {"Response": resp}
            if lab and lab != rd:
                st["Round_Add"] = lab
            why = []
            if fp:
                why.append("same country, same vaccine, same %d districts as %s (%.0f%% match)" % (len(G), resp, fp[0] * 100))
            if ltxt:
                why.append(ltxt)
            check = ""
            if not ltxt and fp:
                same_round = fp[2][(fp[2]["_rd"] == rd) & (fp[2]["_d"] != dk) & fp[2]["_nd"].isin(G)]
                if len(same_round):
                    check = ("this Response already has %s for the same districts on %s: the block (%s) is probably "
                             "the next round; Round_Add was not changed" % (rd, same_round["_d"].iloc[0], dk))
            for i in idx:
                out[i] = {"set": dict(st), "why": "; ".join(why), "check": check, "lookup": bool(ltxt),
                          "confidence": "high" if ltxt else ("medium" if check else "high")}
            continue
        if gi.empty:
            continue
        for i in idx:
            prov, dist = _nk(cur.at[i, "Province"]), _nk(cur.at[i, "District"])
            gp = gi[gi["_np"] == prov]
            n_p = gp["_nd"].nunique()
            cand = None
            try:
                d0 = pd.Timestamp(dk)
                if footprint and n_p and dist:
                    cands = []
                    for (resp, cd, cr), grp in gp.groupby(["Response", "_d", "_rd"]):
                        if not (set(grp["_vt"]) & vts) or not cd:
                            continue
                        if abs((pd.Timestamp(cd) - d0).days) > 90:
                            continue
                        covered = set(grp["_nd"])
                        if dist in covered or len(covered) / n_p < 0.9:
                            continue
                        cands.append((resp, cd, cr, len(covered), n_p, grp))
                    if len(cands) == 1:
                        cand = cands[0]
            except Exception:
                cand = None
            if ldate and cand and ldate[0] != _rk(cand[0]):
                out[i] = {"unresolved": "the lookup says %s (round starting %s), the province says %s: check with the country"
                                        % (ldate[1], dk, cand[0])}
                continue
            if cand:
                resp, cd, cr, k, n_p, grp = cand
                st = {"Response": resp, "SIA_date": cd, "Round_Add": cr}
                for extra in ("Year", "SIAMTH"):   # date-derived columns, only when they hold plain values
                    if extra in cur.columns:
                        vals = cur.loc[grp.index, extra].dropna()
                        mine = cur.at[i, extra]
                        if len(vals) and not str(vals.iloc[0]).startswith("=") and str(vals.iloc[0]) != str(mine) \
                                and not str(mine).startswith("="):
                            st[extra] = vals.iloc[0].item() if hasattr(vals.iloc[0], "item") else vals.iloc[0]
                lab, ltxt = lk_info(_rk(resp), cd)
                check = ""
                if lab and lab != cr:
                    check = "the lookup numbers this round %s, the data %s: Round_Add kept as in the data" % (lab, cr)
                why = ("%s %s (%s) covers %d of the %d districts of %s and is the only round within 90 days that lacks %s; "
                       "the row's date and round came with the wrong Response" % (resp, cr, cd, k, n_p, cur.at[i, "Province"], cur.at[i, "District"]))
                if ltxt:
                    why += "; " + ltxt
                out[i] = {"set": st, "why": why, "check": check, "lookup": bool(ltxt), "confidence": "high" if ltxt else "medium"}
            elif ldate:
                rk_ = ldate[0]
                lab, ltxt = lk_info(rk_, dk)
                st = {"Response": spell.get(rk_) or ldate[1]}
                if lab and lab != rd:
                    st["Round_Add"] = lab
                out[i] = {"set": st, "why": ltxt or "lookup", "check": "", "lookup": True, "confidence": "high"}
    return out


def repair_copy_errors(df, backup_files, src, cache_dir, infer=True, lk=None):
    """Find and fix the copy / drag-fill errors in Response and Vaccine_type, in place. Returns the repair log.

    Response: a name that starts with the code of another country (e.g. the South Sudan name copied onto DRC
    rows). Vaccine_type: nOPV3 ... nOPV15 (a drag-fill of nOPV2). Each is restored, in this order, from
      1. the same row in the newest earlier backup that holds a valid value,
      2. the other rows of the same country, date and round (majority),
      3. for Vaccine_type, what the Response name says (bOPV / nOPV); otherwise nOPV2.
    Errors that were already in the oldest backup cannot be restored that way. For a Response of another country
    the campaign is then inferred from the footprint (infer_campaigns): a block of rows takes the Response of the
    same country's campaign that covers the same districts; a single row takes the campaign round that the other
    districts of its province have and it lacks (the date and the round, copied from the other country together
    with the Response, are corrected too). These are marked INFERRED, with the reason; whatever is still unclear is
    listed as unresolved."""
    cur = _prep(df)
    snaps = []
    if backup_files:
        me = _md5(src)
        for p in backup_files[-8:]:
            try:
                if os.path.getsize(p) == os.path.getsize(src) and _md5(p) == me:
                    continue   # the backup of this very file
                snaps.append((os.path.basename(p), _prep(_load_backup(p, cache_dir))))
            except Exception as e:  # a damaged backup must not stop the update
                print("backup skipped (%s): %s" % (os.path.basename(p), e), flush=True)
    if snaps:
        print("repair: reference = %d earlier backup(s)" % len(snaps), flush=True)
    else:
        print("repair: no earlier backup found, the file itself is the only reference", flush=True)
    owners = _owners([s[1] for s in snaps] if snaps else [cur])

    def resp_ok(prefix, ck):
        return prefix is None or prefix not in owners or owners[prefix] == ck

    # the lookup says to which country each campaign name belongs: this also finds copies that are already in
    # every backup (then the code of the name is shared by two countries and the backups cannot tell the owner)
    lk_owner = lk.groupby("_rk")["_ck"].agg(set).to_dict() if lk is not None else {}

    def name_ok(r, ck):
        if not isinstance(r, str):
            return True
        o = lk_owner.get(_rk(r))
        return resp_ok(_prefix(r), ck) and (o is None or _lkc(ck) in o)

    cur["_resp_ok"] = [name_ok(r, c) for c, r in zip(cur["_ck"], cur["Response"])]
    bad_resp = ~cur["_resp_ok"]
    vt = cur["Vaccine_type"].map(lambda x: x.strip() if isinstance(x, str) else x)
    bad_vt = vt.map(lambda x: bool(isinstance(x, str) and KNOWN_BAD_NOPV.match(x) and KNOWN_BAD_NOPV.match(x).group(1) != "2"))
    log = []

    # one lookup table per snapshot: key -> value
    look = []
    for name, s in snaps:
        look.append((name, dict(zip(s["_kA"], zip(s["Response"], s["Vaccine_type"]))),
                     dict(zip(s["_kB"], zip(s["Response"], s["Vaccine_type"])))))

    def from_backups(i, which, valid):
        for name, la, lb in reversed(look):
            for tbl in (la, lb):
                v = tbl.get(cur.at[i, "_kA" if tbl is la else "_kB"])
                if v is not None:
                    cand = v[which]
                    if isinstance(cand, str) and valid(cand):
                        return cand, "backup " + name
        return None, None

    def peers(i, column, valid):
        ck, dk, rd = cur.at[i, "_ck"], _dkey(cur.at[i, "SIA_date"]), str(cur.at[i, "Round_Add"])
        d_ = cur["SIA_date"].map(_dkey)
        for sel in ((cur["_ck"] == ck) & (d_ == dk) & (cur["Round_Add"].astype(str) == rd),
                    (cur["_ck"] == ck) & (d_ == dk)):
            sel = sel & ~bad_resp & ~bad_vt
            vals = cur.loc[sel, column].dropna()
            vals = vals[vals.map(valid)]
            if len(vals) >= 3:
                top = vals.value_counts()
                if top.iloc[0] / len(vals) >= 0.8:
                    return top.index[0], "other rows of the same country, date and round"
        return None, None

    def rowinfo(i):
        return {"Excel row": int(df.at[i, "_row"]), "Country": df.at[i, "Country"], "Province": df.at[i, "Province"],
                "District": df.at[i, "District"], "SIA_date": _dkey(df.at[i, "SIA_date"])}

    # Response: 1. the lookup (planned and implemented rounds) is the primary source, 2. the backups, 3. other rows of
    # the same country, date and round, 4. the footprint of the campaign in the data (only if nothing else answers)
    bad_idx = list(df.index[bad_resp])
    lk_inf = infer_campaigns(cur, bad_idx, bad_resp, bad_vt, lk, footprint=infer) if (bad_idx and (lk is not None or infer)) else {}

    def apply_inferred(i, row, inf, label):
        for col, val in inf["set"].items():
            old = df.at[i, col] if col in df.columns else None
            if col == "SIA_date":
                df.at[i, col] = pd.Timestamp(val).to_pydatetime()
            else:
                df.at[i, col] = val
            r2 = dict(row if col == "Response" else rowinfo(i), Column=col,
                      Found=(_dkey(old) if col == "SIA_date" else old),
                      Rule=row["Rule"] if col == "Response" else "copied together with the Response from another country",
                      Restored=val, Source=label + inf["why"], Confidence=inf["confidence"],
                      Check=inf.get("check", ""))
            log.append(r2)

    for i in bad_idx:
        found = df.at[i, "Response"]
        ck = cur.at[i, "_ck"]
        valid = lambda r, ck=ck: name_ok(r, ck)
        row = dict(rowinfo(i), Column="Response", Found=found,
                   Rule="Response name belongs to another country" +
                        (" (found with the lookup)" if lk_owner.get(_rk(found)) and _prefix(found) not in owners else ""))
        inf = lk_inf.get(i)
        if inf and "set" in inf and inf.get("lookup"):
            apply_inferred(i, row, inf, "lookup: ")
            continue
        new, src_ = from_backups(i, 0, valid)
        if new is None:
            new, src_ = peers(i, "Response", valid)
        if new is not None:
            df.at[i, "Response"] = new
            row.update({"Restored": new, "Source": src_})
            log.append(row)
        elif inf and "set" in inf:
            apply_inferred(i, row, inf, "INFERRED: ")
        else:
            why_ = ("UNRESOLVED: " + inf["unresolved"]) if inf else \
                "UNRESOLVED: no lookup round, earlier version or other rows give the right name"
            row.update({"Restored": "", "Source": why_})
            log.append(row)

    # every restored Response is checked against the lookup, whatever source it came from
    if lk is not None:
        keyset = set(zip(lk["_ck"], lk["_rk"], lk["_d"]))
        for r in log:
            if r["Column"] == "Response" and r["Restored"] and not r.get("Lookup"):
                i = int(r["Excel row"]) - 2
                ok_ = (_lkc(cur.at[i, "_ck"]), _rk(r["Restored"]), _dkey(df.at[i, "SIA_date"])) in keyset
                r["Lookup"] = "confirmed: this round starts on that date" if ok_ else "not confirmed (round or date not in the lookup)"

    def vt_valid(v):
        return not (KNOWN_BAD_NOPV.match(v.strip()) and KNOWN_BAD_NOPV.match(v.strip()).group(1) != "2")

    for i in df.index[bad_vt]:
        found = df.at[i, "Vaccine_type"]
        new, src_ = from_backups(i, 1, vt_valid)
        if new is None:
            r = str(df.at[i, "Response"]).lower()
            if "bopv" in r and "nopv" not in r:
                new, src_ = "bOPV", "Response name"
            elif "nopv" in r:
                new, src_ = "nOPV2", "Response name"
            else:
                new, src_ = peers(i, "Vaccine_type", vt_valid)
                if new is None:
                    new, src_ = "nOPV2", "default (drag-fill of nOPV2)"
        df.at[i, "Vaccine_type"] = new
        log.append(dict(rowinfo(i), Column="Vaccine_type", Found=found, Rule="nOPV3 ... nOPV15: drag-fill of nOPV2",
                        Restored=new, Source=src_))
    fixed = sum(1 for r in log if r["Restored"])
    print("repair: %d copy/drag-fill errors found, %d restored, %d unresolved" % (len(log), fixed, len(log) - fixed),
          flush=True)
    return log


# =====================================================================================================
# Data dictionary (written at the end of the run, from the finished data)
# =====================================================================================================
VACC = [  # (suffix, long name)
    ("Polio", "polio"), ("Measles", "measles"), ("YF", "yellow fever"), ("Tetanus", "tetanus"),
    ("Mening", "meningitis"), ("VitA", "vitamin A"), ("Deparasit", "deworming"),
    ("Albendazole", "albendazole"), ("MII", "insecticide-treated nets (MII)"),
    ("AntMal", "antimalarial drugs"), ("Cholera", "cholera"), ("LLIN", "long-lasting insecticidal nets"),
    ("IPV", "IPV"), ("fIPV", "fractional IPV"), ("Other_A", "other intervention A"), ("Other_B", "other intervention B"),
]
GROUPS = {
    "id": "Identification", "integ": "Integration with other interventions", "pop": "Target population",
    "age": "Target age group", "polio": "Polio vaccination", "other": "Other interventions",
    "aefi": "Safety and evaluation", "campaign": "Campaign / response", "coadmin": "Co-administration (nOPV2 + bOPV)",
    "tech": "Technical / helper",
}
COLUMNS = {}


def add(name, group, desc, notes=""):
    COLUMNS[name] = (GROUPS[group], desc, notes)


add("Ctry_Code", "id", "Country code (3 letters, e.g. GUI, NIE, DRC).", "43 codes for 38 country names: 10 countries have two codes (e.g. DRC, Kenya, Tanzania).")
add("Country", "id", "Country name in capitals.", "Some countries appear under two names (CAMEROON / CAMEROUN, CENTRAL AFRICAN REPUBLIC / RÉPUBLIQUE CENTRAFRIQUE): group them before counting countries.")
add("Province", "id", "First administrative level below the country (province, region, state).", "Names are cleaned against the geographic reference; keep spelling as is when joining to the shapefiles.")
add("District", "id", "Health district / second administrative level in which the SIA took place. One row = one district in one campaign round and vaccine.", "Empty for a few records.")
add("Admin_3", "id", "Third administrative level (sub-district), when the data is reported at that level.", "Almost always empty (21 records).")
for suf, nm in VACC:
    add("Integ" + suf, "integ", "Whether %s was given together with the polio SIA (Yes / No)." % nm,
        "Values written in different ways (Yes, yes, YES, No, NO): compare case-insensitively." if suf in ("Polio", "Measles") else "")
add("Main_Act", "integ", "Main activity of the campaign: Polio, or Polio_MR_Integrated (polio with measles-rubella).", "'polio' also appears in lower case.")
add("TotPop", "pop", "Total population of the district used for planning the campaign.", "Can be smaller than the vaccinated total: see the coverage rules.")
for suf, nm in VACC:
    add("Pop" + suf, "pop", "Target population for %s in the district." % nm,
        "PopPolio is the denominator of CV_nOPV2 and CV_bOPV. Corrected where clearly wrong, see sheet Coverage rules." if suf == "Polio" else "")
for suf, nm in VACC:
    add("AgeG" + suf, "age", "Age group targeted for %s." % nm,
        "Polio: mostly 0-59m; also 0-10y, 0-14y, 0-119m. Written in several ways (0-119m / 0-119mth)." if suf == "Polio" else
        ("Contains 0 when measles was not part of the campaign." if suf == "Measles" else ""))
add("Nbr0dosesVaccPolio_0_11M", "polio", "Children aged 0-11 months vaccinated who had never received a polio dose (zero-dose children).")
add("Nbr_1doses_Plus_VacPolio_0_11M", "polio", "Children aged 0-11 months vaccinated who had already received 1 or more polio doses.")
add("Nbr0dosesVacPolio_12_59M", "polio", "Children aged 12-59 months vaccinated who had never received a polio dose (zero-dose children).")
add("Nbr_1doses_Plus_VacPolio_12_59M", "polio", "Children aged 12-59 months vaccinated who had already received 1 or more polio doses.")
add("Total_Nbr_0doseVaccPolio_0_59M", "polio", "Zero-dose children vaccinated, 0-59 months (sum of the two age groups).")
add("Prop_0dosesPolio", "polio", "Proportion of the vaccinated children who were zero-dose.", "Ratio, not %. A few values above 1 are data-entry errors.")
add("Total_Nbr_1dose_Plus_vaccPolio_0_59M", "polio", "Children with 1 or more previous doses vaccinated, 0-59 months (sum of the two age groups).")
add("TotalNbrVaccPolio", "polio", "Total number of children vaccinated with polio vaccine in the district and round, as reported. For a co-administered record (Vaccine_type 'nOPV2 & bOPV') it is the larger of Tot_Vacc_nOPV2 and Total_Vacc_bOPV, so a child vaccinated with both vaccines is counted once.",
    "Reported value; about 60% of records equal zero-dose + one-or-more-dose children. Not the sum of the two vaccines. Placed just before Response.")
add("CVPolio", "polio", "Polio vaccination coverage as stored in the source (ratio: 1.00 = 100%).", "Not always equal to TotalNbrVaccPolio / PopPolio. Used to correct PopPolio where it was clearly wrong (coverage above 1.5 or below 0.5). When it was empty it is filled with TotalNbrVaccPolio / PopPolio (target-only records stay empty). Some values are above 1.5.")
add("Doses.UsedPolio", "polio", "Polio vaccine doses used.", "Almost always empty.")
add("WastRPolio", "polio", "Polio vaccine wastage rate.", "Mostly empty.")
add("NbrVacMeasles", "other", "Children vaccinated against measles during the campaign.")
add("CVMeasles", "other", "Measles vaccination coverage (ratio).")
add("ZeroDoseMeasles", "other", "Zero-dose children among those vaccinated against measles.")
add("PropZeroDoseMeasles", "other", "Proportion of zero-dose children among those vaccinated against measles.")
add("UsedMeasles", "other", "Measles vaccine doses used.")
add("WastRMeasles", "other", "Measles vaccine wastage rate.")
add("NbrVacYF", "other", "Children vaccinated against yellow fever.")
add("CVYF", "other", "Yellow fever coverage (ratio).")
add("ZeroDoseYF", "other", "Zero-dose children among those vaccinated against yellow fever.")
add("PropZeroDoseYF", "other", "Proportion of zero-dose children, yellow fever.")
add("UsedYF", "other", "Yellow fever doses used.")
add("WastRYF", "other", "Yellow fever wastage rate.")
add("NbrVacTetanus_1st_Dose", "other", "People vaccinated with the 1st tetanus dose.")
add("NbrVacTetanus_2nd_Dose", "other", "People vaccinated with the 2nd tetanus dose.")
add("NbrVacTetanus_3rd_Dose", "other", "People vaccinated with the 3rd tetanus dose.")
add("TotalNbrVacTetanus", "other", "Total vaccinated against tetanus.")
add("CVTetanus", "other", "Tetanus coverage (ratio).")
add("UsedTetanus", "other", "Tetanus doses used.")
add("WastRTetanus", "other", "Tetanus wastage rate.")
for suf, nm, nbr in [("Mening", "meningitis", "NbrVacMening"), ("VitA", "vitamin A", "NbrVacVitA"), ("Deparasit", "deworming", "NbrVacDeparasit"),
                     ("Albendazole", "albendazole", "NbrVacAlbendazole"), ("MII", "insecticide-treated nets", "NbrVacMII"),
                     ("AntMal", "antimalarial drugs", "NbrVacAntMal")]:
    add(nbr, "other", "People reached with %s." % nm)
    add("CV" + suf, "other", "Coverage for %s (ratio)." % nm)
    add("Used" + suf, "other", "Quantity of %s used." % nm)
    add("WastR" + suf, "other", "Wastage rate, %s." % nm)
add("ZeroDoseCholera", "other", "Zero-dose children among those vaccinated against cholera.")
add("NbrVacCholera", "other", "People vaccinated against cholera.")
add("PropZeroDoseCholera", "other", "Proportion of zero-dose children, cholera.")
add("CVCholera", "other", "Cholera vaccination coverage (ratio).")
add("UsedCholera", "other", "Cholera doses used.")
add("WastRCholera", "other", "Cholera wastage rate.")
add("ZeroLLIN", "other", "Households without any net reached by the net distribution.")
add("NbrLLINDistributed", "other", "Long-lasting insecticidal nets distributed.")
add("PropZeroLLIN", "other", "Proportion of households without a net.")
add("CVLLIN", "other", "Net distribution coverage (ratio).")
add("UsedLLIN", "other", "Nets used.")
add("WastRLLIN", "other", "Net wastage rate.")
add("Total_IPV1_0_6wk", "other", "Children given the 1st IPV dose, 0-6 weeks.")
add("Total_IPV2_6_14wk", "other", "Children given the 2nd IPV dose, 6-14 weeks.")
add("TotalNbrVaccIPV", "other", "Total children vaccinated with IPV.")
add("CVIPV", "other", "IPV coverage (ratio).")
add("UsedIPV", "other", "IPV doses used.")
add("WastRIPV", "other", "IPV wastage rate.")
add("Total_fIPV", "other", "Children vaccinated with fractional IPV.")
add("CVfIPV", "other", "fIPV coverage (ratio).")
add("UsedfIPV", "other", "fIPV doses used.")
add("WastRfIPV", "other", "fIPV wastage rate.")
for x in ("A", "B"):
    add("NbrVacOther_" + x, "other", "People reached by other intervention %s." % x)
    add("CVOther_" + x, "other", "Coverage of other intervention %s (ratio)." % x)
    add("UsedOther_" + x, "other", "Quantity used, other intervention %s." % x)
    add("WastROther_" + x, "other", "Wastage rate, other intervention %s." % x)
add("AEFIMin", "aefi", "Minor adverse events following immunisation (AEFI) reported.", "Mostly empty.")
add("AEFISerious", "aefi", "Serious adverse events following immunisation reported.", "Mostly empty (58 records).")
add("Death", "aefi", "Deaths reported after vaccination.", "Empty.")
add("NumberMonitored", "aefi", "Children checked in the post-campaign evaluation (monitoring).", "Empty. Independent monitoring data is in the IM repository.")
add("NumberVaccinated_Evaluation", "aefi", "Children found vaccinated in the post-campaign evaluation.", "Empty.")
add("Prop_Missed_Eval", "aefi", "Proportion of children found missed in the post-campaign evaluation.", "Empty.")
add("SIA_date", "campaign", "Start date of the SIA round in the district.", "Date.")
add("Date_Report", "campaign", "Date the report was received.", "Date. Filled for about a quarter of the records.")
add("SIA_Integrated", "campaign", "Whether the SIA was integrated with another intervention (Yes / No).", "Written as Yes / No / NO.")
add("TypeofActivity", "campaign", "Type of activity: Reactive, Preventive, SIA, NIDs, Outbreak Response, Polio_MR_Integrated, ...", "Filled for about a quarter of the records. 'Poilo_MR_Integrated' is a spelling error for Polio_MR_Integrated.")
add("EpidNumber", "campaign", "EPID number of the outbreak case or environmental sample that triggered an outbreak response.", "Filled only for response campaigns (about 470 records).")
add("DateOfOnset", "campaign", "Date of onset of the case that triggered the response.", "Stored in mixed formats (dates and Excel serial numbers saved as text, for example 45828).")
add("Year", "campaign", "Year of the campaign (year of SIA_date).")
add("Round_Add", "campaign", "Campaign round: Rnd0 to Rnd6.", "Round of the response, not of the calendar year.")
add("Tot_Vacc_nOPV2", "coadmin", "Children vaccinated with nOPV2 in the district and round.", "Empty when no nOPV2 was given. Combined records: same value as TotalNbrVaccPolio unless the country gave a separate figure.")
add("CV_nOPV2", "coadmin", "nOPV2 coverage = Tot_Vacc_nOPV2 / PopPolio (ratio: 1.00 = 100%).", "Above 1.5 means the reported population is probably too low. Such records are listed for the countries.")
add("Total_Vacc_bOPV", "coadmin", "Children vaccinated with bOPV in the district and round.", "Empty when no bOPV was given. Combined records: same value as TotalNbrVaccPolio unless the country gave a separate figure.")
add("CV_bOPV", "coadmin", "bOPV coverage = Total_Vacc_bOPV / PopPolio (ratio).", "Same caution as CV_nOPV2.")
add("Vaccine_type", "coadmin", "Vaccine used: nOPV2, bOPV, mOPV2, MR, or 'nOPV2 & bOPV' when both were given in the same campaign (co-administration).", "See sheet Vaccine_type. Placed just before Response.")
add("Response", "campaign", "Name of the response / campaign, e.g. GUI-2025-01-NID_bOPV-nOPV (country code, year, number, type, vaccine).", "Empty for about 1,900 older records. A campaign is co-administered when the name contains both 'nOPV' and 'bOPV', or when a district has an nOPV2 and a bOPV record for the same Response, year and round (some DRC and South Sudan responses do not name both vaccines).")
add("Index", "tech", "Running record number inside the data-entry batch of a campaign (1 = first row).", "Internal; no analytical meaning.")
add("SIAMTH", "tech", "Month of the SIA (1-12).", "Same as the month of SIA_date for 98% of records.")
add("check_dp", "tech", "Helper key used to find duplicated records: formula joining Country, Province, District, SIA_date, TotalNbrVaccPolio and Response.", "Internal. A live formula in the admin data; it follows the columns, so keep it when columns are added.")
add("AFRO_block", "tech", "Programme block of the country: WA (West Africa), LCB (Lake Chad Basin), ESA (East and Southern Africa), DRC, CA (Central Africa).", "'LBC' appears for 25 records of the Central African Republic: it should be LCB.")
add("prov_candidates", "tech", "Helper column used when the province names were repaired.", "Empty.")

VACCINE_TYPES = [
    ("nOPV2", "Novel oral polio vaccine type 2 given alone. One record per district and round."),
    ("bOPV", "Bivalent oral polio vaccine (types 1 and 3) given alone. One record per district and round."),
    ("nOPV2 & bOPV", "Co-administration: nOPV2 and bOPV in the same campaign. One record per district and round, with both vaccine totals."),
    ("nOPV2_bOPV_IPV", "Co-administration of nOPV2, bOPV and IPV (Nigeria, Kebbi, 2026). Treated as a combined record."),
    ("mOPV2", "Monovalent oral polio vaccine type 2."),
    ("MR", "Measles-rubella vaccine given in an integrated polio / MR campaign. TotalNbrVaccPolio is 0 on these records; use NbrVacMeasles."),
]

COADMIN = [
    ("How to recognise a co-administered campaign", "The Response name contains both 'nOPV' and 'bOPV', or the district has an nOPV2 record and a bOPV record for the same Response, year and round, or Vaccine_type already holds both vaccines."),
    ("Reported as separate records", "An nOPV2 record and a bOPV record of the same district, round and Response are merged into one record: Tot_Vacc_nOPV2 and Total_Vacc_bOPV take the two totals, Vaccine_type = 'nOPV2 & bOPV'."),
    ("Reported as one combined record", "Tot_Vacc_nOPV2 = Total_Vacc_bOPV = TotalNbrVaccPolio by default (the countries are asked to confirm the real split)."),
    ("TotalNbrVaccPolio of a co-administered record", "The larger of Tot_Vacc_nOPV2 and Total_Vacc_bOPV (the same children are vaccinated with both, so the totals are not added)."),
    ("Not co-administered", "Vaccine_type unchanged. TotalNbrVaccPolio goes in the column of the vaccine used (nOPV2 or bOPV); the other column stays empty."),
    ("Single record in a co-administered Response", "Where other districts report separate nOPV2 and bOPV records but a district has only nOPV2, it is kept as nOPV2 only (no bOPV was reported there)."),
    ("Coverage", "CV_nOPV2 = Tot_Vacc_nOPV2 / PopPolio and CV_bOPV = Total_Vacc_bOPV / PopPolio."),
    ("Vaccine_type clean-up", "'Nopv2' and 'nOPV3' to 'nOPV14' (drag-fill errors) were corrected to 'nOPV2'; 'bopv' to 'bOPV'."),
]

QUALITY = [
    ("Coverage above 1.5", "More children vaccinated than the target population: the population is probably too low. Records are listed for the countries to correct."),
    ("PopPolio from CVPolio", "Where vaccinated / PopPolio is above 1.5 and the stored CVPolio is plausible, PopPolio = vaccinated / CVPolio (whole number). Also when the coverage is above 5 and CVPolio is less than half of it."),
    ("PopPolio from CVPolio (coverage below 0.5)", "Where vaccinated / PopPolio is below 0.5 and the stored CVPolio is plausible (0.5 to 1.5), PopPolio = vaccinated / CVPolio (whole number). In these records PopPolio was usually the total population or a province figure repeated in every district."),
    ("PopPolio from the district's other rounds", "Where coverage is still above 1.5, PopPolio = median of the same district's other polio records in the same year, when this gives a coverage between 0.5 and 1.5."),
    ("PopPolio from a model", "Where it is still above 1.5, PopPolio is estimated by a model trained on the district's own history (typical error about 10%), when this gives a coverage between 0.5 and 1.5. These are estimates, not reported data; countries are asked to confirm them."),
    ("Whole numbers", "PopPolio is always a whole number of children."),
    ("Records still above 1.5", "Left as reported (no reliable history). Listed for the countries."),
    ("Audit", "Every changed record is listed in the audit workbook produced with the data (sheet PopPolio_corrected, with the old value and the reason)."),
]

CODES = [
    ("Round_Add", "Rnd0 to Rnd6", "Campaign round."),
    ("AFRO_block", "WA", "West Africa"), ("AFRO_block", "LCB", "Lake Chad Basin"), ("AFRO_block", "ESA", "East and Southern Africa"),
    ("AFRO_block", "DRC", "Democratic Republic of the Congo"), ("AFRO_block", "CA", "Central Africa"),
    ("Integ* flags", "Yes / No", "Written in several ways (Yes, yes, YES, No, NO): compare case-insensitively."),
    ("Coverage columns", "ratio", "1.00 = 100%. Values are not capped."),
    ("Dates", "date", "SIA_date, Date_Report, DateOfOnset."),
]


def kind(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return "Date"
    if pd.api.types.is_numeric_dtype(s):
        return "Number"
    return "Text"


def build_dictionary(df, dst):
    """Write the data dictionary workbook for the admin data in `df` (completeness, ranges and examples are
    computed from the data; the descriptions are in COLUMNS above)."""
    print("building the data dictionary", flush=True)
    missing = [c for c in df.columns if c not in COLUMNS]
    if missing:
        print("note: no description yet for these columns (marked in the dictionary; add them to COLUMNS in this "
              "script): " + ", ".join(missing))
        for c in missing:
            COLUMNS[c] = ("Other", "No description yet.", "Column not known to the script: description to be added.")
    extra = [c for c in COLUMNS if c not in df.columns]
    if extra:
        print("note: described but not in the data:", ", ".join(extra))

    rows = []
    for j, c in enumerate(df.columns, start=1):
        s = df[c]
        grp, desc, notes = COLUMNS[c]
        k = kind(s)
        nn = s.notna().sum()
        ex = s.dropna().iloc[0] if nn else ""
        if isinstance(ex, pd.Timestamp):
            ex = ex.strftime("%Y-%m-%d")
        rng = ""
        if k == "Number" and nn:
            rng = "%s to %s" % (format(s.min(), ",.4g"), format(s.max(), ",.10g"))
        elif k == "Text" and nn and s.nunique() <= 12:
            rng = " | ".join(map(str, s.value_counts().index[:12]))
        elif k == "Date" and nn:
            rng = "%s to %s" % (s.min().strftime("%Y-%m-%d"), s.max().strftime("%Y-%m-%d"))
        rows.append({"No": j, "Column": c, "Excel column": L(j), "Group": grp, "Description": desc, "Type": k,
                     "Filled (%)": round(100 * nn / len(df), 1), "Distinct values": int(s.nunique()),
                     "Range / values": rng, "Example": ex, "Notes / known issues": notes})
    dd = pd.DataFrame(rows)

    vt = df["Vaccine_type"].value_counts().rename_axis("Vaccine_type").reset_index(name="Records")
    meaning = dict(VACCINE_TYPES)
    vt["Meaning"] = vt["Vaccine_type"].map(meaning).fillna("")
    vt = pd.concat([vt, pd.DataFrame([{"Vaccine_type": k, "Records": 0, "Meaning": v} for k, v in VACCINE_TYPES
                                      if k not in set(vt["Vaccine_type"])])], ignore_index=True)

    readme = pd.DataFrame({"": [
        "Data dictionary of AFRO_admin_data.xlsx (WHO AFRO administrative vaccination data, polio SIAs).",
        "One row of the data = one district, one campaign round and one vaccine (or one combined nOPV2 & bOPV record).",
        "Records: {:,}   |   Columns: {}   |   Countries: {}   |   Years: {} to {}".format(
            len(df), df.shape[1], df["Country"].nunique(), int(df["Year"].min()), int(df["Year"].max())),
        "Sheets: Data dictionary (every column), Vaccine_type (values), Co-administration (how the new columns are filled), "
        "Coverage rules (data-quality corrections), Codes.",
        "Co-administration columns Tot_Vacc_nOPV2, CV_nOPV2, Total_Vacc_bOPV, CV_bOPV, Vaccine_type, TotalNbrVaccPolio and CVPolio sit just before Response.",
        "Coverage is a ratio (1.00 = 100%). Coverage above 1.5 is not credible and is under review with the countries.",
        "Built by add_coadmin_columns.py from the data file; completeness and ranges are computed from the data.",
    ]})

    with pd.ExcelWriter(dst, engine="openpyxl") as xw:
        readme.to_excel(xw, sheet_name="Read me", index=False)
        dd.to_excel(xw, sheet_name="Data dictionary", index=False)
        vt.to_excel(xw, sheet_name="Vaccine_type", index=False)
        pd.DataFrame(COADMIN, columns=["Topic", "Rule"]).to_excel(xw, sheet_name="Co-administration", index=False)
        pd.DataFrame(QUALITY, columns=["Topic", "Rule"]).to_excel(xw, sheet_name="Coverage rules", index=False)
        pd.DataFrame(CODES, columns=["Column", "Value", "Meaning"]).to_excel(xw, sheet_name="Codes", index=False)
        widths = {"Read me": [150], "Data dictionary": [5, 34, 9, 24, 70, 8, 9, 10, 34, 24, 60],
                  "Vaccine_type": [20, 10, 90], "Co-administration": [42, 120], "Coverage rules": [38, 120], "Codes": [20, 20, 70]}
        for name, ws in xw.sheets.items():
            ws.sheet_view.showGridLines = False
            for j, w in enumerate(widths[name], start=1):
                ws.column_dimensions[L(j)].width = w
            if name != "Read me":
                for c in ws[1]:
                    c.font = Font(bold=True, color="FFFFFF")
                    c.fill = PatternFill("solid", fgColor="005C97")
                    c.alignment = Alignment(vertical="center", wrap_text=True)
                ws.freeze_panes = "C2" if name == "Data dictionary" else "A2"
                ws.auto_filter.ref = ws.dimensions
            for row in ws.iter_rows(min_row=1 if name == "Read me" else 2):
                for c in row:
                    c.alignment = Alignment(vertical="top", wrap_text=True)
        # emphasise the co-administration group
        ws = xw.sheets["Data dictionary"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 4).value == GROUPS["coadmin"]:
                for j in range(1, 12):
                    ws.cell(r, j).fill = PatternFill("solid", fgColor="E8F3FA")
    print("written:", dst, "|", len(dd), "columns described")


if __name__ == "__main__":
    main()
