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

    # copy / drag-fill errors (Response from another country, nOPV3..nOPV15): restored from the backups
    repairs = []
    repl = {}
    if not a.no_repair:
        repairs = repair_copy_errors(df, find_backups(a.backups, a.src), a.src, os.path.join(a.out_dir, "_backup_cache"),
                                     infer=not a.no_infer)
        for r in repairs:
            if r["Restored"] and r["Column"] in ("Response", "Round_Add", "Year", "SIAMTH"):
                repl[(int(r["Excel row"]) - 2, r["Column"])] = r["Restored"]
            elif r["Restored"] and r["Column"] == "SIA_date":
                repl[(int(r["Excel row"]) - 2, "SIA_date")] = pd.Timestamp(r["Restored"]).to_pydatetime()

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
    cand = (is_n | is_b) & resp_both & ~vt_both
    # a district that has an nOPV2 record AND a bOPV record for the same Response, year and round is
    # co-administered even when the Response name does not mention both vaccines (DRC, South Sudan)
    cand_any = (is_n | is_b) & ~vt_both
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

    with np.errstate(divide="ignore", invalid="ignore"):
        pop_ok = pop.where(pop > 0)
        out_cvn = out_tn / pop_ok
        out_cvb = out_tb / pop_ok
    # no vaccinated total yet (target only): no coverage either
    no_total = out_total.isna() & (out_tn.fillna(0) == 0) & (out_tb.fillna(0) == 0)
    out_cvn[no_total] = np.nan
    out_cvb[no_total] = np.nan

    keep = ~drop
    print("pairs merged:", int(drop.sum()), "-> rows out:", int(keep.sum()), flush=True)

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

    # a vaccine that was not used has a total of 0, not a blank (the convention now used in the SharePoint file)
    vals = {
        "Tot_Vacc_nOPV2": out_tn.fillna(0), "CV_nOPV2": out_cvn,
        "Total_Vacc_bOPV": out_tb.fillna(0), "CV_bOPV": out_cvb,
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
    if a.audit:
        write_audit(a.audit, df, approach, vt_orig, out_vt, vt_changed, total, out_total, out_tn, out_tb,
                    drop, merged_log, keep, paired, pop_fixed, pop_before, pop, cvp_all, pop_reason, out_cvn, out_cvb, model_info, repairs)

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
    if partial:
        print("WARNING: only %d countries in this file: do NOT upload the output as AFRO_admin_data.xlsx." % n_countries)
    else:
        print("Next: open AFRO_admin_data.xlsx in Excel and save once, then upload it to SharePoint.")


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
                pop_reason=None, cv_n=None, cv_b=None, model_info=None, repairs=None):
    cols = ["_row", "Country", "Province", "District", "Admin_3", "Year", "Round_Add", "Response", "SIA_date"]
    a = df[cols].copy()
    a.rename(columns={"_row": "Original_row"}, inplace=True)
    a["Approach"] = approach
    a["Vaccine_type_before"] = vt_orig
    a["Vaccine_type_after"] = out_vt.where(keep, "(merged into the other record)")
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


def infer_campaigns(cur, pending, bad_resp, bad_vt):
    """Right campaign for rows whose Response comes from another country and that no backup can restore.

    The evidence is the footprint of the campaign, nothing else is used:
      block (5 rows or more with the same wrong Response, date and round): the Response of the same country
        whose districts are the same districts (at least 80% both ways) and the same vaccine. Date and round
        are the block's own; a warning is added when the round was already used by that Response on the same
        districts (then it is probably the next round) - this is not changed.
      single row (or a few rows): one campaign round of the same country, vaccine and province, within 90 days
        of the row, that covers at least 90% of the province's districts and does not have this district. The
        row is that district's missing record: its Response, SIA_date and Round_Add are all taken from it,
        because the date and the round were copied from the other country together with the Response.
    Only a unique answer is used. Returns {row index: {set, why, confidence, check}}."""
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
    groups = {}
    for i in pending:
        groups.setdefault((cur.at[i, "_ck"], cur.at[i, "Response"], _dkey(cur.at[i, "SIA_date"]),
                           _s(cur.at[i, "Round_Add"])), []).append(i)
    for (ck, found, dk, rd), idx in groups.items():
        gi = good[good["_ck"] == ck]
        if gi.empty:
            continue
        vts = {x.strip() if isinstance(x, str) else "" for x in cur.loc[idx, "Vaccine_type"]}
        if len(idx) >= 5:
            G = {_nk(x) for x in cur.loc[idx, "District"]}
            scored = []
            for resp, grp in gi.groupby("Response"):
                if not (set(grp["_vt"]) & vts):
                    continue
                C = set(grp["_nd"])
                inter = len(G & C)
                scored.append((min(inter / len(G), inter / max(len(C), 1)), resp, grp))
            scored.sort(key=lambda t: -t[0])
            if scored and scored[0][0] >= 0.8 and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.1):
                score, resp, grp = scored[0]
                same_round = grp[(grp["_rd"] == rd) & (grp["_d"] != dk) & grp["_nd"].isin(G)]
                check = ""
                if len(same_round):
                    check = ("this Response already has %s for the same districts on %s: the block (%s) is probably "
                             "the next round; Round_Add was not changed" % (rd, same_round["_d"].iloc[0], dk))
                for i in idx:
                    out[i] = {"set": {"Response": resp},
                              "why": "same country, same vaccine, same %d districts as %s (%.0f%% match)"
                                     % (len(G), resp, score * 100),
                              "confidence": "medium" if check else "high", "check": check}
            continue
        for i in idx:
            prov, dist = _nk(cur.at[i, "Province"]), _nk(cur.at[i, "District"])
            gp = gi[gi["_np"] == prov]
            n_p = gp["_nd"].nunique()
            try:
                d0 = pd.Timestamp(dk)
            except Exception:
                continue
            if not n_p or not dist:
                continue
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
                resp, cd, cr, k, n_p, grp = cands[0]
                st = {"Response": resp, "SIA_date": cd, "Round_Add": cr}
                for extra in ("Year", "SIAMTH"):   # date-derived columns, only when they hold plain values
                    if extra in cur.columns:
                        vals = cur.loc[grp.index, extra].dropna()
                        mine = cur.at[i, extra]
                        if len(vals) and not str(vals.iloc[0]).startswith("=") and str(vals.iloc[0]) != str(mine) \
                                and not str(mine).startswith("="):
                            st[extra] = vals.iloc[0].item() if hasattr(vals.iloc[0], "item") else vals.iloc[0]
                out[i] = {"set": st, "confidence": "medium",
                          "why": "%s %s (%s) covers %d of the %d districts of %s and is the only round within 90 days "
                                 "that lacks %s; the row's date and round came with the wrong Response"
                                 % (resp, cr, cd, k, n_p, cur.at[i, "Province"], cur.at[i, "District"])}
    return out


def repair_copy_errors(df, backup_files, src, cache_dir, infer=True):
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

    cur["_resp_ok"] = [resp_ok(p, c) if isinstance(r, str) else True
                       for p, c, r in zip(cur["_pre"], cur["_ck"], cur["Response"])]
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

    pending = []
    for i in df.index[bad_resp]:
        found = df.at[i, "Response"]
        ck = cur.at[i, "_ck"]
        valid = lambda r, ck=ck: resp_ok(_prefix(r), ck)
        new, src_ = from_backups(i, 0, valid)
        if new is None:
            new, src_ = peers(i, "Response", valid)
        row = dict(rowinfo(i), Column="Response", Found=found, Rule="Response name belongs to another country")
        if new is None:
            pending.append((i, row))
        else:
            df.at[i, "Response"] = new
            row.update({"Restored": new, "Source": src_})
            log.append(row)

    inferred = infer_campaigns(cur, [i for i, _ in pending], bad_resp, bad_vt) if (infer and pending) else {}
    for i, row in pending:
        inf = inferred.get(i)
        if inf is None:
            row.update({"Restored": "", "Source": "UNRESOLVED: no earlier version or other rows give the right name"})
            log.append(row)
            continue
        for col, val in inf["set"].items():
            old = df.at[i, col] if col in df.columns else None
            if col == "SIA_date":
                df.at[i, col] = pd.Timestamp(val).to_pydatetime()
            else:
                df.at[i, col] = val
            r2 = dict(row if col == "Response" else rowinfo(i), Column=col,
                      Found=(_dkey(old) if col == "SIA_date" else old),
                      Rule=row["Rule"] if col == "Response" else "copied together with the Response from another country",
                      Restored=val, Source="INFERRED: " + inf["why"], Confidence=inf["confidence"],
                      Check=inf.get("check", ""))
            log.append(r2)

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
add("CVPolio", "polio", "Polio vaccination coverage as stored in the source (ratio: 1.00 = 100%).", "Not always equal to TotalNbrVaccPolio / PopPolio. Used to correct PopPolio where it was clearly wrong (coverage above 1.5 or below 0.5). Some values are above 1.5.")
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
add("Tot_Vacc_nOPV2", "coadmin", "Children vaccinated with nOPV2 in the district and round.", "0 when no nOPV2 was given. Combined records: same value as TotalNbrVaccPolio unless the country gave a separate figure.")
add("CV_nOPV2", "coadmin", "nOPV2 coverage = Tot_Vacc_nOPV2 / PopPolio (ratio: 1.00 = 100%).", "Above 1.5 means the reported population is probably too low. Such records are listed for the countries.")
add("Total_Vacc_bOPV", "coadmin", "Children vaccinated with bOPV in the district and round.", "0 when no bOPV was given. Combined records: same value as TotalNbrVaccPolio unless the country gave a separate figure.")
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
