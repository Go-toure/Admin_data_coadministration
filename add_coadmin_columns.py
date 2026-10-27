#!/usr/bin/env python3
"""
add_coadmin_columns.py  --  co-administration columns for AFRO_admin_data.xlsx
=============================================================================

Why
---
Some co-administered SIAs (nOPV2 + bOPV) are reported in three different ways
(feedback from Valentina Bigoni):

  * separate vaccine-specific records  - one nOPV2 record + one bOPV record
  * combined records                   - one record, Vaccine_type "nOPV2 & bOPV"
  * mixed practices                    - both, in different campaigns / rounds

A campaign is co-administered when its Response name contains both "nOPV" and
"bOPV".  This script makes every record carry the same structure, so
co-administration can be reported correctly:

    ... Year | Round_Add | Tot_Vacc_nOPV2 | CV_nOPV2 | Total_Vacc_bOPV | CV_bOPV |
        Vaccine_type | TotalNbrVaccPolio | Response | ...

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

The script is safe to run again on its own output.

Usage
-----
  python add_coadmin_columns.py INPUT.xlsx OUTPUT.xlsx [--audit AUDIT.xlsx]
"""
import argparse
import re
import sys
from copy import copy

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

SHEET = "Sheet1"
SUMMARY_SHEET = "Children Reached"
BLOCK = ["Tot_Vacc_nOPV2", "CV_nOPV2", "Total_Vacc_bOPV", "CV_bOPV",
         "Vaccine_type", "TotalNbrVaccPolio"]
BEFORE = "Response"
KEY = ["Country", "Province", "District", "Admin_3", "Year", "Round_Add", "Response"]
COMBINED = "nOPV2 & bOPV"
CV_HIGH = 1.5
CV_EXTREME = 5.0   # coverage above this is not credible even when CVPolio is also above CV_HIGH

A_NONE = "Not co-administered"
A_COMBINED = "Combined record (as reported)"
A_MERGED = "Separate vaccine records merged into one"
A_RESP = "Co-administered by Response (single record)"
A_KEPT = "Single-vaccine record kept (no bOPV record in a co-administered Response)"


def num(s):
    return pd.to_numeric(s, errors="coerce")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--audit", help="write an audit workbook here")
    ap.add_argument("--model-poppolio", action="store_true",
                    help="after --fix-poppolio: estimate PopPolio of the records still above %.1f with a model "
                         "trained on the district's own history (needs scikit-learn)" % CV_HIGH)
    ap.add_argument("--fix-poppolio", action="store_true",
                    help="correct PopPolio where it is clearly wrong: when vaccinated/PopPolio is above %.1f "
                         "but the stored CVPolio is plausible (<= %.1f), PopPolio is replaced by "
                         "TotalNbrVaccPolio / CVPolio, rounded to a whole number; any other decimal PopPolio is "
                         "rounded to a whole number too" % (CV_HIGH, CV_HIGH))
    a = ap.parse_args()

    print("loading", a.src, flush=True)
    wb = openpyxl.load_workbook(a.src)
    ws = wb[SHEET]
    header = []
    for c in ws[1]:
        header.append(c.value)
    while header and header[-1] is None:
        header.pop()
    ncol = len(header)
    old_letter = {h: get_column_letter(i + 1) for i, h in enumerate(header)}
    for need in ["Country", "Response", "Vaccine_type", "TotalNbrVaccPolio", "PopPolio", "Round_Add", "Year"]:
        if need not in header:
            sys.exit("missing column: " + need)

    cells = list(ws.iter_rows(min_row=2, max_col=ncol))
    cells = [r for r in cells if any(c.value is not None for c in r)]
    n = len(cells)
    df = pd.DataFrame([[c.value for c in r] for r in cells], columns=header, dtype=object)
    df["_row"] = np.arange(2, n + 2)
    print("rows:", n, "cols:", ncol, flush=True)

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
    df["_key"] = key_str
    occ = pd.Series(0, index=df.index)
    occ[cand] = df[cand].groupby(["_key", vt_norm[cand]]).cumcount()
    pk = key_str + "#" + occ.astype(str)

    n_set = set(pk[cand & is_n]); b_set = set(pk[cand & is_b])
    paired_keys = n_set & b_set
    paired = cand & pk.isin(paired_keys)
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
        out_total[base] = np.nanmax([t_n, t_b])
        out_vt[base] = COMBINED
        drop[other] = True
        merged_log.append((df.at[base, "_row"], df.at[other, "_row"]))

    # optional correction of clearly wrong target populations (see --fix-poppolio)
    pop_fixed = pd.Series(False, index=df.index)
    pop_reason = pd.Series("", index=df.index, dtype=object)
    pop_new = pop.copy()
    if a.fix_poppolio:
        cvp = num(df["CVPolio"])
        with np.errstate(divide="ignore", invalid="ignore"):
            cv_now = out_total / pop.where(pop > 0)
            implied = (out_total / cvp.where(cvp > 0)).round()
        cond1 = (cv_now > CV_HIGH) & (cvp > 0) & (cvp <= CV_HIGH)
        # extreme coverage while CVPolio is less than half of it: the stored CVPolio used a much better population
        cond2 = (cv_now > CV_EXTREME) & (cvp > 0) & (cvp < cv_now / 2)
        pop_fixed = (cond1 | cond2) & implied.notna() & (implied > 0) & ~drop
        pop_new[pop_fixed] = implied[pop_fixed]
        pop_reason[pop_fixed] = "Population implied by CVPolio (vaccinated / CVPolio)"
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
                if h == "PopPolio" and bool(pop_fixed.iat[i]):
                    c = ws_new.cell(row=r_out, column=j + 1, value=int(pop.iat[i]))
                    c._style = copy(oc._style)
                    continue
                if oc.value is None:
                    continue
                c = ws_new.cell(row=r_out, column=j + 1, value=oc.value)
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

    # ----------------------------------------------------------------- audit
    if a.audit:
        write_audit(a.audit, df, approach, vt_orig, out_vt, vt_changed, total, out_total, out_tn, out_tb,
                    drop, merged_log, keep, paired, pop_fixed, pop_before, pop, num(df["CVPolio"]), pop_reason, out_cvn, out_cvb, model_info)

    # ------------------------------------------------------------- self-check
    chk = pd.read_excel(a.dst, sheet_name=SHEET)
    assert len(chk) == int(keep.sum()), "row count"
    assert list(chk.columns[:0]) == []
    assert chk.columns.tolist() == new_header, "column order"
    pos = chk.columns.get_loc(BEFORE)
    assert chk.columns[pos - 6:pos].tolist() == BLOCK, "block must sit right before Response"
    co = chk["Vaccine_type"].astype(str).str.lower().str.contains("nopv") & chk["Vaccine_type"].astype(str).str.lower().str.contains("bopv")
    mx = chk[["Tot_Vacc_nOPV2", "Total_Vacc_bOPV"]].max(axis=1)
    assert (chk.loc[co, "TotalNbrVaccPolio"] == mx[co]).all(), "co-admin total must be the max"
    print("self-check OK: %d rows, co-administered records: %d" % (len(chk), int(co.sum())))
    print(approach.value_counts().to_string())


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


def write_audit(path, df, approach, vt_orig, out_vt, vt_changed, total, out_total, out_tn, out_tb,
                drop, merged_log, keep, paired, pop_fixed=None, pop_before=None, pop_after=None, cvp=None,
                pop_reason=None, cv_n=None, cv_b=None, model_info=None):
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


if __name__ == "__main__":
    main()
