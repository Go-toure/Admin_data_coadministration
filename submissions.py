"""Submissions intake: the files that countries send (one SharePoint folder) are inspected, normalised and compared with the live
admin data, and the ones an operator accepts are added to the next build.

    python admin.py submissions inspect [--folder DIR]   read every file of the folder, write the report, the prepared records and the ledger
    python admin.py submissions list                     what the ledger knows
    python admin.py submissions accept NAME [--note ..]  the file goes into the next build (never one that is blocked)
    python admin.py submissions reject NAME [--note ..]

What the inspection does for each file
  1. layout: the template (the 164 columns of the live file), the template with other spellings of the headers, a file whose headers can
     be mapped by a short list of certain synonyms, or a file that cannot be used (it is rejected with the reason);
  2. normalisation (each change is listed): capitals and spaces, country and province spelling of the geography reference, vaccine and round
     spelling of the live file, dates (text, Excel serial numbers), numbers written as text, percentages, Year / SIAMTH / Ctry_Code / AFRO_block,
     the values the live file always has for the columns a country does not report, an empty Response filled from the campaign lookup;
  3. the records are checked: key fields, impossible dates and numbers, the portal quality checks (with the live records of the same
     countries as context), values never seen in the live file, the campaigns against the lookup, duplicates inside the file;
  4. every record is compared with the live file: new / already there with the same figures / already there with other figures
     (those are only listed, they are never changed by a submission).
Nothing is written to the live file here: the builder adds the accepted new records (add_coadmin_columns.py --add-records)."""
import datetime as dt
import hashlib
import json
import re
import unicodedata
from collections import Counter
from pathlib import Path

import numpy as np
import pandas as pd

KEY_COLS = ["Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response"]
FIGURE_COLS = ["TotPop", "PopPolio", "TotalNbrVaccPolio", "Tot_Vacc_nOPV2", "Total_Vacc_bOPV", "CVPolio"]
TEXT_COLS = ["Country", "Province", "District", "Admin_3", "Response", "Vaccine_type", "Round_Add"]
DATE_COLS = ["SIA_date", "Date_Report", "DateOfOnset"]
CAT_COLS = ["Vaccine_type", "Round_Add", "AgeGPolio", "Main_Act"]
NO_DEFAULT = set(KEY_COLS + FIGURE_COLS + ["Admin_3", "Date_Report", "DateOfOnset", "Year", "SIAMTH", "Index", "check_dp",
                                            "prov_candidates", "Ctry_Code", "AFRO_block", "CV_nOPV2", "CV_bOPV", "EpidNumber"])
EXTENSIONS = (".xlsx", ".xlsm", ".csv")
EMPTY_WORDS = {"", "-", "--", "na", "n/a", "nan", "null", "none", "#n/a", "nd", "n.d.", "#div/0!"}
# headers that mean the same as a template column (compared without case, accents, spaces, punctuation). Only what is certain.
SYNONYMS = {"round": "Round_Add", "rnd": "Round_Add", "roundnumber": "Round_Add", "rounds": "Round_Add",
            "vaccine": "Vaccine_type", "vaccinetype": "Vaccine_type", "vaccineused": "Vaccine_type",
            "campaign": "Response", "campaignname": "Response", "responsename": "Response", "obrname": "Response",
            "admin1": "Province", "adm1": "Province", "admin2": "District", "adm2": "District",
            "startdate": "SIA_date", "campaignstartdate": "SIA_date", "datestart": "SIA_date", "siastartdate": "SIA_date", "campaignstart": "SIA_date", "startofcampaign": "SIA_date", "campaigndate": "SIA_date",
            "targetpopulation": "PopPolio", "targetpop": "PopPolio", "totalpopulation": "TotPop", "totalpop": "TotPop",
            "totalvaccinated": "TotalNbrVaccPolio", "childrenvaccinated": "TotalNbrVaccPolio", "vaccinated": "TotalNbrVaccPolio",
            "coverage": "CVPolio", "admincoverage": "CVPolio", "reportdate": "Date_Report", "datereport": "Date_Report",
            "countrycode": "Ctry_Code", "iso3": "Ctry_Code"}


# ----------------------------------------------------------------------------------------------- helpers
def cn(x):
    """A header for comparison: no accents, case, spaces or punctuation (SIA-date = SIA_date = sia date)."""
    t = unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", t)


def sha256_of(path, n=1 << 20):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(n)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def list_files(folder):
    out = []
    for p in sorted(Path(folder).rglob("*")):
        if p.is_file() and p.suffix.lower() in EXTENSIONS and not p.name.startswith(("~$", ".")):
            out.append(p)
    return out


def is_empty(v):
    return v is None or (isinstance(v, float) and np.isnan(v)) or (isinstance(v, str) and v.strip().lower() in EMPTY_WORDS) \
        or v is pd.NaT


def num_value(v, pct_ok=False):
    """A cell as a number: text with spaces, thousands separators, a decimal comma or a % sign is understood. Returns (value, how) with
    how = '' (already a number), a short description of what was done, or None when it is not a number."""
    if is_empty(v):
        return np.nan, ""
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        return float(v), ""
    s = str(v).strip().replace(" ", " ")
    pct = s.endswith("%")
    t = s.rstrip("%").strip().replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", t):
        x, how = float(t.replace(",", "")), "thousands separator removed"
    elif re.fullmatch(r"-?\d+,\d{1,2}", t):
        x, how = float(t.replace(",", ".")), "decimal comma"
    else:
        try:
            x, how = float(t), "number written as text"
        except ValueError:
            return np.nan, None
    if pct:
        if not pct_ok:
            return x, "percent sign kept as it is"
        x, how = x / 100.0, "percent written as a fraction"
    return x, how


def excel_serial(x):
    try:
        x = float(x)
    except (TypeError, ValueError):
        return pd.NaT
    if 20000 <= x <= 80000:
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=x)
    return pd.NaT


def to_text(df, num_cols=()):
    """A typed frame as text, the way admin.load_norm gives the live file (dates as YYYY-MM-DD, numbers without trailing zeros)."""
    cols = {}
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            cols[c] = s.dt.strftime("%Y-%m-%d").fillna("")
        elif c in num_cols or pd.api.types.is_numeric_dtype(s):
            n = pd.to_numeric(s, errors="coerce")
            cols[c] = n.map(lambda x: format(x, ".10g") if x == x else "")
        else:
            cols[c] = s.map(lambda x: "" if is_empty(x) else str(x))
    return pd.DataFrame(cols, index=df.index)


# ----------------------------------------------------------------------------------------------- what the live file teaches
def learn_profile(live, builder, geo):
    """Everything a submission is judged against, learnt from the live file (text frame)."""
    prof = {"columns": list(live.columns), "values": {}, "numeric": [], "upper": [], "country_map": {}, "defaults": {}, "country_defaults": {}}
    for c in CAT_COLS:
        if c in live.columns:
            prof["values"][c] = set(v for v in live[c].astype(str).str.strip().unique() if v)
    for c in live.columns:
        if c in TEXT_COLS or c in DATE_COLS or c in ("check_dp", "prov_candidates"):
            continue
        s = live[c].replace("", np.nan).dropna()
        if len(s) and pd.to_numeric(s, errors="coerce").notna().mean() >= 0.95:
            prof["numeric"].append(c)
    for c in ("Country", "Province", "District"):
        s = live[c].replace("", np.nan).dropna().astype(str)
        if len(s) and (s == s.str.upper()).mean() >= 0.95:
            prof["upper"].append(c)
    ck = live["Country"].map(builder._qcountry)
    for c in ("Ctry_Code", "AFRO_block"):
        if c in live.columns:
            d = live[c].replace("", np.nan).groupby(ck).agg(lambda x: x.dropna().value_counts().index[0] if x.notna().any() else None)
            prof["country_map"][c] = {k: v for k, v in d.items() if v}
    for c in live.columns:
        if c in NO_DEFAULT:
            continue
        s = live[c].replace("", np.nan)
        if s.notna().mean() < 0.98:
            continue
        vc = s.dropna().value_counts()
        if len(vc) and vc.iloc[0] / vc.sum() >= 0.9:
            prof["defaults"][c] = vc.index[0]
        else:
            per = {}
            for k, g in s.groupby(ck):
                v = g.dropna().value_counts()
                if len(v) and v.iloc[0] / v.sum() >= 0.9:
                    per[k] = v.index[0]
            if per:
                prof["country_defaults"][c] = per
    return prof


# ----------------------------------------------------------------------------------------------- layout
def read_head(path, nrows=25):
    """{sheet: DataFrame of the first rows without header}."""
    if str(path).lower().endswith(".csv"):
        return {"(csv)": _read_csv(path, nrows)}
    xl = pd.ExcelFile(path)
    out = {}
    for sn in xl.sheet_names:
        try:
            out[sn] = pd.read_excel(xl, sheet_name=sn, header=None, nrows=nrows, dtype=object)
        except Exception:
            continue
    return out


def _read_csv(path, nrows=None):
    last = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return pd.read_csv(path, header=None, dtype=object, sep=None, engine="python", encoding=enc, nrows=nrows)
        except Exception as e:
            last = e
    raise last


def detect_layout(path, template_cols):
    """Find the sheet and the header row, and map the headers to the template. Returns a dict (see keys)."""
    tn = {cn(c): c for c in template_cols}
    best = None
    for sn, raw in read_head(path).items():
        for r in range(min(len(raw), 25)):
            vals = [v for v in raw.iloc[r].tolist() if not is_empty(v)]
            hits = sum(1 for v in vals if cn(v) in tn or cn(v) in SYNONYMS)
            if best is None or hits > best[0]:
                best = (hits, sn, r)
    lay = {"sheet": None, "header_row": None, "map": {}, "synonym": [], "ignored": [], "missing_keys": list(KEY_COLS), "class": "not usable",
           "template_columns_found": 0, "notes": []}
    if best is None or best[0] < 3:
        lay["notes"].append("no row with column names of the template was found in the first 25 rows of any sheet")
        return lay
    _, sn, r = best
    lay["sheet"], lay["header_row"] = sn, r
    raw = read_head(path)[sn]
    seen = {}
    for j, v in enumerate(raw.iloc[r].tolist()):
        if is_empty(v):
            continue
        k = cn(v)
        tgt, certain = (tn[k], True) if k in tn else ((SYNONYMS[k], False) if k in SYNONYMS else (None, None))
        if tgt is None:
            lay["ignored"].append(str(v))
        elif tgt in seen:
            lay["ignored"].append("%s (second column for %s)" % (v, tgt))
            lay["notes"].append("two columns for %s: %s and %s; the first is used" % (tgt, seen[tgt], v))
        else:
            seen[tgt] = str(v)
            lay["map"][j] = tgt
            if not certain:
                lay["synonym"].append("%s -> %s" % (v, tgt))
    lay["missing_keys"] = [k for k in KEY_COLS if k not in seen]
    lay["template_columns_found"] = len(seen)
    if lay["missing_keys"]:
        lay["class"] = "not usable"
        lay["notes"].append("key columns missing: " + ", ".join(lay["missing_keys"]))
    elif lay["synonym"]:
        lay["class"] = "mapped (headers differ)"
    elif len(seen) >= 0.9 * len(template_cols):
        lay["class"] = "template"
    else:
        lay["class"] = "template columns, fewer than the template"
    return lay


def read_records(path, lay):
    if str(path).lower().endswith(".csv"):
        raw = _read_csv(path)
    else:
        raw = pd.read_excel(path, sheet_name=lay["sheet"], header=None, dtype=object)
    r = lay["header_row"]
    body = raw.iloc[r + 1:].copy()
    body["_src_row"] = np.arange(r + 2, r + 2 + len(body))     # the row in the file as Excel shows it
    keep = {j: t for j, t in lay["map"].items()}
    out = pd.DataFrame({t: body.iloc[:, j].values for j, t in keep.items()}, index=body.index)
    out["_src_row"] = body["_src_row"].values
    return out.reset_index(drop=True)


# ----------------------------------------------------------------------------------------------- one file
class Notes:
    def __init__(self, cap=30000):
        self.rows, self.counts, self.cap = [], Counter(), cap

    def add(self, src_row, col, found, new, rule):
        self.counts[(rule, col)] += 1
        if len(self.rows) < self.cap:
            self.rows.append({"Row_in_file": src_row, "Column": col, "Found": found, "Set_to": new, "Rule": rule})


def prepare_file(path, ctx, prof, template_cols):
    """Inspect one file. Returns (summary dict, prepared DataFrame in template layout or None, detail dict of frames)."""
    b = ctx.builder
    name = Path(path).name
    s = {"name": name, "size": Path(path).stat().st_size, "hash": sha256_of(path), "layout": "", "sheet": "", "header_row": "",
         "rows_read": 0, "rows_used": 0, "excluded": 0, "new": 0, "existing_same": 0, "existing_different": 0, "existing_units": 0, "in_file_duplicates_removed": 0,
         "qc_errors": 0, "qc_warnings": 0, "countries": "", "status": "blocked", "why": ""}
    det = {"layout": None, "notes": None, "excluded": [], "issues": None, "compare": None, "diffs": None, "campaigns": None, "dups": None}
    try:
        lay = detect_layout(path, template_cols)
    except Exception as e:
        s["why"] = "the file could not be opened (%s)" % type(e).__name__
        det["layout"] = {"class": "not readable", "map": {}, "synonym": [], "ignored": [], "missing_keys": [], "notes": [str(e)[:200]]}
        return s, None, det
    det["layout"] = lay
    s["layout"], s["sheet"], s["header_row"] = lay["class"], lay["sheet"], (lay["header_row"] + 1) if lay["header_row"] is not None else ""
    if lay["class"] == "not usable":
        s["why"] = "; ".join(lay["notes"])
        return s, None, det
    raw = read_records(path, lay)
    s["rows_read"] = len(raw)
    notes = Notes()
    P = pd.DataFrame({c: raw[c] if c in raw.columns else None for c in template_cols}, dtype=object)
    P["_src_row"] = raw["_src_row"].values
    # rows with nothing in the key and figure columns are empty lines; a "TOTAL" line is not a record
    core = [c for c in KEY_COLS + ["PopPolio", "TotalNbrVaccPolio"] if c in raw.columns]
    empty_line = raw[core].map(is_empty).all(axis=1).values
    totals = np.zeros(len(P), bool)
    for c in ("Country", "Province", "District"):
        totals |= P[c].map(lambda v: isinstance(v, str) and bool(re.match(r"^\s*(grand\s+)?total\b", v, re.I))).values
    for i in np.where(totals & ~empty_line)[0]:
        det["excluded"].append({"Row_in_file": int(P["_src_row"].iat[i]), "Reason": "a total line, not a record"})
    P = P[~empty_line & ~totals].reset_index(drop=True)
    # text
    for c in TEXT_COLS:
        for i, v in P[c].items():
            if is_empty(v):
                P.at[i, c] = None
                continue
            t = re.sub(r"\s+", " ", re.sub(r"_x000D_", " ", str(v))).strip()
            if t != v:
                notes.add(P["_src_row"].iat[i], c, v, t, "spaces / control characters")
            P.at[i, c] = t
    # country and province: the geography reference's spelling
    if ctx.geo is not None:
        std = ctx.geo.drop_duplicates("c_key").set_index("c_key")["country"].to_dict()
        for i, v in P["Country"].items():
            if v is None:
                continue
            t = std.get(b._qcountry(v))
            if t and t != v:
                notes.add(P["_src_row"].iat[i], "Country", v, t, "country name of the geography reference")
                P.at[i, "Country"] = t
    for c in prof["upper"]:
        for i, v in P[c].items():
            if isinstance(v, str) and v != v.upper():
                notes.add(P["_src_row"].iat[i], c, v, v.upper(), "capitals, like the live file")
                P.at[i, c] = v.upper()
    if ctx.geo is not None:
        try:
            for i, (old, new) in b.fix_province_spelling(P, ctx.geo).items():
                notes.add(P["_src_row"].iat[i], "Province", old, new, "province spelling of the geography reference")
                P.at[i, "Province"] = new
        except Exception:
            pass
    # vaccine and round, in the spelling of the live file
    vmap = {cn(v): v for v in prof["values"].get("Vaccine_type", ())}
    for i, v in P["Vaccine_type"].items():
        if isinstance(v, str):
            t = vmap.get(cn(v))
            if t and t != v:
                notes.add(P["_src_row"].iat[i], "Vaccine_type", v, t, "vaccine spelling of the live file")
                P.at[i, "Vaccine_type"] = t
    for i, v in P["Round_Add"].items():
        if v is None:
            continue
        m = re.fullmatch(r"(?i)\s*(?:rnd|round|r)?\s*0*(\d+)\.?0*\s*", str(v))
        if m:
            t = "Rnd%d" % int(m.group(1))
            if t != v:
                notes.add(P["_src_row"].iat[i], "Round_Add", v, t, "round written as Rnd<number>")
                P.at[i, "Round_Add"] = t
    # dates
    for c in DATE_COLS:
        cache = {}
        vals = []
        for i, v in P[c].items():
            if is_empty(v):
                vals.append(pd.NaT)
                continue
            key = repr(v)
            if key not in cache:
                if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
                    cache[key] = excel_serial(v)
                else:
                    cache[key] = b.parse_sia_text(v)
            t = cache[key]
            if not isinstance(v, (pd.Timestamp, dt.datetime, dt.date)) and t is not pd.NaT and not pd.isna(t):
                notes.add(P["_src_row"].iat[i], c, v, t.strftime("%Y-%m-%d"), "date written as text or as a number")
            vals.append(t)
        P[c] = pd.to_datetime(pd.Series(vals, index=P.index), errors="coerce")
    # numbers
    bad_num = []
    for c in prof["numeric"]:
        if c not in P.columns:
            continue
        pct_ok = c.startswith("CV") or c == "CVPolio"
        out = []
        for i, v in P[c].items():
            x, how = num_value(v, pct_ok)
            if how is None:
                bad_num.append({"Row_in_file": int(P["_src_row"].iat[i]), "Column": c, "Value": v})
                x = np.nan
            elif how and how != "percent sign kept as it is":
                notes.add(P["_src_row"].iat[i], c, v, x, how)
            elif how:
                notes.add(P["_src_row"].iat[i], c, v, x, how)
            out.append(x)
        P[c] = pd.Series(out, index=P.index, dtype=float)
    det["bad_numbers"] = pd.DataFrame(bad_num)
    # an empty Response: the lookup (the same rules as the build), then the names of the lookup
    if ctx.lookup is not None:
        try:
            for i, inf in b.fill_blank_response(P, ctx.lookup).items():
                P.at[i, "Response"] = inf["set"]["Response"]
                notes.add(P["_src_row"].iat[i], "Response", "", inf["set"]["Response"], inf["why"])
            for i, inf in b.align_response_names(P, ctx.lookup, None).items():
                old = P.at[i, "Response"]
                P.at[i, "Response"] = inf["set"]["Response"]
                notes.add(P["_src_row"].iat[i], "Response", old, inf["set"]["Response"], inf["why"])
        except Exception as e:
            det.setdefault("warnings", []).append("Response filling from the lookup failed: %s" % e)
    # rows that cannot be used
    bad = []
    for i in P.index:
        miss = [c for c in KEY_COLS if (pd.isna(P.at[i, c]) if c == "SIA_date" else is_empty(P.at[i, c]))]
        if miss:
            bad.append((i, "key field empty or unreadable: " + ", ".join(miss)))
    for i, why in bad:
        det["excluded"].append({"Row_in_file": int(P["_src_row"].iat[i]), "Reason": why})
    P = P.drop(index=[i for i, _ in bad]).reset_index(drop=True)
    # duplicates inside the file: exact copies are dropped
    if len(P):
        T = to_text(P.drop(columns="_src_row"), prof["numeric"])
        cd = b.classify_duplicates(T)
        ex = cd["cls"] == "exact"
        drop = ex & (cd.groupby(["key", "sig"]).cumcount() > 0)
        det["dups"] = pd.DataFrame({"Row_in_file": P["_src_row"].values, "Class": cd["cls"].values, "Records_with_this_key": cd["n"].values})
        det["dups"] = det["dups"][det["dups"]["Class"] != ""]
        s["in_file_duplicates_removed"] = int(drop.sum())
        for i in np.where(drop.values)[0]:
            det["excluded"].append({"Row_in_file": int(P["_src_row"].iat[i]), "Reason": "exact copy of another record of the file"})
        P = P[~drop.values].reset_index(drop=True)
    s["rows_used"] = len(P)
    s["excluded"] = s["rows_read"] - len(P)
    if not len(P):
        s["why"] = "no usable record"
        det["notes"] = pd.DataFrame(notes.rows)
        return s, None, det
    s["countries"] = ", ".join(sorted(P["Country"].dropna().unique()))
    # comparison with the live file (before the defaults are filled in: only what the country reported is compared)
    cmp_, diffs = compare_with_live(P, ctx, prof)
    det["compare"], det["diffs"] = cmp_, diffs
    s["new"] = int((cmp_["Status"] == "new").sum())
    s["existing_same"] = int((cmp_["Status"] == "already in the live file, same figures").sum())
    s["existing_different"] = int((cmp_["Status"] == "already in the live file, other figures").sum())
    s["existing_units"] = int(cmp_["Status"].str.startswith("same names as several").sum())
    # derived columns and the live file's constants
    yr = P["SIA_date"].dt.year.astype(float)
    mo = P["SIA_date"].dt.month.astype(float)
    for col, ser in (("Year", yr), ("SIAMTH", mo)):
        if col in P.columns:
            cur = pd.to_numeric(P[col], errors="coerce")
            for i in np.where((cur != ser).values)[0]:
                notes.add(P["_src_row"].iat[i], col, None if pd.isna(cur.iat[i]) else cur.iat[i], ser.iat[i], "derived from SIA_date")
            P[col] = ser
    ckeys = P["Country"].map(b._qcountry)
    for col, mp in prof["country_map"].items():
        if col in P.columns:
            for i in P.index:
                v = mp.get(ckeys.iat[i])
                if not v:
                    continue
                cur = P.at[i, col]
                if is_empty(cur):
                    P.at[i, col] = v
                    notes.add(P["_src_row"].iat[i], col, "", v, "country code / block of the live file")
                elif str(cur).strip() != str(v):
                    P.at[i, col] = v
                    notes.add(P["_src_row"].iat[i], col, cur, v, "code of another country replaced by the code of the live file")
    filled = Counter()
    for col in template_cols:
        if col not in prof["defaults"] and col not in prof["country_defaults"]:
            continue
        for i in P.index[P[col].map(is_empty)]:
            v = prof["defaults"].get(col) or prof["country_defaults"].get(col, {}).get(ckeys.iat[i])
            if v is not None:
                P.at[i, col] = float(v) if col in prof["numeric"] else v
                filled[col] += 1
    det["defaults"] = pd.DataFrame([{"Column": c, "Records_filled": n, "Value": prof["defaults"].get(c, "(per country)")} for c, n in filled.items()])
    det["notes"] = pd.DataFrame(notes.rows)
    det["note_counts"] = pd.DataFrame([{"Rule": r, "Column": c, "Cells": n} for (r, c), n in notes.counts.most_common()])
    # values never seen in the live file
    unseen = []
    for c in CAT_COLS:
        if c in P.columns and prof["values"].get(c):
            for v, k in P[c].dropna().astype(str).value_counts().items():
                if v not in prof["values"][c]:
                    unseen.append({"Column": c, "Value": v, "Records": int(k), "Message": "never seen in the live file"})
    det["unseen"] = pd.DataFrame(unseen)
    # the portal quality checks on the records, with the live records of the same countries as context
    det["issues"] = quality_issues(P, ctx, prof)
    iss = det["issues"]
    s["qc_errors"] = int(iss.loc[iss["severity"] == "Error", "Excel_row"].nunique()) if len(iss) else 0
    s["qc_warnings"] = int(iss.loc[iss["severity"] == "Warning", "Excel_row"].nunique()) if len(iss) else 0
    # campaigns against the lookup
    if ctx.lookup is not None:
        try:
            rep, _ = b.lookup_crosscheck(P, ctx.lookup)
            det["campaigns"] = rep
        except Exception:
            det["campaigns"] = None
    # the verdict
    reasons = []
    if s["excluded"]:
        reasons.append("%d row(s) left out" % s["excluded"])
    if s["qc_errors"]:
        reasons.append("%d record(s) with an Error" % s["qc_errors"])
    if s["existing_different"]:
        reasons.append("%d record(s) already in the live file with other figures (not changed)" % s["existing_different"])
    if s["existing_units"]:
        reasons.append("%d record(s) with the same names as several live records and none identical (not added)" % s["existing_units"])
    if lay["class"] != "template":
        reasons.append("layout: " + lay["class"])
    if len(det["unseen"]):
        reasons.append("values never seen in the live file")
    if det["bad_numbers"] is not None and len(det["bad_numbers"]):
        reasons.append("%d cell(s) that are not numbers" % len(det["bad_numbers"]))
    if s["qc_warnings"] > 0.25 * s["rows_used"]:
        reasons.append("more than 25% of the records have a warning")
    if s["new"] == 0:
        reasons.append("no new record")
    s["status"] = "ready" if not reasons else "review"
    s["why"] = "; ".join(reasons)
    P = P.drop(columns="_src_row")
    P = P[template_cols]
    return s, P, det


def compare_with_live(P, ctx, prof):
    b = ctx.builder
    T = to_text(P.drop(columns="_src_row"), prof["numeric"])
    key, blank = b.dup_key(T)
    rows, diffs = [], []
    cols = [c for c in FIGURE_COLS + ["Date_Report"] if c in P.columns]
    for i in range(len(P)):
        k = key.iat[i]
        cand = ctx.live_index.get(k)
        if not cand:
            st, note = "new", ""
        else:
            best, best_d = None, None
            for j in cand:
                d = []
                for c in cols:
                    a, bv = P.at[i, c], ctx.live.at[j, c] if c in ctx.live.columns else ""
                    if c == "Date_Report":
                        a_s = "" if pd.isna(a) else pd.Timestamp(a).strftime("%Y-%m-%d")
                        b_s = str(bv)[:10]
                        if a_s and a_s != b_s:
                            d.append((c, bv, a_s))
                        continue
                    x = pd.to_numeric(pd.Series([bv]), errors="coerce").iat[0]
                    if pd.isna(a):
                        continue                      # the country did not report it: nothing to compare
                    if pd.isna(x) or abs(float(a) - float(x)) > 0.5 * (0.01 if c == "CVPolio" else 1):
                        d.append((c, "" if pd.isna(x) else x, a))
                if best_d is None or len(d) < len(best_d):
                    best, best_d = j, d
            if not best_d:
                st, note = "already in the live file, same figures", "live row %d" % (best + 2)
            elif len(cand) > 1:
                st, note = "same names as several live records, none identical (not added)", "live rows " + ", ".join(str(j + 2) for j in cand[:6])
            else:
                st, note = "already in the live file, other figures", "live row %d" % (best + 2)
            if best_d and len(cand) == 1:
                for c, lv, sv in best_d:
                    diffs.append({"Row_in_file": int(P["_src_row"].iat[i]), "Country": P.at[i, "Country"], "District": P.at[i, "District"],
                                  "SIA_date": P.at[i, "SIA_date"].strftime("%Y-%m-%d"), "Response": P.at[i, "Response"],
                                  "Column": c, "Live_value": lv, "Submitted_value": sv, "Live_row": best + 2})
        rows.append({"Row_in_file": int(P["_src_row"].iat[i]), "Country": P.at[i, "Country"], "Province": P.at[i, "Province"],
                     "District": P.at[i, "District"], "SIA_date": P.at[i, "SIA_date"].strftime("%Y-%m-%d"),
                     "Round_Add": P.at[i, "Round_Add"], "Vaccine_type": P.at[i, "Vaccine_type"], "Response": P.at[i, "Response"],
                     "Status": st, "Note": note})
    return pd.DataFrame(rows), pd.DataFrame(diffs)


def quality_issues(P, ctx, prof):
    b = ctx.builder
    ctry = set(P["Country"].dropna().unique())
    live_ctx = ctx.live[ctx.live["Country"].isin(ctry)].copy()
    live_ctx = typed(live_ctx, prof)
    q = P.drop(columns="_src_row")
    try:
        res = b.qc_run(q, ctx.geo, b.qc_config(), context=live_ctx if len(live_ctx) else None)
    except Exception as e:
        return pd.DataFrame([{"severity": "Warning", "check": "inspection", "Country": "", "message": "quality checks could not run: %s" % e,
                              "Excel_row": ""}])
    iss = res["issues"].copy()
    iss["Excel_row"] = iss["idx"].map(lambda i: int(P["_src_row"].iat[i]) if i < len(P) else "")
    iss = iss[~iss["check"].isin(["duplicate_key", "same_key_other_figures"])]
    return iss[["severity", "check", "check_label", "Country", "Province", "District", "SIA_date", "Round", "Vaccine", "Response",
                "message", "current", "suggested", "Excel_row"]].reset_index(drop=True)


def typed(live_text, prof):
    t = live_text.copy()
    for c in prof["numeric"]:
        if c in t.columns:
            t[c] = pd.to_numeric(t[c], errors="coerce")
    for c in DATE_COLS:
        if c in t.columns:
            t[c] = pd.to_datetime(t[c], errors="coerce")
    return t


# ----------------------------------------------------------------------------------------------- the folder
def inspect_folder(ctx, folder, out_root, template_cols, log=print):
    """Inspect every file of a folder. Writes out_root/Submissions_report.xlsx, out_root/prepared/<hash>.xlsx and updates the ledger."""
    out_root = Path(out_root)
    (out_root / "prepared").mkdir(parents=True, exist_ok=True)
    ledger = load_ledger(out_root)
    files = list_files(folder)
    log("%d file(s) in %s" % (len(files), folder))
    prof = learn_profile(ctx.live, ctx.builder, ctx.geo)
    rows, details = [], {}
    seen_hashes = set()
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    for p in files:
        h = sha256_of(p)
        seen_hashes.add(h)
        old = ledger["files"].get(h)
        log("  inspecting %s" % p.name)
        s, P, det = prepare_file(p, ctx, prof, template_cols)
        s["path"] = str(Path(p).relative_to(folder)).replace("\\", "/")
        prepared = None
        if P is not None:
            prepared = "prepared/%s.xlsx" % h[:12]
            P.to_excel(out_root / prepared, sheet_name="Sheet1", index=False)
        e = old or {"decision": None, "decided_by": "", "decided_at": "", "note": "", "first_seen": now}
        e.update({k: s[k] for k in ("name", "path", "size", "hash", "layout", "rows_read", "rows_used", "excluded", "new", "existing_same",
                                    "existing_different", "existing_units", "qc_errors", "qc_warnings", "countries", "status", "why")})
        e.update({"prepared": prepared, "last_inspected": now, "present": True})
        ledger["files"][h] = e
        rows.append(dict(s, decision=e["decision"] or "", decided_by=e["decided_by"]))
        details[h] = (s, det)
        log("    %s: %s | %d rows, %d new, %d already there (%d with other figures)%s" %
            (s["status"].upper(), s["layout"], s["rows_read"], s["new"], s["existing_same"] + s["existing_different"], s["existing_different"],
             (" | " + s["why"]) if s["why"] else ""))
    # files that were seen before and are not in the folder any more; a newer version of the same name supersedes the older one
    for h, e in ledger["files"].items():
        e["present"] = h in seen_hashes
        newer = [x for k, x in ledger["files"].items() if k != h and x.get("name") == e.get("name") and x.get("last_inspected", "") > e.get("last_inspected", "")
                 and k in seen_hashes]
        e["superseded"] = bool(newer) and h not in seen_hashes
    save_ledger(out_root, ledger)
    write_report(out_root / "Submissions_report.xlsx", rows, details, ctx)
    log("report: %s" % (out_root / "Submissions_report.xlsx"))
    return rows


def write_report(path, rows, details, ctx):
    ov = pd.DataFrame(rows)
    cols = ["name", "path", "status", "decision", "layout", "countries", "rows_read", "rows_used", "excluded", "new", "existing_same",
            "existing_different", "existing_units", "in_file_duplicates_removed", "qc_errors", "qc_warnings", "why", "hash"]
    ov = ov[[c for c in cols if c in ov.columns]].rename(columns={"name": "File", "path": "Path", "status": "Status", "decision": "Decision",
        "layout": "Layout", "countries": "Countries", "rows_read": "Rows read", "rows_used": "Rows used", "excluded": "Rows left out",
        "new": "New records", "existing_same": "Already there (same)", "existing_different": "Already there (other figures)", "existing_units": "Same names as several live records (not added)",
        "in_file_duplicates_removed": "Exact copies removed", "qc_errors": "Records with an Error", "qc_warnings": "Records with a Warning",
        "why": "Why not ready", "hash": "Content hash"})
    sheets = {"Overview": ov}

    def stack(key, extra=None):
        parts = []
        for h, (s, det) in details.items():
            d = det.get(key)
            if d is not None and len(d):
                d = d.copy()
                d.insert(0, "File", s["name"])
                parts.append(d)
        return pd.concat(parts, ignore_index=True) if parts else None
    lay = []
    for h, (s, det) in details.items():
        L = det["layout"] or {}
        lay.append({"File": s["name"], "Layout": L.get("class", ""), "Sheet": L.get("sheet", ""), "Header row": s["header_row"],
                    "Template columns found": L.get("template_columns_found", ""), "Missing key columns": ", ".join(L.get("missing_keys", [])),
                    "Mapped by synonym (verify)": "; ".join(L.get("synonym", [])), "Columns not in the template (ignored)": "; ".join(L.get("ignored", [])[:40]),
                    "Notes": "; ".join(L.get("notes", []))})
    sheets["Layout"] = pd.DataFrame(lay)
    for name, key in (("Records", "compare"), ("Differences_with_live", "diffs"), ("Issues", "issues"), ("Normalised_summary", "note_counts"),
                      ("Normalised_cells", "notes"), ("Defaults_from_live", "defaults"), ("Left_out", "excluded_df"), ("Not_numbers", "bad_numbers"),
                      ("Values_not_seen", "unseen"), ("Duplicates_in_file", "dups"), ("Campaigns_vs_lookup", "campaigns")):
        if key == "excluded_df":
            parts = []
            for h, (s, det) in details.items():
                if det.get("excluded"):
                    d = pd.DataFrame(det["excluded"])
                    d.insert(0, "File", s["name"])
                    parts.append(d)
            d = pd.concat(parts, ignore_index=True) if parts else None
        else:
            d = stack(key)
        if d is not None and len(d):
            sheets[name] = d
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for n, d in sheets.items():
            d.head(200000).to_excel(xw, sheet_name=n[:31], index=False)
            ws = xw.sheets[n[:31]]
            ws.freeze_panes = "A2"
            for k, c in enumerate(d.columns, 1):
                w = 14
                try:
                    w = min(60, max(12, len(str(c)) + 2, int(d[c].astype(str).str.len().quantile(0.9)) + 2))
                except Exception:
                    pass
                ws.column_dimensions[ws.cell(1, k).column_letter].width = w


# ----------------------------------------------------------------------------------------------- the ledger
def load_ledger(out_root):
    p = Path(out_root) / "ledger.json"
    if p.is_file():
        try:
            d = json.loads(p.read_text(encoding="utf-8"))
            d.setdefault("files", {})
            return d
        except ValueError:
            pass
    return {"files": {}}


def save_ledger(out_root, ledger):
    Path(out_root).mkdir(parents=True, exist_ok=True)
    (Path(out_root) / "ledger.json").write_text(json.dumps(ledger, indent=1, ensure_ascii=False), encoding="utf-8")


def find_entry(ledger, target):
    t = (target or "").strip().lower()
    hits = [h for h, e in ledger["files"].items() if h.lower().startswith(t) or e.get("name", "").lower() == t]
    if not hits:
        hits = [h for h, e in ledger["files"].items() if t in e.get("name", "").lower()]
    live = [h for h in hits if ledger["files"][h].get("present", True) and not ledger["files"][h].get("superseded")]
    hits = live or hits
    return hits[0] if len(hits) == 1 else (hits if hits else None)


def decide(out_root, target, decision, by, note=""):
    ledger = load_ledger(out_root)
    h = find_entry(ledger, target)
    if h is None:
        return False, "no file in the ledger matches '%s'" % target
    if isinstance(h, list):
        return False, "'%s' matches %d files; use more of the name or the content hash: %s" % (
            target, len(h), ", ".join(ledger["files"][x]["name"] + " (" + x[:8] + ")" for x in h[:6]))
    e = ledger["files"][h]
    if decision == "accepted":
        if e.get("status") == "blocked" or not e.get("prepared"):
            return False, "%s is blocked (%s): it cannot be accepted" % (e["name"], e.get("why") or "no usable record")
        if not (Path(out_root) / e["prepared"]).is_file():
            return False, "the prepared records of %s are missing: inspect the folder again" % e["name"]
    e.update({"decision": decision, "decided_by": by, "decided_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M"), "note": note})
    save_ledger(out_root, ledger)
    return True, "%s: %s by %s" % (e["name"], decision, by)


def accepted_files(out_root):
    """Prepared record files of the accepted files: [(name, path)]."""
    ledger = load_ledger(out_root)
    out = []
    for h, e in ledger["files"].items():
        if e.get("decision") == "accepted" and not e.get("superseded") and e.get("prepared") and e.get("status") != "blocked":
            p = Path(out_root) / e["prepared"]
            if p.is_file():
                out.append((e["name"], str(p)))
    return out
