"""Submissions intake: the files that countries send (one SharePoint folder) are inspected, normalised and compared with the live
admin data, and the ones an operator accepts are added to the next build.

    python admin.py submissions inspect [--folder DIR]   read every file of the folder, write the report, the prepared records and the ledger
    python admin.py submissions list                     what the ledger knows
    python admin.py submissions accept NAME [--note ..]  the file goes into the next build (never one that is blocked)
    python admin.py submissions reject NAME [--note ..]

What the inspection does for each file
  1. layout: the template (the columns of the live file, with those of the V2 template), the template with other spellings of the headers, a file whose headers can
     be mapped by a short list of certain synonyms, or a file that cannot be used (it is rejected with the reason);
  2. normalisation (each change is listed): capitals and spaces, country and province spelling of the geography reference, vaccine and round
     spelling of the live file, dates (text, Excel serial numbers), numbers written as text, percentages, Year / SIAMTH / Ctry_Code / AFRO_block,
     the values the live file always has for the columns a country does not report, an empty Response filled from the campaign lookup;
  3. the records are checked: key fields, impossible dates and numbers, the portal quality checks (with the live records of the same
     countries as context), values never seen in the live file, the campaigns against the lookup, duplicates inside the file;
  4. every record is compared with the live file: new / already there with the same figures / already there with other figures
     (those are only listed, they are never changed by a submission).
Nothing is written to the live file here: the builder adds the accepted new records (add_coadmin_columns.py --add-records).
Integrated campaigns (an SIA led by polio that also gave measles-rubella, yellow fever, vitamin A, IPV ...)
  The interventions of a file are recognised from the FIGURES of their blocks (children vaccinated / treated / doses delivered or used above zero), not from the
  flags, which the countries often leave wrong. The flags follow the figures and SIA_Integrated becomes Yes (normalise_integration; sheet Integrated_interventions).
  A record that is already in the live file with its polio figures only is completed: the empty cells of the integrated blocks are filled, nothing else is changed
  (integration_additions; sheet Integrated_additions, written in the prepared file and applied by the build: add_coadmin_columns.enrich_records). Figures both
  have that differ are listed (Integrated_differences) and left alone. A campaign written with its round ("...-Round 2") and a round 0 that is only the template's
  default are read from the lookup (split_round_from_response, zero_round_placeholder).
"""
import warnings
import datetime as dt
import difflib
import hashlib
import json
import os
import re
import shutil
import unicodedata
from collections import Counter
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)
KEY_COLS = ["Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response"]
FIGURE_COLS = ["TotPop", "PopPolio", "TotalNbrVaccPolio", "Tot_Vacc_nOPV2", "Total_Vacc_bOPV", "CVPolio"]
TEXT_COLS = ["Country", "Province", "District", "Admin_3", "Response", "Vaccine_type", "Round_Add"]
DATE_COLS = ["SIA_date", "Date_Report", "DateOfOnset"]
CAT_COLS = ["Vaccine_type", "Round_Add", "AgeGPolio", "Main_Act"]
NO_DEFAULT = set(KEY_COLS + FIGURE_COLS + ["Admin_3", "Date_Report", "DateOfOnset", "Year", "SIAMTH", "Index", "check_dp",
                                            "prov_candidates", "Ctry_Code", "AFRO_block", "CV_nOPV2", "CV_bOPV", "EpidNumber"])
EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv")
ARCHIVES = (".zip",)                       # opened: the files inside are inspected as if they were in the folder
NOT_OPENED = (".7z", ".rar", ".tar", ".gz")   # archives that are not opened (extract them and put the files in the folder)
EMPTY_WORDS = {"", "-", "--", "na", "n/a", "nan", "null", "none", "#n/a", "nd", "n.d.", "#div/0!"}
# headers that mean the same as a template column (compared without case, accents, spaces, punctuation). Only what is certain.
SYNONYMS = {"ctryname": "Country", "provname": "Province", "distname": "District", "typeopvvaccine": "Vaccine_type",
            "nbr0dosesvacpolio011m": "Nbr0dosesVaccPolio_0_11M", "totalnbr0dosepolio": "Total_Nbr_0doseVaccPolio_0_59M",
            "totalnbr1dosepluspolio": "Total_Nbr_1dose_Plus_vaccPolio_0_59M", "totalnbrvacpolio": "TotalNbrVaccPolio",
            "usedpolio": "Doses.UsedPolio",
            "round": "Round_Add", "rnd": "Round_Add", "roundnumber": "Round_Add", "rounds": "Round_Add",
            "vaccine": "Vaccine_type", "vaccinetype": "Vaccine_type", "vaccineused": "Vaccine_type",
            "campaign": "Response", "campaignname": "Response", "responsename": "Response", "obrname": "Response",
            "admin1": "Province", "adm1": "Province", "admin2": "District", "adm2": "District",
            "startdate": "SIA_date", "campaignstartdate": "SIA_date", "datestart": "SIA_date", "siastartdate": "SIA_date", "campaignstart": "SIA_date", "startofcampaign": "SIA_date", "campaigndate": "SIA_date",
            "targetpopulation": "PopPolio", "targetpop": "PopPolio", "totalpopulation": "TotPop", "totalpop": "TotPop",
            "totalvaccinated": "TotalNbrVaccPolio", "childrenvaccinated": "TotalNbrVaccPolio", "vaccinated": "TotalNbrVaccPolio",
            "coverage": "CVPolio", "admincoverage": "CVPolio", "reportdate": "Date_Report", "datereport": "Date_Report",
            "countrycode": "Ctry_Code", "iso3": "Ctry_Code", "siacode": "Response", "obr": "Response",
            "number0dosesvaccinated011m": "Nbr0dosesVaccPolio_0_11M", "numbernotdosevaccinated011m": "Nbr_1doses_Plus_VacPolio_0_11M",
            "number0dosesvaccinated1259m": "Nbr0dosesVacPolio_12_59M", "numbernotdosevaccinated1259m": "Nbr_1doses_Plus_VacPolio_12_59M",
            "childrenvaccinated059months": "TotalNbrVaccPolio", "dosesusedpolio": "Doses.UsedPolio", "wastratepolio": "WastRPolio",
            "siasstartdate": "SIA_date", "under5targetted": "PopPolio", "under5targeted": "PopPolio"}


# ----------------------------------------------------------------------------------------------- helpers
V2_NEW = set()   # the columns of the V2 template the live file does not fill yet (set by inspect_folder): they do not count when a file is judged to follow the template
LEARNED = {}     # headers confirmed by the operators (inputs/header_synonyms.json: {"header as written": "template column"})


def cn(x):
    """A header for comparison: no accents, case, spaces or punctuation (SIA-date = SIA_date = sia date)."""
    t = unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().lower()
    t = t.replace("<=", "le").replace(">=", "ge").replace("<", "lt").replace(">", "gt")      # Total_IPV<=1Y and Total_IPV>1Y are two columns
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


def list_files(folder, scratch=None):
    """Every data file below a folder, zip archives opened (the scratch folder holds the extracted files):
    ([(path on disk, path shown: 'Country/archive.zip/inner.xlsx')], [(shown path, why it is not read)])."""
    import zipfile
    folder = Path(folder)
    scratch = Path(scratch) if scratch else None
    out, other = [], []

    def walk_dir(d, shown_prefix, depth):
        for p in sorted(Path(d).rglob("*")):
            if not p.is_file() or p.name.startswith(("~$", ".")) or "__MACOSX" in p.parts:
                continue
            handle(p, shown_prefix + str(p.relative_to(d)).replace("\\", "/"), depth)

    def handle(p, shown, depth):
        ext = p.suffix.lower()
        if ext in EXTENSIONS:
            out.append((p, shown))
        elif ext in ARCHIVES:
            if scratch is None or depth >= 3:
                other.append((shown, "zip archive not opened (too deeply nested)" if scratch else "zip archive not opened"))
                return
            try:
                with zipfile.ZipFile(p) as z:
                    if sum(i.file_size for i in z.infolist()) > 800 * 1024 * 1024:
                        other.append((shown, "zip archive larger than 800 MB once opened: not opened"))
                        return
                    dest = scratch / ("%s_%d" % (sha256_of(p)[:12], depth))
                    if not dest.is_dir():
                        z.extractall(dest)
            except Exception as e:
                other.append((shown, "zip archive could not be opened (%s)" % type(e).__name__))
                return
            walk_dir(dest, shown + "/", depth + 1)
        elif ext in NOT_OPENED:
            other.append((shown, "%s archive: not opened (extract it and put its files in the folder)" % ext))
        else:
            other.append((shown, "not a spreadsheet (%s)" % (ext or "no extension")))

    walk_dir(folder, "", 0)
    return out, other


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
        if c in TEXT_COLS or c in DATE_COLS or c in ("check_dp", "prov_candidates") or c.startswith("AgeG"):
            continue                               # (AgeG<intervention> is an age-group label: 9M-14Y, <5Y)
        s = live[c].replace("", np.nan).dropna()
        if len(s) and pd.to_numeric(s, errors="coerce").notna().mean() >= 0.95:
            prof["numeric"].append(c)
    for c in ("Country", "Province", "District"):
        s = live[c].replace("", np.nan).dropna().astype(str)
        if len(s) and (s == s.str.upper()).mean() >= 0.95:
            prof["upper"].append(c)
    for c in getattr(builder, "TEMPLATE_V2_NEW", {}):          # columns of the V2 template the live file does not fill yet: figures, except the flags and the age groups
        if c in live.columns and c not in prof["numeric"] and not c.startswith(("Integ", "AgeG")):
            prof["numeric"].append(c)
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
    """{sheet: DataFrame of the first rows without header}. Large workbooks are read row by row (not loaded whole)."""
    low = str(path).lower()
    if low.endswith(".csv"):
        return {"(csv)": _read_csv(path, nrows)}
    out = {}
    if low.endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            for ws in wb.worksheets:
                rows = []
                for r in ws.iter_rows(min_row=1, max_row=nrows, values_only=True):
                    rows.append(list(r))
                if rows:
                    w = max(len(r) for r in rows)
                    out[ws.title] = pd.DataFrame([r + [None] * (w - len(r)) for r in rows], dtype=object)
        finally:
            wb.close()
        return out
    eng = "pyxlsb" if low.endswith(".xlsb") else None
    if eng:
        try:
            import pyxlsb  # noqa: F401
        except ImportError:
            raise RuntimeError("an .xlsb file needs the package pyxlsb (pip install pyxlsb), or save it as .xlsx")
    xl = pd.ExcelFile(path, engine=eng)
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


def _map_header(raw, r, tn, template_cols, assume):
    """Map the header row r of one sheet to the template. Returns a layout dict."""
    lay = {"sheet": None, "header_row": r, "map": {}, "synonym": [], "ignored": [], "missing_keys": [], "class": "not usable",
           "template_columns_found": 0, "notes": [], "extra": []}
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
    lay["missing_keys"] = [k for k in KEY_COLS if k not in seen and k not in assume]
    lay["country_from_folder"] = "Country" not in seen and "Country" in assume
    lay["template_columns_found"] = len(seen)
    if lay["missing_keys"]:
        lay["class"] = "not usable"
        lay["notes"].append("key columns missing: " + ", ".join(lay["missing_keys"]))
    elif lay["synonym"] or lay["country_from_folder"]:
        lay["class"] = "mapped (headers differ)"
    elif len(seen) >= 0.7 * len([c for c in template_cols if c not in V2_NEW]):
        lay["class"] = "template"
    else:
        lay["class"] = "template columns, fewer than the template"
    return lay


def detect_layout(path, template_cols, assume=()):
    """Find the sheets with records and their header rows, and map the headers to the template. Every sheet that has the columns of the
    template (or certain synonyms) is read: the best one is the layout itself, the others are in lay['extra']. 'assume' = key columns
    that may be missing because they are known from elsewhere (the Country of the folder)."""
    tn = {cn(c): c for c in template_cols}
    heads = read_head(path)
    found = []                                   # (hits, sheet, header row) of the best header row of each sheet
    for sn, raw in heads.items():
        best = None
        for r in range(min(len(raw), 25)):
            vals = [v for v in raw.iloc[r].tolist() if not is_empty(v)]
            hits = sum(1 for v in vals if cn(v) in tn or cn(v) in SYNONYMS)
            if best is None or hits > best[0]:
                best = (hits, sn, r)
        if best and best[0] >= 3:
            found.append(best)
    lay = {"sheet": None, "header_row": None, "map": {}, "synonym": [], "ignored": [], "missing_keys": list(KEY_COLS), "class": "not usable",
           "template_columns_found": 0, "notes": [], "extra": []}
    if not found:
        lay["notes"].append("no row with column names of the template was found in the first 25 rows of any sheet")
        return lay
    found.sort(key=lambda x: -x[0])
    layouts = []
    for hits, sn, r in found:
        L = _map_header(heads[sn], r, tn, template_cols, assume)
        L["sheet"] = sn
        layouts.append(L)
    usable = [L for L in layouts if L["class"] != "not usable"]
    if not usable:
        lay = layouts[0]
        return lay
    lay = usable[0]
    lay["extra"] = usable[1:]
    for L in layouts:
        if L["class"] == "not usable":
            lay["notes"].append("sheet '%s' not read: %s" % (L["sheet"], "; ".join(L["notes"])))
            lay["skipped_sheets"] = lay.get("skipped_sheets", []) + [L["sheet"]]
    if any(L["class"] == "mapped (headers differ)" for L in usable):
        lay["class"] = "mapped (headers differ)"
    return lay


def _read_sheet(path, L):
    if str(path).lower().endswith(".csv"):
        raw = _read_csv(path)
    else:
        raw = pd.read_excel(path, sheet_name=L["sheet"], header=None, dtype=object, engine="pyxlsb" if str(path).lower().endswith(".xlsb") else None)
    r = L["header_row"]
    body = raw.iloc[r + 1:].copy()
    src = np.arange(r + 2, r + 2 + len(body))     # the row in the file as Excel shows it
    out = pd.DataFrame({t: body.iloc[:, j].values for j, t in L["map"].items()}, index=body.index)
    out["_src_row"] = src
    out["_sheet"] = L["sheet"]
    return out.reset_index(drop=True)


def read_records(path, lay):
    parts = [_read_sheet(path, L) for L in [lay] + list(lay.get("extra") or [])]
    return parts[0] if len(parts) == 1 else pd.concat(parts, ignore_index=True)


def rowref(P, i):
    """The row of record i in the file: 15, or 'Sheet2!15' when records come from several sheets."""
    r = int(P["_src_row"].iat[i])
    if "_sheet" in P.columns and P["_sheet"].nunique() > 1:
        return "%s!%d" % (P["_sheet"].iat[i], r)
    return r


class Notes:
    def __init__(self, cap=30000):
        self.rows, self.counts, self.cap = [], Counter(), cap

    def add(self, src_row, col, found, new, rule):
        self.counts[(rule, col)] += 1
        if len(self.rows) < self.cap:
            self.rows.append({"Row_in_file": src_row, "Column": col, "Found": found, "Set_to": new, "Rule": rule})


def round_offsets(ctx):
    """How the live file numbers the rounds against the lookup: {'by_campaign': {campaign key: offset}, 'common': most frequent offset}."""
    if getattr(ctx, "_round_off", None) is None:
        b = ctx.builder
        try:
            L = ctx.live
            pr = pd.DataFrame({"_rk": L["Response"].map(b._rk), "_d": pd.to_datetime(L["SIA_date"], errors="coerce").dt.strftime("%Y-%m-%d"),
                               "rn": L["Round_Add"].map(b._rnum)}).dropna(subset=["_d"]).drop_duplicates()
            by = b._round_offsets(pr, ctx.lookup)
        except Exception:
            by = {}
        common = Counter(by.values()).most_common(1)[0][0] if by else 0
        ctx._round_off = {"by_campaign": by, "common": common}
    return ctx._round_off


# ----------------------------------------------------------------------------------------------- campaign hints, eTally
VACC_ALIASES = {"nvpo2": "nOPV2", "nopv2": "nOPV2", "vpo2": "nOPV2", "bvpo": "bOPV", "vpob": "bOPV", "bopv": "bOPV", "vpot": "tOPV",
                "topv": "tOPV", "mvpo2": "mOPV2", "mopv2": "mOPV2"}


def vaccine_alias(v):
    return VACC_ALIASES.get(re.sub(r"[^a-z0-9]", "", str(v).lower()))


def name_hints(path, sheets=()):
    """What the file name and the sheet names say about the campaign: a date (of the report, usually), a round, a vaccine."""
    nm = Path(path).stem
    out = {"date": None, "round": None, "vaccine": None, "year": None}
    for rx, order in ((r"(?<!\d)(0[1-9]|[12]\d|3[01])(0[1-9]|1[0-2])(20\d\d)(?!\d)", "dmy"), (r"(?<!\d)(20\d\d)[-_.]?(0[1-9]|1[0-2])[-_.]?(0[1-9]|[12]\d|3[01])(?!\d)", "ymd"),
                      (r"(?<!\d)(0?[1-9]|[12]\d|3[01])[-_. ](0?[1-9]|1[0-2])[-_. ](20\d\d)(?!\d)", "dmy")):
        m = re.search(rx, nm)
        if m:
            g = [int(x) for x in m.groups()]
            try:
                out["date"] = dt.date(g[2], g[1], g[0]) if order == "dmy" else dt.date(g[0], g[1], g[2])
            except ValueError:
                pass
            if out["date"]:
                break
    ys = re.findall(r"(?<!\d)(20[12]\d)(?!\d)", nm)
    if ys:
        out["year"] = int(ys[-1])
    m = re.search(r"(?<![a-z])(?:r|rnd|round|rd|cycle)[\s_\-]*([1-9])(?!\d)", nm, re.I)
    if m:
        out["round"] = int(m.group(1))
    toks = {t.lower() for t in re.findall(r"[nbmt]OPV2?|[nbmt]VPO2?", " ".join([nm] + [str(x) for x in sheets]), re.I)}
    toks = {vaccine_alias(t) for t in toks} - {None}
    if len(toks) == 1:
        out["vaccine"] = toks.pop()
    return out


def live_spelling(ctx):
    """{campaign key: the spelling of its Response in the live file}"""
    if getattr(ctx, "_resp_spell", None) is None:
        b = ctx.builder
        s = ctx.live["Response"].dropna()
        ctx._resp_spell = s.groupby(s.map(b._rk)).agg(lambda x: x.value_counts().index[0]).to_dict() if len(s) else {}
    return ctx._resp_spell


WIRE_MAX_BACK = 60       # days: a record belongs to the previous planned campaign when that one started at most this long before (older: a round missing from the lookup is more likely)
WIRE_MAX_LEAD = 21       # days: a campaign that ran earlier than planned (its first vaccination day before the planned start)


def wire_by_date(lc, starts, ref, rn, offs):
    """Which round of the lookup a record belongs to, from its date alone (a tally has no campaign name).
    lc: the lookup rounds of the country (and vaccine) that may apply; starts: their planned start dates; ref: the date of the
    records; rn: the round number the file gives (or None); offs: the numbering offsets of the live file.
    Rule: a campaign keeps its PLANNED start date, also when it was conducted later; so data dated on or after a planned start belong to
    that planned campaign until the next planned start, and data between two campaigns belong to the previous one (at most
    WIRE_MAX_BACK days back). A campaign can also run EARLIER than planned: when the next planned start is at most WIRE_MAX_LEAD days
    ahead and the previous campaign started longer ago than that, the next one is taken. When both are within WIRE_MAX_LEAD days there is
    no safe answer. Only one campaign (Response + round) may start on the chosen date.
    Returns (index of the lookup row, a text saying how) or None."""
    if lc is None or not len(lc):
        return None
    if rn is not None:
        def same_round(row):
            if pd.isna(row["_rn"]):
                return False
            off = offs["by_campaign"].get(row["_rk"], offs["common"])
            return int(row["_rn"]) + off == rn or int(row["_rn"]) == rn
        lc = lc[lc.apply(same_round, axis=1)]
        if lc.empty:
            return None
    st = starts.loc[lc.index]
    delta = (st - ref).dt.days                     # < 0: the planned start is before the record
    back, ahead = delta[delta <= 0], delta[delta > 0]
    floor_d = back.max() if len(back) else None
    next_d = ahead.min() if len(ahead) else None
    gap = -int(floor_d) if floor_d is not None else None
    lead = int(next_d) if next_d is not None else None

    def one(d_):
        rows = lc.loc[delta[delta == d_].index]
        keys = list(dict.fromkeys(zip(rows["_rk"], rows["_rn"])))
        return rows.index[0] if len(keys) == 1 else None
    if gap == 0:
        i = one(floor_d)
        return (i, "the lookup round that starts on the date of the records") if i is not None else None
    near_next = lead is not None and lead <= WIRE_MAX_LEAD
    near_floor = gap is not None and gap <= WIRE_MAX_LEAD
    if near_next and near_floor:
        return None
    if near_next:
        i = one(next_d)
        return (i, "the campaign ran %d day(s) before its planned start in the lookup" % lead) if i is not None else None
    if gap is not None and gap <= WIRE_MAX_BACK:
        i = one(floor_d)
        return (i, "the records are dated %d day(s) after the planned start of the previous campaign in the lookup and before the next one" % gap) if i is not None else None
    return None


def wire_next_round(lc, starts, ref, rn, offs):
    """The records say a round number (the round column, a hint, or the file name: 'Cycle_2') that the lookup does not have yet.
    When the campaign that started last before the date of the records is the one whose LAST lookup round is rn - 1, and that round
    started at most WIRE_MAX_BACK days before the records, the round exists but the lookup was not updated: the records are the next round
    of the same campaign (same Response). Returns {'x': the lookup row of the last known round, 'last': its number, 'gap': days} or None."""
    if lc is None or not len(lc) or rn is None:
        return None
    st = starts.loc[lc.index]
    back = (ref - st).dt.days
    before = back[back >= 0]
    if before.empty:
        return None
    last_start = st.loc[before.idxmin()]
    cand = lc.loc[before[before == before.min()].index]
    rks = list(dict.fromkeys(cand["_rk"]))
    if len(rks) != 1 or int(before.min()) > WIRE_MAX_BACK:
        return None
    rk = rks[0]
    rows = lc[lc["_rk"] == rk]
    off = offs["by_campaign"].get(rk, offs["common"])
    nums = (rows["_rn"] + off).dropna().astype(int)
    if nums.empty or rn != int(nums.max()) + 1:
        return None
    x = rows.loc[nums.idxmax()]
    if pd.Timestamp(x["_d"]) > ref or int((ref - pd.Timestamp(x["_d"])).days) > WIRE_MAX_BACK:
        return None
    return {"x": x, "last": int(nums.max()), "gap": int((ref - pd.Timestamp(x["_d"])).days)}


ROUND_SUFFIX = re.compile(r"(?i)^(.*?)[\s_\-]*[\s_\-(]\s*(?:round|rnd|rd|r)[\s_\-.]*0*(\d+)\)?\s*$")


def split_round_from_response(P, ctx, notes):
    """A Response that carries its round at the end ('ANG-2026-08-02_nOPV_NID-Round 2', 'NIE-2026-01_nOPV Rnd1'): when what stands before the round is a
    campaign of the lookup for that country, the Response becomes that campaign (as the lookup writes it, in the live file's spelling) and the round is kept
    apart. Returns {row index: round number}; the round is written to Round_Add by round_from_hint when the file gives no round."""
    hint = {}
    lk, b = ctx.lookup, ctx.builder
    if lk is None or "Response" not in P.columns:
        return hint
    spell = live_spelling(ctx)
    for i in P.index:
        r_ = P.at[i, "Response"]
        if is_empty(r_) or is_empty(P.at[i, "Country"]):
            continue
        m = ROUND_SUFFIX.match(str(r_).strip())
        if not m or not m.group(1).strip():
            continue
        base, n = m.group(1).strip(" _-"), int(m.group(2))
        rk = b._rk(base)
        if not (lk[(lk["_ck"] == b._lkc(P.at[i, "Country"])) & (lk["_rk"] == rk)].shape[0]):
            continue                                    # what remains is not a campaign of the lookup: the Response is left as written
        if b._rk(r_) in set(lk["_rk"]):
            continue                                    # the whole text is itself a name of the lookup
        name = spell.get(rk) or lk[lk["_rk"] == rk].iloc[0]["OBR Name"]
        notes.add(rowref(P, i), "Response", r_, name, "the round (%d) was written at the end of the campaign name: campaign and round separated" % n)
        P.at[i, "Response"] = name
        hint[i] = n
    return hint


def zero_round_placeholder(P, ctx, notes):
    """The template is sent with 0 in the round (and the vaccine) when the country did not fill them. A round 0 is a real round only for a campaign whose lookup
    has a round 0 (the R0 campaigns): for a campaign of the lookup that has none, the 0 is a placeholder, the round is left empty and found from the lookup
    (the date of the records, or the round written in the campaign name)."""
    if "Round_Add" not in P.columns or "Response" not in P.columns:
        return
    lk, b = ctx.lookup, ctx.builder
    for i in P.index[P["Round_Add"].map(lambda v: isinstance(v, str) and re.fullmatch(r"(?i)\s*rnd\s*0+\s*", v) is not None)]:
        r_, c_ = P.at[i, "Response"], P.at[i, "Country"]
        if is_empty(r_) or is_empty(c_):
            continue
        lc = lk[(lk["_ck"] == b._lkc(c_)) & (lk["_rk"] == b._rk(r_))]
        if lc.empty or (lc["_rn"] == 0).any():
            continue
        notes.add(rowref(P, i), "Round_Add", P.at[i, "Round_Add"], "", "0 is not a round of this campaign (template default): left empty, taken from the lookup")
        P.at[i, "Round_Add"] = None


def round_from_hint(P, ctx, notes, hint):
    """The round number written in the campaign name, for the records whose round is still empty: the round of the lookup with that number
    (the date of the records agrees when the lookup has it), written the way the live file numbers this campaign."""
    if not hint or "Round_Add" not in P.columns:
        return
    lk, b = ctx.lookup, ctx.builder
    offs = round_offsets(ctx)
    for i, n in hint.items():
        if not is_empty(P.at[i, "Round_Add"]):
            continue
        rk = b._rk(P.at[i, "Response"])
        lc = lk[(lk["_ck"] == b._lkc(P.at[i, "Country"])) & (lk["_rk"] == rk) & (lk["_rn"] == n)]
        if lc["_rn"].nunique() != 1:
            continue
        off = offs["by_campaign"].get(rk, offs["common"])
        P.at[i, "Round_Add"] = "Rnd%d" % (n + off)
        notes.add(rowref(P, i), "Round_Add", "", P.at[i, "Round_Add"], "round %d of the lookup, written in the campaign name (the live file numbers this campaign %+d from the lookup)" % (n, off))


def align_round_to_lookup(P, ctx, notes):
    """THE LOOKUP IS THE AUTHORITY: a record whose campaign is one of the lookup is written with the lookup's name for the Response, and with the round
    the lookup gives to that campaign and start date ('Rnd' + its number) whatever number the file gives (Angola's file says Round Number 3 for the
    lookup's Round 1: it is Rnd1). The start dates are standardised before (standardize_to_lookup). The live file follows the same rule (the build changes
    its records). Returns the list of texts, one per campaign round."""
    out = []
    lk = ctx.lookup
    if lk is None or not {"Country", "Response", "SIA_date", "Round_Add"} <= set(P.columns) or not len(P):
        return out
    b = ctx.builder
    T = P[["Country", "Response", "SIA_date", "Round_Add"]]
    names, nl = b.lookup_response_names(T, lk)
    for i in P.index:
        if not is_empty(P.at[i, "Response"]) and str(names.at[i]) != str(P.at[i, "Response"]):
            notes.add(rowref(P, i), "Response", P.at[i, "Response"], names.at[i], "campaign name as written in the lookup")
            P.at[i, "Response"] = names.at[i]
    for _, r in nl.iterrows():
        out.append("campaign name: %s -> %s (%d record(s))" % (r["Found"], r["Restored"], r["Records"]))
    # a start date a few days from the planned start of the lookup: the planned start (same rule as the live file)
    cols = [c for c in ("Country", "Response", "SIA_date", "Round_Add", "Province", "District", "Admin_3", "Vaccine_type") if c in P.columns]
    dts, dl, _sk = b.lookup_dates(P[cols], lk)
    for i in P.index:
        if not pd.isna(P.at[i, "SIA_date"]) and b._dkey(dts.at[i]) != b._dkey(P.at[i, "SIA_date"]):
            notes.add(rowref(P, i), "SIA_date", b._dkey(P.at[i, "SIA_date"]), b._dkey(dts.at[i]), "planned start date of the lookup (within %d days; the lookup is the authority)" % b.LOOKUP_DATE_DAYS)
            P.at[i, "SIA_date"] = pd.Timestamp(dts.at[i])
    for _, r in dl.iterrows():
        out.append("%s: date %s -> planned start %s (%d record(s))" % (r["Response"], r["Found"], r["Restored"], r["Records"]))
    new, log = b.lookup_round_labels(P[["Country", "Response", "SIA_date", "Round_Add"]], lk)
    for i in P.index:
        old = P.at[i, "Round_Add"]
        if is_empty(old) or str(new.at[i]) == str(old):
            continue
        P.at[i, "Round_Add"] = new.at[i]
        notes.add(rowref(P, i), "Round_Add", old, new.at[i], "round number of the lookup (the lookup is the authority): %s started %s is %s in the lookup" %
                  (P.at[i, "Response"], pd.Timestamp(P.at[i, "SIA_date"]).strftime("%Y-%m-%d"), new.at[i]))
    for _, r in log.iterrows():
        out.append("%s %s: the file says %s, the lookup's round is %s (%d record(s))" % (r["Response"], r["SIA_date"], r["Found"], r["Restored"], r["Records"]))
    return out


def standardize_to_lookup(P, ctx, notes):
    """Records that name a campaign of the lookup but carry another date: the campaign keeps its PLANNED start date (the lookup's), also when it was
    conducted earlier or later. Same country and same Response as the lookup, a date that is not a planned start of that Response, one safe answer
    from wire_by_date (the round the file gives is respected): SIA_date becomes the planned start, an empty round is filled, the date in the file goes
    in the notes. Returns the list of texts (one per country, Response and date), to confirm."""
    out = []
    lk = ctx.lookup
    if lk is None or not {"Country", "Response", "SIA_date"} <= set(P.columns) or not len(P):
        return out
    b = ctx.builder
    starts = pd.to_datetime(lk["_d"], errors="coerce")
    offs = round_offsets(ctx)
    ck = P["Country"].map(b._lkc)
    rk = P["Response"].map(lambda x: b._rk(x) if not is_empty(x) else "")
    ok = P["SIA_date"].notna() & (rk != "")
    for (c, r_, d_), g in P[ok].groupby([ck[ok], rk[ok], P.loc[ok, "SIA_date"].dt.strftime("%Y-%m-%d")]):
        lc = lk[(lk["_ck"] == c) & (lk["_rk"] == r_)]
        if lc.empty or d_ in set(lc["_d"]):
            continue
        rns = {b._rnum(x) for x in g["Round_Add"] if not is_empty(x)} if "Round_Add" in g.columns else set()
        rns.discard(None)
        rn = next(iter(rns)) if len(rns) == 1 else None
        if len(rns) > 1:
            continue
        try:
            i = wire_by_date(lc, starts, pd.Timestamp(d_), rn, offs)
        except Exception:
            i = None
        if not i:
            continue
        j, how = i
        x = lk.loc[j]
        new_d = pd.Timestamp(x["_d"])
        if pd.isna(new_d):
            continue
        off = offs["by_campaign"].get(x["_rk"], offs["common"])
        for k in g.index:
            P.at[k, "SIA_date"] = new_d
            notes.add(rowref(P, k), "SIA_date", d_, new_d.strftime("%Y-%m-%d"), "planned start date of the lookup kept (%s)" % how)
            if "Round_Add" in P.columns and is_empty(P.at[k, "Round_Add"]) and not pd.isna(x["_rn"]):
                P.at[k, "Round_Add"] = "Rnd%d" % (int(x["_rn"]) + off)
                notes.add(rowref(P, k), "Round_Add", "", P.at[k, "Round_Add"], "round of the lookup")
        out.append("%s %s: %d record(s) dated %s now carry the planned start %s (%s)" % (x["OBR Name"], x["Round Number"], len(g), d_, new_d.strftime("%Y-%m-%d"), how))
    return out


def infer_campaign(P, ctx, prof, hints, notes):
    """Records without a Response: find the campaign in the lookup from what the file does say (country, SIA_date, round, vaccine; or, with no
    date, the date / round / vaccine written in the file name). Only a unique answer is used, and it is always reported for the operator to
    confirm. Returns the list of messages."""
    b, lk, out = ctx.builder, ctx.lookup, []
    if lk is None or "Response" not in P.columns or not len(P):
        return out
    vmap0 = {cn(v): v for v in prof["values"].get("Vaccine_type", ())}
    # a Response is written but not the vaccine: the vaccine of that campaign in the lookup (when it has only one) or the one in the sheet / file name
    nov = P["Vaccine_type"].map(is_empty) & ~P["Response"].map(is_empty)
    if nov.any():
        n_set = 0
        for i in P.index[nov]:
            rk = b._rk(P.at[i, "Response"])
            lc = lk[(lk["_rk"] == rk) & (lk["_ck"] == b._lkc(P.at[i, "Country"]))] if not is_empty(P.at[i, "Country"]) else lk[lk["_rk"] == rk]
            vals = list(dict.fromkeys(str(v) for v in lc["Vaccines"].dropna())) if "Vaccines" in lc.columns else []
            vt = vmap0.get(cn(vals[0])) if len(vals) == 1 else None
            how = "vaccine of this campaign in the lookup"
            if not vt and hints.get("vaccine"):
                vt, how = vmap0.get(cn(hints["vaccine"])), "vaccine written in the file or sheet name"
            if vt:
                P.at[i, "Vaccine_type"] = vt
                notes.add(rowref(P, i), "Vaccine_type", "", vt, how)
                n_set += 1
        if n_set:
            out.append("vaccine of %d record(s) taken from the lookup / the file name" % n_set)
    # a Response and a round are written but no date: the start date of that round in the lookup
    nod = P["SIA_date"].isna() & ~P["Response"].map(is_empty) & ~P["Round_Add"].map(is_empty)
    if nod.any():
        offs0, cache, n_set = round_offsets(ctx), {}, 0
        for i in P.index[nod]:
            rk, rn = b._rk(P.at[i, "Response"]), b._rnum(P.at[i, "Round_Add"])
            ck0 = b._lkc(P.at[i, "Country"]) if not is_empty(P.at[i, "Country"]) else None
            key = (rk, rn, ck0)
            if key not in cache:
                lc = lk[lk["_rk"] == rk]
                if ck0:
                    lc = lc[lc["_ck"] == ck0]
                off = offs0["by_campaign"].get(rk, offs0["common"])
                m = lc[[(not pd.isna(r)) and (int(r) + off == rn or int(r) == rn) for r in lc["_rn"]]]
                if not len(m) and lc["_d"].nunique() == 1:        # the lookup has one round for this campaign and the file numbers it differently
                    m = lc
                cache[key] = (m.iloc[0]["_d"], m.iloc[0]["Round Number"], m.iloc[0]["OBR Name"]) if len(m) and m["_d"].nunique() == 1 else None
            if cache[key]:
                d_, rname, obr = cache[key]
                P.at[i, "SIA_date"] = pd.Timestamp(d_)
                notes.add(rowref(P, i), "SIA_date", "", d_, "start date of %s %s in the lookup (the file gives no date)" % (obr, rname))
                n_set += 1
        if n_set:
            out.append("SIA_date of %d record(s) taken from the lookup (start of the round)" % n_set)
    blank = P["Response"].map(is_empty)
    if not blank.any():
        return out
    offs, spell = round_offsets(ctx), live_spelling(ctx)
    starts = pd.to_datetime(lk["_d"], errors="coerce")
    keyf = pd.DataFrame({"c": P["Country"].map(lambda v: "" if is_empty(v) else b._lkc(v)),
                         "d": P["SIA_date"].map(lambda x: "" if pd.isna(x) else pd.Timestamp(x).strftime("%Y-%m-%d")),
                         "r": P["Round_Add"].map(lambda x: "" if is_empty(x) else str(x)),
                         "v": P["Vaccine_type"].map(lambda x: "" if is_empty(x) else str(x))})
    vmap = {cn(v): v for v in prof["values"].get("Vaccine_type", ())}
    named = ~(P["Province"].map(is_empty) | P["District"].map(is_empty))        # a line without province or district is not a record
    for (c, d_, r_, v_), g in keyf[blank & named].groupby(["c", "d", "r", "v"]):
        if not c:
            continue
        lc0 = lc = lk[lk["_ck"] == c]
        vs = b._vset(v_) if v_ else set()
        if not vs and hints.get("vaccine"):
            vs = b._vset(hints["vaccine"])
        vnote = ""
        if vs and len(lc):
            lc = lc0[[bool(x & vs) or not x for x in lc0["_vs"]]]
            if lc.empty:
                lc, vnote = lc0, "; the vaccine of the file (%s) is not the lookup's for this campaign" % (v_ or hints.get("vaccine"))
        if lc.empty:
            continue
        rn = b._rnum(r_) if r_ else hints.get("round")
        if d_:
            ref = pd.Timestamp(d_)
            t0, t1, src = ref - pd.Timedelta(days=45), ref + pd.Timedelta(days=45), "the SIA_date of the records"
        elif hints.get("date"):
            ref = pd.Timestamp(hints["date"])
            t0, t1, src = ref - pd.Timedelta(days=100), ref + pd.Timedelta(days=7), "the date in the file name (a campaign starts before its report)"
        elif hints.get("year"):
            ref = None
            t0, t1, src = pd.Timestamp(hints["year"], 1, 1), pd.Timestamp(hints["year"], 12, 31), "the year in the file name"
        else:
            continue
        try:
            wired = wire_by_date(lc, starts, ref, rn, offs) if (d_ and ref is not None) else None
        except Exception:
            wired = None            # the earlier inference (window and nearest round) then decides, as before
        nxt = None
        if wired is None and d_ and ref is not None and rn is not None:
            try:
                nxt = wire_next_round(lc, starts, ref, rn, offs)
            except Exception:
                nxt = None
        if nxt is not None:
            lc, src = lc.loc[[nxt["x"].name]], "the date of the records"
        elif wired is not None:
            lc, src = lc.loc[[wired[0]]], "the date of the records: " + wired[1]
        else:
            lc = lc[(starts.loc[lc.index] >= t0) & (starts.loc[lc.index] <= t1)]
            if rn is not None and len(lc):
                def same_round(row):
                    if pd.isna(row["_rn"]):
                        return False
                    off = offs["by_campaign"].get(row["_rk"], offs["common"])
                    return int(row["_rn"]) + off == rn or int(row["_rn"]) == rn
                lc = lc[lc.apply(same_round, axis=1)]
            if lc.empty:
                continue
            keys = list(dict.fromkeys(zip(lc["_rk"], lc["_rn"])))
            if len(keys) != 1:
                if ref is None:
                    continue
                dd = (starts.loc[lc.index] - ref).dt.days.abs().sort_values()
                if len(dd) > 1 and dd.iloc[1] - dd.iloc[0] >= 20:
                    lc = lc.loc[[dd.index[0]]]
                else:
                    continue
        x = lc.iloc[0]
        rk = x["_rk"]
        name = spell.get(rk) or x["OBR Name"]
        off = offs["by_campaign"].get(rk, offs["common"])
        why = "campaign inferred: %s %s (starts %s) in the lookup, from %s%s" % (x["OBR Name"], x["Round Number"], x["Round Start Date"], src, vnote)
        rname_out = x["Round Number"]
        if nxt is not None:
            rname_out = "Rnd%d" % rn
            why = ("round %d of %s is not in the lookup yet (the lookup has rounds up to %d, the last starting %s): taken as the next round of "
                   "the same campaign, with the date of the records%s" % (rn, x["OBR Name"], nxt["last"], x["Round Start Date"], vnote))
            gl = getattr(ctx, "_lookup_gaps", None)
            if gl is None:
                gl = ctx._lookup_gaps = []
            gl.append({"Country": x["Country"], "Response": name, "Round_missing_in_lookup": rn, "Date_in_the_data": d_,
                       "Last_round_in_lookup": nxt["last"], "Last_round_start": x["Round Start Date"], "Records": len(g)})
        vt = None
        if "Vaccines" in lk.columns and not is_empty(x.get("Vaccines")):
            vt = vmap.get(cn(x["Vaccines"]))
        for i in g.index:
            P.at[i, "Response"] = name
            notes.add(rowref(P, i), "Response", "", name, why)
            if is_empty(P.at[i, "Round_Add"]) and (nxt is not None or not pd.isna(x["_rn"])):
                P.at[i, "Round_Add"] = "Rnd%d" % (rn if nxt is not None else int(x["_rn"]) + off)
                notes.add(rowref(P, i), "Round_Add", "", P.at[i, "Round_Add"], why)
            if pd.isna(P.at[i, "SIA_date"]):
                P.at[i, "SIA_date"] = pd.Timestamp(x["_d"])
                notes.add(rowref(P, i), "SIA_date", "", x["_d"], why)
            elif wired is not None and pd.Timestamp(P.at[i, "SIA_date"]).normalize() != pd.Timestamp(x["_d"]).normalize():
                old_d = pd.Timestamp(P.at[i, "SIA_date"]).strftime("%Y-%m-%d")
                P.at[i, "SIA_date"] = pd.Timestamp(x["_d"])          # the planned date of the lookup is kept (the file's own date is in the note)
                notes.add(rowref(P, i), "SIA_date", old_d, x["_d"], why + "; the planned start of the lookup is kept, the file says " + old_d)
            if is_empty(P.at[i, "Vaccine_type"]) and vt:
                P.at[i, "Vaccine_type"] = vt
                notes.add(rowref(P, i), "Vaccine_type", "", vt, why)
        out.append("%s %s (%d record(s); from %s%s)" % (name, rname_out, len(g), src + ("; round not in the lookup yet" if nxt is not None else ""), vnote))
    return out


def campaign_from_live(P, ctx, notes):
    """Records with no campaign (Response, round, vaccine or date missing): the live file may already hold the same records. Each record is matched
    on country, province, district and the number of children vaccinated; when most of the matches point to one campaign, the missing fields are
    taken from it. Only fields that are empty are filled, and the operator is always asked to confirm. Returns the list of messages."""
    b, out = ctx.builder, []
    need = ["Response", "Round_Add", "Vaccine_type"]
    if ctx.live is None or "TotalNbrVaccPolio" not in P.columns or not len(P) or "Response" not in P.columns:
        return out
    miss = P["Response"].map(is_empty) | P["Round_Add"].map(is_empty) | P["Vaccine_type"].map(is_empty) | P["SIA_date"].isna()
    named = ~(P["Province"].map(is_empty) | P["District"].map(is_empty))
    todo = P.index[miss & named]
    if not len(todo):
        return out
    pidx, L = place_index(ctx), ctx.live
    votes, per = Counter(), {}
    for i in todo:
        tot = pd.to_numeric(P.at[i, "TotalNbrVaccPolio"], errors="coerce")
        if pd.isna(tot) or tot <= 0:
            continue
        hit = place_hits(ctx, b._qcountry(P.at[i, "Country"]), cn(P.at[i, "Province"]), cn(P.at[i, "District"]))
        rows = [h for h in hit if (not pd.isna(h[3])) and abs(h[3] - tot) < 0.5 and not is_empty(h[5])]
        keys = {(str(h[5]), str(h[6])) for h in rows}
        if len(keys) == 1:
            k = next(iter(keys))
            votes[k] += 1
            per[i] = (k, rows[0][0])
    n_hit = sum(votes.values())
    if not n_hit:
        return out
    (rs, ro), top = votes.most_common(1)[0]
    if top < 0.6 * n_hit or n_hit < max(1, 0.2 * len(todo)):
        out.append("live file: %d of %d record(s) without campaign match a live record, but not one campaign (%s): not used" %
                   (n_hit, len(todo), ", ".join("%s %s x%d" % (k[0], k[1], v) for k, v in votes.most_common(3))))
        return out
    sel = L[(L["Response"].astype(str) == rs) & (L["Round_Add"].astype(str) == ro)]
    ctr = {b._qcountry(c) for c in P.loc[todo, "Country"].dropna().unique()}
    sel_c = sel[sel["Country"].map(b._qcountry).isin(ctr)]
    sel = sel_c if len(sel_c) else sel
    def mode(col):
        v = sel[col].dropna().astype(str)
        v = v[v.str.strip() != ""]
        return v.value_counts().index[0] if len(v) else None
    vacc = mode("Vaccine_type")
    d_mode = pd.to_datetime(sel["SIA_date"], errors="coerce").dropna()
    d_mode = d_mode.value_counts().index[0] if len(d_mode) else None
    why = "campaign taken from the live file: %d of %d record(s) without a campaign have the same place and number vaccinated in %s %s" % (top, len(todo), rs, ro)
    for i in todo:
        for col, val in (("Response", rs), ("Round_Add", ro), ("Vaccine_type", vacc)):
            if val and is_empty(P.at[i, col]):
                P.at[i, col] = val
                notes.add(rowref(P, i), col, "", val, why)
        if P.at[i, "SIA_date"] is pd.NaT or pd.isna(P.at[i, "SIA_date"]):
            d_ = pd.to_datetime(L.at[per[i][1], "SIA_date"], errors="coerce") if i in per and per[i][0] == (rs, ro) else d_mode
            if d_ is not None and not pd.isna(d_):
                P.at[i, "SIA_date"] = pd.Timestamp(d_)
                notes.add(rowref(P, i), "SIA_date", "", pd.Timestamp(d_).strftime("%Y-%m-%d"), why)
    other = int((~P.loc[todo, "Round_Add"].map(is_empty) & (P.loc[todo, "Round_Add"].astype(str) != ro)).sum())
    out.append("%s %s (%d of %d record(s) matched in the live file; %d fill-in(s) for the rest%s)" % (
        rs, ro, top, len(todo), len(todo) - top,
        "; the file itself writes another round for %d record(s), kept as written: check" % other if other else ""))
    return out


def subdistrict_check(P, ctx, notes):
    """A file of health-facility / sub-district lines (Admin_3 filled, each district on many lines) cannot be added to the district-level live file. When the
    district totals (children vaccinated and target population) equal those of one campaign already in the live file, the file is its detail: the campaign
    is filled in and the lines are reported as already there. Returns a dict (Response, Round_Add, districts, matched, message) or None."""
    b = ctx.builder
    if ctx.live is None or "Admin_3" not in P.columns or "TotalNbrVaccPolio" not in P.columns or not len(P):
        return None
    a3 = ~P["Admin_3"].map(is_empty)
    keys = list(zip(P["Country"].map(lambda v: b._qcountry(v) if not is_empty(v) else ""), P["Province"].map(cn), P["District"].map(cn)))
    if a3.mean() < 0.5 or len(P) < 1.5 * len(set(keys)):
        return None
    tot = pd.to_numeric(P["TotalNbrVaccPolio"], errors="coerce")
    pop = pd.to_numeric(P["PopPolio"], errors="coerce") if "PopPolio" in P.columns else pd.Series(np.nan, index=P.index)
    g = pd.DataFrame({"k": keys, "t": tot, "p": pop}).groupby("k").agg(t=("t", "sum"), p=("p", "sum"))
    pidx, votes = place_index(ctx), Counter()
    for k, r in g.iterrows():
        cs = {(str(h[5]), str(h[6])) for h in place_hits(ctx, *k) if (not pd.isna(h[3])) and abs(h[3] - r["t"]) < 0.5
              and (pd.isna(h[4]) or r["p"] == 0 or abs(h[4] - r["p"]) < 0.5) and not is_empty(h[5])}
        if len(cs) == 1:
            votes[next(iter(cs))] += 1
    if not votes:
        return None
    (rs, ro), top = votes.most_common(1)[0]
    if top < 3 or top < 0.4 * len(g) or top < 0.9 * sum(votes.values()):
        return None
    msg = ("health-facility lines (%d lines in %d districts); the district totals equal %s %s already in the live file (%d of %d districts; the others are spelled differently or are older district names), "
           "so this is its detail: nothing to add" % (len(P), len(g), rs, ro, top, len(g)))
    return {"Response": rs, "Round_Add": ro, "districts": len(g), "matched": top, "message": msg}


def fill_down_sheet(P, notes):
    """Lines of a sheet that lack the campaign while the other lines of the same sheet all carry one and the same campaign (Response, round, date):
    the campaign is filled down (the person who filled the sheet did not copy it to the last lines). Reported for the operator to confirm."""
    out = []
    if "_sheet" not in P.columns or not len(P):
        return out
    for sh, g in P.groupby("_sheet", dropna=False):
        full = g[~g["Response"].map(is_empty) & ~g["Round_Add"].map(is_empty) & g["SIA_date"].notna()]
        todo = g[g["Response"].map(is_empty) & ~(g["Province"].map(is_empty) | g["District"].map(is_empty))]
        if len(full) < 3 or not len(todo):
            continue
        trip = set(zip(full["Response"].astype(str), full["Round_Add"].astype(str), full["SIA_date"]))
        if len(trip) != 1:
            continue
        rs, ro, dt_ = next(iter(trip))
        why = "campaign filled down from the other %d line(s) of sheet %s (all give %s %s, %s)" % (len(full), sh, rs, ro, pd.Timestamp(dt_).strftime("%Y-%m-%d"))
        for i in todo.index:
            for col, val in (("Response", rs), ("Round_Add", ro), ("SIA_date", pd.Timestamp(dt_))):
                if (P.at[i, col] is pd.NaT or pd.isna(P.at[i, col])) if col == "SIA_date" else is_empty(P.at[i, col]):
                    P.at[i, col] = val
                    notes.add(rowref(P, i), col, "", val, why)
        out.append("%s %s: %d line(s) of sheet %s had no campaign, filled down from the other lines" % (rs, ro, len(todo), sh))
    return out


ET_ZD = ("h2h_zero_dose_sum", "fp_opv_ri_zd_vaccinated", "dopv_vaccinated_zd_vaccinated", "spt_opv_vac_spt_zd_vaccinated", "tt_opv_tt_zd_vaccinated")


def etally_find(path):
    """The sheet of an eTally export (one line per settlement and day: state, lga, vacdate, total_children_vaccinated ...), or None."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            hdr = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            low = [str(h).strip().lower() if h is not None else "" for h in (hdr or ())]
            if all(k in low for k in ("state", "lga", "vacdate")) and any(h.startswith("total_children_vaccinated") for h in low):
                return ws.title
    finally:
        wb.close()
    return None


def etally_layout(path, sheet, folder_country):
    """Add the settlement lines up by state and LGA for the campaign: one record per LGA, in the template's columns.
    Returns (layout dict, records frame)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    it = ws.iter_rows(values_only=True)
    hdr = [str(h).strip().lower() if h is not None else "" for h in next(it)]
    first = {}
    for j, h in enumerate(hdr):
        if h and h not in first:
            first[h] = j
    cols = [w for w in ("id", "state", "lga", "vacdate", "total_children_vaccinated", "total_children_vaccinated_bopv", "total_vaccine_doses_used") + ET_ZD if w in first]
    ix = [first[w] for w in cols]
    rows = [[r[j] if j < len(r) else None for j in ix] + [n] for n, r in enumerate(it, start=2)]
    wb.close()
    d = pd.DataFrame(rows, columns=cols + ["_row"])
    n0 = len(d)
    notes = []
    if "id" in d.columns:
        k = d["id"].notna() & d.duplicated("id", keep="last")
        if k.any():
            notes.append("%d settlement line(s) with an id already seen (edited submissions): the last one is kept" % int(k.sum()))
        d = d[~k]
    for c in cols:
        if c not in ("id", "state", "lga", "vacdate"):
            d[c] = pd.to_numeric(d[c], errors="coerce")
    d["vacdate"] = pd.to_datetime(d["vacdate"], errors="coerce")
    nolga = d["state"].map(is_empty) | d["lga"].map(is_empty)
    if nolga.any():
        notes.append("%d settlement line(s) without state or LGA left out" % int(nolga.sum()))
    d = d[~nolga]
    med = d["vacdate"].dropna().sort_values().iloc[len(d["vacdate"].dropna()) // 2]
    ok = ((d["vacdate"] - med).abs() <= pd.Timedelta(days=20))
    if (~ok).any():
        notes.append("%d settlement line(s) with a vaccination date more than 20 days from the campaign (%s) left out" % (int((~ok).sum()), med.strftime("%Y-%m-%d")))
    d = d[ok]
    cnt = d["vacdate"].dt.normalize().value_counts().sort_index()
    start = cnt[cnt >= max(10, 0.01 * len(d))].index.min()          # the first day on which the teams really report (a few test lines come earlier)
    early = d["vacdate"] < start
    if early.any():
        notes.append("%d settlement line(s) dated before the first vaccination day (%s) left out (test entries)" % (int(early.sum()), start.strftime("%Y-%m-%d")))
        d = d[~early]
    agg = {"_row": "min"}
    out_cols = {"total_children_vaccinated": "TotalNbrVaccPolio", "total_vaccine_doses_used": "Doses.UsedPolio", "total_children_vaccinated_bopv": "Total_Vacc_bOPV"}
    for c in out_cols:
        if c in d.columns:
            agg[c] = lambda x: x.sum(min_count=1)
    zd = [c for c in ET_ZD if c in d.columns]
    for c in zd:
        agg[c] = lambda x: x.sum(min_count=1)
    g = d.groupby(["state", "lga"]).agg(agg).reset_index()
    g["_n"] = d.groupby(["state", "lga"]).size().values
    rec = pd.DataFrame({"Country": folder_country, "Province": g["state"], "District": g["lga"], "SIA_date": start})
    for c, t in out_cols.items():
        if c in g.columns:
            rec[t] = g[c]
    if zd:
        z = g[zd].sum(axis=1, min_count=1)
        rec["Total_Nbr_0doseVaccPolio_0_59M"] = z
        rec["Total_Nbr_1dose_Plus_vaccPolio_0_59M"] = rec["TotalNbrVaccPolio"] - z
    rec["_src_row"] = g["_row"].values
    rec["_sheet"] = sheet
    notes.insert(0, "eTally export: %d settlement lines added up into %d LGA records for the campaign that starts %s (first vaccination day)" % (len(d), len(rec), start.strftime("%Y-%m-%d")))
    notes.append("figures added up from the settlement lines: children vaccinated, doses used, zero-dose children (household, fixed post, other teams); 1-dose-plus = total minus zero-dose; the target population is not in an eTally file")
    lay = {"sheet": sheet, "header_row": 0, "map": {}, "synonym": [], "ignored": [], "missing_keys": [], "class": "eTally (settlement lines added up by LGA)",
           "template_columns_found": len(rec.columns) - 2, "notes": notes, "extra": [], "etally": True}
    return lay, rec


def population_from_live(ctx):
    """{(country, province, district): the PopPolio of the latest live record}"""
    if getattr(ctx, "_pop_live", None) is None:
        b, L = ctx.builder, ctx.live
        t = pd.DataFrame({"c": L["Country"].map(b._qcountry), "p": L["Province"].map(cn), "d": L["District"].map(cn),
                          "dt": pd.to_datetime(L["SIA_date"], errors="coerce"), "pop": pd.to_numeric(L["PopPolio"], errors="coerce")}).dropna(subset=["pop", "dt"])
        t = t.sort_values("dt").drop_duplicates(["c", "p", "d"], keep="last")
        ctx._pop_live = {(r.c, r.p, r.d): r.pop for r in t.itertuples()}
    return ctx._pop_live


# ----------------------------------------------------------------------------------------------- district spelling, call-in workbooks
def live_districts(ctx):
    """{(country, province): {district key: spelling in the live file}}"""
    if getattr(ctx, "_live_dist", None) is None:
        b, L = ctx.builder, ctx.live
        out = {}
        for c, p_, d_ in zip(L["Country"].map(b._qcountry), L["Province"].map(cn), L["District"]):
            if isinstance(d_, str) and d_.strip():
                out.setdefault((c, p_), {}).setdefault(cn(d_), d_)
        ctx._live_dist = out
    return ctx._live_dist


def align_districts(P, ctx, notes):
    """A district that is not in the live file but has a close spelling there (GIREI / GIRIE) takes the live spelling. Only when exactly one
    live name is close, that live name is not also a district of the file, and the province has live records."""
    b, ld = ctx.builder, live_districts(ctx)
    n = 0
    keys = pd.DataFrame({"c": P["Country"].map(lambda v: "" if is_empty(v) else b._qcountry(v)), "p": P["Province"].map(lambda v: "" if is_empty(v) else cn(v))})
    for (c, p_), g in keys.groupby(["c", "p"]):
        live = ld.get((c, p_))
        if not live:
            continue
        mine = {cn(x) for x in P.loc[g.index, "District"].dropna()}
        free = [x for x in live if x not in mine]
        for dname in sorted({x for x in P.loc[g.index, "District"].dropna()}):
            k = cn(dname)
            if k in live or not free:
                continue
            best = difflib.get_close_matches(k, free, n=2, cutoff=0.8)
            if len(best) == 2 and difflib.SequenceMatcher(None, k, best[0]).ratio() - difflib.SequenceMatcher(None, k, best[1]).ratio() < 0.05:
                continue
            if best:
                new = live[best[0]]
                for i in P.index[(keys["c"] == c) & (keys["p"] == p_) & (P["District"] == dname)]:
                    notes.add(rowref(P, i), "District", dname, new, "district spelling of the live file (close match)")
                    P.at[i, "District"] = new
                    n += 1
    return n


def read_callin(path):
    """A daily call-in workbook (one sheet per state, nOPV and bOPV): {State, LGA, Target, Total, Used} per LGA."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True, max_row=400))
            hi = next((i for i, r in enumerate(rows[:6]) if r and any(isinstance(x, str) and x.strip() == "State" for x in r)), None)
            if hi is None:
                continue
            h = [str(x).replace("\n", " ").strip().lower() if x is not None else "" for x in rows[hi]]
            col = lambda pre: next((j for j, x in enumerate(h) if x.startswith(pre)), None)
            cs, cl, ct, cc, cu = col("state"), col("lga"), col("target pop"), col("total children imm"), col("total vaccine us")
            if None in (cs, cl, cc):
                continue
            for r in rows[hi + 1:]:
                if r and isinstance(r[cl], str) and r[cl].strip() and isinstance(r[cs], str) and r[cs].strip().lower() != "total":
                    g = lambda c: (r[c] if c is not None and c < len(r) and isinstance(r[c], (int, float)) else np.nan)
                    out.append({"State": r[cs].strip(), "LGA": re.sub(r"\s+lga$", "", r[cl].strip(), flags=re.I), "Target": g(ct), "Total": g(cc), "Used": g(cu)})
    finally:
        wb.close()
    c = pd.DataFrame(out)
    if not len(c):
        return c
    c = c.groupby(["State", "LGA"], as_index=False).agg(Target=("Target", "max"), Total=("Total", lambda x: x.sum(min_count=1)), Used=("Used", lambda x: x.sum(min_count=1)))
    c["k"] = [(cn(a), cn(b_)) for a, b_ in zip(c["State"], c["LGA"])]
    return c


def use_callin(P, ctx, notes, det, name):
    """eTally records: the target population and a cross-check come from the country's call-in workbook for the same campaign (the one whose
    figures agree with the eTally for at least half of the LGAs). Rule: children vaccinated = eTally, target population = call-in."""
    allf = getattr(ctx, "_callin_files", None) or {}
    ctry = P["Country"].dropna().mode()
    files = allf.get(ctry.iat[0], []) if len(ctry) else []
    want = ((getattr(ctx, "_campaign_hints", None) or {}).get(name) or {}).get("callin")
    if want:
        files = [f for f in files if Path(f).name == want]
    elif (getattr(ctx, "_etally_names", None) or {}).get(ctry.iat[0] if len(ctry) else None, 1) > 1:
        return            # several eTally files in the folder: say which call-in workbook goes with this one in campaign_hints.json ("callin")
    cache = getattr(ctx, "_callin_cache", None)
    if cache is None:
        cache = ctx._callin_cache = {}
    k = [(cn(a), cn(b_)) for a, b_ in zip(P["Province"], P["District"])]
    best = None
    for cp in files:
        if cp not in cache:
            try:
                cache[cp] = read_callin(cp)
            except Exception:
                cache[cp] = pd.DataFrame()
        c = cache[cp]
        if not len(c):
            continue
        m = pd.DataFrame({"k": k, "et": pd.to_numeric(P["TotalNbrVaccPolio"], errors="coerce").values}).merge(c[["k", "Target", "Total"]], on="k", how="left")
        ok = m.dropna(subset=["et", "Total"])
        ok = ok[ok["Total"] > 0]
        if len(ok) < 20:
            continue
        agree = float(((ok["et"] / ok["Total"] - 1).abs() <= 0.2).mean())
        if (agree >= 0.5 or want) and (best is None or agree > best[0]):
            best = (agree, cp, c)
    if best is None:
        return
    agree, cp, c = best
    cmp_ = pd.DataFrame({"State": P["Province"].values, "LGA": P["District"].values, "k": k, "eTally": pd.to_numeric(P["TotalNbrVaccPolio"], errors="coerce").values}).merge(
        c[["k", "Target", "Total"]].rename(columns={"Total": "Call_in", "Target": "Call_in_target"}), on="k", how="left")
    cmp_["Gap_pct"] = ((cmp_["eTally"] / cmp_["Call_in"].where(cmp_["Call_in"] > 0) - 1) * 100).round(1)
    cmp_["Verdict"] = np.where(cmp_["Call_in"].isna() | (cmp_["Call_in"] <= 0), "no call-in figure", np.where(cmp_["Gap_pct"].abs() <= 5, "agree (within 5%)",
                               np.where(cmp_["Gap_pct"].abs() <= 20, "differ 5-20%", "differ more than 20%: confirm with the country")))
    det["callin_check"] = cmp_.drop(columns="k")
    det["callin_name"] = Path(cp).name
    n = 0
    for i in P.index:
        t = cmp_.at[i, "Call_in_target"]
        if not pd.isna(t) and t > 0:
            old = P.at[i, "PopPolio"]
            if is_empty(old) or pd.isna(old) or abs(float(old) - float(t)) > 0.5:
                notes.add(rowref(P, i), "PopPolio", "" if is_empty(old) or pd.isna(old) else old, float(t), "target population of the call-in workbook %s (same campaign)" % Path(cp).name)
            P.at[i, "PopPolio"] = float(t)
            n += 1
    det["callin_summary"] = "eTally against the call-in workbook %s: %d of %d LGAs within 20%%, %d differ by more than 20%%; target population of %d LGAs from the call-in" % (
        Path(cp).name, int(round(agree * (cmp_["Call_in"] > 0).sum())), int((cmp_["Call_in"] > 0).sum()), int(cmp_["Verdict"].str.startswith("differ more").sum()), n)


# ----------------------------------------------------------------------------------------------- tally sheets: a check, not a source
def tally_read(path):
    """A Nigeria tally file read as a check (never as admin records): the LGA summaries ('... TallySheet Data Summary', sheet LGA), the ward tally
    (State, Lga, Ward, TOTALOPV ...) and the tally details (sheet TD2_MainDATA). Returns {"kind", "vaccine", "frame"} with one line per LGA
    (State, LGA, Target, Total, Total2, Wards), or None when the file is not one of these."""
    name = Path(path).name
    low = name.lower()
    if "tally" not in low:
        return None
    vacc = "nOPV2" if "nopv2" in low and "bopv" not in low else "bOPV" if "bopv" in low else "MR" if ("measles" in low or "rubella" in low) else None
    out = None
    if low.endswith((".xlsx", ".xlsm")) and "summary" in low and "utilization" not in low:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            if "LGA" not in wb.sheetnames:
                return None
            rows = list(wb["LGA"].iter_rows(values_only=True))
        finally:
            wb.close()
        recs = []
        for r in rows:
            if r and len(r) > 3 and isinstance(r[0], str) and isinstance(r[1], str) and isinstance(r[2], (int, float)) and r[0].strip().lower() != "state":
                nums = [x for x in r[2:] if isinstance(x, (int, float))]
                if len(nums) >= 4:
                    recs.append({"State": r[0].strip(), "LGA": r[1].strip(), "Target": nums[2], "Total": nums[-2], "Total2": np.nan, "Wards": np.nan})
        if recs:
            out = ("LGA summary", pd.DataFrame(recs))
    if out is None:
        try:
            head = pd.read_excel(path, nrows=0)
            cols = {str(c).strip().lower(): c for c in head.columns}
            if {"state", "lga", "ward", "totalopv"} <= set(cols):
                w = pd.read_excel(path)
                w = w.rename(columns={cols[k]: k for k in cols})
                for c in ("totalopv", "eligcild"):
                    if c in w.columns:
                        w[c] = pd.to_numeric(w[c], errors="coerce")
                g = w.groupby(["state", "lga"]).agg(Total=("totalopv", "sum"), Target=("eligcild", "sum") if "eligcild" in w.columns else ("totalopv", "size"),
                                                    Wards=("ward", "count")).reset_index().rename(columns={"state": "State", "lga": "LGA"})
                g["Total2"] = np.nan
                if "eligcild" not in w.columns:
                    g["Target"] = np.nan
                out = ("ward tally", g)
        except Exception:
            out = None
    if out is None and low.endswith((".xls", ".xlsx")):
        try:
            x = pd.read_excel(path, sheet_name="TD2_MainDATA")
            cols = {str(c).strip().lower(): c for c in x.columns}
            if {"state", "lga", "ward"} <= set(cols) and "total_nopv" in cols:
                for c in ("targetpop6wto59mths", "total_nopv", "total_bopv"):
                    if c in cols:
                        x[cols[c]] = pd.to_numeric(x[cols[c]], errors="coerce")
                x["_t"] = x[cols["total_nopv"]].fillna(0) + (x[cols["total_bopv"]].fillna(0) if "total_bopv" in cols else 0)
                g = x.groupby([cols["state"], cols["lga"]]).agg(Target=(cols["targetpop6wto59mths"], "sum") if "targetpop6wto59mths" in cols else ("_t", "size"),
                                                                 Total=(cols["total_nopv"], "sum"), Total2=("_t", "sum"), Wards=(cols["ward"], "count")).reset_index()
                g.columns = ["State", "LGA", "Target", "Total", "Total2", "Wards"]
                out = ("tally details", g)
        except Exception:
            out = None
    if out is None:
        return None
    kind, g = out
    g = g[g["State"].map(lambda v: isinstance(v, str) and v.strip() != "") & g["LGA"].map(lambda v: isinstance(v, str) and v.strip() != "")]
    return {"kind": kind, "vaccine": vacc, "frame": g.reset_index(drop=True)} if len(g) else None


def _tally_candidates(ctx, country, vacc):
    """Groups of records of the country to compare a tally with: {label: DataFrame(key -> Total, Pop)}, from the live file and from the files of the folder."""
    b = ctx.builder
    cands = {}

    def add(label, D):
        D = D.dropna(subset=["Total"])
        if not len(D):
            return
        g = D.groupby("key").agg(Total=("Total", "sum"), Pop=("Pop", "sum"))
        cands[label] = g

    def frame(src):
        keys = [(cn(a), cn(c)) for a, c in zip(src["Province"], src["District"])]
        return pd.DataFrame({"key": keys, "Total": pd.to_numeric(src["TotalNbrVaccPolio"], errors="coerce").values,
                             "Pop": pd.to_numeric(src["PopPolio"], errors="coerce").values if "PopPolio" in src.columns else np.nan})

    def vfilter(v):
        if not vacc or vacc == "MR":
            return True
        fam = vaccine_family(b)(v)
        return (vacc.lower() in {x.lower() for x in fam}) or (not fam)

    L = ctx.live[ctx.live["Country"].map(b._qcountry) == country]
    if len(L):
        d = pd.to_datetime(L["SIA_date"], errors="coerce")
        for (rs, ro, dd), g in L.assign(_d=d).groupby([L["Response"].astype(str), L["Round_Add"].astype(str), "_d"], dropna=True):
            g = g[g["Vaccine_type"].map(vfilter)]
            if len(g) >= 3:
                D = frame(g)
                D["F"] = 0
                add("live file: %s %s (%s)" % (rs, ro, pd.Timestamp(dd).strftime("%Y-%m-%d")), D)
    for h, (nm, R) in (getattr(ctx, "_file_records", {}) or {}).items():
        R = R[R["Country"].map(lambda v: b._qcountry(v) if not is_empty(v) else "") == country]
        if not len(R):
            continue
        R = R.assign(_d=pd.to_datetime(R["SIA_date"], errors="coerce"))
        for (rs, ro, dd), g in R.groupby([R["Response"].astype(str), R["Round_Add"].astype(str), "_d"], dropna=False):
            g = g[g["Vaccine_type"].map(vfilter)]
            if len(g) >= 3:
                add("file %s: %s %s (%s)" % (nm, rs, ro, "" if pd.isna(dd) else pd.Timestamp(dd).strftime("%Y-%m-%d")), frame(g))
    return cands


def tally_compare(ctx, name, T, country):
    """The LGA lines of a tally file against the best matching live campaign and the best matching file of the folder. Returns (DataFrame, message)."""
    F = T["frame"].copy()
    if T["vaccine"] == "MR":
        return pd.DataFrame(), "TALLY SHEET used as a check, not as admin data (%s, %d LGAs): measles-rubella, the admin files carry no measles figures to compare it with" % (T["kind"], len(F))
    cands = _tally_candidates(ctx, country, T["vaccine"])
    universe = {}
    for g in cands.values():
        for k in g.index:
            universe.setdefault(k[0], set()).add(k[1])
    keys, fuzzy = [], 0
    for s_, l_ in zip(F["State"], F["LGA"]):
        k = (cn(s_), cn(re.sub(r"\s+lga$", "", l_, flags=re.I)))
        if k[1] not in universe.get(k[0], ()):
            near = difflib.get_close_matches(k[1], list(universe.get(k[0], ())), n=2, cutoff=0.8)
            if near and (len(near) == 1 or difflib.SequenceMatcher(None, k[1], near[0]).ratio() - difflib.SequenceMatcher(None, k[1], near[1]).ratio() >= 0.05):
                k, fuzzy = (k[0], near[0]), fuzzy + 1
        keys.append(k)
    F["key"] = keys

    def score(g, col):
        m = F.join(g, on="key", rsuffix="_o") if False else F.merge(g, left_on="key", right_index=True, how="left", suffixes=("", "_o"))
        ok = m[col].notna() & m["Total_o"].notna()
        return int((ok & ((m[col] - m["Total_o"]).abs() < 0.5)).sum()), int(ok.sum()), m
    best = {"live": None, "file": None}
    for lab, g in cands.items():
        grp = "live" if lab.startswith("live file") else "file"
        for col in [c for c in ("Total", "Total2") if F[c].notna().any()]:
            ident, both, m = score(g, col)
            ok_ = m[col].notna() & m["Total_o"].notna()
            w5 = int((ok_ & ((m[col] / m["Total_o"].replace(0, np.nan) - 1).abs() <= 0.05)).sum())
            relevant = ident >= 0.1 * len(F) or w5 >= 0.5 * len(F)          # a tally that does not resemble a campaign at all is not compared with it
            if both >= max(3, 0.1 * len(F)) and relevant and (best[grp] is None or (ident, both) > (best[grp][0], best[grp][1])):
                best[grp] = (ident, both, lab, col, m)
    parts, rows = [], []
    for grp in ("live", "file"):
        bst = best[grp]
        if not bst:
            parts.append("%s: no campaign with matching figures (the tally may cover a part only, e.g. a revaccination)" % ("live file" if grp == "live" else "files of the folder"))
            continue
        ident, both, lab, col, m = bst
        for _, r in m.iterrows():
            o, t = r.get("Total_o"), r[col]
            if pd.isna(o) or pd.isna(t):
                gap, v = np.nan, "LGA not in the other source" if pd.isna(o) else "no figure"
            else:
                gap = (t / o - 1) * 100 if o else np.nan
                v = "identical" if abs(t - o) < 0.5 else "within 1%" if abs(gap) <= 1 else "differ 1-5%" if abs(gap) <= 5 else "differ more than 5%: check"
            pg = (r["Target"] / r["Pop"] - 1) * 100 if not pd.isna(r["Target"]) and not pd.isna(r.get("Pop")) and r.get("Pop") else np.nan
            rows.append({"State": r["State"], "LGA": r["LGA"], "Tally_total": t, "Tally_target": r["Target"], "Compared_with": lab, "Other_total": o,
                         "Other_target": r.get("Pop"), "Total_gap_pct": round(gap, 1) if not pd.isna(gap) else np.nan,
                         "Target_gap_pct": round(pg, 1) if not pd.isna(pg) else np.nan, "Verdict": v})
        sub = pd.DataFrame(rows[-len(m):])
        parts.append("%s: %d of %d LGAs identical, %d within 5%%, %d differ more than 5%%, %d not found" % (
            lab, int((sub["Verdict"] == "identical").sum()), len(sub), int(sub["Verdict"].isin(["identical", "within 1%", "differ 1-5%"]).sum()),
            int(sub["Verdict"].str.startswith("differ more").sum()), int(sub["Verdict"].str.startswith("LGA not").sum())))
    msg = "TALLY SHEET used as a check, not as admin data (%s, %d LGAs%s). Against %s" % (T["kind"], len(F), ", %d LGA spelling(s) matched" % fuzzy if fuzzy else "", "; ".join(parts))
    return pd.DataFrame(rows), msg


def wide_rounds_layout(path, folder_country):
    """A sheet with the rounds side by side (Province, District, then per round a target population and a number vaccinated; the campaign dates written above
    the round labels): unpivoted into one record per district and round. Returns (layout, records) or (None, None)."""
    if not str(path).lower().endswith((".xlsx", ".xlsm")):
        return None, None
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets[:3]:
            rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=400, values_only=True)]
            head = rows[:8]
            lab = next((i for i, r in enumerate(head) if sum(1 for c in r if isinstance(c, str) and re.fullmatch(r"R\s*\d", c.strip(), re.I)) >= 4), None)
            if lab is None or not head or not any(isinstance(c, str) and re.match(r"province", c.strip(), re.I) for c in head[0] if c is not None):
                continue
            hdr = head[lab + 1] if lab + 1 < len(head) else []
            kinds = {}
            for j, c in enumerate(hdr):
                t = str(c).lower() if c is not None else ""
                if "target" in t or "pop" in t:
                    kinds[j] = "pop"
                elif "vaccinated" in t or "vaccin" in t:
                    kinds[j] = "vacc"
            rl = {j: re.sub(r"\s", "", str(c)).upper() for j, c in enumerate(head[lab]) if isinstance(c, str) and re.fullmatch(r"R\s*\d", c.strip(), re.I)}
            datemap = {}
            for r in head[:lab]:
                for j, c in enumerate(r):
                    if isinstance(c, str) and j in rl:
                        m = re.search(r"(\d{1,2})\s*-\s*\d{1,2}\s*([A-Za-z]+)\s*(\d{4})", c)
                        if m:
                            dd = pd.to_datetime("%s %s %s" % (m.group(1), m.group(2)[:3], m.group(3)), errors="coerce", format="%d %b %Y")
                            if not pd.isna(dd):
                                datemap[rl[j]] = dd
            pops = {rl[j]: j for j, k in kinds.items() if k == "pop" and j in rl}
            vacs = {rl[j]: j for j, k in kinds.items() if k == "vacc" and j in rl}
            rounds = [r for r in rl.values() if r in pops and r in vacs]
            if not rounds:
                continue
            prov, recs = None, []
            for i, r in enumerate(rows[lab + 2:], lab + 3):
                if r and r[0] is not None and str(r[0]).strip():
                    prov = str(r[0]).strip()
                dist = r[1] if len(r) > 1 else None
                if not prov or not isinstance(dist, str) or not dist.strip() or re.match(r"\s*(grand\s+|sub\s*)?total", dist, re.I) or re.match(r"\s*(grand\s+|sub\s*)?total", prov, re.I):
                    continue
                for rd in sorted(set(rounds)):
                    v = r[vacs[rd]] if vacs[rd] < len(r) else None
                    p = r[pops[rd]] if pops[rd] < len(r) else None
                    if not isinstance(v, (int, float)):
                        continue
                    recs.append({"Country": folder_country, "Province": prov, "District": dist.strip(), "SIA_date": datemap.get(rd, pd.NaT),
                                 "Round_Add": "Rnd" + rd[1:], "PopPolio": p if isinstance(p, (int, float)) else np.nan, "TotalNbrVaccPolio": v,
                                 "_src_row": i, "_sheet": ws.title})
            if len(recs) >= 5:
                rec = pd.DataFrame(recs)
                notes = ["rounds side by side in one sheet (%s): unpivoted into %d records (one per district and round with a number vaccinated); campaign dates read from the sheet: %s" % (
                    ", ".join(sorted(set(rounds))), len(rec), ", ".join("%s %s" % (k, v.strftime("%Y-%m-%d")) for k, v in sorted(datemap.items())) or "none")]
                lay = {"sheet": ws.title, "header_row": lab, "map": {}, "synonym": [], "ignored": [], "missing_keys": [], "class": "rounds side by side (unpivoted)",
                       "template_columns_found": len(rec.columns) - 2, "notes": notes, "extra": [], "etally": False}
                return lay, rec
    finally:
        wb.close()
    return None, None


KINDS = (("ipdredo", "a revisit (REDO) tally of the ward teams"), ("ipdvaccine", "a vaccine vials tally"), ("callindatasource", "a call-in summary of other interventions (NTD, SMC, routine immunisation)"),
         ("lqas", "an LQAS file"), ("monitor", "a monitoring (IM) file"), ("_im_", "a monitoring (IM) file"), ("im_polio", "a monitoring (IM) file"),
         ("etally", "an eTally file"), ("tally", "a tally sheet"), ("aefi", "an AEFI file"), ("supplies", "a supplies file"), ("dqc", "a list of data-quality issues"),
         ("synthese", "a monitoring summary"), ("resultats", "a results summary"), ("daily call", "a call-in report"), ("vaccine utilization", "a vaccine utilisation summary"))


def guess_kind(path, lay):
    import unicodedata
    n = "".join(c for c in unicodedata.normalize("NFD", Path(path).name.lower()) if unicodedata.category(c) != "Mn")
    for k, v in KINDS:
        if k in n:
            return v
    return ""


def prepare_file(path, ctx, prof, template_cols, folder_country=None):
    """Inspect one file. Returns (summary dict, prepared DataFrame in template layout or None, detail dict of frames)."""
    b = ctx.builder
    name = Path(path).name
    s = {"name": name, "size": Path(path).stat().st_size, "hash": sha256_of(path), "layout": "", "sheet": "", "header_row": "",
         "rows_read": 0, "rows_used": 0, "excluded": 0, "new": 0, "existing_same": 0, "existing_different": 0, "existing_units": 0, "possible_dup": 0, "in_file_duplicates_removed": 0,
         "qc_errors": 0, "qc_warnings": 0, "countries": "", "status": "blocked", "why": ""}
    det = {"layout": None, "notes": None, "excluded": [], "issues": None, "compare": None, "diffs": None, "campaigns": None, "dups": None}
    lay, et_raw = None, None
    if "tally" in name.lower() and "etally" not in name.lower() and folder_country:
        try:
            tl = tally_read(path)
        except Exception:
            tl = None
        if tl is not None:
            if not hasattr(ctx, "_tallies"):
                ctx._tallies = {}
            ctx._tallies[s["hash"]] = (name, tl, folder_country)
            s["layout"], s["status"], s["countries"], s["rows_read"] = "tally sheet (%s)" % tl["kind"], "nothing new", folder_country, len(tl["frame"])
            s["why"] = "TALLY SHEET used as a check, not as admin data (%s, %d LGAs): see the Tally_check sheet" % (tl["kind"], len(tl["frame"]))
            det["layout"] = {"class": "tally sheet", "map": {}, "synonym": [], "ignored": [], "missing_keys": [], "notes": [], "header_row": None, "sheet": ""}
            return s, None, det
    if "tally" in name.lower() and name.lower().endswith((".xlsx", ".xlsm")):
        try:
            sh = etally_find(path)
            if sh:
                lay, et_raw = etally_layout(path, sh, folder_country)
        except Exception:
            lay, et_raw = None, None
    if lay is None and folder_country:
        try:
            lay, et_raw = wide_rounds_layout(path, folder_country)
        except Exception:
            lay, et_raw = None, None
    try:
        lay = lay or detect_layout(path, template_cols, assume=(("Country",) if folder_country else ()) + ("Round_Add", "Response"))
    except Exception as e:
        kind_ = guess_kind(path, None)
        damaged = isinstance(e, (AssertionError, IndexError)) and str(path).lower().endswith(".xls")
        if kind_:
            s["codes"] = ["not_admin"]
        s["why"] = ("NOT ADMIN DATA: looks like %s; " % kind_ if kind_ else "") + (
            "the old .xls file is damaged (its list of texts is cut; open it in Excel and save it again as .xlsx if it matters)" if damaged
            else "the file could not be opened (%s)" % type(e).__name__)
        det["layout"] = {"class": "not readable", "map": {}, "synonym": [], "ignored": [], "missing_keys": [], "notes": [str(e)[:200]]}
        return s, None, det
    det["layout"] = lay
    s["layout"], s["header_row"] = lay["class"], (lay["header_row"] + 1) if lay["header_row"] is not None else ""
    s["sheet"] = ", ".join(str(x) for x in ([lay["sheet"]] + [L["sheet"] for L in lay.get("extra") or []]) if x is not None)
    if lay["class"] == "not usable":
        kind = guess_kind(path, lay)
        if kind:
            s["codes"] = ["not_admin"]
            s["why"] = "NOT ADMIN DATA: looks like %s; %s" % (kind, "; ".join(lay["notes"]))
        else:
            s["why"] = "different layout from the AFRO template, not read (admin data in another template is kept visible, never ignored): " + "; ".join(lay["notes"])
        return s, None, det
    raw = et_raw if et_raw is not None else read_records(path, lay)
    s["rows_read"] = len(raw)
    notes = Notes()
    P = pd.DataFrame({c: raw[c] if c in raw.columns else None for c in template_cols}, dtype=object)
    P["_src_row"] = raw["_src_row"].values
    P["_sheet"] = raw["_sheet"].values
    # rows with nothing in the key and figure columns are empty lines; a "TOTAL" line is not a record
    core = [c for c in ("Province", "District", "SIA_date", "PopPolio", "TotalNbrVaccPolio") if c in raw.columns]
    empty_line = raw[core].map(is_empty).all(axis=1).values       # a pre-filled Country / code alone does not make a record
    # a line with neither province nor district is not a record (title lines, summary tables of other sheets)
    if "Province" in raw.columns and "District" in raw.columns:
        empty_line = empty_line | (raw["Province"].map(is_empty) & raw["District"].map(is_empty)).values
    # a line pre-filled with names and a target population but with no date and no vaccination figure was not reported: not a record
    figs = [c for c in ("TotalNbrVaccPolio", "Total_Nbr_0doseVaccPolio_0_59M", "Total_Nbr_1dose_Plus_vaccPolio_0_59M", "Doses.UsedPolio") if c in raw.columns]
    if figs:
        nodate = raw["SIA_date"].map(is_empty).values if "SIA_date" in raw.columns else np.ones(len(raw), bool)
        unrep = nodate & raw[figs].map(lambda v: is_empty(v) or str(v).strip() in ("0", "0.0")).all(axis=1).values & ~empty_line
        if unrep.any():
            lay["notes"].append("%d pre-filled line(s) with no date and no vaccination figure were ignored (not reported yet)" % int(unrep.sum()))
            empty_line = empty_line | unrep
    totals = np.zeros(len(P), bool)
    for c in ("Country", "Province", "District"):
        totals |= P[c].map(lambda v: isinstance(v, str) and bool(re.match(r"^\s*(grand\s+)?total\b", v, re.I))).values
    for i in np.where(totals & ~empty_line)[0]:
        det["excluded"].append({"Row_in_file": rowref(P, i), "Reason": "a total line, not a record"})
    P = P[~empty_line & ~totals].reset_index(drop=True)
    # the Country of the folder fills a Country column that is missing or empty
    if folder_country:
        for i in P.index[P["Country"].map(is_empty)]:
            P.at[i, "Country"] = folder_country
            notes.add(rowref(P, i), "Country", "", folder_country, "country of the folder the file is in")
    # text
    for c in TEXT_COLS:
        for i, v in P[c].items():
            if is_empty(v):
                P.at[i, c] = None
                continue
            t = re.sub(r"\s+", " ", re.sub(r"_x000D_", " ", str(v))).strip()
            if t != v:
                notes.add(rowref(P, i), c, v, t, "spaces / control characters")
            P.at[i, c] = t
    # country and province: the geography reference's spelling
    if ctx.geo is not None:
        std = ctx.geo.drop_duplicates("c_key").set_index("c_key")["country"].to_dict()
        for i, v in P["Country"].items():
            if v is None:
                continue
            t = std.get(b._qcountry(v))
            if t is None and folder_country:
                t = folder_country
                notes.add(rowref(P, i), "Country", v, t, "country name not recognised: the country of the folder is used")
                P.at[i, "Country"] = t
                continue
            if t and t != v:
                notes.add(rowref(P, i), "Country", v, t, "country name of the geography reference")
                P.at[i, "Country"] = t
    if "Admin_3" in P.columns:
        # a sub-district is a name: a figure in that column is a number written in the wrong place (a total shifted from the neighbouring column)
        for i, v in P["Admin_3"].items():
            if not is_empty(v) and re.fullmatch(r"[\d.,\s]+", str(v)):
                notes.add(rowref(P, i), "Admin_3", v, "", "a number is not a sub-district name (figure typed in the wrong column): left empty")
                P.at[i, "Admin_3"] = None
    for c in prof["upper"]:
        for i, v in P[c].items():
            if isinstance(v, str) and v != v.upper():
                notes.add(rowref(P, i), c, v, v.upper(), "capitals, like the live file")
                P.at[i, c] = v.upper()
    if ctx.geo is not None:
        try:
            for i, (old, new) in b.fix_province_spelling(P, ctx.geo).items():
                notes.add(rowref(P, i), "Province", old, new, "province spelling of the geography reference")
                P.at[i, "Province"] = new
        except Exception:
            pass
    # a district with a close spelling in the live file (same province) takes the live spelling
    if ctx.live is not None:
        try:
            align_districts(P, ctx, notes)
        except Exception as e:
            det.setdefault("warnings", []).append("District spelling check failed: %s" % e)
    # vaccine and round, in the spelling of the live file
    vmap = {cn(v): v for v in prof["values"].get("Vaccine_type", ())}
    for i, v in P["Vaccine_type"].items():
        if v is not None and (re.fullmatch(r"[\d.\s]+", str(v)) or str(v).strip().lower() in EMPTY_WORDS):
            notes.add(rowref(P, i), "Vaccine_type", v, "", "not a vaccine name: left empty")
            P.at[i, "Vaccine_type"] = None
            continue
        a = vaccine_alias(v) if v is not None else None
        if a and a != v:
            notes.add(rowref(P, i), "Vaccine_type", v, a, "vaccine name written another way (French / short form)")
            P.at[i, "Vaccine_type"] = v = a
        if isinstance(v, str):
            t = vmap.get(cn(v))
            if t and t != v:
                notes.add(rowref(P, i), "Vaccine_type", v, t, "vaccine spelling of the live file")
                P.at[i, "Vaccine_type"] = t
    if "AgeGPolio" in P.columns:
        # under five years, written another way than the live file's 0-59m (the V2 template says <5Y)
        for i, v in P["AgeGPolio"].items():
            if isinstance(v, str) and re.fullmatch(r"(?i)\s*(<\s*5\s*(y|yr|yrs|years?|ans)?|under\s*5\s*(y|yrs?|years?)?|0\s*-\s*4\s*(y|yrs?|years?|ans)?)\s*", v):
                notes.add(rowref(P, i), "AgeGPolio", v, "0-59m", "under five years, written the way the live file writes it")
                P.at[i, "AgeGPolio"] = "0-59m"
    for i, v in P["Round_Add"].items():
        if v is None:
            continue
        m = re.fullmatch(r"(?i)\s*(?:rnd|round|r)?\s*0*(\d+)\.?0*\s*", str(v))
        if m:
            t = "Rnd%d" % int(m.group(1))
            if t != v:
                notes.add(rowref(P, i), "Round_Add", v, t, "round written as Rnd<number>")
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
                notes.add(rowref(P, i), c, v, t.strftime("%Y-%m-%d"), "date written as text or as a number")
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
                bad_num.append({"Row_in_file": rowref(P, i), "Column": c, "Value": v})
                x = np.nan
            elif how and how != "percent sign kept as it is":
                notes.add(rowref(P, i), c, v, x, how)
            elif how:
                notes.add(rowref(P, i), c, v, x, how)
            out.append(x)
        P[c] = pd.Series(out, index=P.index, dtype=float)
    det["bad_numbers"] = pd.DataFrame(bad_num)
    s["bad_cols"] = dict(Counter(b_["Column"] for b_ in bad_num).most_common(12))
    # eTally has no target population: the latest one of the same LGA in the live file
    if lay.get("etally"):
        pops, nfill = population_from_live(ctx), 0
        for i in P.index:
            if is_empty(P.at[i, "PopPolio"]) or pd.isna(P.at[i, "PopPolio"]):
                v = pops.get((b._qcountry(P.at[i, "Country"]), cn(P.at[i, "Province"]), cn(P.at[i, "District"])))
                if v is not None:
                    P.at[i, "PopPolio"] = float(v)
                    nfill += 1
                    notes.add(rowref(P, i), "PopPolio", "", float(v), "target population of the latest live record of this LGA (an eTally file has none)")
        det.setdefault("warnings", []).append("%d of %d LGA(s) got their target population from the live file" % (nfill, len(P)))
        try:
            use_callin(P, ctx, notes, det, name)
        except Exception as e:
            det.setdefault("warnings", []).append("Call-in comparison failed: %s" % e)
    # campaign values the operator gave for this file (campaign_hints.json: {"file name": {"Response": .., "Round_Add": .., "Vaccine_type": .., "SIA_date": ..}})
    given = (getattr(ctx, "_campaign_hints", None) or {}).get(name)
    if given:
        for col, val in given.items():
            if col not in P.columns:
                continue
            if col == "SIA_date":
                val = pd.to_datetime(val, errors="coerce")
            for i in P.index[P[col].map(lambda x: pd.isna(x) if col == "SIA_date" else is_empty(x))]:
                P.at[i, col] = val
                notes.add(rowref(P, i), col, "", val if col != "SIA_date" else val.strftime("%Y-%m-%d"), "given by the operator for this file (campaign_hints.json)")
    # a Response written with its round ("ANG-2026-08-02_nOPV_NID-Round 2"): the campaign of the lookup and the round apart
    round_hint = {}
    if ctx.lookup is not None:
        try:
            round_hint = split_round_from_response(P, ctx, notes)
        except Exception as e:
            det.setdefault("warnings", []).append("Round written in the campaign name: %s" % e)
    if ctx.lookup is not None:
        try:
            zero_round_placeholder(P, ctx, notes)
        except Exception as e:
            det.setdefault("warnings", []).append("Round 0 check failed: %s" % e)
    # an empty Response: the lookup (the same rules as the build), then the names of the lookup
    if ctx.lookup is not None:
        try:
            dated = P[P["SIA_date"].notna()]
            for i, inf in b.fill_blank_response(dated, ctx.lookup).items():
                P.at[i, "Response"] = inf["set"]["Response"]
                notes.add(rowref(P, i), "Response", "", inf["set"]["Response"], inf["why"])
            for i, inf in b.align_response_names(P[P["SIA_date"].notna()], ctx.lookup, None).items():
                old = P.at[i, "Response"]
                P.at[i, "Response"] = inf["set"]["Response"]
                notes.add(rowref(P, i), "Response", old, inf["set"]["Response"], inf["why"])
        except Exception as e:
            det.setdefault("warnings", []).append("Response filling from the lookup failed: %s" % e)
    # an empty round: the round of the lookup for this country, campaign and start date (the old AFRO template has no round column)
    if ctx.lookup is not None and "Round_Add" in P.columns:
        try:
            lk = ctx.lookup
            for i in P.index[P["Round_Add"].map(is_empty)]:
                resp, d_, c_ = P.at[i, "Response"], P.at[i, "SIA_date"], P.at[i, "Country"]
                if is_empty(resp) or pd.isna(d_) or is_empty(c_):
                    continue
                e = lk[(lk["_ck"] == b._lkc(c_)) & (lk["_rk"] == b._rk(resp)) & (lk["_d"] == pd.Timestamp(d_).strftime("%Y-%m-%d"))]
                rn = list(dict.fromkeys(e["_rn"].dropna())) if len(e) else []
                if len(rn) == 1:
                    rk = b._rk(resp)
                    offs = round_offsets(ctx)
                    off = offs["by_campaign"].get(rk)
                    how = "the data numbers this campaign's rounds %+d from the lookup" % off if off is not None else \
                        "assumed numbering of the live file (%+d from the lookup)" % offs["common"]
                    off = offs["common"] if off is None else off
                    P.at[i, "Round_Add"] = "Rnd%d" % (int(rn[0]) + off)
                    notes.add(rowref(P, i), "Round_Add", "", P.at[i, "Round_Add"], "round taken from the lookup (country, campaign, start date); " + how)
        except Exception as e:
            det.setdefault("warnings", []).append("Round filling from the lookup failed: %s" % e)
        try:
            round_from_hint(P, ctx, notes, round_hint)
        except Exception as e:
            det.setdefault("warnings", []).append("Round from the campaign name failed: %s" % e)
    # a named campaign of the lookup with another date than its planned start: the planned start is kept (the file's date goes in the notes)
    if ctx.lookup is not None:
        try:
            det["date_std"] = standardize_to_lookup(P, ctx, notes)
        except Exception as e:
            det.setdefault("warnings", []).append("Standardising the dates with the lookup failed: %s" % e)
    # still no Response: look for the campaign in the lookup from what the file says (date, round, vaccine, file name)
    try:
        det["facility"] = subdistrict_check(P, ctx, notes)
        if det["facility"]:
            f = det["facility"]
            sel = ctx.live[(ctx.live["Response"].astype(str) == f["Response"]) & (ctx.live["Round_Add"].astype(str) == f["Round_Add"])]
            sel = sel[sel["Country"].map(b._qcountry).isin({b._qcountry(c) for c in P["Country"].dropna().unique()})]
            for col in ("Response", "Round_Add", "Vaccine_type", "SIA_date"):
                vals = sel[col].dropna().astype(str)
                vals = vals[vals.str.strip() != ""]
                if len(vals):
                    v = vals.value_counts().index[0]
                    v = pd.Timestamp(v) if col == "SIA_date" else v
                    m_ = P[col].isna() if col == "SIA_date" else P[col].map(is_empty)
                    for i in P.index[m_]:
                        P.at[i, col] = v
                        notes.add(rowref(P, i), col, "", v, "campaign taken from the live file (the district totals equal it)")
    except Exception as e:
        det["facility"] = None
        det.setdefault("warnings", []).append("Facility-level check failed: %s" % e)
    try:
        det["from_live"] = campaign_from_live(P, ctx, notes)
        det["filled_down"] = fill_down_sheet(P, notes)
        det["inferred"] = infer_campaign(P, ctx, prof, name_hints(path, [lay.get("sheet")] + [L.get("sheet") for L in lay.get("extra") or []]), notes)
    except Exception as e:
        det["inferred"] = []
        det.setdefault("warnings", []).append("Campaign search in the lookup failed: %s" % e)
    # pre-filled lines with no campaign and no vaccination figure at all: planned, not reported yet
    if "TotalNbrVaccPolio" in P.columns and len(P):
        fig0 = pd.to_numeric(P["TotalNbrVaccPolio"], errors="coerce").fillna(0) == 0
        unrep2 = fig0 & P["SIA_date"].isna() & P["Response"].map(is_empty) & P["Round_Add"].map(is_empty) & ~P["Province"].map(is_empty) & ~P["District"].map(is_empty)
        if unrep2.any():
            lay["notes"].append("%d pre-filled line(s) with no campaign and no vaccination figure were ignored (not reported yet)" % int(unrep2.sum()))
            P = P[~unrep2]
    # rows that cannot be used
    bad, camp_missing = [], Counter()
    for i in P.index:
        miss = [c for c in KEY_COLS if (pd.isna(P.at[i, c]) if c == "SIA_date" else is_empty(P.at[i, c]))]
        if miss:
            bad.append((i, "key field empty or unreadable: " + ", ".join(miss)))
            if set(miss) <= {"SIA_date", "Round_Add", "Vaccine_type", "Response"}:
                camp_missing[tuple(miss)] += 1
    if camp_missing:
        # the records are complete except for the campaign: tell the operator what the file says and which campaigns the lookup has
        h = name_hints(path, [lay.get("sheet")] + [L.get("sheet") for L in lay.get("extra") or []])
        cs = P["Country"].dropna().unique()
        cand = ""
        if ctx.lookup is not None and len(cs) == 1:
            lc = ctx.lookup[ctx.lookup["_ck"] == b._lkc(cs[0])].sort_values("_d", ascending=False).head(6)
            cand = " | ".join("%s %s (%s)" % (r["OBR Name"], r["Round Number"], r["Round Start Date"]) for _, r in lc.iterrows())
        fields = sorted({c for k in camp_missing for c in k})
        det["to_set"] = pd.DataFrame([{"Country": ", ".join(map(str, cs)), "Records_without_campaign": sum(camp_missing.values()), "Missing": ", ".join(fields),
                                       "File_says": "; ".join("%s %s" % (k, v) for k, v in h.items() if v) or "nothing usable",
                                       "Latest_campaigns_in_lookup": cand,
                                       "To_give": '"%s": {"Response": "...", "Round_Add": "Rnd1", "Vaccine_type": "...", "SIA_date": "YYYY-MM-DD"}' % name}])
    for i, why in bad:
        det["excluded"].append({"Row_in_file": rowref(P, i), "Reason": why})
    P = P.drop(index=[i for i, _ in bad]).reset_index(drop=True)
    # duplicates inside the file: exact copies are dropped
    if len(P):
        T = to_text(P.drop(columns=["_src_row", "_sheet"]), prof["numeric"])
        cd = b.classify_duplicates(T)
        ex = cd["cls"] == "exact"
        drop = ex & (cd.groupby(["key", "sig"]).cumcount() > 0)
        det["dups"] = pd.DataFrame({"Row_in_file": [rowref(P, j) for j in range(len(P))], "Class": cd["cls"].values, "Records_with_this_key": cd["n"].values})
        det["dups"] = det["dups"][det["dups"]["Class"] != ""]
        s["in_file_duplicates_removed"] = int(drop.sum())
        for i in np.where(drop.values)[0]:
            det["excluded"].append({"Row_in_file": rowref(P, i), "Reason": "exact copy of another record of the file"})
        P = P[~drop.values].reset_index(drop=True)
    s["rows_used"] = len(P)
    s["excluded"] = len(det["excluded"])
    if s["rows_read"] - len(P) - len(det["excluded"]) > 0:
        lay["notes"].append("%d empty line(s) ignored" % (s["rows_read"] - len(P) - len(det["excluded"])))
    if not len(P):
        s["why"] = "no usable record"
        det["notes"] = pd.DataFrame(notes.rows)
        return s, None, det
    s["countries"] = ", ".join(sorted(P["Country"].dropna().unique()))
    try:
        det["round_std"] = align_round_to_lookup(P, ctx, notes)
        for t in det["round_std"]:
            lay["notes"].append("round number set from the lookup: " + t)
    except Exception as e:
        det.setdefault("warnings", []).append("Aligning the round numbers with the lookup failed: %s" % e)
    # comparison with the live file (before the defaults are filled in: only what the country reported is compared)
    det["date_check"] = date_year_check(P, notes)
    det["vaccine_check"] = vaccine_vs_lookup(P, ctx)
    try:
        normalise_integration(P, ctx, notes, det, s)
    except Exception as e:
        s["integrated"] = ""
        det.setdefault("warnings", []).append("Integrated interventions check failed: %s" % e)
    cmp_, diffs = compare_with_live(P, ctx, prof)
    det["compare"], det["diffs"] = cmp_, diffs
    try:
        integration_additions(P, cmp_, ctx, det, s)
    except Exception as e:
        det.setdefault("warnings", []).append("Integrated data for live records failed: %s" % e)
    det["older"] = older_campaign_check(P, cmp_)
    det["_builder"] = b
    if det.get("facility"):
        cmp_["Status"] = "already in the live file, same figures"
        cmp_["Note"] = "health-facility line; its district total equals the live file's " + det["facility"]["Response"] + " " + det["facility"]["Round_Add"]
    s["new"] = int((cmp_["Status"] == "new").sum())
    s["existing_same"] = int((cmp_["Status"] == "already in the live file, same figures").sum())
    s["existing_different"] = int((cmp_["Status"] == "already in the live file, other figures").sum())
    s["existing_units"] = int(cmp_["Status"].str.startswith("same names as several").sum())
    # a record with the same place and the same figures as a live record of another name, round or date is a duplicate: it is removed (never added)
    s["aliases"] = []
    if "Live_campaign" in cmp_.columns and (cmp_["Live_campaign"] != "").any():
        al = cmp_[cmp_["Live_campaign"] != ""].assign(_rs=lambda d: d["Response"].astype(str), _ro=lambda d: d["Round_Add"].astype(str))
        for (rs_, ro_, ls_, lo_), g in al.groupby(["_rs", "_ro", "Live_campaign", "Live_round"]):
            s["aliases"].append({"file_name": rs_, "file_round": ro_, "live_name": ls_, "live_round": lo_, "records": int(len(g))})
    s["dup_removed"] = int(cmp_["Status"].str.startswith("possible duplicate: same figures").sum())
    s["possible_dup"] = int(cmp_["Status"].str.startswith("possible duplicate").sum()) - s["dup_removed"]
    # derived columns and the live file's constants
    yr = P["SIA_date"].dt.year.astype(float)
    mo = P["SIA_date"].dt.month.astype(float)
    for col, ser in (("Year", yr), ("SIAMTH", mo)):
        if col in P.columns:
            cur = pd.to_numeric(P[col], errors="coerce")
            for i in np.where((cur != ser).values)[0]:
                notes.add(rowref(P, i), col, None if pd.isna(cur.iat[i]) else cur.iat[i], ser.iat[i], "derived from SIA_date")
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
                    notes.add(rowref(P, i), col, "", v, "country code / block of the live file")
                elif str(cur).strip() != str(v):
                    P.at[i, col] = v
                    notes.add(rowref(P, i), col, cur, v, "code of another country replaced by the code of the live file")
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
    s["qc_warnings"] = int(iss.loc[(iss["severity"] == "Warning") & ~iss["check"].isin(["coverage_ceiling", "coverage_high"]), "Excel_row"].nunique()) if len(iss) else 0
    # campaigns against the lookup
    if ctx.lookup is not None:
        try:
            rep, _ = b.lookup_crosscheck(P, ctx.lookup)
            det["campaigns"] = rep
        except Exception:
            det["campaigns"] = None
    # the verdict (every finding also has a code, kept in the ledger: the automatic rules read the codes, never the text)
    reasons, codes, infos = [], [], []
    if s["excluded"]:
        reasons.append("%d row(s) left out" % s["excluded"]); codes.append("left_out")
    if s["qc_errors"]:
        reasons.append("%d record(s) with an Error" % s["qc_errors"]); codes.append("error")
    if s["existing_different"]:
        reasons.append("%d record(s) already in the live file with other figures (not changed)" % s["existing_different"]); codes.append("other_figures")
    if s["existing_units"]:
        reasons.append("%d record(s) with the same names as several live records and none identical (not added)" % s["existing_units"]); codes.append("several_units")
    if s.get("aliases"):
        infos.append("records matched to a live record under another campaign name or round (the live record is kept as it is): " +
                     "; ".join("%s %s -> %s %s (%d)" % (a_["file_name"], a_["file_round"], a_["live_name"], a_["live_round"], a_["records"]) for a_ in s["aliases"][:4]))
    if s.get("dup_removed"):
        infos.append("%d duplicate record(s) removed (same place and figures as a live record under another campaign name, round or date): see the Records sheet" % s["dup_removed"])
    for m_ in det.get("date_std", []) or []:
        reasons.append("date standardised to the lookup (please confirm): " + m_); codes.append("date_standardized")
    if s.get("possible_dup"):
        codes.append("possible_dup")
        reasons.append("%d record(s) that look like live records under another campaign name, round or date (not added): see the Records sheet" % s["possible_dup"])
    for m_ in det.get("date_check", {}).get("fixed", []):
        reasons.append("year of the date corrected (please confirm): " + m_); codes.append("year_fixed")
    for m_ in det.get("date_check", {}).get("far", []):
        reasons.append(m_ + ": check"); codes.append("date_far")
    for m_ in det.get("older", []):
        reasons.append(m_); codes.append("older_campaign")
    for m_ in det.get("vaccine_check", []):
        reasons.append("vaccine against the lookup: " + m_); codes.append("vaccine_vs_lookup")
    if lay["class"] != "template":
        reasons.append("layout: " + lay["class"]); codes.append("layout_mapped" if str(lay["class"]).startswith("mapped") else ("etally" if lay.get("etally") else "layout_other"))
    if det.get("from_live"):
        codes.append("campaign_from_live")
        reasons.append("campaign details taken from the live file (same place and number vaccinated), not written in the file (please confirm): " + "; ".join(det["from_live"][:3]))
    if det.get("filled_down"):
        codes.append("campaign_filled_down")
        reasons.append("campaign filled down inside the sheet, not written in those lines (please confirm): " + "; ".join(det["filled_down"][:3]))
    if det.get("inferred"):
        codes.append("campaign_inferred")
        if any("round not in the lookup yet" in str(m_) for m_ in det["inferred"]):
            codes.append("campaign_next_round")    # the lookup has no such round: a person confirms, and the lookup needs the round
        if any("planned start of the previous campaign" in str(m_) for m_ in det["inferred"]):
            codes.append("campaign_by_floor")      # dated after a planned start, before the next one: it can be a round the lookup does not have: a person confirms
        reasons.append("campaign details taken from the lookup / the file name, not written in the file (please confirm): " + "; ".join(det["inferred"][:3]))
    if lay.get("etally"):
        reasons.append("; ".join(lay["notes"][:1])); codes.append("etally")
    if det.get("callin_summary"):
        reasons.append(det["callin_summary"]); codes.append("callin")
    # a file that is mostly another country's records than the folder it is in is in the wrong place
    if folder_country and len(P) >= 10:
        oc = P["Country"].map(lambda v: v != folder_country)
        if oc.mean() > 0.5:
            top_c = P.loc[oc, "Country"].value_counts()
            hint = ("; the records are of %s: move the file to that folder" % top_c.index[0]) if len(top_c) and top_c.iat[0] >= 0.8 * int(oc.sum()) and len(P) < 50 else "; not usable, delete it from the folder"
            reasons.insert(0, "WRONG FILE OR WRONG FOLDER: %d of %d records are of other countries than the folder (%s)%s" % (int(oc.sum()), len(P), folder_country, hint))
            s["wrong_place"] = True; codes.append("wrong_place")
    if lay.get("skipped_sheets"):
        codes.append("sheets_skipped")
        reasons.append("sheet(s) with column names but not usable, not read: %s" % ", ".join(str(x) for x in lay["skipped_sheets"]))
    if folder_country and len(P):
        other = P["Country"].map(lambda v: v != folder_country)
        if other.any():
            others = P.loc[other, "Country"].value_counts().head(3)
            codes.append("other_country")
            reasons.append("%d record(s) of another country than the folder (%s): %s" % (
                int(other.sum()), folder_country, ", ".join("%s %d" % (k, v) for k, v in others.items())))
    if len(det["unseen"]):
        reasons.append("values never seen in the live file"); codes.append("unseen")
    if det["bad_numbers"] is not None and len(det["bad_numbers"]):
        codes.append("non_numbers")
        reasons.append("%d cell(s) that are not numbers" % len(det["bad_numbers"]))
    if len(iss) and "check" in iss.columns:
        nc_ = int(iss.loc[iss["check"] == "coverage_ceiling", "Excel_row"].nunique())
        s["coverage_ceiling"] = nc_
        if nc_:       # for reference only: the live file has its own coverage columns, a high coverage never stops a record
            infos.append("%d record(s) with coverage above the ceiling (kept; for reference)" % nc_)
    if s["qc_warnings"] > 0.25 * s["rows_used"]:
        reasons.append("more than 25% of the records have a warning"); codes.append("warnings")
    if s["new"] == 0:
        reasons.append("no new record"); codes.append("no_new")
    if s.get("integrated"):
        reasons.insert(0, s["integrated"]); codes.append("integrated")
    if s.get("integ_cells"):
        reasons.append("INTEGRATED DATA TO ADD to %d record(s) already in the live file (%d cell(s); the polio figures are not touched): sheet Integrated_additions" % (s["integ_records"], s["integ_cells"]))
        codes.append("integrated_to_add")
    if len(det.get("integ_diff") if det.get("integ_diff") is not None else []):
        infos.append("%d integrated figure(s) differ from the live record (not changed): sheet Integrated_differences" % len(det["integ_diff"]))
    s["codes"] = codes
    # the lookup as evidence: every campaign round of the file is a round of the lookup (same country, campaign, date and round)
    s["lookup_rounds"], s["lookup_ok"] = 0, 0
    try:
        cp = det.get("campaigns")
        if cp is not None and len(cp):
            s["lookup_rounds"], s["lookup_ok"] = int(len(cp)), int(cp["Status"].eq("OK").sum())
    except Exception:
        pass                    # no evidence: the rule then simply does not apply
    s["status"] = "ready" if not reasons else "review"
    if s.get("wrong_place"):
        s["status"] = "blocked"
    if not s.get("wrong_place") and (reasons == ["no new record"] or (s["new"] == 0 and not s["qc_errors"] and not s["existing_different"] and not s["existing_units"] and not s.get("possible_dup") and not s.get("integ_cells"))):
        s["status"] = "nothing new"
    s["why"] = "; ".join(reasons + (["for reference: " + "; ".join(infos)] if infos else []))
    if det.get("facility"):
        s["status"], s["why"] = "nothing new", det["facility"]["message"] + ("; " + s["why"] if s["why"] else "")
    if not hasattr(ctx, "_file_records"):
        ctx._file_records = {}
    ctx._file_records[s["hash"]] = (name, P[[c for c in ("Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response", "TotalNbrVaccPolio", "PopPolio") if c in P.columns]].copy())
    # nothing is lost: every record that is not a duplicate goes in; a record that fails a validation is held back with its reason (list: Held_back_records)
    err_rows = {}
    if len(iss) and "severity" in iss.columns:
        for _, r_ in iss[iss["severity"] == "Error"].iterrows():
            err_rows.setdefault(r_["Excel_row"], []).append(str(r_.get("check_label") or r_.get("check")))
    # a record is held for a text where a figure of the target or of the children vaccinated should be; text in a calculated column (coverage, wastage,
    # proportions, which the build recalculates, or an age-group label) does not hold it (a #DIV/0! of the country's own formulas, for example)
    core_ = {"TotPop", "PopPolio", "TotalNbrVaccPolio", "Tot_Vacc_nOPV2", "Total_Vacc_bOPV"}
    bn_ = det["bad_numbers"]
    bad_rows = set(bn_.loc[bn_["Column"].isin(core_), "Row_in_file"]) if bn_ is not None and len(bn_) else set()
    hold = []
    for i in range(len(P)):
        r_ = []
        rr_ = rowref(P, i)
        if rr_ in err_rows:
            r_.append("validation error: " + ", ".join(sorted(set(err_rows[rr_]))))
        if rr_ in bad_rows:
            r_.append("the target or the number vaccinated is not a number")
        miss = [c for c in ("Country", "Province", "District", "SIA_date", "Round_Add", "Response", "Vaccine_type") if c in P.columns and
                (pd.isna(P.at[i, c]) if c == "SIA_date" else is_empty(P.at[i, c]))]
        if miss:
            r_.append("missing after standardising: " + ", ".join(miss))
        if folder_country and P.at[i, "Country"] != folder_country:
            r_.append("record of another country than the folder")
        hold.append("; ".join(r_))
    hold = np.array(hold, dtype=object)
    isnew = (cmp_["Status"] == "new").values
    other = cmp_["Status"].str.startswith("possible duplicate: same place") | cmp_["Status"].str.startswith("same names as several")
    ho = P[other.values].drop(columns=["_src_row", "_sheet"])[template_cols].copy()
    ho["Hold_reason"] = (cmp_.loc[other.values, "Status"] + ": " + cmp_.loc[other.values, "Note"].astype(str)).values
    co = cmp_.loc[other.values]
    ho["Triage"] = [triage_label(a, b, c, d, e) for a, b, c, d, e in zip(co["Status"], co["Fig_file_vacc"], co["Fig_live_vacc"], co["Fig_file_pop"], co["Fig_live_pop"])]
    ho["Live_vacc"], ho["Live_pop"] = co["Fig_live_vacc"].values, co["Fig_live_pop"].values
    ho["Held_key"] = [held_key(r.Country, r.Province, r.District, r.SIA_date, r.Response, r.Round_Add, r.Vaccine_type) for r in ho.itertuples()]
    closed_ = getattr(ctx, "_closed_keys", set())
    if closed_:                                  # closed by a person as 'duplicate, not added': listed in s['closed_dup'], no longer held
        was = len(ho)
        ho = ho[~ho["Held_key"].isin(closed_)].copy()
        s["closed_dup"] = was - len(ho)
    det["hold"] = hold[isnew]
    det["held_other"] = ho
    s["clean_new"] = int((hold[isnew] == "").sum())
    s["held_new"] = int((hold[isnew] != "").sum()) + len(ho)
    P = P[isnew]                 # only the records the live file does not have go to the build
    P = P.drop(columns=["_src_row", "_sheet"])
    P = P[template_cols]
    return s, P, det


def date_year_check(P, notes):
    """The date of a record one year away from the date of its campaign round (typing 2025 for 2026): when the day and month are
    within 3 days of the date most records of that round have, the year is corrected (listed, to confirm); other dates far from
    the round's date are only reported."""
    out = {"fixed": [], "far": []}
    if not {"SIA_date", "Response", "Round_Add"} <= set(P.columns) or not len(P):
        return out
    keys = P["Response"].fillna("").astype(str) + "|" + P["Round_Add"].fillna("").astype(str)
    for k, g in P.groupby(keys):
        gd = g["SIA_date"].dropna()
        if len(gd) < 5:
            continue
        vc = gd.value_counts()
        mode = vc.index[0]
        if vc.iat[0] < 0.6 * len(gd):
            continue
        fixed, far = [], []
        for i in gd.index[gd != mode]:
            d = P.at[i, "SIA_date"]
            delta = abs((d - mode).days)
            if delta <= 45:
                continue
            new = None
            try:
                new = d.replace(year=mode.year)
            except ValueError:
                pass
            if new is not None and d.year != mode.year and abs((new - mode).days) <= 3:
                P.at[i, "SIA_date"] = new
                notes.add(rowref(P, i), "SIA_date", d.strftime("%Y-%m-%d"), new.strftime("%Y-%m-%d"), "year of the date corrected (the day and month are those of the campaign round)")
                fixed.append(i)
            elif delta > 45:
                far.append(i)
        resp = k.split("|")[0] + " " + k.split("|")[1]
        if fixed:
            out["fixed"].append("%d record(s) of %s had the year of their date changed to %d, the year of the round (%s)" % (
                len(fixed), resp, mode.year, mode.strftime("%Y-%m-%d")))
        if far:
            out["far"].append("%d record(s) of %s dated more than 45 days from the date of the round (%s)" % (len(far), resp, mode.strftime("%Y-%m-%d")))
    return out


def older_campaign_check(P, cmp_):
    """New records of a round dated much earlier than the rest of the file: usually an older round pasted into the workbook."""
    msgs = []
    if not {"SIA_date", "Response", "Round_Add"} <= set(P.columns) or not len(P):
        return msgs
    new = (cmp_["Status"] == "new").values
    keys = P["Response"].fillna("").astype(str) + "|" + P["Round_Add"].fillna("").astype(str)
    med = P.groupby(keys)["SIA_date"].median().dropna()
    if len(med) < 2:
        return msgs
    last = med.max()
    for k, d in med.items():
        n = int(((keys == k).values & new).sum())
        if n >= 5 and (last - d).days >= 150:
            msgs.append("%d new record(s) of %s dated %s, %d days before the rest of the file (%s): check that they belong in this file" % (
                n, k.replace("|", " "), d.strftime("%Y-%m-%d"), (last - d).days, last.strftime("%Y-%m-%d")))
    return msgs


def overlap_between_files(details, out_root, rows, ledger, log=print):
    """The new records that two files both carry: the build keeps the first accepted file's copy, so the report says which files repeat each other."""
    import itertools
    seen, parts = {}, []
    order = [(h, sd) for h, sd in details.items() if sd[0].get("new")]
    for h, (s, det) in order:
        pth = Path(out_root) / ("prepared/%s.xlsx" % h[:12])
        if not pth.is_file():
            continue
        try:
            d = pd.read_excel(pth, dtype=object)
        except Exception:
            continue
        keys, blank = None, None
        try:
            keys, blank = det["_builder"].dup_key(d[[c for c in ("Country", "Province", "District", "Admin_3", "SIA_date", "Round_Add", "Vaccine_type", "Response") if c in d.columns]].copy())
        except Exception:
            continue
        figs = pd.Series("", index=d.index)
        for c in ("PopPolio", "TotalNbrVaccPolio"):
            if c in d.columns:
                figs = figs + "|" + pd.to_numeric(d[c], errors="coerce").round(2).astype(str)
        cnt = {}
        for k, f in zip(keys[~blank], figs[~blank]):
            o = seen.get(k)
            if o is not None and o[0] != h:
                c = cnt.setdefault(o[0], [0, 0])
                c[0] += 1
                c[1] += int(o[1] == f)
        for k, f in zip(keys[~blank], figs[~blank]):
            seen.setdefault(k, (h, f))
        tot = sum(v[0] for v in cnt.values())
        same_tot = sum(v[1] for v in cnt.values())
        s["overlap"] = tot - same_tot          # a copy with the same figures is a duplicate: the build keeps one, nobody needs to look
        s["overlap_same"] = same_tot
        for h2, (n, same) in cnt.items():
            parts.append({"File": s["name"], "Also_in_earlier_file": details[h2][0]["name"], "New_records_in_both": n, "Same_figures": same, "Other_figures": n - same})
        if tot:
            msg = "%d of its %d new record(s) are also new in %s (the build keeps the first accepted copy)" % (
                tot, s["new"], "; ".join("%s (%d)" % (details[h2][0]["name"], v[0]) for h2, v in sorted(cnt.items(), key=lambda kv: -kv[1][0])[:3]))
            if s["overlap"] == 0:
                msg = "for reference: " + msg.replace("(the build keeps the first accepted copy)", "(same figures: duplicates, the build keeps one)")
            s["why"] = (s["why"] + "; " if s["why"] else "") + msg
            if s["status"] == "ready" and s["overlap"]:
                s["status"] = "review"
            for r in rows:
                if r.get("hash") == h:
                    r["why"], r["status"] = s["why"], s["status"]
            if h in ledger["files"]:
                ledger["files"][h]["why"], ledger["files"][h]["status"] = s["why"], s["status"]
                ledger["files"][h]["overlap"] = s["overlap"]
                if s["overlap"]:
                    ledger["files"][h]["codes"] = list(ledger["files"][h].get("codes") or []) + ["overlap"]
            det["_overlap_msg"] = msg
    return pd.DataFrame(parts)


def vaccine_vs_lookup(P, ctx):
    """The vaccine written in the file against the vaccine the lookup has for the same campaign (mOPV2 for nOPV2 is a typo until proven otherwise)."""
    msgs = []
    lk = getattr(ctx, "lookup", None)
    if lk is None or not {"Response", "Vaccine_type", "Country"} <= set(P.columns) or "_vs" not in getattr(lk, "columns", []):
        return msgs
    b = ctx.builder
    for (c, resp), g in P.groupby([P["Country"].fillna(""), P["Response"].fillna("")]):
        if not c or not resp:
            continue
        e = lk[(lk["_ck"] == b._lkc(c)) & (lk["_rk"] == b._rk(resp))]
        allowed = set().union(*e["_vs"]) if len(e) else set()
        used = set().union(*[b._vset(v) for v in g["Vaccine_type"].dropna()]) if g["Vaccine_type"].notna().any() else set()
        if allowed and used and not (used & allowed):
            vals = ", ".join(sorted(g["Vaccine_type"].dropna().astype(str).unique())[:3])
            msgs.append("%s: the file writes the vaccine as %s (%d record(s)), the lookup has %s: a typo?" % (resp, vals, len(g), ", ".join(sorted(allowed))))
    return msgs


def place_index(ctx):
    """{(country, province, district): [(live label, date, vaccine family, total vaccinated, target population, Response, Round_Add)]}"""
    if getattr(ctx, "_place_idx", None) is None:
        b, L = ctx.builder, ctx.live
        cc, pp, dd = L["Country"].map(b._qcountry), L["Province"].map(cn), L["District"].map(cn)
        dts = pd.to_datetime(L["SIA_date"], errors="coerce")
        tot = pd.to_numeric(L["TotalNbrVaccPolio"], errors="coerce")
        pop = pd.to_numeric(L["PopPolio"], errors="coerce")
        vf = L["Vaccine_type"].map(vaccine_family(b))
        idx = {}
        for lab, a, p_, d_, t_, tt, po, v, rs, ro in zip(L.index, cc, pp, dd, dts, tot, pop, vf, L["Response"], L["Round_Add"]):
            if not pd.isna(t_):
                idx.setdefault((a, p_, d_), []).append((lab, t_, v, tt, po, rs, ro))
        ctx._place_idx = idx
    return ctx._place_idx


def place_hits(ctx, country, province, district):
    """The live records of a place: those of the same country, province and district, and those of the same district name filed under another province
    (the live file may use an older or longer province name: 'Central' / 'Central Region', 'Toliary' / 'Androy'). The figures decide in the callers."""
    idx = place_index(ctx)
    if getattr(ctx, "_cd_idx", None) is None:
        cd = {}
        for (c, p_, d_), v in idx.items():
            cd.setdefault((c, d_), []).extend(v)
        ctx._cd_idx = cd
    return list(ctx._cd_idx.get((country, district), []))


def vaccine_family(b):
    return lambda v: frozenset(b._vset(v)) or frozenset({cn(v)}) if isinstance(v, str) and v.strip() else frozenset()


def lookup_round_at(ctx, response, country, date):
    """The round number (numbered as the live file does) that the lookup gives for this campaign and country on this date (3 days of tolerance), or None."""
    lk, b = getattr(ctx, "lookup", None), ctx.builder
    if lk is None or is_empty(response) or is_empty(country) or pd.isna(date):
        return None
    e = lk[(lk["_ck"] == b._lkc(country)) & (lk["_rk"] == b._rk(response))]
    if e.empty:
        return None
    d = (pd.to_datetime(e["_d"], errors="coerce") - pd.Timestamp(date)).abs().dt.days
    e = e[d <= 3]
    rn = list(dict.fromkeys(e["_rn"].dropna()))
    if len(rn) != 1:
        return None
    offs = round_offsets(ctx)
    return int(rn[0]) + offs["by_campaign"].get(e.iloc[0]["_rk"], offs["common"])


def normalise_integration(P, ctx, notes, det, s):
    """INTEGRATED CAMPAIGNS. Recognises, from the FIGURES, which interventions other than polio the file carries (children vaccinated / treated / doses delivered or
    used above zero in the block of the intervention), whatever its flag says: a data focal point may leave 'No' on a block he filled, or 'Yes' on one he did not.
    The flags are written the way the figures say (Yes where there are figures; a 'Yes' with no figure anywhere in the file is left and reported), SIA_Integrated
    becomes Yes on the records with an intervention, and every change is listed in the notes. Returns the summary (a text for the report); det['integrated'] holds
    the table per intervention."""
    b = ctx.builder
    if not len(P):
        return ""
    ev, fl = b.integrated_state(P)
    rows, parts = [], []
    for name, sp in b.INTEGRATED.items():
        n_fig = int(ev[name].sum())
        flag = sp["flag"]
        n_yes = int((fl[name] == "Yes").sum())
        n_fixed = 0
        if flag in P.columns:
            for i in P.index[ev[name].values]:
                cur = P.at[i, flag]
                if b.integ_flag_text(cur) != "Yes":
                    notes.add(rowref(P, i), flag, "" if is_empty(cur) else cur, "Yes", "the file has %s figures on this record: the flag follows the figures" % name)
                    P.at[i, flag] = "Yes"
                    n_fixed += 1
                elif cur != "Yes":
                    notes.add(rowref(P, i), flag, cur, "Yes", "flag written the way the live file writes it")
                    P.at[i, flag] = "Yes"
        elsewhere = [c for c in sp["block"] if c not in ctx.live.columns and c in P.columns]
        if n_fig or n_yes:
            rows.append({"Intervention": name, "Records with figures": n_fig, "Flag said Yes": n_yes, "Flag corrected to Yes": n_fixed,
                         "Flag Yes but no figures": int(((fl[name] == "Yes") & ~ev[name]).sum()),
                         "Columns the live file does not have": ", ".join(elsewhere) if n_fig and elsewhere else ""})
        if n_fig:
            parts.append("%s %d%s" % (name, n_fig, (" (flag said No on %d)" % n_fixed) if n_fixed else ""))
    if parts and "SIA_Integrated" in P.columns:
        anyev = ev.any(axis=1)
        for i in P.index[anyev.values]:
            cur = P.at[i, "SIA_Integrated"]
            if b.integ_flag_text(cur) != "Yes":
                notes.add(rowref(P, i), "SIA_Integrated", "" if is_empty(cur) else cur, "Yes", "the record carries other interventions than polio")
                P.at[i, "SIA_Integrated"] = "Yes"
    det["integrated"] = pd.DataFrame(rows)
    s["integrated_n"] = int(ev.any(axis=1).sum())
    msg = ""
    if parts:
        msg = "INTEGRATED CAMPAIGN: %s (records with figures)" % "; ".join(parts)
        lost = [r for r in rows if r["Columns the live file does not have"]]
        if lost:
            msg += "; figures in columns the live file does not have yet (not kept): " + "; ".join("%s: %s" % (r["Intervention"], r["Columns the live file does not have"]) for r in lost)
    noflag = [r["Intervention"] for r in rows if r["Flag Yes but no figures"] and not r["Records with figures"]]
    if noflag:
        det.setdefault("warnings", []).append("flag says Yes but there is no figure for: " + ", ".join(noflag))
    s["integrated"] = msg
    return msg


def integration_additions(P, cmp_, ctx, det, s):
    """The records of the file that are already in the live file: what the file adds as integrated data (interventions with figures in the file and none in the
    live record), cell by cell. det['integ_add'] is what the build applies (sheet Integrated_additions of the prepared file); det['integ_diff'] lists the
    interventions that both have, with other figures (never changed). The live record is identified by its own values, so the build finds it again."""
    b = ctx.builder
    s["integ_cells"] = s["integ_records"] = 0
    if "Live_index" not in cmp_.columns or not len(P):
        return
    ident_cols = ["Country", "Province", "District", "Admin_3", "SIA_date", "Round_Add", "Vaccine_type", "Response"]
    lines, diffs = [], []
    recs = set()
    for pos in range(len(P)):
        li = cmp_["Live_index"].iat[pos]
        if li == "" or pd.isna(li):
            continue
        li = int(li)
        if li not in ctx.live.index:
            continue
        lr = ctx.live.loc[li]
        frow = P.iloc[pos].to_dict()
        live_row = {c: lr[c] for c in ctx.live.columns}
        ident = {c: ("" if is_empty(lr.get(c)) else (str(lr[c])[:10] if c == "SIA_date" else lr[c])) for c in ident_cols if c in ctx.live.columns}
        ident.update({"Row_in_file": rowref(P, pos), "Live_row": li + 2})
        l_, d_ = b.integrated_additions(frow, live_row, ident)
        if l_:
            recs.add(li)
            lines += l_
        diffs += d_
    det["integ_add"] = pd.DataFrame(lines)
    det["integ_diff"] = pd.DataFrame(diffs)
    s["integ_cells"], s["integ_records"] = len(lines), len(recs)
    if len(lines):
        ants = sorted({a for l in lines for a in str(l["Antigen"]).split(", ") if a})
        mark = {i: True for i in recs}
        for pos in range(len(P)):
            li = cmp_["Live_index"].iat[pos]
            if li != "" and not pd.isna(li) and int(li) in mark:
                cmp_.at[cmp_.index[pos], "Note"] = (str(cmp_["Note"].iat[pos]) + "; " if str(cmp_["Note"].iat[pos]) else "") + "integrated data to add to the live record"


def compare_with_live(P, ctx, prof):
    b = ctx.builder
    T = to_text(P.drop(columns=["_src_row", "_sheet"]), prof["numeric"])
    key, blank = b.dup_key(T)
    rows, diffs = [], []
    cols = [c for c in FIGURE_COLS + ["Date_Report"] if c in P.columns]
    for i in range(len(P)):
        alias = None
        figs = None
        k = key.iat[i]
        cand = ctx.live_index.get(k)
        if not cand:
            st, note = "new", ""
            # same place, same vaccine, and the same figures (or the same date) in the live file under another campaign name, round or date
            pl = place_hits(ctx, b._qcountry(P.at[i, "Country"]), cn(P.at[i, "Province"]), cn(P.at[i, "District"]))
            # the same campaign and round at the same place, only the date, the province or the vaccine name differs: it is the live record
            same_c = [h for h in pl if (not is_empty(P.at[i, "Response"])) and str(h[5]).strip().lower() == str(P.at[i, "Response"]).strip().lower()
                      and str(h[6]).strip().lower() == str(P.at[i, "Round_Add"]).strip().lower()] if pl else []
            aligned = False
            # the same place on the same date (3 days at most; as many days as the live file may have been moved to the planned start of the lookup) with the same vaccine, and only one such live record: it is that live record under another
            # campaign name or round (CAM-2025-11-01 for CAM-2026-01, NOPV2 for bOPV-nOPV2, Rnd3 for Rnd1); the live record is not changed
            if len(same_c) != 1 and pl and not pd.isna(P.at[i, "SIA_date"]):
                fam0 = vaccine_family(b)(P.at[i, "Vaccine_type"])
                cand = [h for h in pl if fam0 and h[2] and (fam0 & h[2]) and abs((pd.Timestamp(P.at[i, "SIA_date"]) - h[1]).days) <= max(3, getattr(b, "LOOKUP_DATE_DAYS", 3))]
                if len(cand) == 1:
                    ro_f, ro_l = P.at[i, "Round_Add"], cand[0][6]
                    same_r = (not is_empty(ro_f)) and str(ro_f).strip().lower() == str(ro_l).strip().lower()
                    # a round that differs is accepted only when the file's round is not a round at all (a date, Rnd0 typed for none), or when the lookup
                    # gives the live record's round for this campaign and date: otherwise it may be another round, and the record is held, never dropped
                    if same_r or is_empty(ro_f) or b._rnum(str(ro_f)) is None or lookup_round_at(ctx, P.at[i, "Response"], P.at[i, "Country"], P.at[i, "SIA_date"]) == b._rnum(str(ro_l)):
                        same_c, aligned = cand, True
            if len(same_c) == 1:
                lab, d_, v, tt, po, rs, ro = same_c[0]
                tot = P.at[i, "TotalNbrVaccPolio"] if "TotalNbrVaccPolio" in P.columns else np.nan
                pop = P.at[i, "PopPolio"] if "PopPolio" in P.columns else np.nan
                dt_ = P.at[i, "SIA_date"]
                dif = []
                if not pd.isna(tot) and (pd.isna(tt) or abs(tot - tt) > 0.5):
                    dif.append(("TotalNbrVaccPolio", "" if pd.isna(tt) else tt, tot))
                if not pd.isna(pop) and (pd.isna(po) or abs(pop - po) > 0.5):
                    dif.append(("PopPolio", "" if pd.isna(po) else po, pop))
                st = "already in the live file, other figures" if dif else "already in the live file, same figures"
                note = "live row %d: same place, campaign and round (%s %s); the date, province or vaccine name differs%s" % (
                    lab + 2, rs, ro, "" if pd.isna(dt_) or abs((pd.Timestamp(dt_) - d_).days) == 0 else " (file %s, live %s)" % (pd.Timestamp(dt_).strftime("%Y-%m-%d"), d_.strftime("%Y-%m-%d")))
                if aligned:
                    note = "live row %d: same place, date and vaccine; the campaign name or round differs (file %s %s, live %s %s)" % (
                        lab + 2, P.at[i, "Response"], P.at[i, "Round_Add"], rs, ro)
                    alias = (str(P.at[i, "Response"]), str(P.at[i, "Round_Add"]), str(rs), str(ro))
                for c_, lv, sv in dif:
                    diffs.append({"Row_in_file": rowref(P, i), "Country": P.at[i, "Country"], "District": P.at[i, "District"],
                                  "SIA_date": "" if pd.isna(dt_) else pd.Timestamp(dt_).strftime("%Y-%m-%d"), "Response": P.at[i, "Response"],
                                  "Column": c_, "Live_value": lv, "Submitted_value": sv, "Live_row": lab + 2})
                pl = []
            if pl and not pd.isna(P.at[i, "SIA_date"]):
                fam = vaccine_family(b)(P.at[i, "Vaccine_type"])
                tot, pop = P.at[i, "TotalNbrVaccPolio"] if "TotalNbrVaccPolio" in P.columns else np.nan, P.at[i, "PopPolio"] if "PopPolio" in P.columns else np.nan
                for lab, d_, v, tt, po, rs, ro in pl:
                    same_camp = (not is_empty(P.at[i, "Response"])) and str(rs).strip().lower() == str(P.at[i, "Response"]).strip().lower() \
                        and str(ro).strip().lower() == str(P.at[i, "Round_Add"]).strip().lower()
                    if not (fam & v if fam and v else False) and not same_camp:
                        continue
                    gap = abs((pd.Timestamp(P.at[i, "SIA_date"]) - d_).days)
                    same_fig = (not pd.isna(tot)) and (not pd.isna(tt)) and tot > 0 and abs(tot - tt) < 0.5
                    if same_fig and gap <= 120:
                        st = "possible duplicate: same figures as a live record of another campaign name, round or date"
                        note = "live row %d (%s %s, %s, %d day(s) apart)" % (lab + 2, rs, ro, d_.strftime("%Y-%m-%d"), gap)
                        if not (fam & v if fam and v else False):
                            note += "; same campaign and round, but the vaccine is named differently (file %s, live %s)" % (P.at[i, "Vaccine_type"], "/".join(sorted(v)) or "?")
                        if not (pd.isna(pop) or pd.isna(po)) and abs(pop - po) >= 0.5:
                            note += "; same children vaccinated but the target population differs (file %d, live %d)" % (pop, po)
                        break
                    if gap <= 10 and not same_fig and st == "new":
                        st = "possible duplicate: same place, date and vaccine as a live record of another campaign name or round, other figures"
                        note = "live row %d (%s %s, %s)" % (lab + 2, rs, ro, d_.strftime("%Y-%m-%d"))
                        figs = (tot, tt, pop, po)          # the figures of both records: the triage of the held records (triage_label)
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
                    diffs.append({"Row_in_file": rowref(P, i), "Country": P.at[i, "Country"], "District": P.at[i, "District"],
                                  "SIA_date": P.at[i, "SIA_date"].strftime("%Y-%m-%d"), "Response": P.at[i, "Response"],
                                  "Column": c, "Live_value": lv, "Submitted_value": sv, "Live_row": best + 2})
        live_ix = None
        if st.startswith("already in the live file"):
            mrow = re.search(r"live row (\d+)", note or "")
            live_ix = int(mrow.group(1)) - 2 if mrow else None
        rows.append({"Row_in_file": rowref(P, i), "Live_index": "" if live_ix is None else live_ix, "Country": P.at[i, "Country"], "Province": P.at[i, "Province"],
                     "District": P.at[i, "District"], "SIA_date": P.at[i, "SIA_date"].strftime("%Y-%m-%d"),
                     "Round_Add": P.at[i, "Round_Add"], "Vaccine_type": P.at[i, "Vaccine_type"], "Response": P.at[i, "Response"],
                     "Status": st, "Note": note, "Live_campaign": alias[2] if alias else "", "Live_round": alias[3] if alias else "",
                     "Fig_file_vacc": "" if figs is None or pd.isna(figs[0]) else figs[0], "Fig_live_vacc": "" if figs is None or pd.isna(figs[1]) else figs[1],
                     "Fig_file_pop": "" if figs is None or pd.isna(figs[2]) else figs[2], "Fig_live_pop": "" if figs is None or pd.isna(figs[3]) else figs[3]})
    return pd.DataFrame(rows), pd.DataFrame(diffs)


TRIAGE_SAME = "Probably the same campaign re-sent"
TRIAGE_FAR = "Figures differ a lot from the live record"
TRIAGE_LOOK = "Needs a look"


def triage_label(status, file_vacc, live_vacc, file_pop, live_pop):
    """A held record that looks like a live record under another campaign name, round or date, sorted for the person who closes it:
    the target within 1% and the children vaccinated within 5% of the live record = probably the same campaign sent again; the target more than 25%
    or the vaccinated more than 50% away = the figures differ a lot (another campaign or a correction); anything else needs a look.
    A record with the same names as several live records always needs a look. The triage only sorts: nothing is closed by it."""
    if not str(status).startswith("possible duplicate"):
        return TRIAGE_LOOK

    def num(x):
        try:
            return float(x)
        except (TypeError, ValueError):
            return float("nan")
    fv, lv, fp, lp = num(file_vacc), num(live_vacc), num(file_pop), num(live_pop)
    vd = abs(fv - lv) / lv if lv == lv and lv > 0 and fv == fv else None
    pdiff = abs(fp - lp) / lp if lp == lp and lp > 0 and fp == fp else None
    if vd is None and pdiff is None:
        return TRIAGE_LOOK
    if (pdiff is not None and pdiff > 0.25) or (vd is not None and vd > 0.5):
        return TRIAGE_FAR
    if (pdiff is None or pdiff <= 0.01) and (vd is None or vd <= 0.05):
        return TRIAGE_SAME
    return TRIAGE_LOOK


def held_key(country, province, district, date, response, round_, vaccine):
    """The identity of a record for the 'closed' list (held_closed.json): the same place, date, campaign, round and vaccine."""
    def t(x):
        return "" if x is None or (isinstance(x, float) and x != x) else str(x).strip().lower()
    d = t(date)[:10]
    return "|".join([t(country), t(province), t(district), d, t(response), t(round_), t(vaccine)])


def load_closed_keys(out_root):
    """The held records a person closed as 'duplicate, not added' (kept by the Operations Center in submissions/held_closed.json): {key}."""
    try:
        d = json.loads((Path(out_root) / "held_closed.json").read_text(encoding="utf-8"))
        return {k for g in d.get("closed", []) for k in g.get("keys", [])}
    except (OSError, ValueError, AttributeError):
        return set()


def quality_issues(P, ctx, prof):
    b = ctx.builder
    ctry = set(P["Country"].dropna().unique())
    live_ctx = ctx.live[ctx.live["Country"].isin(ctry)].copy()
    live_ctx = typed(live_ctx, prof)
    q = P.drop(columns=["_src_row", "_sheet"])
    try:
        res = b.qc_run(q, ctx.geo, b.qc_config(), context=live_ctx if len(live_ctx) else None)
    except Exception as e:
        return pd.DataFrame([{"severity": "Warning", "check": "inspection", "Country": "", "message": "quality checks could not run: %s" % e,
                              "Excel_row": ""}])
    iss = res["issues"].copy()
    iss["Excel_row"] = iss["idx"].map(lambda i: rowref(P, i) if i < len(P) else "")
    iss = iss[~iss["check"].isin(["duplicate_key", "same_key_other_figures"])]
    return iss[["severity", "check", "check_label", "Country", "Province", "District", "SIA_date", "Round", "Vaccine", "Response",
                "message", "current", "suggested", "Excel_row"]].reset_index(drop=True)


def typed(live_text, prof):
    t = live_text.copy()
    for c in prof["numeric"]:
        if c in t.columns:
            t[c] = pd.to_numeric(t[c], errors="coerce")
    import warnings
    with warnings.catch_warnings():       # mixed date formats are read one by one on purpose: no need to warn for each column
        warnings.simplefilter("ignore", UserWarning)
        for c in DATE_COLS:
            if c in t.columns:
                t[c] = pd.to_datetime(t[c], errors="coerce")
    return t


# ----------------------------------------------------------------------------------------------- the folder
def folder_country_of(rel_path, geo, b):
    """The country a sub-folder is named after ('Chad Admin Data', 'Algeria Nov_Dec_2025 Admin Data', 'CAR Admin Data'): the standard
    country name of the geography reference, or None."""
    if geo is None:
        return None
    parts = Path(rel_path).parts[:-1]
    if not parts:
        return None
    std = geo.drop_duplicates("c_key").set_index("c_key")["country"].to_dict()
    alias = {"CAR": "CENTRAL AFRICAN REPUBLIC", "DRC": "DEMOCRATIC REPUBLIC OF THE CONGO", "TANZANIA": "UNITED REPUBLIC OF TANZANIA",
             "COTE D IVOIRE": "COTE D IVOIRE", "IVORY COAST": "COTE D IVOIRE", "CAPE VERDE": "CABO VERDE", "SWAZILAND": "ESWATINI"}
    name = re.sub(r"[^A-Z]+", " ", str(parts[0]).upper().replace("'", " ")).strip()
    cands = sorted(set(list(std) + list(alias)), key=lambda k: -len(k))
    for k in cands:
        if re.search(r"(^| )%s( |$)" % re.escape(k), name):
            return std.get(k) or std.get(b._qcountry(alias[k]))
    return None


# ---- incremental inspection: the result of a file is kept while nothing it depends on changed ----------------------------------------------
CACHE_VERSION = 3


def _frame_hash(df):
    """A hash of the whole content of a table (None when it cannot be hashed: the cache is then not used)."""
    try:
        return hashlib.sha256(pd.util.hash_pandas_object(df.astype(str), index=False).values.tobytes()).hexdigest()[:20]
    except Exception:
        return None


def inspect_context_key(ctx, template_cols, out_root, files):
    """Everything an inspection result depends on besides the file itself: the engine code, the live file, the lookup, the geo reference, the learned headers, the
    campaign hints, the held records closed by a person, and the call-in / eTally files of the folder. None when one part cannot be fingerprinted."""
    parts = [str(CACHE_VERSION), "day:" + dt.date.today().isoformat()]      # some checks depend on today's date
    for mod in (Path(__file__), Path(getattr(ctx.builder, "__file__", "") or "")):
        try:
            parts.append(sha256_of(mod))
        except OSError:
            return None
    for name, df in (("live", ctx.live), ("lookup", ctx.lookup), ("geo", ctx.geo)):
        if df is None:
            parts.append(name + ":none")
            continue
        hh = _frame_hash(df)
        if hh is None:
            return None
        parts.append("%s:%s:%s" % (name, len(df), hh))
    parts.append("cols:" + "|".join(map(str, template_cols)))
    parts.append("syn:" + json.dumps(sorted(SYNONYMS.items())))
    parts.append("hints:" + json.dumps(getattr(ctx, "_campaign_hints", {}), sort_keys=True, default=str))
    parts.append("closed:" + str(len(getattr(ctx, "_closed_keys", ()) or ())) + ":" + hashlib.sha256("|".join(sorted(getattr(ctx, "_closed_keys", ()) or ())).encode()).hexdigest()[:12])
    for p0, rel0 in sorted(files, key=lambda x: x[1]):
        nm = Path(rel0).name.lower()
        if re.search(r"call[\s_\-]*in", nm) or "etally" in nm or "tally" in nm:
            try:
                parts.append("ctxfile:%s:%s" % (rel0, sha256_of(p0)))
            except OSError:
                return None
    return hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()[:24]


def _cache_file(out_root, ctx_key, h, rel, fc):
    k = hashlib.sha256(("%s|%s|%s|%s" % (ctx_key, h, rel, fc or "")).encode("utf-8")).hexdigest()[:24]
    return Path(out_root) / "_inspect_cache" / (k + ".pkl")


def _cache_load(out_root, ctx_key, h, rel, fc):
    import pickle
    if not ctx_key:
        return None
    f = _cache_file(out_root, ctx_key, h, rel, fc)
    try:
        with open(f, "rb") as fh:
            c = pickle.load(fh)
    except Exception:
        return None
    for key in ("prepared", "clean_p"):                 # the files the result points to must still be there
        if c.get(key) and not (Path(out_root) / c[key]).is_file():
            return None
    if c.get("held_n") and not (Path(out_root) / ("prepared/%s_held.xlsx" % h[:12])).is_file():
        return None
    return c


def _cache_save(out_root, ctx_key, h, rel, fc, payload):
    import pickle
    if not ctx_key:
        return
    f = _cache_file(out_root, ctx_key, h, rel, fc)
    try:
        f.parent.mkdir(parents=True, exist_ok=True)
        tmp = f.with_suffix(".tmp")
        with open(tmp, "wb") as fh:
            pickle.dump(payload, fh, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(f)
    except Exception:
        pass


def _cache_sweep(out_root, keep_names):
    """Only the results of the last inspection are useful: a result no longer used is removed."""
    d = Path(out_root) / "_inspect_cache"
    if not d.is_dir():
        return
    for f in d.iterdir():
        if f.name not in keep_names:
            try:
                f.unlink()
            except OSError:
                pass


def inspect_folder(ctx, folder, out_root, template_cols, log=print):
    """Inspect every file of a folder. Writes out_root/Submissions_report.xlsx, out_root/prepared/<hash>.xlsx and updates the ledger."""
    V2_NEW.clear()
    V2_NEW.update(getattr(ctx.builder, "TEMPLATE_V2_NEW", {}))
    out_root = Path(out_root)
    (out_root / "prepared").mkdir(parents=True, exist_ok=True)
    ledger = load_ledger(out_root)
    lp = out_root / "header_synonyms.json"
    if lp.is_file():
        try:
            tn_ = {cn(c): c for c in template_cols}
            learned = {cn(k): v for k, v in json.loads(lp.read_text(encoding="utf-8")).items() if v in template_cols}
            SYNONYMS.update(learned)
            log("learned headers: %d from %s" % (len(learned), lp.name))
        except (OSError, ValueError):
            log("WARNING: %s could not be read" % lp)
    hp = out_root / "campaign_hints.json"
    ctx._campaign_hints = {}
    if hp.is_file():
        try:
            ctx._campaign_hints = {k: v for k, v in json.loads(hp.read_text(encoding="utf-8")).items() if isinstance(v, dict)}
            log("campaign values given for %d file(s) in %s" % (len(ctx._campaign_hints), hp.name))
        except (OSError, ValueError):
            log("WARNING: %s could not be read" % hp)
    ctx._closed_keys = load_closed_keys(out_root)           # held records a person closed as 'duplicate, not added' (held_closed.json)
    if ctx._closed_keys:
        log("%d held record(s) were closed as 'duplicate, not added' by a person (held_closed.json): they are not held again." % len(ctx._closed_keys))
    scratch = out_root / "_unzipped"
    if scratch.is_dir():
        shutil.rmtree(scratch, ignore_errors=True)
    scratch.mkdir(parents=True, exist_ok=True)
    files, not_read = list_files(folder, scratch)
    log("%d data file(s) in %s (%d other file(s) not read)" % (len(files), folder, len(not_read)))
    ctx._callin_files = {}
    for p0, rel0 in files:
        if re.search(r"call[\s_\-]*in", Path(rel0).name, re.I) and str(p0).lower().endswith((".xlsx", ".xlsm")):
            fc0 = folder_country_of(rel0, ctx.geo, ctx.builder)
            if fc0:
                ctx._callin_files.setdefault(fc0, []).append(p0)
    ctx._etally_names = {}
    for p0, rel0 in files:
        if "etally" in Path(rel0).name.lower():
            fc0 = folder_country_of(rel0, ctx.geo, ctx.builder)
            ctx._etally_names.setdefault(fc0, set()).add(Path(rel0).name.lower())
    ctx._etally_names = {k: len(v) for k, v in ctx._etally_names.items()}
    prof = learn_profile(ctx.live, ctx.builder, ctx.geo)
    rows, details = [], {}
    seen_hashes = set()
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    use_cache = os.environ.get("OPS_NO_INSPECT_CACHE", "") not in ("1", "true", "yes")
    ctx_key = inspect_context_key(ctx, template_cols, out_root, files) if use_cache else None
    if use_cache and not ctx_key:
        log("The inspection results cannot be reused this time (a part of the context could not be fingerprinted): every file is read.")
    n_reused, cache_names = 0, set()
    for n_file, (p, rel) in enumerate(files, 1):
        h = sha256_of(p)
        seen_hashes.add(h)
        old = ledger["files"].get(h)
        fc = folder_country_of(rel, ctx.geo, ctx.builder)
        cached = _cache_load(out_root, ctx_key, h, rel, fc)
        if cached is not None:
            n_reused += 1
            cache_names.add(_cache_file(out_root, ctx_key, h, rel, fc).name)
            log("  [%d/%d] unchanged since the last inspection, result reused: %s" % (n_file, len(files), rel))
            s, det = cached["s"], cached["det"]
            det["_builder"] = ctx.builder
            prepared, clean_p, clean_n, held_n = cached["prepared"], cached["clean_p"], cached["clean_n"], cached["held_n"]
            if cached.get("tally"):
                if not hasattr(ctx, "_tallies"):
                    ctx._tallies = {}
                ctx._tallies[h] = cached["tally"]
            if cached.get("records") is not None:
                if not hasattr(ctx, "_file_records"):
                    ctx._file_records = {}
                ctx._file_records[h] = cached["records"]
            if cached.get("gaps"):
                ctx._lookup_gaps = list(getattr(ctx, "_lookup_gaps", None) or []) + list(cached["gaps"])
        else:
            log("  [%d/%d] inspecting %s%s" % (n_file, len(files), rel, ("  (folder country: %s)" % fc) if fc else ""))
            n_gaps0 = len(getattr(ctx, "_lookup_gaps", None) or [])
            s, P, det = prepare_file(p, ctx, prof, template_cols, folder_country=fc)
            s["path"] = rel
            prepared = None
            if P is not None:
                prepared = "prepared/%s.xlsx" % h[:12]
                with pd.ExcelWriter(out_root / prepared, engine="openpyxl") as xw_:
                    P.to_excel(xw_, sheet_name="Sheet1", index=False)
                    ia_ = det.get("integ_add")
                    if ia_ is not None and len(ia_):         # integrated data for records the live file already has: the build applies it (add_coadmin_columns.enrich_records)
                        ia_.to_excel(xw_, sheet_name="Integrated_additions", index=False)
            clean_p, clean_n, held_n = None, 0, 0
            if P is not None and det.get("hold") is not None and len(det["hold"]) == len(P):
                hold_ = pd.Series(det["hold"], index=P.index)
                Pc = P[hold_ == ""]
                Ph = pd.concat([P[hold_ != ""].assign(Hold_reason=hold_[hold_ != ""]), det["held_other"]], ignore_index=True)
                if len(Pc):
                    clean_p = "prepared/%s_clean.xlsx" % h[:12]
                    Pc.to_excel(out_root / clean_p, sheet_name="Sheet1", index=False)
                clean_n, held_n = len(Pc), len(Ph)
                det["held_all"] = Ph
                if len(Ph):
                    Ph.to_excel(out_root / ("prepared/%s_held.xlsx" % h[:12]), sheet_name="Sheet1", index=False)
            if ctx_key:
                cache_names.add(_cache_file(out_root, ctx_key, h, rel, fc).name)
                _cache_save(out_root, ctx_key, h, rel, fc, {
                    "s": s, "det": {k: v for k, v in det.items() if k != "_builder"}, "prepared": prepared, "clean_p": clean_p, "clean_n": clean_n, "held_n": held_n,
                    "tally": (getattr(ctx, "_tallies", {}) or {}).get(h), "records": (getattr(ctx, "_file_records", {}) or {}).get(h),
                    "gaps": list((getattr(ctx, "_lookup_gaps", None) or [])[n_gaps0:])})
        e = old or {"decision": None, "decided_by": "", "decided_at": "", "note": "", "first_seen": now}
        stem_ = name_stem(Path(rel).name)
        if not e.get("decision"):
            prev_ = earlier_decisions(ledger, h, stem_, fc)
            if prev_:
                p0_ = prev_[0]
                same_shape = p0_.get("layout") == s.get("layout") and p0_.get("rows_read") == s.get("rows_read")
                txt_ = "a file with the same name pattern ('%s') was %s by %s on %s%s" % (
                    p0_.get("name"), p0_["decision"], p0_.get("decided_by") or "?", (p0_.get("decided_at") or "")[:10],
                    (": " + p0_["note"]) if p0_.get("note") else "")
                s["why"] = (s["why"] + "; " if s["why"] else "") + txt_
                if p0_["decision"] == "rejected" and same_shape and s["status"] != "blocked":
                    e.update({"decision": "rejected", "decided_by": "app (same layout and number of rows as the rejected file)", "decided_at": now,
                              "note": "decided from the earlier rejection of " + str(p0_.get("name"))})
                    s["why"] += " -> rejected again by the app (same layout and number of rows); accept it by hand if it is right"
        e.update({k: s[k] for k in ("name", "path", "size", "hash", "layout", "rows_read", "rows_used", "excluded", "new", "existing_same",
                                    "existing_different", "existing_units", "qc_errors", "qc_warnings", "countries", "status", "why")})
        e["possible_dup"] = s.get("possible_dup", 0)
        e["codes"], e["lookup_rounds"], e["lookup_ok"] = s.get("codes"), s.get("lookup_rounds", 0), s.get("lookup_ok", 0)
        e["stem"], e["country_folder"] = stem_, fc or ""
        e.update({"prepared": prepared, "last_inspected": now, "present": True, "prepared_clean": clean_p, "clean_new": clean_n, "held_new": held_n,
                  "aliases": s.get("aliases", []), "bad_cols": s.get("bad_cols", {})})
        ledger["files"][h] = e
        rows.append(dict(s, decision=e["decision"] or "", decided_by=e["decided_by"]))
        details[h] = (s, det)
        log("    %s: %s | %d rows, %d new, %d already there (%d with other figures)%s" %
            (s["status"].upper(), s["layout"], s["rows_read"], s["new"], s["existing_same"] + s["existing_different"], s["existing_different"],
             (" | " + s["why"]) if s["why"] else ""))
    ctx._file_records = getattr(ctx, "_file_records", {})
    for i_, ((h_, (s_, det_))) in enumerate(details.items()):
        tl_ = (getattr(ctx, "_tallies", {}) or {}).get(h_)
        if not tl_:
            continue
        try:
            tdf, msg = tally_compare(ctx, tl_[0], tl_[1], ctx.builder._qcountry(tl_[2]))
            det_["tally_check"] = tdf
            s_["why"] = msg
            rows[i_]["why"] = msg
            ledger["files"][h_]["why"] = msg
        except Exception as ex:
            log("    tally check failed for %s: %s" % (tl_[0], ex))
    try:
        ov_ = overlap_between_files(details, out_root, rows, ledger, log)
        if len(ov_):
            ctx._overlap = ov_
            log("%d file(s) repeat new records of an earlier file (sheet Overlap_between_files)" % ov_["File"].nunique())
    except Exception as ex:
        ctx._overlap = None
        log("overlap check failed: %s" % ex)
    # files that were seen before and are not in the folder any more; a newer version of the same name supersedes the older one
    for h, e in ledger["files"].items():
        e["present"] = h in seen_hashes
        newer = [x for k, x in ledger["files"].items() if k != h and x.get("path", x.get("name")) == e.get("path", e.get("name")) and x.get("last_inspected", "") > e.get("last_inspected", "")
                 and k in seen_hashes]
        e["superseded"] = bool(newer) and h not in seen_hashes
    save_ledger(out_root, ledger)
    shutil.rmtree(scratch, ignore_errors=True)
    if ctx_key:
        _cache_sweep(out_root, cache_names)
        log("Incremental inspection: %d of %d file(s) were unchanged (same file, same live file, lookup and rules) and their result was reused; %d read again." % (
            n_reused, len(files), len(files) - n_reused))
    write_report(out_root / "Submissions_report.xlsx", rows, details, ctx, not_read)
    log("report: %s" % (out_root / "Submissions_report.xlsx"))
    return rows


def write_report(path, rows, details, ctx, not_read=()):
    ov = pd.DataFrame(rows)
    cols = ["name", "path", "status", "decision", "layout", "countries", "rows_read", "rows_used", "excluded", "new", "existing_same",
            "existing_different", "existing_units", "possible_dup", "in_file_duplicates_removed", "clean_new", "held_new", "qc_errors", "qc_warnings", "integrated", "integ_records", "why", "hash"]
    ov = ov[[c for c in cols if c in ov.columns]].rename(columns={"name": "File", "path": "Path", "status": "Status", "decision": "Decision",
        "layout": "Layout", "countries": "Countries", "rows_read": "Rows read", "rows_used": "Rows used", "excluded": "Rows left out",
        "new": "New records", "existing_same": "Already there (same)", "existing_different": "Already there (other figures)", "existing_units": "Same names as several live records (not added)", "possible_dup": "Possible duplicates (not added)",
        "in_file_duplicates_removed": "Exact copies removed", "clean_new": "Clean new records (go in)", "held_new": "Held back (see Held_back_records)", "qc_errors": "Records with an Error", "integrated": "Integrated interventions (from the figures)", "integ_records": "Live records to complete with integrated data", "qc_warnings": "Records with a Warning",
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
        lay.append({"File": s["name"], "Layout": L.get("class", ""), "Sheet": s.get("sheet", L.get("sheet", "")), "Header row": s["header_row"],
                    "Template columns found": L.get("template_columns_found", ""), "Missing key columns": ", ".join(L.get("missing_keys", [])),
                    "Mapped by synonym (verify)": "; ".join(L.get("synonym", [])), "Columns not in the template (ignored)": "; ".join(L.get("ignored", [])[:40]),
                    "Notes": "; ".join(L.get("notes", []))})
    sheets["Layout"] = pd.DataFrame(lay)
    # every header the app could not map, over all the files: the way it learns. Copy a line into header_synonyms.json to teach it.
    import difflib
    tcols = {cn(c): c for c in (ctx.live.columns if getattr(ctx, "live", None) is not None else [])}
    known_integ = {cn(c) for sp in getattr(ctx.builder, "INTEGRATED", {}).values() for c in sp["block"]}   # columns of the V2 template the live file does not have yet: known, not to be mapped on another intervention
    cnt = {}
    for h, (s, det) in details.items():
        L = det["layout"] or {}
        if L.get("class") in ("not readable", "not usable") and not L.get("map"):
            continue                                   # a file that is not admin data teaches nothing
        for x in L.get("ignored", []):
            if "(second column for" in str(x) or cn(x) in tcols or cn(x) in SYNONYMS or cn(x) == "checkdp" or cn(x) in known_integ:
                continue                              # a repeated header or a column the app already knows
            x = str(x)
            c = cnt.setdefault(x, {"files": set(), "n": 0})
            c["files"].add(s["name"])
    if cnt:
        rows_h = []
        for x, c in sorted(cnt.items(), key=lambda kv: -len(kv[1]["files"])):
            near = difflib.get_close_matches(cn(x), list(tcols), n=1, cutoff=0.8)
            rows_h.append({"Header": x, "In files": len(c["files"]), "Example file": sorted(c["files"])[0],
                           "Closest template column": tcols[near[0]] if near else "", "To teach it": '"%s": "%s"' % (x, tcols[near[0]]) if near else ""})
        sheets["Headers_to_learn"] = pd.DataFrame(rows_h)
    gaps = getattr(ctx, "_lookup_gaps", None)
    if gaps:
        sheets["Lookup_gaps"] = pd.DataFrame(gaps).drop_duplicates()
    for name, key in (("Campaign_to_set", "to_set"), ("eTally_vs_callin", "callin_check"), ("Tally_check", "tally_check"), ("Records", "compare"), ("Differences_with_live", "diffs"), ("Issues", "issues"), ("Normalised_summary", "note_counts"),
                      ("Integrated_interventions", "integrated"), ("Integrated_additions", "integ_add"), ("Integrated_differences", "integ_diff"), ("Normalised_cells", "notes"), ("Defaults_from_live", "defaults"), ("Held_back_records", "held_all"), ("Left_out", "excluded_df"), ("Not_numbers", "bad_numbers"),
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
    dec = []
    for h, (s_, det_) in details.items():
        base = {"File": s_["name"], "Status": s_["status"]}
        for m_ in det_.get("from_live") or []:
            dec.append(dict(base, Question="Campaign details were not in the file; taken from the live file", Proposed=m_, Where="campaign_hints.json (or accept to keep it)"))
        for m_ in det_.get("filled_down") or []:
            dec.append(dict(base, Question="Campaign copied down inside the sheet", Proposed=m_, Where="accept if right"))
        for m_ in det_.get("inferred") or []:
            dec.append(dict(base, Question="Campaign details taken from the lookup or the file name", Proposed=m_, Where="campaign_hints.json"))
        for m_ in det_.get("vaccine_check") or []:
            dec.append(dict(base, Question="Vaccine differs from the lookup", Proposed=m_, Where="ask the country, or correct in the file"))
        for m_ in (det_.get("date_check") or {}).get("fixed", []):
            dec.append(dict(base, Question="Year of a date corrected", Proposed=m_, Where="see Normalised_cells"))
        for m_ in (det_.get("date_check") or {}).get("far", []):
            dec.append(dict(base, Question="Date far from the round", Proposed=m_, Where="ask the country"))
        for m_ in det_.get("older") or []:
            dec.append(dict(base, Question="Records of an older campaign in the file", Proposed=m_, Where="reject the file or ask the country"))
    if dec:
        sheets["Decisions_needed"] = pd.DataFrame(dec)
    if getattr(ctx, "_overlap", None) is not None and len(ctx._overlap):
        sheets["Overlap_between_files"] = ctx._overlap
    if len(not_read):
        sheets["Files_not_read"] = pd.DataFrame(list(not_read), columns=["File", "Why"])
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


def name_stem(name):
    """A file name without its extension, digits, dates, months, version words: 'Togo_Admin_R2_final_v3.xlsx' and 'togo admin r4 (copy).xlsx' meet."""
    b = re.sub(r"\.[A-Za-z0-9]{2,5}$", "", str(name or "").lower())
    b = re.sub(r"[_\-\.\(\)]+", " ", b)
    b = re.sub(r"\b(final|updated|update|copy|new|version|revised|corrected|clean|cleaned|v\d+|jan\w*|feb\w*|mar\w*|apr\w*|may|jun\w*|jul\w*|aug\w*|sep\w*|oct\w*|nov\w*|dec\w*)\b", " ", b)
    b = re.sub(r"\b(round|rnd|cycle|phase|r)\s*(\d+)\b", r"\1\2", b)      # the round number stays: round 1 and round 2 are not the same file
    b = re.sub(r"\b\d+\b", " ", b)
    return re.sub(r"\s+", " ", b).strip()


def earlier_decisions(ledger, h, stem, country):
    """The decisions already taken on files with the same name pattern (same country folder): [(entry)] newest first."""
    out = []
    for k, e in ledger["files"].items():
        if k == h or not e.get("decision"):
            continue
        if e.get("stem") and e.get("stem") == stem and (e.get("country_folder") or "") == (country or ""):
            out.append(e)
    return sorted(out, key=lambda e: e.get("decided_at", ""), reverse=True)


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
    extra = ""
    if decision == "accepted":
        try:
            extra = learn_campaign(out_root, e)
        except Exception as ex:
            extra = ""
    return True, "%s: %s by %s%s" % (e["name"], decision, by, extra)


def learn_campaign(out_root, e):
    """An accepted file whose campaign details were not written in it: the details the operator accepted are kept in campaign_hints.json
    (for this file name), so the next inspection of the same file does not depend on the live file or the lookup."""
    why = e.get("why") or ""
    if not any(t in why for t in ("campaign details taken from", "campaign filled down")):
        return ""
    pth = Path(out_root) / (e.get("prepared") or "")
    if not pth.is_file():
        return ""
    d = pd.read_excel(pth, dtype=object)
    hint = {}
    for col in ("Response", "Round_Add", "Vaccine_type"):
        if col in d.columns and d[col].notna().any():
            vc = d[col].dropna().astype(str).value_counts()
            if len(vc) == 1:
                hint[col] = vc.index[0]
    if "SIA_date" in d.columns and pd.to_datetime(d["SIA_date"], errors="coerce").notna().any():
        vd = pd.to_datetime(d["SIA_date"], errors="coerce").dropna().dt.strftime("%Y-%m-%d").value_counts()
        if len(vd) == 1:
            hint["SIA_date"] = vd.index[0]
    if not hint:
        return ""
    hp = Path(out_root) / "campaign_hints.json"
    try:
        cur = json.loads(hp.read_text(encoding="utf-8")) if hp.is_file() else {}
    except ValueError:
        cur = {}
    if e["name"] in cur:
        return ""
    cur[e["name"]] = hint
    hp.write_text(json.dumps(cur, indent=1, ensure_ascii=False), encoding="utf-8")
    return " (campaign details kept in campaign_hints.json: %s)" % ", ".join("%s %s" % (k, v) for k, v in hint.items())


# ------------------------------------------------------------------------------------------------ automatic rule 1
# A file nobody has decided on is taken into a build by itself ONLY when everything the inspection found about it is on this list:
# findings about records that are already in the live file (which no submission ever changes) or about the layout, never about
# the new records themselves. Anything else (an Error, a date far from the campaign, a campaign guessed from the live file or the
# file name, values never seen, numbers that are not numbers, a coverage above the ceiling, a possible duplicate, records
# that are also new in another file, an eTally conversion ...) stays for a person. A decision of a person always wins.
AUTO_RULE = "automatic rule 1"
AUTO_INFO = ("row(s) left out",
             "already in the live file with other figures (not changed)",
             "layout: mapped (headers differ)",
             "of the records have a warning")


def auto_reasons_ok(why):
    """True when every finding in 'why' is one of the harmless ones above."""
    parts = [x.strip() for x in (why or "").split("; ") if x.strip()]
    return all(any(t in x for t in AUTO_INFO) for x in parts)


# Automatic rule 2 (needs the codes of an inspection made by this version): the same, and when EVERY campaign round of the file is a
# round of the lookup (same country, campaign, date and round) and the vaccine agrees with the lookup, a campaign that the file did not
# write but the lookup gave, a year typed wrongly that the lookup's date confirms, and an eTally export added up by LGA are harmless too.
AUTO_CODES = {"left_out", "other_figures", "layout_mapped", "warnings"}
AUTO_CODES_LOOKUP = {"campaign_inferred", "year_fixed", "etally", "date_standardized"}


def auto_codes_ok(e):
    """None when the ledger entry has no codes (an older inspection: use the text rule), else True / False."""
    codes = e.get("codes")
    if codes is None:
        return None
    ok = set(AUTO_CODES)
    lk = int(e.get("lookup_rounds") or 0)
    if lk and int(e.get("lookup_ok") or 0) == lk and "vaccine_vs_lookup" not in codes:
        ok |= AUTO_CODES_LOOKUP
    return all(c in ok for c in codes)


# Automatic rule 3, the engine learns from the people: a file with the same name pattern, in the same country folder and with the same layout as a file
# that a PERSON accepted is accepted again, when it shows no kind of finding that the accepted one did not show and none of the HARD ones.
AUTO_HARD = {"integrated", "integrated_to_add", "error", "several_units", "possible_dup", "overlap", "wrong_place", "other_country", "non_numbers", "vaccine_vs_lookup", "sheets_skipped"}


def precedent_for(e, ledger):
    """The newest earlier file accepted by a person that this entry repeats (same pattern, folder and layout, no new kind of finding), or None."""
    codes = e.get("codes")
    if codes is None or not e.get("stem") or set(codes) & AUTO_HARD or e.get("overlap"):
        return None
    for p in earlier_decisions(ledger, None, e.get("stem"), e.get("country_folder")):
        who = str(p.get("decided_by") or "")
        if p.get("decision") != "accepted" or who.startswith("engine") or p.get("layout") != e.get("layout"):
            continue
        if set(codes) <= (set(p.get("codes") or []) | AUTO_CODES | AUTO_CODES_LOOKUP):
            return p
    return None


def auto_candidates(out_root, ledger=None):
    """[(name, path, note)] of the files rule 1 takes into the build: not decided by a person, not blocked, not replaced by a newer
    version, still in the folder, with new records, no Error, no record with several live units, no possible duplicate, no overlap
    with another file, and only the harmless findings of AUTO_INFO."""
    ledger = ledger or load_ledger(out_root)
    out = []
    for h, e in ledger["files"].items():
        if e.get("decision") or e.get("superseded") or not e.get("present", True) or e.get("status") in ("blocked", "nothing new"):
            continue
        if not e.get("prepared") or (e.get("new") or 0) <= 0:
            continue
        if e.get("qc_errors") or e.get("existing_units") or e.get("possible_dup") or e.get("overlap"):
            continue
        by_codes = auto_codes_ok(e)
        prec = precedent_for(e, ledger) if by_codes is False else None
        if prec is not None:
            p = Path(out_root) / e["prepared"]
            if p.is_file():
                out.append((e["name"], str(p), "%s: %d new record(s), the same kind of file as '%s', accepted by %s on %s" % (
                    AUTO_RULE, int(e.get("new") or 0), str(prec.get("name"))[:40], prec.get("decided_by") or "?", str(prec.get("decided_at") or "")[:10])))
            continue
        if by_codes is False or (by_codes is None and not auto_reasons_ok(e.get("why"))):
            continue
        lk = int(e.get("lookup_rounds") or 0)
        how = "only harmless findings"
        if by_codes and lk and int(e.get("lookup_ok") or 0) == lk and set(e.get("codes") or []) & AUTO_CODES_LOOKUP:
            how = "campaign confirmed by the lookup (%d of %d round(s))" % (int(e["lookup_ok"]), lk)
        p = Path(out_root) / e["prepared"]
        if p.is_file():
            out.append((e["name"], str(p), "%s: %d new record(s), %s" % (AUTO_RULE, int(e.get("new") or 0), how)))
    taken = {e.get("prepared") for e in ledger["files"].values() if any(Path(pp).name == Path(str(e.get("prepared"))).name for _, pp, _ in out)}
    out += auto_clean_candidates(out_root, ledger, taken)
    return out


AUTO_RULE4 = "automatic rule 4"
AUTO_HARD4 = {"wrong_place", "vaccine_vs_lookup"}


def auto_clean_candidates(out_root, ledger, taken):
    """Rule 4, nothing is lost: of a file that no rule takes whole, the records that passed every validation and the standardisation and are not duplicates
    (file prepared/<hash>_clean.xlsx) go in; the others are held back with their reason (Held_back_records). Never a file a person decided, never a file in the
    wrong folder or with a vaccine that disagrees with the lookup (the whole file waits for a person)."""
    out = []
    for h, e in ledger["files"].items():
        if e.get("decision") or e.get("superseded") or not e.get("present", True) or e.get("status") in ("blocked", "nothing new"):
            continue
        if e.get("prepared") in taken or not e.get("prepared_clean") or int(e.get("clean_new") or 0) <= 0:
            continue
        codes = e.get("codes")
        if codes is None or set(codes) & AUTO_HARD4:
            continue
        p = Path(out_root) / e["prepared_clean"]
        if p.is_file():
            out.append((e["name"], str(p), "%s: %d clean new record(s) (validated, standardised, not duplicates); %d held back, see Held_back_records" % (
                AUTO_RULE4, int(e["clean_new"]), int(e.get("held_new") or 0))))
    return out


CODE_MEANING = {
    "not_admin": ("a file that is not admin data (IM, LQAS, AEFI, supplies, call-in, summaries)", "nobody: ignored"),
    "overlap": ("the new records are also new in another file", "engine: a copy with the same figures is harmless; other figures need a person"),
    "coverage_ceiling": ("records with more vaccinated than the target allows (reference only, no longer stops a file)", "nobody: the live file has its own coverage columns"),
    "date_standardized": ("records given the planned start date of the lookup for their campaign", "the engine when the lookup confirms it"),
    "possible_dup": ("records that look like live records under another campaign name, round or date", "a person"),
    "unseen": ("names or values never seen in the live file", "a person (new spelling or new area)"),
    "campaign_from_live": ("campaign details taken from the live file, not written in the file", "a person, once (remembered)"),
    "campaign_inferred": ("campaign taken from the lookup / the file name", "the engine when the lookup confirms it"),
    "campaign_by_floor": ("dated after a planned start, before the next one", "a person, once (remembered)"),
    "campaign_next_round": ("round missing from the lookup, taken as the next round of the same campaign", "a person, and the lookup needs the round"),
    "error": ("records with an Error", "the country"),
    "several_units": ("same names as several live records, none identical", "a person"),
    "sheets_skipped": ("a sheet with column names that could not be read", "the country or a person"),
    "non_numbers": ("cells that are not numbers", "the country"),
    "year_fixed": ("year of a date corrected", "the engine when the lookup confirms it"),
    "date_far": ("a date far from the campaign", "a person"),
    "older_campaign": ("records of an older campaign inside the file", "a person"),
    "etally": ("eTally export added up by LGA", "the engine when the lookup confirms the campaign"),
    "callin": ("eTally compared with the call-in workbook", "a person"),
    "layout_other": ("a layout that is not the template", "a person"),
    "vaccine_vs_lookup": ("the vaccine differs from the lookup", "the country or a person"),
    "other_country": ("records of another country than the folder", "a person"),
    "wrong_place": ("a file in the wrong folder", "move or delete the file"),
    "integrated": ("an integrated campaign: other interventions than polio recognised from the figures (flags corrected)", "a person"),
    "integrated_to_add": ("integrated data for records the live file already has (polio only so far)", "a person"),
}


def blockers_of(e):
    """The kinds of finding that keep a file from an automatic decision, in plain words: [code]. None when the entry has no codes (inspect again)."""
    codes = e.get("codes")
    if codes is None:
        return None
    ok = set(AUTO_CODES)
    lk = int(e.get("lookup_rounds") or 0)
    if lk and int(e.get("lookup_ok") or 0) == lk and "vaccine_vs_lookup" not in codes:
        ok |= AUTO_CODES_LOOKUP
    out = [c for c in dict.fromkeys(codes) if c not in ok and c not in ("no_new",)]
    if e.get("overlap") and "overlap" not in out:
        out.append("overlap")
    return out


def summary(out_root, audit=None):
    """What a person would otherwise work out by hand: where the new records are, what keeps them from going in by themselves, who can fix it,
    and which rounds of the lookup that have no data in the live file are in a submitted file. Returns {sheet: DataFrame}."""
    ledger = load_ledger(out_root)
    auto_notes = {n: note for n, _, note in auto_candidates(out_root, ledger)}
    auto = set(auto_notes)
    rows = []
    for h, e in ledger["files"].items():
        if not e.get("present", True) or e.get("superseded") or (e.get("new") or 0) <= 0:
            continue
        bl = blockers_of(e)
        rows.append({"File": e["name"], "Country": e.get("country_folder") or "", "New": int(e["new"]), "Status": e.get("status"),
                     "Decision": e.get("decision") or "", "Auto": e["name"] in auto, "Blockers": bl,
                     "Auto_new": (int(e.get("clean_new") or 0) if str(auto_notes.get(e["name"], "")).startswith(AUTO_RULE4) else int(e["new"])) if e["name"] in auto else 0,
                     "Held_back": int(e.get("held_new") or 0)})
    f = pd.DataFrame(rows)
    out = {}
    if len(f):
        f["Blocker_list"] = f["Blockers"].map(lambda b: "inspect again (no codes)" if b is None else ", ".join(b))
        bl_rows = []
        for code in sorted({c for b in f["Blockers"] if b for c in b}):
            sel = f[f["Blockers"].map(lambda b: bool(b) and code in b) & ~f["Auto"] & (f["Decision"] == "")]
            m = CODE_MEANING.get(code, (code, "a person"))
            bl_rows.append({"Finding": code, "Meaning": m[0], "Who can fix it": m[1], "Files": len(sel), "New_records_affected": int(sel["New"].sum())})
        out["Blockers"] = pd.DataFrame(bl_rows).sort_values("New_records_affected", ascending=False) if bl_rows else pd.DataFrame()
        g = []
        for c, d in f.groupby("Country"):
            todo = d[~d["Auto"] & (d["Decision"] == "")]
            top = pd.Series([x for b in todo["Blockers"] if b for x in b]).value_counts()
            g.append({"Country": c, "Files_with_new_records": len(d), "New_records": int(d["New"].sum()), "Taken_by_rules": int(d["Auto_new"].sum()), "Held_back_records": int(d["Held_back"].sum()),
                      "Decided": int(d.loc[d["Decision"] != "", "New"].sum()), "Waiting_for_a_person": int((d["New"] - d["Auto_new"])[d["Decision"] == ""].sum()),
                      "Main_finding": top.index[0] if len(top) else ""})
        out["By_country"] = pd.DataFrame(g).sort_values("New_records", ascending=False)
        out["Files"] = f.drop(columns=["Blockers"]).sort_values(["Country", "New"], ascending=[True, False])
    # lookup rounds without data in the live file, found in a submitted file?
    rep = Path(out_root) / "Submissions_report.xlsx"
    if audit and Path(audit).is_file() and rep.is_file():
        try:
            nd = pd.read_excel(audit, sheet_name="Lookup_no_data")
            cv = pd.read_excel(rep, sheet_name="Campaigns_vs_lookup")
        except Exception:
            nd, cv = None, None
        if nd is not None and len(nd):
            def rk(x):
                return re.sub(r"[^a-z0-9]", "", str(x).lower())
            def rnum(x):
                m = re.search(r"(\d+)", str(x))
                return int(m.group(1)) if m else None
            cv = cv.assign(_k=cv["Response"].map(rk), _ld=pd.to_datetime(cv["Lookup_start_date"].astype(str).str[:10], errors="coerce"))
            st = {e["name"]: e for e in ledger["files"].values()}
            rec = []
            for _, r in nd.iterrows():
                _sd = pd.to_datetime(str(r["Round Start Date"])[:10], errors="coerce")
                hit = cv[(cv["_k"] == rk(r["OBR Name"])) & cv["_ld"].notna() & ((cv["_ld"] - _sd).abs().dt.days <= 7)]
                files = [x for x in hit["File"].dropna().unique() if x in st and st[x].get("present", True) and not st[x].get("superseded")]
                n_new = max([int(st[x].get("new") or 0) for x in files] or [0])
                rec.append({"Country": r["Country"], "Campaign": r["OBR Name"], "Round": r["Round Number"], "Planned_start": str(r["Round Start Date"])[:10],
                            "In_a_submitted_file": bool(files), "Files": "; ".join(files[:3]), "New_records_in_best_file": n_new,
                            "Next_step": ("the file brings new records: decide it" if n_new else ("the file has the records but all are in the live file or held back: see its findings" if files else "not in any submitted file: ask the country"))})
            r_ = pd.DataFrame(rec)
            r_["_d"] = r_["Planned_start"]
            out["Missing_rounds"] = r_[r_["_d"] >= "2025-01-01"].sort_values(["Country", "_d"]).drop(columns="_d")
    return out


def writable(path, log=print):
    """The path itself, or the same name with the time added when the file is locked (open in Excel): a report is never lost because it was open."""
    p = Path(path)
    try:
        with open(p, "ab"):
            pass
        return p
    except PermissionError:
        alt = p.with_name("%s_%s%s" % (p.stem, dt.datetime.now().strftime("%H%M%S"), p.suffix))
        log("NOTE: %s is open in another program (Excel?): written as %s instead; close it and run again to replace it." % (p.name, alt.name))
        return alt
    except OSError:
        return p


PENDING_DAYS = 15


def _pkey(country, campaign, rnd):
    return "%s|%s|%s" % (str(country).strip().upper(), str(campaign).strip(), str(rnd).strip())


def pending(out_root, nodata, days=PENDING_DAYS, since="2025-01-01", today=None, update=True, source="", live=None, near_days=10, near_min=5):
    """PENDING ADMIN DATA. A lookup round whose admin data are not in the live file `days` days after its start (15 by default) is marked pending, so the
    data can be followed until they arrive. `nodata` is the table 'lookup rounds with no records in the data' (builder.lookup_crosscheck(live, lookup)[1]:
    Country, OBR Name, Round Number, Vaccines, Round Start Date, Round Status, Response_in_data).
    What is pending is told apart: nothing received at all (ask the country), or received in a submitted file but not in the live file yet (a decision, the build or a
    correction is what is missing). A register (pending_register.json) keeps the date each round was first flagged, when it was resolved (the data
    arrived in the live file) and the rounds a person closed (the round did not take place); it is the follow-up list. Rounds that started before `since` are
    kept in a reference sheet, they are old gaps and not pending. With `live` (the live file) a round of a country that has live records within `near_days` days of the
    planned start (at least `near_min`, 5) is not asked of the country: the data may be there under another campaign name or date, so it goes to the sheet Check_other_name (a person aligns
    the names), never to the country's list. Returns {sheet name: DataFrame}."""
    out_root = Path(out_root)
    near = {}
    if live is not None and len(live) and {"Country", "SIA_date"} <= set(live.columns):
        ld = pd.to_datetime(live["SIA_date"].astype(str).str[:10], errors="coerce")
        for c_, g_ in pd.DataFrame({"c": live["Country"].astype(str).str.strip().str.upper(), "d": ld}).dropna().groupby("c"):
            near[c_] = np.sort(g_["d"].values.astype("datetime64[D]"))
    today = pd.Timestamp(today) if today else pd.Timestamp(dt.date.today())
    reg_path = out_root / "pending_register.json"
    try:
        reg = json.loads(reg_path.read_text(encoding="utf-8")) if reg_path.is_file() else {}
    except Exception:
        reg = {}
    reg.setdefault("rounds", {})
    ledger = load_ledger(out_root)
    st = {e["name"]: e for e in ledger["files"].values()}
    rep = out_root / "Submissions_report.xlsx"
    cv = None
    if rep.is_file():
        try:
            cv = pd.read_excel(rep, sheet_name="Campaigns_vs_lookup")
        except Exception:
            cv = None

    def rk(x):
        return re.sub(r"[^a-z0-9]", "", str(x).lower())

    def rnum(x):
        m = re.search(r"(\d+)", str(x))
        return int(m.group(1)) if m else None
    if cv is not None and len(cv):
        # a file is matched to a lookup round by the lookup date of its records, not by the round number the file writes (the live file may number rounds
        # differently from the lookup: a file that says Round 3 on 6 October is the lookup's Round 1)
        cv = cv.assign(_k=cv["Response"].map(rk), _ld=pd.to_datetime(cv["Lookup_start_date"].astype(str).str[:10], errors="coerce"))
    inlive = {}      # (campaign, round number) -> dates of the live records: the same round at another date (a typo in the month or the day)
    if live is not None and len(live) and {"Response", "Round_Add", "SIA_date"} <= set(live.columns):
        _d = pd.to_datetime(live["SIA_date"].astype(str).str[:10], errors="coerce")
        for a_, b_, d_ in zip(live["Response"], live["Round_Add"], _d):
            if not pd.isna(d_):
                inlive.setdefault((rk(a_), rnum(b_)), set()).add(d_)
    rows, old, notyet = [], [], []
    for _, r in (nodata if nodata is not None else pd.DataFrame()).iterrows():
        start = pd.to_datetime(str(r["Round Start Date"])[:10], errors="coerce")
        if pd.isna(start):
            continue
        key = _pkey(r["Country"], r["OBR Name"], r["Round Number"])
        base = {"Country": str(r["Country"]).strip().upper(), "Campaign": r["OBR Name"], "Round": r["Round Number"], "Vaccines": r.get("Vaccines", ""),
                "Planned_start": start.strftime("%Y-%m-%d"), "Due_date": (start + pd.Timedelta(days=days)).strftime("%Y-%m-%d"),
                "Live_file": "campaign is in the live file under other dates" if str(r.get("Response_in_data", "")).startswith("yes") else "no records of this campaign"}
        if start < pd.Timestamp(since):
            old.append(dict(base, Days_since_start=int((today - start).days)))
            continue
        if start + pd.Timedelta(days=days) > today:
            notyet.append(dict(base, Days_to_due=int((start + pd.Timedelta(days=days) - today).days)))
            continue
        files, n_new = [], 0
        if cv is not None and len(cv):
            hit = cv[(cv["_k"] == rk(r["OBR Name"])) & cv["_ld"].notna() & ((cv["_ld"] - start).abs().dt.days <= 7)]
            files = [x for x in hit["File"].dropna().unique() if x in st and st[x].get("present", True) and not st[x].get("superseded")]
            n_new = max([int(st[x].get("new") or 0) for x in files] or [0])
        n_near = 0
        if near.get(base["Country"]) is not None:
            arr = near[base["Country"]]
            n_near = int(((arr >= np.datetime64(start.date() - pd.Timedelta(days=near_days).to_pytimedelta())) & (arr <= np.datetime64(start.date() + pd.Timedelta(days=near_days).to_pytimedelta()))).sum())
        if any(abs((d_ - start).days) <= 35 for d_ in inlive.get((rk(r["OBR Name"]), rnum(r["Round Number"])), ())):
            status, nxt = "In the live file under another date", ("The live file has this campaign and round, but not at the lookup's planned start date (%s): check the date of the live records "
                                                                  "(a typo in the date is the usual cause); a submitted file may carry other figures for it" % start.strftime("%Y-%m-%d"))
        elif not files and n_near >= near_min:
            status, nxt = "Check: live records of the country near the start date", ("The live file has %d record(s) of this country within %d days of the planned start: the data may be there under another campaign name or date; "
                                                                                      "align the name (python admin.py find --country ... --date ...) before asking the country" % (n_near, near_days))
        elif not files:
            status, nxt = "Pending: nothing received", "Ask the country for the admin data of this round (AFRO template), or for confirmation that the round did not take place"
        elif n_new:
            status, nxt = "Received, not in the live file yet", "A submitted file has new records of this round: they go in with the next build (automatic rules or a decision in the Operations Center)"
        else:
            status, nxt = "Received, not in the live file yet", "A submitted file has records of this round but none can be added now (held back or already there under another name): see the country's feedback"
        age = int((today - start).days)
        rows.append(dict(base, Days_since_start=age, Days_overdue=age - days, Age_group=("15-30 days" if age <= 30 else "31-90 days" if age <= 90 else "over 90 days"),
                         Status=status, Live_records_near_start=n_near, Files="; ".join(files[:3]), New_records_in_best_file=n_new, Next_step=nxt, _key=key))
    cur = pd.DataFrame(rows)
    # the register: first flagged, resolved, closed
    now = today.strftime("%Y-%m-%d")
    seen = set(cur["_key"]) if len(cur) else set()
    if update:
        for k in seen:
            e = reg["rounds"].setdefault(k, {"first_flagged": now})
            e["last_seen"] = now
            e.pop("resolved", None)
        for k, e in reg["rounds"].items():
            if k not in seen and not e.get("resolved") and not e.get("closed"):
                e["resolved"] = now
        reg["updated"] = now
        reg["source"] = source
        reg["days"] = days
        reg_path.write_text(json.dumps(reg, indent=1, ensure_ascii=False), encoding="utf-8")
    if len(cur):
        cur["First_flagged"] = cur["_key"].map(lambda k: reg["rounds"].get(k, {}).get("first_flagged", now))
        cur["_closed"] = cur["_key"].map(lambda k: bool(reg["rounds"].get(k, {}).get("closed")))
    out = {}
    cols = ["Country", "Campaign", "Round", "Vaccines", "Planned_start", "Due_date", "Days_since_start", "Days_overdue", "Age_group", "Status", "Live_file",
            "Live_records_near_start", "First_flagged", "Files", "New_records_in_best_file", "Next_step"]
    live_ = cur[~cur["_closed"]] if len(cur) else cur
    pend = live_[live_["Status"] == "Pending: nothing received"] if len(live_) else live_
    chk = live_[live_["Status"].str.startswith("Check:")] if len(live_) else live_
    dd = live_[live_["Status"].str.startswith("In the live file")] if len(live_) else live_
    recv = live_[live_["Status"].str.startswith("Received")] if len(live_) else live_
    out["Pending_now"] = pend[cols].sort_values(["Country", "Planned_start"]) if len(pend) else pd.DataFrame(columns=cols)
    out["Date_differs_in_live"] = dd[cols].sort_values(["Country", "Planned_start"]) if len(dd) else pd.DataFrame(columns=cols)
    out["Check_other_name"] = chk[cols].sort_values(["Country", "Planned_start"]) if len(chk) else pd.DataFrame(columns=cols)
    out["Received_not_in_live"] = recv[cols].sort_values(["Country", "Planned_start"]) if len(recv) else pd.DataFrame(columns=cols)
    out["By_country"] = (pend.groupby("Country").agg(Rounds_pending=("Campaign", "size"), Oldest_start=("Planned_start", "min"), Max_days_since_start=("Days_since_start", "max"))
                         .reset_index().sort_values("Rounds_pending", ascending=False)) if len(pend) else pd.DataFrame(columns=["Country", "Rounds_pending", "Oldest_start", "Max_days_since_start"])
    res = [{"Round": k, "First_flagged": e.get("first_flagged"), "Resolved_on": e.get("resolved")} for k, e in reg["rounds"].items() if e.get("resolved")]
    out["Resolved"] = pd.DataFrame(res, columns=["Round", "First_flagged", "Resolved_on"])
    clo = [{"Round": k, "Closed_on": e["closed"].get("on"), "By": e["closed"].get("by"), "Note": e["closed"].get("note")} for k, e in reg["rounds"].items() if e.get("closed")]
    out["Closed"] = pd.DataFrame(clo, columns=["Round", "Closed_on", "By", "Note"])
    out["Not_yet_due"] = pd.DataFrame(notyet).sort_values("Planned_start") if notyet else pd.DataFrame(columns=["Country", "Campaign", "Round", "Planned_start", "Due_date", "Days_to_due"])
    out["Older_rounds_no_data"] = pd.DataFrame(old).sort_values(["Country", "Planned_start"]) if old else pd.DataFrame(columns=["Country", "Campaign", "Round", "Planned_start"])
    out = {"Read_me": pd.DataFrame({"Pending admin data": [
        "A round of the lookup (GPEI) is PENDING when its admin data are not in the live file %d days after its planned start date (checked on %s, rounds that started on or after %s)." % (days, now, since),
        "Pending_now: nothing received for the round, in no submitted file, and no live record of the country near its start date: ask the country. Date_differs_in_live: the live file has the campaign and round but at another date than the lookup (a person corrects the date). Check_other_name: nothing received, but the live file has records of the country near the start date (the round may be there under another campaign name or date): a person checks and aligns before the country is asked. Received_not_in_live: a submitted file holds records of the round, they are not in the live file yet (next build, a decision, or a correction by the country).",
        "Not_yet_due: rounds that started less than %d days ago, to watch. Older_rounds_no_data: rounds before %s with no data, kept for reference. Resolved: the data arrived in the live file. Closed: a person said the round did not take place or is not expected." % (days, since),
        "First_flagged is the date the engine first saw the round pending (pending_register.json, kept in the workspace). To close a round: python admin.py submissions pending --close \"COUNTRY|campaign|Round n\" --note \"why\".",
        "Source of the live and lookup rounds: %s." % (source or "the last inspection")]}), **out}
    return out


def _scope_index(live, scope, ck):
    """Districts of the live records per (country, campaign, round, date) and the districts the scope plans per (country, campaign, round, start date).
    A district is its geo_code when there is one, else its upper-case name. Returns (live index, scope index) or (None, None) when either is missing."""
    import pandas as pd
    if live is None or scope is None or not len(live) or not len(scope):
        return None, None
    need_l, need_s = {"Country", "Response", "Round_Add", "SIA_date", "District"}, {"country", "response", "roundNumber", "round_start_date", "district"}
    if not need_l <= set(live.columns) or not need_s <= set(scope.columns):
        return None, None
    def dk(df, code, name):
        n = "~" + df[name].fillna("").astype(str).str.strip().str.upper()
        return df[code].where(df[code].notna() & (df[code].astype(str).str.strip() != ""), n) if code in df.columns else n
    L = pd.DataFrame({"c": live["Country"].map(ck), "r": live["Response"].astype(str), "n": live["Round_Add"].astype(str),
                      "d": pd.to_datetime(live["SIA_date"], errors="coerce").dt.strftime("%Y-%m-%d"), "k": dk(live, "geo_code", "District")})
    S = pd.DataFrame({"c": scope["country"].map(ck), "r": scope["response"].astype(str), "n": scope["roundNumber"].astype(str),
                      "d": pd.to_datetime(scope["round_start_date"], errors="coerce"), "k": dk(scope, "geo_code", "district")})
    li = {key: set(g["k"]) for key, g in L.dropna(subset=["d"]).groupby(["c", "r", "n", "d"])}
    si = {}
    for (c, r, n, d), g in S.dropna(subset=["d"]).groupby(["c", "r", "n", "d"]):
        si.setdefault(c, []).append((r, n, d, set(g["k"])))
    return li, si


def _scope_windows(scope, ck, rk):
    """Round windows the scope carries (round_end_date, mop_up_*; added from the GPEI API by scripts/enrich_scope.py): {(country, campaign, round): (start, end, [(mop-up start, mop-up end)])}.
    Empty when the scope has no round_end_date."""
    import pandas as pd
    if scope is None or not len(scope) or "round_end_date" not in scope.columns:
        return {}
    S = scope.copy()
    S["_c"] = S["country"].map(ck)
    S["_r"] = S["response"].map(rk)
    S["_s"] = pd.to_datetime(S["round_start_date"], errors="coerce")
    S["_e"] = pd.to_datetime(S["round_end_date"], errors="coerce")
    for c in ("mop_up_start_date", "mop_up_end_date"):
        S["_" + c] = pd.to_datetime(S[c], errors="coerce") if c in S.columns else pd.NaT
    out = {}
    for (c, r, n), g in S.dropna(subset=["_s", "_e"]).groupby(["_c", "_r", "roundNumber"]):
        x = g.iloc[0]
        mop = [(x["_mop_up_start_date"], x["_mop_up_end_date"])] if pd.notna(x["_mop_up_start_date"]) and pd.notna(x["_mop_up_end_date"]) else []
        out[(c, r, str(n))] = (x["_s"], x["_e"], mop)
    return out


def window_accepts(win, li, si, country_key, camp, rnd, date, rk, slack_days=35, min_districts=5, strong=0.9):
    """Can the GPEI round window place a live round the lookup could not? Returns the reason, or "". (1) the live date is inside the round's window (start to end, or the mop-up window): the lookup only holds the start date;
    (2) the date is within `slack_days` days of the window and the live districts are those the scope plans for that very round (>= `strong`, at least `min_districts` districts): the round ran earlier or later than planned."""
    import pandas as pd
    w = win.get((country_key, rk(camp), str(rnd)))
    if not w:
        return ""
    s, e, mop = w
    d = pd.Timestamp(date)
    if s <= d <= e:
        return "the live date %s is inside the round window in GPEI (%s to %s)" % (str(d)[:10], str(s)[:10], str(e)[:10])
    for ms, me in mop:
        if ms <= d <= me:
            return "the live date %s is inside the mop-up window in GPEI (%s to %s)" % (str(d)[:10], str(ms)[:10], str(me)[:10])
    gap = (s - d).days if d < s else (d - e).days
    L = (li or {}).get((country_key, camp, str(rnd), str(d)[:10]), set())
    if gap <= slack_days and len(L) >= min_districts:
        for r, n, dd, S in si.get(country_key, []) if si else []:
            if rk(r) == rk(camp) and str(n) == str(rnd):
                if len(L & S) / len(L) >= strong:
                    return "the live districts are those planned for this round, which GPEI dates %s to %s (the live date is %d days %s)" % (str(s)[:10], str(e)[:10], gap, "before" if d < s else "after")
                break
    return ""


def scope_evidence(li, si, country_key, camp, rnd, date, days=60, min_districts=5, strong=0.9, same_name=False, rk=None):
    """Which campaign does the scope plan for the districts of these live records? Returns (live districts, text, target campaign or "", share of its planned districts that have records, tied campaigns). The text says: the live districts are those planned for ONE
    campaign (share of the live districts that are planned in it, share of its planned districts that have records); or the scope lists the same districts for two or more
    campaigns (the campaign is entered twice); or no planned campaign of the country within `days` days matches. Rounds of one campaign share their districts, so the round
    whose start date is nearest is taken for each campaign. Too few districts (< min_districts) cannot tell campaigns apart."""
    import pandas as pd
    if li is None:
        return "", "scope not available", "", 0, []
    L = li.get((country_key, camp, rnd, str(date)[:10]), set())
    if len(L) < min_districts:
        return len(L), "too few districts (%d) to tell the campaigns apart" % len(L), "", 0, []
    d0 = pd.Timestamp(date)
    if same_name:
        # the campaign has a name the lookup knows: what does the scope say about THIS campaign (its rounds and districts)?
        rk = rk or (lambda s: str(s).strip().upper())
        mine = [(n, d, S) for r, n, d, S in si.get(country_key, []) if rk(r) == rk(camp)]
        if not mine:
            return len(L), "the scope does not list this campaign", "", 0, []
        union = set().union(*[S for _, _, S in mine])
        return len(L), "the scope lists %s: %s; %d%% of the %d live districts are planned in it" % (
            camp, "; ".join("%s %s (%d districts)" % (n, str(d)[:10], len(S)) for n, d, S in sorted(mine, key=lambda t: t[1])),
            round(100 * len(L & union) / len(L)), len(L)), "", 0, []
    best = {}
    for r, n, d, S in si.get(country_key, []):
        gap = abs((d - d0).days)
        if gap > days:
            continue
        inter = len(L & S)
        if r not in best or gap < best[r][0]:
            best[r] = (gap, n, d, inter / len(L), inter / max(len(S), 1), len(S))
    cand = sorted(((v[3], v[4], r, v) for r, v in best.items() if v[3] >= strong), key=lambda t: (-t[0], -t[1]))
    if not cand:
        return len(L), "no campaign planned in the scope within %d days has %d%% of these %d districts" % (days, int(strong * 100), len(L)), "", 0, []
    top = [t for t in cand if abs(t[0] - cand[0][0]) < 0.02 and abs(t[1] - cand[0][1]) < 0.02]
    if len(top) > 1:
        return len(L), "the scope plans the same districts for %d campaigns (%s): the campaign is entered more than once" % (
            len(top), "; ".join("%s %s, %s" % (t[2], t[3][1], str(t[3][2])[:10]) for t in top)), "", 0, [t[2] for t in top]
    o, cv, r, v = cand[0]
    also = ("; also compatible: " + ", ".join(t[2] for t in cand[1:3])) if len(cand) > 1 else ""
    return len(L), "the districts are those planned for %s %s (start %s): %d%% of the %d live districts are planned in it, %d%% of its %d planned districts have records%s" % (
        r, v[1], str(v[2])[:10], round(o * 100), len(L), round(cv * 100), v[5], also), r, cv, []


def lookup_issues(rep, lookup, since="2025-01-01", ckey=None, rkey=None, near_days=7, same_day=3, live=None, scope=None, accepted=None):
    """LOOKUP ISSUES. Campaign rounds that are in the live file but that the lookup cannot place, written for the people who keep the lookup (the GPEI side) and,
    when it helps, the country. `rep` is the first table of builder.lookup_crosscheck(live, lookup) (one row per campaign round of the data with its status).
    Only two statuses are issues of the lookup: 'Response not in the lookup' and 'Date not in the lookup for this Response' (a difference of up to 7 days is a
    date the build corrects, not an issue). Each is told apart: the campaign is not in the lookup at all; the name differs from ONE lookup campaign that starts within
    `near_days` days (a proposal, a person confirms); TWO or more lookup campaigns start within `near_days` days (the data cannot say which one); the date is the start
    date of ANOTHER lookup campaign of the country; or the lookup lists the campaign but not this round or date. The engine never renames on these findings. Returns
    (issues of rounds starting since `since`, older ones), the same columns. When the scope carries the GPEI round windows (round_end_date) a round whose date the window explains is not an issue: it is appended to the list `accepted`, when given.
    Campaigns GPEI flags as test (is_test) are left out of the scope evidence."""
    import pandas as pd
    cols = ["Country", "Campaign_in_live", "Round_in_live", "Start_in_live", "Records", "Issue", "Lookup_says", "Live_districts", "Scope_says", "Scope_target", "Scope_coverage_pct", "Scope_tie", "Proposal", "Ask", "Message"]
    empty = pd.DataFrame(columns=cols)
    if rep is None or not len(rep) or lookup is None or not len(lookup):
        return empty, empty.copy()
    ck = ckey or (lambda s: str(s).strip().upper())
    rk = rkey or (lambda s: str(s).strip().upper())
    if scope is not None and "is_test" in scope.columns:
        scope = scope[scope["is_test"].astype(str).str.lower() != "true"]
    li, si = _scope_index(live, scope, ck)
    win = _scope_windows(scope, ck, rk)
    lk = lookup.copy()
    lk["_dt"] = pd.to_datetime(lk["_d"], errors="coerce")
    lk["_c"] = lk["Country"].map(ck)
    lk["_n"] = lk["OBR Name"].map(rk)

    def short(x):
        return "%s %s (%s, %s)" % (x["OBR Name"], x["Round Number"], str(x["Round Start Date"])[:10], x.get("Vaccines", ""))
    rows = []
    sel = rep[rep["Status"].isin(["Response not in the lookup", "Date not in the lookup for this Response"])]
    for r in sel.to_dict("records"):
        d = pd.to_datetime(r["SIA_date"], errors="coerce")
        if pd.isna(d):
            continue
        camp, rnd, n = str(r["Response"]), str(r["Round_Add"]), int(r["Records"])
        mine = lk[lk["_c"] == ck(r["Country"])]
        gap = (mine["_dt"] - d).abs().dt.days
        near = mine[gap <= near_days]
        me = "%s %s starting %s (%d record%s)" % (camp, rnd, str(d)[:10], n, "" if n == 1 else "s")
        if r["Status"] == "Response not in the lookup":
            names = sorted(set(near["OBR Name"].astype(str)))
            if not len(names):
                issue, says = "Campaign not in the lookup", "no campaign of the country starts within %d days" % near_days
                prop, ask = "", "Lookup maintainers (add the campaign) and the country (is it a campaign that took place?)"
                msg = "%s is in the live file, but the lookup has no campaign of that name and none of %s starting within %d days. Please add it to the lookup, or tell us which lookup campaign it is." % (me, r["Country"], near_days)
            elif len(names) == 1:
                x = near.iloc[(near["_dt"] - d).abs().argmin()]
                issue, says = "Name differs from the lookup", short(x)
                prop, ask = "If it is the same campaign, use the lookup name %s (a person confirms and decides the correction)" % x["OBR Name"], "Lookup maintainers and the country (confirm they are the same campaign)"
                msg = "%s is in the live file, the lookup has no campaign of that name. The only campaign of the country starting within %d days is %s. Is it the same campaign? If yes, the live name will be aligned to the lookup." % (me, near_days, short(x))
            else:
                issue, says = "Two or more lookup campaigns on this date", "; ".join(short(x) for _, x in near.sort_values("_dt").iterrows())
                prop, ask = "Cannot be decided from the data (names, vaccines and dates do not separate them)", "Lookup maintainers (are they one campaign entered twice, or two?) and the country (which one the records belong to)"
                msg = "%s is in the live file, the lookup has no campaign of that name and %d lookup campaigns start within %d days: %s. Which one do the records belong to? If the lookup holds the same campaign twice, please remove one." % (me, len(names), near_days, says)
        else:
            has = mine[mine["_n"] == rk(camp)]
            other = mine[(gap <= same_day) & (mine["_n"] != rk(camp))]
            lists = "; ".join("%s (%s)" % (x["Round Number"], str(x["Round Start Date"])[:10]) for _, x in has.sort_values("_dt").iterrows())
            if len(other):
                issue = "The date belongs to another lookup campaign"
                says = "%s starts on this date; %s is listed as %s" % ("; ".join(short(x) for _, x in other.iterrows()), camp, lists or "not at all")
                prop, ask = "Cannot be decided from the data", "Lookup maintainers and the country (which campaign the records belong to)"
                msg = "%s: the date is the start date of %s in the lookup, while the name in the live file is %s, which the lookup lists as %s. Which one is right?" % (me, "; ".join(short(x) for _, x in other.iterrows()), camp, lists or "not at all")
            else:
                why = window_accepts(win, li, si, ck(r["Country"]), camp, rnd, d, rk) if win else ""
                if why:
                    if accepted is not None:
                        accepted.append({"Country": r["Country"], "Campaign_in_live": camp, "Round_in_live": rnd, "Start_in_live": str(d)[:10], "Records": n,
                                         "Lookup_says": "the lookup lists %s" % lists, "Engine_reading": why})
                    continue
                issue, says = "Round or date not in the lookup", "the lookup lists %s (nearest %s day(s) away)" % (lists, r.get("Days_difference"))
                prop, ask = "", "Lookup maintainers (add or correct the round) and the country (confirm the date)"
                msg = "%s: the lookup lists %s, nearest %s day(s) away. Please add the round to the lookup, or confirm that the date in the live file is wrong." % (me, lists or "this campaign without rounds", r.get("Days_difference"))
        nd, sev, tgt, cvg, tie = scope_evidence(li, si, ck(r["Country"]), camp, rnd, d, same_name=(issue == "Round or date not in the lookup"), rk=rk)
        if li is not None and sev.startswith("the districts are those planned"):
            msg += " Scope: %s." % sev
        elif li is not None and sev.startswith("the scope lists"):
            msg += " Scope: %s." % sev
        elif li is not None and "entered more than once" in sev:
            msg += " Scope: %s." % sev
            prop = "The scope and the lookup both hold the campaign more than once: one entry should be removed at the source"
        rows.append({"Country": r["Country"], "Campaign_in_live": camp, "Round_in_live": rnd, "Start_in_live": str(d)[:10], "Records": n, "Issue": issue,
                     "Lookup_says": says, "Live_districts": nd, "Scope_says": sev, "Scope_target": tgt, "Scope_coverage_pct": round(100 * cvg), "Scope_tie": " | ".join(tie), "Proposal": prop, "Ask": ask, "Message": msg})
    out = pd.DataFrame(rows, columns=cols)
    if not len(out):
        return empty, empty.copy()
    out = out.sort_values(["Start_in_live", "Country"], ascending=[False, True])
    recent = out["Start_in_live"] >= str(since)
    return out[recent].reset_index(drop=True), out[~recent].reset_index(drop=True)


NAME_ISSUES = ("Name differs from the lookup", "Campaign not in the lookup", "Two or more lookup campaigns on this date", "The date belongs to another lookup campaign")


def propose_fixes(issues, lookup, ckey=None, rkey=None, choose=None, min_districts=20, min_coverage=50, near_days=7, rep=None, vset=None, live=None):  # rep is kept for callers that pass it
    """FROM THE LOOKUP ISSUES TO DECIDED CHANGES. For the rows whose scope evidence is strong (the districts are those planned for ONE campaign of the lookup, and either at
    least `min_districts` live districts or `min_coverage`% of its planned districts have records) it proposes: (1) a corrections line that renames the live campaign to the
    lookup's name (campaign-wide, Country + Response + Round); (2) a lookup change `add` when the lookup lacks the round of that campaign the live records belong to (no round
    of it within `near_days` days and not that round number). For a campaign the scope plans TWICE (the lookup entered it twice: same districts, same start date, same round
    and vaccines) the entry whose planning month (YYYY-MM in its name) is the month of the start date is kept and the other removed from the lookup, when that is unique;
    `choose` = {live campaign: lookup campaign to keep} overrides it.
    Each proposal has a Tier. `auto` = the engine can decide it alone: strong evidence AND the rename changes nothing in how the engine treats the vaccines. The
    co-administration rule reads the NAME of the campaign (a name with nOPV and bOPV makes a single-vaccine record without a partner a co-administered one), so a rename is `auto`
    when both names say the same about the vaccines, or when no single-vaccine record is left without a partner in its district and date; otherwise `review`, with the reason in
    Tier_note (for example the 264 Nigerian records of NIE-2026-01-SNID, whose new name names both vaccines). `vset` and `live` give the vaccines of the live records. Nothing is applied here. Returns (corrections rows, lookup change rows, not proposed [(row text, reason)])."""
    import re
    import pandas as pd
    ck = ckey or (lambda s_: str(s_).strip().upper())
    rk = rkey or (lambda s_: str(s_).strip().upper())
    vs = vset or (lambda s_: set())
    choose = {k.strip().upper(): v.strip() for k, v in (choose or {}).items()}
    lk = lookup.copy()
    lk["_c"] = lk["Country"].map(ck)
    lk["_n"] = lk["OBR Name"].map(rk)
    lk["_dt"] = pd.to_datetime(lk["_d"], errors="coerce")
    has_live = live is not None and len(live) and {"Country", "Response", "Round_Add", "Vaccine_type", "District", "SIA_date"} <= set(live.columns)
    if has_live:
        live = live.assign(_c=live["Country"].map(ck))

    def both(n_):
        """The engine's co-administration rule looks at the NAME of the campaign: a name that holds nOPV and bOPV makes a single-vaccine record without a partner a co-administered one."""
        l_ = str(n_).lower()
        return "nopv" in l_ and "bopv" in l_

    def unpaired(c_, camp_, rnd_):
        """Single-vaccine records of the campaign that have no record of the other vaccine in the same district and date: the only ones the name can change. None when not known."""
        if not has_live:
            return None
        g_ = live[(live["_c"] == c_) & (live["Response"].astype(str) == camp_) & (live["Round_Add"].astype(str) == rnd_)]
        sets_ = g_["Vaccine_type"].map(lambda x: frozenset(vs(x)))
        g_ = g_[sets_.map(len) == 1].assign(_v=[next(iter(x)) for x in sets_[sets_.map(len) == 1]])
        if not len(g_):
            return 0
        # a campaign that shows the "separate records" practice (a district with an nOPV2 record AND a bOPV record on the same date, or a combined record with two different
        # vaccinated totals) keeps its single-vaccine records as they are, whatever its name says
        pairs_ = g_.groupby([g_["District"].astype(str), g_["SIA_date"].astype(str)])["_v"].transform("nunique")
        if (pairs_ >= 2).any():
            return 0
        if {"Tot_Vacc_nOPV2", "Total_Vacc_bOPV"} <= set(live.columns):
            allg = live[(live["_c"] == c_) & (live["Response"].astype(str) == camp_) & (live["Round_Add"].astype(str) == rnd_)]
            two = allg[allg["Vaccine_type"].map(lambda x: len(vs(x)) == 2)]
            if len(two) and (two["Tot_Vacc_nOPV2"].astype(str) != two["Total_Vacc_bOPV"].astype(str)).any():
                return 0
        per = g_.groupby([g_["Province"].astype(str) if "Province" in g_.columns else "", g_["District"].astype(str), g_["SIA_date"].astype(str)])["_v"].transform("nunique")
        return int((per < 2).sum())
    corr, chg, skipped, seen = [], [], [], {}

    def lk_name(name, c):
        h = lk[(lk["_c"] == c) & (lk["_n"] == rk(name))]
        return (h["OBR Name"].iloc[0], h) if len(h) else (None, h)

    def month_gap(name, d0):
        m = re.search(r"(20\d\d)-(\d\d)", str(name))
        if not m:
            return None
        return abs((d0.year - int(m.group(1))) * 12 + d0.month - int(m.group(2)))
    for r in issues.to_dict("records"):
        who = "%s | %s %s | %s" % (r["Country"], r["Campaign_in_live"], r["Round_in_live"], r["Start_in_live"])
        if r["Issue"] not in NAME_ISSUES:
            continue
        c = ck(r["Country"])
        d0 = pd.Timestamp(r["Start_in_live"])
        tie = [t for t in str(r.get("Scope_tie") or "").split(" | ") if t]
        target, drop, tie_note, by_choice = r.get("Scope_target") or "", [], "", False
        if tie:
            keep = choose.get(str(r["Campaign_in_live"]).strip().upper())
            if keep:
                if keep not in tie:
                    skipped.append((who, "--choose names %s, which is not one of the tied campaigns (%s)" % (keep, "; ".join(tie))))
                    continue
                by_choice = True
            else:
                rows_ = pd.concat([lk_name(t, c)[1] for t in tie])
                same = rows_["Vaccines"].astype(str).nunique() == 1 and rows_["_d"].nunique() == 1 and rows_["_rn"].nunique() == 1
                gaps = {t: month_gap(t, d0) for t in tie}
                best = sorted((g, t) for t, g in gaps.items() if g is not None)
                if same and len(best) == len(tie) and (len(best) == 1 or best[0][0] < best[1][0]):
                    keep = best[0][1]
                    tie_note = "kept %s: the planning month in its name is the start month (%s); the other entries are the same campaign (same start date, round, vaccines, districts)" % (keep, str(d0)[:7])
                else:
                    skipped.append((who, "the scope plans the same districts for %d campaigns (%s) and nothing in the lookup separates them: say which to keep with --choose \"%s=<lookup name>\"" % (
                        len(tie), "; ".join(tie), r["Campaign_in_live"])))
                    continue
            target, drop = keep, [t for t in tie if t != keep]
        if not target:
            skipped.append((who, "no campaign of the scope matches the districts (%s)" % str(r.get("Scope_says"))[:90]))
            continue
        nd, cvg = int(r["Live_districts"] or 0), float(r.get("Scope_coverage_pct") or 0)
        if not tie and nd < min_districts and cvg < min_coverage:
            skipped.append((who, "weak evidence: %d districts, %d%% of the planned districts of %s have records" % (nd, cvg, target)))
            continue
        name, hit = lk_name(target, c)
        if name is None:
            skipped.append((who, "%s is planned in the scope but is not in the lookup" % target))
            continue
        key = (c, str(r["Campaign_in_live"]), str(r["Round_in_live"]))
        if key in seen and seen[key] != name:
            skipped.append((who, "the same live campaign and round would go to two names (%s and %s)" % (seen[key], name)))
            continue
        # tier: the vaccines of the records against the vaccines the lookup lists for the target
        if by_choice:
            tier, tnote = "decided", "chosen with --choose"
        elif both(name) == both(r["Campaign_in_live"]):
            tier, tnote = "auto", "the engine treats the vaccines the same under both names"
        else:
            unp = unpaired(c, str(r["Campaign_in_live"]), str(r["Round_in_live"]))
            if unp is None:
                tier, tnote = "review", "the vaccines of the records are not known"
            elif unp == 0:
                tier, tnote = "auto", "no single-vaccine record is left without a partner, so the name changes nothing for the vaccines"
            else:
                tier, tnote = "review", "%d single-vaccine record(s) have no record of the other vaccine in the same district and date; the name %s %s, so the engine would treat them as %s" % (
                    unp, name, "names both vaccines" if both(name) else "no longer names both vaccines", "co-administered (it copies the vaccinated figure to the other vaccine)" if both(name) else "single-vaccine again")
        if str(r["Campaign_in_live"]) != name and key not in seen:
            corr.append({"Country": r["Country"], "Province": "", "District": "", "Response": r["Campaign_in_live"], "Round_Add": r["Round_in_live"], "SIA_date": "",
                         "Column": "Response", "Found": r["Campaign_in_live"], "Restored": name,
                         "Reason": (tie_note or ("Scope: " + str(r["Scope_says"])))[:300], "Tier": tier, "Tier_note": tnote})
        seen[key] = name
        for dn in drop:
            nm, hh = lk_name(dn, c)
            for _, x in hh.iterrows():
                chg.append({"Action": "remove", "Country": x["Country"], "OBR Name": x["OBR Name"], "Round Number": x["Round Number"], "Round Start Date": str(x["Round Start Date"])[:10],
                            "Vaccines": x.get("Vaccines", ""), "Round Status": x.get("Round Status", ""), "New_Round_Start_Date": "",
                            "Reason": "Same districts and dates as %s in the scope and the lookup: the campaign was entered twice" % name, "Tier": tier, "Tier_note": tnote})
        rn = pd.to_numeric(pd.Series([str(r["Round_in_live"])]).str.extract(r"(\d+)")[0], errors="coerce").iloc[0]
        near = ((hit["_dt"] - d0).abs().dt.days <= near_days).any()
        if not near and not pd.isna(rn) and rn >= 1 and not (hit["_rn"] == rn).any():
            chg.append({"Action": "add", "Country": hit["Country"].iloc[0], "OBR Name": name, "Round Number": "Rnd%d" % int(rn), "Round Start Date": str(d0)[:10],
                        "Vaccines": hit["Vaccines"].iloc[0] if "Vaccines" in hit.columns else "", "Round Status": "Finished", "New_Round_Start_Date": "",
                        "Reason": "Live records (%d districts, all planned in %s in the scope) of a round the lookup does not list" % (nd, name), "Tier": tier, "Tier_note": tnote})
    cc = pd.DataFrame(corr)
    ch = pd.DataFrame(chg).drop_duplicates() if chg else pd.DataFrame(chg)
    return cc, ch, skipped


def pending_close(out_root, targets, note="", by=""):
    """Mark pending rounds as closed (the round did not take place, or no data are expected). A target is a piece of 'COUNTRY|campaign|Round n' (not case sensitive);
    it must match exactly one round of the register. Returns [(target, message)]."""
    out_root = Path(out_root)
    reg_path = out_root / "pending_register.json"
    reg = json.loads(reg_path.read_text(encoding="utf-8")) if reg_path.is_file() else {"rounds": {}}
    res = []
    for t in targets:
        hits = [k for k in reg["rounds"] if t.lower() in k.lower()]
        if len(hits) != 1:
            res.append((t, "NOT DONE: %s" % ("no pending round matches (run 'submissions pending' first)" if not hits else "%d rounds match, be more precise: %s" % (len(hits), "; ".join(hits[:4])))))
            continue
        reg["rounds"][hits[0]]["closed"] = {"on": dt.datetime.now().strftime("%Y-%m-%d"), "by": by, "note": note}
        res.append((t, "closed: %s" % hits[0]))
    reg_path.write_text(json.dumps(reg, indent=1, ensure_ascii=False), encoding="utf-8")
    return res


FR_MATTER = {"target population missing": "population cible manquante", "failed a check": "contrôle non satisfait", "empty fields": "champs vides",
             "text instead of numbers": "texte à la place de nombres", "other country": "autre pays", "unit not clear": "unité pas claire",
             "same place and date, other figures": "même lieu et même date, autres chiffres", "campaign name differs": "nom de campagne différent",
             "round differs": "round différent", "vaccine differs from the lookup": "vaccin différent du lookup", "wrong folder": "mauvais dossier",
             "sheet not read": "feuille non lue", "error values in calculated columns": "valeurs en erreur dans des colonnes calculées",
             "data not received": "données non reçues"}


def _top(series, n=3):
    vc = series.dropna().astype(str).value_counts().head(n)
    return ", ".join("%s %d" % (k, v) for k, v in vc.items())


def feedback(out_root, ledger=None):
    """What each country has to correct, confirm or send again, written by the engine from the held-back records and the findings of the inspection.
    Returns (DataFrame with one line per country, file and matter, markdown text). Plain words: the people who read it are the countries."""
    ledger = ledger or load_ledger(out_root)
    rows = []

    def add(country, file, matter, records, text, action, fr=None):
        rows.append({"Country": country, "File": file, "Matter": matter, "Records": records, "What we found": text, "What the country does": action,
                     "Constat": fr[0] if fr else "", "Ce que le pays fait": fr[1] if fr else "", "Sujet": FR_MATTER.get(matter, matter) if fr else ""})

    for h, e in ledger["files"].items():
        if e.get("superseded") or not e.get("present", True) or e.get("status") == "nothing new" and not e.get("aliases"):
            continue
        ctry = e.get("country_folder") or e.get("countries") or "?"
        name = e["name"]
        codes = set(e.get("codes") or [])
        if "not_admin" in codes:
            continue
        # held-back records of the file, by reason
        hp = Path(out_root) / str(e.get("prepared") or "x").replace(".xlsx", "_held.xlsx")
        if e.get("prepared") and hp.is_file() and int(e.get("held_new") or 0) > 0:
            try:
                d = pd.read_excel(hp, dtype=object)
            except Exception:
                d = pd.DataFrame()
            if len(d):
                d["_r"] = d["Hold_reason"].astype(str)
                kinds = [("validation error", "validation"), ("missing after standardising", "empty"), ("target or the number vaccinated is not a number", "text"),
                         ("record of another country", "country"), ("same names as several live records", "several"), ("possible duplicate", "dup")]
                for key, kind in kinds:
                    g = d[d["_r"].str.contains(key, regex=False)]
                    if not len(g):
                        continue
                    camp = ", ".join(sorted({"%s %s" % (a, b) for a, b in zip(g["Response"].astype(str), g["Round_Add"].astype(str))})[:3])
                    where = _top(g["Province"])
                    n = len(g)
                    if kind == "validation":
                        lab = sorted({x.split("validation error: ")[1].split("; ")[0] for x in g["_r"] if "validation error: " in x})
                        if any("Target population missing" in x for x in lab):
                            add(ctry, name, "target population missing", n, "%d record(s) of %s have no target population (%s)" % (n, camp, where),
                                "Fill in the target population (PopPolio) of these districts and send the file again",
                                fr=("%d enregistrement(s) de %s n'ont pas de population cible (%s)" % (n, camp, where), "Renseigner la population cible (PopPolio) de ces districts et renvoyer le fichier"))
                        else:
                            add(ctry, name, "failed a check", n, "%d record(s) of %s fail the check: %s (%s)" % (n, camp, "; ".join(lab) or "see the Issues sheet", where),
                                "Correct these records and send the file again",
                                fr=("%d enregistrement(s) de %s ne passent pas le contrôle : %s (%s)" % (n, camp, "; ".join(lab) or "voir la feuille Issues", where), "Corriger ces enregistrements et renvoyer le fichier"))
                    elif kind == "empty":
                        cols = sorted({c.strip() for x in g["_r"] for c in x.split("missing after standardising: ")[1].split(";")[0].split(",")})
                        add(ctry, name, "empty fields", n, "%d record(s) of %s still have no %s after the lookup was used (%s)" % (n, camp, ", ".join(cols), where),
                            "Fill in %s for these records" % ", ".join(cols),
                            fr=("%d enregistrement(s) de %s n'ont toujours pas de %s après usage du lookup (%s)" % (n, camp, ", ".join(cols), where), "Renseigner %s pour ces enregistrements" % ", ".join(cols)))
                    elif kind == "text":
                        add(ctry, name, "text instead of numbers", n, "%d record(s) of %s have text where the target or the number vaccinated should be (%s)" % (n, camp, where),
                            "Replace the text by the number",
                            fr=("%d enregistrement(s) de %s contiennent du texte à la place de la population cible ou du nombre de vaccinés (%s)" % (n, camp, where), "Remplacer le texte par le nombre"))
                    elif kind == "country":
                        add(ctry, name, "other country", n, "%d record(s) carry another country than the folder (%s)" % (n, _top(g["Country"])),
                            "Correct the country or send these records in their own country's folder",
                            fr=("%d enregistrement(s) portent un autre pays que le dossier (%s)" % (n, _top(g["Country"])), "Corriger le pays ou envoyer ces enregistrements dans le dossier de leur pays"))
                    elif kind == "several":
                        ex = ", ".join(sorted({str(x) for x in g["District"].dropna()})[:4])
                        add(ctry, name, "unit not clear", n, "%d record(s) of %s have a province and district that match several records of the live file (for example %s)" % (n, camp, ex),
                            "Say which unit (health zone, sub-district or other level) each of these rows is",
                            fr=("%d enregistrement(s) de %s ont une province et un district qui correspondent à plusieurs enregistrements du fichier de référence (par exemple %s)" % (n, camp, ex), "Indiquer quelle unité (zone de santé, sous-district ou autre niveau) correspond à chacune de ces lignes"))
                    elif kind == "dup":
                        live = sorted(set(re.findall(r"\(([^,()]+ Rnd\d+), \d{4}-\d\d-\d\d\)", " ".join(g["_r"]))))[:3]
                        add(ctry, name, "same place and date, other figures", n,
                            "%d record(s) of %s look like records already in the live file (%s) with other figures" % (n, camp, ", ".join(live) or "another campaign name or round"),
                            "Confirm the campaign and round, and which figures are the final ones",
                            fr=("%d enregistrement(s) de %s ressemblent à des enregistrements déjà présents (%s) avec d'autres chiffres" % (n, camp, ", ".join(live) or "autre nom de campagne ou round"), "Confirmer la campagne et le round, et quels chiffres sont les chiffres définitifs"))
        # records matched to the live file under another campaign name or round
        for a_ in e.get("aliases") or []:
            same_round = str(a_["file_round"]).strip().lower() == str(a_["live_round"]).strip().lower()
            if same_round:
                add(ctry, name, "campaign name differs", a_["records"], "%d record(s) are written as %s %s; the live file has the same districts and dates as %s %s" % (
                    a_["records"], a_["file_name"], a_["file_round"], a_["live_name"], a_["live_round"]),
                    "Use the live campaign name in the next file (nothing else is needed: the records were matched automatically)",
                    fr=("%d enregistrement(s) sont écrits %s %s ; le fichier de référence a les mêmes districts et dates sous %s %s" % (a_["records"], a_["file_name"], a_["file_round"], a_["live_name"], a_["live_round"]), "Utiliser le nom de campagne du fichier de référence dans le prochain fichier (rien d'autre n'est nécessaire : les enregistrements ont été rapprochés automatiquement)"))
            else:
                add(ctry, name, "round differs", a_["records"], "%d record(s) are written as %s %s; the live file has the same districts and dates as %s %s" % (
                    a_["records"], a_["file_name"], a_["file_round"], a_["live_name"], a_["live_round"]),
                    "Confirm the round (the file says %s, the live file says %s) and which figures are the final ones" % (a_["file_round"], a_["live_round"]),
                    fr=("%d enregistrement(s) sont écrits %s %s ; le fichier de référence a les mêmes districts et dates sous %s %s" % (a_["records"], a_["file_name"], a_["file_round"], a_["live_name"], a_["live_round"]), "Confirmer le round (le fichier indique %s, le fichier de référence indique %s) et quels chiffres sont les chiffres définitifs" % (a_["file_round"], a_["live_round"])))
        why = str(e.get("why") or "")
        m = re.search(r"vaccine against the lookup: ([^;]+?)(?:; |$)", why)
        if "vaccine_vs_lookup" in codes and m:
            add(ctry, name, "vaccine differs from the lookup", "", m.group(1).strip(), "Confirm the vaccine used in this campaign",
                fr=("le fichier écrit un vaccin différent de celui du lookup : %s" % m.group(1).strip(), "Confirmer le vaccin utilisé dans cette campagne"))
        if "wrong_place" in codes:
            add(ctry, name, "wrong folder", "", why.split(";")[0], "Send the file in the folder of the right country, or tell us it is not needed",
                fr=(why.split(";")[0], "Envoyer le fichier dans le dossier du bon pays, ou nous dire qu'il n'est pas nécessaire"))
        if "sheets_skipped" in codes:
            m = re.search(r"not read: ([^;]+)", why)
            add(ctry, name, "sheet not read", "", "sheet(s) with column names that could not be read: %s" % (m.group(1) if m else "?"),
                "Tell us if these sheets hold admin data; if so, put them in the AFRO template",
                fr=("feuille(s) avec des noms de colonnes qui n'ont pas pu être lues : %s" % (m.group(1) if m else "?"), "Dites-nous si ces feuilles contiennent des données administratives ; si oui, les mettre dans le modèle AFRO"))
        if "non_numbers" in codes:
            bc = e.get("bad_cols") or {}
            core_ = {"TotPop", "PopPolio", "TotalNbrVaccPolio", "Tot_Vacc_nOPV2", "Total_Vacc_bOPV"}
            hard = {c: n for c, n in bc.items() if c in core_}
            soft = {c: n for c, n in bc.items() if c not in core_}
            if hard and not any(r["File"] == name and r["Matter"] == "text instead of numbers" for r in rows):
                add(ctry, name, "text instead of numbers", sum(hard.values()), "text where a number is expected in %s" % ", ".join("%s (%d cells)" % kv for kv in hard.items()),
                    "Replace the text by the number",
                    fr=("texte à la place d'un nombre dans %s" % ", ".join("%s (%d cellules)" % kv for kv in hard.items()), "Remplacer le texte par le nombre"))
            if soft or not bc:
                add(ctry, name, "error values in calculated columns", "", ("%d cells that are not numbers in %s" % (sum(soft.values()), ", ".join(sorted(soft)[:6]))) if soft else
                    (re.search(r"\d+ cell\(s\) that are not numbers", why).group(0) if re.search(r"\d+ cell\(s\) that are not numbers", why) else "cells that are not numbers"),
                    "For information only: these cells are in calculated columns (coverage, wastage, proportions, month) or in columns of other vaccines, not in the polio target or "
                    "vaccinated figures, so nothing is held back and no new file is needed; the country can clean them at the source (an error such as #DIV/0! in its own formulas)",
                    fr=("%d cellules qui ne sont pas des nombres dans %s" % (sum(soft.values()), ", ".join(sorted(soft)[:6])) if soft else "cellules qui ne sont pas des nombres",
                        "À titre d'information : ces cellules sont dans des colonnes calculées ou de vaccins autres que la polio, pas dans la population cible ni le nombre de vaccinés de la polio ; rien n'est bloqué et aucun nouveau fichier n'est nécessaire"))
    # pending admin data: rounds of the lookup with nothing received 15 days after their start (written by 'inspect' / 'pending' in Pending_admin_data.xlsx)
    pp = Path(out_root) / "Pending_admin_data.xlsx"
    if pp.is_file():
        try:
            pn = pd.read_excel(pp, sheet_name="Pending_now", dtype=object)
        except Exception:
            pn = pd.DataFrame()
        days = PENDING_DAYS
        try:
            days = int(json.loads((Path(out_root) / "pending_register.json").read_text(encoding="utf-8")).get("days") or PENDING_DAYS)
        except Exception:
            pass
        for _, r in pn.iterrows():
            lab = "%s %s" % (r["Campaign"], r["Round"])
            add(str(r["Country"]).strip().upper(), "(no file received)", "data not received", "",
                "No admin data received for %s (planned start %s, %s days ago; the data are due %d days after the start)" % (lab, r["Planned_start"], r["Days_since_start"], days),
                "Send the admin data of this round (AFRO template), or tell us if the round did not take place",
                fr=("Aucune donnée administrative reçue pour %s (début prévu le %s, il y a %s jours ; les données sont attendues %d jours après le début)" % (lab, r["Planned_start"], r["Days_since_start"], days),
                    "Envoyer les données administratives de ce round (modèle AFRO), ou nous dire si le round n'a pas eu lieu"))
    f = pd.DataFrame(rows, columns=["Country", "File", "Matter", "Records", "What we found", "What the country does", "Constat", "Ce que le pays fait", "Sujet"]).drop_duplicates(
        ["Country", "File", "Matter", "What we found"]).sort_values(["Country", "File", "Matter"])
    lines = ["# What each country has to correct or confirm", "",
             "Written by the engine from the last inspection (`python admin.py submissions feedback`). Records that passed every check are already in the live file; this list is about the others.", ""]
    for c, d in f.groupby("Country", sort=True):
        lines.append("## %s" % c)
        lines.append("")
        for _, r in d.iterrows():
            lines.append("- **%s** (%s): %s. %s." % (r["Matter"], r["File"], r["What we found"], r["What the country does"]))
        lines.append("")
    if not len(f):
        lines.append("Nothing to send back.")
    return f, "\n".join(lines) + "\n"


FRENCH_COUNTRIES = {"ALGERIA", "BENIN", "BURKINA FASO", "BURUNDI", "CAMEROON", "CENTRAL AFRICAN REPUBLIC", "CHAD", "COMOROS", "CONGO", "COTE D'IVOIRE",
                    "CÔTE D'IVOIRE", "DEMOCRATIC REPUBLIC OF THE CONGO", "GABON", "GUINEA", "MADAGASCAR", "MALI", "MAURITANIA", "NIGER", "SENEGAL", "TOGO"}
FR_REASON = (("validation error: ", "erreur de validation : "), ("Target population missing or zero", "population cible manquante ou nulle"),
             ("Missing key field", "champ clé manquant"), ("missing after standardising: ", "manquant après standardisation : "),
             ("the target or the number vaccinated is not a number", "la population cible ou le nombre de vaccinés n'est pas un nombre"),
             ("record of another country than the folder", "enregistrement d'un autre pays que le dossier"),
             ("same names as several live records, none identical (not added): live rows", "mêmes noms que plusieurs enregistrements du fichier de référence, aucun identique (non ajouté) : lignes"),
             ("possible duplicate: same place, date and vaccine as a live record of another campaign name or round, other figures: live row",
              "doublon possible : même lieu, date et vaccin qu'un enregistrement d'une autre campagne ou d'un autre round, autres chiffres : ligne"))


def _fr_reason(t):
    t = str(t)
    for a, b in FR_REASON:
        t = t.replace(a, b)
    return t


def feedback_files(out_root, folder, ledger=None):
    """One workbook per country, ready to send: sheet 'What to do' (what we found and what the country does, by file) and sheet 'Records' (the held-back
    records themselves: province, district, campaign, round, date, the figures and the reason), so the focal point can correct the rows and send them back.
    Returns [(country, path)]."""
    ledger = ledger or load_ledger(out_root)
    f, _ = feedback(out_root, ledger)
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    # the workbooks of an earlier run go to _previous: only the current ones are in the folder, so nothing stale is sent by mistake
    old = list(folder.glob("Feedback_*.xlsx"))
    if old:
        (folder / "_previous").mkdir(exist_ok=True)
        for o_ in old:
            try:
                shutil.move(str(o_), str(folder / "_previous" / o_.name))
            except Exception:
                pass
    keep = ["File", "Province", "District", "Admin_3", "SIA_date", "Response", "Round_Add", "Vaccine_type", "PopPolio", "TotalNbrVaccPolio", "Hold_reason"]
    out = []
    today = dt.datetime.now().strftime("%Y-%m-%d")
    INFO_ONLY = {"campaign name differs", "error values in calculated columns"}
    for c, d in f.groupby("Country", sort=True):
        d = d.copy()
        d.insert(2, "Action needed", d["Matter"].map(lambda m: "No, for information" if m in INFO_ONLY else "Yes"))
        d = d.sort_values(["Action needed", "File"], ascending=[False, True])
        needs = bool((d["Action needed"] == "Yes").any())
        recs = []
        for h, e in ledger["files"].items():
            if (e.get("country_folder") or "") != c or e.get("superseded") or not e.get("present", True) or not e.get("prepared"):
                continue
            hp = Path(out_root) / str(e["prepared"]).replace(".xlsx", "_held.xlsx")
            if hp.is_file() and int(e.get("held_new") or 0) > 0:
                try:
                    x = pd.read_excel(hp, dtype=object)
                    x.insert(0, "File", e["name"])
                    recs.append(x)
                except Exception:
                    pass
        R = pd.concat(recs, ignore_index=True) if recs else pd.DataFrame(columns=keep)
        R = R[[k for k in keep if k in R.columns]].drop_duplicates()
        if "SIA_date" in R.columns:
            R["SIA_date"] = pd.to_datetime(R["SIA_date"], errors="coerce").dt.strftime("%Y-%m-%d")
        french = str(c).upper() in FRENCH_COUNTRIES
        has_files = any((e.get("country_folder") or "") == c and e.get("present", True) and not e.get("superseded") for e in ledger["files"].values())
        if french and "Hold_reason" in R.columns:
            R["Pourquoi non ajouté"] = R["Hold_reason"].map(_fr_reason)
        R = R.rename(columns={"Hold_reason": "Why it was not added"})
        D_en = d.drop(columns=["Country", "Constat", "Ce que le pays fait", "Sujet"])
        D_fr = d.assign(**{"Action requise": d["Action needed"].map(lambda x: "Non, pour information" if x.startswith("No") else "Oui")})[
            ["File", "Action requise", "Sujet", "Records", "Constat", "Ce que le pays fait"]].rename(columns={"File": "Fichier", "Records": "Enregistrements"})
        intro = pd.DataFrame({"Admin data, %s, %s" % (c, today): [
            ("Thank you for the administrative vaccination data you sent. The records that passed all checks are already in the AFRO admin data file."
             if has_files else "This workbook lists the campaign rounds of your country for which we have not received administrative vaccination data."),
            "This workbook lists what we could not add yet, what we have not received, and what we ask you to correct, confirm or send.",
            ("Matter 'data not received': the round started more than %d days ago and its admin data are not in the AFRO admin data file; please send them (AFRO template), or tell us if the round did not take place." % PENDING_DAYS)
            if (d["Matter"] == "data not received").any() else "",
            "Sheet 'What to do': one line per file and matter ('Action needed': Yes = we ask you to correct or confirm; No = for your information). Sheet 'Records': the rows concerned (when there are any), with the reason.",
            "Please send the corrected file, or answer the questions in an email; nothing in the live data is changed without your confirmation."] + ([
            "Version française : merci pour les données administratives de vaccination envoyées. Les enregistrements qui ont passé tous les contrôles sont déjà dans le fichier AFRO.",
            "Ce classeur liste ce que nous n'avons pas pu ajouter pour l'instant et ce que nous vous demandons de corriger ou de confirmer (feuille « À faire » ; feuille « Records » : les lignes concernées).",
            "Merci de renvoyer le fichier corrigé ou de répondre par courriel ; rien n'est modifié dans les données sans votre confirmation."] if french else [])})
        intro = intro[intro.iloc[:, 0].astype(str).str.len() > 0]
        path = writable(folder / ("Feedback_%s_%s%s.xlsx" % (re.sub(r"[^A-Za-z0-9]+", "_", c).strip("_"), today, "" if needs else "_information_only")))
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            intro.to_excel(xw, sheet_name="Read me", index=False)
            D_en.to_excel(xw, sheet_name="What to do", index=False)
            if french:
                D_fr.to_excel(xw, sheet_name="À faire", index=False)
            if len(R):
                R.to_excel(xw, sheet_name="Records", index=False)
            for ws in xw.book.worksheets:
                for col in ws.columns:
                    w = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col[:60])
                    ws.column_dimensions[col[0].column_letter].width = min(max(10, w + 2), 70)
                from openpyxl.styles import Alignment, Font
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical="top")
        out.append((c, str(path)))
    return out


def accepted_files(out_root, auto=False):
    """Prepared record files of the accepted files: [(name, path)]. With auto=True the files of automatic rule 1 are added
    (never a file a person rejected or accepted otherwise: they are already decided)."""
    ledger = load_ledger(out_root)
    out = []
    for h, e in ledger["files"].items():
        if e.get("decision") == "accepted" and not e.get("superseded") and e.get("prepared") and e.get("status") != "blocked":
            p = Path(out_root) / e["prepared"]
            if p.is_file():
                out.append((e["name"], str(p)))
    if auto:
        out += [(n, p) for n, p, _ in auto_candidates(out_root, ledger)]
    return out
