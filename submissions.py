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
import warnings
import datetime as dt
import difflib
import hashlib
import json
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
LEARNED = {}     # headers confirmed by the operators (inputs/header_synonyms.json: {"header as written": "template column"})


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
    elif len(seen) >= 0.7 * len(template_cols):
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
        vt = None
        if "Vaccines" in lk.columns and not is_empty(x.get("Vaccines")):
            vt = vmap.get(cn(x["Vaccines"]))
        for i in g.index:
            P.at[i, "Response"] = name
            notes.add(rowref(P, i), "Response", "", name, why)
            if is_empty(P.at[i, "Round_Add"]) and not pd.isna(x["_rn"]):
                P.at[i, "Round_Add"] = "Rnd%d" % (int(x["_rn"]) + off)
                notes.add(rowref(P, i), "Round_Add", "", P.at[i, "Round_Add"], why)
            if pd.isna(P.at[i, "SIA_date"]):
                P.at[i, "SIA_date"] = pd.Timestamp(x["_d"])
                notes.add(rowref(P, i), "SIA_date", "", x["_d"], why)
            if is_empty(P.at[i, "Vaccine_type"]) and vt:
                P.at[i, "Vaccine_type"] = vt
                notes.add(rowref(P, i), "Vaccine_type", "", vt, why)
        out.append("%s %s (%d record(s); from %s%s)" % (name, x["Round Number"], len(g), src, vnote))
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


KINDS = (("lqas", "an LQAS file"), ("monitor", "a monitoring (IM) file"), ("_im_", "a monitoring (IM) file"), ("im_polio", "a monitoring (IM) file"),
         ("etally", "an eTally file"), ("tally", "a tally sheet"), ("aefi", "an AEFI file"), ("supplies", "a supplies file"), ("dqc", "a list of data-quality issues"),
         ("synthese", "a monitoring summary"), ("resultats", "a results summary"), ("daily call", "a call-in report"), ("vaccine utilization", "a vaccine utilisation summary"))


def guess_kind(path, lay):
    n = Path(path).name.lower()
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
    try:
        lay = lay or detect_layout(path, template_cols, assume=(("Country",) if folder_country else ()) + ("Round_Add", "Response"))
    except Exception as e:
        s["why"] = "the file could not be opened (%s)" % type(e).__name__
        det["layout"] = {"class": "not readable", "map": {}, "synonym": [], "ignored": [], "missing_keys": [], "notes": [str(e)[:200]]}
        return s, None, det
    det["layout"] = lay
    s["layout"], s["header_row"] = lay["class"], (lay["header_row"] + 1) if lay["header_row"] is not None else ""
    s["sheet"] = ", ".join(str(x) for x in ([lay["sheet"]] + [L["sheet"] for L in lay.get("extra") or []]) if x is not None)
    if lay["class"] == "not usable":
        kind = guess_kind(path, lay)
        s["why"] = ("looks like %s, not admin data; " % kind if kind else "") + "; ".join(lay["notes"])
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
    # comparison with the live file (before the defaults are filled in: only what the country reported is compared)
    cmp_, diffs = compare_with_live(P, ctx, prof)
    det["compare"], det["diffs"] = cmp_, diffs
    if det.get("facility"):
        cmp_["Status"] = "already in the live file, same figures"
        cmp_["Note"] = "health-facility line; its district total equals the live file's " + det["facility"]["Response"] + " " + det["facility"]["Round_Add"]
    s["new"] = int((cmp_["Status"] == "new").sum())
    s["existing_same"] = int((cmp_["Status"] == "already in the live file, same figures").sum())
    s["existing_different"] = int((cmp_["Status"] == "already in the live file, other figures").sum())
    s["existing_units"] = int(cmp_["Status"].str.startswith("same names as several").sum())
    s["possible_dup"] = int(cmp_["Status"].str.startswith("possible duplicate").sum())
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
    if s.get("possible_dup"):
        reasons.append("%d record(s) that look like live records under another campaign name, round or date (not added): see the Records sheet" % s["possible_dup"])
    if lay["class"] != "template":
        reasons.append("layout: " + lay["class"])
    if det.get("from_live"):
        reasons.append("campaign details taken from the live file (same place and number vaccinated), not written in the file (please confirm): " + "; ".join(det["from_live"][:3]))
    if det.get("filled_down"):
        reasons.append("campaign filled down inside the sheet, not written in those lines (please confirm): " + "; ".join(det["filled_down"][:3]))
    if det.get("inferred"):
        reasons.append("campaign details taken from the lookup / the file name, not written in the file (please confirm): " + "; ".join(det["inferred"][:3]))
    if lay.get("etally"):
        reasons.append("; ".join(lay["notes"][:1]))
    if det.get("callin_summary"):
        reasons.append(det["callin_summary"])
    # a file that is mostly another country's records than the folder it is in is in the wrong place
    if folder_country and len(P) >= 10:
        oc = P["Country"].map(lambda v: v != folder_country)
        if oc.mean() > 0.5:
            top_c = P.loc[oc, "Country"].value_counts()
            hint = ("; the records are of %s: move the file to that folder" % top_c.index[0]) if len(top_c) and top_c.iat[0] >= 0.8 * int(oc.sum()) and len(P) < 50 else "; not usable, delete it from the folder"
            reasons.insert(0, "WRONG FILE OR WRONG FOLDER: %d of %d records are of other countries than the folder (%s)%s" % (int(oc.sum()), len(P), folder_country, hint))
            s["wrong_place"] = True
    if lay.get("skipped_sheets"):
        reasons.append("sheet(s) with column names but not usable, not read: %s" % ", ".join(str(x) for x in lay["skipped_sheets"]))
    if folder_country and len(P):
        other = P["Country"].map(lambda v: v != folder_country)
        if other.any():
            others = P.loc[other, "Country"].value_counts().head(3)
            reasons.append("%d record(s) of another country than the folder (%s): %s" % (
                int(other.sum()), folder_country, ", ".join("%s %d" % (k, v) for k, v in others.items())))
    if len(det["unseen"]):
        reasons.append("values never seen in the live file")
    if det["bad_numbers"] is not None and len(det["bad_numbers"]):
        reasons.append("%d cell(s) that are not numbers" % len(det["bad_numbers"]))
    if s["qc_warnings"] > 0.25 * s["rows_used"]:
        reasons.append("more than 25% of the records have a warning")
    if s["new"] == 0:
        reasons.append("no new record")
    s["status"] = "ready" if not reasons else "review"
    if s.get("wrong_place"):
        s["status"] = "blocked"
    if not s.get("wrong_place") and (reasons == ["no new record"] or (s["new"] == 0 and not s["qc_errors"] and not s["existing_different"] and not s["existing_units"] and not s.get("possible_dup"))):
        s["status"] = "nothing new"
    s["why"] = "; ".join(reasons)
    if det.get("facility"):
        s["status"], s["why"] = "nothing new", det["facility"]["message"] + ("; " + s["why"] if s["why"] else "")
    if not hasattr(ctx, "_file_records"):
        ctx._file_records = {}
    ctx._file_records[s["hash"]] = (name, P[[c for c in ("Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response", "TotalNbrVaccPolio", "PopPolio") if c in P.columns]].copy())
    P = P[(cmp_["Status"] == "new").values]                 # only the records the live file does not have go to the build
    P = P.drop(columns=["_src_row", "_sheet"])
    P = P[template_cols]
    return s, P, det


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


def compare_with_live(P, ctx, prof):
    b = ctx.builder
    T = to_text(P.drop(columns=["_src_row", "_sheet"]), prof["numeric"])
    key, blank = b.dup_key(T)
    rows, diffs = [], []
    cols = [c for c in FIGURE_COLS + ["Date_Report"] if c in P.columns]
    for i in range(len(P)):
        k = key.iat[i]
        cand = ctx.live_index.get(k)
        if not cand:
            st, note = "new", ""
            # same place, same vaccine, and the same figures (or the same date) in the live file under another campaign name, round or date
            pl = place_hits(ctx, b._qcountry(P.at[i, "Country"]), cn(P.at[i, "Province"]), cn(P.at[i, "District"]))
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
        rows.append({"Row_in_file": rowref(P, i), "Country": P.at[i, "Country"], "Province": P.at[i, "Province"],
                     "District": P.at[i, "District"], "SIA_date": P.at[i, "SIA_date"].strftime("%Y-%m-%d"),
                     "Round_Add": P.at[i, "Round_Add"], "Vaccine_type": P.at[i, "Vaccine_type"], "Response": P.at[i, "Response"],
                     "Status": st, "Note": note})
    return pd.DataFrame(rows), pd.DataFrame(diffs)


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


def inspect_folder(ctx, folder, out_root, template_cols, log=print):
    """Inspect every file of a folder. Writes out_root/Submissions_report.xlsx, out_root/prepared/<hash>.xlsx and updates the ledger."""
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
    for n_file, (p, rel) in enumerate(files, 1):
        h = sha256_of(p)
        seen_hashes.add(h)
        old = ledger["files"].get(h)
        fc = folder_country_of(rel, ctx.geo, ctx.builder)
        log("  [%d/%d] inspecting %s%s" % (n_file, len(files), rel, ("  (folder country: %s)" % fc) if fc else ""))
        s, P, det = prepare_file(p, ctx, prof, template_cols, folder_country=fc)
        s["path"] = rel
        prepared = None
        if P is not None:
            prepared = "prepared/%s.xlsx" % h[:12]
            P.to_excel(out_root / prepared, sheet_name="Sheet1", index=False)
        e = old or {"decision": None, "decided_by": "", "decided_at": "", "note": "", "first_seen": now}
        e.update({k: s[k] for k in ("name", "path", "size", "hash", "layout", "rows_read", "rows_used", "excluded", "new", "existing_same",
                                    "existing_different", "existing_units", "qc_errors", "qc_warnings", "countries", "status", "why")})
        e["possible_dup"] = s.get("possible_dup", 0)
        e.update({"prepared": prepared, "last_inspected": now, "present": True})
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
    # files that were seen before and are not in the folder any more; a newer version of the same name supersedes the older one
    for h, e in ledger["files"].items():
        e["present"] = h in seen_hashes
        newer = [x for k, x in ledger["files"].items() if k != h and x.get("path", x.get("name")) == e.get("path", e.get("name")) and x.get("last_inspected", "") > e.get("last_inspected", "")
                 and k in seen_hashes]
        e["superseded"] = bool(newer) and h not in seen_hashes
    save_ledger(out_root, ledger)
    shutil.rmtree(scratch, ignore_errors=True)
    write_report(out_root / "Submissions_report.xlsx", rows, details, ctx, not_read)
    log("report: %s" % (out_root / "Submissions_report.xlsx"))
    return rows


def write_report(path, rows, details, ctx, not_read=()):
    ov = pd.DataFrame(rows)
    cols = ["name", "path", "status", "decision", "layout", "countries", "rows_read", "rows_used", "excluded", "new", "existing_same",
            "existing_different", "existing_units", "possible_dup", "in_file_duplicates_removed", "qc_errors", "qc_warnings", "why", "hash"]
    ov = ov[[c for c in cols if c in ov.columns]].rename(columns={"name": "File", "path": "Path", "status": "Status", "decision": "Decision",
        "layout": "Layout", "countries": "Countries", "rows_read": "Rows read", "rows_used": "Rows used", "excluded": "Rows left out",
        "new": "New records", "existing_same": "Already there (same)", "existing_different": "Already there (other figures)", "existing_units": "Same names as several live records (not added)", "possible_dup": "Possible duplicates (not added)",
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
        lay.append({"File": s["name"], "Layout": L.get("class", ""), "Sheet": s.get("sheet", L.get("sheet", "")), "Header row": s["header_row"],
                    "Template columns found": L.get("template_columns_found", ""), "Missing key columns": ", ".join(L.get("missing_keys", [])),
                    "Mapped by synonym (verify)": "; ".join(L.get("synonym", [])), "Columns not in the template (ignored)": "; ".join(L.get("ignored", [])[:40]),
                    "Notes": "; ".join(L.get("notes", []))})
    sheets["Layout"] = pd.DataFrame(lay)
    # every header the app could not map, over all the files: the way it learns. Copy a line into header_synonyms.json to teach it.
    import difflib
    tcols = {cn(c): c for c in (ctx.live.columns if getattr(ctx, "live", None) is not None else [])}
    cnt = {}
    for h, (s, det) in details.items():
        L = det["layout"] or {}
        if L.get("class") in ("not readable", "not usable") and not L.get("map"):
            continue                                   # a file that is not admin data teaches nothing
        for x in L.get("ignored", []):
            if "(second column for" in str(x) or cn(x) in tcols or cn(x) in SYNONYMS or cn(x) == "checkdp":
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
    for name, key in (("Campaign_to_set", "to_set"), ("eTally_vs_callin", "callin_check"), ("Tally_check", "tally_check"), ("Records", "compare"), ("Differences_with_live", "diffs"), ("Issues", "issues"), ("Normalised_summary", "note_counts"),
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
