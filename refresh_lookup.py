#!/usr/bin/env python3
"""Refresh lookup.xlsx from the GPEI (poliogo.afro.who.int) API: the same data, rules and checks as the scope pipeline
(scope/scripts/prepare_the_AFRO_SIA_Dashboard_input.py, function fetch_gpei_data_via_api), restricted to the lookup.

    python refresh_lookup.py                      # refresh lookup.xlsx next to this script
    python refresh_lookup.py --out PATH --dry-run # fetch and compare, write nothing

The lookup has one row per Polio campaign round that is Finished or In Progress (Round Number, OBR Name, Vaccines, Country,
Round Start Date, Round Status). It needs GPEI_API_TOKEN (https://poliogo.afro.who.int/dashboard/api/apitoken), taken from the
environment or from a secrets.env file (--env-file, or ADMIN_GPEI_ENV_FILE). Nothing is replaced unless the new file passes the
checks: expected columns, statuses, at least 20 rows, not less than 50% of the previous rows. The previous file is kept in
lookup_backups/ (newest 5). If nothing changed, the file is left alone.

Used by admin.py (lookup, build, inspect). Exit code: 0 refreshed or unchanged, 1 failed (the old lookup stays in place)."""
import argparse
import hashlib
import os
import shutil
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
API = os.environ.get("GPEI_API_BASE_URL", "https://poliogo.afro.who.int")
REQUIRED = ["Country", "OBR Name", "Vaccines", "Round Number", "Round Start Date"]
COLUMNS = ["Round Number", "OBR Name", "Vaccines", "Country", "Round Start Date", "Round Status"]
STATUS_COL = "Round Status"
ALLOWED = {"Finished", "In Progress"}
MIN_FRACTION = 0.5
MIN_ROWS = 20
KEEP_BACKUPS = 5


class Failed(Exception):
    pass


def say(msg, emit=None):
    (emit or (lambda m: print(m, flush=True)))(msg)


# ---------------------------------------------------------------------------------------------- token
def read_token(env_file=None):
    """GPEI_API_TOKEN from the environment, else from the first secrets.env that has it."""
    t = os.environ.get("GPEI_API_TOKEN", "").strip()
    if t:
        return t
    for f in [env_file, os.environ.get("ADMIN_GPEI_ENV_FILE")]:
        if not f or not Path(f).is_file():
            continue
        for line in Path(f).read_text(encoding="utf-8-sig", errors="replace").splitlines():
            line = line.strip()
            if line.startswith("GPEI_API_TOKEN") and "=" in line:
                v = line.split("=", 1)[1].strip().strip('"').strip("'")
                if v:
                    return v
    return ""


# ---------------------------------------------------------------------------------------------- rules shared with the scope pipeline
def is_polio_campaign(entry):
    return any(ct.get("slug") == "polio" for ct in (entry.get("campaign_types") or []))


def round_status(rnd, today):
    """Finished / In Progress from the round dates; planned or not yet started rounds are not in the lookup."""
    if rnd.get("is_planned"):
        return None
    started, ended = rnd.get("started_at"), rnd.get("ended_at")
    if not started or date.fromisoformat(started) > today:
        return None
    if ended:
        return "Finished" if date.fromisoformat(ended) < today else "In Progress"
    return "In Progress"


def clean_vaccine_names(raw):
    """Drop a vaccine token that is contained in another token of the same value ('nOPV2, nOPV2 & bOPV' -> 'nOPV2 & bOPV')."""
    if not isinstance(raw, str) or not raw:
        return raw if isinstance(raw, str) else ""
    tokens = [t.strip() for t in raw.split(",") if t.strip()]
    if not tokens:
        return raw
    kept = [t for i, t in enumerate(tokens)
            if not any(i != j and t != o and t in o for j, o in enumerate(tokens))]
    seen, out = set(), []
    for t in kept:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return ", ".join(out) if out else raw


def fetch_campaigns(headers, fieldset, page_limit=100, max_pages=500, emit=None):
    import requests
    by_id, page = {}, 1
    while page <= max_pages:
        r = requests.get(API + "/api/polio/campaigns/", headers=headers, timeout=60,
                         params={"fieldset": fieldset, "limit": page_limit, "page": page, "campaign_category": "all",
                                 "show_test": "true", "is_embedded": "false", "enabled": "true"})
        if r.status_code == 404 and page > 1:       # the previous page was the last one (exact multiple of page_limit)
            page -= 1
            break
        r.raise_for_status()
        rows = r.json().get("campaigns", [])
        for c in rows:
            by_id[c["id"]] = c
        if len(rows) < page_limit:                 # end of data is the short page, not the has_next flag
            break
        page += 1
    else:
        say("WARNING: lookup fetch hit the %d-page safety cap; the data may be truncated." % max_pages, emit)
    say("  GPEI API, fieldset=%s: %d campaigns." % (fieldset, len(by_id)), emit)
    return by_id


def fetch_lookup(token, emit=None):
    """The lookup as a DataFrame, straight from the API."""
    import pandas as pd
    headers = {"Authorization": "Bearer " + token, "Accept": "application/json"}
    try:
        lst = fetch_campaigns(headers, "list", emit=emit)
        det = fetch_campaigns(headers, "detail", emit=emit)
    except Exception as e:
        raise Failed("could not read the GPEI API (%s: %s)" % (type(e).__name__, str(e)[:160]))
    today = date.today()
    rows = []
    for cid, entry in lst.items():
        if not is_polio_campaign(entry):
            continue
        d = det.get(cid)
        if not d:
            continue
        for rnd in d.get("rounds", []):
            st = round_status(rnd, today)
            if st and rnd.get("started_at"):
                rows.append({"Round Number": "Round %s" % rnd.get("number"), "OBR Name": entry.get("obr_name"),
                             "Vaccines": clean_vaccine_names(rnd.get("vaccine_names") or ""),
                             "Country": entry.get("top_level_org_unit_name"), "Round Start Date": rnd["started_at"],
                             "Round Status": st})
    return pd.DataFrame(rows, columns=COLUMNS)


# ---------------------------------------------------------------------------------------------- checks and comparison
def validate(df, previous):
    miss = [c for c in REQUIRED if c not in df.columns]
    if miss:
        raise Failed("the new lookup lacks the columns %s" % miss)
    bad = set(df[STATUS_COL].dropna().unique()) - ALLOWED
    if bad:
        raise Failed("the new lookup has statuses %s (only Finished / In Progress expected)" % sorted(bad))
    if len(df) < MIN_ROWS:
        raise Failed("only %d rows in the new lookup (at least %d expected)" % (len(df), MIN_ROWS))
    if previous is not None and len(previous) > 0 and len(df) < len(previous) * MIN_FRACTION:
        raise Failed("the new lookup has %d rows against %d before (less than %d%%): probably a partial answer from the API"
                     % (len(df), len(previous), int(MIN_FRACTION * 100)))


def _key(df):
    d = df.reindex(columns=COLUMNS).fillna("").astype(str)       # an empty cell is NaN once read back from Excel
    d = d.apply(lambda c: c.str.strip())
    return d.sort_values(COLUMNS).reset_index(drop=True)


def rows_set(df):
    d = _key(df)
    return set(map(tuple, d.values.tolist()))


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def describe(path):
    """{'rows','sha256'} of an existing lookup file, or None."""
    p = Path(path)
    if not p.is_file():
        return None
    import pandas as pd
    try:
        return {"rows": int(len(pd.read_excel(p))), "sha256": sha(p)}
    except Exception:
        return {"rows": None, "sha256": sha(p)}


def refresh(path, token, backup_dir=None, dry_run=False, emit=None):
    """Refresh the lookup file. Returns a dict: status = updated | unchanged | failed, rows, added, removed, reason."""
    import pandas as pd
    path = Path(path)
    res = {"status": "failed", "rows": None, "added": 0, "removed": 0, "reason": "", "at": datetime.now().isoformat(timespec="seconds")}
    if not token:
        res["reason"] = "GPEI_API_TOKEN is not set"
        return res
    try:
        new = fetch_lookup(token, emit)
        old = None
        if path.is_file():
            try:
                old = pd.read_excel(path)
            except Exception:
                old = None
        validate(new, old)
        res["rows"] = int(len(new))
        a = rows_set(new)
        b = rows_set(old) if old is not None else set()
        res["added"], res["removed"] = len(a - b), len(b - a)
        if old is not None and a == b:
            res["status"] = "unchanged"
            return res
        res["status"] = "updated"
        if dry_run:
            return res
        if path.is_file() and backup_dir:
            bd = Path(backup_dir)
            bd.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, bd / ("lookup_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S")))
            for old_copy in sorted(bd.glob("lookup_*.xlsx"))[:-KEEP_BACKUPS]:
                old_copy.unlink()
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=str(path.parent))
        os.close(fd)
        try:
            new.to_excel(tmp, index=False)
            os.replace(tmp, path)
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)
        return res
    except Failed as e:
        res["status"], res["reason"] = "failed", str(e)
    except PermissionError as e:
        res["status"], res["reason"] = "failed", "could not write %s (is it open in Excel?)" % path.name
    except Exception as e:
        res["status"], res["reason"] = "failed", "%s: %s" % (type(e).__name__, str(e)[:160])
    return res


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=str(HERE / "lookup.xlsx"))
    ap.add_argument("--env-file")
    ap.add_argument("--backup-dir")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    r = refresh(a.out, read_token(a.env_file), a.backup_dir or str(Path(a.out).parent / "lookup_backups"), a.dry_run)
    if r["status"] == "failed":
        print("Lookup NOT refreshed: %s. The lookup in place is unchanged." % r["reason"])
        sys.exit(1)
    if r["status"] == "unchanged":
        print("Lookup unchanged (%d rounds)." % r["rows"])
    else:
        print("Lookup %s: %d rounds (%d new or changed, %d gone)." % ("would be updated" if a.dry_run else "updated",
                                                                     r["rows"], r["added"], r["removed"]))


if __name__ == "__main__":
    main()
