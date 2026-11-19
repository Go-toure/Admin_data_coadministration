#!/usr/bin/env python3
"""
admin.py - one command for the whole AFRO admin data cycle.

    python admin.py status            where everything stands (runs, backups, git, folders)
    python admin.py backup            save the live SharePoint file locally (and list the versions)
    python admin.py lookup            refresh lookup.xlsx from the GPEI API (also done at the start of build and inspect)
    python admin.py build             build the next version from the LIVE file (lookup refreshed, then backup first)
    python admin.py check             verify the newest build against the file it started from
    python admin.py excel             store the check_dp values in the build (written directly; --use-excel to use Excel)
    python admin.py publish           replace the file AND the data dictionary on SharePoint (all gates must be green)
    python admin.py dictionary        push only the data dictionary of the last published build (retry / refresh)
    python admin.py all               backup -> build -> check -> excel -> publish
    python admin.py runs              list the builds
    python admin.py inspect           quality + hygiene + lookup checks on a file (the live file by default)
    python admin.py find ...          look up records (country, province, district, response, date, round) in any version
    python admin.py compare A B       what differs between two versions (builds v24, backups b9, files)
    python admin.py rules             what the system knows and fixes, and the decisions behind it
    python admin.py commit -m "msg" FILE [FILE ...]    git add + commit + push (code only, never data)
    python admin.py sync-scope        copy the scripts / the published file to the scope folders
    python admin.py sync-ops          copy corrections.csv, lookup_overrides.csv and scope_.csv to the AdminOps folder (inputs/) for the Operations Center
    python admin.py doctor            check python packages, folders, Excel

Every build lives in its own folder, workspace/runs/vNN, with a manifest.json that records what it started from,
what was done and what is still to do. publish refuses to run unless the build was checked, saved in Excel and the
live file is still the one the build started from (so nobody's newer edit is ever overwritten).

Settings: admin_settings.json (paths only, no secrets). The SharePoint credentials stay in the backup tool's
config/secrets.env, which this script never reads.
"""
import argparse
import csv
import datetime as dt
import hashlib
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
SETTINGS_FILE = Path(os.environ["ADMIN_SETTINGS"]) if os.environ.get("ADMIN_SETTINGS") else HERE / "admin_settings.json"
DEFAULTS = {
    "backup_repo": "../SharePoint_version_backup",
    "workspace": "workspace",
    "version_floor": 23,                      # the next build is max(floor, existing runs) + 1
    "builder": "add_coadmin_columns.py",
    "publisher": "publish_admin_data.py",
    "lookup": "lookup.xlsx",
    "refresh_lookup": True,                   # refresh lookup.xlsx from the GPEI API at the start of build and inspect
    "gpei_env_file": "",                      # secrets.env holding GPEI_API_TOKEN (default: the scope folder's config/secrets.env)
    "geo": "geo_reference.csv",
    "afrogdb_dir": "",                        # AFRO-Geo-Reference folder (default: next to this repository)
    "district_source": "",                    # optional recovered file used by --district-source
    "submissions_dir": "workspace/submissions_in",   # the files of the countries: 'submissions pull' copies them here from SharePoint
    "submissions_auto": False,                # true: the build also takes the files of automatic rule 1 (see submissions.py), no decision needed
    "lookup_fix_auto": True,                  # before each build the engine wires what it can decide alone (strong scope evidence, vaccine treatment unchanged) into corrections.csv / lookup_overrides.csv
    "lookup_overrides_file": "workspace/lookup_overrides.csv",   # changes a person decided to the lookup (a campaign entered twice removed, a round added, a start date moved); applied every time the lookup is read
    "corrections_file": "workspace/corrections.csv",   # corrections a person decided; the build applies them (a cell changes only while it holds the value in Found)
    "pending_days": 15,                       # a lookup round with no admin data in the live file this many days after its start is marked pending (submissions pending)
    "pending_since": "2025-01-01",            # rounds that started before this date are old gaps (reference sheet), not pending
    "submissions_max_mb": 60,                 # a file of the submissions folder larger than this is not copied
    "build_extra_args": [],
    "max_row_drop_pct": 2.0,
    "key_columns": ["Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response"],
    "scope_script_dir": "",
    "scope_data_dir": "",
    "scope_file": "",
    "excel_timeout_min": 30,
    "excel_method": "python",                 # "python": values written into the file, no Excel needed; "excel": desktop Excel
    "excel_visible": True,                    # show the Excel window while it saves (any pop-up can then be answered)
}
EXPECTED_CHANGE = {"geo_code", "geo_match", "Response", "Round_Add", "District", "Province", "Country", "SIA_date", "Vaccine_type", "PopPolio",
                   "CVPolio", "TotalNbrVaccPolio", "Tot_Vacc_nOPV2", "CV_nOPV2", "Total_Vacc_bOPV", "CV_bOPV", "TotPop",
                   "AgeGPolio", "Admin_3", "Year", "Index", "SIAMTH"}
DATA_EXT = (".xlsx", ".xls", ".csv", ".pkl", ".rds", ".zip")
DATA_OK = ("lookup.xlsx", "geo_reference.csv")
FORBIDDEN_MSG = ("co-authored-by", "claude-session")


# ----------------------------------------------------------------------------------------------- settings / logging
def load_settings():
    s = dict(DEFAULTS)
    if SETTINGS_FILE.is_file():
        try:
            s.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
        except Exception as e:
            die("admin_settings.json is not valid JSON: %s" % e)
    return s


S = load_settings()


def P(rel):
    """A path from the settings: absolute stays, relative is taken from this folder."""
    p = Path(str(rel))
    return p if p.is_absolute() else (HERE / p).resolve()


WORK = P(S["workspace"])
RUNS = WORK / "runs"
STATE = WORK / "state.json"
DICT_BACKUPS = WORK / "dictionary_backups"      # the previous live dictionary is kept here before it is replaced
DICT_FAILED_RC = 3                              # publish script: admin data published, dictionary push failed
LOGDIR = Path(os.environ["ADMIN_LOGDIR"]) if os.environ.get("ADMIN_LOGDIR") else HERE / "logs"
BACKUP_REPO = P(S["backup_repo"])


def _last_part(recorded):
    return re.split(r"[\\/]", str(recorded))[-1]


def run_path(rd, recorded):
    """A file named in a manifest. The manifest may have been written on another computer (an absolute path that
    does not exist here), so the file is also looked for by name in the run folder."""
    p = Path(str(recorded))
    if p.is_file():
        return p
    return Path(rd) / _last_part(recorded)


def base_path_of(m):
    """The file a build started from: the recorded path, else the same file name in the local backups, else the
    backup version recorded in the manifest. None when it is not here."""
    b = m.get("base", {})
    p = Path(str(b.get("path", "")))
    if b.get("path") and p.is_file():
        return p
    snaps = BACKUP_REPO / "snapshots" / "admin"
    if b.get("path") and (snaps / _last_part(b["path"])).is_file():
        return snaps / _last_part(b["path"])
    for r in catalog_rows("admin"):
        if b.get("backup_version") and str(r["version"]) == str(b["backup_version"]) and (BACKUP_REPO / r["snapshot"]).is_file():
            return BACKUP_REPO / r["snapshot"]
    return None


def log(msg="", quiet=False):
    line = "%s  %s" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg) if msg else ""
    if not quiet:
        print(line, flush=True)
    try:
        LOGDIR.mkdir(exist_ok=True)
        with open(LOGDIR / ("admin_%s.log" % dt.date.today().strftime("%Y%m%d")), "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def die(msg, code=2):
    print("\nSTOPPED: " + msg + "\n", flush=True)
    try:
        log("STOPPED: " + msg, quiet=True)
    except Exception:
        pass
    sys.exit(code)


def run_cmd(cmd, cwd=None, tee=None, check=False):
    """Run a command, stream its output, optionally copy it to a file. Returns (returncode, text)."""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    p = subprocess.Popen([str(c) for c in cmd], cwd=str(cwd) if cwd else None, stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace", env=env)
    out = []
    f = open(tee, "w", encoding="utf-8") if tee else None
    try:
        for line in p.stdout:
            out.append(line)
            print(line, end="", flush=True)
            if f:
                f.write(line)
    finally:
        if f:
            f.close()
    p.wait()
    if check and p.returncode:
        die("a step failed (exit code %s): %s" % (p.returncode, " ".join(str(c) for c in cmd[:3])))
    return p.returncode, "".join(out)


# ----------------------------------------------------------------------------------------------- state and manifests
def read_json(p, default=None):
    try:
        return json.loads(Path(p).read_text(encoding="utf-8"))
    except Exception:
        return default


def write_json(p, obj):
    Path(p).parent.mkdir(parents=True, exist_ok=True)
    Path(p).write_text(json.dumps(obj, indent=2, ensure_ascii=False, default=str), encoding="utf-8")


def run_dirs():
    out = []
    if RUNS.is_dir():
        for d in RUNS.iterdir():
            m = re.match(r"^v(\d+)$", d.name)
            if m and d.is_dir():
                out.append((int(m.group(1)), d))
    return sorted(out)


def next_version():
    nums = [n for n, _ in run_dirs()] + [int(S["version_floor"])]
    return max(nums) + 1


def latest_run(unpublished=False):
    for n, d in reversed(run_dirs()):
        m = read_json(d / "manifest.json", {})
        if unpublished and m.get("published"):
            continue
        return n, d, m
    return None, None, None


def save_manifest(d, m):
    write_json(Path(d) / "manifest.json", m)


def now():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ----------------------------------------------------------------------------------------------- the SharePoint backup tool
def catalog_rows(key="admin"):
    p = BACKUP_REPO / "catalog" / "versions.csv"
    if not p.is_file():
        return []
    with open(p, newline="", encoding="utf-8") as f:
        rows = [r for r in csv.DictReader(f) if r.get("key") == key]
    return sorted(rows, key=lambda r: int(r["version"]))


def latest_snapshot():
    rows = catalog_rows("admin")
    if not rows:
        return None, None
    r = rows[-1]
    return r, BACKUP_REPO / r["snapshot"]


def do_backup(show=True):
    tool = BACKUP_REPO / "sp_versions.py"
    if not tool.is_file():
        die("the backup tool was not found: %s (see backup_repo in admin_settings.json)" % tool)
    log("Backing up the live SharePoint file ...")
    rc, _ = run_cmd([sys.executable, tool, "backup", "admin"], cwd=BACKUP_REPO)
    if rc:
        die("the SharePoint backup failed - nothing was changed. Check the message above.")
    if show:
        run_cmd([sys.executable, tool, "list", "admin"], cwd=BACKUP_REPO)
    return latest_snapshot()


# ----------------------------------------------------------------------------------------------- data helpers
def _pd():
    import pandas as pd
    return pd


def load_norm(path, sheet=0):
    """The first sheet as text, numbers written the same way whatever Excel did to them, text without the control
    characters / extra spaces that Excel or a browser add, without check_dp."""
    pd = _pd()
    df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False, na_values=[""])
    if "check_dp" in df.columns:
        df = df.drop(columns=["check_dp"])
    for c in df.columns:
        col = df[c]
        num = pd.to_numeric(col, errors="coerce")
        ok = num.notna()
        if ok.any():
            df[c] = col.where(~ok, num.map(lambda x: format(x, ".10g") if x == x else x))
        col = df[c]
        dirty = col.str.contains(r"_x000D_|\s\s|^\s|\s$|[\r\n\t]", regex=True, na=False)
        if dirty.any():
            df.loc[dirty, c] = col[dirty].str.replace(r"_x000D_|\s+", " ", regex=True).str.strip()
    return df.fillna("")


def fingerprint(df):
    pd = _pd()
    h = hashlib.sha256()
    h.update("|".join(map(str, df.columns)).encode())
    h.update(pd.util.hash_pandas_object(df, index=False).values.tobytes())
    return h.hexdigest()


def import_submissions():
    path = HERE / "submissions.py"
    spec = importlib.util.spec_from_file_location("coadmin_submissions", path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def import_builder():
    """The builder's own helper functions (lookup reading, country keys) are reused, never copied."""
    path = P(S["builder"])
    spec = importlib.util.spec_from_file_location("coadmin_builder", path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


# ----------------------------------------------------------------------------------------------- status
def git(*args, cwd=HERE):
    try:
        r = subprocess.run(["git", *args], cwd=str(cwd), capture_output=True, text=True, encoding="utf-8", errors="replace")
        return r.returncode, (r.stdout + r.stderr).strip()
    except FileNotFoundError:
        return 127, "git is not installed"


def cmd_status(a):
    print("\n=== AFRO admin data: status ===\n")
    print("Folder    : %s" % HERE)
    print("Workspace : %s" % WORK)
    print("Backups   : %s%s" % (BACKUP_REPO, "" if BACKUP_REPO.is_dir() else "   <-- NOT FOUND"))
    if a.live:
        do_backup(show=False)
    rows = catalog_rows("admin")
    if rows:
        r = rows[-1]
        print("\nLatest local backup of the live file: version %s, %s rows, saved %s UTC, by %s"
              % (r["version"], r["rows"], r["saved_at_utc"].replace("T", " ").replace("Z", ""), r["modified_by"]))
        if not a.live:
            print("  (run 'status --live' or 'backup' to read SharePoint now)")
    else:
        print("\nNo local backup yet: run 'backup'.")
    print("\nBuilds:")
    runs = run_dirs()
    if not runs:
        print("  none yet: run 'build'")
    for n, d in runs[-5:]:
        m = read_json(d / "manifest.json", {})
        print("  v%-3d %s" % (n, describe_run(m)))
    st = read_json(STATE, {})
    if st.get("published"):
        print("\nLast published: v%s on %s (SharePoint backup version %s)"
              % (st["published"].get("version"), st["published"].get("at"), st["published"].get("backup_version")))
    for pn, pd in reversed(run_dirs()):
        pm = read_json(pd / "manifest.json", {})
        if pm.get("published"):
            ds = pm["published"].get("dictionary")
            if ds:
                print("Data dictionary: %s (v%d)%s" % (ds, pn, "  <- run 'python admin.py dictionary'" if ds == "failed" else ""))
            break
    n, d, m = latest_run(unpublished=True)
    if n:
        print("\nNext step for v%d: %s" % (n, next_step(m)))
    else:
        print("\nNext step: build")
    print("\nGit:")
    rc, out = git("status", "--short")
    if rc:
        print("  (not available here: %s)" % out.splitlines()[0][:80])
    else:
        changed = [l for l in out.splitlines() if l.strip()]
        print("  uncommitted: %s" % (("%d item(s)" % len(changed)) if changed else "nothing"))
        for l in changed[:8]:
            print("    " + l)
        rc, out = git("log", "@{u}..HEAD", "--oneline")
        if rc == 0:
            print("  not pushed : %s" % ("%d commit(s)" % len(out.splitlines()) if out.strip() else "nothing"))
    left = sorted(p.name for p in HERE.glob("_history*") if p.is_dir())
    if left:
        print("\nLeftover folders (delete when their commit is done): %s" % ", ".join(left))
        print("  rm -r " + " ".join(left))
    print()


def describe_run(m):
    bits = []
    if m.get("build", {}).get("ok"):
        bits.append("built")
    c = m.get("check")
    if c:
        bits.append("check %s" % ("OK" if c.get("ok") else "FAILED"))
    if m.get("excel", {}).get("ok"):
        bits.append("saved in Excel")
    if m.get("published"):
        bits.append("PUBLISHED %s" % m["published"].get("at", ""))
    return ", ".join(bits) or "started"


def next_step(m):
    if not m.get("build", {}).get("ok"):
        return "build (the last one did not finish)"
    if not m.get("check"):
        return "check"
    if not m["check"].get("ok"):
        return "fix the problems in the check report, then build again"
    if not m.get("excel", {}).get("ok"):
        return "excel (open and save once)"
    return "publish"


def cmd_runs(a):
    for n, d in run_dirs():
        m = read_json(d / "manifest.json", {})
        b = m.get("base", {})
        print("v%-3d %-19s base: backup v%s (%s rows) | rows out %s | errors %s | %s"
              % (n, m.get("created", ""), b.get("backup_version", "?"), b.get("rows", "?"),
                 m.get("build", {}).get("rows_out", "?"), m.get("build", {}).get("errors", "?"), describe_run(m)))


# ----------------------------------------------------------------------------------------------- backup
def cmd_backup(a):
    do_backup(show=True)


# ----------------------------------------------------------------------------------------------- lookup refresh
def lookup_env_file():
    f = S.get("gpei_env_file")
    if f:
        return P(f)
    sd = S.get("scope_script_dir")
    return (Path(sd).parent / "config" / "secrets.env") if sd else None


def lookup_stamp():
    """rows + sha256 of the lookup in place (recorded in the manifest of a build)."""
    try:
        import refresh_lookup as rl
        return rl.describe(P(S["lookup"])) or {}
    except Exception:
        return {}


def refresh_lookup_step(force=False, dry_run=False):
    """Refresh lookup.xlsx from the GPEI API with the same rules and checks as the scope pipeline. It never stops the cycle:
    if the API cannot be read, the lookup in place is used and the log says so loudly. Returns a dict for the manifest."""
    lk = P(S["lookup"])
    if not (force or S.get("refresh_lookup", True)):
        return {"status": "off"}
    try:
        import refresh_lookup as rl
    except Exception as e:
        log("WARNING: the lookup was NOT refreshed (refresh_lookup.py cannot be loaded: %s). The lookup in place is used." % e)
        return {"status": "failed", "reason": str(e)}
    log("Refreshing the lookup from the GPEI API ...")
    env = lookup_env_file()
    r = rl.refresh(lk, rl.read_token(str(env) if env else None), str(WORK / "lookup_backups"), dry_run=dry_run,
                   emit=lambda m: log(m))
    if r["status"] == "failed":
        have = rl.describe(lk)
        when = dt.datetime.fromtimestamp(lk.stat().st_mtime).strftime("%Y-%m-%d %H:%M") if lk.is_file() else "none"
        log("WARNING: the lookup was NOT refreshed (%s). The lookup in place is used (%s rounds, last changed %s)."
            % (r["reason"], have["rows"] if have else "?", when))
    elif r["status"] == "unchanged":
        log("Lookup is up to date (%d rounds)." % r["rows"])
    else:
        log("Lookup %s: %d rounds (%d new or changed, %d gone)." %
            ("would be refreshed" if dry_run else "refreshed from the GPEI API", r["rows"], r["added"], r["removed"]))
    r.update(lookup_stamp() if r["status"] != "failed" or lk.is_file() else {})
    return r


def cmd_lookup(a):
    r = refresh_lookup_step(force=True, dry_run=a.dry_run)
    if r["status"] == "failed":
        sys.exit(1)


# ----------------------------------------------------------------------------------------------- build
def base_info(path):
    """Rows, columns and fingerprint of a workbook, read by a separate process: the table is large, and a Python process
    does not give such memory back, which the builder started right after would have to share the machine with."""
    try:
        p = subprocess.run([sys.executable, str(Path(__file__).resolve()), "baseinfo", str(path)], capture_output=True,
                           text=True, encoding="utf-8", errors="replace", env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        for line in p.stdout.splitlines():
            if line.startswith("BASEINFO "):
                return json.loads(line[len("BASEINFO "):])
    except Exception:
        pass
    df = load_norm(Path(path))
    info = {"rows": len(df), "columns": list(df.columns), "fingerprint": fingerprint(df)}
    del df
    return info


def cmd_baseinfo(a):
    df = load_norm(Path(a.path))
    print("BASEINFO " + json.dumps({"rows": len(df), "columns": list(df.columns), "fingerprint": fingerprint(df)}), flush=True)


def cmd_build(a):
    if not getattr(a, "no_refresh_lookup", False):
        lk_info = refresh_lookup_step()
    else:
        lk_info = {"status": "skipped"}
    for need in (P(S["builder"]), P(S["lookup"])):
        if not need.is_file():
            die("file not found: %s" % need)
    if a.base:
        base = Path(a.base)
        if not base.is_file():
            die("the base file does not exist: %s" % base)
        row = None
    else:
        row, base = (do_backup(show=False) if not a.no_backup else latest_snapshot())
        if row is None or not base.is_file():
            die("no local backup of the live file: run 'backup' first.")
    if S.get("lookup_fix_auto") and not getattr(a, "no_auto_wire", False):
        try:
            log("Lookup wiring: what the evidence lets the engine decide alone (strong scope evidence, same vaccines) ...")
            cmd_lookup_fix(argparse.Namespace(action="auto", by="", choose=[], older=False, source=str(base)))
        except SystemExit:
            raise
        except Exception as e:
            log("Lookup wiring skipped (%s): the build goes on without it." % e)
    n = next_version()
    rd = RUNS / ("v%d" % n)
    if rd.exists():
        die("%s already exists." % rd)
    rd.mkdir(parents=True)
    out = rd / "out"
    log("Build v%d, starting from %s" % (n, base.name))
    log("Reading the base file (for the check later) ...")
    bi = base_info(base)
    m = {"version": n, "created": now(),
         "base": {"path": str(base), "backup_version": row["version"] if row else None,
                  "sha256": row["sha256"] if row else None, "rows": bi["rows"], "columns": bi["columns"],
                  "fingerprint": bi["fingerprint"], "modified_by": row["modified_by"] if row else None},
         "lookup": dict(lk_info, **lookup_stamp()),
         "build": {}}
    save_manifest(rd, m)
    cmd = [sys.executable, P(S["builder"]), base, "--out-dir", out, "--lookup", P(S["lookup"])]
    if lookup_overrides_path():
        cmd += ["--lookup-overrides", lookup_overrides_path()]
    geo = P(S["geo"])
    if geo.is_file():
        cmd += ["--geo", geo]
    if S.get("afrogdb_dir") and P(S["afrogdb_dir"]).is_dir():
        cmd += ["--afrogdb", P(S["afrogdb_dir"])]
    snaps = BACKUP_REPO / "snapshots" / "admin"
    if snaps.is_dir():
        cmd += ["--backups", snaps]
    ds = S.get("district_source")
    if ds and P(ds).is_file():
        cmd += ["--district-source", P(ds)]
    if not getattr(a, "no_submissions", False):
        sm = import_submissions()
        st = "local" if getattr(a, "no_ledger_pull", False) else submissions_state_pull()
        if st == "adminops":
            log("Submissions: the decisions of the Operations Center were read.")
        else:
            log("WARNING: the decisions of the Operations Center on the submissions were NOT read (%s). The ledger kept on this PC is used: "
                "files accepted in the app are not in this build." % st)
        m["build"]["submissions_ledger"] = st
        use_auto = bool(S.get("submissions_auto")) and not getattr(a, "no_auto_submissions", False)
        acc = sm.accepted_files(WORK / "submissions", auto=use_auto)
        if use_auto:
            auto_names = [n for n, _, _ in sm.auto_candidates(WORK / "submissions")]
            m["build"]["submissions_auto"] = auto_names
            if auto_names:
                log("Submissions: %d file(s) taken by the automatic rules 1 to 4 (nobody decided on them; a file with findings brings only its clean records, the rest is held back, see Held_back_records): %s" % (len(auto_names), ", ".join(auto_names)))
        if acc:
            log("Submissions: %d accepted file(s) will be added to this build: %s" % (len(acc), ", ".join(n for n, _ in acc)))
            for _, pth in acc:
                cmd += ["--add-records", pth]
            m["build"]["submissions"] = [n for n, _ in acc]
    cf = P(S.get("corrections_file")) if S.get("corrections_file") else None
    if cf and cf.is_file():
        cmd += ["--corrections", str(cf)]
        m["build"]["corrections"] = cf.name
        log("Decided corrections: %s will be applied (a cell changes only while it still holds the value it was found with)." % cf.name)
    cmd += list(S.get("build_extra_args") or []) + list(a.extra or [])
    m["build"]["command"] = " ".join(str(c) for c in cmd)
    rc, text = run_cmd(cmd, cwd=HERE, tee=rd / "build.log")
    if rc:
        m["build"]["ok"] = False
        save_manifest(rd, m)
        die("the build failed (see %s)." % (rd / "build.log"))
    f_up = out / "upload_to_sharepoint" / "AFRO_admin_data.xlsx"
    if not f_up.is_file():
        die("the build produced no AFRO_admin_data.xlsx (see %s)." % (rd / "build.log"))
    # versioned names in the run folder (what the publish step and the team look at)
    final = rd / ("AFRO_admin_data_v%d.xlsx" % n)
    shutil.copy2(f_up, final)
    for src, name in ((out / "upload_to_sharepoint" / "AFRO_admin_data_dictionary.xlsx", "AFRO_admin_data_dictionary_v%d.xlsx" % n),
                      (out / "review" / "AFRO_admin_data_coadmin_audit.xlsx", "AFRO_admin_data_coadmin_audit_v%d.xlsx" % n),
                      (out / "review" / "AFRO_admin_data_quality_report.xlsx", "AFRO_admin_data_quality_report_v%d.xlsx" % n),
                      (out / "review" / "Feedback_suspicious_coverage.xlsx", "Feedback_suspicious_coverage_v%d.xlsx" % n),
                      (out / "review" / "Fix_list_by_country.xlsx", "Fix_list_by_country_v%d.xlsx" % n)):
        if src.is_file():
            shutil.copy2(src, rd / name)
    mm = re.search(r"\((\d+) records with an Error", text) or re.search(r"(\d+) records with an Error", text)
    mr = re.search(r"rows out: (\d+)", text)
    ma = re.search(r"records added from submissions: (\d+)", text)
    m["build"].update({"ok": True, "finished": now(), "file": str(final), "errors": int(mm.group(1)) if mm else None,
                       "rows_out": int(mr.group(1)) if mr else None, "added_records": int(ma.group(1)) if ma else 0,
                       "summary": [l.strip() for l in text.splitlines()
                                   if re.search(r"filled|renamed|corrected|removed|merged|restored|repair", l)][:25]})
    save_manifest(rd, m)
    log("Build v%d done: %s" % (n, final))
    log("Next: python admin.py check")


# ----------------------------------------------------------------------------------------------- check
def response_country_mismatches(df, builder, lk):
    """Rows whose Response is a campaign name of the lookup that belongs to another country."""
    pd = _pd()
    owner = {}
    for rk_, ck_ in zip(lk["_rk"], lk["_ck"]):
        owner.setdefault(rk_, set()).add(ck_)
    pair = df[["Country", "Response"]].drop_duplicates()
    bad = {}
    for c, r in zip(pair["Country"], pair["Response"]):
        rk_ = builder._rk(r)
        if rk_ in owner and builder._lkc(c) not in owner[rk_]:
            bad[(c, r)] = int(((df["Country"] == c) & (df["Response"] == r)).sum())
    return bad


def cmd_check(a):
    n, rd, m = (None, None, None)
    if a.run:
        rd = RUNS / ("v%d" % a.run)
        n, m = a.run, read_json(rd / "manifest.json", {})
    else:
        n, rd, m = latest_run()
    if not n or not m.get("build", {}).get("ok"):
        die("no finished build to check: run 'build' first.")
    final = run_path(rd, m["build"]["file"])
    log("Checking v%d (%s) against the file it started from ..." % (n, final.name))
    was = (m.get("lookup") or {}).get("sha256")
    now_ = lookup_stamp().get("sha256")
    if was and now_ and was != now_:
        log("NOTE: the lookup changed after v%d was built. The checks below use the lookup in place now; run 'build' again "
            "if the new lookup should be used for the repairs." % n)
    pd = _pd()
    base_path = base_path_of(m)
    if base_path is None:
        die("the file v%d started from is not available here (backup version %s)." % (n, m["base"].get("backup_version")))
    base = load_norm(base_path)
    new = load_norm(final)
    builder = import_builder()
    lk = builder.load_lookup(str(P(S["lookup"])), lookup_overrides_path())
    hard, warn, info = [], [], []

    # 1. header
    miss = [c for c in base.columns if c not in new.columns]
    added = [c for c in new.columns if c not in base.columns]
    if miss:
        hard.append("columns missing from the new file: %s" % ", ".join(miss[:10]))
    if added:
        warn.append("new columns: %s" % ", ".join(added[:10]))
    # 2. rows
    drop_pct = (len(base) - len(new)) / max(len(base), 1) * 100
    added = int(m["build"].get("added_records") or 0)
    info.append("rows: %d -> %d (%+d)" % (len(base), len(new), len(new) - len(base)))
    if added:
        info.append("records added from accepted submissions (%s): %d" % (", ".join(m["build"].get("submissions") or []), added))
    if m["build"].get("submissions_auto"):
        info.append("files taken by the automatic rules (no person decided on them): %s" % ", ".join(m["build"]["submissions_auto"]))
    led = m["build"].get("submissions_ledger")
    if led and led != "adminops":
        warn.append("the decisions on the submissions were not read from the Operations Center (%s): files accepted in the app are not in this build" % led)
    if drop_pct > float(S["max_row_drop_pct"]):
        hard.append("%.1f%% of the rows disappeared (limit %.1f%%)" % (drop_pct, float(S["max_row_drop_pct"])))
    # 3. key fields blank
    keys = [k for k in S["key_columns"] if k in new.columns and k in base.columns]

    def blanks(df):
        return int((df[keys].apply(lambda c: c.astype(str).str.strip() == "")).any(axis=1).sum())
    b0, b1 = blanks(base), blanks(new)
    info.append("records with an empty key field: %d -> %d" % (b0, b1))
    if b1 > b0:
        hard.append("more records with an empty key field than before (%d -> %d)" % (b0, b1))
    # 4. Response that belongs to another country
    if "Response" in new.columns and "Country" in new.columns:
        m0 = response_country_mismatches(base, builder, lk)
        m1 = response_country_mismatches(new, builder, lk)
        info.append("records with a Response of another country (per the lookup): %d -> %d"
                    % (sum(m0.values()), sum(m1.values())))
        if sum(m1.values()) > sum(m0.values()):
            hard.append("more records with another country's Response (%d -> %d)" % (sum(m0.values()), sum(m1.values())))
        elif m1:
            warn.append("still %d record(s) with another country's Response: %s"
                        % (sum(m1.values()), "; ".join("%s / %s (%d)" % (k[0], k[1], v) for k, v in list(m1.items())[:6])))
    # 5. duplicates: a record entered twice (same names, same population and results) is a duplicate; records with the same names
    #    but other figures are different units (the file records no sub-district), counted apart and never a warning
    c0 = builder.classify_duplicates(base)["cls"].value_counts()
    c1 = builder.classify_duplicates(new)["cls"].value_counts()
    d0 = int(c0.get("exact", 0) + c0.get("same_record", 0))
    d1 = int(c1.get("exact", 0) + c1.get("same_record", 0))
    info.append("duplicate records (the same record entered twice): %d -> %d" % (d0, d1))
    info.append("records with the same names but other figures (different units, not duplicates): %d -> %d"
                % (int(c0.get("other_figures", 0)), int(c1.get("other_figures", 0))))
    if d1 > d0:
        warn.append("more duplicate records than before (%d -> %d)" % (d0, d1))
    # 6. errors reported by the build
    prev = None
    for k, d in reversed(run_dirs()):
        if k < n:
            e = read_json(d / "manifest.json", {}).get("build", {}).get("errors")
            if e is not None:
                prev = e
                break
    e1 = m["build"].get("errors")
    info.append("quality errors: %s -> %s" % (prev if prev is not None else "?", e1 if e1 is not None else "?"))
    if prev is not None and e1 is not None and e1 > prev:
        warn.append("more quality errors than the previous build (%d -> %d)" % (prev, e1))
    # 7. what changed, cell by cell (only when the rows line up)
    changes = {}
    if len(base) == len(new) - added or len(base) == len(new):
        new_cmp = new.iloc[:len(base)] if len(new) != len(base) else new
        common = [c for c in new.columns if c in base.columns]
        for c in common:
            k = int((base[c].values != new_cmp[c].values).sum())
            if k:
                changes[c] = k
        info.append("cells changed vs the base: %s" % (", ".join("%s %d" % kv for kv in sorted(changes.items(), key=lambda x: -x[1])[:12]) or "none"))
        big = [c for c in ("TotalNbrVaccPolio", "PopPolio") if changes.get(c, 0) > 0.01 * len(new)]
        if big:
            warn.append("more than 1%% of the records changed in: %s" % ", ".join(big))
        odd = [c for c, k in changes.items() if c not in EXPECTED_CHANGE and k > 0.005 * len(new)]
        if odd:
            warn.append("columns the builder does not normally touch changed in more than 0.5%% of the records: %s "
                        "(look at 'compare live vNN --col NAME' before publishing)" % ", ".join(odd[:8]))
    else:
        info.append("rows differ in number, so no cell-by-cell comparison")
    # 8. AFROGDB geo_code: the share of records with a code must not fall (names are not changed by this rule)
    if "geo_code" in new.columns:
        s1 = float(new["geo_code"].astype(str).str.strip().replace({"nan": "", "None": ""}).ne("").mean() * 100)
        s0 = float(base["geo_code"].astype(str).str.strip().replace({"nan": "", "None": ""}).ne("").mean() * 100) if "geo_code" in base.columns else None
        info.append("records with an AFROGDB geo_code: %s -> %.1f%%" % ("none" if s0 is None else "%.1f%%" % s0, s1))
        if s0 is not None and s1 < s0 - 1.0:
            warn.append("fewer records with an AFROGDB geo_code than before (%.1f%% -> %.1f%%): look at afrogdb_unmatched.csv" % (s0, s1))
        if s0 is None:
            info.append("first build with geo_code: the 'new columns' note above is expected")
    # 9. no record lost: every record of the base (same country, province, district, date and vaccine) is still in the build, apart from the duplicates removed
    try:
        rk_ = [k for k in ("Country", "Province", "District", "SIA_date", "Vaccine_type") if k in new.columns and k in base.columns]
        if rk_:
            kb = base[rk_].astype(str).agg("|".join, axis=1).value_counts()
            kn = new[rk_].astype(str).agg("|".join, axis=1).value_counts()
            lost = int((kb - kn.reindex(kb.index).fillna(0)).clip(lower=0).sum())
            allowed = max(d0 - d1, 0)
            extra_lost = max(lost - allowed, 0)
            info.append("records of the base no longer in the build (same country, province, district, date, vaccine), apart from %d removed duplicate(s): %d"
                        % (allowed, extra_lost))
            lost_pct = extra_lost / max(len(base), 1) * 100
            if lost_pct > 1.0:
                warn.append("%d record(s) of the base (%.1f%%) are no longer in the build, beyond the duplicates removed: look at 'compare live v%d' before publishing" % (extra_lost, lost_pct, n))
            elif extra_lost > max(10, 0.001 * len(base)):
                warn.append("%d record(s) of the base are no longer in the build, beyond the duplicates removed" % extra_lost)
    except Exception as e_:
        info.append("records of the base still in the build: could not be checked (%s)" % e_)
    if getattr(a, "strict", False):
        # the signs of damage become stops (the chronic findings, such as records that still have another country's Response, stay warnings)
        sign = ("more than 1% of the records changed", "columns the builder does not normally touch", "more duplicate records", "more quality errors",
                "fewer records with an AFROGDB geo_code", "the decisions on the submissions were not read", "of the base", "no longer in the build")
        promoted = [w for w in warn if any(t in w for t in sign)]
        for w in promoted:
            warn.remove(w)
            hard.append("(strict) " + w)
    ok = not hard
    rep = ["# Check of v%d" % n, "", "Started from backup version %s (%s)." % (m["base"].get("backup_version"), m["base"].get("modified_by")),
           "", "Result: **%s**" % ("OK" if ok else "FAILED"), ""]
    if hard:
        rep += ["## Stops", ""] + ["- " + x for x in hard] + [""]
    if warn:
        rep += ["## To look at", ""] + ["- " + x for x in warn] + [""]
    rep += ["## Figures", ""] + ["- " + x for x in info] + [""]
    if changes:
        rep += ["## Cells changed per column", ""] + ["- %s: %d" % kv for kv in sorted(changes.items(), key=lambda x: -x[1])] + [""]
    (rd / "check_report.md").write_text("\n".join(rep), encoding="utf-8")
    m["check"] = {"ok": ok, "at": now(), "stops": hard, "warnings": warn, "info": info}
    save_manifest(rd, m)
    print()
    for x in info:
        print("  " + x)
    for x in warn:
        print("  WARNING: " + x)
    for x in hard:
        print("  STOP   : " + x)
    print("\n  Check %s  (report: %s)\n" % ("OK" if ok else "FAILED", rd / "check_report.md"))
    if not ok:
        sys.exit(3)
    log("Next: python admin.py excel")


# ----------------------------------------------------------------------------------------------- excel
PS_SAVE = r'''
# Opens the workbook in Excel, lets it calculate, saves and quits. Every Excel call is retried: right after a big
# open, Excel often answers "busy" (0x800AC472) for a few seconds.
$path = $args[0]
$visible = ($args[1] -eq "1")
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $visible
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$wb = $null
$code = 0
function Step([string]$name, [scriptblock]$sb, [int]$tries = 12) {
  for ($i = 1; $i -le $tries; $i++) {
    try { $null = (& $sb); Write-Host ("STEP ok: " + $name); return $true }
    catch { Start-Sleep -Seconds 5 }
  }
  Write-Host ("STEP FAILED: " + $name)
  return $false
}
Step "ready" { if (-not $xl.Ready) { throw "not ready" } } 24 | Out-Null
if (-not (Step "open" { $global:wb = $xl.Workbooks.Open($path, 0, $false) })) { $code = 1 }
if ($code -eq 0) {
  Step "automatic calculation" { $xl.Calculation = -4105 } | Out-Null
  Step "calculate" { $xl.Calculate() } | Out-Null
  for ($i = 0; $i -lt 120; $i++) {                      # wait until Excel has finished calculating (up to 10 min)
    try { if ($xl.CalculationState -eq 0) { break } } catch {}
    Start-Sleep -Seconds 5
  }
  if (-not (Step "save" { $global:wb.Saved = $false; $global:wb.Save() })) { $code = 1 }
  Step "close" { $global:wb.Close($false) } | Out-Null
}
Step "quit" { $xl.Quit() } 6 | Out-Null
try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
exit $code
'''


def store_check_dp_values(path):
    """Store the value of every check_dp formula next to the formula, without Excel.

    The builder writes check_dp as a live formula (Country & Province & District & SIA_date & total & Response) with no
    stored result. Excel normally stores it when the file is saved once. This does the same thing directly in the file:
    the formula stays, its result is written beside it (dates as Excel's serial number, numbers in General format),
    so every reader (Excel, R, pandas) sees the value. Returns (number of cells filled, the file written: the same
    path, or '<name>_stored.xlsx' when Windows keeps the original locked)."""
    import zipfile
    from xml.sax.saxutils import escape, unescape
    path = Path(path)
    tmp = path.with_name(path.stem + ".tmp_values.xlsx")
    with zipfile.ZipFile(path) as zin:
        sheets = sorted(n for n in zin.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
        if not sheets:
            raise ValueError("no worksheet found in the file")
        name = sheets[0]
        xml = zin.read(name).decode("utf-8")
        if "<sharedStrings" in xml or 't="s"' in xml:
            raise ValueError("this file uses shared strings: open it in Excel and save it instead")
        cell_re = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
        row_re = re.compile(r"(<row [^>]*>)(.*?)(</row>)", re.S)
        n_done = 0

        def text_of(attrs, inner):
            if not inner:
                return ""
            if "inlineStr" in attrs:
                m = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
                return unescape(m.group(1)) if m else ""
            m = re.search(r"<v>(.*?)</v>", inner, re.S)
            if not m or m.group(1) == "":
                return ""
            v = unescape(m.group(1))
            if 't="b"' in attrs:
                return "TRUE" if v in ("1", "true") else "FALSE"
            if 't="str"' in attrs or 't="e"' in attrs:
                return v
            try:
                x = float(v)
                return str(int(x)) if x == int(x) and abs(x) < 1e15 else ("%.15g" % x)
            except ValueError:
                return v

        def do_row(m):
            nonlocal n_done
            head, body, tail = m.group(1), m.group(2), m.group(3)
            if "<f>" not in body:
                return m.group(0)
            cells = {}
            for c in cell_re.finditer(body):
                cells[c.group(1)] = text_of(c.group(3), c.group(4))

            def fill(c):
                nonlocal n_done
                inner = c.group(4) or ""
                f = re.match(r"<f>([^<]*)</f><v\s*/?>(?:</v>)?$", inner)
                if not f:
                    return c.group(0)
                refs = re.findall(r"([A-Z]+)\d+", unescape(f.group(1)))
                val = "".join(cells.get(r, "") for r in refs)
                attrs = re.sub(r'\s+t="[^"]*"', "", c.group(3))
                n_done += 1
                return '<c r="%s%s"%s t="str"><f>%s</f><v>%s</v></c>' % (c.group(1), c.group(2), attrs, f.group(1), escape(val))
            return head + cell_re.sub(fill, body) + tail
        new_xml = row_re.sub(do_row, xml)
        if n_done == 0:
            raise ValueError("no formula without a stored value was found")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, new_xml.encode("utf-8") if item.filename == name else zin.read(item.filename))
    for _ in range(4):                       # the target can be held for a moment by Excel, a preview pane, a sync tool
        try:
            os.replace(tmp, path)
            return n_done, path
        except PermissionError:
            time.sleep(2)
    # still locked: keep the result under a new name next to it (the manifest then points to it)
    alt = path.with_name(path.stem + "_stored.xlsx")
    if alt.exists():
        try:
            alt.unlink()
        except PermissionError:
            alt = path.with_name(path.stem + "_stored_%s.xlsx" % dt.datetime.now().strftime("%H%M%S"))
    os.replace(tmp, alt)
    return n_done, alt


def checkdp_filled(path):
    """True when the check_dp column holds stored values (not empty) at the top, the middle and the end."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        it = ws.iter_rows(values_only=True)
        head = next(it)
        if "check_dp" not in head:
            return False, "no check_dp column"
        j = head.index("check_dp")
        vals = [r[j] if j < len(r) else None for r in it]
    finally:
        wb.close()
    if not vals:
        return False, "no data rows"
    probe = vals[:200] + vals[len(vals) // 2: len(vals) // 2 + 200] + vals[-200:]
    empty = sum(1 for v in probe if v in (None, ""))
    return empty == 0, "%d of %d probed cells empty" % (empty, len(probe))


def cmd_excel(a):
    """Store the check_dp values. Default: written straight into the file, formula kept (no Excel needed).
    --use-excel: desktop Excel opens, calculates and saves it (falls back to the direct method if Excel fails).
    --verify-only: you saved the file yourself; only check the values."""
    n, rd, m = latest_run(unpublished=True)
    if not n:
        die("no build waiting: run 'build' first.")
    if not m.get("check", {}).get("ok"):
        die("v%d has not passed 'check' yet." % n)
    f = run_path(rd, m["build"]["file"])
    f_holder = [f]
    method = "manual"

    def verify():
        try:
            return checkdp_filled(f_holder[0])
        except Exception as e:
            return False, "the file could not be read (%s): close any Excel window that still has it open" % type(e).__name__

    def python_store():
        log("Storing the check_dp values in the file (the formulas stay) ...")
        try:
            k, newf = store_check_dp_values(f)
        except Exception as e:
            die("could not store the values: %s. Open the file in Excel, save once and run 'excel --verify-only'." % e)
        log("check_dp: %d values stored." % k)
        if Path(newf) != f:
            log("The original file is locked by another program, so the result is saved as %s." % Path(newf).name)
            m["build"]["file"] = str(newf)
            f_holder[0] = Path(newf)

    if not a.verify_only:
        want_excel = getattr(a, "use_excel", False) or str(S.get("excel_method", "python")).lower() == "excel"
        if not want_excel:
            ok, why = verify()
            if not ok:
                python_store()
            method = "values stored by python (formulas kept)"
        else:
            if os.name != "nt":
                die("the Excel method needs Windows with Excel.")
            ps = rd / "_save_in_excel.ps1"
            ps.write_text(PS_SAVE, encoding="utf-8")
            log("Opening %s in Excel, recalculating and saving (a minute or two) ..." % f.name)
            try:
                r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps), str(f),
                                    "1" if S.get("excel_visible", True) else "0"],
                                   capture_output=True, text=True, errors="replace",
                                   timeout=int(S["excel_timeout_min"]) * 60)
                if r.returncode:
                    log("Excel reported a problem: %s" % (r.stdout.strip() or r.stderr.strip())[-300:].replace("\n", " | "))
            except subprocess.TimeoutExpired:
                log("Excel did not finish in %s minutes." % S["excel_timeout_min"])
            try:
                ps.unlink()
            except Exception:
                pass
            ok, why = verify()
            method = "excel automation"
            if not ok:
                log("Excel did not store the values (a copy of Office that is not activated opens files read-only). "
                    "Storing them directly instead.")
                python_store()
                method = "values stored by python (formulas kept) after Excel failed"
    ok, why = verify()
    m["excel"] = {"ok": ok, "at": now(), "method": method, "detail": why}
    save_manifest(rd, m)
    if not ok:
        die("the check_dp values are still empty (%s). Open the file in Excel, save it, and run 'excel --verify-only'." % why)
    log("check_dp OK: %s (%s)." % (why, method))
    log("Next: python admin.py publish")


# ----------------------------------------------------------------------------------------------- publish
def verify_published(built, snap, row):
    """After the publish: the backup just taken of the live file must be the file that was built (same bytes, or at least the same content).
    Returns (True / False / None, message). A False means the file on SharePoint is NOT what was built: look at it before anyone uses it."""
    try:
        if not snap or not Path(snap).is_file():
            return None, "the new live file could not be read back (no backup of it): check it by hand"
        if hashlib.sha256(Path(built).read_bytes()).hexdigest() == hashlib.sha256(Path(snap).read_bytes()).hexdigest():
            return True, "the live file on SharePoint is byte for byte the build (backup version %s)" % (row or {}).get("version")
        if fingerprint(load_norm(built)) == fingerprint(load_norm(snap)):
            return True, "the live file on SharePoint has the content of the build (backup version %s; the bytes differ only in how SharePoint stores the file)" % (row or {}).get("version")
        return False, "the live file on SharePoint (backup version %s) does NOT have the content of the build" % (row or {}).get("version")
    except Exception as e:
        return None, "the check of the published file could not run (%s): check it by hand" % e


def cmd_publish(a):
    n, rd, m = latest_run(unpublished=True)
    if not n:
        die("no build waiting to be published.")
    f = run_path(rd, m["build"]["file"])
    if not m.get("check", {}).get("ok"):
        die("v%d has not passed 'check' (run it, and fix what it stops on)." % n)
    if not m.get("excel", {}).get("ok"):
        die("v%d was not saved in Excel yet (run 'excel')." % n)
    if not f.is_file():
        die("the build file is missing: %s" % f)
    ok, why = checkdp_filled(f)
    if not ok:
        die("the check_dp values of %s are empty (%s). Run 'excel' again." % (f.name, why))
    # the live file must still be the one the build started from
    row, snap = do_backup(show=False)
    log("Comparing the live file with the one v%d started from ..." % n)
    live = load_norm(snap)
    bp = base_path_of(m)
    base_fp = fingerprint(load_norm(bp)) if bp else m["base"]["fingerprint"]   # same normalisation on both sides
    same = fingerprint(live) == base_fp
    if not same:
        base_cols = m["base"]["columns"]
        extra = [c for c in live.columns if c not in base_cols]
        die("the live SharePoint file changed after v%d was built (now backup version %s, saved by %s). "
            "Publishing now would overwrite that change. Run 'build' again: it starts from the current live file.%s"
            % (n, row["version"], row["modified_by"], (" New columns: %s." % ", ".join(extra)) if extra else ""))
    log("The live file is the one v%d started from (backup version %s)." % (n, row["version"]))
    pub = P(S["publisher"])
    cmd = [sys.executable, pub, "--file", f, "--backup-repo", BACKUP_REPO]
    if not getattr(a, "ask", False) and "--yes" not in (a.extra or []):
        cmd.append("--yes")        # every gate has passed: go ahead without typing YES (publish --ask to be asked)
    extra = list(a.extra or [])
    res_json = rd / "publish_result.json"
    if res_json.exists():
        res_json.unlink()
    cmd += ["--result-json", res_json, "--dictionary-backup-dir", DICT_BACKUPS]
    dfile = rd / ("AFRO_admin_data_dictionary_v%d.xlsx" % n)
    if "--no-dictionary" not in extra and "--dictionary" not in extra:
        if dfile.is_file():
            cmd += ["--dictionary", dfile]
        else:
            log("WARNING: %s is not in the run folder: the data dictionary will not be pushed." % dfile.name)
    cmd += extra
    rc, _ = run_cmd(cmd, cwd=HERE, tee=rd / "publish.log")
    if rc and rc != DICT_FAILED_RC:
        die("the publish script stopped (exit %s). Nothing is marked as published; see %s." % (rc, rd / "publish.log"))
    res = read_json(res_json, {})
    row2, snap2 = latest_snapshot()
    m["published"] = {"at": now(), "backup_version": row2["version"] if row2 else None, "file": str(f),
                      "dictionary": res.get("dictionary", "failed" if rc else "unknown"),
                      "dictionary_message": res.get("dictionary_message", "")}
    m["published"]["verified"], m["published"]["verify_message"] = verify_published(f, snap2, row2)
    save_manifest(rd, m)
    st = read_json(STATE, {})
    st["published"] = {"version": n, "at": m["published"]["at"], "backup_version": m["published"]["backup_version"]}
    write_json(STATE, st)
    log("v%d is live. Local backup of the new live file: version %s." % (n, m["published"]["backup_version"]))
    if m["published"]["verified"]:
        log("VERIFIED: %s." % m["published"]["verify_message"])
    else:
        log("WARNING - NOT VERIFIED: %s." % m["published"]["verify_message"])
    dstat = m["published"]["dictionary"]
    if dstat == "published":
        log("The data dictionary on SharePoint is now AFRO_admin_data_dictionary_v%d.xlsx." % n)
    elif dstat == "unchanged":
        log("The data dictionary on SharePoint was already up to date.")
    elif dstat == "failed":
        log("WARNING: the data dictionary was NOT pushed (%s). The admin data is fine. Fix the cause, then run: "
            "python admin.py dictionary" % m["published"]["dictionary_message"])
    else:
        log("The data dictionary was not pushed (%s)." % dstat)
    log("Next: commit the code if it changed ('commit'), and tell the team the new version is live.")


def cmd_dictionary(a):
    """Push only the data dictionary of the last published build (or of the newest build if none was published)."""
    pick = None
    for n, d in reversed(run_dirs()):
        m = read_json(d / "manifest.json", {})
        if m.get("published"):
            pick = (n, d, m)
            break
    if not pick:
        n, d, m = latest_run()
        pick = (n, d, m) if n else None
    if not pick:
        die("no build yet: run 'build' first.")
    n, rd, m = pick
    dfile = rd / ("AFRO_admin_data_dictionary_v%d.xlsx" % n)
    if not dfile.is_file():
        die("%s does not exist (the build of v%d produced no dictionary)." % (dfile.name, n))
    log("Pushing %s to SharePoint ..." % dfile.name)
    res_json = rd / "dictionary_result.json"
    cmd = [sys.executable, P(S["publisher"]), "--dictionary-only", "--dictionary", dfile, "--yes",
           "--result-json", res_json, "--dictionary-backup-dir", DICT_BACKUPS, "--backup-repo", BACKUP_REPO]
    if getattr(a, "dry_run", False):
        cmd.append("--dry-run")
    cmd += list(a.extra or [])
    rc, _ = run_cmd(cmd, cwd=HERE, tee=rd / "dictionary.log")
    res = read_json(res_json, {})
    if rc:
        die("the dictionary was not pushed (see %s)." % (rd / "dictionary.log"))
    if not getattr(a, "dry_run", False):
        if m.get("published"):
            new_state = res.get("dictionary", "unknown")
            old_state = m["published"].get("dictionary")
            # "unchanged" only says the dictionary was already current: keep "published" if that is what was recorded
            if not (new_state == "unchanged" and old_state in ("published", "unchanged")):
                m["published"]["dictionary"] = new_state
                m["published"]["dictionary_message"] = res.get("dictionary_message", "")
        save_manifest(rd, m)
        log("Dictionary: %s." % res.get("dictionary", "done"))


# ----------------------------------------------------------------------------------------------- all
def cmd_all(a):
    class X:  # reuse the sub-commands with their defaults
        pass
    b = X(); b.base = None; b.no_backup = False; b.extra = a.extra
    cmd_build(b)
    c = X(); c.run = None; c.strict = False
    cmd_check(c)
    e = X(); e.verify_only = False; e.use_excel = False
    cmd_excel(e)
    p = X(); p.extra = []; p.ask = False
    cmd_publish(p)


# ----------------------------------------------------------------------------------------------- git
def cmd_commit(a):
    if not a.files:
        die("name the files to commit, e.g.: python admin.py commit -m \"message\" add_coadmin_columns.py README.md")
    low = a.message.lower()
    if any(x in low for x in FORBIDDEN_MSG):
        die("the message must not contain Co-Authored-By / Claude-Session lines.")
    for f in a.files:
        name = Path(f).name.lower()
        if name.endswith(DATA_EXT) and name not in DATA_OK:
            die("%s is a data file: data is never committed (only lookup.xlsx and geo_reference.csv)." % f)
        if "secrets" in name:
            die("%s looks like a secrets file: never committed." % f)
    rc, out = git("add", "--", *a.files)
    if rc:
        die("git add failed: %s" % out)
    rc, out = git("diff", "--cached", "--quiet")
    if rc == 0:
        die("nothing to commit for these files (no change).")
    rc, out = git("commit", "-m", a.message)
    print(out)
    if rc:
        die("git commit failed.")
    if not a.no_push:
        rc, out = git("push")
        print(out)
        if rc:
            die("the commit is made but the push failed: run 'git push' yourself.")
    log("Committed%s: %s" % ("" if a.no_push else " and pushed", a.message))


# ----------------------------------------------------------------------------------------------- scope sync, doctor
def cmd_sync_scope(a):
    sd, dd = S.get("scope_script_dir"), S.get("scope_data_dir")
    if not (sd or dd):
        die("set scope_script_dir / scope_data_dir in admin_settings.json first.")
    if sd:
        Path(sd).mkdir(parents=True, exist_ok=True)
        for name in (S["builder"], S["publisher"]):
            shutil.copy2(P(name), Path(sd) / Path(name).name)
            log("copied %s -> %s" % (name, sd))
    n, d, m = latest_run()
    if dd and n:
        up = Path(dd) / "upload_to_sharepoint"
        up.mkdir(parents=True, exist_ok=True)
        for p in d.glob("*_v%d.xlsx" % n):
            tgt = up if ("AFRO_admin_data_v" in p.name or "dictionary" in p.name) else Path(dd)
            shutil.copy2(p, tgt / p.name)
        log("copied the v%d files to %s" % (n, dd))


def cmd_sync_ops(a):
    """Copy the decision files (corrections.csv, lookup_overrides.csv, scope_.csv) to the inputs/ folder of the AdminOps folder, so the Operations Center wires the data the way
    this PC does. Unchanged files are skipped; a file in the store is saved in workspace/ops_backups before it is replaced. --dry-run lists what would change."""
    files = {}
    for name, path in (("corrections.csv", S.get("corrections_file")), ("lookup_overrides.csv", S.get("lookup_overrides_file")), ("scope_.csv", scope_path())):
        if path and P(path).is_file():
            files[name] = str(P(path))
        else:
            log("%s: no local file (%s), not copied." % (name, path or "not set"))
    if not files:
        die("nothing to copy.")
    tool = HERE / "sync_ops.py"
    cmd = [sys.executable, str(tool), "--state", str(WORK / "sync_ops_state.json"), "--backup-dir", str(WORK / "ops_backups"),
           "--env-file", str(BACKUP_REPO / "config" / "secrets.env")]
    for n, f in files.items():
        cmd += ["--file", "%s=%s" % (n, f)]
    if getattr(a, "dry_run", False):
        cmd.append("--dry-run")
    if getattr(a, "force", False):
        cmd.append("--force")
    log("Copying the decision files to the AdminOps folder (inputs/) ...")
    rc, text = run_cmd(cmd, cwd=HERE)
    if rc != 0:
        die("the copy to AdminOps did not complete; see above. Nothing else was changed.")


def cmd_doctor(a):
    ok = True

    def line(good, msg):
        nonlocal ok
        ok &= good
        print("  [%s] %s" % ("ok" if good else "!!", msg))
    print("\nChecking the set-up:\n")
    line(sys.version_info >= (3, 9), "python %s" % sys.version.split()[0])
    for mod in ("pandas", "openpyxl", "numpy", "sklearn"):
        try:
            __import__(mod)
            line(True, "package %s" % mod)
        except Exception:
            line(mod == "sklearn", "package %s%s" % (mod, " (optional)" if mod == "sklearn" else " - install with: pip install " + mod))
    for label, p in (("builder", P(S["builder"])), ("publisher", P(S["publisher"])), ("lookup", P(S["lookup"])),
                     ("geo reference", P(S["geo"])), ("backup tool", BACKUP_REPO / "sp_versions.py")):
        line(p.is_file(), "%s: %s" % (label, p))
    line((BACKUP_REPO / "config" / "secrets.env").is_file(), "SharePoint credentials file present (not read here)")
    line(shutil.which("git") is not None, "git")
    if os.name == "nt":
        try:
            r = subprocess.run(["powershell", "-NoProfile", "-Command",
                                "[bool](Get-ItemProperty 'HKLM:\\SOFTWARE\\Classes\\Excel.Application' -ErrorAction SilentlyContinue)"],
                               capture_output=True, text=True, timeout=30)
            line("True" in r.stdout, "Excel installed (for the automatic save)")
        except Exception:
            line(False, "Excel could not be detected")
    else:
        print("  [--] Excel step is for Windows only")
    print("\n  %s\n" % ("All good." if ok else "Fix the lines marked !! first."))


# ----------------------------------------------------------------------------------------------- sources
SHOW_COLS = ["Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response", "PopPolio",
             "TotalNbrVaccPolio", "CVPolio"]


def resolve_source(spec):
    """'live' (newest local backup), 'vNN' (a build), 'bN' (backup version N) or a file path."""
    if not spec or spec == "live":
        row, path = latest_snapshot()
        if row is None:
            die("no local backup yet: run 'backup'.")
        return path, "backup version %s (%s)" % (row["version"], row["modified_by"])
    m = re.match(r"^v(\d+)$", spec)
    if m:
        d = RUNS / ("v%s" % m.group(1))
        mf = read_json(d / "manifest.json", {})
        if not mf.get("build", {}).get("file"):
            die("build %s not found." % spec)
        return run_path(d, mf["build"]["file"]), "build %s" % spec
    m = re.match(r"^b(\d+)$", spec)
    if m:
        for r in catalog_rows("admin"):
            if r["version"] == m.group(1):
                return BACKUP_REPO / r["snapshot"], "backup version %s (%s)" % (r["version"], r["modified_by"])
        die("backup version %s not found." % m.group(1))
    p = Path(spec)
    if not p.is_file():
        die("file not found: %s" % spec)
    return p, p.name


# ----------------------------------------------------------------------------------------------- inspect
def hygiene(df, builder, lk):
    """Checks that are not in the portal quality report: stray characters, names, the lookup."""
    out = []
    for c in ("Country", "Province", "District", "Response", "Vaccine_type", "Round_Add"):
        if c in df.columns:
            bad = df[c].astype(str).str.contains(r"_x000D_|\s\s|^\s|\s$|[\r\n\t]", regex=True, na=False)
            if bad.any():
                out.append(("Stray characters or spaces in %s" % c, int(bad.sum()), ", ".join(sorted(set(df.loc[bad, c].astype(str).str.strip()))[:4])))
    if "Response" in df.columns:
        bad = response_country_mismatches(df, builder, lk)
        if bad:
            out.append(("Response of another country (per the lookup)", sum(bad.values()),
                        "; ".join("%s / %s" % k for k in list(bad)[:4])))
    for c in ("Country", "Province", "District", "SIA_date", "Round_Add", "Vaccine_type", "Response"):
        if c in df.columns:
            e = int((df[c].astype(str).str.strip() == "").sum())
            if e:
                out.append(("Empty %s" % c, e, ""))
    return out


def cmd_inspect(a):
    if not getattr(a, "no_refresh_lookup", False):
        refresh_lookup_step()
    path, label = resolve_source(a.source)
    log("Inspecting %s" % label)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    od = WORK / "inspect" / stamp
    od.mkdir(parents=True, exist_ok=True)
    cmd = [sys.executable, P(S["builder"]), path, "--check-only", "--out-dir", od, "--lookup", P(S["lookup"])]
    if lookup_overrides_path():
        cmd += ["--lookup-overrides", lookup_overrides_path()]
    if P(S["geo"]).is_file():
        cmd += ["--geo", P(S["geo"])]
    if a.country:
        cmd += ["--country", a.country]
    rc, text = run_cmd(cmd, cwd=HERE, tee=od / "inspect.log")
    if rc:
        die("the quality check failed (see %s)." % (od / "inspect.log"))
    pd = _pd()
    builder = import_builder()
    lk = builder.load_lookup(str(P(S["lookup"])), lookup_overrides_path())
    df = load_norm(path)
    if a.country:
        df = df[df["Country"].map(builder._qcountry) == builder._qcountry(a.country)]
    print("\nHygiene checks:")
    hy = hygiene(df, builder, lk)
    for t, k, ex in hy:
        print("  %-50s %6d  %s" % (t, k, ex))
    if not hy:
        print("  nothing found")
    rep, nodata = builder.lookup_crosscheck(df, lk)
    bad = rep[rep["Status"] != "OK"]
    print("\nCampaign rounds against the lookup: %d of %d agree" % (len(rep) - len(bad), len(rep)))
    for st, k in bad["Status"].value_counts().items():
        print("  %-55s %4d rounds" % (st, k))
    with pd.ExcelWriter(od / "lookup_check.xlsx") as w:
        rep.to_excel(w, sheet_name="Lookup_check", index=False)
        nodata.to_excel(w, sheet_name="Lookup_rounds_without_data", index=False)
    print("\nReports in %s" % od)
    print("  AFRO_admin_data_quality_report.xlsx (portal checks), lookup_check.xlsx, inspect.log\n")


# ----------------------------------------------------------------------------------------------- duplicates
def cmd_duplicates(a):
    """List the duplicates of a file and what a build with --resolve-duplicates would do with each, so the fix can be decided
    before it is made. Nothing is changed."""
    path, label = resolve_source(a.source)
    log("Looking for duplicates in %s" % label)
    pd = _pd()
    builder = import_builder()
    df = load_norm(path)
    if a.country:
        df = df[df["Country"].map(builder._qcountry) == builder._qcountry(a.country)]
    df = df.copy()
    df["_row"] = df.index + 2          # the Excel row number
    drop, exact, groups = builder.find_duplicates(df, "complete" if a.keep != "first" else "first")
    n_exact = len(exact)
    same = groups[groups["Class"].str.startswith("same record")] if len(groups) else groups
    other = groups[~groups["Class"].str.startswith("same record")] if len(groups) else groups
    print()
    print("%s: %d records%s" % (label, len(df), (" of %s" % a.country) if a.country else ""))
    print("  exact copies (every value the same; removed by every build):        %d" % n_exact)
    print("  records entered twice (same names, same population and results):    %d records in %d groups"
          % (len(same), same["duplicate_group"].nunique() if len(same) else 0))
    print("  same names but other figures (different units, NOT duplicates):     %d records in %d groups"
          % (len(other), other["duplicate_group"].nunique() if len(other) else 0))
    if len(same):
        print("\nRecords entered twice (kept = the %s):" % ("most complete record, then the first" if a.keep != "first" else "first record"))
        for gid, g in same.groupby("duplicate_group", sort=False):
            r0 = g.iloc[0]
            print("\n  %s | %s | %s | %s | %s | %s | %s" % (r0["Country"], r0["Province"], r0["District"], r0["SIA_date"],
                                                         r0["Round_Add"], r0["Vaccine_type"], r0["Response"]))
            for r in g[g["Class"].str.startswith("same record")].itertuples():
                print("     row %-6s vaccinated %-8s target %-8s filled cells %-4s -> %s" %
                      (r.Excel_row, r.TotalNbrVaccPolio, r.PopPolio, r.Filled_cells, "KEEP" if r.Suggested == "keep" else "REMOVE"))
            print("     cells that differ: %s" % (g.iloc[0]["Columns_that_differ"] or "none"))
        print("\nTo remove them: build with the option 'remove the extra copy' (admin.py build -- --resolve-duplicates complete).")
        print("Or correct the rows listed above in the SharePoint file and publish again.")
    elif not n_exact:
        print("\nNo record is entered twice. Nothing to fix.")
    if not a.no_file:
        od = WORK / "inspect" / dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        od.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(od / "Duplicates.xlsx") as w:
            (groups if len(groups) else pd.DataFrame({"note": ["no duplicates"]})).to_excel(w, sheet_name="Duplicate_groups", index=False)
            if n_exact:
                exact.to_excel(w, sheet_name="Exact_copies", index=False)
        print("\nThe list is in %s" % (od / "Duplicates.xlsx"))
    print()


# ----------------------------------------------------------------------------------------------- submissions
def _submissions_ctx(builder):
    from types import SimpleNamespace
    path, label = resolve_source("live")
    log("Reading the live file (%s) ..." % label)
    live = load_norm(path)
    gp, lp = P(S["geo"]), P(S["lookup"])
    geo = builder.load_geo(str(gp)) if gp.is_file() else None
    lk = builder.load_lookup(str(lp), lookup_overrides_path()) if lp.is_file() else None
    key, blank = builder.dup_key(live)
    index = {}
    for i, k, bl in zip(live.index, key.values, blank.values):
        if not bl:
            index.setdefault(k, []).append(i)
    return SimpleNamespace(builder=builder, live=live, live_index=index, geo=geo, lookup=lk, live_label=label)


def submissions_pull_step(dry_run=False, explicit=False):
    """Copy the countries' files from the SharePoint folder Admin_data_submissions to submissions_dir (read only on SharePoint).
    pull_submissions.py reads the link and the sign-in values itself (the backup tool's secrets.env); this script never does.
    Returns True when the local copy is up to date."""
    tool = HERE / "pull_submissions.py"
    if not tool.is_file():
        if explicit:
            log("pull_submissions.py is not here: nothing to pull with.")
            return False
        return True            # a copy of the engine without the pull (the Operations Center reads the folder itself): use the files in place
    dest = P(S["submissions_dir"])
    cmd = [sys.executable, str(tool), "--dest", str(dest), "--max-mb", str(S.get("submissions_max_mb", 60)),
           "--env-file", str(BACKUP_REPO / "config" / "secrets.env")]
    if dry_run:
        cmd.append("--dry-run")
    log("Reading the files of the submissions folder from SharePoint ...")
    rc, _ = run_cmd(cmd, cwd=HERE)
    return rc == 0


def submissions_state_pull(dry_run=False):
    """Copy the DECISIONS on the submissions (ledger, report, prepared records) from the AdminOps folder of the Operations Center
    to workspace/submissions, read only: the app decides, the PC build adds exactly what the app accepted. The ledger in place is kept
    in workspace/submissions/ledger_backups. Returns 'adminops' (read), 'no link' (ADMIN_STORE_SHARE_URL not set) or 'failed'."""
    if os.environ.get("OPS_LEDGER_FROM_APP"):
        return "adminops"      # run by the Operations Center: the app reads its own ledger itself just before the job (workspace/submissions/ledger.json)
    tool = HERE / "pull_submissions.py"
    if not tool.is_file():
        return "no tool"
    cmd = [sys.executable, str(tool), "--ledger", "--dest", str(WORK / "submissions"),
           "--env-file", str(BACKUP_REPO / "config" / "secrets.env")]
    if dry_run:
        cmd.append("--dry-run")
    log("Reading the decisions on the submissions from the AdminOps folder ...")
    rc, text = run_cmd(cmd, cwd=HERE)
    if rc == 0:
        return "adminops"
    return "no link" if "no share link" in text else "failed"


def write_pending(sm, builder, ctx, out_root):
    """Pending admin data: lookup rounds with no records in the live file `pending_days` days after their start. Writes Pending_admin_data.xlsx and the register
    (pending_register.json) and prints the counts. The country feedback reads the sheet Pending_now, so this runs before it."""
    import pandas as pd
    if ctx.lookup is None:
        log("pending admin data: there is no lookup, nothing to compare with.")
        return None
    cc = builder.lookup_crosscheck(ctx.live, ctx.lookup)
    nodata = cc[1]
    days, since = int(S.get("pending_days") or 15), str(S.get("pending_since") or "2025-01-01")
    res = sm.pending(out_root, nodata, days=days, since=since, update=True, source="the live file (%s) and the lookup of this run" % ctx.live_label, live=ctx.live)
    scope = read_scope()
    accepted = []
    li_new, li_old = sm.lookup_issues(cc[0], ctx.lookup, since=since, ckey=builder._lkc, rkey=builder._rk, live=ctx.live, scope=scope, accepted=accepted)
    res["Lookup_issues"], res["Lookup_issues_older"] = li_new, li_old
    if accepted:
        res["Lookup_accepted"] = pd.DataFrame(accepted).sort_values("Start_in_live", ascending=False).reset_index(drop=True)
        print("Lookup: %d round(s) the lookup does not date exactly are placed by the GPEI round window (live date inside the window, or the districts are those of that round): sheet Lookup_accepted, nothing to ask." % len(accepted))
    target = sm.writable(Path(out_root) / "Pending_admin_data.xlsx", log)
    with pd.ExcelWriter(target, engine="openpyxl") as xw:
        for name_, d_ in res.items():
            d_.to_excel(xw, sheet_name=name_[:31], index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                w = max(len(str(c.value)) if c.value is not None else 0 for c in col[:60])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, w + 2), 70)
    pn, rn_ = res["Pending_now"], res["Received_not_in_live"]
    print("(%d more round(s) have live records of the country near the start date: sheet Check_other_name, they are not asked of the country until checked.)" % len(res["Check_other_name"]))
    print("\nPending admin data (no records in the live file %d days after the round start, rounds since %s): %d round(s) with nothing received in %d country(ies); "
          "%d more received in a submitted file but not in the live file yet; %d resolved since the register started; %d not yet due." % (
              days, since, len(pn), pn["Country"].nunique() if len(pn) else 0, len(rn_), len(res["Resolved"]), len(res["Not_yet_due"])))
    print("Lookup issues (rounds in the live file that the lookup cannot place, for the lookup maintainers; nothing is changed on them): %d since %s, %d older (sheets Lookup_issues, Lookup_issues_older)." % (len(li_new), since, len(li_old)))
    if len(li_new):
        print(li_new["Issue"].value_counts().to_string())
    if len(res["By_country"]):
        print(res["By_country"].head(15).to_string(index=False))
    print("Written: %s (register: %s)\n" % (target, Path(out_root) / "pending_register.json"))
    return res


def lookup_overrides_path():
    """The file of lookup changes a person decided (setting lookup_overrides_file), or None when there is none."""
    f = S.get("lookup_overrides_file")
    p = P(f) if f else None
    return str(p) if p and p.is_file() else None


def scope_path():
    """The path of scope_.csv: the setting scope_file, else one level above scope_data_dir ('' when neither is set)."""
    return S.get("scope_file") or (str(P(S.get("scope_data_dir")).parent / "scope_.csv") if S.get("scope_data_dir") else "")


def read_scope():
    """scope_.csv (districts planned per campaign) from the setting scope_file, or from one level above scope_data_dir; None when there is none."""
    pd = _pd()
    sf = scope_path()
    if sf and Path(sf).is_file():
        try:
            df = pd.read_csv(sf, dtype=str)
            log("Scope file %s (%d rows) is used as evidence (districts planned per campaign)." % (Path(sf).name, len(df)))
            return df
        except Exception as e:
            log("The scope file %s could not be read (%s): without it." % (sf, e))
    else:
        log("No scope file (setting scope_file, or scope_.csv one level above scope_data_dir): without the scope evidence.")
    return None


def cmd_lookup_fix(a):
    """From the lookup issues to DECIDED changes. `propose` (default) lists, from the live file, the lookup and the scope, the corrections (live campaign renamed to the lookup
    name) and the lookup changes (a round added, a campaign entered twice removed) that the evidence supports, and writes them to workspace/submissions/lookup_fix_proposals.xlsx: nothing is
    applied. `adopt --by NAME` writes them as decided lines into corrections.csv and lookup_overrides.csv (the build applies both); old files are kept as .bak copies."""
    pd = _pd()
    sm = import_submissions()
    builder = import_builder()
    path, label = resolve_source(getattr(a, "source", None) or "live")
    log("Reading the live file (%s) ..." % label)
    live = load_norm(path)
    lk = builder.load_lookup(str(P(S["lookup"])), lookup_overrides_path())
    scope = read_scope()
    if scope is None:
        if a.action == "auto":
            log("Lookup wiring: no scope file, nothing decided by the engine.")
            return
        die("lookup-fix needs the scope file: set scope_file (or scope_data_dir) in admin_settings.json.")
    cc = builder.lookup_crosscheck(live, lk)
    since = str(S.get("pending_since") or "2025-01-01")
    n_, o_ = sm.lookup_issues(cc[0], lk, since=since, ckey=builder._lkc, rkey=builder._rk, live=live, scope=scope)
    issues = pd.concat([n_, o_], ignore_index=True) if getattr(a, "older", False) else n_   # older rounds (before pending_since) only with --older: they were not reviewed
    choose = {}
    for c in (a.choose or []):
        if "=" not in c:
            die("--choose wants LIVE CAMPAIGN=LOOKUP CAMPAIGN, got: %s" % c)
        k, v = c.split("=", 1)
        choose[k.strip()] = v.strip()
    corr, chg, skipped = sm.propose_fixes(issues, lk, ckey=builder._lkc, rkey=builder._rk, choose=choose, rep=cc[0], vset=builder._vset, live=live)
    if a.action == "auto":
        review = {"corr": corr[corr["Tier"] == "review"] if len(corr) else corr, "chg": chg[chg["Tier"] == "review"] if len(chg) else chg}
        corr = corr[corr["Tier"].isin(["auto", "decided"])] if len(corr) else corr
        chg = chg[chg["Tier"].isin(["auto", "decided"])] if len(chg) else chg
    else:
        review = None
    print("\nLookup issues considered: %d (rounds since %s%s). Supported by the scope evidence: %d correction(s) of a live campaign name, %d lookup change(s). Not proposed: %d.\n" % (
        len(issues), since, ", and the older ones" if getattr(a, "older", False) else "; %d older ones are not proposed (--older)" % len(o_), len(corr), len(chg), len(skipped)))
    if len(corr):
        print("Corrections (live campaign -> lookup name):")
        print(corr[["Country", "Response", "Round_Add", "Restored"]].to_string(index=False))
    if len(chg):
        print("\nLookup changes:")
        print(chg[["Action", "Country", "OBR Name", "Round Number", "Round Start Date"]].to_string(index=False))
    if skipped:
        print("\nNot proposed:")
        for w_, why in skipped:
            print("  %s\n      %s" % (w_, why))
    out = WORK / "submissions" / "lookup_fix_proposals.xlsx"      # in the synced folder: the Operations Center shows it
    out.parent.mkdir(parents=True, exist_ok=True)
    target = sm.writable(out, log)
    with pd.ExcelWriter(target, engine="openpyxl") as xw:
        (corr if len(corr) else pd.DataFrame(columns=["Country", "Response", "Round_Add", "Restored"])).to_excel(xw, sheet_name="Corrections", index=False)
        (chg if len(chg) else pd.DataFrame(columns=builder.OVERRIDE_COLS)).to_excel(xw, sheet_name="Lookup_changes", index=False)
        pd.DataFrame(skipped, columns=["Round", "Why not proposed"]).to_excel(xw, sheet_name="Not_proposed", index=False)
        issues.to_excel(xw, sheet_name="Lookup_issues", index=False)
    print("\nWritten: %s" % target)
    if review is not None and (len(review["corr"]) or len(review["chg"])):
        print("\nLeft to a person (a rename that would change how the vaccines are treated, or other doubt):")
        for _, r_ in review["corr"].iterrows():
            print("  %s | %s %s -> %s | %s" % (r_["Country"], r_["Response"], r_["Round_Add"], r_["Restored"], r_["Tier_note"]))
    if a.action not in ("adopt", "auto"):
        print("Nothing was changed. To decide them: python admin.py lookup-fix adopt --by \"Your name\"  (a campaign the scope plans twice needs --choose \"LIVE NAME=LOOKUP NAME TO KEEP\").\n")
        return
    if a.action == "auto" and not a.by:
        a.by = "engine rule (lookup wiring: strong scope evidence, vaccine treatment unchanged)"
    if not a.by:
        die("adopt needs --by \"name\": the decision is written with the name of the person who decides.")
    import datetime as dt
    stamp, today = dt.datetime.now().strftime("%Y%m%d_%H%M%S"), dt.datetime.now().strftime("%Y-%m-%d")

    def append(path_s, new, keycols, cols):
        p = P(path_s)
        cols = list(dict.fromkeys(cols))
        old = pd.read_csv(p, dtype=str, keep_default_na=False, encoding="utf-8-sig") if p.is_file() else pd.DataFrame(columns=cols)
        if p.is_file():
            import shutil
            shutil.copyfile(p, str(p) + ".bak_" + stamp)
        new = new.drop(columns=[c for c in ("Tier", "Tier_note") if c in new.columns]).copy()
        new["Decided_by"], new["Decided_on"] = a.by, today
        have = set(tuple(str(x) for x in r) for r in old.reindex(columns=keycols).fillna("").values.tolist())
        keep = new[[tuple(str(x) for x in r) not in have for r in new.reindex(columns=keycols).fillna("").values.tolist()]]
        allc = list(dict.fromkeys(list(old.columns) + list(cols)))
        res = pd.concat([old, keep], ignore_index=True).reindex(columns=allc).fillna("")
        p.parent.mkdir(parents=True, exist_ok=True)
        res.to_csv(p, index=False, encoding="utf-8-sig")
        return len(keep), len(new) - len(keep), p
    if len(corr):
        n1, d1, p1 = append(S.get("corrections_file") or "workspace/corrections.csv", corr, ["Country", "Response", "Round_Add", "Column", "Found", "Restored"],
                            ["Country", "Province", "District", "Response", "Round_Add", "SIA_date", "Column", "Found", "Restored", "Reason", "Decided_by", "Decided_on"])
        print("corrections: %d line(s) added to %s, %d already there." % (n1, p1.name, d1))
    if len(chg):
        n2, d2, p2 = append(S.get("lookup_overrides_file") or "workspace/lookup_overrides.csv", chg, ["Action", "Country", "OBR Name", "Round Number", "Round Start Date"],
                            builder.OVERRIDE_COLS + ["Decided_by", "Decided_on"])
        print("lookup changes: %d line(s) added to %s, %d already there." % (n2, p2.name, d2))
    print("Decided by %s on %s. The next build applies them; `python admin.py corrections` shows what the corrections would change in the live file.\n" % (a.by, today))


def cmd_corrections(a):
    """Shows what the corrections file would do to the live file, changing nothing: applied / already in place / skipped, with the reason."""
    cf = P(S.get("corrections_file"))
    if not cf.is_file():
        print("\nThere is no corrections file (%s)." % cf)
        return
    builder = import_builder()
    path, label = resolve_source("live")
    log("Reading the live file (%s) ..." % label)
    live = load_norm(path)
    ch, rep_ = builder.apply_corrections(live, str(cf))
    print("\nCorrections file %s against the live file (%s):\n  would change %d cell(s); already in place %d; skipped %d" % (cf.name, label, rep_["applied"], rep_["already"], len(rep_["skipped"])))
    for line, why in rep_["skipped"][:20]:
        print("    line %s: %s" % (line, why))
    if ch:
        import pandas as pd
        t = pd.DataFrame([{"Line": l_, "Country": live.at[i, "Country"], "District": live.at[i, "District"], "Response": live.at[i, "Response"], "Round": live.at[i, "Round_Add"],
                           "Column": c_, "Found": f_, "New": v_} for i, c_, v_, f_, w_, l_ in ch])
        print("\n" + t.head(12).to_string(index=False) + ("\n  ... %d more" % (len(t) - 12) if len(t) > 12 else ""))
    print("\nThe build applies them (and lists them in the repairs of the audit file). Nothing was changed now.\n")


def cmd_submissions(a):
    sm = import_submissions()
    out_root = WORK / "submissions"
    if a.action in ("list", "auto", "summary", "feedback", "pending", "accept", "reject") and not getattr(a, "no_pull", False) and not getattr(a, "local", False):
        if a.action in ("accept", "reject"):
            die("accept and reject are made in the Operations Center (Submissions tab): the PC reads those decisions and does not make "
                "its own, so that there is one ledger. To decide here anyway, add --local (the decision then stays on this PC only).")
        st = submissions_state_pull()
        if st != "adminops":
            log("NOTE: the ledger of the Operations Center was not read (%s): this is the ledger kept on this PC." % st)
    if a.action == "pending":
        if a.close:
            if not a.targets:
                die("give the round to close, as a piece of 'COUNTRY|campaign|Round n' (see the sheet Pending_now).")
            for t_, m_ in sm.pending_close(out_root, a.targets, a.note or "", a.by or os.environ.get("USERNAME") or os.environ.get("USER") or "operator"):
                print("  %s: %s" % (t_, m_))
            print("Run 'python admin.py submissions pending' to refresh Pending_admin_data.xlsx and the country feedback.")
            return
        builder = import_builder()
        ctx = _submissions_ctx(builder)
        write_pending(sm, builder, ctx, out_root)
        return
    if a.action == "feedback":
        import pandas as pd
        f, md = sm.feedback(out_root)
        (out_root / "Feedback_to_countries.md").write_text(md, encoding="utf-8")
        f.to_excel(sm.writable(out_root / "Feedback_to_countries.xlsx", log), sheet_name="Feedback", index=False)
        print("\n" + md)
        print("Written: %s and %s (one line per country, file and matter).\n" % (out_root / "Feedback_to_countries.md", out_root / "Feedback_to_countries.xlsx"))
        fl = sm.feedback_files(out_root, out_root / "feedback_by_country")
        print("One workbook per country, ready to send, in %s:" % (out_root / "feedback_by_country"))
        for c_, p_ in fl:
            print("  %-34s %s" % (c_, Path(p_).name))
        print()
        return
    if a.action == "summary":
        audits = sorted(WORK.glob("runs/v*/AFRO_admin_data_coadmin_audit_v*.xlsx"), key=lambda p: int(re.search(r"_v(\d+)\.xlsx$", p.name).group(1)))
        import pandas as pd
        res = sm.summary(out_root, audits[-1] if audits else None)
        if not res:
            print("\nNothing to summarise: inspect the files first.\n")
            return
        target = sm.writable(out_root / "Submissions_summary.xlsx", log)
        with pd.ExcelWriter(target, engine="openpyxl") as xw:
            for name, d in res.items():
                d.to_excel(xw, sheet_name=name[:31], index=False)
        bc = res.get("By_country")
        print("\nNew records waiting, by country (taken_by_rules = automatic rules, waiting = needs a person):\n")
        if bc is not None:
            print(bc.head(15).to_string(index=False))
        bl = res.get("Blockers")
        if bl is not None and len(bl):
            print("\nWhat holds them back (records affected, who can fix it):\n")
            print(bl.head(12)[["Finding", "Files", "New_records_affected", "Who can fix it"]].to_string(index=False))
        mr = res.get("Missing_rounds")
        if mr is not None and len(mr):
            print("\nLookup rounds since 2025 with no data in the live file: %d; in a submitted file: %d; not in any file (ask the country): %d" % (
                len(mr), int(mr["In_a_submitted_file"].sum()), int((~mr["In_a_submitted_file"]).sum())))
        print("\nFull tables: %s\n" % target)
        return
    if a.action == "auto":
        c = sm.auto_candidates(out_root)
        led = sm.load_ledger(out_root)["files"]
        print("\n%s would take %d file(s) into a build (nobody has decided on them; a person's decision always wins):\n" % (sm.AUTO_RULE, len(c)))
        for n, pth, note in sorted(c):
            print("  %-80s %s" % (n[:80], note.split(": ", 1)[1]))
        print("\nEverything else needs a person. Turn it on with \"submissions_auto\": true in admin_settings.json.\n")
        return
    if a.action == "pull":
        if not submissions_pull_step(dry_run=a.dry_run, explicit=True):
            sys.exit(1)
        return
    out_root = WORK / "submissions"
    who = a.by or os.environ.get("USERNAME") or os.environ.get("USER") or "operator"
    if a.action == "list":
        led = sm.load_ledger(out_root)
        if not led["files"]:
            print("\nNo file inspected yet: run 'admin.py submissions inspect'.\n")
            return
        print()
        for h, e in sorted(led["files"].items(), key=lambda x: x[1].get("name", "")):
            print("  %-45s %-8s %-9s new %-5s same %-5s other figures %-4s %s" % (
                e["name"][:45], e.get("status", ""), e.get("decision") or "-", e.get("new", ""), e.get("existing_same", ""),
                e.get("existing_different", ""), "(superseded)" if e.get("superseded") else ("" if e.get("present", True) else "(not in the folder)")))
        print()
        return
    if a.action in ("accept", "reject"):
        if not a.targets:
            die("give the file name (or the start of its content hash).")
        bad = 0
        for t in a.targets:
            ok, msg = sm.decide(out_root, t, "accepted" if a.action == "accept" else "rejected", who, a.note or "")
            print(("  " if ok else "  NOT DONE: ") + msg)
            bad += 0 if ok else 1
        if bad:
            sys.exit(2)
        return
    # inspect
    if not getattr(a, "no_refresh_lookup", False):
        refresh_lookup_step()
    folder = Path(a.folder) if a.folder else P(S["submissions_dir"])
    if not a.folder and not getattr(a, "no_pull", False):
        st = submissions_state_pull()          # earlier decisions stay: the ledger of the app first, then the files
        if st != "adminops":
            log("NOTE: the ledger of the Operations Center was not read (%s): the ledger kept on this PC is used." % st)
        if not submissions_pull_step():
            log("WARNING: the submissions folder could NOT be read from SharePoint: %s" % (
                "the copy already in %s is used and may be out of date." % folder if folder.is_dir() else "there is no local copy either."))
    if not folder.is_dir():
        die("the submissions folder does not exist: %s (run 'admin.py submissions pull', or give --folder)." % folder)
    builder = import_builder()
    ctx = _submissions_ctx(builder)
    cols = list(ctx.live.columns)
    rows = sm.inspect_folder(ctx, folder, out_root, cols, log=log)
    ready = sum(1 for r in rows if r["status"] == "ready")
    n_na = sum(1 for r in rows if str(r.get("why") or "").startswith("NOT ADMIN DATA"))
    print("\n%d file(s): %d ready, %d to review, %d blocked (%d of them not admin data: IM, LQAS, AEFI, supplies, call-in, summaries), %d with nothing new (tally sheets are checks)." % (
        len(rows), ready, sum(1 for r in rows if r["status"] == "review"), sum(1 for r in rows if r["status"] == "blocked"), n_na, sum(1 for r in rows if r["status"] == "nothing new")))
    print("Records: %d new and clean (they go in by themselves when submissions_auto is on), %d held back with a reason (sheet Held_back_records)." % (
        sum(int(r.get("clean_new") or 0) for r in rows), sum(int(r.get("held_new") or 0) for r in rows)))
    # pending admin data (a round with no data in the live file 15 days after its start), then the summary and the feedback to the countries:
    # they are part of every inspection, nobody has to remember a second command
    try:
        write_pending(sm, builder, ctx, out_root)
    except Exception as ex:
        log("the pending admin data could not be written (%s): run 'submissions pending'" % ex)
    try:
        import pandas as pd
        audits = sorted(WORK.glob("runs/v*/AFRO_admin_data_coadmin_audit_v*.xlsx"), key=lambda p: int(re.search(r"_v(\d+)\.xlsx$", p.name).group(1)))
        res = sm.summary(out_root, audits[-1] if audits else None)
        if res:
            with pd.ExcelWriter(sm.writable(out_root / "Submissions_summary.xlsx", log), engine="openpyxl") as xw:
                for name_, d_ in res.items():
                    d_.to_excel(xw, sheet_name=name_[:31], index=False)
        f_, md_ = sm.feedback(out_root)
        (out_root / "Feedback_to_countries.md").write_text(md_, encoding="utf-8")
        f_.to_excel(sm.writable(out_root / "Feedback_to_countries.xlsx", log), sheet_name="Feedback", index=False)
        fl_ = sm.feedback_files(out_root, out_root / "feedback_by_country")
        print("Also written: %s, %s and %s; one workbook per country (%d) in %s." % (out_root / "Submissions_summary.xlsx", out_root / "Feedback_to_countries.xlsx",
              out_root / "Feedback_to_countries.md", len(fl_), out_root / "feedback_by_country"))
    except Exception as ex:
        log("the summary and the feedback could not be written (%s): run 'submissions summary' and 'submissions feedback'" % ex)
    print("Next: look at %s. Accept or reject in the Operations Center (Submissions tab): the build here reads those decisions. "
          "What 'inspect' wrote on this PC is a preview; it is replaced by the app's ledger at the next build.\n" % (out_root / "Submissions_report.xlsx"))


# ----------------------------------------------------------------------------------------------- find
def cmd_find(a):
    path, label = resolve_source(a.source)
    df = load_norm(path)
    df.index = df.index + 2          # the Excel row number
    m = _pd().Series(True, index=df.index)
    for col, val in (("Country", a.country), ("Province", a.province), ("District", a.district),
                     ("Response", a.response), ("Vaccine_type", a.vaccine)):
        if val and col in df.columns:
            m &= df[col].str.upper().str.contains(re.escape(val.upper()), regex=True, na=False)
    if a.round and "Round_Add" in df.columns:
        m &= df["Round_Add"].str.upper() == a.round.upper()
    if a.date and "SIA_date" in df.columns:
        m &= df["SIA_date"].str.startswith(a.date)
    r = df[m]
    cols = [c for c in SHOW_COLS if c in r.columns] + [c for c in (a.cols or []) if c in r.columns]
    print("\n%s: %d record(s) match\n" % (label, len(r)))
    if r.empty:
        return
    if a.summary:
        g = r.assign(_t=_pd().to_numeric(r.get("TotalNbrVaccPolio"), errors="coerce")) \
             .groupby([c for c in ("Country", "Response", "Round_Add", "SIA_date", "Vaccine_type") if c in r.columns], dropna=False) \
             .agg(records=("Country", "size"), vaccinated=("_t", "sum")).reset_index()
        g["SIA_date"] = g["SIA_date"].str[:10]
        print(g.to_string(index=False))
    else:
        t = r[cols].head(a.limit).copy()
        if "SIA_date" in t.columns:
            t["SIA_date"] = t["SIA_date"].str[:10]
        print(t.to_string())
        if len(r) > a.limit:
            print("\n... %d more (use --limit N, or --summary)" % (len(r) - a.limit))
    print()


# ----------------------------------------------------------------------------------------------- compare
def cmd_compare(a):
    pa, la = resolve_source(a.a)
    pb, lb = resolve_source(a.b)
    log("Comparing %s with %s" % (la, lb))
    pd = _pd()
    A, B = load_norm(pa), load_norm(pb)
    print("\n  A = %s: %d records, %d columns" % (la, len(A), len(A.columns)))
    print("  B = %s: %d records, %d columns" % (lb, len(B), len(B.columns)))
    only_a = [c for c in A.columns if c not in B.columns]
    only_b = [c for c in B.columns if c not in A.columns]
    if only_a:
        print("  columns only in A: %s" % ", ".join(only_a[:12]))
    if only_b:
        print("  columns only in B: %s" % ", ".join(only_b[:12]))
    common = [c for c in A.columns if c in B.columns]
    keys = [k for k in S["key_columns"] if k in common]
    out = WORK / "compare"
    out.mkdir(parents=True, exist_ok=True)
    name = out / ("compare_%s_vs_%s.xlsx" % (re.sub(r"\W+", "", a.a or "live"), re.sub(r"\W+", "", a.b or "live")))
    sheets = {}
    if len(A) == len(B):
        diffs = {}
        for c in common:
            ne = A[c].values != B[c].values
            if ne.any():
                diffs[c] = ne
        print("\n  same number of records: cells that differ, by column")
        if not diffs:
            print("    none: the two versions hold the same data")
        rows = []
        for c, ne in sorted(diffs.items(), key=lambda kv: -kv[1].sum()):
            print("    %-28s %6d" % (c, ne.sum()))
            if a.col is None or a.col == c:
                idx = ne.nonzero()[0][: a.limit if a.col else 200]
                for i in idx:
                    rows.append({"Excel_row": i + 2, "Column": c, "A": A[c].values[i], "B": B[c].values[i],
                                 **{k: A[k].values[i] for k in keys}})
        if rows:
            sheets["Changed_cells"] = pd.DataFrame(rows)
    else:
        def keyed(D):
            k = D[keys].agg("|".join, axis=1)
            return k + "|#" + k.groupby(k).cumcount().astype(str)
        ka, kb = keyed(A), keyed(B)
        ra, rb = A[ka.isin(set(kb)) == False], B[kb.isin(set(ka)) == False]
        print("\n  records only in A: %d    only in B: %d    in both: %d" % (len(ra), len(rb), len(A) - len(ra)))
        sheets["Only_in_A"] = ra[keys + [c for c in ("PopPolio", "TotalNbrVaccPolio") if c in ra.columns]].head(5000)
        sheets["Only_in_B"] = rb[keys + [c for c in ("PopPolio", "TotalNbrVaccPolio") if c in rb.columns]].head(5000)
        for lab, D in (("A", ra), ("B", rb)):
            if len(D):
                print("    only in %s, by Country / Response:" % lab)
                print("      " + D.groupby(["Country", "Response"]).size().sort_values(ascending=False).head(8).to_string().replace("\n", "\n      "))
    if sheets:
        with pd.ExcelWriter(name) as w:
            for k, v in sheets.items():
                v.to_excel(w, sheet_name=k, index=False)
        print("\n  details: %s\n" % name)


# ----------------------------------------------------------------------------------------------- rules / knowledge
RULES = [
    # (id, area, what the system does, where it comes from / why, switch)
    ("coadmin", "Co-administration", "Same district, round and Response with an nOPV2 record and a bOPV record are merged into one 'nOPV2 & bOPV' record; TotalNbrVaccPolio = the larger of the two (same children). A combined record gets the same total in both columns until the country gives the split.",
     "Feedback from Valentina Bigoni, WHO AFRO", "(always)"),
    ("unused-vaccine", "Co-administration", "A vaccine that was not used has an EMPTY total and coverage (not 0). Real reported zeros are kept.", "Decision 30 Oct 2026 (Ahmed's convention)", "(always)"),
    ("poppolio", "Population", "PopPolio is corrected from vaccinated / CVPolio when coverage is impossible (above 1.5, or below 0.5 with a plausible CVPolio); decimals rounded; otherwise the district median of the year; then a gradient-boosting estimate (listed for confirmation).", "Quality work with the focal points", "--no-fix-poppolio, --no-model-poppolio"),
    ("copy-repair", "Copy / drag-fill errors", "A Response of another country, or nOPV3..nOPV15, is restored in this order: the LOOKUP (country + vaccine + start date), the newest earlier backup, the other rows of the same country/date/round, the campaign footprint. Only unique answers are used; the rest is UNRESOLVED.", "The lookup is the primary reference (your rule)", "--no-repair, --no-infer"),
    ("copied-rounds", "Copy / drag-fill errors", "Rounds that the lookup does not have and that only copy another round of the same Response and date are removed and listed.", "Found in the 2026 data", "--keep-copied-rounds"),
    ("response-fill", "Key fields", "An empty Response is filled from the lookup (exact start date, or within 21 days with the same round number), only when unique.", "Portal check 'Missing key field'", "--no-fill-response"),
    ("response-names", "Key fields", "A Response that is not a lookup name but starts the name of the one lookup round of that country within 7 days is renamed to the lookup's spelling (MLW-2026-02-1_nOPV -> MLW-2026-02-1_nOPV-NID).", "Decision 30 Oct 2026: follow the lookup", "--no-align-response-names, --align-response-names COUNTRY"),
    ("cvpolio", "Coverage", "An empty CVPolio is filled with TotalNbrVaccPolio / PopPolio. Records with only a target (no vaccinated) stay empty.", "Request of 28 Oct 2026", "--no-fill-cvpolio"),
    ("district-fill", "Districts", "An empty District is filled from a recovered file, matched on country + Response + round + province: same PopPolio and total; same PopPolio and total within 10%; same total; total + the target implied by the coverage; the last unmatched pair; same PopPolio or total with unique pairing.", "Ethiopia file AFRO-ETH FIVE REGIONS R1 (282 of 282 filled)", "--district-source FILE"),
    ("addis", "Districts", "Addis Ababa rows that repeat the sub-city name (NIFAS SILK LAFTO x13) are renamed '<SUB-CITY> WOREDA NN' from the file, only when every row of the sub-city matches a different woreda. Rows without a source (round 2 and 4 sub-cities of round 1) are left; they stay possible duplicates until a woreda-level file exists.", "Request of 29 Oct 2026", "(with --district-source)"),
    ("afrogdb-code", "Geography", "Two columns, geo_code and geo_match, are added at the end of the sheet: the AFROGDB district code of the Country / Province / District, found with the AFRO-Geo-Reference rules (exact name, same name ignoring accents and punctuation, the one district of that name in another province, or a reviewed crosswalk decision; then the extended layers of crosswalk_extended.csv: manual decisions, Ethiopia woreda to zone, relaxed, old-shape overlay, name similarity; geo_match says which, and the last two are marked probable). The names themselves are NOT changed. A name that is not found has an empty code and is listed in review/afrogdb_unmatched.csv of the build (names and counts only). The check warns when the share of records with a code falls.", "Request of 16 Nov 2026 (same reference as LQAS, IM and SCOPE)", "--no-afrogdb-codes, --no-afrogdb-extended, --afrogdb DIR"),
    ("province", "Geography", "A Province that belongs to another country (per geo_reference.csv) is corrected when certain.", "Admin portal checks", "--no-fix-province"),
    ("country-names", "Geography", "A Country spelled another way (CAMEROUN, REPUBLIQUE CENTRAFRIQUE, TANZANIA ...) gets the standard name of the geography reference (CAMEROON, CENTRAL AFRICAN REPUBLIC, UNITED REPUBLIC OF TANZANIA). Only names of a fixed list whose standard name is in the reference.", "Request of 4 Nov 2026", "--no-fix-country-names"),
    ("province-spelling", "Geography", "A Province that is the same name as an Admin1 of the geography reference written differently (spaces, accents, punctuation, filler words: MAI NDOMBE = MAINDOMBE, PROVINCIA DE NAMPULA = NAMPULA, WADIFIRA = WADI FIRA, RS 3 = RS3) gets the reference's spelling. Exact match only, and only when the district does not belong to another province of the reference. Near matches and new Admin1s (Cuando, Sidama ...) are never changed: they are listed in the fix list.", "Request of 4 Nov 2026", "--no-fix-province-spelling"),
    ("digit-error", "Population", "A target population that is 10 or 100 times off (a missing or extra digit) is corrected only when the reported coverage (vaccinated / CVPolio) implies the corrected figure within 10% AND that figure is 0.6 - 1.7 times the median target of the same district's other campaigns. Both must agree; otherwise it stays flagged.", "Request of 4 Nov 2026", "--no-fix-digit-error"),
    ("fix-list", "Checks", "What cannot be corrected safely (missing key fields, report dates, target above total population, coverage above the ceiling, very different targets, coverage that differs, provinces to confirm, combined and separate records, records entered twice) is written to Fix_list_by_country.xlsx: one sheet per problem, with the row, the column to correct, the current and the suggested value. Cut per country in the app.", "Request of 4 Nov 2026", "--no-fix-list"),
    ("submissions", "Intake", "The files the countries send (SharePoint folder Admin_data_submissions, copied read-only to workspace/submissions_in by 'admin.py submissions pull', which 'submissions inspect' runs first) are inspected one by one. The decisions (accept / reject, who, the note) are made in the Operations Center and kept in its AdminOps folder; a build on the PC reads that ledger first (read only, ADMIN_STORE_SHARE_URL) so that it adds exactly what the app accepted, and says so in the manifest and in the check when it could not: layout (template, template with other header spellings, mapped by certain synonyms, or not usable), normalisation of every cell (capitals, spaces, country / province / vaccine / round spelling, dates, numbers as text, percentages, Year / SIAMTH / Ctry_Code / AFRO_block, the constants the live file has), the portal quality checks, the campaigns against the lookup, duplicates inside the file, and every record against the live file (new / same / other figures). A file with no campaign name (a Nigeria eTally or tally) is wired to the lookup by its date: a campaign keeps its planned start date, so data dated on or after a planned start belong to that campaign until the next planned start (at most 60 days back), data of a campaign that ran up to 21 days before its planned start belong to that campaign, and anything else stays unwired for a person; Response, round and the planned date are written and the date in the file is kept in the note. When the file gives a round the lookup does not have (Cycle_2 in the file name, a round column, a hint) and the campaign that started last before the date has its last lookup round just below it, within 60 days, the records are the next round of the same campaign: same Response, that round, the date of the records; the lookup is then missing a round, which is listed in the report sheet Lookup_gaps, and a person confirms. An operator accepts or rejects each file; only accepted files are used, and only their NEW records are added at the end of the next build, before every other rule. Records already in the live file are never changed by a submission: the differences are listed.", "Request of 5 Nov 2026", "admin.py submissions inspect | list | accept | reject; build --no-submissions"),
    ("submissions-clean", "Intake", "What the intake no longer asks a person about (16 Nov 2026). NO LOSS (17 Nov 2026, automatic rule 4): every new record that passed the validation and the standardisation and is not a duplicate goes into the live file, also when its file has other findings; a record that fails (an Error of the quality checks, the target or the number vaccinated is not a number (text in a calculated column does not hold a record), a key field still empty, another country than the folder) is held back with its reason in the sheet Held_back_records and in prepared/<hash>_held.xlsx, as are records that look like live ones with other figures or match several live records. Files in the wrong folder or with a vaccine that disagrees with the lookup wait whole. NAME ALIGNMENT (17 Nov 2026): a record at the same place, within 3 days of the same date and with the same vaccine as exactly one live record is that live record under another campaign name (CAM-2025-11-01 for CAM-2026-01, NOPV2 for bOPV-nOPV2): it is not added, not held, and the pair is listed as an alias in the inspection. A DIFFERENT ROUND is accepted as the same record only when the file gives no real round (a date, empty) or when the lookup gives the live record's round for that campaign and date; otherwise the record is held for the country to confirm, never dropped (the live file may number a campaign's rounds differently from the lookup). FEEDBACK TO THE COUNTRIES: `python admin.py submissions feedback` writes Feedback_to_countries.md and .xlsx (one line per country, file and matter: target population missing, failed check, empty fields, text instead of numbers, unit not clear, same place and date with other figures, campaign name or round differs, vaccine differs from the lookup, sheet not read, wrong folder, calculated columns with errors), from the held-back records and the findings, in plain words with what the country does; nobody has to write it by hand. It also writes one workbook per country, ready to send, in workspace/submissions/feedback_by_country (sheet What to do with an Action needed column, sheet Records with the held-back rows and the reason; a country with nothing to do gets a file marked information_only); the francophone countries (list FRENCH_COUNTRIES in submissions.py) get the sheet À faire and a French read-me beside the English ones; every inspect writes them too, so one command (inspect, which also pulls the files from SharePoint) gives the report, the summary, the feedback and the country workbooks. NOT ADMIN DATA: a file with no admin columns whose name says IM, LQAS, AEFI, supplies, call-in or a summary is labelled NOT ADMIN DATA and ignored; a file with no template columns and no such name is NOT ignored (it is reported as another layout). (1) Coverage above the ceiling is for reference only: the live file has its own coverage columns, so the records are kept and the inspection only says how many there are. (2) Duplicates are removed: a record with the same place and the same figures as a live record under another campaign name, round or date is not added (the report counts it as removed); a copy of the same new records in two files with the same figures is kept once by the build. Only records with other figures than their twin are left for a person. (3) New records are standardised with the lookup: an empty or near-spelled Response gets the lookup name, an empty round gets the lookup round, and a named campaign whose records carry another date gets the lookup's PLANNED start date (the file's date goes in the notes); this uses the same date rules as the wiring of records without a campaign name, and the round the file gives is respected. Standardised dates are listed (code date_standardized) and the lookup-confirmed rule can accept them.", "Request of 16 Nov 2026", "none: always on in submissions inspect"),
    ("submissions-pending", "Intake", "PENDING ADMIN DATA (17 Nov 2026). A round of the lookup (GPEI) whose admin data are not in the live file 15 days after its planned start date is marked pending, so the data can be followed until they arrive. Every `submissions inspect` (and `submissions pending`, which only reads the live file and the lookup, about a minute) compares the live file with the lookup of the run and writes Pending_admin_data.xlsx in workspace/submissions: Pending_now (nothing received for the round, in no submitted file, and no live record of the country near its start date: ask the country), Date_differs_in_live (the live file has the campaign and round at another date than the lookup, within 35 days: a date to correct, not missing data; a file is matched to a lookup round by the lookup date of its records, never by the round number it writes), Check_other_name (nothing received but the live file has at least 5 records of the country within 10 days of the start: the round may be there under another campaign name or date; a person aligns the name before the country is asked), Received_not_in_live (a submitted file holds records of the round that are not in the live file yet: the next build, a decision or a correction is what is missing), Not_yet_due (started less than 15 days ago), Resolved (the data arrived), Closed, Older_rounds_no_data (rounds before pending_since, reference only), By_country. Each round has an age group (15-30, 31-90, over 90 days) and the date the engine first flagged it. pending_register.json (workspace/submissions) is the follow-up list: first flagged, resolved when the round is no longer missing, closed when a person says the round did not take place (`python admin.py submissions pending --close \"COUNTRY|campaign|Round n\" --note \"why\"`, then run `submissions pending` again). The rounds of Pending_now are also written into the feedback to the countries (matter: data not received, in English and, for the francophone countries, in French: send the admin data of the round in the AFRO template or tell us the round did not take place); a country with only pending rounds gets a workbook too. The workbooks of an earlier run move to feedback_by_country/_previous so that only the current ones are in the folder; a report that is open in Excel is written under the same name with the time added.", "Request of 17 Nov 2026", "settings pending_days (default 15) and pending_since (default 2025-01-01)"),
    ("submissions-lookup-issues", "Intake", "LOOKUP ISSUES (17 Nov 2026). Campaign rounds that are in the live file but that the lookup (GPEI) cannot place are listed in Pending_admin_data.xlsx, sheets Lookup_issues (rounds starting since pending_since) and Lookup_issues_older, with a message for the lookup maintainers and, when it helps, the country. Five kinds: the campaign is not in the lookup at all; the name differs from ONE lookup campaign starting within 7 days (a proposal, a person confirms); TWO or more lookup campaigns start within 7 days (e.g. DRC 28 Aug 2025: two campaigns, same date, same vaccines, so the data cannot say which one); the date is the start date of ANOTHER lookup campaign of the country; the lookup lists the campaign but not this round or date (e.g. Nigeria Round 2). A difference of up to 7 days is a date the build corrects, not an issue. SCOPE EVIDENCE: when the scope file is available (setting scope_file, or scope_.csv one level above scope_data_dir) each row also compares the DISTRICTS of the live records with the districts the scope plans for each campaign of the country within 60 days (geo_code, else the district name): the districts are those planned for ONE campaign (a strong proposal, never applied); the scope plans the same districts for two campaigns (the campaign is entered twice, e.g. DRC 28 Aug 2025: 518 identical districts); or, for a campaign the lookup knows, what the scope lists for it (rounds, districts). Fewer than 5 districts cannot tell campaigns apart. The engine never renames a campaign on its own because of these findings: the lookup is the reference, and a case is a decision for a person (a line in workspace/corrections.csv, a change in lookup_overrides.csv: `python admin.py lookup-fix adopt` writes the ones the scope evidence supports) or a fix at the source.", "Found 17 Nov 2026 (DRC and Nigeria rounds the lookup does not place); scope evidence added the same day. GPEI ROUND WINDOWS (18 Nov 2026): when scope_.csv carries the round end dates taken from the GPEI API (scripts/enrich_scope.py, run by the scope refresh), a live round whose date the window explains is NOT an issue: the date is inside the round window or its mop-up window, or it is within 35 days of the window and the live districts (5 or more, 90% or more) are those planned for that very round. These are listed in the sheet Lookup_accepted with the reason. Campaigns GPEI flags as test are left out of the scope evidence", "(always, with the pending admin data)"),
    ("sync-ops", "Intake", "COPY TO THE OPERATIONS CENTER (18 Nov 2026). `python admin.py sync-ops` copies corrections.csv, lookup_overrides.csv and scope_.csv to the folder inputs/ of the AdminOps folder (share link ADMIN_STORE_SHARE_URL), with the same app sign-in as the backup tool and the Operations Center, so the app wires the data the way this PC does. Only those three files are written. A file that has not changed since the last copy is skipped (workspace/sync_ops_state.json); before a file in the store is replaced its current content is saved in workspace/ops_backups; after the upload the size in the store is checked. --dry-run lists what would change, --force copies even if unchanged. If the app registration cannot write to the folder the command says so and nothing is copied.", "sync_ops.py, admin.py cmd_sync_ops", "(on demand)"),
    ("check-strict", "Check", "OPTIONAL STRICT CHECK (18 Nov 2026). The normal check (used by `all` and by the Operations Center) stops on the hard findings and only warns on the others. `python admin.py check --strict`, on demand only and never used by `all` or the app, turns the signs of damage into stops too: records of the base no longer in the build beyond the duplicates removed (same country, province, district, date and vaccine), columns the builder does not normally touch changed in more than 0.5% of the records, more than 1% of the records changed in the population or doses, more duplicates, more quality errors, fewer AFROGDB geo codes, or the submissions' decisions not read. Without --strict the same findings are listed as warnings (and the check always reports how many records of the base are no longer in the build). The chronic findings (a record that still has another country's Response) stay warnings: they are listed, not blocking. Nothing is published unless the check passed.", "admin.py check --strict", "--strict (on demand)"),
    ("publish-verified", "Publish", "VERIFIED PUBLISH (18 Nov 2026). After the publish the engine reads back the live file (the backup it takes of it) and compares it with the build: the same bytes, or at least the same content. The result is written in the manifest (published.verified) and in the log as VERIFIED or WARNING - NOT VERIFIED. A file that is not verified must be looked at before anyone uses it. The live file is also compared with the one the build started from just before the publish, so a change made by someone else in between stops the publish.", "admin.py publish", "always on"),
    ("lookup-overrides", "Build", "DECIDED CHANGES TO THE LOOKUP (17 Nov 2026). The lookup is the GPEI calendar and is refreshed from GPEI, so a change made in lookup.xlsx would be lost at the next refresh. A change a person decides is a line in workspace/lookup_overrides.csv (setting lookup_overrides_file; columns Action, Country, OBR Name, Round Number, Round Start Date, Vaccines, Round Status, New_Round_Start_Date, Reason, Decided_by, Decided_on) and is applied every time the lookup is read (the build, the check, the inspection, the pending data): `remove` drops a campaign entered twice, `add` adds a round the lookup lacks (not when it already has that round of that campaign or one within 7 days), `move` changes a start date. A line already made by GPEI counts as already in the lookup; a line that matches nothing is reported, never guessed. `python admin.py lookup-fix` (propose, nothing is changed) lists what the lookup issues and the scope evidence support: the live campaigns to rename to the lookup name (corrections lines) and the lookup changes; `lookup-fix adopt --by NAME [--choose \"LIVE NAME=LOOKUP NAME TO KEEP\"]` writes them as decided lines into corrections.csv and lookup_overrides.csv (old files kept as .bak copies; lines already there are not repeated); the next build applies both. Only rounds since pending_since are proposed (--older for the rest, not reviewed). Proposed only when the scope evidence is strong: the districts are those planned for ONE campaign, and at least 20 live districts or 50% of its planned districts have records; a campaign the scope plans twice needs --choose. AUTOMATIC TIER (17 Nov 2026): `lookup-fix auto`, run by the build itself before it starts while the setting lookup_fix_auto is on (default), adopts only what the engine can decide alone: strong scope evidence AND a rename that does not change how the vaccines are treated (the co-administration rule reads the NAME of the campaign: a name that holds nOPV and bOPV makes a single-vaccine record without a partner co-administered, so a rename is automatic when both names say the same about the vaccines, or when no unpaired single-vaccine record would be affected, e.g. the campaign keeps separate nOPV2 and bOPV records), and, for a campaign the lookup holds twice (same start date, round, vaccines, districts), keeps the entry whose planning month in its name is the start month and removes the others. The lines are written with Decided_by 'engine rule (lookup wiring ...)' and appear in the audit like any decided correction. What changes the vaccines (e.g. NIE-2026-01-SNID, 264 records) stays a proposal for a person. Checked on v35 (17 Nov 2026): 10 renames (2,515 records) and 2 lookup changes (remove DRC-2025-06-01_nOPV_NID Round 1 of 28 Aug 2025 as the duplicate of DRC-2025-08-01_n-bOPV_NID; add Round 2 of DRC-2025-11-01_n-bOPV_NID on 2025-12-11) bring the lookup issues since 2025 from 19 to 6, the rest being gaps at the source (Nigeria Round 2 of NIE-2025-08-01_n-bOPV_sNID and others).", "Request of 17 Nov 2026 (wire the DRC cases, update the lookup)", "setting lookup_overrides_file (default workspace/lookup_overrides.csv; no file, no changes)"),
    ("build-corrections", "Build", "DECIDED CORRECTIONS (17 Nov 2026). The engine does not change a live record on its own when a submitted file has other figures or another date for it; a person decides, and the decision is a line in workspace/corrections.csv (columns Country, Province, District, Response, Round_Add, SIA_date, Column, Found, Restored, Reason, Decided_by, Decided_on; one line per cell, the record is found by its identity and the date it has now). The build applies the file (--corrections, setting corrections_file) before the duplicate check and lists every change in the Repairs of the audit file (rule: Decided correction, with the reason and who decided). A cell is changed only while it still holds the value in Found: a line already applied is counted as already in place (also when the record's date was corrected by the same file), and a record or cell changed since is skipped and reported, never overwritten. `python admin.py corrections` shows what the file would change in the live file (applied, already in place, skipped with the reason) and changes nothing. A line with District empty is campaign-wide: it changes every record of Country + Response + Round_Add (optionally one SIA_date) and needs Found, which guards each record; it is how a campaign gets the name the lookup has. Uses so far (17 Nov 2026): Zimbabwe Round 2 of ZIM-2026-03-01_nOPV_sNID (27 districts), dated 2026-06-02 as the lookup says instead of 2026-07-02, with the vaccinated figures of the file Zimbabwe R1 and R2 adminstrative coverage 2026 (vaccinated, nOPV2 total, coverage recalculated; the target population of the live record is kept); Zimbabwe Round 1 figures of the same file; Zimbabwe Round 3 renamed from the placeholder POLIO to the lookup campaign; five DRC campaigns renamed to the lookup name (29 Jan, 26 Feb, 16 Apr 2026, June 2025 Round 2, June 2026 Round 2; 1,282 records). Two DRC cases are NOT renamed: 28 Aug 2025 (the lookup has two campaigns on that date) and 3 Sept 2026 (the live name matches the other lookup campaign of that name).", "Request of 17 Nov 2026", "setting corrections_file (default workspace/corrections.csv; no file, no corrections)"),
    ("submissions-auto", "Intake", "Automatic rule 1 (switch: submissions_auto in admin_settings.json, off by default). A file nobody has decided on is taken into the build by itself only when every finding of its inspection is harmless: records already in the live file with other figures (never changed by a submission), a header mapping that was certain, records left out, many warnings. A file with an Error, a coverage above the ceiling, a date far from its campaign, a campaign guessed from the live file or the file name, values never seen, cells that are not numbers, a possible duplicate, records also new in another file or an eTally conversion always waits for a person. A decision of a person, accepted or rejected, always wins. Rule 3 (the engine learns from the people): a file with the same name pattern, country folder and layout as a file a person accepted is taken again when it shows no kind of finding that one did not show and none of the hard ones (an Error, several live units, a possible duplicate, an overlap, another country, cells that are not numbers, a vaccine that differs from the lookup, an unreadable sheet). admin.py submissions summary lists, by country, where the new records are, what holds them back and who can fix it, and which lookup rounds with no data are in a submitted file. Rule 2 (needs an inspection made by this version): when every campaign round of the file is a round of the lookup (same country, campaign, date and round) and the vaccine agrees, a campaign taken from the lookup or the file name, a year typed wrongly and an eTally export added up by LGA are harmless too. The files taken are listed in the log, in the manifest and in the check.", "Request of 16 Nov 2026", "--no-auto-submissions; admin.py submissions auto lists what the rule would take"),
    ("text-clean", "Hygiene", "Stray control characters (carriage returns), line breaks and double / edge spaces in Country, Province, District, Response, Vaccine_type and Round_Add are removed.", "Found in District (KAJIADO EAST) 30 Oct 2026", "--no-clean-text"),
    ("dates", "Hygiene", "SIA_date written as text ('18-sept-20') becomes a real date.", "Admin portal", "(always)"),
    ("duplicates", "Hygiene", "Exact copies are removed. A record entered twice (same key AND the same population and results) is reported as Duplicate record and removed only with --resolve-duplicates. Records with the same key but other population or results are different units (most records have no sub-district): never removed, reported only as information (Same names, other figures). Key = Country, Province, District, [Admin_3], SIA_date, Round_Add, Vaccine_type, Response, compared without accents, case or punctuation.", "Admin portal", "--no-dedupe, --resolve-duplicates; admin.py duplicates lists them and what a fix would do; the Operations Center has the same choice on the Cycle tab"),
    ("quality", "Checks", "The portal quality checks (missing key field, coverage above the ceiling, target jumps, report date, province/country, co-administration consistency, vaccine spelling, duplicates) produce a report with Errors / Warnings / Info.", "Admin portal (run_admin_quality_check.R)", "--no-quality"),
]
DECISIONS = [
    "The lookup file is the primary reference for campaign names, dates and rounds; backups come second.",
    "Never guess: an answer is used only when it is unique; otherwise the record is left and listed.",
    "A build always starts from the live SharePoint file, so nobody's newer edits are lost.",
    "SharePoint keeps no version history (the local backup has it); one published file, replaced in place.",
    "The data dictionary (AFRO_admin_data_dictionary.xlsx) is pushed with a publish only when its content changed (compared cell by cell); then it replaces the existing SharePoint file, and the previous live one is kept in workspace/dictionary_backups. A dictionary problem never undoes the published data (exit 3, retry with `admin.py dictionary`).",
    "The Operations Center app keeps only the newest 3 backups of the live file on SharePoint (the PC keeps all of them); the dictionary is pushed only when it changed.",
    "Data is never committed to git (only lookup.xlsx and geo_reference.csv). No attribution lines in commit messages.",
    "Each code change is its own commit, followed by its own push.",
    "Open questions kept visible: Addis Ababa round 2 and 4 sub-cities lack woreda names (need a woreda-level file); Tegedie exists in two zones of Amhara with the same name; about 1,460 Nigeria 2022-Jan 2023 records have no Response in the lookup.",
]


def rules_markdown():
    L = ["# What the admin data system knows", "", "Generated by `python admin.py rules --markdown`. The rules live in `add_coadmin_columns.py`; this is their catalogue.", ""]
    area = None
    for rid, ar, what, src, sw in RULES:
        if ar != area:
            L += ["", "## " + ar, ""]
            area = ar
        L += ["- **%s**: %s  \n  Source: %s. Switch: `%s`." % (rid, what, src, sw)]
    L += ["", "## Decisions and standing rules", ""] + ["- " + d for d in DECISIONS] + [""]
    return "\n".join(L)


def cmd_rules(a):
    if a.markdown:
        out = HERE / "KNOWLEDGE.md"
        out.write_text(rules_markdown(), encoding="utf-8")
        print("written: %s" % out)
        return
    print()
    for rid, ar, what, src, sw in RULES:
        if a.name and a.name.lower() not in (rid + " " + ar + " " + what).lower():
            continue
        print("[%s] %s\n    %s\n    source: %s | switch: %s\n" % (rid, ar, what, src, sw))
    if not a.name:
        print("Decisions and standing rules:")
        for d in DECISIONS:
            print("  - " + d)
        print()


# ----------------------------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd")
    s = sub.add_parser("status"); s.add_argument("--live", action="store_true", help="read SharePoint first (runs the backup)")
    sub.add_parser("runs")
    sub.add_parser("backup")
    s = sub.add_parser("lookup", help="refresh lookup.xlsx from the GPEI API"); s.add_argument("--dry-run", action="store_true")
    s = sub.add_parser("build"); s.add_argument("--base", help="start from this file instead of the latest backup of the live file")
    s.add_argument("--no-auto-wire", action="store_true", help="do not let the engine wire the lookup cases it can decide alone (setting lookup_fix_auto) before this build")
    s.add_argument("--no-refresh-lookup", action="store_true", help="use lookup.xlsx as it is (default: refreshed from the GPEI API first)")
    s.add_argument("--no-backup", action="store_true", help="do not read SharePoint, use the newest local backup")
    s.add_argument("--no-submissions", action="store_true", help="do not add the accepted submissions to this build")
    s.add_argument("--no-auto-submissions", action="store_true", help="do not take the files of automatic rule 1 (setting submissions_auto)")
    s.add_argument("--no-ledger-pull", action="store_true", help="do not read the decisions of the Operations Center first, use the ledger kept on this PC")
    s.add_argument("extra", nargs="*", help="extra options for the builder, after --")
    s = sub.add_parser("baseinfo", help=argparse.SUPPRESS); s.add_argument("path")
    s = sub.add_parser("check"); s.add_argument("--run", type=int, help="version number (default: the newest)")
    s.add_argument("--strict", action="store_true", help="on demand: also stop (not only warn) on the signs of damage: records of the base lost, columns the builder does not touch changed, "
                   "more duplicates or quality errors, fewer geo codes, the submissions' decisions not read. Not used by 'all'")
    s = sub.add_parser("excel"); s.add_argument("--verify-only", action="store_true", help="you saved the file in Excel yourself")
    s.add_argument("--use-excel", action="store_true", help="let desktop Excel store the values (default: written directly, no Excel)")
    s = sub.add_parser("publish"); s.add_argument("--ask", action="store_true", help="ask for YES before replacing the file")
    s.add_argument("extra", nargs="*", help="extra options for the publish script, after --")
    s = sub.add_parser("dictionary", help="push only the data dictionary of the last published build")
    s.add_argument("--dry-run", action="store_true"); s.add_argument("extra", nargs="*")
    s = sub.add_parser("all"); s.add_argument("extra", nargs="*")
    s = sub.add_parser("commit"); s.add_argument("-m", "--message", required=True); s.add_argument("--no-push", action="store_true")
    s.add_argument("files", nargs="*")
    sub.add_parser("sync-scope")
    so = sub.add_parser("sync-ops"); so.add_argument("--dry-run", action="store_true"); so.add_argument("--force", action="store_true")
    sub.add_parser("doctor")
    s = sub.add_parser("inspect"); s.add_argument("--source", help="live (default), vNN build, bN backup version, or a file")
    s.add_argument("--no-refresh-lookup", action="store_true")
    s.add_argument("--country")
    s = sub.add_parser("submissions", help="inspect the files the countries sent, accept or reject them (accepted ones go into the next build)")
    s.add_argument("action", choices=["pull", "inspect", "list", "auto", "summary", "feedback", "pending", "accept", "reject"])
    s.add_argument("targets", nargs="*", help="accept / reject: the file name or the start of its content hash")
    s.add_argument("--folder", help="inspect: the folder with the files (default: the submissions folder of the workspace, filled from SharePoint)")
    s.add_argument("--no-pull", action="store_true", help="inspect: do not read the SharePoint folder first, use the local copy as it is; list: do not read the app's ledger")
    s.add_argument("--local", action="store_true", help="accept / reject: decide on this PC only (normally decisions are made in the Operations Center)")
    s.add_argument("--dry-run", action="store_true", help="pull: list what would be copied, write nothing")
    s.add_argument("--note", help="accept / reject: a note kept in the ledger"); s.add_argument("--by", help="who decides (default: the user name)")
    s.add_argument("--no-refresh-lookup", action="store_true")
    s.add_argument("--close", action="store_true", help="pending: mark the given rounds ('COUNTRY|campaign|Round n') as closed: the round did not take place or no data are expected")
    s = sub.add_parser("lookup-fix", help="from the lookup issues and the scope to decided changes: propose (default, nothing is changed) or adopt (writes corrections.csv and lookup_overrides.csv)")
    s.add_argument("action", nargs="?", choices=["propose", "adopt", "auto"], default="propose",
                   help="auto = adopt only what the engine can decide alone (strong scope evidence, vaccine treatment unchanged); the build does it itself when lookup_fix_auto is on")
    s.add_argument("--by", help="adopt: who decides (written in the decided lines)")
    s.add_argument("--older", action="store_true", help="also propose for rounds before pending_since (not reviewed: check each line)")
    s.add_argument("--choose", action="append", help="a campaign the scope plans twice: LIVE NAME=LOOKUP NAME TO KEEP (the other entry is removed from the lookup); repeatable")
    sub.add_parser("corrections", help="show what the corrections file (decided corrections) would change in the live file; nothing is changed")
    s = sub.add_parser("duplicates", help="list the duplicates of a file and what removing them would do (nothing is changed)")
    s.add_argument("--source", help="live (default), vNN build, bN backup version, or a file"); s.add_argument("--country")
    s.add_argument("--keep", choices=["complete", "first"], default="complete"); s.add_argument("--no-file", action="store_true")
    s = sub.add_parser("find"); s.add_argument("--source"); s.add_argument("--country"); s.add_argument("--province")
    s.add_argument("--district"); s.add_argument("--response"); s.add_argument("--vaccine"); s.add_argument("--round")
    s.add_argument("--date", help="YYYY-MM-DD or YYYY-MM"); s.add_argument("--limit", type=int, default=40)
    s.add_argument("--summary", action="store_true", help="count per Response / round / date instead of listing records")
    s.add_argument("--cols", nargs="*", help="more columns to show")
    s = sub.add_parser("compare"); s.add_argument("a", help="live, vNN, bN or a file"); s.add_argument("b")
    s.add_argument("--col", help="list the changed cells of this column"); s.add_argument("--limit", type=int, default=200)
    s = sub.add_parser("rules"); s.add_argument("name", nargs="?"); s.add_argument("--markdown", action="store_true")
    a = ap.parse_args()
    if not a.cmd:
        ap.print_help()
        return
    fn = {"status": cmd_status, "runs": cmd_runs, "backup": cmd_backup, "lookup": cmd_lookup, "build": cmd_build, "check": cmd_check,
          "excel": cmd_excel, "publish": cmd_publish, "dictionary": cmd_dictionary, "all": cmd_all, "commit": cmd_commit,
          "sync-scope": cmd_sync_scope, "sync-ops": cmd_sync_ops, "doctor": cmd_doctor, "inspect": cmd_inspect, "find": cmd_find,
          "compare": cmd_compare, "rules": cmd_rules, "baseinfo": cmd_baseinfo, "duplicates": cmd_duplicates,
          "submissions": cmd_submissions, "corrections": cmd_corrections, "lookup-fix": cmd_lookup_fix}[a.cmd]
    log("admin.py %s" % " ".join(sys.argv[1:]), quiet=True)
    try:
        fn(a)
    except KeyboardInterrupt:
        die("interrupted.", 130)


if __name__ == "__main__":
    main()
