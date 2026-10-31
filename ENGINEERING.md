# Admin data: one system, one command

Everything for `AFRO_admin_data.xlsx` is run from **this folder** with one launcher:

```
admin.bat <step>            (or: python admin.py <step>)
```

Nothing else needs to be run from other directories. The only other folder involved is the SharePoint backup tool
(`SharePoint_version_backup`), which `admin.py` calls for you (`backup_repo` in `admin_settings.json`).

## The cycle

| Step | What it does | Stops if |
|---|---|---|
| `status` | Where everything stands: latest backup, builds, what to do next, git, leftover folders. `status --live` reads SharePoint first. | |
| `backup` | Saves the live SharePoint file locally (every real version is kept) and lists the versions. | SharePoint cannot be read |
| `build` | Backs up the live file, then builds the next version **starting from that live file**, so everyone's latest edits are kept. Writes `workspace/runs/vNN/`. | the builder fails |
| `check` | Compares the build with the file it started from: columns, rows, empty key fields, Response of another country, duplicates, errors, cells changed. Writes `check_report.md`. | columns lost, > 2% rows lost, more empty key fields, more wrong-country Responses |
| `excel` | Stores the `check_dp` values in the build (the formulas stay): written straight into the file, no Excel needed, about 15 seconds. `excel --use-excel` lets desktop Excel do it instead (falls back to the direct method if Excel fails, for example when Office is not activated). `excel --verify-only` if you saved it yourself. | `check_dp` still empty |
| `publish` | Re-reads SharePoint, confirms it is still the file the build started from, then runs the publish script (replaces the file, deletes the SharePoint versions, backs up again) and finally **replaces the data dictionary** on SharePoint with the one built with the run, but only when its content changed (compared cell by cell; if it did not change nothing is pushed) (`AFRO_admin_data_dictionary_vNN.xlsx`): same file, same link; the previous live dictionary is first copied to `workspace/dictionary_backups/`. If only the dictionary fails (locked, link) the data stays published, the manifest says `dictionary: failed` and `admin.bat dictionary` repeats just that step. `publish -- --no-dictionary` skips it. No question is asked: when every check passes it goes ahead (`publish --ask` to be asked). | not checked, not saved in Excel, **the live file changed since the build** |
| `dictionary` | Pushes only the data dictionary of the last published build, if it changed, into the existing SharePoint `AFRO_admin_data_dictionary.xlsx` (retry after a failure, or refresh the dictionary alone; `dictionary --dry-run` checks without writing). | the dictionary file is missing, the file is locked/changed on SharePoint |
| `all` | backup -> build -> check -> excel -> publish | any step above |
| `runs` | One line per build | |
| `inspect` | Inspects any version (`--source live`, `vNN`, `bN` or a file; `--country X`): the portal quality checks, hygiene (stray characters, empty key fields, a Response that belongs to another country) and the campaign rounds against the lookup. Reports go to `workspace/inspect/`. | |
| `find` | Looks up records in any version: `find --country cameroon --date 2026-04 --summary`, `find --district kajiado`. Replaces ad-hoc scripts run from other folders. | |
| `compare A B` | What differs between two versions (`compare b6 b7`: what Ahmed changed; `compare live v24`: what the build changes). Counts per column, details in `workspace/compare/`. | |
| `rules` | What the system knows and fixes, with the reason and the switch of each rule, and the standing decisions (`rules coadmin`, `rules --markdown` writes KNOWLEDGE.md). | |
| `commit -m "msg" FILE...` | `git add` + `commit` + `push` of the named files. Refuses data files and secrets, and refuses Co-Authored-By / Claude-Session lines. | |
| `sync-scope` | Copies the scripts and the newest build to the scope folders (paths in settings) | |
| `doctor` | Checks python packages, files, git, Excel | |

A normal update is:

```
admin.bat status --live
admin.bat build
admin.bat check
admin.bat excel
admin.bat publish
```

or simply `admin.bat all`.

## Where things are

```
Admin_data_coadministration/
  admin.py, admin.bat, admin_settings.json     the system and its paths (no secrets)
  add_coadmin_columns.py                       the builder (all the data rules)
  publish_admin_data.py, publish_admin_data.bat
  lookup.xlsx, geo_reference.csv               reference files (the only data kept in git)
  inputs/                                      recovered / source files (e.g. the Ethiopia file); not in git
  workspace/runs/vNN/                          one folder per build; not in git
      AFRO_admin_data_vNN.xlsx                 the file that gets published
      AFRO_admin_data_dictionary_vNN.xlsx, ..._coadmin_audit_vNN.xlsx, ..._quality_report_vNN.xlsx,
      Feedback_suspicious_coverage_vNN.xlsx
      build.log, check_report.md, publish.log
      manifest.json                            what it started from, what was done, what is left
  workspace/inspect/, workspace/compare/       reports of inspect and compare
  workspace/dictionary_backups/                previous live dictionaries (kept before each replacement)
  workspace/state.json                         the last published version
  KNOWLEDGE.md                                 the rules and decisions (generated by: admin.bat rules --markdown)
  logs/admin_YYYYMMDD.log                      every command, with the time
```

## Safety rules built in

1. **A build always starts from the live file** (the newest local backup after a fresh backup), never from an old
   copy. This is what keeps Ahmed's (or anyone's) newer edits.
2. **Publish refuses** if the live file changed after the build: it compares the content (ignoring how Excel stored
   it), so opening the file in a browser does not block you, a real edit does. Rebuild and you are safe again.
3. **Nothing is published** unless the check passed and the `check_dp` values are stored.
4. **Every step is logged** and every build keeps its manifest: you can see months later which file a version came from.
5. **Data never goes to git**: `commit` refuses spreadsheets (except `lookup.xlsx`, `geo_reference.csv`) and secrets.
6. The SharePoint credentials stay in the backup tool's `config/secrets.env`; `admin.py` never reads them.

## Settings (`admin_settings.json`)

| Setting | Meaning |
|---|---|
| `backup_repo` | folder of the SharePoint backup tool |
| `workspace` | where builds are kept |
| `version_floor` | the next build is the highest of this and the existing builds, plus 1 |
| `district_source` | optional recovered file for empty District values (used if the file exists) |
| `build_extra_args` | extra options always passed to the builder |
| `max_row_drop_pct` | the check stops if more rows than this disappear |
| `scope_script_dir`, `scope_data_dir` | where `sync-scope` copies to |

## When something stops

Each stop says what happened and what to do. The common ones:

- *The live SharePoint file changed after the build*: someone saved a new version; run `build` again.
- *check FAILED*: read `workspace/runs/vNN/check_report.md`; fix the cause (a rule in the builder, a reference file) and `build` again.
- *Excel could not save*: open the file yourself, save once, run `excel --verify-only`.
- Do not open the file in Excel Online: it autosaves and creates a SharePoint version.

## Cross-checking without leaving this folder

| You want to know | Command |
|---|---|
| Is the live file what I published? | `admin.bat compare live v24` (no differences = yes) |
| What did someone change on SharePoint? | `admin.bat compare b6 b7` (backup versions, see `backup`) |
| Where is this campaign / district in the data? | `admin.bat find --response CAM-2026 --summary` |
| Is a country clean? | `admin.bat inspect --country CAMEROON` |
| What does the build change? | `admin.bat check` (report) or `admin.bat compare live v24` |
| Why does the system do X? | `admin.bat rules X` |

## Adding a new fix

1. Put the rule in `add_coadmin_columns.py` (the builder), with a switch to turn it off, and log it in the audit sheet `Repairs`.
2. Add one line to `RULES` in `admin.py` (what, why, switch) and run `admin.bat rules --markdown`.
3. If it can go wrong silently, add the matching test to `check` (for example "no Response of another country").
4. `build`, `check`, `compare live vNN`: look at what changed before publishing.
5. Commit the code: `admin.bat commit -m "message" add_coadmin_columns.py admin.py KNOWLEDGE.md`.
