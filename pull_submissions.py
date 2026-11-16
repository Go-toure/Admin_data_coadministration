#!/usr/bin/env python3
"""Copy the files the countries sent from the SharePoint folder Admin_data_submissions to a local folder, so that
'admin.py submissions inspect' can read them without anyone downloading them by hand.

    python pull_submissions.py --dest workspace/submissions_in            # the share link comes from ADMIN_SUBMISSIONS_SHARE_URL
    python pull_submissions.py --dest PATH --share-url LINK --dry-run     # list what would be copied, write nothing

The SharePoint folder is only READ, never changed. The link and the sign-in values (SHAREPOINT_TENANT_ID / _CLIENT_ID /
_CLIENT_SECRET) are taken from the environment, else from the backup tool's config/secrets.env (--env-file, or
ADMIN_SUBMISSIONS_ENV_FILE). Nothing secret is written anywhere by this script.

What is copied: .xlsx .xlsm .xls .xlsb .csv .zip (zip archives are opened by the engine), also from sub-folders; not Office
lock files (~$...) and not files larger than --max-mb (default 60). A file is copied again only when its size or its
modified time on SharePoint changed (the signatures are kept in <dest>/.sources.json). A file that is gone from the SharePoint
folder is removed from the local copy, so the local folder always shows what the folder shows; if the folder looks empty
while the local copy is not, nothing is removed (a wrong link must not wipe the copy).

A second mode, --ledger, reads the DECISIONS instead of the files: the folder submissions/ of the AdminOps folder of the Operations
Center (link in ADMIN_STORE_SHARE_URL) holds the ledger (who accepted or rejected which file), the report and the prepared record
files. They are copied (read only) to --dest, so that a build on the PC adds exactly the files accepted in the app. The ledger
in place is first kept in <dest>/ledger_backups (newest 5). Prepared files that the app removed are removed here too.

Used by admin.py (submissions pull, submissions inspect, build). Exit code: 0 done, 1 failed (the local copy stays as it was)."""
import argparse
import base64
import json
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
GRAPH = os.environ.get("PUBLISH_GRAPH_BASE", "https://graph.microsoft.com/v1.0")
TOKEN_URL = os.environ.get("PUBLISH_TOKEN_URL", "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token")
EXT = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv", ".zip")
META = ".sources.json"
LEDGER_META = ".sources_store.json"
KEEP_LEDGER_BACKUPS = 5


class Failed(Exception):
    pass


def say(msg, emit=None):
    (emit or (lambda m: print(m, flush=True)))(msg)


def read_env(env_file=None, keys=()):
    """The values of keys from the environment, else from the first secrets.env that has them. Returns a dict."""
    out = {k: os.environ.get(k, "").strip() for k in keys}
    files = [env_file, os.environ.get("ADMIN_SUBMISSIONS_ENV_FILE"), HERE.parent / "SharePoint_version_backup" / "config" / "secrets.env"]
    for f in files:
        if not f or not Path(f).is_file() or all(out.values()):
            continue
        for line in Path(f).read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                k = k.strip()
                if k in out and not out[k]:
                    out[k] = v.strip().strip("\"'")
    return out


class GraphApi:
    """Read-only Microsoft Graph access (app-only sign-in, the same values as the backup tool)."""

    def __init__(self, tenant, client, secret):
        import requests
        self.r = requests
        try:
            resp = requests.post(TOKEN_URL.format(tenant=tenant), timeout=30, data={
                "client_id": client, "client_secret": secret, "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials"})
        except requests.exceptions.RequestException as e:
            raise Failed("no connection to SharePoint (%s)." % type(e).__name__)
        if resp.status_code != 200:
            raise Failed("could not sign in to SharePoint (HTTP %s): check the three SHAREPOINT_* values." % resp.status_code)
        self.h = {"Authorization": "Bearer " + resp.json()["access_token"]}

    def _get(self, url):
        import time
        last = None
        for attempt in range(4):
            try:
                resp = self.r.get(url, headers=self.h, timeout=120)
            except self.r.exceptions.RequestException as e:
                last = e
                time.sleep(2 * (attempt + 1))
                continue
            if resp.status_code in (429, 502, 503, 504):
                time.sleep(max(int(resp.headers.get("Retry-After", 0) or 0), 4 * (attempt + 1)))
                last = resp
                continue
            return resp
        raise Failed("network error talking to SharePoint (%s)." % (type(last).__name__ if last is not None else "unknown"))

    def root(self, share_url):
        b64 = base64.b64encode(share_url.encode()).decode().rstrip("=").replace("/", "_").replace("+", "-")
        resp = self._get("%s/shares/u!%s/driveItem" % (GRAPH, b64))
        if resp.status_code != 200:
            raise Failed("cannot open the share link (HTTP %s): is it the link of the folder, and can the app read it?" % resp.status_code)
        it = resp.json()
        if "folder" not in it:
            raise Failed("the share link points to a file, not to a folder.")
        return {"drive": it["parentReference"]["driveId"], "id": it["id"], "name": it.get("name", "")}

    def child(self, drive, parent_id, name):
        from urllib.parse import quote
        resp = self._get("%s/drives/%s/items/%s:/%s" % (GRAPH, drive, parent_id, quote(name)))
        if resp.status_code == 404:
            return None
        if resp.status_code != 200:
            raise Failed("cannot open the folder %s (HTTP %s)." % (name, resp.status_code))
        it = resp.json()
        return it["id"] if "folder" in it else None

    def children(self, drive, item_id):
        out, url = [], "%s/drives/%s/items/%s/children?$top=200" % (GRAPH, drive, item_id)
        while url:
            resp = self._get(url)
            if resp.status_code != 200:
                raise Failed("cannot list a folder (HTTP %s)." % resp.status_code)
            b = resp.json()
            for c in b.get("value", []):
                out.append({"name": c["name"], "size": int(c.get("size") or 0), "is_dir": "folder" in c, "id": c["id"],
                            "modified": c.get("lastModifiedDateTime", ""), "url": c.get("@microsoft.graph.downloadUrl")})
            url = b.get("@odata.nextLink")
        return out

    def download(self, drive, item):
        if item.get("url"):
            resp = self.r.get(item["url"], timeout=300)
        else:
            resp = self._get("%s/drives/%s/items/%s/content" % (GRAPH, drive, item["id"]))
        if resp.status_code != 200:
            raise Failed("cannot read %s (HTTP %s)." % (item["name"], resp.status_code))
        return resp.content


def pull(api, share_url, dest, max_mb=60, dry_run=False, emit=None, subpath=None, exts=EXT, prune_prefix=None,
         meta_name=META, backup_dir=None, label="Submissions folder"):
    """Mirror the files of a SharePoint folder (or of its sub-folder subpath) into dest. exts=None takes every file.
    Only local files under prune_prefix (None: all) are removed when they are gone from SharePoint.
    api has root(url), child(drive, id, name), children(drive, id), download(drive, item)."""
    dest = Path(dest)
    root = api.root(share_url)
    drive = root["drive"]
    start = root["id"]
    if subpath:
        start = api.child(drive, root["id"], subpath)
        if start is None:
            say("%s: no folder '%s' in '%s' yet: nothing to read." % (label, subpath, root["name"]), emit)
            return {"files": 0, "copied": 0, "removed": 0, "skipped": [], "dry_run": dry_run, "missing": True}
    metap = dest / meta_name
    try:
        meta = json.loads(metap.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        meta = {}
    seen, got, skipped, other, todo = {}, 0, [], {}, []

    def walk(item_id, rel):
        for c in api.children(drive, item_id):
            r = (rel + "/" + c["name"]) if rel else c["name"]
            if c["is_dir"]:
                walk(c["id"], r)
                continue
            if c["name"].startswith(("~$", ".")):
                continue
            if exts is not None and not c["name"].lower().endswith(exts):
                ext = Path(c["name"].lower()).suffix or "(no extension)"
                other[ext] = other.get(ext, 0) + 1
                continue
            if c["size"] > max_mb * 1048576:
                skipped.append("%s (%.0f MB)" % (r, c["size"] / 1048576))
                continue
            seen[r] = [c["size"], c["modified"]]
            # the ledger is small and can have been changed here by a local 'inspect': it is always compared with the one in the app
            always = bool(backup_dir) and Path(r).name == "ledger.json"
            if always or not ((dest / r).is_file() and meta.get(r) == seen[r]):
                todo.append((r, c))
    walk(start, "")
    local = [p.relative_to(dest).as_posix() for p in dest.rglob("*")
             if p.is_file() and not p.name.startswith(".sources") and not p.name.endswith(".part")] if dest.is_dir() else []
    gone = [r for r in local if r not in seen and (prune_prefix is None or r.startswith(prune_prefix))]
    if not seen and gone:
        say("WARNING: the SharePoint folder shows no data file but the local copy has %d: nothing was changed. "
            "Check the share link." % len(gone), emit)
        return {"files": 0, "copied": 0, "removed": 0, "skipped": skipped, "dry_run": dry_run, "unchanged_empty_listing": True}
    if dry_run:
        for r, c in todo:
            say("   would copy: %s" % r, emit)
        for r in gone:
            say("   would remove from the local copy: %s" % r, emit)
    else:
        dest.mkdir(parents=True, exist_ok=True)
        for r, c in todo:
            target = dest / r
            target.parent.mkdir(parents=True, exist_ok=True)
            data = api.download(drive, c)
            if target.is_file() and target.read_bytes() == data:
                meta[r] = seen[r]
                continue                                  # the same content: nothing to write
            if backup_dir and target.name == "ledger.json" and target.is_file() and target.read_bytes() != data:
                keep_backup(target, backup_dir)
            tmp = target.with_name(target.name + ".part")
            tmp.write_bytes(data)
            os.replace(tmp, target)
            meta[r] = seen[r]
            got += 1
            say("   read from the submissions folder: %s" % r, emit)
        for r in gone:
            try:
                (dest / r).unlink()
            except OSError:
                pass
            meta.pop(r, None)
        meta = {k: v for k, v in meta.items() if k in seen}
        metap.write_text(json.dumps(meta), encoding="utf-8")
    say("%s '%s': %d file(s) in the folder, %d %s, %d %s." % (
        label, root["name"], len(seen), len(todo) if dry_run else got, "to copy" if dry_run else "copied now",
        len(gone), "to remove from the local copy" if dry_run else "removed from the local copy"), emit)
    if other:
        say("   not data files, not read: %s." % ", ".join("%d %s" % (n, e) for e, n in sorted(other.items())), emit)
    for x in skipped:
        say("   NOT READ, larger than %d MB: %s" % (max_mb, x), emit)
    return {"files": len(seen), "copied": got, "removed": len(gone), "skipped": skipped, "dry_run": dry_run}


def keep_backup(path, backup_dir):
    import shutil
    from datetime import datetime
    Path(backup_dir).mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, Path(backup_dir) / ("ledger_%s.json" % datetime.now().strftime("%Y%m%d_%H%M%S")))
    old = sorted(Path(backup_dir).glob("ledger_*.json"))
    for f in old[:-KEEP_LEDGER_BACKUPS]:
        try:
            f.unlink()
        except OSError:
            pass


def run(dest, share_url=None, env_file=None, max_mb=60, dry_run=False, emit=None, api=None, ledger=False):
    link_key = "ADMIN_STORE_SHARE_URL" if ledger else "ADMIN_SUBMISSIONS_SHARE_URL"
    keys = (link_key, "SHAREPOINT_TENANT_ID", "SHAREPOINT_CLIENT_ID", "SHAREPOINT_CLIENT_SECRET")
    env = read_env(env_file, keys)
    url = (share_url or env[link_key]).strip()
    if not url:
        raise Failed("no share link: set %s in the backup tool's config/secrets.env (or give --share-url)." % link_key)
    if api is None:
        if not all(env[k] for k in keys[1:]):
            raise Failed("SHAREPOINT_TENANT_ID / _CLIENT_ID / _CLIENT_SECRET are not set (the backup tool's config/secrets.env).")
        api = GraphApi(env["SHAREPOINT_TENANT_ID"], env["SHAREPOINT_CLIENT_ID"], env["SHAREPOINT_CLIENT_SECRET"])
    if ledger:
        return pull(api, url, dest, max_mb=max_mb, dry_run=dry_run, emit=emit, subpath="submissions", exts=None,
                    prune_prefix="prepared/", meta_name=LEDGER_META, backup_dir=Path(dest) / "ledger_backups",
                    label="AdminOps folder (decisions on the submissions)")
    return pull(api, url, dest, max_mb=max_mb, dry_run=dry_run, emit=emit)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--dest", required=True, help="the local folder that receives the files")
    ap.add_argument("--share-url", help="the share link of the folder (default: ADMIN_SUBMISSIONS_SHARE_URL)")
    ap.add_argument("--env-file", help="a secrets.env (default: the backup tool's config/secrets.env)")
    ap.add_argument("--max-mb", type=int, default=60)
    ap.add_argument("--dry-run", action="store_true", help="list what would change, write nothing")
    ap.add_argument("--ledger", action="store_true", help="read the decisions (ledger, report, prepared records) of the AdminOps "
                                                         "folder (ADMIN_STORE_SHARE_URL) instead of the countries' files")
    a = ap.parse_args(argv)
    try:
        run(a.dest, a.share_url, a.env_file, a.max_mb, a.dry_run, ledger=a.ledger)
    except Failed as e:
        print("Submissions pull FAILED: %s The local copy was not changed." % e, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
