#!/usr/bin/env python3
"""Copy the engine's decision files to the Operations Center's AdminOps folder (inputs/), so the app wires the data the same way as the PC.

    corrections.csv       the decided corrections (campaign renames, dates, vaccines)
    lookup_overrides.csv  the decided lookup changes (rounds added, campaigns entered twice removed)
    scope_.csv            the districts planned per campaign, with the GPEI round facts

Called by `python admin.py sync-ops`. Only these files are written, only into the folder inputs/ of the AdminOps folder (share link ADMIN_STORE_SHARE_URL),
with the same app sign-in the backup tool and the Operations Center use. A file that has not changed since the last copy is skipped. Before a file in
the store is replaced, its current content is saved in the local backup folder. SharePoint keeps its version history.
"""
import argparse
import datetime as dt
import hashlib
import json
import sys
from pathlib import Path
from urllib.parse import quote

import pull_submissions as ps

SIMPLE_MAX = 4 * 1024 * 1024
CHUNK = 320 * 1024 * 32
FOLDER = "inputs"


class WriteApi(ps.GraphApi):
    def _put(self, url, data, headers=None):
        import time
        last = None
        for attempt in range(4):
            try:
                resp = self.r.put(url, headers=dict(self.h, **(headers or {})), data=data, timeout=300)
            except self.r.exceptions.RequestException as e:
                last = e
                time.sleep(3 * (attempt + 1))
                continue
            if resp.status_code in (429, 502, 503, 504):
                time.sleep(max(int(resp.headers.get("Retry-After", 0) or 0), 4 * (attempt + 1)))
                last = resp
                continue
            return resp
        raise ps.Failed("network error writing to SharePoint (%s)." % (type(last).__name__ if last is not None else "unknown"))

    def item(self, drive, root_id, rel):
        resp = self._get("%s/drives/%s/items/%s:/%s" % (ps.GRAPH, drive, root_id, quote(rel)))
        if resp.status_code == 404:
            return None
        if resp.status_code != 200:
            raise ps.Failed("cannot look at %s (HTTP %s)." % (rel, resp.status_code))
        return resp.json()

    def write(self, drive, root_id, rel, data):
        base = "%s/drives/%s/items/%s:/%s" % (ps.GRAPH, drive, root_id, quote(rel))
        if len(data) <= SIMPLE_MAX:
            r = self._put(base + ":/content", data, {"Content-Type": "application/octet-stream"})
            if r.status_code in (403, 401):
                raise ps.Failed("SharePoint refused the write (HTTP %s): the app registration can read the AdminOps folder but not write to it. Copy the files by hand." % r.status_code)
            if r.status_code not in (200, 201):
                raise ps.Failed("cannot write %s (HTTP %s)." % (rel, r.status_code))
            return
        r = self.r.post(base + ":/createUploadSession", headers=self.h, json={"item": {"@microsoft.graph.conflictBehavior": "replace"}}, timeout=60)
        if r.status_code in (403, 401):
            raise ps.Failed("SharePoint refused the write (HTTP %s): the app registration can read the AdminOps folder but not write to it. Copy the files by hand." % r.status_code)
        if r.status_code not in (200, 201):
            raise ps.Failed("cannot start the upload of %s (HTTP %s)." % (rel, r.status_code))
        url, size, start = r.json()["uploadUrl"], len(data), 0
        while start < size:
            chunk = data[start:start + CHUNK]
            rr = self.r.put(url, data=chunk, timeout=300, headers={"Content-Length": str(len(chunk)),
                                                                  "Content-Range": "bytes %d-%d/%d" % (start, start + len(chunk) - 1, size)})
            if rr.status_code not in (200, 201, 202):
                raise ps.Failed("the upload of %s failed (HTTP %s)." % (rel, rr.status_code))
            start += len(chunk)


def sha(b):
    return hashlib.sha256(b).hexdigest()


def run(files, state_path, backup_dir, env_file=None, dry_run=False, force=False, share_url=None):
    """files: {name in the store: local Path}. Returns a list of (name, what happened)."""
    keys = ("ADMIN_STORE_SHARE_URL", "SHAREPOINT_TENANT_ID", "SHAREPOINT_CLIENT_ID", "SHAREPOINT_CLIENT_SECRET")
    env = ps.read_env(env_file, keys)
    url = (share_url or env["ADMIN_STORE_SHARE_URL"]).strip()
    if not url:
        raise ps.Failed("no share link: set ADMIN_STORE_SHARE_URL in the backup tool's config/secrets.env.")
    if not all(env[k] for k in keys[1:]):
        raise ps.Failed("SHAREPOINT_TENANT_ID / _CLIENT_ID / _CLIENT_SECRET are not set (the backup tool's config/secrets.env).")
    api = WriteApi(env["SHAREPOINT_TENANT_ID"], env["SHAREPOINT_CLIENT_ID"], env["SHAREPOINT_CLIENT_SECRET"])
    root = api.root(url)
    state = json.loads(Path(state_path).read_text(encoding="utf-8")) if Path(state_path).is_file() else {}
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = []
    for name, local in files.items():
        local = Path(local)
        if not local.is_file():
            out.append((name, "no local file: skipped"))
            continue
        data = local.read_bytes()
        h = sha(data)
        rel = "%s/%s" % (FOLDER, name)
        remote = api.item(root["drive"], root["id"], rel)
        rsize = int(remote.get("size") or 0) if remote else None
        if not force and remote and state.get(name, {}).get("sha256") == h and rsize == len(data):
            out.append((name, "unchanged (%d bytes), skipped" % len(data)))
            continue
        if remote and rsize == len(data) and state.get(name, {}).get("sha256") is None:
            # never copied by this command: if the stored copy is identical, only record it
            if sha(api.download(root["drive"], {"id": remote["id"], "name": name, "url": remote.get("@microsoft.graph.downloadUrl")})) == h:
                state[name] = {"sha256": h, "size": len(data), "on": stamp}
                out.append((name, "already identical in the store, recorded"))
                continue
        if dry_run:
            out.append((name, "WOULD %s (%d bytes)" % ("replace" if remote else "create", len(data))))
            continue
        if remote:
            keep = Path(backup_dir) / stamp
            keep.mkdir(parents=True, exist_ok=True)
            (keep / name).write_bytes(api.download(root["drive"], {"id": remote["id"], "name": name, "url": remote.get("@microsoft.graph.downloadUrl")}))
        api.write(root["drive"], root["id"], rel, data)
        after = api.item(root["drive"], root["id"], rel)
        if not after or int(after.get("size") or 0) != len(data):
            raise ps.Failed("%s: after the upload the store holds %s bytes, expected %d." % (name, (after or {}).get("size"), len(data)))
        state[name] = {"sha256": h, "size": len(data), "on": stamp}
        out.append((name, "%s (%d bytes), checked" % ("replaced" if remote else "created", len(data))))
    if not dry_run:
        Path(state_path).write_text(json.dumps(state, indent=1), encoding="utf-8")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--file", action="append", required=True, help="NAME=PATH, repeatable")
    ap.add_argument("--state", required=True)
    ap.add_argument("--backup-dir", required=True)
    ap.add_argument("--env-file")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args(argv)
    files = dict(f.split("=", 1) for f in a.file)
    try:
        res = run(files, a.state, a.backup_dir, a.env_file, a.dry_run, a.force)
    except ps.Failed as e:
        print("Sync to AdminOps FAILED: %s Nothing more was written." % e, file=sys.stderr)
        return 1
    for n, w in res:
        print("  %-22s %s" % (n, w))
    return 0


if __name__ == "__main__":
    sys.exit(main())
