#!/usr/bin/env python3
"""Upload ONLY abstract_log.db to the existing GitHub Release, replacing the old asset.

Small file (~0.5M). Idempotent: deletes the existing 'abstract_log.db' asset then
uploads the fresh one. reports.db is NEVER touched here.

Token from git credential. Default dry-run; pass --apply to actually upload.
"""
import os
import sys
import json
import sqlite3
import subprocess
import urllib.request
import urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
OWNER = "LeahLiuL"
REPO = "Claw-Report"
TAG = "db-snapshot-2026-11-23"
DB = os.path.join(HERE, "rob_data", "abstract_log.db")


def get_token():
    p = subprocess.run(
        ["git", "credential", "fill"],
        input="protocol=https\nhost=github.com\n\n",
        capture_output=True, text=True, cwd=HERE,
    )
    for line in p.stdout.splitlines():
        if line.startswith("password="):
            return line.split("=", 1)[1]
    return ""


def api(method, url, token, data=None, raw=None, ctype="application/json"):
    body = json.dumps(data).encode("utf-8") if data is not None else raw
    req = urllib.request.Request(url, data=body, method=method)
    req.add_header("Authorization", "Bearer " + token)
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    if body is not None:
        req.add_header("Content-Type", ctype)
    with urllib.request.urlopen(req, timeout=120) as r:
        txt = r.read().decode("utf-8")
        return json.loads(txt) if txt else None


def checkpoint(db_path):
    """Force WAL -> main db so the copied/uploaded file has all rows."""
    c = sqlite3.connect(db_path)
    try:
        mode = c.execute("PRAGMA journal_mode").fetchone()[0]
        if str(mode).lower() == "wal":
            c.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        n = c.execute("SELECT count(*) FROM voyages").fetchone()[0]
        legs = c.execute("SELECT count(*) FROM legs").fetchone()[0]
        print("[checkpoint] journal=%s voyages=%d legs=%d" % (mode, n, legs))
        return n
    finally:
        c.close()


def main():
    apply = "--apply" in sys.argv
    token = get_token()
    if not token:
        print("ERROR: cannot read token from git credential")
        return 1
    api_base = "https://api.github.com/repos/%s/%s" % (OWNER, REPO)

    checkpoint(DB)
    size = os.path.getsize(DB)
    if not apply:
        print("DRY-RUN. Would upload %s (%.2f MB) to tag %s" % (
            os.path.basename(DB), size / 1e6, TAG))
        return 0

    rel = api("GET", api_base + "/releases/tags/" + TAG, token)
    rel_id = rel["id"]
    assets = api("GET", api_base + "/releases/%s/assets" % rel_id, token)
    old = [a for a in assets if a["name"] == "abstract_log.db"]
    for a in old:
        print("DELETE old asset id=%s" % a["id"])
        api("DELETE", api_base + "/releases/assets/%s" % a["id"], token)

    upload_url = "https://uploads.github.com/repos/%s/%s/releases/%s/assets?name=abstract_log.db" % (
        OWNER, REPO, rel_id)
    print("Uploading abstract_log.db (%.2f MB) ..." % (size / 1e6), flush=True)
    with open(DB, "rb") as fh:
        raw = fh.read()
    asset = api("POST", upload_url, token, raw=raw, ctype="application/octet-stream")
    print("  OK state=%s -> %s" % (asset.get("state"), asset.get("browser_download_url")))
    print("DONE")
    return 0


if __name__ == "__main__":
    sys.exit(main())
