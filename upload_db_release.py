#!/usr/bin/env python3
"""Create a GitHub Release and upload the two SQLite databases as assets.
Token is read from git credential (never hardcoded/logged).
Default dry-run; pass --apply to actually create/upload.
"""
import os
import json
import subprocess
import urllib.request
import urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
OWNER = "LeahLiuL"
REPO = "Claw-Report"
TAG = "db-snapshot-2026-11-23"
RELEASE_NAME = "Claw databases snapshot 2026-11-23"
BODY = (
    "Database snapshot: reports.db (full-field report warehouse) and "
    "abstract_log.db (Abstract Log voyage/leg tables). "
    "Contains real operational data - internal use only."
)
ASSETS = [
    os.path.join(HERE, "rob_data", "abstract_log.db"),
    os.path.join(HERE, "rob_data", "reports.db"),
]


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
    body = None
    if data is not None:
        body = json.dumps(data).encode("utf-8")
    elif raw is not None:
        body = raw
    req = urllib.request.Request(url, data=body, method=method)
    req.add_header("Authorization", "Bearer " + token)
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    if body is not None:
        req.add_header("Content-Type", ctype)
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read().decode("utf-8"))


def main():
    import sys
    apply = "--apply" in sys.argv
    token = get_token()
    if not token:
        print("ERROR: cannot read token from git credential")
        return
    api_base = "https://api.github.com/repos/%s/%s" % (OWNER, REPO)

    if not apply:
        print("DRY-RUN (no changes). Files to upload:")
        for f in ASSETS:
            print("  %s  %.1f MB" % (os.path.basename(f), os.path.getsize(f) / 1e6))
        print("Release tag:", TAG)
        print("Add --apply to create release and upload.")
        return

    # find existing release by tag (idempotent)
    rel = None
    try:
        rel = api("GET", api_base + "/releases/tags/" + TAG, token)
        print("Existing release found (id=%s), reusing." % rel["id"])
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    if rel is None:
        rel = api("POST", api_base + "/releases", token, data={
            "tag_name": TAG,
            "name": RELEASE_NAME,
            "body": BODY,
            "draft": False,
            "prerelease": False,
        })
        print("Release created: id=%s url=%s" % (rel["id"], rel["html_url"]))

    upload_base = "https://uploads.github.com/repos/%s/%s/releases/%s/assets" % (
        OWNER, REPO, rel["id"])

    # skip already-uploaded assets
    existing = {a["name"] for a in api("GET", api_base + "/releases/%s/assets" % rel["id"], token)}
    for f in ASSETS:
        name = os.path.basename(f)
        if name in existing:
            print("SKIP already uploaded:", name)
            continue
        size = os.path.getsize(f)
        print("Uploading %s (%.1f MB) ..." % (name, size / 1e6), flush=True)
        with open(f, "rb") as fh:
            raw = fh.read()
        asset = api("POST", upload_base + "?name=" + name, token, raw=raw,
                    ctype="application/octet-stream")
        print("  OK ->", asset["browser_download_url"])

    print("\nRelease page:", rel["html_url"])
    print("DONE")


if __name__ == "__main__":
    main()
