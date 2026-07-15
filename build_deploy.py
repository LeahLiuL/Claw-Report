#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Incremental rebuild + deploy pipeline for the CULines Vessel Bapfile static site.

The FTP source is a ROLLING WINDOW that overlaps previous data, so we ACCUMULATE
(new rows added, exact-duplicate rows skipped) — never a full replace.

Steps:
  0) Download the latest Vessel Bapfile.xlsx from the company SFTP server
     (sftp_fetch.py) — the real data source, NOT a local copy. VPN required.
  1) Append the xlsx into bapfile.db, skipping exact-duplicate rows
     (process_all.py --append). On first run it also de-duplicates the existing db.
  2) Generate static shards (gen_static.py)
  3) Deploy to GitHub Pages (deploy_site.py)

Intended to be run by a WorkBuddy automation on a schedule so the public site
stays in sync with the spreadsheet on the FTP server. Requires VPN to reach
10.5.4.2 and git push credentials to be available in the local environment.
"""
import os, subprocess, sys

REPO = os.path.dirname(os.path.abspath(__file__))
PY   = sys.executable
DB   = os.path.join(REPO, "bapfile.db")
SITE = os.path.join(REPO, "site")


def run(cmd):
    print("+", " ".join(cmd))
    subprocess.run(cmd, check=True, cwd=REPO)


def main():
    print("[0/4] downloading latest Vessel Bapfile.xlsx from SFTP ...")
    run([PY, os.path.join(REPO, "sftp_fetch.py")])

    print("[1/4] appending latest xlsx into bapfile.db (dedup exact-duplicate rows) ...")
    run([PY, os.path.join(REPO, "process_all.py"), "--append"])

    print("[2/4] generating static shards ...")
    run([PY, os.path.join(REPO, "gen_static.py"), "--db", DB, "--out", SITE])

    print("[3/4] deploying to gh-pages ...")
    run([PY, os.path.join(REPO, "deploy_site.py")])

    print("ALL DONE")


if __name__ == "__main__":
    main()
