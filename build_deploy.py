#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Full rebuild + deploy pipeline for the CULines Vessel Bapfile static site.

Steps:
  0) Download the latest Vessel Bapfile.xlsx from the company SFTP server
     (sftp_fetch.py) — the real data source, NOT a local copy. VPN required.
  1) Rebuild bapfile.db from the downloaded xlsx (process_all.py)
  2) Generate static shards (gen_static.py)
  3) Deploy to GitHub Pages (deploy_site.py)

Intended to be run by a WorkBuddy automation on a schedule so the public site
stays in sync with the spreadsheet on the FTP server. Requires VPN to reach
10.5.4.2 and git push credentials to be available in the local environment.
"""
import os, subprocess, sys

REPO = r"C:/Users/leahliu/Claw-Report"
PY   = r"C:/Users/leahliu/.workbuddy/binaries/python/versions/3.13.12/python.exe"
DB   = os.path.join(REPO, "bapfile.db")
SITE = os.path.join(REPO, "site")


def run(cmd):
    print("+", " ".join(cmd))
    subprocess.run(cmd, check=True, cwd=REPO)


def main():
    print("[0/4] downloading latest Vessel Bapfile.xlsx from SFTP ...")
    run([PY, os.path.join(REPO, "sftp_fetch.py")])

    print("[1/4] rebuilding bapfile.db from xlsx ...")
    run([PY, os.path.join(REPO, "process_all.py")])

    print("[2/4] generating static shards ...")
    run([PY, os.path.join(REPO, "gen_static.py"), "--db", DB, "--out", SITE])

    print("[3/4] deploying to gh-pages ...")
    run([PY, os.path.join(REPO, "deploy_site.py")])

    print("ALL DONE")


if __name__ == "__main__":
    main()
