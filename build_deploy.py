#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Incremental rebuild + deploy pipeline for the CULines Vessel Bapfile static site.

The FTP source is a ROLLING WINDOW that overlaps previous data, so we ACCUMULATE
(new rows added, exact-duplicate rows skipped) — never a full replace.

Steps:
  0) Download the latest Vessel Bapfile.xlsx from the company SFTP server
     (sftp_fetch.py) — the real data source, NOT a local copy. VPN required.
     FALLBACK: if sftp_fetch.py fails because the server-side file is a
     truncated streaming ZIP (complete byte count but missing central
     directory), use _resume_download_wk37.py (resume-aware downloader) then
     repair_bapfile.py (rebuild valid ZIP) and feed the repaired file into
     step 1 via the BAPFILE_LOCAL env var. This fallback was proven on
     2026-09-22 / 09-23 and is now wired in automatically.
  1) Append the xlsx into bapfile.db, skipping exact-duplicate rows
     (process_all.py --append). On first run it also de-duplicates the existing db.
  2) Generate static shards (gen_static.py)
  3) Deploy to GitHub Pages (deploy_site.py)

CONCURRENCY: two schedulers run this script at 23:00 (the WorkBuddy automation
and the user's auto_bapfile.py Windows task). A pipeline lock serializes them:
  - shared state (bapfile.db, site/ git repo, the .dl resume temp) must not be
    touched by two runs at once;
  - the second run is still worth completing afterwards: append dedups to 0 new
    rows and the deploy is a manifest-timestamp-only push (idempotent).
A lock older than 40 minutes is considered stale (crashed run) and is broken.

Intended to be run by a WorkBuddy automation on a schedule so the public site
stays in sync with the spreadsheet on the FTP server. Requires VPN to reach
10.5.4.2 and git push credentials to be available in the local environment.
"""
import os, subprocess, sys, time

REPO = os.path.dirname(os.path.abspath(__file__))
PY   = sys.executable
DB   = os.path.join(REPO, "bapfile.db")
SITE = os.path.join(REPO, "site")
LOCK = os.path.join(REPO, ".bapfile_pipeline.lock")

CULINES_DIR = r"C:\CULINES\Claw Report"
REMOTE_XLSX = os.path.join(CULINES_DIR, "Vessel Bapfile.xlsx")
REPAIRED_XLSX = os.path.join(CULINES_DIR, "Vessel Bapfile.repaired.xlsx")


def run(cmd, extra_env=None):
    print("+", " ".join(cmd))
    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)
    subprocess.run(cmd, check=True, cwd=REPO, env=env)


def acquire_lock(timeout_s=900, stale_s=2400):
    """Wait up to timeout_s for the pipeline lock. Stale locks are broken."""
    deadline = time.time() + timeout_s
    while True:
        try:
            fd = os.open(LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            return True
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(LOCK) > stale_s:
                    print(f"[lock] stale lock ({stale_s//60}min) — breaking it")
                    os.remove(LOCK)
                    continue
            except OSError:
                continue  # lock vanished; retry immediately
            if time.time() > deadline:
                return False
            print("[lock] another pipeline run is active — waiting 30s ...")
            time.sleep(30)


def release_lock():
    try:
        os.remove(LOCK)
    except OSError:
        pass


def fetch_source():
    """Step 0: get a usable local xlsx. Returns the path to feed into step 1.

    Normal path: sftp_fetch.py downloads to the repo root (returns None ->
    process_all.py uses its default SRC).
    Fallback path (server-side truncated ZIP): resume-download + repair, and
    return the repaired path so step 1 reads it via BAPFILE_LOCAL.
    """
    print("[0/4] downloading latest Vessel Bapfile.xlsx from SFTP ...")
    try:
        run([PY, os.path.join(REPO, "sftp_fetch.py")])
        return None
    except subprocess.CalledProcessError:
        print("[0/4] sftp_fetch.py failed — trying resume-download + repair fallback ...")
        run([PY, os.path.join(REPO, "_resume_download_wk37.py")])
        # _resume_download_wk37.py exits 0 even when the ZIP is truncated
        # (prints REPAIR_NEEDED); repair_bapfile.py rebuilds a valid xlsx.
        run([PY, os.path.join(REPO, "repair_bapfile.py"), REMOTE_XLSX, REPAIRED_XLSX])
        print("[0/4] fallback OK — repaired xlsx will be appended: " + REPAIRED_XLSX)
        return REPAIRED_XLSX


def main():
    if not acquire_lock():
        print("[lock] TIMEOUT: another pipeline run held the lock for 15min — aborting this run")
        sys.exit(3)
    try:
        src = fetch_source()

        print("[1/4] appending latest xlsx into bapfile.db (dedup exact-duplicate rows) ...")
        run([PY, os.path.join(REPO, "process_all.py"), "--append"],
            extra_env=({"BAPFILE_LOCAL": src} if src else None))

        print("[2/4] generating static shards ...")
        run([PY, os.path.join(REPO, "gen_static.py"), "--db", DB, "--out", SITE])

        print("[3/4] deploying to gh-pages ...")
        run([PY, os.path.join(REPO, "deploy_site.py")])

        print("ALL DONE")
    finally:
        release_lock()


if __name__ == "__main__":
    main()
