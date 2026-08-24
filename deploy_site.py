#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deploy Claw-Report/site to the public GitHub Pages repo via an isolated clone.

The public site lives in its OWN repo (LeahLiuL/cul-bapfile-site) so the Pages
source branch is clean and under our control. We keep the heavy pipeline source
(server.py, process_all.py, bapfile.db, ...) in the private Claw-Report repo and
only publish the generated static site here.

IMPORTANT (corporate AV quirks on this machine):
  * The antivirus does per-file scanning on .git writes, which makes bulk file
    operations (clone, reset --hard of ~2200 files) extremely slow and prone to
    time-outs. We use `git reset --soft` (HEAD-only, no file rewrite) instead of
    `reset --hard` to sync with origin/main, halving the file-write volume.
  * The remote-tracking ref `origin/main` is NOT reliably written to disk here
    (AV intercepts the packed-refs update, so `rev-parse origin/main` can stay
    stale even after a successful fetch). We therefore NEVER trust `origin/main`;
    instead we fetch the true remote tip via `git ls-remote` and reset onto that
    exact SHA. This avoids both the slow reset AND the non-fast-forward rejects.
  * git ops are retried with a generous timeout, and any stale .git/index.lock
    left by an interrupted run is cleared (both at start and right before reset).
"""
import os, sys, shutil, time, subprocess

REPO = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(REPO, "site")
WT   = os.environ.get("DEPLOY_WT") or os.path.join(os.path.dirname(REPO), "cul-bapfile-site")
BRANCH = "main"
URL  = "https://github.com/LeahLiuL/cul-bapfile-site.git"


def run_git(args, cwd, retries=3, timeout=600):
    """Run a git command with retry + timeout. Returns True on success."""
    for attempt in range(1, retries + 1):
        print("  git", " ".join(args), "@", os.path.basename(cwd), "(try %d)" % attempt)
        try:
            subprocess.run(["git"] + args, cwd=cwd, check=True, timeout=timeout)
            return True
        except subprocess.TimeoutExpired:
            print("  ! git timed out, retrying...")
        except subprocess.CalledProcessError as e:
            print("  ! git exited %s, retrying..." % e.returncode)
    return False


def clear_stale_lock(wt_git):
    lock = os.path.join(wt_git, "index.lock")
    if os.path.exists(lock):
        try:
            os.remove(lock)
            print("[deploy] removed stale .git/index.lock")
        except OSError:
            pass


def remote_main_sha(cwd):
    """Return the true SHA of origin/main via ls-remote (no local ref trust)."""
    out = subprocess.run(["git", "ls-remote", URL, "refs/heads/" + BRANCH],
                         cwd=cwd, capture_output=True, text=True).stdout.strip()
    if not out:
        return None
    return out.split()[0]


def main():
    if not os.path.isdir(SITE):
        print("ERROR: site/ not found. Run gen_static.py first."); sys.exit(1)

    wt_git = os.path.join(WT, ".git")
    if not os.path.exists(wt_git):
        # fresh clone (slow under corporate AV; pipeline tolerates it)
        if os.path.exists(WT):
            shutil.rmtree(WT)
        if not run_git(["clone", URL, WT], cwd=REPO):
            print("ERROR: clone failed"); sys.exit(1)
    else:
        clear_stale_lock(wt_git)
        if not run_git(["fetch", "origin"], cwd=WT):
            print("ERROR: fetch failed"); sys.exit(1)

    rsha = remote_main_sha(WT)
    if not rsha:
        print("ERROR: cannot resolve remote main SHA"); sys.exit(1)
    # reset --SOFT (instant: HEAD-only, no file rewrite) moves main to the true
    # remote tip. Working tree keeps the prior published content, which we
    # overwrite with the freshly generated site below. Commit then fast-forwards.
    clear_stale_lock(wt_git)  # fetch may have left a lock if interrupted
    if not run_git(["reset", "--soft", rsha], cwd=WT):
        print("ERROR: reset failed"); sys.exit(1)

    # copy fresh site content (skip generation/deploy logs)
    for item in os.listdir(SITE):
        if item.endswith(".log"):
            continue
        s = os.path.join(SITE, item)
        d = os.path.join(WT, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, dirs_exist_ok=True)
        else:
            shutil.copy2(s, d)

    # NOTE: we intentionally do NOT prune files that vanished from the generated
    # site. The shard file set (cont prefixes / lanes / months) is stable across
    # daily runs, so stale files are effectively never produced; and a bulk
    # on-disk delete would trip the environment's safe-delete guard and break
    # this unattended nightly job. A few unused files in the repo are harmless
    # for a static site. If a shard is ever renamed, clean up manually once.

    # commit & push
    if not run_git(["add", "-A"], cwd=WT):
        print("ERROR: add failed"); sys.exit(1)
    st = subprocess.run(["git", "status", "--porcelain"], cwd=WT,
                        capture_output=True, text=True).stdout.strip()
    if not st:
        print("[deploy] no changes, skip")
        return
    msg = "deploy static bapfile site " + time.strftime("%Y-%m-%d %H:%M:%S")
    if not run_git(["commit", "-m", msg], cwd=WT):
        print("ERROR: commit failed"); sys.exit(1)
    if not run_git(["push", "-u", "origin", BRANCH], cwd=WT):
        print("ERROR: push failed (will retry next run)")
        sys.exit(1)
    print("[deploy] DONE ->", URL.replace(".git", "") + "/tree/" + BRANCH)
    print("[deploy] Pages  -> https://leahliul.github.io/cul-bapfile-site/")


if __name__ == "__main__":
    main()
