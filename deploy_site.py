#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deploy Claw-Report/site to the gh-pages branch via an isolated git worktree.

Why a worktree: the main working tree holds source (server.py, process_all.py, ...)
and the 725 MB bapfile.db. We never touch it. Instead we create a separate worktree
at C:/Users/leahliu/claw-pages checked out to gh-pages, copy the static site there,
commit and push. Idempotent: safe to run again for updates.
"""
import os, sys, shutil, time, subprocess

REPO = r"C:/Users/leahliu/Claw-Report"
SITE = os.path.join(REPO, "site")
WT   = r"C:/Users/leahliu/claw-pages"
BRANCH = "gh-pages"
URL  = "https://github.com/LeahLiuL/Claw-Report.git"


def git(args, cwd):
    print("  git", " ".join(args), "@", os.path.basename(cwd))
    subprocess.run(["git"] + args, cwd=cwd, check=True)


def main():
    if not os.path.isdir(SITE):
        print("ERROR: site/ not found. Run gen_static.py first."); sys.exit(1)

    # 1. ensure worktree exists
    wt_git = os.path.join(WT, ".git")
    if os.path.exists(wt_git):
        print("[deploy] reusing worktree", WT)
        git(["fetch", "origin"], cwd=WT)
        git(["checkout", BRANCH], cwd=WT)
        git(["reset", "--hard", "origin/" + BRANCH], cwd=WT)
    else:
        if os.path.exists(WT):
            shutil.rmtree(WT)
        git(["fetch", "origin"], cwd=REPO)
        res = subprocess.run(["git", "ls-remote", "--heads", URL, BRANCH],
                             capture_output=True, text=True)
        if res.stdout.strip():
            git(["worktree", "add", WT, BRANCH], cwd=REPO)
        else:
            git(["worktree", "add", "-b", BRANCH, WT], cwd=REPO)
            git(["rm", "-rf", "."], cwd=WT)
            git(["commit", "--allow-empty", "-m", "init gh-pages"], cwd=WT)

    # 2. copy fresh site content (skip generation log)
    for item in os.listdir(SITE):
        if item == "gen.log":
            continue
        s = os.path.join(SITE, item)
        d = os.path.join(WT, item)
        if os.path.isdir(s):
            if os.path.exists(d):
                shutil.rmtree(d)
            shutil.copytree(s, d)
        else:
            shutil.copy2(s, d)

    # 3. commit & push
    git(["add", "-A"], cwd=WT)
    st = subprocess.run(["git", "status", "--porcelain"], cwd=WT,
                        capture_output=True, text=True).stdout.strip()
    if not st:
        print("[deploy] no changes, skip")
        return
    git(["commit", "-m", "deploy static bapfile site " + time.strftime("%Y-%m-%d %H:%M:%S")], cwd=WT)
    git(["push", "-u", "origin", BRANCH], cwd=WT)
    print("[deploy] DONE ->", URL.replace(".git", "") + "/tree/" + BRANCH)


if __name__ == "__main__":
    main()
