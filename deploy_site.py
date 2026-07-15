#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deploy Claw-Report/site to the public GitHub Pages repo via an isolated clone.

The public site lives in its OWN repo (LeahLiuL/cul-bapfile-site) so the Pages
source branch is clean and under our control. We keep the heavy pipeline source
(server.py, process_all.py, bapfile.db, ...) in the private Claw-Report repo and
only publish the generated static site here.

Why a separate working clone: the main working tree holds source + the 725 MB
bapfile.db. We never touch it. Instead we keep a clone at C:/Users/leahliu/cul-bapfile-site
checked out to `main`, copy the static site into it, commit and push. Idempotent.
"""
import os, sys, shutil, time, subprocess

REPO = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(REPO, "site")
WT   = os.environ.get("DEPLOY_WT") or os.path.join(os.path.dirname(REPO), "cul-bapfile-site")
BRANCH = "main"
URL  = "https://github.com/LeahLiuL/cul-bapfile-site.git"


def git(args, cwd):
    print("  git", " ".join(args), "@", os.path.basename(cwd))
    subprocess.run(["git"] + args, cwd=cwd, check=True)


def main():
    if not os.path.isdir(SITE):
        print("ERROR: site/ not found. Run gen_static.py first."); sys.exit(1)

    # 1. ensure working clone exists (cloned from the public repo)
    wt_git = os.path.join(WT, ".git")
    if os.path.exists(wt_git):
        print("[deploy] reusing clone", WT)
        git(["fetch", "origin"], cwd=WT)
        git(["checkout", BRANCH], cwd=WT)
        git(["reset", "--hard", "origin/" + BRANCH], cwd=WT)
    else:
        if os.path.exists(WT):
            shutil.rmtree(WT)
        git(["clone", URL, WT], cwd=REPO)
        git(["checkout", BRANCH], cwd=WT)

    # 2. copy fresh site content (skip generation/deploy logs)
    for item in os.listdir(SITE):
        if item.endswith(".log"):
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
    print("[deploy] Pages  -> https://leahliul.github.io/cul-bapfile-site/")


if __name__ == "__main__":
    main()
