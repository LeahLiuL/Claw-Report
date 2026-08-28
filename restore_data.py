#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
从指定 commit 中把 TODAY_DATA 数据块取回，注入当前 HTML。
用途：误用旧数据覆盖后，从 culadmin 的提交里恢复最新数据，同时保留代码修复。
"""
import os, re, sys, subprocess

REPO = os.path.dirname(os.path.abspath(__file__))
HTML = "cul_daily_movement.html"


def run(cmd, check=True):
    p = subprocess.run(cmd, cwd=REPO, shell=True, capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    out = (p.stdout or "") + (p.stderr or "")
    if check and p.returncode != 0:
        raise RuntimeError("failed: {}\n{}".format(cmd, out))
    return out


def balanced_extract(text, name):
    m = re.search(r"^\s*const\s+" + re.escape(name) + r"\s*=\s*", text, re.M)
    if not m:
        return None
    start = m.end()
    if text[start] not in "{[":
        return None
    depth, i, in_str, esc = 0, start, False, False
    while i < len(text):
        ch = text[i]
        if in_str:
            if esc: esc = False
            elif ch == "\\": esc = True
            elif ch == '"': in_str = False
        else:
            if ch == '"': in_str = True
            elif ch in "{[": depth += 1
            elif ch in "}]":
                depth -= 1
                if depth == 0:
                    return text[start:i+1]
        i += 1
    return None


def balanced_replace(text, name, newval):
    m = re.search(r"^\s*const\s+" + re.escape(name) + r"\s*=\s*", text, re.M)
    if not m:
        raise RuntimeError("block not found: " + name)
    start = m.end()
    if text[start] not in "{[":
        raise RuntimeError("bad start for " + name)
    depth, i, in_str, esc = 0, start, False, False
    while i < len(text):
        ch = text[i]
        if in_str:
            if esc: esc = False
            elif ch == "\\": esc = True
            elif ch == '"': in_str = False
        else:
            if ch == '"': in_str = True
            elif ch in "{[": depth += 1
            elif ch in "}]":
                depth -= 1
                if depth == 0:
                    return text[:start] + newval + text[i+1:]
        i += 1
    raise RuntimeError("unbalanced: " + name)


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else None
    if not src:
        print("usage: python restore_data.py <commit-with-good-data>")
        return 1

    print("[restore] reading data from {} ...".format(src))
    src_html = run("git show {}:{}".format(src, HTML))

    if not os.path.exists(os.path.join(REPO, HTML)):
        raise RuntimeError("local HTML missing")
    with open(os.path.join(REPO, HTML), "r", encoding="utf-8") as f:
        cur = f.read()

    for name in ("TODAY_DATA", "SNAPSHOTS"):
        sv = balanced_extract(src_html, name)
        cv = balanced_extract(cur, name)
        if sv is None:
            print("  skip {} (absent in source)".format(name)); continue
        if cv is None:
            print("  skip {} (absent in current)".format(name)); continue
        if sv == cv:
            print("  {} identical".format(name)); continue
        cur = balanced_replace(cur, name, sv)
        print("  {} restored from {}".format(name, src))

    with open(os.path.join(REPO, HTML), "w", encoding="utf-8", newline="") as f:
        f.write(cur)

    # verify
    for token, label in (("AGENT_BY_PORT", "Agent/OP"), ("berthedCalls", "Berth Rate")):
        print("  {} {}".format("OK " if token in cur else "MISSING", label))
    ts = re.search(r'"generatedAt": "([^"]*)"', cur)
    print("[restore] generatedAt now = {}".format(ts.group(1) if ts else "?"))
    print("[restore] done. review with: git diff --stat")
    return 0


if __name__ == "__main__":
    sys.exit(main())
