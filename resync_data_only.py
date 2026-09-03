#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
resync_data_only.py —— 只把"最新数据"搬进 HTML，绝不把旧代码逻辑覆盖回去。

背景
----
culadmin 每天用 auto_update.py 重新生成整个 cul_daily_movement.html 并推送。
它本地的 gen_html.py 可能不是最新版，于是"旧代码生成的新 HTML"会把
我们推上去的代码修复（Agent/OP 列、Berth Rate 口径、Per-Port Rate 等）覆盖掉。

本脚本的解法：**数据与代码解耦**
  - 从 culadmin 最新上传的 HTML 中，只抽取数据块（TODAY_DATA / SNAPSHOTS）
  - 注入到本机用**最新 gen_html.py** 生成的 HTML 中
  - 结果：代码 = 我们修好的最新版；数据 = culadmin 当天最新上传

前置条件（重要）
----------------
  1. 不要在 leahliu 本机跑 gen_html.py 去覆盖 culadmin 的数据 —— 本机数据可能更旧。
     本脚本正是为了避免这件事：它只用本机的**代码**，用远端的**数据**。
  2. culadmin 的 HTML 必须先 fetch 下来。脚本会自动 `git fetch` 并读取远端版本。

用法
----
  python resync_data_only.py            # 同步数据 -> 生成 -> 提交 -> 推送
  python resync_data_only.py --dry-run   # 只比对，不写文件不推送
"""

import os
import re
import subprocess
import sys
import json

REPO = os.path.dirname(os.path.abspath(__file__))
HTML = "cul_daily_movement.html"

# 数据块搬运计划: (块名, 取值来源)
#   remote = 用 culadmin(远端)HTML 的值
#   local  = 保留本机重生成的值
#
# 2026-08-31 实测结论（务必遵守，否则会丢数据）：
#   TODAY_DATA    -> remote  本机 CUL DAILY MOVEMENT.rebuilt.xlsx 是残缺快照
#                            (44 船、无"-已下线"标记)，远端 54 船、含 8 艘已下线
#   MAINT_DATA    -> remote  本机 MAINT 源只有 9283 条，远端 24774 条
#   AGENT_BY_PORT -> local   culadmin 机器读不到 P: 盘联系人表，远端该块是空 {}
#   SNAPSHOTS     -> remote
BLOCK_PLAN = (
    ("TODAY_DATA",    "remote"),
    ("MAINT_DATA",    "remote"),
    ("AGENT_BY_PORT", "local"),
    ("SNAPSHOTS",     "remote"),
)

# 本机重生成会把 maint_snapshot.json 写成残缺版本，必须回退远端版本
SNAPSHOT_FILE = "maint_snapshot.json"


def run(cmd, check=True, capture=True):
    p = subprocess.run(
        cmd, cwd=REPO, shell=True,
        capture_output=capture, text=True, encoding="utf-8", errors="replace",
    )
    out = (p.stdout or "") + (p.stderr or "")
    if check and p.returncode != 0:
        raise RuntimeError("Command failed: {}\n{}".format(cmd, out))
    return p.returncode, out


def log(m):
    print("[resync] {}".format(m), flush=True)


def extract_data_block(text, name):
    """从 HTML 文本中精确抽取 `const NAME = ...;` 这一整段（支持跨行）。

    用括号配平定位结尾，避免正则吃掉后面的内容。
    """
    m = re.search(r"^\s*const\s+" + re.escape(name) + r"\s*=\s*", text, re.M)
    if not m:
        return None
    start = m.end()
    opener = text[start]
    if opener not in "{[":
        return None
    closer = "}" if opener == "{" else "]"
    depth = 0
    i = start
    in_str = False
    esc = False
    while i < len(text):
        ch = text[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch in "{[":
                depth += 1
            elif ch in "}]":
                depth -= 1
                if depth == 0:
                    return text[start:i + 1]
        i += 1
    return None


def replace_data_block(text, name, new_value):
    """把 HTML 中 `const NAME = <旧值>;` 替换为新值。返回新文本。"""
    m = re.search(r"^\s*const\s+" + re.escape(name) + r"\s*=\s*", text, re.M)
    if not m:
        raise RuntimeError("data block not found in local HTML: {}".format(name))
    start = m.end()
    opener = text[start]
    if opener not in "{[":
        raise RuntimeError("unexpected data block start for {}: {!r}".format(name, opener))
    closer = "}" if opener == "{" else "]"
    depth = 0
    i = start
    in_str = False
    esc = False
    while i < len(text):
        ch = text[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch in "{[":
                depth += 1
            elif ch in "}]":
                depth -= 1
                if depth == 0:
                    return text[:start] + new_value + text[i + 1:]
        i += 1
    raise RuntimeError("unbalanced braces for block {}".format(name))


def fetch_remote_html():
    """从远端 main 分支拿到 culadmin 上传的最新 HTML 文本。

    注意：本 sandbox 环境下 `origin/main` 跟踪引用会滞后，必须用 `FETCH_HEAD`
    （git fetch 写入的真实远端 hash）读取，否则会拿到旧版本。
    """
    log("fetch origin ...")
    run("git fetch origin")
    rc, out = run("git show FETCH_HEAD:{}".format(HTML), check=False)
    if rc != 0:
        raise RuntimeError("cannot read {} from FETCH_HEAD: {}".format(HTML, out))
    return out


def main():
    dry = "--dry-run" in sys.argv

    try:
        remote_html = fetch_remote_html()

        # 1) 本机用【最新 gen_html.py】生成 HTML（本机可能数据旧，但代码最新）
        log("regenerating HTML locally with latest gen_html.py ...")
        rc, out = run("python gen_html.py", check=False)
        if rc != 0:
            raise RuntimeError("gen_html.py failed:\n{}".format(out))
        with open(os.path.join(REPO, HTML), "r", encoding="utf-8") as f:
            local_html = f.read()

        # 2) 按 BLOCK_PLAN 逐块决定取远端还是保留本机
        changed = []
        for name, direction in BLOCK_PLAN:
            remote_val = extract_data_block(remote_html, name)
            local_val = extract_data_block(local_html, name)
            if remote_val is None or local_val is None:
                log("  skip {} (remote={} local={})".format(
                    name, remote_val is not None, local_val is not None))
                continue
            if direction == "local":
                log("  {:<16} keep LOCAL  ({} bytes; remote was {} — known to be worse)".format(
                    name, len(local_val), len(remote_val)))
                continue
            if remote_val == local_val:
                log("  {:<16} identical, nothing to do".format(name))
                continue
            local_html = replace_data_block(local_html, name, remote_val)
            changed.append(name)
            log("  {:<16} <- REMOTE ({} bytes, local was {})".format(
                name, len(remote_val), len(local_val)))

        if not changed:
            log("data already up to date; no rewrite needed")
            if dry:
                return 0
            # 即使数据一致，也确保工作区 HTML 是"本机最新代码"生成的
            with open(os.path.join(REPO, HTML), "w", encoding="utf-8", newline="") as f:
                f.write(local_html)
            log("wrote HTML (regenerated with latest code)")
        else:
            if dry:
                log("dry-run: would rewrite {}".format(", ".join(changed)))
                return 0
            with open(os.path.join(REPO, HTML), "w", encoding="utf-8", newline="") as f:
                f.write(local_html)
            log("wrote HTML: latest code + culadmin data ({})".format(", ".join(changed)))

        # 3) 回退 maint_snapshot.json 到远端版本（本机重生成会截断它，必须还原）
        rc, out = run("git show FETCH_HEAD:{}".format(SNAPSHOT_FILE), check=False)
        if rc == 0 and out:
            with open(os.path.join(REPO, SNAPSHOT_FILE), "w", encoding="utf-8", newline="") as f:
                f.write(out)
            log("restored {} from FETCH_HEAD ({} bytes)".format(SNAPSHOT_FILE, len(out)))
        else:
            log("WARN: cannot restore {} — leaving as-is".format(SNAPSHOT_FILE))

        # 4) 校验关键修复仍在（防止 gen_html.py 被回退导致修复丢失）
        #    MUST-PRESENT = 新版特征串；MUST-ABSENT = 已被移除的旧逻辑（出现即代码是旧的）
        must_present = (("AGENT_BY_PORT", "Agent/OP column"),
                        ("berthedCalls", "Berth Rate fix"),
                        ("isDecommissioned", "decommissioned-vessel filter"),
                        ("R_CODE_MAP", "R-code semantic map"),
                        ("portclosure", "Port Closure category"),
                        ("buildVesselLaneRank", "Lane-first sort"),
                        ("defaultFullSort", "Default full sort"))
        must_absent  = (("reasoncode", "legacy reasoncode class removed"),)
        for token, label in must_present:
            if token in local_html:
                log("  OK  {}".format(label))
            else:
                raise RuntimeError("{} missing after regeneration -- aborting to avoid shipping a broken page".format(label))
        for token, label in must_absent:
            if token not in local_html:
                log("  OK  {}".format(label))
            else:
                raise RuntimeError("{} still present in regenerated HTML -- gen_html.py is stale, aborting".format(label))

        # 5) 提交并推送
        #    ⚠️ gen_html.py 必须一起提交！否则远端仓库里的源程序永远是 culadmin 的旧版本，
        #    HTML 只是"产物"被替换、源码没跟上 → culadmin 下次 pull 拿到的还是旧代码，
        #    用它重生成又会覆盖掉修复（2026-09-03 定位到的真正根因）。
        run("git add {} gen_html.py".format(HTML))
        rc, out = run("git diff --cached --quiet", check=False)
        if rc == 0:
            log("no changes to commit")
            return 0
        run('git commit -m "resync: latest culadmin data + latest gen_html.py code (data-only resync, no logic regression)"')
        # 推送前先确认工作区是远端的直系后代（fast-forward 可推）。
        # 刻意不用 git rebase：本 sandbox 环境下 rebase 会损坏 .git/refs 与 objects
        # （已复现 2 次，只能重新 clone 修复）。非直系时报错让人手工处理。
        run("git fetch origin")
        rc, _ = run("git merge-base --is-ancestor FETCH_HEAD HEAD", check=False)
        if rc != 0:
            raise RuntimeError(
                "local branch is NOT a descendant of remote ({}). ".format(HTML) +
                "Run manually: git reset --hard FETCH_HEAD, re-apply your gen_html.py edits, then re-run this script.")
        run("git push origin main")
        log("pushed")
        return 0

    except Exception as e:
        log("FAILED: {}".format(e))
        return 1


if __name__ == "__main__":
    sys.exit(main())
