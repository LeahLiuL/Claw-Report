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

# 需要从远端 HTML 中"原样搬运"的数据块（行首 const 声明）
DATA_BLOCKS = ("TODAY_DATA", "SNAPSHOTS")


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
    """从远端 main 分支拿到 culadmin 上传的最新 HTML 文本。"""
    log("fetch origin ...")
    run("git fetch origin")
    rc, out = run("git show origin/main:{}".format(HTML), check=False)
    if rc != 0:
        raise RuntimeError("cannot read {} from origin/main: {}".format(HTML, out))
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

        # 2) 从远端 HTML 抽取数据块，注入到本地生成的 HTML
        changed = []
        for name in DATA_BLOCKS:
            remote_val = extract_data_block(remote_html, name)
            if remote_val is None:
                log("  skip {} (not present in remote HTML)".format(name))
                continue
            local_val = extract_data_block(local_html, name)
            if local_val is None:
                log("  skip {} (not present in local HTML)".format(name))
                continue
            if remote_val == local_val:
                log("  {} identical, nothing to do".format(name))
                continue
            local_html = replace_data_block(local_html, name, remote_val)
            changed.append(name)
            log("  {} replaced with culadmin's latest data".format(name))

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

        # 3) 校验关键修复仍在（防止 gen_html.py 被回退导致修复丢失）
        for token, label in (("AGENT_BY_PORT", "Agent/OP column"),
                             ("berthedCalls", "Berth Rate fix")):
            if token in local_html:
                log("  OK  {}".format(label))
            else:
                raise RuntimeError("{} missing after regeneration -- aborting to avoid shipping a broken page".format(label))

        # 4) 提交并推送（只推 HTML，不动代码）
        run("git add {}".format(HTML))
        rc, out = run("git diff --cached --quiet", check=False)
        if rc == 0:
            log("no changes to commit")
            return 0
        run('git commit -m "resync: latest culadmin data + latest gen_html.py code (data-only resync, no logic regression)"')
        run("git fetch origin")
        run("git rebase origin/main")
        run("git push origin main")
        log("pushed")
        return 0

    except Exception as e:
        log("FAILED: {}".format(e))
        return 1


if __name__ == "__main__":
    sys.exit(main())
