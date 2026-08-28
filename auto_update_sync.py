#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
auto_update_sync.py —— culadmin 端“先同步代码、再生成、再推送”的安全更新封装。

背景 / 为什么要这个脚本
------------------------
culadmin 的 `auto_update.py` 只提交 `cul_daily_movement.html` 和 `maint_snapshot.json`，
从不更新 `gen_html.py`。导致：
  * 本地 `gen_html.py` 长期停留在旧版本（缺 Agent/OP 列、Berth Rate 旧口径等修复）；
  * 每晚用旧代码重新生成 HTML 并推送，**把其他人推上去的修复直接覆盖掉**。

本脚本在生成之前强制把工作树同步到 origin/main，保证“生成 HTML 的代码”永远是最新版。

用法
----
  python auto_update_sync.py            # 同步 -> 调用 auto_update.py -> 推送
  python auto_update_sync.py --check    # 只做环境/同步自检，不生成不推送

约定
----
  * 与 `auto_update.py` 放在同一目录（Claw-Report 仓库根目录）。
  * 只保留仓库内被版本管理的文件为权威；本地对 .py/.bat 的手改会被 reset 掉（见 SYNC 策略）。
"""

import os
import subprocess
import sys

REPO = os.path.dirname(os.path.abspath(__file__))

# 需要保护的“数据/产物”文件：同步代码时不动它们（它们由脚本自己重新生成）
DATA_ARTIFACTS = (
    "cul_daily_movement.html",
    "maint_snapshot.json",
    "vessel.csv",
    ".cache",
)


def run(cmd, check=True, capture=True):
    """执行命令，返回 (returncode, stdout+stderr)。"""
    p = subprocess.run(
        cmd, cwd=REPO, shell=True,
        capture_output=capture, text=True, encoding="utf-8", errors="replace",
    )
    out = (p.stdout or "") + (p.stderr or "")
    if check and p.returncode != 0:
        raise RuntimeError("Command failed: {}\n{}".format(cmd, out))
    return p.returncode, out


def log(msg):
    print("[sync] {}".format(msg), flush=True)


def step_sync():
    """把工作树同步到 origin/main。核心：本地 gen_html.py 等脚本以远端为准。"""
    log("fetch origin ...")
    run("git fetch origin")

    # 检查是否有本地未提交改动（可能是手改的脚本）
    rc, out = run("git status --porcelain", check=False)
    dirty = [l for l in out.splitlines() if l.strip()]
    if dirty:
        log("local changes detected ({} files)".format(len(dirty)))
        for l in dirty[:20]:
            log("  {}".format(l))
        # 脚本类文件：丢弃本地改动，强制与远端一致（防止旧代码复活）
        # 数据类产物：保留，稍后由脚本重新生成
        run("git checkout -- gen_html.py")
        log("restored gen_html.py from remote (local edits discarded)")
        # 其他 .py/.bat 也一并还原，避免旧逻辑残留
        rc2, out2 = run("git ls-files '*.py' '*.bat'", check=False)
        tracked = [f.strip() for f in out2.splitlines()
                   if f.strip() and f.strip() not in DATA_ARTIFACTS]
        if tracked:
            run("git checkout -- " + " ".join('"{}"'.format(f) for f in tracked))
            log("restored {} tracked scripts from remote".format(len(tracked)))

    # fast-forward 到远端最新
    log("reset working tree to origin/main (fast-forward) ...")
    run("git reset --hard origin/main")
    rc, out = run("git rev-parse HEAD", check=False)
    log("now at {}".format(out.strip()))

    # 确认关键文件已包含最新修复
    verify_fixes()


def verify_fixes():
    """校验 gen_html.py 是否已含关键修复，防止同步失败却静默继续。"""
    try:
        with open(os.path.join(REPO, "gen_html.py"), "r", encoding="utf-8") as f:
            src = f.read()
    except IOError as e:
        raise RuntimeError("cannot read gen_html.py: {}".format(e))

    checks = {
        "Agent/OP column":        "AGENT_BY_PORT" in src,
        "Berth Rate (threshold)": "berthedCalls" in src,
    }
    ok = True
    for name, present in checks.items():
        if present:
            log("  OK  {}".format(name))
        else:
            log("  MISSING  {} -- gen_html.py may be outdated!".format(name))
            ok = False
    if not ok:
        raise RuntimeError("gen_html.py missing expected fixes; abort to avoid overwriting the dashboard with a stale build")


def step_run_auto_update():
    """调用原有的 auto_update.py（保持原有业务逻辑不变）。"""
    if not os.path.exists(os.path.join(REPO, "auto_update.py")):
        raise RuntimeError(
            "auto_update.py not found next to this script.\n"
            "Place auto_update_sync.py in the same folder as auto_update.py."
        )
    log("running auto_update.py ...")
    rc, out = run("python auto_update.py", check=False)
    print(out)
    if rc != 0:
        raise RuntimeError("auto_update.py exited with code {}".format(rc))
    log("auto_update.py finished")


def step_push():
    """提交并推送产物（只提交数据/产物，不提交代码）。"""
    run("git add cul_daily_movement.html maint_snapshot.json")
    rc, out = run("git diff --cached --quiet", check=False)
    if rc == 0:
        log("no changes to commit")
        return
    run('git commit -m "daily update (auto_update_sync): regenerated with latest gen_html.py"')
    # 推送前再拉一次，避免 non-fast-forward
    run("git fetch origin")
    run("git rebase origin/main")
    run("git push origin main")
    log("pushed to origin/main")


def main():
    check_only = "--check" in sys.argv
    try:
        step_sync()
        if check_only:
            log("check mode: sync + verify passed, skipping generation")
            return 0
        step_run_auto_update()
        step_push()
        log("DONE")
        return 0
    except Exception as e:
        log("FAILED: {}".format(e))
        # 不静默失败：非 0 退出，便于计划任务/批处理捕获
        return 1


if __name__ == "__main__":
    sys.exit(main())
