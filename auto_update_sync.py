#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
auto_update_sync.py —— culadmin 端“先同步代码、再生成、再推送”的安全更新封装。

背景 / 为什么要这个脚本
------------------------
culadmin 的 `auto_update.py` 只提交 `cul_daily_movement.html` 和 `maint_snapshot.json`，
从不 `git pull`。导致：
  * 本地 `gen_html.py` 长期停留在旧版本（缺 Agent/OP 列、R 码语义映射、Lane 排序等修复）；
  * 每次用旧代码重新生成 HTML 并推送，**把别人推上去的修复直接覆盖掉**。
    已记录回退事件：2026-09-01 `32e35d1`、2026-09-03 `b6f6a62` / `0f8d9ec`。

本脚本在生成之前强制把工作树同步到远端最新，保证“生成 HTML 的代码”永远是最新版。

用法
----
  python auto_update_sync.py            # 同步 -> 调用 auto_update.py -> 推送
  python auto_update_sync.py --check    # 只做环境/同步自检，不生成不推送

约定
----
  * 与 `auto_update.py` 放在同一目录（Claw-Report 仓库根目录）。
  * 只保留仓库内被版本管理的文件为权威；本地对 .py/.bat 的手改会被 reset 掉（见 SYNC 策略）。
  * 数据/产物文件（HTML、maint_snapshot.json、vessel.csv、.cache）不受影响，由脚本重新生成。
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

COMMIT_MSG = "daily update (auto_update_sync): regenerated with latest gen_html.py"


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


def remote_ref():
    """返回真实的远端最新 commit。

    ⚠️ 不要用 `origin/main`：在部分环境（Windows / 计划任务 / sandbox）里这个跟踪引用
    会**滞后于真实远端**，拿它做 `reset --hard` 会把仓库退回旧版并删掉大量文件。
    （2026-09-03 踩过：origin/main=deba575 而真实远端=6f9a4d2，一次 reset 删了 151 个文件。）
    `git fetch` 之后 FETCH_HEAD 永远是当前这次 fetch 拿到的真实远端 tip。
    """
    rc, out = run("git rev-parse FETCH_HEAD", check=False)
    ref = (out or "").strip()
    if rc != 0 or not ref:
        ref = "origin/main"
    return ref


def step_sync():
    """把工作树同步到远端最新。核心：本地 gen_html.py 等脚本以远端为准。"""
    log("fetch origin ...")
    run("git fetch origin")
    ref = remote_ref()

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
        rc2, out2 = run("git ls-files '*.py' '*.bat'", check=False)
        tracked = [f.strip() for f in out2.splitlines()
                   if f.strip() and f.strip() not in DATA_ARTIFACTS]
        if tracked:
            run("git checkout -- " + " ".join('"{}"'.format(f) for f in tracked))
            log("restored {} tracked scripts from remote".format(len(tracked)))

    log("reset working tree to remote tip ({}) ...".format(ref[:12]))
    run("git reset --hard {}".format(ref))
    rc, out = run("git rev-parse HEAD", check=False)
    log("now at {}".format((out or "").strip()[:12]))

    verify_fixes()


def verify_fixes():
    """校验 gen_html.py 是否已含关键修复，防止同步失败却静默继续。

    MUST-PRESENT: 新版修复的特征字符串，缺任一项说明代码是旧的。
    MUST-ABSENT : 已被移除的旧逻辑，出现任一项同样说明代码是旧的。
    （2026-09-03 更新：补上 R 码语义映射、封港分类、Lane-first 排序、reasoncode 已移除）
    """
    try:
        with open(os.path.join(REPO, "gen_html.py"), "r", encoding="utf-8") as f:
            src = f.read()
    except IOError as e:
        raise RuntimeError("cannot read gen_html.py: {}".format(e))

    must_present = {
        "Agent/OP column":        "AGENT_BY_PORT",
        "Berth Rate (threshold)": "berthedCalls",
        "Decommissioned filter":  "isDecommissioned",
        "R-code semantic map":    "R_CODE_MAP",
        "Port Closure category":  "portclosure",
        "Lane-first sort":        "buildVesselLaneRank",
        "Default full sort":      "defaultFullSort",
    }
    must_absent = {
        "legacy reasoncode class": "reasoncode",   # 9-01 已删除；出现即旧代码
    }

    ok = True
    for name, token in must_present.items():
        if token in src:
            log("  OK       {}".format(name))
        else:
            log("  MISSING  {} (token '{}') -- gen_html.py is outdated!".format(name, token))
            ok = False
    for name, token in must_absent.items():
        if token not in src:
            log("  OK       {} absent".format(name))
        else:
            log("  STALE    {} (token '{}') still present -- gen_html.py is outdated!".format(name, token))
            ok = False
    if not ok:
        raise RuntimeError(
            "gen_html.py missing expected fixes; abort to avoid overwriting "
            "the dashboard with a stale build"
        )


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


def _commit_artifacts():
    """暂存并提交产物。返回 True 表示确实产生了 commit。"""
    run("git add cul_daily_movement.html maint_snapshot.json")
    rc, out = run("git diff --cached --quiet", check=False)
    if rc == 0:
        log("no changes to commit")
        return False
    run('git commit -m "{}"'.format(COMMIT_MSG))
    return True


def step_push():
    """提交并推送产物（只提交数据/产物，不提交代码）。

    ⚠️ 不用 `git rebase`：rebase 在本项目的 Windows/sandbox 环境里多次损坏过仓库
    （丢 .git/refs 与 .git/objects，只能重新 clone 修复）。
    改用更安全的策略：推送前若发现远端有新提交，就撤销本次 commit、
    `reset --hard` 重新对齐，用最新代码**再生成一次**后重新提交。
    产物是纯生成结果，冲突没有意义 —— 重生成永远比合并正确。
    """
    if not _commit_artifacts():
        return

    for attempt in range(3):
        run("git fetch origin")
        ref = remote_ref()
        rc, out = run("git rev-list --count HEAD..{}".format(ref), check=False)
        behind = int((out or "0").strip() or 0)
        if behind == 0:
            break
        log("remote advanced by {} commit(s) -> re-align and regenerate".format(behind))
        run("git reset --hard {}".format(ref))   # 丢弃本次生成结果，回到远端最新
        step_run_auto_update()                    # 用最新代码重新生成
        if not _commit_artifacts():
            return

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
