#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
git_safe_push.py —— Claw-Report 仓库统一安全提交推送器（所有自动化共用）

背景(2026-08-27): 多个自动化(daily movement / ROB / Bunker / Alphaliner / 手动会话)
共用 main 分支。并行提交后 `git push` 被拒(非快进), 本地提交被搁浅; 下一个任务
reset 时这些提交被移出 main(仅存于 backup 分支) —— 表现为"自动化失败 / 成果丢失"。

本脚本协议(幂等可重入, 绝不 force push):
  1. 快照提交: 先把【全部已跟踪改动】(可另加指定路径) commit, 绝不留未提交改动,
     保证后续任何 reset 都不会丢数据;
  2. 权威远端: git ls-remote 取 main 的真实 SHA(本地 origin/main 引用不可信);
  3. 备份: 每轮把 HEAD 备份到 backup/safe-push-prev, 任何时候可恢复;
  4. 同步: 本地领先→直接快进推送; 远端领先→前向 reset; 分叉→rebase 到远端;
     rebase 冲突(罕见, 各自动化改不同文件)→abort→reset 到远端→从备份检出
     本会话产物文件→重新提交再推(生成物以本会话最新为准);
  5. 推送: token 认证(读 .update_token 或环境变量 GITHUB_TOKEN, 绝不硬编码入库);
     被其它自动化抢推则回到 2, 最多 3 轮;
  6. 收尾: 还原 remote URL(防 token 泄漏), 打印中文结果。

用法:
  python git_safe_push.py -m "提交信息"                      # 提交全部已跟踪改动并安全推送
  python git_safe_push.py -m "提交信息" -- reports/ xx.html  # 额外纳入指定路径(新文件)
  python git_safe_push.py --no-commit                        # 不新提交, 只同步并推送已有本地提交
"""
import os, sys, subprocess, argparse

REPO = os.path.dirname(os.path.abspath(__file__))
REMOTE_URL = "https://github.com/LeahLiuL/Claw-Report.git"
TOKEN_FILE = os.path.join(REPO, ".update_token")
BRANCH = "main"
BACKUP = "backup/safe-push-prev"
MAX_ROUNDS = 3


def sh(args, check=True):
    """运行 git 命令并打印输出; check=True 时失败抛异常。"""
    r = subprocess.run(["git"] + args, cwd=REPO, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    print("  >> git " + " ".join(args))
    out = (r.stdout or "") + (r.stderr or "")
    for ln in out.splitlines():
        if ln.strip():
            print("     " + ln)
    if check and r.returncode != 0:
        raise RuntimeError("git 命令失败: git " + " ".join(args))
    return r


def get_token():
    tok = os.environ.get("GITHUB_TOKEN", "").strip()
    if not tok and os.path.exists(TOKEN_FILE):
        try:
            tok = open(TOKEN_FILE, encoding="utf-8").read().strip()
        except Exception:
            tok = ""
    if not tok:
        raise RuntimeError("未找到 GitHub token(.update_token 或环境变量 GITHUB_TOKEN)")
    return tok


def ls_remote_sha():
    r = sh(["ls-remote", "origin", "refs/heads/" + BRANCH], check=False)
    line = (r.stdout or "").strip()
    if r.returncode != 0 or not line:
        raise RuntimeError("ls-remote 失败(网络/远端不可达)")
    return line.split()[0]


def is_ancestor(a, b):
    return subprocess.run(["git", "merge-base", "--is-ancestor", a, b],
                          cwd=REPO, capture_output=True).returncode == 0


def push_with_token(token):
    sh(["remote", "set-url", "origin",
        REMOTE_URL.replace("https://", "https://" + token + "@")])
    try:
        r = sh(["push", "origin", BRANCH], check=False)
        return r.returncode == 0
    finally:
        sh(["remote", "set-url", "origin", REMOTE_URL], check=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-m", "--message", default="wip: safe-push 快照(自动化产物)")
    ap.add_argument("--no-commit", action="store_true",
                    help="不新提交, 只同步并推送已有本地提交")
    ap.add_argument("paths", nargs="*", default=[],
                    help="额外纳入提交的路径(新文件等); 已跟踪改动总会被提交")
    args = ap.parse_args()

    print("== git_safe_push 开始 ==")

    # 1) 快照提交(绝不留未提交的已跟踪改动)
    our_files = []
    if not args.no_commit:
        sh(["add", "-u"])
        if args.paths:
            sh(["add", "--"] + args.paths)
        st = sh(["diff", "--cached", "--name-only"]).stdout.strip()
        if st:
            our_files = [l.strip() for l in st.splitlines() if l.strip()]
            sh(["commit", "-m", args.message])
            print("  已快照提交 %d 个文件: %s" % (len(our_files), ", ".join(our_files[:8])))
        else:
            print("  无待提交改动。")

    token = get_token()

    for rnd in range(1, MAX_ROUNDS + 1):
        print("== 第 %d/%d 轮同步 ==" % (rnd, MAX_ROUNDS))
        server = ls_remote_sha()
        head = sh(["rev-parse", "HEAD"]).stdout.strip()
        sh(["branch", "-f", BACKUP, "HEAD"], check=False)

        if head == server:
            print("  本地与远端一致(%s), 无需推送。" % head[:8])
            print("== git_safe_push 成功 ==")
            return 0

        if is_ancestor(head, server):
            # 远端领先且本地无独有提交 → 前向同步(不丢任何本地提交)
            sh(["fetch", "origin", BRANCH])
            sh(["reset", "--hard", server])
            print("  已前向同步到远端 %s。" % server[:8])
            continue

        if is_ancestor(server, head):
            print("  本地领先, 快进推送。")
        else:
            # 分叉: 本地提交重放到远端之上
            print("  检测到分叉, fetch + rebase 到远端 %s ..." % server[:8])
            sh(["fetch", "origin", BRANCH])
            r = sh(["rebase", server], check=False)
            if r.returncode != 0:
                print("  rebase 冲突, 降级: reset 到远端 + 从备份恢复本会话产物 + 重新提交")
                sh(["rebase", "--abort"], check=False)
                sh(["reset", "--hard", server])
                if our_files:
                    sh(["checkout", BACKUP, "--"] + our_files, check=False)
                    sh(["add", "--"] + our_files)
                    st2 = sh(["diff", "--cached", "--name-only"]).stdout.strip()
                    if st2:
                        sh(["commit", "-m", args.message])
                        print("  产物已按本会话最新版本重新提交。")
                else:
                    print("  无产物文件需恢复(纯同步), 已对齐远端。")

        if push_with_token(token):
            print("  推送成功。")
            print("== git_safe_push 成功 ==")
            return 0
        print("  推送被拒(其它自动化抢先), 重试 ...")

    print("== git_safe_push 失败: %d 轮后仍未推上(网络或持续抢占)。" % MAX_ROUNDS)
    print("   本地提交已备份在分支 %s, 数据未丢失, 下次运行会自动重试 ==" % BACKUP)
    return 1


if __name__ == "__main__":
    sys.exit(main())
