#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
sync_templates.py —— 渲染前把远端 templates/ 同步到工作树（只碰模板，绝不碰数据）

== 为什么需要它 ==
rob_update.bat / rob_fast_update.bat 在渲染前只跑 `git fetch`：
FETCH_HEAD 更新了，但工作树不动（脚本注释里写明了"生成前只 fetch 不动工作树，
避免与其它自动化抢工作树"）。工作树真正对齐要等本轮末尾 git_safe_push.py 的
reset --hard。于是每次渲染用的都是"上一轮结束时"同步到的那版模板 —— 代码
改动永远滞后一班。本脚本补上这一步：fetch 之后、render 之前，只把 templates/
对齐到 FETCH_HEAD。

== 安全边界（写死在代码里，不依赖调用方）==
  1. 只允许写 templates/ 下的文件（扩展名白名单），外加 ALLOW_ROOT_CODE 里显式列出的
     根目录代码脚本（默认只有 voyage_consumption.py —— 航次油耗算法）。
     任何其它路径（rob_data/、*.csv、*.json、*.jsonl、*.xlsx、*.db ...）一律拒写。
     ⚠️ 为什么允许这一个脚本：rob_refresh.py 在 [3/5] 才跑，它 import 的
     voyage_consumption 若还是上一轮的旧版，页面算出来的航次油耗就会滞后一班
     （2026-11-12 期初口径修正后实测：新模板 + 旧算法 = 页面数字会被下一轮覆盖回去）。
  2. 先把所有 blob 读进内存并校验（非空 / 大小 / HTML 结构完整），
     全部通过才开始落盘；任一不通过就跳过该文件，绝不半截写入。
  3. 落盘前把当前文件备份到 rob_data/tpl_backup/<时间戳>/（.gitignore 已排除，
     不会进 git，也不会被 git add -u 带走）。
  4. 原子替换：写 <file>.tplnew 再 os.replace()，避免中断留下半个文件。
  5. 任一写入异常 → 从备份回滚已写的文件，然后 exit 0。
  6. 内容相同（忽略 CRLF 差异）就不写，避免制造无意义的改动和假提交。
  7. 所有非致命问题一律 exit 0：沿用旧模板总比断掉整条流水线好。
     只有"确认可以写且写成功了"才会打印已更新。

用法:
    python sync_templates.py            # 正常同步
    python sync_templates.py --dry-run  # 只报告会改哪些文件, 不落盘
"""
import os
import shutil
import subprocess
import sys
import time

REPO = os.path.dirname(os.path.abspath(__file__))
TPL_DIR = "templates"
ALLOW_EXT = (".html", ".htm", ".css", ".js", ".svg", ".ico")
# 根目录代码脚本白名单（显式列表，不做 *.py 通配 —— 避免把数据/临时脚本带进来）。
ALLOW_ROOT_CODE = ("voyage_consumption.py",)
REF_CANDIDATES = ("FETCH_HEAD", "origin/main")
BACKUP_ROOT = os.path.join("rob_data", "tpl_backup")
KEEP_BACKUPS = 5
MIN_HTML_BYTES = 5000
HTML_MARKERS = ("<html", "</html>", "<script")

# 明确的拒绝清单 —— 双保险。白名单已经只放行 templates/, 这里再挡一次,
# 任何人日后放宽白名单也不会误伤数据文件。
FORBIDDEN_PREFIX = ("rob_data/", ".git/", "history/", "site/", ".workbuddy/")
FORBIDDEN_NAMES = ("rob_results.json", "rob_history.csv", "meta.json")


def log(msg):
    print(msg)
    sys.stdout.flush()


def warn(msg):
    log("[WARN] " + msg)


def info(msg):
    log("[OK] " + msg)


def git(*args):
    return subprocess.run(["git"] + list(args), cwd=REPO,
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def resolve_ref():
    """优先 FETCH_HEAD(刚 fetch 过)，回退 origin/main。返回 (ref, sha) 或 (None, None)。"""
    for ref in REF_CANDIDATES:
        r = git("rev-parse", "--verify", "--quiet", ref + "^{commit}")
        if r.returncode == 0 and r.stdout.strip():
            return ref, r.stdout.decode("utf-8", "replace").strip()
    return None, None


def is_allowed(path):
    """唯一的放行判断：templates/ 前缀（扩展名白名单）或根目录代码白名单，且不在拒绝清单。"""
    norm = path.replace("\\", "/").strip()
    while norm.startswith("./"):
        norm = norm[2:]
    low = norm.lower()
    for pre in FORBIDDEN_PREFIX:
        if low.startswith(pre):
            return False
    for name in FORBIDDEN_NAMES:
        if low == name or low.endswith("/" + name):
            return False
    if norm in ALLOW_ROOT_CODE:
        return True
    if not norm.startswith(TPL_DIR + "/"):
        return False
    if ".." in norm.split("/"):
        return False
    if not low.endswith(ALLOW_EXT):
        return False
    return True


def remote_paths(ref):
    r = git("ls-tree", "-r", "--name-only", ref, "--", TPL_DIR, *ALLOW_ROOT_CODE)
    if r.returncode != 0:
        return []
    out = r.stdout.decode("utf-8", "replace")
    return [p.strip() for p in out.splitlines() if p.strip() and is_allowed(p.strip())]


def read_blob(ref, path):
    return git("cat-file", "blob", "%s:%s" % (ref, path))


def validate(path, data):
    """返回 None 表示通过，否则返回拒绝原因。"""
    if not data:
        return "空文件"
    if path.lower().endswith(".py"):
        # 代码脚本必须语法完整，否则渲染时 import 直接崩，比不更新更糟。
        if len(data) < 2000:
            return "体积异常小 (%d bytes), 疑似被截断" % len(data)
        try:
            compile(data.decode("utf-8", "replace"), path, "exec")
        except SyntaxError as e:
            return "语法错误: line %s %s" % (e.lineno, e.msg)
        return None
    if path.lower().endswith((".html", ".htm")):
        if len(data) < MIN_HTML_BYTES:
            return "体积异常小 (%d bytes), 疑似被截断" % len(data)
        low = data.lower()
        for mk in HTML_MARKERS:
            if mk.encode("ascii") not in low:
                return "缺少结构标记 %r" % mk
    return None


def local_path(path):
    return os.path.join(REPO, path.replace("/", os.sep))


def same_content(a, b):
    return a.replace(b"\r\n", b"\n") == b.replace(b"\r\n", b"\n")


def eol_mode():
    """仓库若开了 core.autocrlf=true，工作树必须是 CRLF，否则会留下 stat-dirty。"""
    r = git("config", "--get", "core.autocrlf")
    val = r.stdout.decode("utf-8", "replace").strip().lower()
    return "crlf" if val == "true" else "asis"


def to_worktree_eol(data, mode, old):
    """按工作树原有行尾写回，避免内容相同却显示 modified。

    git cat-file 拿到的是仓库里的原始字节（LF）。core.autocrlf=true 的仓库
    checkout 出来是 CRLF，若直接写 LF，文件内容等价但 git 会报 stat-dirty；
    随后的 `git add -u && git commit` 会因"没有实际改动"失败，拖垮 safe push。
    """
    if mode != "crlf":
        return data
    if b"\r\n" in data:
        return data
    if old is not None and b"\r\n" not in old:
        return data          # 本地本来就是 LF，别来回翻转
    return data.replace(b"\n", b"\r\n")


def prune_backups(keep_dir):
    """只保留最近 KEEP_BACKUPS 份备份目录。"""
    root = os.path.join(REPO, BACKUP_ROOT)
    if not os.path.isdir(root):
        return
    dirs = []
    for name in os.listdir(root):
        full = os.path.join(root, name)
        if os.path.isdir(full):
            dirs.append(full)
    dirs.sort(reverse=True)
    for old in dirs[KEEP_BACKUPS:]:
        shutil.rmtree(old, ignore_errors=True)


def main():
    dry = "--dry-run" in sys.argv

    ref, sha = resolve_ref()
    if not ref:
        warn("找不到 FETCH_HEAD / origin/main（fetch 可能失败），保持现有模板不动")
        return 0

    paths = remote_paths(ref)
    if not paths:
        warn("%s 里没有 templates/ 内容，保持现有模板不动" % ref)
        return 0

    # ---- 阶段一：全部读入 + 校验，一个不过就跳过该文件 ----
    plan = []
    for p in paths:
        r = read_blob(ref, p)
        if r.returncode != 0:
            warn("读取 %s:%s 失败，跳过（保留现有文件）" % (ref, p))
            continue
        data = r.stdout
        err = validate(p, data)
        if err:
            warn("%s 校验未通过（%s），跳过" % (p, err))
            continue
        plan.append((p, data))

    if not plan:
        warn("没有通过校验的模板，保持现状")
        return 0

    # ---- 阶段二：区分"已一致"和"需要更新" ----
    changed, same = [], []
    for p, data in plan:
        lp = local_path(p)
        old = None
        if os.path.exists(lp):
            try:
                with open(lp, "rb") as f:
                    old = f.read()
            except Exception as e:
                warn("读取本地 %s 失败（%s），跳过该文件" % (p, e))
                continue
        if old is not None and same_content(old, data):
            same.append(p)
        else:
            changed.append((p, data, old))

    if not changed:
        info("模板已是最新（%d 个文件一致），无需改动" % len(same))
        return 0

    if dry:
        info("[dry-run] 远端 %s %s 将更新:" % (ref, sha[:8]))
        for p, data, old in changed:
            info("  %s  %s -> %d bytes" % (p, "new" if old is None else "%d bytes" % len(old), len(data)))
        return 0

    # ---- 阶段三：备份（失败就整体放弃，绝不裸写）----
    stamp = time.strftime("%Y%m%d-%H%M%S")
    bak_dir = os.path.join(REPO, BACKUP_ROOT, stamp)
    backups = {}
    try:
        for p, _data, old in changed:
            if old is None:
                continue
            dst = os.path.join(bak_dir, p.replace("/", os.sep))
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            with open(dst, "wb") as f:
                f.write(old)
            backups[p] = dst
    except Exception as e:
        warn("备份失败（%s），为安全起见不写任何文件" % e)
        return 0

    # ---- 阶段四：原子落盘，失败即回滚 ----
    mode = eol_mode()
    written = []
    try:
        for p, data, old in changed:
            lp = local_path(p)
            os.makedirs(os.path.dirname(lp), exist_ok=True)
            payload = to_worktree_eol(data, mode, old)
            tmp = lp + ".tplnew"
            with open(tmp, "wb") as f:
                f.write(payload)
            os.replace(tmp, lp)
            written.append(p)
    except Exception as e:
        warn("写入失败（%s），回滚已写文件" % e)
        for p in written:
            src = backups.get(p)
            lp = local_path(p)
            if src and os.path.exists(src):
                try:
                    shutil.copy2(src, lp)
                    warn("  已回滚 %s" % p)
                except Exception as e2:
                    warn("  回滚 %s 失败（%s）" % (p, e2))
            elif not src:
                try:
                    os.remove(lp)
                    warn("  已移除新建的 %s" % p)
                except Exception:
                    pass
        return 0

    prune_backups(bak_dir)
    # 刷新 stat 缓存：写入的 mtime 变了，git 首次 status 可能误报 modified
    # （内容其实一致）。这一步只更新索引缓存，不改任何文件内容。
    git("update-index", "-q", "--refresh")
    for p, data, old in changed:
        info("已更新 %s  %s -> %d bytes" % (p, "new" if old is None else "%d bytes" % len(old), len(data)))
    info("来源 %s %s | 行尾 %s | 备份 rob_data/tpl_backup/%s"
         % (ref, sha[:8], "CRLF" if mode == "crlf" else "原样", stamp))
    return 0


if __name__ == "__main__":
    sys.exit(main())
