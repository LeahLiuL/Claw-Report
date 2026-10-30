#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""git_op_lock.py —— Claw-Report 仓库 git 操作互斥锁（防并发损坏 .git）

背景
----
本机有 7 个计划任务 + 手动会话并发操作同一个仓库（ROB 01:00/13:00、
Daily Movement 每 2 小时、Bunker 08:00、Bapfile 23:00、Vessel Departure 14:00），
2026-09-22 / 09-24 两次 .git 对象库损坏（对象凭空消失，只能重克隆修复），
高度疑似后台 git gc 与 pull/rebase 并发所致。所有做 git 写操作的自动化
**必须**先持有本锁。

协议
----
  * 锁文件: 仓库根目录 .git_op.lock，内容 = 持有者名（如 rob_update）
  * acquire: 不存在/已过期(STALE_MIN) → 原子创建(O_CREAT|O_EXCL)；
             已存在且 owner 相同 → 可重入返回(外层持有者继续用)；
             已存在且 owner 不同 → 轮询等待直到 --wait 超时
  * release : 仅当锁内容 = 自己的 owner 时删除（防止误删他人锁）
  * 持有进程崩溃不清锁 → 靠 STALE_MIN 自动过期（30 分钟 > 最长一轮全量刷新）
  * 配套: 仓库已设 `git config gc.auto 0` 停用后台自动 gc

CLI 用法（rob_update.bat 等）
----
  python git_op_lock.py acquire --owner rob_update --wait 300   # 退出码 0=获得/可重入 1=超时
  python git_op_lock.py release --owner rob_update              # 退出码恒 0
  python git_op_lock.py status                                   # 查看当前持有者/年龄

Python 用法（git_safe_push.py 等）
----
  import git_op_lock
  created = git_op_lock.acquire("git_safe_push", wait_s=180)  # True=新建的锁(退出时要释放)
  git_op_lock.release("git_safe_push")                        # 幂等; 非本人锁不动
  # 建议: import atexit; atexit.register(git_op_lock.release, owner)  防多 return 漏释放
"""
import os
import sys
import time
import argparse

REPO = os.path.dirname(os.path.abspath(__file__))
LOCK = os.path.join(REPO, ".git_op.lock")
STALE_MIN = 30      # 锁过期分钟数(> 最长一轮 ROB 全量刷新 ~15min)
POLL_S = 5           # 等待期间轮询间隔


def _age_min():
    """锁文件年龄(分钟); 不存在返回 None。"""
    try:
        return (time.time() - os.path.getmtime(LOCK)) / 60.0
    except OSError:
        return None


def _owner():
    """锁内容原文; 不存在返回 None。"""
    try:
        with open(LOCK, encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return None


def _parse(content):
    """解析锁内容 -> (owner, pid, ts)。

    新格式 'owner|pid|HH:MM'（2026-10-30 起：rob_fast 死进程持锁 30 分钟、
    daily_movement 全量更新等 5 分钟超时的陈旧死锁事故后引入）；
    老格式（纯 owner 名）pid/ts 为 None，仍走 STALE_MIN 兜底。"""
    if not content:
        return None, None, None
    parts = content.split("|")
    owner = parts[0].strip()
    pid = None
    try:
        pid = int(parts[1])
    except (IndexError, ValueError):
        pass
    ts = parts[2].strip() if len(parts) > 2 else None
    return owner, pid, ts


def _pid_alive(pid):
    """Windows 下检测进程是否存活。

    返回 True/False；无法判断(异常/无效 pid)返回 None —— 保守不判死，走原兜底。"""
    if not pid or pid <= 0:
        return None
    try:
        import ctypes
        # PROCESS_QUERY_LIMITED_INFORMATION: 同用户进程下 OpenProcess 失败 ≈ 进程已不存在
        h = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
        if h:
            ctypes.windll.kernel32.CloseHandle(h)
            return True
        return False
    except Exception:
        return None


def status():
    o, a = _owner(), _age_min()
    if o is None:
        print("[lock] free")
        return 0
    owner, pid, ts = _parse(o)
    extra = " pid=%s started=%s" % (pid, ts) if pid else " (old format)"
    dead = "  <-- HOLDER DEAD, will be stolen in seconds" if _pid_alive(pid) is False else ""
    stale = "  <-- STALE, will be stolen" if a >= STALE_MIN else ""
    print("[lock] held by '%s' (age %.1f min%s)%s%s" % (owner, a, extra, dead, stale))
    return 0


def acquire(owner, wait_s=0):
    """获取锁。返回 True=新建(调用方负责 release)，False=可重入(勿 release)。失败抛 LockHeld。"""
    deadline = time.time() + wait_s
    while True:
        age = _age_min()
        if age is None:
            try:
                fd = os.open(LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                try:
                    os.write(fd, ("%s|%d|%s" % (owner, os.getpid(),
                          time.strftime("%H:%M"))).encode("utf-8"))
                finally:
                    os.close(fd)
                return True
            except FileExistsError:
                continue  # 刚被别人抢了, 重新检查
        if age >= STALE_MIN:
            # 过期锁: 删掉重抢(与并发抢锁者由 O_EXCL 决胜负)
            try:
                os.remove(LOCK)
            except OSError:
                pass
            continue
        cur_owner, cur_pid, _cur_ts = _parse(_owner())
        # 新格式且持有进程已死 → 陈旧死锁, 几秒内自愈(2026-10-30 rob_fast 事故)
        if cur_owner and cur_pid is not None and _pid_alive(cur_pid) is False:
            try:
                os.remove(LOCK)
            except OSError:
                pass
            continue
        if cur_owner == owner:
            return False           # 可重入: 外层同 owner 已持有
        if time.time() >= deadline:
            raise LockHeld("lock held by '%s' (age %.1f min), waited %ds" %
                           (cur_owner, age, wait_s))
        time.sleep(POLL_S)


def release(owner):
    """释放锁。仅当锁 owner 匹配才删；不存在或他人持有则不动。幂等。"""
    cur_owner, _p, _t = _parse(_owner())
    if cur_owner is None or cur_owner != owner:
        return
    try:
        os.remove(LOCK)
    except OSError:
        pass


class LockHeld(RuntimeError):
    pass


def main():
    ap = argparse.ArgumentParser(description="Claw-Report git 操作互斥锁")
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("acquire")
    a.add_argument("--owner", required=True)
    a.add_argument("--wait", type=int, default=0, help="等待秒数, 默认 0")
    sub.add_parser("status")
    r = sub.add_parser("release")
    r.add_argument("--owner", required=True)
    args = ap.parse_args()

    try:
        if args.cmd == "acquire":
            acquire(args.owner, wait_s=args.wait)
            print("[lock] acquired by '%s'" % args.owner)
            return 0
        if args.cmd == "release":
            release(args.owner)
            print("[lock] released by '%s'" % args.owner)
            return 0
        return status()
    except LockHeld as e:
        print("[lock] FAILED: %s" % e)
        return 1


if __name__ == "__main__":
    sys.exit(main())
