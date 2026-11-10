# -*- coding: utf-8 -*-
"""Outlook 健康检查与自愈 —— 跑任何 COM 抓取之前先过这一关。

## 为什么需要
Outlook 连续运行数天后会耗尽资源(实测 8 天：句柄 21,959 / 内存 417MB),
此后 `Items.Restrict` 抛 "内存或系统资源不足", 联机存档 `Items.Count` 直接从
66,023 变成 0, 而 `NS.Offline=False`、`ExchangeConnectionMode=700` 仍然显示正常。
结果: 抓取脚本"成功"退出, 但一封报告都没抓到, 网页继续显示旧数据 —— 这种
静默失败比直接报错危险得多。

本模块在抓取前做功能性探针(不是只看 Offline 标志), 发现异常就杀掉重启
OUTLOOK.EXE, 恢复不了就明确失败, 让上层跳过本轮而不是产出空数据。

## 用法
    import outlook_guard
    outlook_guard.ensure_ready()          # 恢复不了抛 RuntimeError

命令行:
    python outlook_guard.py --check       # 只体检, 不改动
    python outlook_guard.py --ensure      # 体检 + 必要时重启(默认)
"""
import argparse
import datetime
import os
import re
import subprocess
import sys
import time

# 预防性重启阈值: 实测崩溃发生在句柄 ~22k / 内存 ~420MB, 留足余量
HANDLE_SOFT_LIMIT = 15000
MEM_SOFT_LIMIT_MB = 900
# 判定"已加载到可用状态"的门槛
READY_MEM_MB = 150
READY_HANDLES = 500

LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "rob_data", "outlook_guard.log")


def _log(msg):
    line = "[%s] %s" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line)
    try:
        d = os.path.dirname(LOG_PATH)
        if d and not os.path.isdir(d):
            os.makedirs(d)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------- 进程状态

def outlook_proc():
    """返回 {'pid','mem_mb','handles'} 或 None(未运行)。"""
    try:
        import psutil  # 有就用, 快
        for p in psutil.process_iter(["name"]):
            if (p.info.get("name") or "").upper() == "OUTLOOK.EXE":
                with p.oneshot():
                    return {"pid": p.pid,
                            "mem_mb": int(p.memory_info().rss / 1048576),
                            "handles": p.num_handles()}
        return None
    except Exception:
        pass
    # 回退: PowerShell(冷启动约 1-2s, 可接受)
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "Get-Process OUTLOOK -ErrorAction SilentlyContinue | "
             "Select-Object -First 1 | ForEach-Object { "
             "\"$($_.Id);$($_.WorkingSet64);$($_.HandleCount)\" }"],
            stderr=subprocess.DEVNULL, timeout=40).decode(errors="replace").strip()
        if not out:
            return None
        parts = out.split(";")
        if len(parts) >= 3:
            return {"pid": int(parts[0]),
                    "mem_mb": int(int(parts[1]) / 1048576),
                    "handles": int(parts[2])}
    except Exception:
        pass
    return None


def _find_outlook_exe():
    """找 OUTLOOK.EXE 全路径。"""
    try:
        import winreg
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for sub in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",):
                try:
                    with winreg.OpenKey(root, sub) as k:
                        path = winreg.QueryValue(k, None)
                        if path and os.path.isfile(path):
                            return path
                except OSError:
                    continue
    except Exception:
        pass
    for base in (r"C:\Program Files\Microsoft Office\root",
                 r"C:\Program Files (x86)\Microsoft Office\root",
                 r"C:\Program Files\Microsoft Office",
                 r"C:\Program Files (x86)\Microsoft Office"):
        for sub in ("Office16", "Office15", ""):
            p = os.path.join(base, sub, "OUTLOOK.EXE") if sub else os.path.join(base, "OUTLOOK.EXE")
            if os.path.isfile(p):
                return p
    return None


# ---------------------------------------------------------------- 连接与探针

def connect(retries=3, delay=5):
    """连上已运行的 Outlook 实例。优先 GetActiveObject, 避免冷启动挂死。"""
    import win32com.client
    last = None
    for i in range(retries):
        try:
            ol = win32com.client.GetActiveObject("Outlook.Application")
            return ol.GetNamespace("MAPI")
        except Exception as e:
            last = e
            # 退一步: Outlook 在运行但未注册到 ROT 时, Dispatch 会复用现有实例
            try:
                ol = win32com.client.Dispatch("Outlook.Application")
                return ol.GetNamespace("MAPI")
            except Exception as e2:
                last = e2
            if i < retries - 1:
                time.sleep(delay)
    raise RuntimeError("无法连接 Outlook: %s" % last)


def probe(ns, days=7):
    """功能性探针。返回 (ok, 描述)。

    刻意不看 Items.Count(联机存档上会永久挂死), 只用带时间窗的 Restrict,
    并且实际读一条记录, 确保属性可访问。
    """
    issues = []

    try:
        if ns.Offline:
            issues.append("Outlook 处于离线模式(读到的都是本地缓存)")
    except Exception as e:
        return False, "NS.Offline 读取失败: %s" % str(e)[:80]

    try:
        mode = ns.ExchangeConnectionMode
        # 0=无Exchange 100=离线 200=缓存离线 300=断开 400+=已连接
        if mode is not None and mode < 400:
            issues.append("Exchange 连接模式=%s(离线/断开)" % mode)
    except Exception:
        pass

    # 主邮箱收件箱做时间窗查询 —— 这是最贴近实际抓取的操作
    n = None
    try:
        inbox = ns.GetDefaultFolder(6)
        cut = (datetime.datetime.now() - datetime.timedelta(days=days)
               ).strftime("%m/%d/%Y %I:%M %p")
        items = inbox.Items.Restrict("[ReceivedTime] >= '%s'" % cut)
        n = items.Count
        if n > 0:
            it = items.GetFirst()          # 真正读一次属性
            _ = (it.Subject, it.ReceivedTime, it.Class)
    except Exception as e:
        msg = str(e)
        if "内存" in msg or "资源" in msg or "-2147024882" in msg:
            return False, "收件箱查询报资源不足(Outlook 已耗尽): %s" % msg[:80]
        return False, "收件箱查询失败: %s" % msg[:100]

    detail = "近%d天主收件箱 %s 封" % (days, n)
    if issues:
        return False, "; ".join(issues) + " | " + detail
    return True, detail


# ---------------------------------------------------------------- 重启

def _launch_detached(exe):
    """以脱离父进程的方式启动 OUTLOOK.EXE。

    坑: 直接 Popen 出来的 Outlook 是调用方进程树的子进程, 脚本一退出(或计划任务
    一结束)就被连带回收 —— 实测 16:25:47 探针还健康, 16:26:00 就"未运行"了。
    必须 DETACHED_PROCESS + CREATE_BREAKAWAY_FROM_JOB 才能真正活下来。
    """
    flags = (0x00000008 |          # DETACHED_PROCESS: 不继承控制台
             0x00000200 |          # CREATE_NEW_PROCESS_GROUP
             0x01000000)           # CREATE_BREAKAWAY_FROM_JOB: 脱离作业对象
    try:
        return subprocess.Popen([exe], close_fds=True, creationflags=flags,
                                stdin=subprocess.DEVNULL,
                                stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL)
    except OSError:
        # BREAKAWAY 被作业对象禁止时会失败, 退一步只用 DETACHED + NEW_GROUP
        return subprocess.Popen([exe], close_fds=True,
                                creationflags=0x00000008 | 0x00000200,
                                stdin=subprocess.DEVNULL,
                                stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL)


def restart(timeout=240):
    """杀掉并重启 Outlook, 等到真正可用。返回 (ok, 描述)。"""
    p = outlook_proc()
    if p:
        _log("结束 OUTLOOK.EXE (pid=%s, %dMB, 句柄 %d)"
             % (p["pid"], p["mem_mb"], p["handles"]))
    try:
        subprocess.run(["taskkill", "/F", "/IM", "OUTLOOK.EXE"],
                       capture_output=True, timeout=60)
    except Exception as e:
        _log("taskkill 失败: %s" % e)
    time.sleep(4)

    exe = _find_outlook_exe()
    if not exe:
        return False, "找不到 OUTLOOK.EXE, 无法自动重启"
    _log("启动 %s" % exe)
    try:
        _launch_detached(exe)
    except Exception as e:
        return False, "启动失败: %s" % e

    deadline = time.time() + timeout
    last = "未就绪"
    while time.time() < deadline:
        time.sleep(6)
        p = outlook_proc()
        if not p:
            last = "进程未出现"
            continue
        if p["mem_mb"] < READY_MEM_MB or p["handles"] < READY_HANDLES:
            last = "加载中(%dMB/句柄%d)" % (p["mem_mb"], p["handles"])
            continue
        try:
            ns = connect(retries=1)
            ok, detail = probe(ns)
            if ok:
                return True, "重启后恢复: " + detail
            last = detail
        except Exception as e:
            last = str(e)[:80]
    return False, "重启后 %.0fs 仍未可用: %s" % (timeout, last)


# ---------------------------------------------------------------- 对外接口

def ensure_ready(max_restarts=1, quiet=False):
    """确保 Outlook 处于可抓取状态。恢复不了抛 RuntimeError。"""
    def say(m):
        if not quiet:
            _log(m)

    p = outlook_proc()
    if not p:
        say("Outlook 未运行, 尝试启动")
        ok, detail = restart()
        if not ok:
            raise RuntimeError(detail)
        say(detail)
        return True

    # 预防性重启: 还没坏但已经接近崩溃阈值, 趁现在重启比半路崩掉好
    if p["handles"] > HANDLE_SOFT_LIMIT or p["mem_mb"] > MEM_SOFT_LIMIT_MB:
        say("Outlook 资源偏高(%dMB / 句柄 %d), 超过预防阈值(%dMB / %d), 提前重启"
            % (p["mem_mb"], p["handles"], MEM_SOFT_LIMIT_MB, HANDLE_SOFT_LIMIT))
        if max_restarts > 0:
            ok, detail = restart()
            if ok:
                say(detail)
                return True
            raise RuntimeError(detail)
        say("达到重启上限, 继续用当前实例(可能失败)")

    used = 0
    while True:
        try:
            ns = connect()
            ok, detail = probe(ns)
        except Exception as e:
            ok, detail = False, str(e)[:120]
        if ok:
            say("Outlook 健康: " + detail)
            return True
        if used >= max_restarts:
            raise RuntimeError("Outlook 不健康且重启次数用尽: " + detail)
        say("Outlook 异常(%s), 重启第 %d 次" % (detail, used + 1))
        ok2, d2 = restart()
        used += 1
        if not ok2:
            raise RuntimeError(d2)
        say(d2)
        return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="只体检, 不改动")
    ap.add_argument("--ensure", action="store_true", help="体检并在必要时重启(默认)")
    ap.add_argument("--max-restarts", type=int, default=1)
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    p = outlook_proc()
    if p:
        print("进程: pid=%s %dMB 句柄=%d" % (p["pid"], p["mem_mb"], p["handles"]))
    else:
        print("进程: 未运行")

    if args.check:
        if not p:
            print("结论: 未运行")
            return 2
        try:
            ok, detail = probe(connect())
        except Exception as e:
            print("结论: 连不上 (%s)" % str(e)[:100])
            return 2
        print("结论: %s | %s" % ("健康" if ok else "异常", detail))
        return 0 if ok else 2

    try:
        ensure_ready(max_restarts=args.max_restarts, quiet=args.quiet)
        return 0
    except Exception as e:
        _log("[FAIL] %s" % e)
        return 1


if __name__ == "__main__":
    sys.exit(main())
