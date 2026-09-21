"""把远端(GitHub/culadmin)的 rob_history.csv 与本地做**并集合并**, 而不是盲目覆盖。

背景:
  ROB 项目是多机协作 —— culadmin 机器上 `rob_update.bat` 每天自动抓 Outlook 并 push 到
  GitHub; 本机(leahliu)如果不 pull, 本地 rob_history.csv 会停留在旧日期, 导致
  rob_refresh.py 生成的页面"缺最近的数据"(看起来像"没记下来", 实际是本机没同步)。

合并规则(保证不丢数据):
  - 唯一键 = (date, vessel, report_time)
  - 远端有、本地无 -> 补入(远端是每日自动抓的, 更新更全)
  - 本地有、远端无 -> 保留(可能是本机刚跑出来的 MISS 记录)
  - 两边都有      -> 取远端, 但若远端该行是空值(found=0)而本地有实际油量, 取本地
  - 合并前自动备份本地文件为 rob_history.csv.bak-YYYYmmdd-HHMM

用法:
  python sync_history_from_git.py            # 合并 rob_history.csv + 同步船期/加油/吃水
  python sync_history_from_git.py --results  # 同时用远端 rob_results.json 覆盖本地
  python sync_history_from_git.py --dry-run  # 只看差异, 不写文件

同步的数据文件(culadmin 才是权威数据源, 本机一律以远端为准):
  - rob_data/rob_history.csv        并集合并(本机独有保留)
  - rob_data/rob_results.json       覆盖(当日快照, 整体重算)
  - cul_daily_movement.html         覆盖(船期; 不拉会导致航次 End 时间算错!)
  - rob_data/bunkering*.json        覆盖(加油量, 来自 Z 盘燃油添加日志)
  - rob_data/draft_history.csv      并集合并(吃水, 累加式)
"""
import os
import sys
import csv
import shutil
import subprocess
import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(BASE, "rob_data")
LOCAL_CSV = os.path.join(ROB_DIR, "rob_history.csv")
REMOTE_CSV = os.path.join(ROB_DIR, "_remote_hist.csv")
LOCAL_RES = os.path.join(ROB_DIR, "rob_results.json")
LOCAL_DM = os.path.join(BASE, "cul_daily_movement.html")
LOCAL_DRAFT = os.path.join(ROB_DIR, "draft_history.csv")
BUNKER_FILES = ["bunkering.json", "bunkering_types.json"]
DRAFT_KEY = ("date", "vessel")
FIELDS = ["date", "vessel", "code", "lane", "pic", "lsfo", "hsfo", "mgo",
          "ulsfo", "bw", "fw", "refeer", "found", "report_time"]
KEY = ("date", "vessel", "report_time")


def sh(cmd):
    r = subprocess.run(cmd, cwd=BASE, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print("[ERR]", " ".join(cmd), "\n", (r.stderr or "")[:500])
        return None
    return r.stdout


def fetch_remote_csv():
    """git fetch + 导出远端 rob_history.csv 到临时文件。"""
    print("git fetch origin main ...")
    sh(["git", "fetch", "origin", "main"])
    out = sh(["git", "show", "FETCH_HEAD:rob_data/rob_history.csv"])
    if not out:
        print("[ERR] 无法读取远端 rob_history.csv")
        return False
    with open(REMOTE_CSV, "w", encoding="utf-8", newline="") as f:
        f.write(out)
    return True


def rkey(r):
    return tuple((r.get(k) or "").strip() for k in KEY)


def has_oil(r):
    """该行是否含实际油量(非空且非 0)。"""
    for k in ("lsfo", "hsfo", "mgo", "ulsfo"):
        v = (r.get(k) or "").strip()
        if v:
            try:
                if float(v) != 0:
                    return True
            except ValueError:
                return True
    return False


def merge(dry=False):
    loc = list(csv.DictReader(open(LOCAL_CSV, encoding="utf-8")))
    rem = list(csv.DictReader(open(REMOTE_CSV, encoding="utf-8")))
    print("本地: %d 行   远端: %d 行" % (len(loc), len(rem)))

    rem_map = {}
    for r in rem:
        rem_map.setdefault(rkey(r), []).append(r)

    merged, taken = [], set()
    n_from_remote = n_keep_local = n_conflict_local = 0
    for r in loc:
        k = rkey(r)
        if k not in rem_map:
            merged.append(r)                      # 本地独有 -> 保留
            n_keep_local += 1
            continue
        o = rem_map[k][0]
        # 远端是空行而本地有油量 -> 保留本地(避免用 MISS 覆盖真实数据)
        if has_oil(r) and not has_oil(o):
            merged.append(r)
            n_conflict_local += 1
        else:
            merged.append(o)
        taken.add(k)
    for k, rows in rem_map.items():
        if k in taken:
            continue
        merged.extend(rows)                       # 远端独有 -> 补入
        n_from_remote += len(rows)

    merged.sort(key=lambda r: ((r.get("date") or ""), (r.get("vessel") or ""),
                               (r.get("report_time") or "")))
    print("远端新增补入: %d    本地独有保留: %d    冲突取本地(远端为空): %d"
          % (n_from_remote, n_keep_local, n_conflict_local))
    print("合并结果: %d 行 (本地 %d + 远端独有 %d)" % (len(merged), len(loc), n_from_remote))

    if dry:
        print("[dry-run] 未写入文件")
        return
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    bak = LOCAL_CSV + ".bak-" + stamp
    shutil.copy2(LOCAL_CSV, bak)
    print("已备份本地 ->", os.path.basename(bak))
    with open(LOCAL_CSV, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
        w.writeheader()
        for r in merged:
            w.writerow({k: (r.get(k) or "") for k in FIELDS})
    print("已写入", LOCAL_CSV)


def sync_results(dry=False):
    """rob_results.json 是"当日最新 ROB 快照", 每次刷新整体重算, 直接取远端最新版。"""
    out = sh(["git", "show", "FETCH_HEAD:rob_data/rob_results.json"])
    if not out:
        print("[WARN] 远端 rob_results.json 读取失败, 跳过")
        return
    if dry:
        print("[dry-run] 将用远端 rob_results.json 覆盖本地")
        return
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    shutil.copy2(LOCAL_RES, LOCAL_RES + ".bak-" + stamp)
    with open(LOCAL_RES, "w", encoding="utf-8", newline="") as f:
        f.write(out)
    print("已用远端 rob_results.json 覆盖本地 (旧文件已备份)")


def sync_simple(remote_path, local_path, dry=False):
    """累加式/快照式数据: 直接取远端覆盖本地(culadmin 是权威来源)。"""
    out = sh(["git", "show", "FETCH_HEAD:" + remote_path])
    if not out:
        print("[WARN] 远端 %s 读取失败, 跳过" % remote_path)
        return False
    name = os.path.basename(local_path)
    if os.path.exists(local_path):
        try:
            if open(local_path, encoding="utf-8", errors="replace").read() == out:
                print("%-28s 已是最新, 跳过" % name)
                return True
        except Exception:
            pass
    if dry:
        print("[dry-run] 将更新 %s (%d 字符)" % (name, len(out)))
        return True
    if os.path.exists(local_path):
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
        shutil.copy2(local_path, local_path + ".bak-" + stamp)
    with open(local_path, "w", encoding="utf-8", newline="") as f:
        f.write(out)
    print("%-28s 已更新 (%d 字符)" % (name, len(out)))
    return True


def merge_draft(dry=False):
    """draft_history.csv 是累加式(每天抓一次), 与本地做并集, 不能直接覆盖。"""
    out = sh(["git", "show", "FETCH_HEAD:rob_data/draft_history.csv"])
    if not out:
        print("[WARN] 远端 draft_history.csv 读取失败, 跳过")
        return
    rem = list(csv.DictReader(out.splitlines()))
    if not os.path.exists(LOCAL_DRAFT):
        if dry:
            print("[dry-run] 将写入 draft_history.csv (%d 行)" % len(rem))
            return
        with open(LOCAL_DRAFT, "w", encoding="utf-8", newline="") as f:
            f.write(out)
        print("draft_history.csv      已写入 (%d 行)" % len(rem))
        return
    loc = list(csv.DictReader(open(LOCAL_DRAFT, encoding="utf-8-sig")))
    fields = list(loc[0].keys()) if loc else list(rem[0].keys())
    m = {tuple((r.get(k) or "").strip() for k in DRAFT_KEY): r for r in loc}
    added = 0
    for r in rem:
        k = tuple((r.get(k) or "").strip() for k in DRAFT_KEY)
        if k not in m:
            m[k] = r
            added += 1
    merged = sorted(m.values(), key=lambda r: ((r.get("date") or ""), (r.get("vessel") or "")))
    print("draft_history.csv      本地 %d 行 + 远端新增 %d 行 = %d 行"
          % (len(loc), added, len(merged)))
    if dry:
        print("[dry-run] 未写入")
        return
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    shutil.copy2(LOCAL_DRAFT, LOCAL_DRAFT + ".bak-" + stamp)
    with open(LOCAL_DRAFT, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in merged:
            w.writerow({k: (r.get(k) or "") for k in fields})
    print("已写入", LOCAL_DRAFT)


if __name__ == "__main__":
    dry = "--dry-run" in sys.argv
    if not fetch_remote_csv():
        sys.exit(1)
    merge(dry=dry)
    if "--results" in sys.argv:
        sync_results(dry=dry)
    # culadmin 是船期/加油/吃水的唯一生产者, 本机一律以远端为准
    sync_simple("cul_daily_movement.html", LOCAL_DM, dry=dry)
    for bf in BUNKER_FILES:
        sync_simple("rob_data/" + bf, os.path.join(ROB_DIR, bf), dry=dry)
    merge_draft(dry=dry)
    if os.path.exists(REMOTE_CSV):
        os.remove(REMOTE_CSV)
    print("DONE")
