# -*- coding: utf-8 -*-
"""
ROB 盘油记录自动刷新（Claw-Report 仓库）

流水线:
  1. 船清单: 解析本仓库 cul_daily_movement.html 的 TODAY_DATA (失败则回退 rob_data/dm_summary.json)
  2. ROB:    Outlook CULINES store 实时抓各船最新 Noon/Berth/Sailing Report
             策略: ① Vessel/<船名> 子文件夹 ② 已知 sender 邮箱 ③ 收件箱主题含船名
             抓不到的船保留上次数据(不丢数据)
  3. 输出:   rob_data/rob_results.json (持久化, 记录每船 sender 便于下次定位)
             rob_oil_report.html (AES 加密网页, 密码见 PASSWORD)
             rob_data/rob_oil_table.xlsx + 本机 C:\\CULINES 同步(仅当该目录存在)

用法:
  python rob_refresh.py               # 完整刷新
  python rob_refresh.py --no-outlook  # 只用已有数据重新生成网页(调试)
  python rob_refresh.py --vessel "CHANG SHENG JI 8"  # 只刷新指定船
"""
import sys, os, re, json, base64, hashlib, argparse, tempfile, csv, time
sys.stdout.reconfigure(encoding="utf-8")
from datetime import datetime, timedelta
# 航次油耗(纯本地计算, 不依赖 Outlook/Excel)
# 注意: 多机协作(本机 leahliu / 自动机 culadmin)时 voyage_consumption.py 可能尚未同步到
# 某台机器, 顶层裸 import 会让整个刷新直接崩掉。这里降级为可选依赖: 缺失时只是没有
# 航次油耗, 其余 ROB 功能照常。
try:
    from voyage_consumption import compute_voyages
except Exception as _e:                                   # pragma: no cover
    compute_voyages = None
    print("[WARN] voyage_consumption 不可用, 跳过航次油耗:", _e)

BASE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(BASE, "rob_data")
DM_HTML = os.path.join(BASE, "cul_daily_movement.html")
RESULTS = os.path.join(ROB_DIR, "rob_results.json")
DM_FALLBACK = os.path.join(ROB_DIR, "dm_summary.json")
SENDER_MAP_FILE = os.path.join(ROB_DIR, "vessel_senders.json")
OUT_HTML = os.path.join(BASE, "rob_oil_report.html")
OUT_XLSX = os.path.join(ROB_DIR, "rob_oil_table.xlsx")
CULINES_DIR = r"C:\CULINES\Claw Report"
HISTORY_DIR = os.path.join(ROB_DIR, "history")
# 累加历史 CSV: 单一权威文件, 顶层 rob_data/rob_history.csv (双机协作版规范, 方案A)
# 14 列, 与另一台每日更新 ROB 的机器完全一致, 避免列错位/互相覆盖。
# 注意: sender 不进 CSV(只保留在 rob_results.json), 否则两机 schema 不一致会写坏累计数据。
HISTORY_CSV = os.path.join(ROB_DIR, "rob_history.csv")
# 累加历史 CSV 的列(逐船逐次快照, 一行一船)
SNAP_FIELDS = ["date", "vessel", "code", "lane", "pic",
               "lsfo", "hsfo", "mgo", "ulsfo", "bw", "fw", "refeer",
               "found", "report_time", "speed", "report_type"]

PASSWORD = "jimmy"          # 网页密码(AES, 源码看不到明文; 注意本仓库公开, 密码也在脚本里)
REPORT_KEYS = ("NOON", "BERTH", "SAILING", "ANCHOR", "DRIFT")
OIL_KEYS = ("LSFO", "HSFO", "MGO", "ULSFO")


def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


# ---------------------------------------------------------------- 1. 船清单
# 已下线船舶: 船名带『已下线』后缀(船期表 cul_daily_movement.html 里就是这么标的)。
# 这些船已退租/停航, 船长不再发 ROB 报告, 留着只会占主表、拉低 Data Integrity 统计。
OFFLINE_MARK = "已下线"


def is_offline(name):
    return OFFLINE_MARK in (name or "")


def drop_offline(fleet):
    keep, dropped = [], []
    for v in fleet:
        if is_offline(v.get("vessel")):
            dropped.append(v["vessel"])
        else:
            keep.append(v)
    if dropped:
        print("Fleet: drop %d offline vessels -> %s"
              % (len(dropped), ", ".join(dropped)))
    return keep


def load_fleet():
    if os.path.exists(DM_HTML):
        try:
            html = open(DM_HTML, encoding="utf-8").read()
            m = re.search(r"const\s+TODAY_DATA\s*=\s*", html)
            if m:
                data, _ = json.JSONDecoder().raw_decode(html[m.end():])
                vs = data.get("vessels", [])
                fleet = [{"vessel": (v.get("vessel") or "").strip(),
                          "code": (v.get("code") or "").strip(),
                          "lane": (v.get("route") or "").strip(),
                          "pic": (v.get("pic") or "").strip()} for v in vs]
                if fleet:
                    print("Fleet: %d vessels from cul_daily_movement.html (date %s)"
                          % (len(fleet), data.get("date")))
                    return drop_offline(fleet)
        except Exception as e:
            print("[WARN] parse cul_daily_movement.html failed:", e)
    if os.path.exists(DM_FALLBACK):
        fleet = json.load(open(DM_FALLBACK, encoding="utf-8"))
        print("Fleet: %d vessels from fallback rob_data/dm_summary.json" % len(fleet))
        return drop_offline(fleet)
    raise SystemExit("No fleet source: neither cul_daily_movement.html nor rob_data/dm_summary.json")


# ---------------------------------------------------------------- 2. Outlook ROB
# Outlook 连接状态(供页面提示): 离线时数据是本地缓存, 看着"抓过了"其实没更新
OL_STATE = {"connected": False, "offline": None, "mode": None}


def connect_outlook():
    """返回主邮箱 store。注意: NS.Stores 里『联机存档 - leahliu@culines.com』也含
    CULINES.COM 且可能排在前面, 其收件箱是空的 —— 必须排除存档, 否则什么都抓不到。"""
    import win32com.client
    OL = win32com.client.Dispatch("Outlook.Application")
    NS = OL.GetNamespace("MAPI")
    # 离线/缓存未同步检测: 这时 COM 能连、也能读到邮件, 但都是本地缓存的旧邮件,
    # 刷新会"成功"却拿不到新报告 —— 必须在页面和日志上暴露, 否则看不出异常
    try:
        OL_STATE["offline"] = bool(NS.Offline)
    except Exception:
        OL_STATE["offline"] = None
    try:
        OL_STATE["mode"] = NS.ExchangeConnectionMode
    except Exception:
        OL_STATE["mode"] = None
    stores = [s for s in NS.Stores if "CULINES.COM" in (s.DisplayName or "").upper()]
    main = [s for s in stores
            if "存档" not in (s.DisplayName or "")
            and "ARCHIVE" not in (s.DisplayName or "").upper()]
    if main:
        OL_STATE["connected"] = True
        return main
    if stores:
        OL_STATE["connected"] = True
        return [stores[0]]
    return []


def get_sender(it):
    """尽量拿到真实 SMTP 发件地址。Exchange 账号下 SenderEmailAddress 常返回
    X.500/空, 必须再走 Sender.GetExchangeUser().PrimarySmtpAddress 兜一层。"""
    for attr in ("SenderEmailAddress",):
        try:
            v = getattr(it, attr, "")
            if v and "@" in v:
                return v
        except Exception:
            pass
    try:
        ex = it.Sender
        if ex:
            smtp = ex.GetExchangeUser().PrimarySmtpAddress
            if smtp and "@" in smtp:
                return smtp
    except Exception:
        pass
    return ""


def load_sender_map():
    """船名(norm) -> 船长邮箱, 固化随仓库走。换电脑不依赖运行时历史。"""
    m = {}
    if os.path.exists(SENDER_MAP_FILE):
        try:
            raw = json.load(open(SENDER_MAP_FILE, encoding="utf-8"))
            for k, v in raw.items():
                if v:
                    m[norm(k)] = v
        except Exception as e:
            print("[WARN] load vessel_senders.json failed:", e)
    return m


def save_sender_map(m):
    try:
        with open(SENDER_MAP_FILE, "w", encoding="utf-8") as f:
            json.dump(m, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("[WARN] save vessel_senders.json failed:", e)


def build_folder_cache(inbox):
    """Vessel 子文件夹: norm(文件夹名) -> folder(递归多层)。
    注: dict 按名字做 key, 多艘船同名子文件夹(如 'Master')会互相覆盖,
    因此全量扫描一律改用 build_folder_list()。"""
    cache = {}
    try:
        vf = inbox.Folders["Vessel"]

        def walk(f):
            for c in f.Folders:
                cache.setdefault(norm(c.Name), c)
                try:
                    walk(c)
                except Exception:
                    pass
        walk(vf)
    except Exception:
        pass
    return cache


def build_folder_list(inbox):
    """收件箱下全部子文件夹对象列表(递归多层, 同名不去重)。
    必须覆盖收件箱顶层而不仅是 Vessel/: 相当多船的报告文件夹(SHTG、ZLST、cusk 等)
    直接位于收件箱顶层, 与 Vessel/ 平级; 且多艘船都有 'Master' 同名子文件夹,
    用 dict 存会互相覆盖导致漏扫。这里保留每一个文件夹对象。"""
    out = []

    def walk(f, depth=0):
        if depth > 4:
            return
        try:
            for c in f.Folders:
                out.append(c)
                walk(c, depth + 1)
        except Exception:
            pass
    walk(inbox)
    return out


def build_sender_index(folders, sender_map, per_folder=300, cap_per_sender=15):
    """按用户提供的船长邮箱, 在全部文件夹里预建索引: sender -> 最近若干封邮件(新在前)。
    用 get_sender() 解析真实 SMTP(CULINES 内部 Exchange 存的是 X.500, 直接按 SMTP
    字符串 Restrict 永远匹配不上), 故这里逐封解析而不是用 Restrict。报告无论落在
    哪个文件夹都能被 sender 精确定位。"""
    targets = set(v.lower() for v in sender_map.values() if v)
    if not targets:
        return {}
    idx = {}
    for fobj in folders:
        try:
            items = fobj.Items
            items.Sort("[ReceivedTime]", True)
        except Exception:
            continue
        cnt = 0
        for it in items:
            if cnt >= per_folder:
                break
            cnt += 1
            sa = get_sender(it)
            if not sa:
                continue
            key = sa.lower()
            if key not in targets:
                continue
            lst = idx.get(key)
            if lst is None:
                lst = []
                idx[key] = lst
            if len(lst) < cap_per_sender:
                lst.append(it)
    return idx


def match_folder(cache, vessel, code=None):
    """返回 (folder, exact)。exact=False 表示前缀匹配的共用/变体文件夹(如 'MEDKON'
    同时放 MEDKON DON / MEDKON LIA 的邮件), 需按主题过滤防误抓。
    code: 船代码。很多船的报告文件夹按代码命名(SHTG / ZLST / cusk), 只按船名匹配不到。"""
    nv = norm(vessel)
    nc = norm(code or "")
    if not nv and not nc:
        return None, False
    # 精确: 船名或代码直接命中文件夹名
    if nv and nv in cache:
        return cache[nv], True
    if nc and nc in cache:
        return cache[nc], True
    # 前缀: 共用/变体文件夹, 需主题过滤
    for k, f in cache.items():
        if len(k) < 6:
            continue
        if nv and (k.startswith(nv) or nv.startswith(k)):
            return f, False
        if nc and (k.startswith(nc) or nc.startswith(k)):
            return f, False
    return None, False


def pick_report_attachment(it, strict=False):
    xlsx = None
    for a in it.Attachments:
        fn = a.FileName
        if fn.lower().endswith(".xlsx") and any(k in fn.upper() for k in REPORT_KEYS):
            return a
    if strict:
        return None
    for a in it.Attachments:
        if a.FileName.lower().endswith(".xlsx"):
            return a
    return None


def extract_rob(att):
    import openpyxl
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        att.SaveAsFile(p)
        wb = openpyxl.load_workbook(p, data_only=True)
    except Exception:
        try:
            os.remove(p)
        except Exception:
            pass
        return {}
    # CUL NOON/BERTH/SAILING 报告的油表行有多种写法, 只认行首 "ROB " 会漏抓:
    #   NOON  : "ROB LSFO | 598.17 | MT | ROB MGO | 372.84"   (一行多个键值对)
    #   BERTH : "POB/FWE/ANCHOR AWEIGH ROB LSFO | 602.35"     (阶段前缀)
    # 规则: 单元格含 "ROB" 且 ROB 之后的 token 是已知油种代码时, 取该单元格右侧第一个
    # 数值。同一油种重复出现时后值覆盖前值(BERTH 按 POB/FWE/ANCHOR AWEIGH 排列,
    # 最晚阶段即最新存量)。
    oil_codes = ("LSFO", "HSFO", "ULSFO", "MGO", "BW", "FW", "REFEER", "REEFER")

    import re as _re

    def _as_num(v):
        """单元格 -> float, 非数字返回 None。支持 '1,234.5' / ' 598.17 ' / 数字类型。"""
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace(",", "")
        if not s:
            return None
        if not _re.match(r"^-?\d+(\.\d+)?$", s):
            return None
        return float(s)

    def _take(cu, i, row):
        if "ROB" not in cu:
            return None
        pos = cu.rfind("ROB")
        tail = cu[pos + 3:].strip()
        if not tail:
            return None
        tok = tail.split()[0].strip(" :|-")
        if tok not in oil_codes:
            return None
        # 向右找数值, 但遇到"有内容的文本"(单位 MT / 下一个键名 'POB ROB HSFO')
        # 必须停止 —— BERTH 报告一行放多组 "键 | 值 | MT", 若跨过 MT 继续扫,
        # 空值会误抓到下一组的 0(例: 'POB ROB LSFO | (空) | MT | POB ROB HSFO | 0')
        # 导致整船 ROB 被清零(2026-08-30 CUL HAIPHONG / KAI DA HONG ZHOU 事故)。
        # 空单元格跳过(合并单元格/排版留白), 非空非数字立即终止。
        for j in range(i + 1, len(row)):
            v = row[j]
            if v is None:
                continue
            if isinstance(v, str) and not v.strip():
                continue
            n = _as_num(v)
            if n is None:
                break
            return (tok, round(n, 3))   # 收敛浮点尾巴(3651.5230000000006 -> 3651.523)
        return None

    # ---- draft 前/中/后吃水: 同行业务模板在 'DRAFT FWD|值|M|DRAFT MID|值|M|DRAFT AFT|值|M'
    #      标签与数值在同一行相邻; 不同船模板可能不同, 这里取"含 DRAFT FWD/MID/AFT
    #      的单元格右侧第一个数值"作为对应吃水。
    draft_labels = ("DRAFT FWD", "DRAFT MID", "DRAFT AFT")

    def _take_draft(cu, i, row):
        if not any(l in cu for l in draft_labels):
            return None
        for l in draft_labels:
            if l in cu:
                for j in range(i + 1, len(row)):
                    v = row[j]
                    if v is None:
                        continue
                    if isinstance(v, str) and not v.strip():
                        continue
                    n = _as_num(v)
                    if n is None:
                        break
                    return (l.replace(" ", "_"), round(n, 2))
                return None
        return None

    rob = {}
    speed_cands = []  # (priority, value): 航速候选, 取优先级最高者
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for i, c in enumerate(row):
                if c and isinstance(c, str):
                    cu = c.upper().strip()
                    got = _take(cu, i, row)
                    if got:
                        k, v = got
                        rob["REFEER" if k == "REEFER" else k] = v
                    gotd = _take_draft(cu, i, row)
                    if gotd:
                        k, v = gotd
                        rob[k] = v
                    # ---- 航速: SPEED TO NEXT PORT(开航报告) / AVG SPEED SINCE LAST
                    #      NOON/SAILING REPORT(在航平均航速); 排除 WIND SPEED ----
                    if "SPEED" in cu and "WIND" not in cu:
                        prio = 0
                        if "TO NEXT PORT" in cu:
                            prio = 2
                        elif "SINCE LAST" in cu:
                            prio = 1
                        for j in range(i + 1, len(row)):
                            v = row[j]
                            if v is None:
                                continue
                            if isinstance(v, str) and not v.strip():
                                continue
                            n = _as_num(v)
                            if n is None:
                                break
                            speed_cands.append((prio, n))
                            break
    if speed_cands:
        speed_cands.sort(key=lambda x: -x[0])
        rob["SPEED"] = round(speed_cands[0][1], 2)
    # 低硫船只报 ULSFO 时归一到 LSFO, 统一主表/趋势口径
    if "ULSFO" in rob and "LSFO" not in rob:
        rob["LSFO"] = rob["ULSFO"]
    try:
        os.remove(p)
    except Exception:
        pass
    return rob


def scan_for_rob(items, max_attach=40, max_walk=1200, subject_token=None):
    """倒序遍历邮件, 返回最新一封报告附件含 ROB 的 (rob, receivedTime, subject, sender)。
    subject_token: 若设置(norm 后的船名), 邮件主题必须含该 token 才候选 —— 用于共用文件夹
    防止误抓其他船的报告(如 MEDKON 文件夹里 MEDKON DON 的邮件)。"""
    tried = walked = 0
    for it in items:
        walked += 1
        if walked > max_walk:
            break
        try:
            if subject_token:
                try:
                    subj0 = it.Subject or ""
                except Exception:
                    subj0 = ""
                if subject_token not in norm(subj0):
                    continue
            att = pick_report_attachment(it)
            if att is None:
                continue
            if tried >= max_attach:
                break
            tried += 1
            rob = extract_rob(att)
            if any(k in rob for k in OIL_KEYS):
                se = get_sender(it)
                try:
                    subj = it.Subject or ""
                except Exception:
                    subj = ""
                return rob, it.ReceivedTime, subj, se
        except Exception:
            continue
    return None


def apply_hit(rec, hit):
    rob, recv, subj, se = hit
    subjU = (subj or "").upper()
    report_type = ""
    for k in REPORT_KEYS:
        if k in subjU:
            report_type = k
            break
    rec.update({
        "rob_lsfo": rob.get("LSFO"), "rob_hsfo": rob.get("HSFO"), "rob_mgo": rob.get("MGO"),
        "rob_ulsfo": rob.get("ULSFO"), "rob_bw": rob.get("BW"), "rob_fw": rob.get("FW"),
        "rob_refeer": rob.get("REFEER"),
        "rob_speed": rob.get("SPEED"),
        "rob_draft_fwd": rob.get("DRAFT_FWD"), "rob_draft_mid": rob.get("DRAFT_MID"),
        "rob_draft_aft": rob.get("DRAFT_AFT"),
        "report_time": recv.strftime("%Y-%m-%d %H:%M:%S"),
        "report_type": report_type,
        "source": subj, "sender": se, "found": True,
    })
    return True


def refresh_vessel(inbox, cache, rec, sender_map=None, sender_index=None, folder_list=None,
                   fleet_norms=None, all_inboxes=None):
    """多来源合并取全局最新: ①船名/代码文件夹树 ②sender 索引 ③收件箱 sender Restrict
    ④主题 Restrict(全文件夹)。所有候选按 EntryID 去重、ReceivedTime 全局倒序后,
    从最新一封往下找第一份能解析出 ROB 的报告 —— 无论报告落在哪个文件夹、由哪个发件人
    发出, 永远取时间上最新的那份。
    (旧实现是"哪个分支先命中就用谁", 文件夹里的旧报告会压过 sender 找到的新报告。)"""
    vname = rec["vessel"]
    nv = norm(vname)
    nc = norm(rec.get("code") or "")
    sender_map = sender_map or {}
    # 有效发件人: 运行时历史 优先, 否则用固化映射(换电脑也能用)
    eff_sender = rec.get("sender") or sender_map.get(nv) or sender_map.get(vname)
    # 该发件人是否对应多艘船(共用邮箱, 如 MEDKON DON/LIA)——需主题过滤防误抓
    shared = sum(1 for v in sender_map.values() if v == eff_sender) if eff_sender else 0

    cands = {}  # entryID -> (receivedTime, item)

    def add_items(items, token=None, limit=400):
        n = 0
        try:
            for it in items:
                n += 1
                if n > limit:
                    break
                try:
                    eid = it.EntryID
                    rt = it.ReceivedTime
                except Exception:
                    continue
                if token is not None:
                    try:
                        subj0 = it.Subject or ""
                    except Exception:
                        subj0 = ""
                    ns = norm(subj0)
                    toks = token if isinstance(token, (list, tuple, set)) else (token,)
                    if not any(t and t in ns for t in toks):
                        continue
                old = cands.get(eid)
                if old is None or rt > old[0]:
                    cands[eid] = (rt, it)
        except Exception:
            pass

    # ① 船名/代码命名的文件夹树(Vessel/ 下 + 收件箱顶层如 SHTG/ZLST)
    folder, exact = match_folder(cache, vname, rec.get("code"))
    if folder is None and folder_list:
        nc2 = norm(rec.get("code") or "")
        for fo in folder_list:
            try:
                fn2 = norm(fo.Name)
            except Exception:
                continue
            if fn2 == nv or (nc2 and fn2 == nc2):
                folder, exact = fo, True
                break
    if folder is not None:
        scan_folders = [folder]
        try:
            stack = list(folder.Folders)
            while stack:
                sub = stack.pop(0)
                scan_folders.append(sub)
                try:
                    stack.extend(sub.Folders)
                except Exception:
                    pass
        except Exception:
            pass
        # 主题过滤判定 —— 三种情况都强制按主题认船, 否则会串船:
        #   a) 前缀匹配到的共用文件夹(不精确命中)
        #   b) 共用发件人(一个邮箱发多艘船)
        #   c) **共用文件夹**: 文件夹名还是别的船名/代码的前缀。
        #      例: MEDKON 文件夹同时放 MEDKON DON 与 MEDKON LIA 的报告, 两者都
        #      "精确"命中该文件夹, 谁也不过滤主题 -> 两船都取到文件夹里最新那份
        #      (2026-08-30 实测两船显示完全相同的 276.4/103.3, 实为 DON 的数据)。
        try:
            fnm = norm(folder.Name)
        except Exception:
            fnm = ""
        shared_folder = bool(fnm) and len(fnm) >= 5 and any(
            o and o != fnm and o != nv and o.startswith(fnm)
            for o in (fleet_norms or ()))
        need_subj = shared_folder or (not exact) or shared > 1
        tokens = [t for t in (nv, nc if len(nc) >= 4 else "") if t]
        token1 = tokens if need_subj else None
        for fo in scan_folders:
            try:
                fo.Items.Sort("[ReceivedTime]", True)
            except Exception:
                pass
            add_items(fo.Items, token=token1)

    # ②' sender 预建索引(收件箱全部子文件夹, 已用 get_sender 解析 X.500 -> SMTP); 共用邮箱跳过
    if eff_sender and sender_index and shared <= 1:
        add_items(sender_index.get(eff_sender.lower()) or [])

    # ② 注: 不在这里用 Restrict("[SenderEmailAddress]='<smtp>'") 兜底 —— Outlook 把发件人
    # 存成 X.500 内部地址, 按 SMTP 字符串过滤永远返回 0(实测验证)。真正有效的 sender 检索
    # 是上面 ②' 的预建索引(逐封用 get_sender 解析 SMTP), 多 store 时 all_inboxes 已包含
    # Vessel Report 等共享邮箱, 一旦加到 Outlook profile 即可自动覆盖。

    # ③ 主题含船名(收件箱 + 全部子文件夹, 含顶层船文件夹/嵌套/同名)
    try:
        token_sql = "@SQL=\"urn:schemas:httpmail:subject\" like '%s'" % vname
        for fo in [inbox] + (list(folder_list) if folder_list else list(cache.values())):
            try:
                items = fo.Items.Restrict(token_sql)
                items.Sort("[ReceivedTime]", True)
                add_items(items, limit=60)
            except Exception:
                continue
    except Exception:
        pass

    if not cands:
        return False
    ordered = sorted(cands.values(), key=lambda x: x[0], reverse=True)
    hit = scan_for_rob([it for _, it in ordered], max_walk=400)
    if hit:
        if eff_sender:
            rec["sender"] = eff_sender
        return apply_hit(rec, hit)
    return False


# ---- 偏旧兜底 ----
# 常规多来源检索每条分支都有扫描上限(每文件夹前 N 封 / sender 最近 15 封), 个别船的
# 报告如果落在很深的层级或被大量别的邮件挡住, 就会一直停在旧日期(例: M. ODYSSEY
# 卡在 08-26)。下面按 ReceivedTime 窗口对全部文件夹 Restrict 一次(不受"前 N 封"限制),
# 再按 sender / 主题精确匹配, 保证只要邮箱里有更新的报告就一定能捞出来。
STALE_HOURS = 26          # 报告时间超过这个小时数(或压根没抓到)才进深度扫描
DEEP_LOOKBACK_DAYS = 6    # 深度扫描的回看天数

# Outlook 同步完整性阈值。坑: Outlook 刚启动时有一段"进程在、NS.Offline=False、但缓存
# 还没同步完"的窗口, 此时 Restrict 只返回已同步的那一小部分邮件 —— 刷新会照常跑完、
# 页面照常生成, 但压根没抓到新报告, 且看不出任何异常(2026-09-22 踩过: 池子只有 624 封,
# 同步完是 4224 封)。正常 6 天窗口应有 3000~4000 封带附件的邮件, 低于此阈值即判定
# 本次抓取不可信, 直接中止、不覆盖页面(宁可留旧版, 也不出假数据)。用 --force 可绕过。
MIN_POOL_WARN = 3000    # 低于此值: 照常写盘, 页面 updated 标记"深度兜底可能失效"
                        # (2026-09-23 校准: 无人值守 01:00 那次池子 1518 却没被标出来 ——
                        #  旧阈值 1500 太高抬贵手了。正常同步完实测 kept=3860 / scanned=4158)
MIN_POOL_ABORT = 400    # 低于此值: 判定严重未同步, 中止不写盘(保留上一版), --force 可绕过
# 池子偏小时的自动重试。实测: 同一份代码, Outlook 同步未完成那次只扫到 863,
# 同步完成后是 3555 —— 状态是能自己恢复的, 所以别一次定生死, 等一会儿重扫。
POOL_RETRY = 3          # 最多扫 3 次(首扫 + 2 次重试)
POOL_RETRY_WAIT = 45    # 每次重试前等待秒数


def is_stale(rec, now=None):
    rt = rec.get("report_time")
    if not rt:
        return True
    try:
        t = datetime.strptime(rt[:19], "%Y-%m-%d %H:%M:%S")
    except Exception:
        return True
    return ((now or datetime.now()) - t).total_seconds() > STALE_HOURS * 3600


def build_deep_pool(inbox, folder_list, lookback_days=None, all_inboxes=None):
    """构建深度扫描候选池(最近 N 天、带附件的邮件), 按 ReceivedTime 倒序。

    池的大小同时用作** Outlook 同步完整性判据**: Outlook 刚启动时进程在、
    NS.Offline=False, 但缓存还没同步完, Restrict 只返回已同步的小部分邮件 ——
    刷新会照常跑完却抓不到新报告。正常 6 天窗口应有 3000~4000 封。
    """
    lookback_days = lookback_days or DEEP_LOOKBACK_DAYS
    cutoff = (datetime.now() - timedelta(days=lookback_days)).strftime("%m/%d/%Y %H:%M %p")
    # 支持多 store: 主 inbox + 其他共享邮箱收件箱 + 全量子文件夹
    _roots = (all_inboxes or [inbox])
    folders = list(_roots) + list(folder_list or [])
    # pool 元素: [received, item, subject_norm, sender_lower(惰性, None=未计算)]
    # 关键性能设计: Subject / Sender 每封邮件只算一次并缓存, 供所有船复用。
    # (旧实现是"每艘船 × 每封邮件"各算一次 get_sender —— X.500 解析极慢,
    #  48 艘船 × N 封 = 上万次 COM 调用, 单次刷新要 30+ 分钟。)
    pool = []
    scanned = 0
    for fo in folders:
        try:
            items = fo.Items.Restrict("[ReceivedTime] >= '%s'" % cutoff)
        except Exception:
            continue
        # 不再对每个文件夹做 items.Sort(COM 端排序很慢), 最后统一在 Python 端排
        try:
            for it in items:
                scanned += 1
                try:
                    if it.Attachments.Count == 0:
                        continue          # ROB 报告必带 Excel 附件, 无附件直接跳过
                except Exception:
                    pass                  # 取不到附件数就保守保留, 不误杀
                try:
                    rt = it.ReceivedTime
                except Exception:
                    continue
                try:
                    subj = norm(it.Subject or "")
                except Exception:
                    subj = ""
                pool.append([rt, it, subj, None])
        except Exception:
            continue
    pool.sort(key=lambda x: x[0], reverse=True)
    return pool, scanned, len(folders)


def deep_refresh_stale(inbox, folder_list, recs, sender_map, lookback_days=None,
                       all_inboxes=None):
    """对"报告偏旧/抓不到"的船做一次全文件夹深度扫描, 返回补抓成功的条数。"""
    if not recs:
        return 0
    lookback_days = lookback_days or DEEP_LOOKBACK_DAYS
    pool, scanned, nfolders = build_deep_pool(inbox, folder_list, lookback_days,
                                               all_inboxes=all_inboxes)
    OL_STATE["pool_size"] = len(pool)
    print("deep scan: %d items kept / %d scanned (last %dd, %d folders), %d stale vessels"
          % (len(pool), scanned, lookback_days, nfolders, len(recs)))
    # Outlook 缓存/联机同步没完成时 Restrict 只返回一小部分, 但状态通常几十秒内会恢复。
    # 池子偏小就等一会儿重扫, 别拿第一次的结果定生死(2026-09-23: 首扫 863, 重扫 3555)。
    _try = 1
    while len(pool) < MIN_POOL_WARN and _try < POOL_RETRY:
        _try += 1
        print("   [pool] 池子偏小, %ds 后第 %d/%d 次重扫 ..." % (POOL_RETRY_WAIT, _try, POOL_RETRY))
        time.sleep(POOL_RETRY_WAIT)
        pool, scanned, nfolders = build_deep_pool(inbox, folder_list, lookback_days)
        OL_STATE["pool_size"] = len(pool)
        print("   [pool] 第 %d 次: kept=%d scanned=%d" % (_try, len(pool), scanned))
    if len(pool) < MIN_POOL_ABORT:
        # 池子小到这个程度, 常规检索多半也不可靠 —— 由 main 决定是否写盘
        OL_STATE["sync_suspect"] = True
        print("!" * 68)
        print("!! [不可信] 邮件池只有 %d 封, 中止阈值 %d —— Outlook 严重未同步"
              % (len(pool), MIN_POOL_ABORT))
        print("!! 本次结果将不写盘、不生成页面(保留上一版)。")
        print("!! 请等 Outlook 同步完成后重跑; 确需写入请加 --force。")
        print("!" * 68)
        return 0
    if len(pool) < MIN_POOL_WARN:
        # 常规检索不依赖 Restrict 仍然有效, 只是深度兜底可能漏 —— 写盘但打标记
        OL_STATE["pool_small"] = True
        print("!! [注意] 邮件池 %d 封, 低于正常值 %d —— Outlook 索引可能未完成,"
              % (len(pool), MIN_POOL_WARN))
        print("!!        深度兜底可能漏抓。常规检索结果仍会写入, 页面将带标记。")
    # ExchangeConnectionMode 是比池子更早暴露的同步信号: 无人值守时 COM 拉起 Outlook,
    # 进程在、Offline=False, 但缓存还在"drizzle"(400) 未同步完, Restrict 搜不全。
    # 实测: 未同步=400, 同步完=700。只打标记, 不中止。
    _mode = OL_STATE.get("mode")
    if isinstance(_mode, int) and 0 < _mode < 500:
        OL_STATE["pool_small"] = True
        print("!! [注意] ExchangeConnectionMode=%d (<500) —— Outlook 缓存同步未完成,"
              % _mode)
        print("!!        Restrict 只能搜到已同步的部分, 可能缺新报告。页面将带标记。")

    def _sender(row):
        """惰性解析发件人并写回缓存(每封邮件最多解析一次)。"""
        if row[3] is None:
            try:
                row[3] = (get_sender(row[1]) or "").lower()
            except Exception:
                row[3] = ""
        return row[3]

    fixed = 0
    for rec in recs:
        vname = rec.get("vessel", "")
        nv, nc = norm(vname), norm(rec.get("code") or "")
        eff = rec.get("sender") or sender_map.get(nv) or sender_map.get(vname)
        shared = sum(1 for v in (sender_map or {}).values() if v == eff) if eff else 0
        # 只需要比当前报告更新的邮件: pool 已按 ReceivedTime 倒序,
        # 一旦遇到不比当前报告新的就可以停 —— 后面只会更旧
        cur_t = None
        if rec.get("report_time"):
            try:
                cur_t = datetime.strptime(rec["report_time"][:19], "%Y-%m-%d %H:%M:%S")
            except Exception:
                cur_t = None
        hits = []
        for row in pool:
            if cur_t is not None:
                try:
                    if row[0].replace(tzinfo=None) <= cur_t:
                        break
                except Exception:
                    pass
            subj = row[2]
            ok = bool(nv and nv in subj) or bool(nc and len(nc) >= 4 and nc in subj)
            if not ok and eff:
                ok = (_sender(row) == eff.lower())
            if ok:
                hits.append(row[1])
        if not hits:
            continue
        hit = scan_for_rob(hits, max_walk=200,
                           subject_token=(nv if (shared > 1 or not eff) else None))
        if hit:
            old_t = rec.get("report_time")
            apply_hit(rec, hit)
            if eff:
                rec["sender"] = eff
            fixed += 1
            print("   [DEEP] %-24s %s -> %s  (LSFO=%s MGO=%s)"
                  % (vname[:24], (old_t or "MISS")[:16], rec["report_time"][:16],
                     rec.get("rob_lsfo"), rec.get("rob_mgo")))
    print("deep scan fixed: %d/%d" % (fixed, len(recs)))
    return fixed


# ---------------------------------------------------------------- 3. AES (CryptoJS 兼容)
def evp_bytes_to_key(password, salt, key_len=32, iv_len=16):
    d = b""
    prev = b""
    while len(d) < key_len + iv_len:
        prev = hashlib.md5(prev + password + salt).digest()
        d += prev
    return d[:key_len], d[key_len:key_len + iv_len]


def cryptojs_encrypt(plaintext, passphrase):
    from Crypto.Cipher import AES
    from Crypto.Random import get_random_bytes
    salt = get_random_bytes(8)
    key, iv = evp_bytes_to_key(passphrase.encode(), salt)
    cipher = AES.new(key, AES.MODE_CBC, iv)
    data = plaintext.encode("utf-8")
    pad = 16 - len(data) % 16
    data += bytes([pad]) * pad
    return base64.b64encode(b"Salted__" + salt + cipher.encrypt(data)).decode()


def cryptojs_decrypt(enc, passphrase):
    """Python 端验证解密(模拟 CryptoJS.AES.decrypt(passphrase 模式))"""
    from Crypto.Cipher import AES
    raw = base64.b64decode(enc)
    assert raw[:8] == b"Salted__", "not CryptoJS format"
    salt = raw[8:16]
    ct = raw[16:]
    key, iv = evp_bytes_to_key(passphrase.encode(), salt)
    pt = AES.new(key, AES.MODE_CBC, iv).decrypt(ct)
    pad = pt[-1]
    return pt[:-pad].decode("utf-8")


# ---------------------------------------------------------------- 4. HTML
# 前端模板独立文件: templates/rob.html
# 抓取逻辑(本文件)与页面(HTML/CSS/JS)彻底解耦 —— 另一台电脑改页面只需动
# templates/rob.html, 不会和抓邮件的 Python 逻辑互相覆盖; 本脚本只负责读模板并
# 把加密数据注入 __ENC__ 占位符。
TEMPLATE_FILE = os.path.join(BASE, "templates", "rob.html")


def load_template():
    if not os.path.exists(TEMPLATE_FILE):
        raise SystemExit("Missing template: %s (前端模板缺失, 请从仓库恢复)" % TEMPLATE_FILE)
    return open(TEMPLATE_FILE, encoding="utf-8").read()


def _num(x):
    try:
        if x is None or x == "":
            return None
        return float(x)
    except Exception:
        return None


def _hours_of(s):
    """'YYYY-MM-DD HH:MM' -> 当天小时数(浮点), 用于挑最接近锚点的报告。"""
    import re as _re
    m = _re.search(r"(\d{1,2}):(\d{2})", s or "")
    if not m:
        return 12.0
    return int(m.group(1)) + int(m.group(2)) / 60.0


def build_html(results):
    # ---- 吃水历史(必须在 vessels 之前解析) ----
    # draft_history.csv 是【累加式】独立文件: date,vessel,draft_fwd,draft_mid,draft_aft
    # 这里同时算出 ①每船各位置/整体的历史最大吃水 ②每船最新一条吃水。
    # ②用于给 vessels 的"当前吃水"兜底 —— 因为 --no-outlook 重跑时 results 里没有
    # rob_draft_* 字段(邮件没重新解析), 若只依赖 results 会把已抓到的吃水抹成 None。
    draft_max, draft_latest = {}, {}
    DRAFT_CSV = os.path.join(ROB_DIR, "draft_history.csv")
    if os.path.exists(DRAFT_CSV):
        try:
            with open(DRAFT_CSV, encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    v = row.get("vessel", "")
                    if is_offline(v):
                        continue
                    d = (row.get("date") or "")[:10]
                    fwd = _num(row.get("draft_fwd"))
                    mid = _num(row.get("draft_mid"))
                    aft = _num(row.get("draft_aft"))
                    cur = draft_max.setdefault(v, {"fwd": None, "fwd_date": "",
                                                  "mid": None, "mid_date": "",
                                                  "aft": None, "aft_date": "",
                                                  "max": None, "max_date": ""})
                    for pos, val, dv in (("fwd", fwd, d), ("mid", mid, d), ("aft", aft, d)):
                        if val is None:
                            continue
                        if cur[pos] is None or val > cur[pos]:
                            cur[pos] = val
                            cur[pos + "_date"] = dv
                        if cur["max"] is None or val > cur["max"]:
                            cur["max"] = val
                            cur["max_date"] = dv
                    # 最新一条(按日期字典序; 同日多行取最后出现的那条)
                    lat = draft_latest.get(v)
                    if lat is None or d >= lat[0]:
                        draft_latest[v] = (d, fwd, mid, aft)
        except Exception as e:
            print("[WARN] read draft_history.csv failed:", e)

    vessels = []
    ordered = sorted(results, key=lambda r: (r.get("lane", ""), r.get("code", "")))
    for i, r in enumerate(ordered, 1):
        # 当前吃水: 优先用本次解析出的值; 没有则从 draft_history.csv 取该船最新一条兜底
        _lat = draft_latest.get(r.get("vessel", "")) or (None, None, None, None)
        vessels.append({
            "seq": i, "vessel": r.get("vessel", ""), "code": r.get("code", ""),
            "lane": r.get("lane", ""), "pic": r.get("pic", ""),
            "rob_lsfo": r.get("rob_lsfo"), "rob_hsfo": r.get("rob_hsfo"),
            "rob_ulsfo": r.get("rob_ulsfo"), "rob_mgo": r.get("rob_mgo"),
            "speed": r.get("rob_speed"),
            "draft_fwd": r.get("rob_draft_fwd") if r.get("rob_draft_fwd") is not None else _lat[1],
            "draft_mid": r.get("rob_draft_mid") if r.get("rob_draft_mid") is not None else _lat[2],
            "draft_aft": r.get("rob_draft_aft") if r.get("rob_draft_aft") is not None else _lat[3],
            "found": bool(r.get("found")),
            "remark": "No ROB report from Master found in mailbox" if not r.get("found") else "",
            "report_time": (r.get("report_time") or "")[:19],
        })
    # ---- 历史存档(供网页趋势/消耗分析) ----
    history = []
    if os.path.exists(HISTORY_CSV):
        try:
            with open(HISTORY_CSV, encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    # 14 列规范: date,vessel,code,lane,pic,lsfo,hsfo,mgo,ulsfo,bw,fw,refeer,found,report_time
                    if is_offline(row.get("vessel", "")):
                        continue          # 已下线船的历史不再展示(CSV 归档行保留, 不删)
                    oils = (_num(row.get("lsfo")), _num(row.get("hsfo")),
                            _num(row.get("mgo")), _num(row.get("ulsfo")))
                    if int(row.get("found") or 0) and all(o is None for o in oils):
                        continue          # 脏行: 标记为抓到但四个油种全空(写入时被清零)
                    history.append({
                        "t": (row.get("report_time") or row.get("date") or "")[:16],
                        "v": row.get("vessel", ""),
                        "c": row.get("code", ""),
                        "l": row.get("lane", ""),
                        "ls": oils[0], "hs": oils[1], "mg": oils[2], "us": oils[3],
                        "bw": _num(row.get("bw")),
                        "fw": _num(row.get("fw")),
                        "sp": _num(row.get("speed")),
                        "rtype": (row.get("report_type") or "").strip(),
                        "f": int(row.get("found") or 0),
                        "rt": (row.get("report_time") or "")[:19],    # 船长报告接收时间
                    })
        except Exception as e:
            print("[WARN] read history csv failed:", e)
    # ---- 加油事件(可选, 由用户后续提供; 当前为空则按 ROB 增幅自动识别) ----
    # 文件格式: {"船名": {"YYYY-MM-DD": 加油量MT, ...}, ...}
    bunkering = {}
    BUNKER_JSON = os.path.join(ROB_DIR, "bunkering.json")
    if os.path.exists(BUNKER_JSON):
        try:
            with open(BUNKER_JSON, encoding="utf-8") as f:
                bunkering = json.load(f)
        except Exception as e:
            print("[WARN] read bunkering.json failed:", e)
    # ---- 分油种加油量(来自《燃油添加日志》, 仅体积 MT) ----
    bunker_types = {}
    BUNKER_TYPES_JSON = os.path.join(ROB_DIR, "bunkering_types.json")
    if os.path.exists(BUNKER_TYPES_JSON):
        try:
            with open(BUNKER_TYPES_JSON, encoding="utf-8") as f:
                bunker_types = json.load(f)
        except Exception as e:
            print("[WARN] read bunkering_types.json failed:", e)
    # ---- 航次油耗: 航次首港 Berth 时间 → 下一航次首港 Berth; 消耗 = ROB 窗口内下降量 ----
    voyages = []
    if compute_voyages:
        try:
            voyages = compute_voyages(bunkering=bunkering, bunker_types=bunker_types)
        except Exception as e:
            print("[WARN] compute voyages failed:", e)
    else:
        print("[WARN] 跳过航次油耗: voyage_consumption 模块不可用")
    # (draft_max / draft_latest 已在 build_html 开头解析, 此处不再重复读取)
    # ---- TDR 设计油耗曲线(来自《Vessel Daily Consumption - TCD Daily Consumption》)
    #     结构: {"vessels": {<code>: {"name":..,"speeds":[{"speed","lsfo","hsfo","mgo"}],"port_stay":{..}}}} ----
    tdr = {}
    TDR_JSON = os.path.join(ROB_DIR, "tdr_consumption.json")
    if os.path.exists(TDR_JSON):
        try:
            tdr = json.load(open(TDR_JSON, encoding="utf-8")).get("vessels", {})
        except Exception as e:
            print("[WARN] read tdr_consumption.json failed:", e)
    _upd = datetime.now().strftime("%Y-%m-%d %H:%M")
    if OL_STATE.get("offline"):
        # Outlook 离线(或缓存未同步): 本次"刷新"读的是本地缓存, 数据可能根本没更新
        _upd += "  [OUTLOOK 离线-数据未更新]"
    elif OL_STATE.get("attempted") and not OL_STATE.get("connected"):
        _upd += "  [未连接 Outlook]"
    if OL_STATE.get("pool_small"):
        _upd += "  [邮件池 %s-深度兜底可能漏抓]" % OL_STATE.get("pool_size")
    payload = {"updated": _upd,
               "vessels": vessels,
               "history": history,
               "bunkering": bunkering,
               "bunkering_types": bunker_types,
               "voyages": voyages,
               "draft_max": draft_max,
               "tdr": tdr}
    enc = cryptojs_encrypt(json.dumps(payload, ensure_ascii=False), PASSWORD)
    # Python 端自校验(确保 JS 端能解开)
    back = cryptojs_decrypt(enc, PASSWORD)
    assert json.loads(back)["updated"] == payload["updated"]
    assert "history" in json.loads(back)
    html = load_template().replace("__ENC__", enc)
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML -> %s (%d vessels, history=%d rows, voyages=%d, encrypted OK, %d bytes)"
          % (OUT_HTML, len(vessels), len(history), len(voyages), len(html)))
    print("HTML -> %s (%d vessels, encrypted OK, %d bytes)" % (OUT_HTML, len(vessels), len(html)))


# ---------------------------------------------------------------- 5. xlsx (可选, 本机)
def build_xlsx(results):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except Exception as e:
        print("[WARN] openpyxl missing, skip xlsx:", e)
        return
    snap = datetime.now().strftime("%Y.%-m.%-d") if os.name != "nt" else datetime.now().strftime("%Y.%#m.%#d")
    ordered = sorted(results, key=lambda r: (r.get("lane", ""), r.get("code", "")))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = snap
    ws.cell(1, 1, snap).font = Font(bold=True, size=14)
    ws.cell(1, 2, "盘油记录（ROB 取自各船船长 Noon/Berth/Sailing Report）").font = Font(bold=True, size=11)
    headers = ["序号", "Vessel Name全称", "Vessel Code", "Lane Code", "燃油负责人", "PIC",
               "ROB LSFO", "ROB HSFO", "ROB ULSFO", "ROB MGO", "订油状态", "订油情况",
               "REMARK", "特殊", "拟采购日期", "ROB报告时间"]
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    def fmt_num(x):
        if x is None:
            return None
        if isinstance(x, float):
            return int(x) if x == int(x) else round(x, 2)
        return x

    r = 3
    for i, rec in enumerate(ordered, 1):
        rowvals = [i, rec.get("vessel"), rec.get("code"), rec.get("lane"), "", rec.get("pic"),
                   fmt_num(rec.get("rob_lsfo")), fmt_num(rec.get("rob_hsfo")),
                   fmt_num(rec.get("rob_ulsfo")), fmt_num(rec.get("rob_mgo")),
                   "", "", "", "", "",
                   (rec.get("report_time") or "")[:19]]
        if not rec.get("found"):
            rowvals[12] = "邮箱未找到船长存油报告"
        for c, val in enumerate(rowvals, 1):
            cell = ws.cell(r, c, val)
            cell.border = border
            if c in (1, 2, 3, 4, 7, 8, 9, 10, 16):
                cell.alignment = Alignment(horizontal="center")
        r += 1
    for col, w in zip("ABCDEFGHIJKLMNO", [6, 22, 16, 12, 12, 16, 11, 11, 11, 10, 18, 28, 8, 14, 22]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"
    try:
        wb.save(OUT_XLSX)
        print("XLSX -> %s" % OUT_XLSX)
    except PermissionError:
        alt = OUT_XLSX.replace(".xlsx", "_new.xlsx")
        wb.save(alt)
        print("[WARN] xlsx locked, saved %s" % alt)
    # 本机 CULINES 同步(仅目录存在时)
    if os.path.isdir(CULINES_DIR):
        import shutil
        try:
            dst = os.path.join(CULINES_DIR, "盘油记录.auto.xlsx")
            shutil.copy2(OUT_XLSX, dst)
            print("Synced -> %s" % dst)
        except Exception as e:
            print("[WARN] CULINES sync failed:", e)


# ---------------------------------------------------------------- 6. 每日历史存档(累加)
def migrate_history_csv():
    """一次性把累加 CSV 表头升级到 16 列(追加 speed/report_type)。已升级则直接返回。
    只重写第一行表头, 数据行不动(旧行由 DictReader 自动用 None 补齐缺失列)。"""
    if not os.path.exists(HISTORY_CSV):
        return
    try:
        with open(HISTORY_CSV, encoding="utf-8", newline="") as f:
            lines = f.read().splitlines()
        if not lines:
            return
        if "report_type" in lines[0]:
            return
        lines[0] = ",".join(SNAP_FIELDS)
        with open(HISTORY_CSV, "w", encoding="utf-8", newline="") as f:
            f.write("\n".join(lines) + "\n")
        print("[INFO] history CSV header migrated -> 16 cols (added report_type)")
    except Exception as e:
        print("[WARN] history csv header migration failed:", e)


def write_daily_history(results):
    """累加式每日历史, 两份产物:
       - rob_data/history/rob_YYYY-MM-DD.json : 当天完整快照(按运行覆盖当天)
       - rob_data/history/rob_history.csv      : 逐船逐次快照, 追加(不覆盖历史,
         可随时按 snapshot_time / vessel 筛选任意历史日)
    """
    os.makedirs(HISTORY_DIR, exist_ok=True)
    now = datetime.now()
    snap_time = now.strftime("%Y-%m-%d %H:%M")
    date_tag = now.strftime("%Y-%m-%d")
    # 1) 当日 json 快照(覆盖当天, 不同日期是不同文件 => 自然累积)
    day_file = os.path.join(HISTORY_DIR, "rob_%s.json" % date_tag)
    try:
        with open(day_file, "w", encoding="utf-8") as f:
            json.dump({"snapshot_time": snap_time, "vessels": results},
                      f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("[WARN] write daily json failed:", e)
    # 2) 累加 csv(幂等, 绝不覆盖历史行; 14 列规范与另一台机器一致):
    #    - 有报告时间的行按 (vessel, report_time) 去重 —— 同一份 Noon 报告无论由
    #      每日运行还是回补(--backfill)、哪台机器写入, 都只留一行;
    #    - MISS 行(report_time 为空)按 (date, vessel) 记录, 每次运行留痕,
    #      供网页 Data Integrity 面板展示缺报历史。
    #    date 列统一取"报告日期"(report_time[:10]), 而非运行日, 避免批量回补/双机
    #    运行日差异导致整月历史塌缩到同一天。
    existing_rt, existing_miss = set(), set()
    if os.path.exists(HISTORY_CSV):
        try:
            with open(HISTORY_CSV, encoding="utf-8", newline="") as f:
                for row in csv.reader(f):
                    if len(row) < 14:
                        continue
                    rt = (row[13] or "")[:19]
                    if rt:
                        existing_rt.add((row[1], rt))
                    else:
                        existing_miss.add((row[0], row[1]))
        except Exception:
            pass
    write_header = not os.path.exists(HISTORY_CSV)
    new_rows = []
    for r in results:
        rt19 = (r.get("report_time") or "")[:19]
        vessel = r.get("vessel", "")
        if rt19:
            if (vessel, rt19) in existing_rt:
                continue
            existing_rt.add((vessel, rt19))
            d = rt19[:10]
        else:
            d = snap_time[:10]
            if (d, vessel) in existing_miss:
                continue
            existing_miss.add((d, vessel))
        new_rows.append([
            d, vessel, r.get("code", ""), r.get("lane", ""),
            r.get("pic", ""),
            r.get("rob_lsfo"), r.get("rob_hsfo"), r.get("rob_mgo"),
            r.get("rob_ulsfo"), r.get("rob_bw"), r.get("rob_fw"),
            r.get("rob_refeer", ""),
            int(bool(r.get("found"))), rt19,
            r.get("rob_speed", ""),
            r.get("report_type", ""),
        ])
    try:
        with open(HISTORY_CSV, "a", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            if write_header:
                w.writerow(SNAP_FIELDS)
            for row in new_rows:
                w.writerow(row)
        if new_rows:
            print("history -> %s (+%d rows, day=%s)" % (HISTORY_CSV, len(new_rows), day_file))
        else:
            print("history -> %s (no new rows this run, cumulative preserved)" % HISTORY_CSV)
    except Exception as e:
        print("[WARN] write history csv failed:", e)


def append_history_rows(recs, fleet_lookup=None):
    """把回补解析到的报告追加进累加 CSV(14 列规范, 与另一台机器一致)。
    去重键 (vessel, report_time): 同一封报告无论由每日运行/回补/哪台机器写入都只留一行。
    一天多份报告(Noon/Berth/Sailing)折叠为"每天每船一条": 取当天最接近 Noon(12:00)
    的那条, 与另一台每日一行模型一致, 也正好是趋势锚点想要的每天同一时刻 ROB。
    date 列 = 报告日期(report_time[:10])。"""
    fleet_lookup = fleet_lookup or {}
    existing = set()
    if os.path.exists(HISTORY_CSV):
        try:
            for row in csv.reader(open(HISTORY_CSV, encoding="utf-8")):
                if len(row) >= 14 and row[13]:
                    existing.add((row[1], row[13][:19]))
        except Exception:
            pass
    best = {}   # (date, vessel) -> (score, rec, rt19, date)
    for r in recs:
        rt19 = (r.get("report_time") or "")[:19]
        if not rt19:
            continue
        if (r["vessel"], rt19) in existing:
            continue
        d = rt19[:10]
        key = (d, r["vessel"])
        score = abs(_hours_of(rt19) - 12.0)
        cur = best.get(key)
        if cur is None or score < cur[0]:
            best[key] = (score, r, rt19, d)
    rows = []
    for (d, vessel), (score, r, rt19, dd) in best.items():
        fl = fleet_lookup.get(norm(r["vessel"]), {})
        rows.append([
            dd, vessel, fl.get("code", r.get("code", "")), fl.get("lane", r.get("lane", "")),
            "",  # pic: 回补不抓船长名, 留空(另一台机器的行会带)
            r.get("rob_lsfo"), r.get("rob_hsfo"), r.get("rob_mgo"),
            r.get("rob_ulsfo"), r.get("rob_bw"), r.get("rob_fw"),
            "",  # refeer: 回补不抓, 留空
            1, rt19,
            r.get("rob_speed", ""),
            r.get("report_type", ""),
        ])
    try:
        write_header = not os.path.exists(HISTORY_CSV)
        with open(HISTORY_CSV, "a", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            if write_header:
                w.writerow(SNAP_FIELDS)
            for row in rows:
                w.writerow(row)
    except Exception as e:
        print("[WARN] append history failed:", e)
    return len(rows)


def write_draft_history(results):
    """累加式吃水历史(独立文件, 不污染双机共用的 14 列 ROB CSV):
       rob_data/draft_history.csv, 列: date,vessel,draft_fwd,draft_mid,draft_aft
       去重键 (date, vessel) —— 同一天同船只留一行(取当天最新一次报告的吃水)。"""
    DRAFT_CSV = os.path.join(ROB_DIR, "draft_history.csv")
    existing = set()
    if os.path.exists(DRAFT_CSV):
        try:
            with open(DRAFT_CSV, encoding="utf-8", newline="") as f:
                for row in csv.reader(f):
                    if len(row) >= 5 and row[1]:
                        existing.add((row[0], row[1]))
        except Exception:
            pass
    new_rows = []
    for r in results:
        rt19 = (r.get("report_time") or "")[:19]
        if not rt19:
            continue
        fwd = _num(r.get("rob_draft_fwd"))
        mid = _num(r.get("rob_draft_mid"))
        aft = _num(r.get("rob_draft_aft"))
        if fwd is None and mid is None and aft is None:
            continue
        d = rt19[:10]
        if (d, r.get("vessel", "")) in existing:
            continue
        existing.add((d, r.get("vessel", "")))
        new_rows.append([d, r.get("vessel", ""), fwd, mid, aft])
    try:
        write_header = not os.path.exists(DRAFT_CSV)
        with open(DRAFT_CSV, "a", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            if write_header:
                w.writerow(["date", "vessel", "draft_fwd", "draft_mid", "draft_aft"])
            for row in new_rows:
                w.writerow(row)
        if new_rows:
            print("draft -> %s (+%d rows)" % (DRAFT_CSV, len(new_rows)))
    except Exception as e:
        print("[WARN] write draft_history.csv failed:", e)


def backfill_history(days, sender_map, fleet_lookup=None):
    """回补过去 days 天的每一份 ROB 报告(Noon/Berth/Sailing), 按船长邮箱逆向映射识别船。
    返回 (recs 已解析, unknown 未知发件人列表)。未知发件人需交用户确认后固化进 vessel_senders.json。"""
    import win32com.client
    stores = connect_outlook()
    if not stores:
        print("[WARN] no CULINES store, backfill aborted"); return [], []
    all_inboxes = []
    folders = []
    seen_folders = set()
    def _walk(f):
        try:
            for c in f.Folders:
                if id(c) in seen_folders:
                    continue
                seen_folders.add(id(c))
                folders.append(c)
                _walk(c)
        except Exception:
            pass
    for store in stores:
        try:
            inbox = store.GetDefaultFolder(6)
            all_inboxes.append(inbox)
            if id(inbox) not in seen_folders:
                seen_folders.add(id(inbox))
                folders.append(inbox)
            _walk(inbox)
        except Exception as e:
            print("[WARN] backfill store %s failed: %s" % (store.DisplayName, e))
    inbox = all_inboxes[0]
    cache = build_folder_cache(inbox)
    rev = {}
    for v, e in sender_map.items():
        if e:
            rev.setdefault(e.lower(), []).append(v)
    cutoff = datetime.now().astimezone() - timedelta(days=days)
    cutoff_naive = cutoff.replace(tzinfo=None)
    filt = cutoff_naive.strftime("%m/%d/%Y %H:%M %p")
    unknown, recs, seen = [], [], set()
    for f in folders:
        try:
            items = f.Items.Restrict("[ReceivedTime] >= '%s'" % filt)
        except Exception:
            try:
                items = f.Items
            except Exception:
                continue
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        in_vessel_tree = (f is not inbox)
        for it in items:
            try:
                rt = it.ReceivedTime
            except Exception:
                continue
            if rt < cutoff:
                continue
            att = pick_report_attachment(it, strict=True)
            if att is None:
                continue
            rob = extract_rob(att)
            if not any(k in rob for k in OIL_KEYS):
                continue
            # 这是一份 ROB 报告, 识别船(按发件人邮箱优先, 其次 Vessel 子文件夹)
            se = get_sender(it)
            vessel = None
            cands = rev.get(se.lower(), []) if se else []
            folder_v = norm(f.Name) if in_vessel_tree else None
            if cands:
                if len(cands) == 1:
                    vessel = cands[0]
                else:
                    try:
                        subj = norm(it.Subject or "")
                    except Exception:
                        subj = ""
                    hit = [v for v in cands if v in subj]
                    vessel = hit[0] if hit else cands[0]
            elif folder_v and folder_v in cache:
                vessel = f.Name
            if not vessel:
                unknown.append({"sender": se, "subject": (it.Subject or "")[:90],
                                "received": rt.strftime("%Y-%m-%d %H:%M")})
                continue
            rep_t = rt.strftime("%Y-%m-%d %H:%M:%S")
            key = (rep_t, vessel)
            if key in seen:
                continue
            seen.add(key)
            recs.append({
                "vessel": vessel, "report_time": rep_t, "sender": se,
                "rob_lsfo": rob.get("LSFO"), "rob_hsfo": rob.get("HSFO"),
                "rob_mgo": rob.get("MGO"), "rob_ulsfo": rob.get("ULSFO"),
                "rob_bw": rob.get("BW"), "rob_fw": rob.get("FW"),
                "rob_draft_fwd": rob.get("DRAFT_FWD"), "rob_draft_mid": rob.get("DRAFT_MID"),
                "rob_draft_aft": rob.get("DRAFT_AFT"),
            })
    return recs, unknown


# ---------------------------------------------------------------- main
def backfill_mode(days):
    sender_map = load_sender_map()
    fleet = load_fleet()
    fleet_lookup = {norm(v["vessel"]): v for v in fleet}
    # norm -> 显示名(fleet + vessel.csv), 回补按发件人映射归船拿到的是 norm 键,
    # 写入历史前统一转显示名, 避免与每日运行的大写船名在趋势页分裂成两条船。
    disp = {norm(v["vessel"]): v["vessel"] for v in fleet}
    try:
        with open(os.path.join(BASE, "vessel.csv"), encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                if len(row) >= 2 and row[0].strip():
                    nv = norm(row[0])
                    disp.setdefault(nv, row[0].strip())
                    if nv not in fleet_lookup:
                        fleet_lookup[nv] = {"vessel": row[0].strip(),
                                            "code": row[1].strip(), "lane": ""}
    except Exception:
        pass
    recs, unknown = backfill_history(days, sender_map, fleet_lookup)
    for r in recs:
        r["vessel"] = disp.get(r["vessel"], r["vessel"].upper())
    added = append_history_rows(recs, fleet_lookup)
    write_draft_history(recs)
    print("backfill: parsed %d reports, appended %d new rows (past %d days)"
          % (len(recs), added, days))
    results = []
    if os.path.exists(RESULTS):
        results = json.load(open(RESULTS, encoding="utf-8"))
    build_html(results)
    build_xlsx(results)
    if unknown:
        print("\n=== UNKNOWN SENDERS (%d) — ROB reports from unmapped captains ==="
              % len(unknown))
        for u in unknown[:300]:
            print("  SENDER=%s | %s | %s" % (u["sender"], u["received"], u["subject"]))
    print("DONE")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-outlook", action="store_true", help="不抓 Outlook, 只重建网页")
    ap.add_argument("--force", action="store_true",
                    help="忽略 Outlook 同步完整性检查(邮件池过小仍写盘, 慎用)")
    ap.add_argument("--vessel", default=None, help="只刷新指定船(名称子串)")
    ap.add_argument("--backfill", type=int, default=0,
                    help="回补过去 N 天的每份 ROB 报告(按船长邮箱识别船), 追加进历史 CSV")
    args = ap.parse_args()
    if args.backfill:
        backfill_mode(args.backfill)
        return

    os.makedirs(ROB_DIR, exist_ok=True)
    fleet = load_fleet()

    results = []
    if os.path.exists(RESULTS):
        results = json.load(open(RESULTS, encoding="utf-8"))
    old_by_vessel = {}
    for r in results:
        old_by_vessel[norm(r.get("vessel", ""))] = r

    merged = []
    for v in fleet:
        old = old_by_vessel.get(norm(v["vessel"]), {})
        merged.append({
            "vessel": v["vessel"],
            "code": v["code"] or old.get("code", ""),
            "lane": v["lane"] or old.get("lane", ""),
            "pic": v["pic"] or old.get("pic", ""),
            "rob_lsfo": old.get("rob_lsfo"), "rob_hsfo": old.get("rob_hsfo"),
            "rob_mgo": old.get("rob_mgo"), "rob_ulsfo": old.get("rob_ulsfo"),
            "rob_bw": old.get("rob_bw"), "rob_fw": old.get("rob_fw"),
            "rob_refeer": old.get("rob_refeer"),
            "rob_speed": old.get("rob_speed"),
            "report_time": old.get("report_time"), "source": old.get("source"),
            "sender": old.get("sender"), "found": old.get("found", False),
        })

    if not args.no_outlook:
        sender_map = load_sender_map()
        OL_STATE["attempted"] = True      # 区分"没连上"和"压根没尝试(--no-outlook 重建)"
        stores = connect_outlook()
        if not stores:
            print("[WARN] CULINES Outlook store not found, keep old data")
        else:
            # 支持多 store(主邮箱 + Vessel Report 等共享邮箱), 分别取收件箱和子文件夹
            all_inboxes = []
            all_folders = []
            for s in stores:
                try:
                    _inb = s.GetDefaultFolder(6)
                    all_inboxes.append(_inb)
                    all_folders.extend(build_folder_list(_inb))
                except Exception as e:
                    print("[WARN] store %s GetDefaultFolder failed: %s" % (s.DisplayName, e))
            inbox = all_inboxes[0]
            if OL_STATE.get("offline"):
                print("!" * 66)
                print("!! [严重] Outlook 处于【离线/缓存未同步】状态 (mode=%s)" % OL_STATE.get("mode"))
                print("!! 本次刷新读到的是本地缓存邮件, 抓不到任何新报告 ——")
                print("!! 页面会显示旧数据但看起来像刚抓过。请连上网络/VPN 后重跑。")
                print("!" * 66)
            elif OL_STATE.get("mode") is not None:
                print("Outlook 在线 (ExchangeConnectionMode=%s, Offline=%s)"
                      % (OL_STATE.get("mode"), OL_STATE.get("offline")))
            cache = build_folder_cache(inbox)
            # 全量文件夹列表(含所有 store 的收件箱顶层船文件夹 + 嵌套 + 同名不去重),
            # 供 sender 索引和主题兜底
            folder_list = all_folders
            sender_index = build_sender_index(all_inboxes + folder_list, sender_map)
            # 全船队 norm(船名 + 船代码): 用于识别"共用文件夹"(MEDKON 里同时有
            # MEDKON DON / MEDKON LIA), 命中这类文件夹时必须按主题认船防串数据
            fleet_norms = set()
            for v in merged:
                fleet_norms.add(norm(v["vessel"]))
                if v.get("code"):
                    fleet_norms.add(norm(v["code"]))
            print("Outlook store OK, Vessel folders: %d, all folders: %d, stores: %d, "
                  "sender map: %d, sender index: %d"
                  % (len(cache), len(folder_list), len(stores), len(sender_map), len(sender_index)))
            n_new = 0
            targets = merged
            if args.vessel:
                targets = [r for r in merged if args.vessel.upper() in r["vessel"].upper()]
            for i, rec in enumerate(targets, 1):
                got = refresh_vessel(inbox, cache, rec, sender_map,
                                     sender_index, folder_list, fleet_norms,
                                     all_inboxes=all_inboxes)
                n_new += 1 if got else 0
                mark = "NEW" if got else ("keep" if rec.get("found") else "MISS")
                print("[%2d/%2d] %-24s %-8s %-5s LSFO=%-8s MGO=%-8s t=%s"
                      % (i, len(targets), rec["vessel"], rec["code"], mark,
                         rec.get("rob_lsfo"), rec.get("rob_mgo"),
                         (rec.get("report_time") or "")[:16]))
                # 自动学习: 成功抓到且有真实发件人, 回写固化映射
                if got:
                    s = rec.get("sender")
                    if s and norm(rec["vessel"]) not in sender_map:
                        sender_map[norm(rec["vessel"])] = s
                        print("   + learned sender for %s: %s" % (rec["vessel"], s))
            # 兜底: 常规检索后仍偏旧/没抓到的船, 再做一次全文件夹深度扫描
            stale = [r for r in targets if is_stale(r)]
            if stale:
                print("stale vessels (older than %dh or MISS): %d -> %s"
                      % (STALE_HOURS, len(stale),
                         ", ".join(r["vessel"] for r in stale)))
                deep_refresh_stale(inbox, folder_list, stale, sender_map,
                                    all_inboxes=all_inboxes)
            save_sender_map(sender_map)
            print("refreshed this run: %d/%d" % (n_new, len(targets)))

    # 同步完整性闸门: Outlook 缓存没同步完时 Restrict 只返回一小部分邮件, 刷新会
    # "成功"却抓不到新报告。此时宁可保留上一版, 也不写假数据覆盖页面。
    if OL_STATE.get("sync_suspect") and not args.force:
        print("\n" + "!" * 68)
        print("!! 已中止: 邮件池 %s 封 < 中止阈值 %d, 判定 Outlook 严重未同步。"
              % (OL_STATE.get("pool_size"), MIN_POOL_ABORT))
        print("!! 未写 rob_results.json / 未生成网页 / 未追加历史 —— 页面保持上一版。")
        print("!! 等 Outlook 同步完再跑一次; 确需写入加 --force。")
        print("!" * 68)
        return 1

    json.dump(merged, open(RESULTS, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    found = sum(1 for r in merged if r.get("found"))
    print("found ROB: %d/%d -> %s" % (found, len(merged), RESULTS))

    migrate_history_csv()
    build_html(merged)
    build_xlsx(merged)
    write_daily_history(merged)
    write_draft_history(merged)
    print("DONE")


if __name__ == "__main__":
    # 用 main() 的返回值作退出码: 同步不完整中止时返回 1, 让 bat / 计划任务能感知失败
    sys.exit(main() or 0)
