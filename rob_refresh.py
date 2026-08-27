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
import sys, os, re, json, base64, hashlib, argparse, tempfile, gc
sys.stdout.reconfigure(encoding="utf-8")
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(BASE, "rob_data")
DM_HTML = os.path.join(BASE, "cul_daily_movement.html")
RESULTS = os.path.join(ROB_DIR, "rob_results.json")
DM_FALLBACK = os.path.join(ROB_DIR, "dm_summary.json")
SENDER_MAP_FILE = os.path.join(ROB_DIR, "vessel_senders.json")
OUT_HTML = os.path.join(BASE, "rob_oil_report.html")
OUT_XLSX = os.path.join(ROB_DIR, "rob_oil_table.xlsx")
CULINES_DIR = r"C:\CULINES\Claw Report"

PASSWORD = "jimmy"          # 网页密码(AES, 源码看不到明文; 注意本仓库公开, 密码也在脚本里)
REPORT_KEYS = ("NOON", "BERTH", "SAILING", "ANCHOR", "DRIFT")
OIL_KEYS = ("LSFO", "HSFO", "MGO")


def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


# ---------------------------------------------------------------- 1. 船清单
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
                    return fleet
        except Exception as e:
            print("[WARN] parse cul_daily_movement.html failed:", e)
    if os.path.exists(DM_FALLBACK):
        fleet = json.load(open(DM_FALLBACK, encoding="utf-8"))
        print("Fleet: %d vessels from fallback rob_data/dm_summary.json" % len(fleet))
        return fleet
    raise SystemExit("No fleet source: neither cul_daily_movement.html nor rob_data/dm_summary.json")


# ---------------------------------------------------------------- 2. Outlook ROB
def connect_outlook():
    """返回主邮箱 store。注意: NS.Stores 里『联机存档 - leahliu@culines.com』也含
    CULINES.COM 且可能排在前面, 其收件箱是空的 —— 必须排除存档, 否则什么都抓不到。"""
    import win32com.client
    OL = win32com.client.Dispatch("Outlook.Application")
    NS = OL.GetNamespace("MAPI")
    stores = [s for s in NS.Stores if "CULINES.COM" in (s.DisplayName or "").upper()]
    main = [s for s in stores
            if "存档" not in (s.DisplayName or "")
            and "ARCHIVE" not in (s.DisplayName or "").upper()]
    if main:
        return main[0]
    return stores[0] if stores else None


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
    """Vessel 子文件夹: norm(文件夹名) -> folder"""
    cache = {}
    try:
        vf = inbox.Folders["Vessel"]
        for c in vf.Folders:
            cache[norm(c.Name)] = c
    except Exception:
        pass
    return cache


def match_folder(cache, vessel):
    """返回 (folder, exact)。exact=False 表示前缀匹配的共用/变体文件夹(如 'MEDKON'
    同时放 MEDKON DON / MEDKON LIA 的邮件), 需按主题过滤防误抓。"""
    nv = norm(vessel)
    if not nv:
        return None, False
    if nv in cache:
        return cache[nv], True
    for k, f in cache.items():
        if len(k) >= 6 and (k.startswith(nv) or nv.startswith(k)):
            return f, False
    return None, False


def pick_report_attachment(it):
    xlsx = None
    for a in it.Attachments:
        fn = a.FileName
        if fn.lower().endswith(".xlsx") and any(k in fn.upper() for k in REPORT_KEYS):
            return a
    for a in it.Attachments:
        if a.FileName.lower().endswith(".xlsx"):
            return a
    return None


def _scan_rob(wb):
    """从已加载 workbook 提取 ROB 油种(扫 'ROB <油种>' 取右一格)。"""
    rob = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for i, c in enumerate(row):
                if c and isinstance(c, str):
                    cu = c.upper().strip()
                    if cu.startswith("ROB ") and i + 1 < len(row):
                        try:
                            rob[cu[4:].strip()] = float(row[i + 1])
                        except Exception:
                            pass
    return rob


def extract_rob(att):
    import openpyxl, io
    # 优先用 att.Content 字节流(不落盘): SaveAsFile 对异常/嵌入附件极易触发
    # Outlook COM 原生崩溃(进程直接被杀, 无 Python 堆栈), 故仅作最后回退。
    try:
        data = att.Content
        if data:
            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
            return _scan_rob(wb)
    except Exception:
        pass
    # 回退: SaveAsFile 落盘后读取(针对 att.Content 不支持的少数附件)
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        att.SaveAsFile(p)
        wb = openpyxl.load_workbook(p, data_only=True)
        return _scan_rob(wb)
    except Exception:
        return {}
    finally:
        try:
            os.remove(p)
        except Exception:
            pass


def scan_for_rob(items, max_attach=20, max_walk=600, subject_tokens=None):
    """倒序遍历邮件(调用方需保证已按 ReceivedTime 倒序), 返回最新一封报告附件含 ROB
    的 (rob, receivedTime, subject, sender)。
    subject_tokens: 若设置(可传单个 token 或 token 列表, norm 后), 邮件主题必须含其中
    至少一个 token 才候选 —— 用于共用/前缀文件夹或共用发件人时防止误抓其他船的报告
    (如 MEDKON 文件夹里 MEDKON DON 的邮件; 或 M. ODYSSEY 主题只写代码 MODS 而非全称)。"""
    if subject_tokens and not isinstance(subject_tokens, (list, tuple, set)):
        subject_tokens = [subject_tokens]
    subject_tokens = [t for t in (subject_tokens or []) if t]
    tried = walked = 0
    for it in items:
        walked += 1
        if walked > max_walk:
            break
        try:
            if subject_tokens:
                try:
                    subj0 = it.Subject or ""
                except Exception:
                    subj0 = ""
                ns = norm(subj0)
                if not any(t in ns for t in subject_tokens):
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
    rec.update({
        "rob_lsfo": rob.get("LSFO"), "rob_hsfo": rob.get("HSFO"), "rob_mgo": rob.get("MGO"),
        "rob_ulsfo": rob.get("ULSFO"), "rob_bw": rob.get("BW"), "rob_fw": rob.get("FW"),
        "rob_refeer": rob.get("REFEER"),
        "report_time": recv.strftime("%Y-%m-%d %H:%M:%S"),
        "source": subj, "sender": se, "found": True,
    })
    return True


def _all_subfolders(parent, max_depth=4):
    """递归 yield parent 下的所有子文件夹(含直接子文件夹)。"""
    stack = [(parent, 0)]
    while stack:
        f, d = stack.pop()
        if d >= max_depth:
            continue
        try:
            for sub in f.Folders:
                yield sub
                stack.append((sub, d + 1))
        except Exception:
            pass


def _search_subject_recursive(inbox, token):
    """在收件箱 + 所有子文件夹中, 按主题含 token(norm 后)倒序返回邮件。
    解决: 报告被归进非 Vessel 子文件夹(如 CUHP\\CUHP Master)或 Vessel 子文件夹
    (如 M. ODYSSEY 在 Vessel\\ODESSY)时, 仅靠 inbox 根目录 Restrict 抓不到的问题。"""
    q = "@SQL=\"urn:schemas:httpmail:subject\" like '%%%s%%'" % token
    out = []
    try:
        folders = [inbox] + list(_all_subfolders(inbox))
    except Exception:
        folders = [inbox]
    for f in folders:
        try:
            items = f.Items.Restrict(q)
            items.Sort("[ReceivedTime]", True)
            for it in items:
                out.append(it)
        except Exception:
            continue
    return out


def _search_sender_recursive(inbox, sender):
    """在收件箱 + 所有子文件夹中, 按发件人邮箱倒序返回邮件(递归, 含子文件夹报告)。"""
    q = "[SenderEmailAddress]='%s'" % sender
    out = []
    try:
        folders = [inbox] + list(_all_subfolders(inbox))
    except Exception:
        folders = [inbox]
    for f in folders:
        try:
            items = f.Items.Restrict(q)
            items.Sort("[ReceivedTime]", True)
            for it in items:
                out.append(it)
        except Exception:
            continue
    return out


def _collect_candidates(inbox, cache, rec, nv, eff_sender, shared):
    """汇总该船所有可能来源的邮件(船文件夹 / 主题含船名 / 主题含代码 / 已知发件人),
    去重后按 ReceivedTime 全局倒序, 返回 (有序邮件列表, 主题过滤 token 列表)。
    关键修复:
      · 旧逻辑在 ② 用 [SenderEmailAddress]='内部邮箱' 做 Restrict —— 内部 Exchange 账号
        的 SenderEmailAddress 是 X.500 而非 SMTP, 查询恒为空, 内邮船只能靠 ③ 兜底;
      · 旧逻辑 ③ 把各文件夹结果按『收件箱→子文件夹』顺序拼接, 未全局排序, 导致子文件夹
        里更新的报告被收件箱里的旧报告压住;
      · 旧逻辑主题搜索只用『去空格小写』token(如 culhochiminh), 但实际报告主题多为
        『CUL HOCHIMINH』『ZHONG LIAN SHAN TOU』(带空格), 落在收件箱(非 Vessel 文件夹)
        的报告因此永远匹配不到 —— 这正是 CUHC / ZLST 卡在旧日期的根因。现补充『保留空格
        小写』token 一起搜, 保证收件箱里的报告也能被命中。
    现统一合并 + 全局倒序 + 取最新一封含 ROB 附件的, 保证拿到『真正最新』的报告。"""
    items = []
    # ① Vessel/<船名> 文件夹(若有): 该船邮件最集中处, 取最近 80 封即可覆盖最新
    folder, exact = match_folder(cache, rec["vessel"])
    if folder is not None:
        try:
            fis = folder.Items
            fis.Sort("[ReceivedTime]", True)
            cnt = 0
            for it in fis:
                items.append(it)
                cnt += 1
                if cnt >= 80:
                    break
        except Exception:
            pass
    # ② 主题搜索: 同时用『去空格小写』和『保留空格小写』两种 token。
    #    船代码(如 cuhc / zlst)通常连写, 去空格 token 即可命中; 船名(如 cul hochiminh)
    #    多带空格, 必须保留空格才能命中收件箱里直接发来的报告。
    name_lower = (rec.get("vessel") or "").lower().strip()
    code = norm(rec.get("code", ""))
    toks = []
    for tok in (nv, code, name_lower):
        if tok and tok not in toks:
            toks.append(tok)
    for tok in toks:
        items += _search_subject_recursive(inbox, tok)
    # ③ 已知发件人(尽力而为: 内部 Exchange 账号 SenderEmailAddress 常为 X.500, Restrict
    #    可能漏, 这里仅作补充来源, 主路径靠主题搜索)
    if eff_sender:
        items += _search_sender_recursive(inbox, eff_sender)
    # 去重 + 全局按 ReceivedTime 倒序
    seen = set()
    uniq = []
    for it in items:
        try:
            eid = it.EntryID
        except Exception:
            eid = id(it)
        if eid in seen:
            continue
        seen.add(eid)
        uniq.append(it)

    def _rt(it):
        try:
            return it.ReceivedTime
        except Exception:
            return datetime.min

    uniq.sort(key=_rt, reverse=True)
    # 共用/前缀文件夹 或 共用发件人 时, 要求主题含船名或代码防误抓(同样要把空格 token 纳入)
    need_filter = (not exact) or (shared > 1)
    tokens = toks if need_filter else None
    return uniq, tokens


def refresh_vessel(inbox, cache, rec, sender_map=None):
    sender_map = sender_map or {}
    nv = norm(rec["vessel"])
    # 有效发件人: 运行时历史 优先, 否则用固化映射(换电脑也能用)
    eff_sender = rec.get("sender") or sender_map.get(nv) or sender_map.get(rec["vessel"])
    shared = sum(1 for v in sender_map.values() if v == eff_sender) if eff_sender else 0
    # 合并所有来源, 全局倒序取最新一封含 ROB 的邮件
    try:
        cands, tokens = _collect_candidates(inbox, cache, rec, nv, eff_sender, shared)
        hit = scan_for_rob(cands, subject_tokens=tokens)
        if hit:
            if eff_sender:
                rec["sender"] = eff_sender
            return apply_hit(rec, hit)
    except Exception as e:
        print("   [WARN] refresh %s failed: %s" % (rec["vessel"], e))
    return False


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
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CUL ROB Bunker Report</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/crypto-js/4.2.0/crypto-js.min.js"
        onerror="var s=document.createElement('script');s.src='https://cdn.bootcdn.net/ajax/libs/crypto-js/4.2.0/crypto-js.min.js';document.head.appendChild(s);"></script>
<script>
/* lazy-load xlsx-js-style for Excel export (same CDN fallback chain as cul_daily_movement.html) */
var _xlsxStyled = true;
function _loadXlsx(cb) {
  if (window.XLSX) { cb(); return; }
  var s = document.createElement('script');
  s.src = 'https://unpkg.com/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
  s.onerror = function() {
    _xlsxStyled = false;
    s.src = 'https://cdn.bootcdn.net/ajax/libs/SheetJS/xlsx.full.min.js';
    s.onerror = function() {
      s.src = 'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js';
      s.onerror = function() { alert('Failed to load Excel library. Please check your network.'); };
      document.head.appendChild(s);
    };
    document.head.appendChild(s);
  };
  s.onload = cb;
  document.head.appendChild(s);
}
</script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; color: #1a2332; }
  /* ---- lock screen ---- */
  #lock { position: fixed; inset: 0; z-index: 999; background: linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);
          display: flex; align-items: center; justify-content: center; }
  .lock-card { background: #fff; border-radius: 10px; padding: 36px 40px; width: 340px;
               box-shadow: 0 12px 40px rgba(0,0,0,.3); text-align: center; }
  .lock-card h2 { font-size: 18px; color: #1F4E79; margin-bottom: 6px; }
  .lock-card .sub { font-size: 12px; color: #8a99ab; margin-bottom: 22px; }
  .lock-card input { width: 100%; padding: 10px 14px; border: 1px solid #c9d5e2; border-radius: 6px;
                     font-size: 14px; outline: none; margin-bottom: 14px; }
  .lock-card input:focus { border-color: #2E75B6; }
  .lock-card button { width: 100%; padding: 10px; border: none; border-radius: 6px; background: #1F4E79;
                      color: #fff; font-size: 14px; font-weight: 600; cursor: pointer; }
  .lock-card button:hover { background: #2E75B6; }
  #lockErr { color: #d64545; font-size: 12px; margin-top: 10px; min-height: 16px; }
  /* ---- main app ---- */
  .header { background: linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%); color: #fff;
            padding: 14px 28px 10px; box-shadow: 0 3px 12px rgba(31,78,121,.35);
            display: flex; justify-content: space-between; align-items: center; }
  .header h1 { font-size: 18px; font-weight: 700; letter-spacing: .5px; }
  .header .sub { font-size: 11px; opacity: .75; margin-top: 2px; }
  .header .updated { font-size: 13px; font-weight: 600; background: rgba(255,255,255,.16);
                     padding: 6px 14px; border-radius: 5px; border: 1px solid rgba(255,255,255,.35); }
  .header-right { display: flex; gap: 10px; align-items: center; }
  .btn-export { background: #F6A623; color: #fff; border: none; border-radius: 5px; padding: 8px 18px;
                font-size: 13px; font-weight: 600; cursor: pointer; transition: .15s; }
  .btn-export:hover { background: #d4891a; }
  .stats { display: flex; gap: 14px; padding: 16px 28px; flex-wrap: wrap; }
  .stat { background: #fff; border-radius: 8px; padding: 12px 20px; min-width: 130px;
          box-shadow: 0 2px 8px rgba(31,78,121,.10); }
  .stat .num { font-size: 22px; font-weight: 700; color: #1F4E79; }
  .stat .num.miss { color: #d64545; }
  .stat .lbl { font-size: 12px; color: #5a6e82; margin-top: 2px; }
  .toolbar { padding: 0 28px 12px; display: flex; gap: 12px; }
  .toolbar input { padding: 8px 14px; border: 1px solid #c9d5e2; border-radius: 6px; font-size: 13px;
                   outline: none; width: 300px; }
  .toolbar input:focus { border-color: #2E75B6; }
  .toolbar .hint { font-size: 12px; color: #8a99ab; align-self: center; }
  .tablewrap { padding: 0 28px 30px; overflow-x: auto; }
  table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 8px; overflow: hidden;
          box-shadow: 0 2px 10px rgba(31,78,121,.10); font-size: 13px; }
  thead th { background: #1F4E79; color: #fff; padding: 10px 10px; font-weight: 600; white-space: nowrap;
             position: sticky; top: 0; cursor: pointer; user-select: none; }
  thead th:hover { background: #2E75B6; }
  tbody td { padding: 8px 10px; border-bottom: 1px solid #eef2f7; white-space: nowrap; }
  tbody tr:nth-child(even) { background: #f8fafc; }
  tbody tr:hover { background: #eaf2fa; }
  td.c { text-align: center; } td.n { text-align: right; font-variant-numeric: tabular-nums; }
  td.miss { color: #d64545; font-size: 12px; }
  td.vessel { font-weight: 600; color: #1F4E79; }
  .footer { padding: 10px 28px 26px; font-size: 11px; color: #8a99ab; }
</style>
</head>
<body>

<div id="lock">
  <div class="lock-card">
    <h2>CUL ROB Bunker Report</h2>
    <div class="sub">This page is encrypted. Enter password to continue.</div>
    <input id="pwd" type="password" placeholder="Password" autofocus>
    <button onclick="doUnlock()">Unlock</button>
    <div id="lockErr"></div>
  </div>
</div>

<div id="app" style="display:none">
  <div class="header">
    <div>
      <h1>CUL ROB Bunker Report</h1>
      <div class="sub">ROB from Masters' Noon / Berth / Sailing Reports · Auto-updated daily at 13:00 / 01:00</div>
    </div>
    <div class="header-right">
      <button class="btn-export" onclick="exportExcel()">Export Excel</button>
      <div class="updated" id="updatedAt"></div>
    </div>
  </div>
  <div class="stats" id="stats"></div>
  <div class="toolbar">
    <input id="q" type="text" placeholder="Search vessel / code / lane / PIC..." oninput="renderRows()">
    <span class="hint">Click column headers to sort</span>
  </div>
  <div class="tablewrap">
    <table id="tbl">
      <thead><tr>
        <th data-k="seq">No.</th><th data-k="vessel">Vessel Name</th><th data-k="code">Vessel Code</th>
        <th data-k="lane">Lane</th><th>Bunker PIC</th><th data-k="pic">PIC</th>
        <th data-k="rob_lsfo" class="n">ROB LSFO</th><th data-k="rob_hsfo" class="n">ROB HSFO</th>
        <th data-k="rob_mgo" class="n">ROB MGO</th><th>Order Status</th><th>Order Details</th>
        <th data-k="remark">REMARK</th><th>Special</th><th>Planned Bunkering Date</th><th data-k="report_time">ROB Report Time</th>
      </tr></thead>
      <tbody id="rows"></tbody>
    </table>
  </div>
  <div class="footer">
    Unit: MT · Source: ROB from Masters' emailed Noon / Berth / Sailing reports ·
    Missing vessels flagged in REMARK · Time = report received time ·
    Authorized personnel only - do not share the password
  </div>
</div>

<script>
var ENC = "__ENC__";
var DATA = null, sortKey = "seq", sortAsc = true;

function tryUnlock(pwd) {
  try {
    var dec = CryptoJS.AES.decrypt(ENC, pwd).toString(CryptoJS.enc.Utf8);
    if (!dec) return null;
    return JSON.parse(dec);
  } catch (e) { return null; }
}
function doUnlock() {
  var pwd = document.getElementById('pwd').value;
  var d = tryUnlock(pwd);
  if (d) {
    DATA = d; enterApp();
  } else {
    document.getElementById('lockErr').textContent = 'Wrong password, please try again';
  }
}
function enterApp() {
  document.getElementById('lock').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  document.getElementById('updatedAt').textContent = 'Updated ' + DATA.updated;
  var found = DATA.vessels.filter(function(v){return v.found;}).length;
  document.getElementById('stats').innerHTML =
    stat(DATA.vessels.length, 'Total Vessels') + stat(found, 'ROB Received') +
    stat(DATA.vessels.length - found, 'Missing / No Report', true);
  renderRows();
  document.querySelectorAll('thead th[data-k]').forEach(function(th) {
    th.onclick = function() {
      var k = th.dataset.k;
      if (sortKey === k) sortAsc = !sortAsc; else { sortKey = k; sortAsc = true; }
      renderRows();
    };
  });
}
function stat(n, lbl, miss) {
  return '<div class="stat"><div class="num' + (miss ? ' miss' : '') + '">' + n +
         '</div><div class="lbl">' + lbl + '</div></div>';
}
function fmt(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return (v % 1 === 0) ? v.toString() : v.toFixed(2);
  return v;
}
function renderRows() {
  var q = (document.getElementById('q').value || '').toLowerCase();
  var vs = DATA.vessels.filter(function(v) {
    return !q || (v.vessel + ' ' + v.code + ' ' + v.lane + ' ' + v.pic).toLowerCase().indexOf(q) >= 0;
  });
  vs = vs.slice().sort(function(a, b) {
    var x = a[sortKey], y = b[sortKey];
    if (x === null || x === undefined) x = '';
    if (y === null || y === undefined) y = '';
    if (typeof x === 'number' && typeof y === 'number') return sortAsc ? x - y : y - x;
    x = String(x); y = String(y);
    return sortAsc ? x.localeCompare(y) : y.localeCompare(x);
  });
  var h = '';
  vs.forEach(function(v) {
    h += '<tr><td class="c">' + v.seq + '</td><td class="vessel">' + v.vessel + '</td>' +
         '<td class="c">' + v.code + '</td><td class="c">' + v.lane + '</td><td></td>' +
         '<td>' + v.pic + '</td>' +
         '<td class="n">' + fmt(v.rob_lsfo) + '</td><td class="n">' + fmt(v.rob_hsfo) + '</td>' +
         '<td class="n">' + fmt(v.rob_mgo) + '</td><td></td><td></td>' +
         '<td class="' + (v.found ? '' : 'miss') + '">' + (v.found ? '' : v.remark) + '</td>' +
         '<td></td><td></td><td class="c">' + v.report_time + '</td></tr>';
  });
  document.getElementById('rows').innerHTML = h;
}
var EXPORT_HEADERS = ['No.', 'Vessel Name', 'Vessel Code', 'Lane', 'Bunker PIC', 'PIC',
                      'ROB LSFO', 'ROB HSFO', 'ROB MGO', 'Order Status', 'Order Details',
                      'REMARK', 'Special', 'Planned Bunkering Date', 'ROB Report Time'];
function exportExcel() {
  if (!window.XLSX) { _loadXlsx(exportExcel); return; }
  var q = (document.getElementById('q').value || '').toLowerCase();
  var vs = DATA.vessels.filter(function(v) {
    return !q || (v.vessel + ' ' + v.code + ' ' + v.lane + ' ' + v.pic).toLowerCase().indexOf(q) >= 0;
  });
  var rows = [EXPORT_HEADERS.slice()];
  vs.forEach(function(v) {
    rows.push([v.seq, v.vessel, v.code, v.lane, '', v.pic,
               v.rob_lsfo === null ? '' : v.rob_lsfo,
               v.rob_hsfo === null ? '' : v.rob_hsfo,
               v.rob_mgo === null ? '' : v.rob_mgo,
               '', '', v.found ? '' : v.remark, '', '', v.report_time]);
  });
  var ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{wch:5},{wch:22},{wch:14},{wch:8},{wch:12},{wch:16},
                 {wch:10},{wch:10},{wch:10},{wch:13},{wch:14},
                 {wch:36},{wch:8},{wch:21},{wch:20}];
  if (_xlsxStyled) {
    for (var c = 0; c < EXPORT_HEADERS.length; c++) {
      var cell = ws[XLSX.utils.encode_cell({r: 0, c: c})];
      if (cell) cell.s = { fill: { fgColor: { rgb: '1F4E79' } },
                           font: { color: { rgb: 'FFFFFF' }, bold: true },
                           alignment: { horizontal: 'center' } };
    }
  }
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'ROB Report');
  XLSX.writeFile(wb, 'CUL_ROB_Report_' + DATA.updated.replace(/[^0-9]/g, '').slice(0, 12) + '.xlsx');
}
document.getElementById('pwd').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') doUnlock();
});
</script>
</body>
</html>
"""


def build_html(results):
    vessels = []
    ordered = sorted(results, key=lambda r: (r.get("lane", ""), r.get("code", "")))
    for i, r in enumerate(ordered, 1):
        vessels.append({
            "seq": i, "vessel": r.get("vessel", ""), "code": r.get("code", ""),
            "lane": r.get("lane", ""), "pic": r.get("pic", ""),
            "rob_lsfo": r.get("rob_lsfo"), "rob_hsfo": r.get("rob_hsfo"),
            "rob_mgo": r.get("rob_mgo"),
            "found": bool(r.get("found")),
            "remark": "No ROB report from Master found in mailbox" if not r.get("found") else "",
            "report_time": (r.get("report_time") or "")[:19],
        })
    payload = {"updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
               "vessels": vessels}
    enc = cryptojs_encrypt(json.dumps(payload, ensure_ascii=False), PASSWORD)
    # Python 端自校验(确保 JS 端能解开)
    back = cryptojs_decrypt(enc, PASSWORD)
    assert json.loads(back)["updated"] == payload["updated"]
    html = HTML_TEMPLATE.replace("__ENC__", enc)
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
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
               "ROB LSFO", "ROB HSFO", "ROB MGO", "订油状态", "订油情况", "REMARK",
               "特殊", "拟采购日期", "ROB报告时间"]
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
                   fmt_num(rec.get("rob_lsfo")), fmt_num(rec.get("rob_hsfo")), fmt_num(rec.get("rob_mgo")),
                   "", "", "", "", "",
                   (rec.get("report_time") or "")[:19]]
        if not rec.get("found"):
            rowvals[11] = "邮箱未找到船长存油报告"
        for c, val in enumerate(rowvals, 1):
            cell = ws.cell(r, c, val)
            cell.border = border
            if c in (1, 2, 3, 4, 7, 8, 9, 15):
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


# ---------------------------------------------------------------- main
# ---------------------------------------------------------------- 3b. 每日存档(供分析统计) + MISS 警告
def write_daily_history(merged, today=None):
    """把当天每艘船的 ROB 写入累计历史 CSV(供后续分析统计), 并生成当日完整快照 JSON。
    每天每艘船只保留一行(同日多次运行取最新)。"""
    import csv
    if today is None:
        today = datetime.now().strftime("%Y-%m-%d")
    hist = os.path.join(ROB_DIR, "rob_history.csv")
    cols = ["date", "vessel", "code", "lane", "pic",
            "lsfo", "hsfo", "mgo", "ulsfo", "bw", "fw", "refeer",
            "found", "report_time"]
    rows = {}
    if os.path.exists(hist):
        try:
            with open(hist, encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    rows[(row["date"], row["vessel"])] = row
        except Exception:
            rows = {}
    for r in merged:
        v = r.get("vessel", "")
        rows[(today, v)] = {
            "date": today,
            "vessel": v,
            "code": r.get("code", ""),
            "lane": r.get("lane", ""),
            "pic": r.get("pic", ""),
            "lsfo": r.get("rob_lsfo"),
            "hsfo": r.get("rob_hsfo"),
            "mgo": r.get("rob_mgo"),
            "ulsfo": r.get("rob_ulsfo"),
            "bw": r.get("rob_bw"),
            "fw": r.get("rob_fw"),
            "refeer": r.get("rob_refeer"),
            "found": "1" if r.get("found") else "0",
            "report_time": (r.get("report_time") or "")[:19],
        }
    out = sorted(rows.values(), key=lambda x: (x["date"], x["vessel"]))
    with open(hist, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(out)
    print("HISTORY -> %s (%d rows, today %s)" % (hist, len(out), today))
    # 当日完整快照(JSON, 含 sender/source 便于追溯)
    snap = os.path.join(ROB_DIR, "history", "rob_%s.json" % today)
    os.makedirs(os.path.dirname(snap), exist_ok=True)
    payload = {
        "date": today,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "vessels": merged,
    }
    with open(snap, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print("SNAPSHOT -> %s" % snap)


def write_miss_report(merged, today=None):
    if today is None:
        today = datetime.now().strftime("%Y-%m-%d")
    miss = [r for r in merged if not r.get("found")]
    path = os.path.join(ROB_DIR, "miss_vessels_%s.txt" % today)
    jpath = os.path.join(ROB_DIR, "miss_vessels_%s.json" % today)
    with open(path, "w", encoding="utf-8") as f:
        f.write("MISS VESSELS @ %s (%d)\n" % (today, len(miss)))
        for r in miss:
            f.write("- %s (%s)  last=%s\n"
                    % (r.get("vessel"), r.get("code"), r.get("report_time") or "NONE"))
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump({"date": today, "miss": miss}, f, ensure_ascii=False, indent=2)
    if miss:
        print("[MISS WARNING] %d vessels have NO ROB report today:" % len(miss))
        for r in miss:
            print("   - %s (%s) last=%s"
                  % (r.get("vessel"), r.get("code"), r.get("report_time") or "NONE"))
    else:
        print("[MISS WARNING] all vessels have ROB today")
    return len(miss)


def write_stale_report(merged, today=None):
    """列出『已找到报告但报告时间不是最近两天』的船 —— 这些船可能没抓到最新那份,
    需人工复核(确认船长 26 号确实发了、且主题/代码可被搜到)。"""
    if today is None:
        today = datetime.now().strftime("%Y-%m-%d")
    from datetime import timedelta
    try:
        cutoff = (datetime.strptime(today, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d")
    except Exception:
        cutoff = today
    stale = []
    for r in merged:
        if not r.get("found"):
            continue
        rt = (r.get("report_time") or "")[:10]
        if rt and rt < cutoff:
            stale.append(r)
    path = os.path.join(ROB_DIR, "stale_vessels_%s.txt" % today)
    with open(path, "w", encoding="utf-8") as f:
        f.write("STALE / NOT-LATEST ROB @ %s (%d)  (cutoff < %s)\n"
                % (today, len(stale), cutoff))
        for r in stale:
            f.write("- %s (%s)  report_time=%s  sender=%s\n"
                    % (r.get("vessel"), r.get("code"),
                       r.get("report_time") or "NONE", r.get("sender") or "?"))
    if stale:
        print("[STALE CHECK] %d vessels have ROB older than %s (verify nothing newer was sent):"
              % (len(stale), cutoff))
        for r in stale:
            print("   - %s (%s) report=%s sender=%s"
                  % (r.get("vessel"), r.get("code"),
                     r.get("report_time") or "NONE", r.get("sender") or "?"))
    else:
        print("[STALE CHECK] all found vessels have ROB within last 2 days")
    return len(stale)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-outlook", action="store_true", help="不抓 Outlook, 只重建网页")
    ap.add_argument("--vessel", default=None, help="只刷新指定船(名称子串)")
    ap.add_argument("--start", type=int, default=1, help="分批: 从第 N 艘(1-based)开始")
    ap.add_argument("--limit", type=int, default=0, help="分批: 最多处理 N 艘(0=全部)")
    args = ap.parse_args()

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
            "report_time": old.get("report_time"), "source": old.get("source"),
            "sender": old.get("sender"), "found": old.get("found", False),
        })

    if not args.no_outlook:
        sender_map = load_sender_map()
        store = connect_outlook()
        if store is None:
            print("[WARN] CULINES Outlook store not found, keep old data")
        else:
            inbox = store.GetDefaultFolder(6)
            cache = build_folder_cache(inbox)
            print("Outlook store OK, Vessel folders: %d, sender map: %d"
                  % (len(cache), len(sender_map)))
            n_new = 0
            targets = merged
            if args.vessel:
                targets = [r for r in merged if args.vessel.upper() in r["vessel"].upper()]
            if args.start > 1:
                targets = targets[args.start - 1:]
            if args.limit > 0:
                targets = targets[:args.limit]
            for i, rec in enumerate(targets, 1):
                got = refresh_vessel(inbox, cache, rec, sender_map)
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
                # 每船抓完即落盘 + 释放 COM 对象, 防 Outlook 原生崩溃丢失整轮成果
                json.dump(merged, open(RESULTS, "w", encoding="utf-8"),
                          ensure_ascii=False, indent=2)
                gc.collect()
            save_sender_map(sender_map)
            print("refreshed this run: %d/%d" % (n_new, len(targets)))

    json.dump(merged, open(RESULTS, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    found = sum(1 for r in merged if r.get("found"))
    print("found ROB: %d/%d -> %s" % (found, len(merged), RESULTS))

    build_html(merged)
    build_xlsx(merged)
    # 每日历史存档(供分析统计) + MISS 船清单每日警告
    today = datetime.now().strftime("%Y-%m-%d")
    write_daily_history(merged, today)
    write_miss_report(merged, today)
    write_stale_report(merged, today)
    print("DONE")


if __name__ == "__main__":
    main()
