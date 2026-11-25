# -*- coding: utf-8 -*-
"""
一次性改名迁移: EVERLASTING HARVEST (EVHA) -> THE HARMONY (THMY)
同一艘船, 2026-11-14 起船长改用 theharmony@gtmailplus.com 发报告。
- rob_results.json: 只改 vessel / code / sender 三个键, 保留旧 ROB 数值与吃水,
  保证历史连续; report_time/source/report_type 留给随后的 Outlook 采集刷新覆盖。
- vessel_senders.json: 登记新发件人 norm(theharmony) -> 新邮箱; 旧键保留不删。
幂等: 若已改名则跳过。
"""
import os, json, shutil, sys

BASE = os.path.dirname(os.path.abspath(__file__))
ROB = os.path.join(BASE, "rob_data")
RESULTS = os.path.join(ROB, "rob_results.json")
SENDERS = os.path.join(ROB, "vessel_senders.json")

OLD_V, NEW_V = "EVERLASTING HARVEST", "THE HARMONY"
OLD_C, NEW_C = "EVHA", "THMY"
OLD_SENDER = "everlasting.harvest@tcshipping.com.cn"
NEW_SENDER = "theharmony@gtmailplus.com"


def norm(s):
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())


def main():
    # --- results ---
    rr = json.load(open(RESULTS, encoding="utf-8"))
    hit = None
    for r in rr:
        if (r.get("vessel") == OLD_V or r.get("code") == OLD_C):
            hit = r
            break
    if hit is None:
        # 已是新名?
        already = [r for r in rr if r.get("vessel") == NEW_V or r.get("code") == NEW_C]
        if already:
            print("results: already renamed -> %s / %s" % (already[0]["vessel"], already[0]["code"]))
        else:
            print("[ERR] results: neither old nor new vessel found", file=sys.stderr)
            return 1
    else:
        hit["vessel"] = NEW_V
        hit["code"] = NEW_C
        hit["sender"] = NEW_SENDER
        json.dump(rr, open(RESULTS, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print("results: renamed %s/%s -> %s/%s, sender -> %s"
              % (OLD_V, OLD_C, NEW_V, NEW_C, NEW_SENDER))
        print("  preserved ROB: lsfo=%s mgo=%s bw=%s time=%s"
              % (hit.get("rob_lsfo"), hit.get("rob_mgo"), hit.get("rob_bw"), hit.get("report_time")))

    # --- senders ---
    vs = json.load(open(SENDERS, encoding="utf-8"))
    vs[norm(NEW_V)] = NEW_SENDER
    # 旧邮箱也指向新船名(历史邮件兜底)
    vs.setdefault(norm(OLD_V), NEW_SENDER)
    json.dump(vs, open(SENDERS, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print("senders: %s -> %s (old key kept)" % (norm(NEW_V), NEW_SENDER))
    return 0


if __name__ == "__main__":
    sys.exit(main())
