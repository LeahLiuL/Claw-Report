# -*- coding: utf-8 -*-
"""诊断4: 为什么 rob_refresh 的池子(904)比诊断脚本(3795)小 4 倍?
对照实验: 同一个会话里, 用两种 cutoff 格式分别 Restrict 计数。
"""
import sys, os, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R

store = R.connect_outlook()
NS = store.Session
print("mode =", NS.ExchangeConnectionMode, " offline =", NS.Offline, " DEEP_DAYS =", R.DEEP_LOOKBACK_DAYS)
inbox = store.GetDefaultFolder(6)
folder_list = R.build_folder_list(inbox)
print("folders =", len(folder_list))

now = datetime.datetime.now()
fmt_rob = (now - datetime.timedelta(days=R.DEEP_LOOKBACK_DAYS)).strftime("%m/%d/%Y %H:%M %p")
fmt_alt = (now - datetime.timedelta(days=R.DEEP_LOOKBACK_DAYS)).strftime("%m/%d/%Y %H:%M")
print("cutoff(rob, 带%p) =", repr(fmt_rob))
print("cutoff(alt, 无%p) =", repr(fmt_alt))

def cnt(cut):
    n = 0
    bad = 0
    for f in [inbox] + list(folder_list):
        try:
            n += f.Items.Restrict("[ReceivedTime] >= '%s'" % cut).Count
        except Exception:
            bad += 1
    return n, bad

print("Count with rob cutoff :", cnt(fmt_rob))
print("Count with alt cutoff :", cnt(fmt_alt))

# 真跑一次 build_deep_pool
pool, scanned, nf = R.build_deep_pool(inbox, folder_list)
print("build_deep_pool -> kept=%d scanned=%d folders=%d" % (len(pool), scanned, nf))
