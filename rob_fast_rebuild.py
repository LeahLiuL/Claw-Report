# -*- coding: utf-8 -*-
"""Fast-path rebuild: regenerate HTML/XLSX/history WITHOUT going through
rob_refresh.main(), so the 20-key records (incl. rob_draft_fwd/mid/aft) survive.

Why this exists
---------------
`rob_refresh.py --no-outlook` rebuilds rob_results.json from a 17-field subset
merge list, which SILENTLY DROPS `rob_draft_fwd/mid/aft` for all 51 vessels.
Hit on 2026-10-22: after the summary-email fast sync + `--no-outlook`, drafts
were gone for every vessel and had to be restored from backup.

This module keeps a sidecar `rob_data/draft_cache.json` (vessel -> drafts) and
re-injects any missing draft values before rebuilding, so the fast path can run
every day without ever losing drafts again.
"""
import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as rr

DRAFT_CACHE = os.path.join(rr.ROB_DIR, "draft_cache.json")
DRAFT_KEYS = ("rob_draft_fwd", "rob_draft_mid", "rob_draft_aft")


def load_cache():
    if os.path.exists(DRAFT_CACHE):
        try:
            return json.load(open(DRAFT_CACHE, encoding="utf-8"))
        except Exception as e:
            print("[WARN] draft cache unreadable (%s), rebuild from git HEAD" % e)
    return {}


def main():
    results = json.load(open(rr.RESULTS, encoding="utf-8"))
    cache = load_cache()

    filled = learned = 0
    for r in results:
        v = r.get("vessel")
        if not v:
            continue
        cur = {k: r.get(k) for k in DRAFT_KEYS}
        if all(x is not None for x in cur.values()):
            # 有完整吃水 -> 回写缓存(顺带更新)
            cache[v] = cur
            learned += 1
            continue
        old = cache.get(v)
        if not old:
            continue
        for k in DRAFT_KEYS:
            if r.get(k) is None and old.get(k) is not None:
                r[k] = old[k]
                filled += 1

    print("draft cache: %d vessels, learned %d, re-injected %d fields"
          % (len(cache), learned, filled))
    missing = sum(1 for r in results
                  if any(r.get(k) is None for k in DRAFT_KEYS))
    print("vessels still without full drafts: %d (no history yet = normal)" % missing)

    try:
        json.dump(cache, open(DRAFT_CACHE, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    except Exception as e:
        print("[WARN] save draft_cache failed:", e)

    json.dump(results, open(rr.RESULTS, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)

    rr.write_daily_history(results)
    rr.write_draft_history(results)
    rr.build_html(results)
    rr.build_xlsx(results)
    print("FAST REBUILD DONE")


if __name__ == "__main__":
    main()
