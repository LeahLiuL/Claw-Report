#!/usr/bin/env python3
"""Generate a realistic sample-data.json to demo the Container Explorer,
including deliberately duplicated container numbers so the dedup rule is visible.

Produces sample-data.json + meta.json in the repo root.
"""
import json
import random
from datetime import datetime, timezone

random.seed(42)

PORTS = ["CNSHA", "CNNGB", "CNHKG", "SGSIN", "NLRTM", "DEHAM", "USLAX", "USNYC", "AEJEA", "EGPSD"]
LANES = ["AEX1", "AEX2", "TPE", "AUE", "MEX", "MED"]
VVDS = ["COSCO2607E", "EVERGREEN2201W", "MAERSK3308E", "MSC4412W", "OOCL5519E"]
SLOT = ["MAERSK", "MSC", "COSCO", "EVERGREEN", "OOCL"]
CONT = ["MAERSK", "CMA-CGM", "COSCO", "EVERGREEN", "HAPAG"]
TYPES = ["22G1", "40GP", "40HC", "20RF", "40RH", "20OT", "40FR"]

def rand_cont():
    prefix = random.choice(["MSKU", "CMAU", "COSU", "EGHU", "HLXU"])
    return prefix + str(random.randint(100000, 999999))

def pick_weight():
    return random.choice([0, 4200, 8500, 12500, 16800, 21000, 24500])

def flag():
    return random.choice(["Y", "N", "N", "N"])

rows = []
for _ in range(140):
    pol = random.choice(PORTS)
    pod = random.choice([p for p in PORTS if p != pol])
    rows.append({
        "VVD": random.choice(VVDS),
        "LANE": random.choice(LANES),
        "CONT NO.": rand_cont(),
        "FULL/EMPTY": random.choice(["F", "E"]),
        "TYPE/SIZE": random.choice(TYPES),
        "POL": pol,
        "POD": pod,
        "CONT_Weight": pick_weight(),
        "AWK": flag(),
        "DG": flag(),
        "RF": "Y" if "RF" in random.choice(TYPES) else random.choice(["Y", "N"]),
        "BB": flag(),
        "SLOT_OPR": random.choice(SLOT),
        "CONT_OPR": random.choice(CONT),
        "REVENUE MONTH": f"2026-{random.randint(1,7):02d}",
        "TARGET_PORT": random.choice([pol, pod, random.choice(PORTS)]),
    })

# --- Inject duplicates to demo dedup ---
# Container A: two rows, only the TARGET_PORT==POL one should win
rows.append({"VVD":"COSCO2607E","LANE":"AEX1","CONT NO.":"DEMO000111","FULL/EMPTY":"F","TYPE/SIZE":"40HC",
  "POL":"CNSHA","POD":"NLRTM","CONT_Weight":18500,"AWK":"N","DG":"N","RF":"N","BB":"N",
  "SLOT_OPR":"COSCO","CONT_OPR":"COSCO","REVENUE MONTH":"2026-03","TARGET_PORT":"NLRTM"})
rows.append({"VVD":"COSCO2607E","LANE":"AEX1","CONT NO.":"DEMO000111","FULL/EMPTY":"F","TYPE/SIZE":"40HC",
  "POL":"CNSHA","POD":"NLRTM","CONT_Weight":18500,"AWK":"N","DG":"N","RF":"N","BB":"N",
  "SLOT_OPR":"COSCO","CONT_OPR":"COSCO","REVENUE MONTH":"2026-05","TARGET_PORT":"CNSHA"})  # TARGET==POL wins

# Container B: multiple TARGET==POL rows -> take max REVENUE MONTH
rows.append({"VVD":"MAERSK3308E","LANE":"TPE","CONT NO.":"DEMO000222","FULL/EMPTY":"E","TYPE/SIZE":"20OT",
  "POL":"SGSIN","POD":"USLAX","CONT_Weight":0,"AWK":"Y","DG":"N","RF":"N","BB":"N",
  "SLOT_OPR":"MAERSK","CONT_OPR":"MAERSK","REVENUE MONTH":"2026-02","TARGET_PORT":"SGSIN"})
rows.append({"VVD":"MAERSK3308E","LANE":"TPE","CONT NO.":"DEMO000222","FULL/EMPTY":"E","TYPE/SIZE":"20OT",
  "POL":"SGSIN","POD":"USLAX","CONT_Weight":0,"AWK":"Y","DG":"N","RF":"N","BB":"N",
  "SLOT_OPR":"MAERSK","CONT_OPR":"MAERSK","REVENUE MONTH":"2026-06","TARGET_PORT":"SGSIN"})  # max month wins
rows.append({"VVD":"MAERSK3308E","LANE":"TPE","CONT NO.":"DEMO000222","FULL/EMPTY":"E","TYPE/SIZE":"20OT",
  "POL":"SGSIN","POD":"USLAX","CONT_Weight":0,"AWK":"Y","DG":"N","RF":"N","BB":"N",
  "SLOT_OPR":"MAERSK","CONT_OPR":"MAERSK","REVENUE MONTH":"2026-04","TARGET_PORT":"SGSIN"})

# Container C: no TARGET==POL -> take max REVENUE MONTH among all
rows.append({"VVD":"MSC4412W","LANE":"MEX","CONT NO.":"DEMO000333","FULL/EMPTY":"F","TYPE/SIZE":"40FR",
  "POL":"DEHAM","POD":"USNYC","CONT_Weight":24500,"AWK":"N","DG":"Y","RF":"N","BB":"Y",
  "SLOT_OPR":"MSC","CONT_OPR":"MSC","REVENUE MONTH":"2026-01","TARGET_PORT":"USNYC"})
rows.append({"VVD":"MSC4412W","LANE":"MEX","CONT NO.":"DEMO000333","FULL/EMPTY":"F","TYPE/SIZE":"40FR",
  "POL":"DEHAM","POD":"USNYC","CONT_Weight":24500,"AWK":"N","DG":"Y","RF":"N","BB":"Y",
  "SLOT_OPR":"MSC","CONT_OPR":"MSC","REVENUE MONTH":"2026-07","TARGET_PORT":"USNYC"})  # max month wins

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(rows, f, ensure_ascii=False, separators=(",", ":"))

meta = {
    "generatedAt": datetime.now(timezone.utc).isoformat(),
    "source": "sample generator (demo)",
    "rowCount": len(rows),
    "columns": ["VVD","LANE","CONT NO.","FULL/EMPTY","TYPE/SIZE","POL","POD","CONT_Weight",
                "AWK","DG","RF","BB","SLOT_OPR","CONT_OPR","REVENUE MONTH","TARGET_PORT"],
    "missingColumns": [],
}
with open("meta.json", "w", encoding="utf-8") as f:
    json.dump(meta, f, ensure_ascii=False, indent=2)

print(f"[ok] generated {len(rows)} sample rows -> data.json")
