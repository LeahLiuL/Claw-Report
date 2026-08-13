# BOA Berth-On-Arrival — Lane ↔ Trade & Port ↔ Region Mapping

> Status: current as of **2026-08-13** (data window: 船期统计 202607)
> Applies to: `cul_daily_movement.html` → **BOA Berth-On-Arrival Stats** section (Port Wait tab)

## 1. What this document is

The BOA stats page groups every call by **Lane** and by **Trade** (and by Port/Region).
Lane and Trade are **two separate dimensions**:

- **Lane** = the route/lane read from the Daily Movement vessel block header (column A of `CUL DAILY MOVEMENT.rebuilt.xlsx`), e.g. `AEM`, `ST3`, `REX`.
- **Trade** = the higher-level trade group a lane belongs to, e.g. `MD` (Mediterranean), `ME` (Middle East), `TH` (Thailand).

Mapping rules below define **which Trade a Lane belongs to** (and which Region a Port belongs to).

## 2. Lane → Trade (50 entries)

### 2.1 From sheet `Port & Lane Mapping` (42 entries)

Read at runtime from:
`P:\04 上海操作中心\01 船期管理科\船期管理\准班率BOA\2026\船期统计 202607.xlsx`

| Lane | Trade | Lane | Trade | Lane | Trade |
|------|-------|------|-------|------|-------|
| AEM | MD | AEX | EU | AG2 | ME |
| AGX | ME | CCT | TH | CES | EU |
| CGX | ME | CHT | TH | CIS | IN |
| CP1 | PH | CPX | MN | CST | TH |
| CV3 | VN | CVT | TH | CVX | VN |
| CVX2 | VN | CVX3 | VN | HDT | TW |
| IMR | ME | ISS | IN | JPS | ME |
| NP2 | PH | NSCT1 | TW | NSX | TH |
| RBC1 | TH | REX | ME | SCT | TH |
| SCT2 | TH | SGX | ME | SHX | VN |
| SJA | ME | SL1 | TH | ST3 | TW |
| STD | TW | STX | TH | SV2 | VN |
| SVG | VN | TP1 | US | TPC | TP |
| TPN | TP | TPX | TP | VGX | ME |

### 2.2 Supplemental mappings (8 entries, code-level fallback)

These lanes exist in Daily Movement but are **not covered by the mapping sheet**.
They are hardcoded in `gen_html.py` → `BOA_LANE_TRADE_FALLBACK` with the rationale:

| Lane | Trade | Reason |
|------|-------|--------|
| `RTS` | ME | Red Sea – Middle East ports (SAJED/EGSOK/OMSOH); 船期统计 maps RTS→ME |
| `NAX` | ME→**MD** | User confirmed **NAX = NAF → MD**; Med–N.Africa line (TRALI/TRIST) overlaps AEM(MD) ports |
| `RES` | ME | 船期统计 maps RES→ME (JOAQJ/EGSOK/SAJED) |
| `HLX` | ME | 船期统计 maps HLX→ME (THLCH/SGSIN/INNSA/PKKHI) |
| `GTS` | ME | KR TASMAN calls SAJED/YEADE/EGSOK/OMSOH (Middle East/Red Sea); lanes on same region SGX/CGX→ME |
| `CGS` | ME | Calls mainly AEKLF; lanes calling AEKLF in mapping sheet (SGX/CGX) are ME |
| `CST/SL1` | TH | Combined lane `CST/SL1`; mapping sheet has CST→TH and SL1→TH |
| `ZGCD` | MD | **ZGCD is a vessel name (ZHONG GU CHENG DU), not a lane**; it belongs to AEM (MD) — see §4 |

## 3. Port → Region (94 entries)

### 3.1 From sheet `Port & Lane Mapping` (75 entries)

| Region | Ports |
|--------|-------|
| **China Mainland** | CNGCT CNHMN CNHUA CNHUI CNNAS CNNGB CNSHA CNSHH CNSHK CNSWA CNTAO CNTNJ CNWIT CNXGG CNXMN CNXNA CNXNG CNYPN CNYTN |
| **China HK & TW** | HKHKG TWKEL TWKHH TWTPE TWTXG |
| **Intra Asia** | AEJEA EGALY EGSOK EGSUE IDJKT INMUN INNSA KHKOS KRPUS MYPKG MYPKN MYPKW OMSOH PHMNL PHMNN PHSPS PKKHI QAHMD SADMM SAJED SGSIN THBKK THBKS THLCH THSCS TRGEB VNDAD VNDAN VNHCM VNHPH VNSGN VNVUT YEADE |
| **Europe** | BEANR DEHAM GBSOU GBTIL GRPIR ILASD ILHFA NLAMS NLRTM TRALI TRIST TRIZT TRMER |
| **AF** | DJJIB SDPZU |
| **US** | USLAX USLGB USOAK |

### 3.2 Supplemental port mappings (19 entries, code-level fallback)

| Port | Region | Reason |
|------|--------|--------|
| DZALG | AF | Algeria — Africa |
| AOAQJ | AF | Angola — Africa |
| LYMRA / LYBEN | AF | Libya — Africa |
| TNRDS | AF | Tunisia — Africa |
| AEKLF | Intra Asia | Khalifa, UAE |
| EGSUZ / EGSGA / EGSAF / EGDAM / JOAQJ | Intra Asia | Egypt/Red Sea area |
| SAGIZ | Intra Asia | Suez area |
| INKDL | Intra Asia | Kandla, India |
| MALTA | Europe | Malta |
| GRSKG | Europe | Thessaloniki, Greece |
| TUZLA | Europe | Turkey (near Istanbul), consistent with TRIST/TRIZT |
| CNDCB | China Mainland | Dacheng Bay? — China |
| THPAT / THSSW | Intra Asia | Thailand ports (consistent with THBKK/THSCS/THLCH/THBKS) |

## 4. Dirty-data rule: `ZGCD` is a vessel, not a lane

- Full name: **ZHONG GU CHENG DU (中谷成都)**, vessel code `ZGCD`.
- In the source Excel, one vessel block header (row ~2630) had **column A filled with the vessel code `ZGCD`** instead of the lane (`AEM`), creating a "lane = ZGCD" artifact.
- Correct handling: **do not hardcode** `ZGCD → AEM`. Instead, when a vessel block's A column equals its I column (vessel code), treat it as dirty and **resolve the lane from the same vessel's overlapping/time-nearest block** (row 1899, lane `AEM`). The `ZGCD` lane thus falls into **MD** via `AEM`.

## 5. BOA metrics definition (for reference)

- **Call** = a schedule row with numeric WAIT, inside the selected date range, non-bunkering port.
- **Berth** = WAIT ≤ threshold: **CNSHA ≤ 12h, all other ports ≤ 6h**.
- **Over** = WAIT beyond threshold; range chips (All / 6–24h / 24–48h / 48h+) filter the over count only.
- **Rate** = Berth / (Berth + Over).
- Trade/Region lookup happens **at runtime in the browser** from the embedded `laneTradeMap` / `portRegionMap`; rows that miss both the sheet and the fallback show **Unknown**.

## 6. How the mapping is loaded (and how to update it)

```
gen_html.py → load_boa_mappings()
  1. start from BOA_LANE_TRADE_FALLBACK / BOA_PORT_REGION_FALLBACK (this document's content)
  2. try to read P:...\船期统计 202607.xlsx → sheet "Port & Lane Mapping"
     - every lane→trade / port→region row OVERWRITES the fallback
  3. if the P drive / file is unreachable → fallback is used, still 0 Unknown
  4. both dicts are embedded into the HTML as TODAY_DATA.laneTradeMap / portRegionMap
```

**To update**: edit the sheet `Port & Lane Mapping` (source of truth), then rerun
`python gen_html.py`. If new lanes/ports still show Unknown, add them to the
supplemental sections in `gen_html.py` (§2.2 / §3.2) and update this document.

> ⚠️ GitHub Pages auto-update job: when the P drive is unreachable during the
> nightly job, the published HTML carries the **fallback** mappings (50 lanes /
> 94 ports). This is by design — it keeps the page at 0 Unknown.
