# ROB 盘油 · TDR vs Actual 数据源说明（给 culadmin 每日抓取用）

> 目的：在网页「⚡ TDR vs Actual」tab 上，把**实际每日消耗**与**TDR 设计消耗（按当天实际航速查表）**做时间序列对比。
> 蓝线 = 实际每日消耗（ROB 按锚定小时对齐后相邻日差值，与 Trend 的 Mode:Consumption 同口径）。
> 红线 = 用当天 NOON 报告的航速，在《Vessel Daily Consumption - TCD Daily Consumption》设计油耗表插值得到的设计消耗。

## 一、每次抓取需要落地的字段

`rob_update.bat` 已经会抓取每封报告（NOON / BERTH / SAILING）的下列信息，并写入 `rob_data/rob_history.csv`：

| 字段 | 含义 | 备注 |
|------|------|------|
| `date` | 报告日期（report_time 前 10 位） | |
| `vessel` / `code` | 船名 / 船代码 | |
| `lane` / `pic` | 航线 / 负责人 | |
| `lsfo` / `hsfo` / `mgo` / `ulsfo` | 当时 ROB（MT） | 低硫船的 ULSFO 归一到 LSFO |
| `bw` / `fw` / `refeer` | 压载水 / 淡水 / 冻柜 | |
| `found` | 是否抓到 ROB | |
| `report_time` | 报告时间（含时分秒） | 去重键之一 |
| **`speed`** | 报告里的航速（kn） | **关键字段**：NOON 取 `AVG SPEED SINCE LAST NOON`，SAILING 取 `SPEED TO NEXT PORT`，排除 `WIND SPEED` |
| **`report_type`** | 报告类型 | `NOON` / `BERTH` / `SAILING` / `ANCHOR` / `DRIFT`，从邮件主题识别 |

> 当前 `rob_history.csv` 已升级到 **16 列**（末尾追加 `speed`、`report_type`）。新抓取的报告会自动带这两个值；2026-09-24 之前的历史行这两个字段为空，需要回填（见第三节）。

## 二、网页怎么用这些数据

1. 对所选船 + 油种，把它的所有 ROB 报告按锚定小时（默认 08:00）对齐，相邻日 ROB 差＝当日实际消耗 → 蓝线。
2. 取该船每天 **NOON 报告**的 `speed`（若有），在 TDR 表里按航速线性插值出该天的设计消耗 → 红线。
3. 两线同 X 轴（日期）叠加对比：红线高于蓝线＝实际比设计偏耗油。

TDR 设计油耗来源：
- 原始：`C:\CULINES\Claw Report\Vessel Daily Consumption-TCD Daily Consumption.csv`
- 已解析为：`rob_data/tdr_consumption.json`（48 艘船，每船 `speeds:[{speed, lsfo, hsfo, mgo}]` + `port_stay`）

## 三、历史行回填（重要）

旧历史行缺 `speed` / `report_type`，导致红线无法绘制。请在 culadmin 机器上做一次回填：

1. **Outlook 仍在**：之前的报告邮件大多还在 Outlook 里。运行一次带 Outlook 抓取的回补（不要 `--no-outlook`），让 `rob_refresh.py` 重新解析邮件，并把解析到的 `speed` / `report_type` 写回 `rob_history.csv`（按 `(vessel, report_time)` 去重，只补缺失列，不新增重复行）。
2. 之后每日定时运行 `rob_update.bat` 即可持续累积新数据，无需再回填。

> 注意：若一次全量回补太慢（邮箱邮件多），可只回补最近 1–2 个月，足以让红线在近期出现。

## 四、文件清单（均在 `C:\Users\culadmin\Claw-Report`）

- `rob_refresh.py` — 抓取 + 渲染主脚本（已含 speed / report_type 抓取与 TDR 注入逻辑）
- `templates/rob.html` — 前端（含「⚡ TDR vs Actual」tab 的 `renderTDR()`）
- `rob_data/rob_history.csv` — 累加历史（16 列，含 speed / report_type）
- `rob_data/rob_results.json` — 每船最新快照
- `rob_data/history/rob_YYYY-MM-DD.json` — 每日快照
- `rob_data/tdr_consumption.json` — TDR 设计油耗（48 船）
- `rob_data/TDR_Daily_Consumption.csv` — TDR 源文件留档

## 五、提交 / 推送约定

- 代码与说明文件（`.py` / `.html` / 本 `.md`）由本机(Leah)改完推 GitHub。
- `rob_update.bat` 跑出的数据文件（CSV / JSON / 生成的 `rob_oil_report.html`）由 culadmin 每日 13:00 / 01:00 推送，不要在本机手动推送数据。
