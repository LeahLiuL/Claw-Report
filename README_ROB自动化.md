# ROB 盘油记录网页 — 每日自动更新部署说明（culadmin）

> 网页地址（密码保护，密码 jimmy）：
> **https://leahliul.github.io/Claw-Report/rob_oil_report.html**
>
> 自动更新时间：每天 **13:00** 和 **01:00**（Windows 计划任务 `ClawReport_ROB_1300` / `ClawReport_ROB_0100`）。
> 每次自动完成：拉最新脚本 → 从 Outlook 抓各船最新 Noon/Berth/Sailing Report 的 ROB →
> 更新加密网页 → 推送 GitHub → 网页自动刷新。

---

## 前提条件（一次性确认）

| 条件 | 说明 |
|---|---|
| Git + Python 3 | culadmin 已有（daily_update.bat 在用），无需重装 |
| Python 依赖 | `pywin32`（Outlook）、`pycryptodome`（网页加密）、`openpyxl` —— 安装脚本会自动装 |
| **Outlook 全天开着** | 必须已配置 **leahliu@CULINES.COM** 邮箱（能收到船长报告），且登录状态保持。脚本通过本地 Outlook 抓邮件，Outlook 没开就抓不到（会保留上次数据并在日志里提示） |
| **不能注销/关机** | 计划任务在用户会话里跑 Outlook COM；锁定屏幕没关系，注销不行 |
| git push 凭证 | culadmin 已配置（daily_update 在用）。若 push 报错，手动进仓库目录跑一次 `git push` 完成登录 |

## 首次安装（只做一次）

```bat
cd C:\Users\culadmin\Claw-Report
git pull
install_rob_task.bat        REM 双击运行：装依赖 + 注册计划任务
```

## 验证

1. 立即手动触发一次：`schtasks /Run /TN "ClawReport_ROB_1300"`
2. 等约 2-5 分钟，看日志：`rob_data\autoupdate.log` 末尾应有 `DONE`
3. 打开网页输入密码 → 能看到 45 艘船的 ROB 表格、右上角更新时间

## 日常

**什么都不用做。** 网页每天 13:00 / 01:00 自动更新。
某船一直没有 ROB（如 DONG FANG MING HAI、NILEDUTCH LION）会一直显示 REMARK「邮箱未找到船长存油报告」——等船长发报告后会自动补上。

## 故障排查

| 现象 | 处理 |
|---|---|
| 网页没更新 | 看 `rob_data\autoupdate.log` 末尾的 `[ERROR]` 行 |
| `git pull failed` | 仓库里有本地改动（可能手动改过文件），进目录 `git stash` 后重试 |
| `rob_refresh.py failed` | 多半是 Outlook 没开或未登录；打开 Outlook 后 `schtasks /Run /TN "ClawReport_ROB_1300"` 重跑 |
| `git push failed` | git 凭证过期；手动 `git push` 一次重新登录 GitHub |
| 网页打开是旧数据 | GitHub CDN 缓存，等 5-10 分钟或强制刷新（Ctrl+F5） |

## 文件结构（抓取逻辑 vs 页面模板）

| 文件 | 职责 | 谁改 |
|---|---|---|
| `rob_refresh.py` | Outlook 抓邮件、解析 ROB、写 JSON/CSV、注入数据生成网页 | 数据侧 |
| `templates/rob.html` | 网页前端（HTML/CSS/JS，含 `__ENC__` 占位符） | 页面侧 |
| `rob_data/rob_results.json` | 每船最新一条 ROB（覆写） | 自动生成 |
| `rob_data/rob_history.csv` | 逐船逐次快照，**累加**（按 `vessel+report_time` 幂等去重） | 自动生成 |

两台电脑分工改页面 / 改数据互不冲突：改页面只动 `templates/rob.html`，改抓取逻辑只动 `rob_refresh.py`。

## 抓最新的三重保障

1. **sender 全局索引**：按 `rob_data/vessel_senders.json` 里 44 个船长邮箱，在全部文件夹预建索引（内部 Exchange 的 X.500 地址会解析回 SMTP）。
2. **多来源合并**：文件夹树 + sender 索引 + 收件箱 Restrict + 全文件夹主题搜索，候选按 `ReceivedTime` **全局倒序**，永远取时间上最新的那份。
3. **偏旧深度兜底**：常规检索后仍是 MISS、或报告时间超过 **26 小时**的船，再按 `ReceivedTime` 窗口对全部 169 个文件夹 Restrict 扫描一次（不受"每文件夹前 N 封"上限影响）。
   例：2026-08-31 自动把卡在 08-26 的 M. ODYSSEY 补到 08-30。

## 主表列（改列要同时改 4 处）

`No. / Vessel Name / Vessel Code / Lane / Bunker PIC / PIC / ROB LSFO / ROB HSFO /
ROB ULSFO / ROB MGO / Order Status / Order Details / REMARK / Special / Planned Bunkering Date / ROB Report Time`

改列必须同步：`templates/rob.html` 的表头 `<th>`、行渲染 `<td>`、`EXPORT_HEADERS`+导出行+列宽，
以及 `rob_refresh.py` 的 `build_html()` payload 和 `build_xlsx()` 表头/数据行。
（2026-08-31 加 ULSFO 列时就是改这 5 处；表头列数与行 `<td>` 数必须相等，否则整表错位。）

> 为什么需要 ULSFO 列：ASR、MEDKON DON 等船实际烧 ULSFO，主表只有三列时它们在主表里
> 显示成 LSFO 0（船上其实有 158.3 / 468.9 MT）。

## 已知注意事项

- **共用文件夹会串船数据**：如 MEDKON 文件夹同时放 MEDKON DON / MEDKON LIA 的报告，
  命中这类文件夹时会强制按主题（船名或船代码）认船。新增船文件夹时注意别让两船共用一个
  且文件夹名是其中一艘名字的前缀。
- **已下线船不显示**：船名带「已下线」后缀的船（船期表里的标记）在读入船清单时直接剔除，不进主表/趋势/Data Integrity（历史 CSV 归档行保留，不删）。
- **两台电脑不要同时跑**：只让 culadmin 的计划任务自动跑。leahliu 本机想手动刷新，跑 `python rob_refresh.py` 生成页面即可，但**不要同时 push**（会互相冲突）。
- **密码说明**：页面数据经 AES 加密，网页源码看不到明文。但本仓库是公开的，密码写在脚本里（jimmy）——知道仓库地址的人可以推出密码。这是「防路人」级别，不是安全级别；如需更强隔离请把仓库转 Private（GitHub Pages Private 仓库需 Pro 账号）。
- **新船 / 船退出**：船清单每天从 `cul_daily_movement.html`（Daily Movement 网页数据）自动解析，新船自动加入、退出的船自动消失，无需改脚本。
- **共用邮件文件夹**：如 "MEDKON" 文件夹同时放两船邮件，脚本已按主题过滤防误抓；若发现某船数据异常，对照该船最近邮件主题确认。
- **Excel 副本**：culadmin 上会生成 `rob_data\rob_oil_table.xlsx`（不入库）；leahliu 本机额外同步到 `C:\CULINES\Claw Report\盘油记录.auto.xlsx`。
