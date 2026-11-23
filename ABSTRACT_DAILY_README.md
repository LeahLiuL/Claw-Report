# Abstract Log 每日自动更新 — 操作说明

## 一句话

每天自动从 Outlook 抓 Abstract Log 邮件 → 更新 `abstract_log.db`（0.5M）→ 传到 GitHub → 另一台电脑下载后生成页面。

**reports.db（189M）不参与每日更新**，保持 11-23 快照不变。

---

## 一、这台电脑（culadmin，有 Outlook）——不用做任何事 ✅

Abstract 更新**已经嵌入现有的 ROB 计划任务**，跟着 ROB 一起跑，无需新建任务、无需双击任何东西：

| 现有任务 | 频率 | 触发 |
|---|---|---|
| `ClawReport_ROB_0100` | 每天 01:00 | ROB 跑完 → 自动接着跑 Abstract |
| `ClawReport_ROB_1300` | 每天 13:00 | 同上 |

实现位置：`rob_update.bat` 在 `[3/5]` 后调用 `abstract_core.bat`（采集 → 上传 → 重生成公告页），
随后 ROB 原有的 `git_safe_push` 会把更新后的 `DATABASES.html` 一起推送。

> Abstract 步骤失败不会影响 ROB；反之亦然。日志独立：`rob_data\abstract_daily.log`

想立即手动跑一次（在仓库目录）：

```
python abstract_daily_sync.py --days 5 --apply
python upload_abstract_db.py --apply
python databases_page_gen.py
```

> 备注：也保留了一个独立安装脚本 `install_abstract_task.bat`（每天 06:00 单独跑），
> 但**正常不需要用**——已经嵌入 ROB 任务了。只有当你想把 Abstract 和 ROB 彻底分开时才需要它。

---

## 二、另一台电脑（生成页面用）——每天做的事

每天从固定地址重新下载数据库（覆盖本地旧文件）：

```
https://github.com/LeahLiuL/Claw-Report/releases/download/db-snapshot-2026-11-23/abstract_log.db
```

地址**永远不变**，每次下载都是最新数据（覆盖同名 asset）。然后跑生成脚本：

```
python abstract_log_report_gen.py
```

生成 `abstract_log_report.html`，用浏览器打开即可。

> 想省事的话，可以在另一台电脑也写个 bat：`curl 下载 db → python 生成页面`，双击就更新。

---

## 三、线上地址（都验证过可用）

| 内容 | 地址 |
|---|---|
| Abstract Log 分析页 | https://leahliul.github.io/Claw-Report/abstract_log_report.html |
| 数据库公告中心（含下载链接） | https://leahliul.github.io/Claw-Report/DATABASES.html |
| abstract_log.db 下载 | https://github.com/LeahLiuL/Claw-Report/releases/download/db-snapshot-2026-11-23/abstract_log.db |

> 注意：分析页 `abstract_log_report.html` 本身**不再每天自动重建**（按你的要求）。
> 它显示的是生成那一刻的数据。要看最新数据 = 下载最新 db → 自己生成页面。
> 公告页 `DATABASES.html` 每天会更新 `last_parse` 时间戳，可用来确认数据新鲜度。

---

## 四、文件清单

| 文件 | 作用 |
|---|---|
| `abstract_daily_sync.py` | 从 Outlook 抓 abstract 邮件 → 增量入库（核心采集） |
| `upload_abstract_db.py` | checkpoint WAL + 覆盖上传线上 db（不碰 reports.db） |
| `databases_page_gen.py` | 重新生成公告页（含新时间戳） |
| `abstract_core.bat` | 核心三步的 bat（无 git 操作），被 rob_update.bat 调用 ← **实际每天运行的就是它** |
| `rob_update.bat` | ROB 定时任务，末尾已嵌入 abstract_core 调用（无需新建任务） |
| `abstract_daily_update.bat` | 独立完整入口（含 git push），备用 |
| `install_abstract_task.bat` | 独立任务安装脚本，备用（正常不需要） |

---

## 五、关键设计（为什么这样）

1. **mail_id 复用**旧规则（EntryID md5 前 20），同一封邮件覆盖不重复。
2. **附件智能识别**：主题含 ABSTRACT 的邮件，附件可能是 CII/NOON；先按文件名找 ABSTRACT，找不到就逐个 xlsx 校验模板。
3. **纯回复邮件自动排除**（只有签名图片，无 xlsx）。
4. **Outlook 缓存门控**：mode>=700 才扫描，没同步完退出让下一轮回补（--days 5 窗口）。
5. **上传前 WAL checkpoint**：否则复制/上传的 db 可能缺数据。
6. **git 锁**：和 ROB 共用锁机制，owner=abstract，避免并发冲突。

---

## 六、验证记录（2026-11-23）

- dry-run：扫 42 文件夹，识别 6 封（4 封已入库跳过、2 封为纯回复无附件），逻辑正确。
- apply 验证（临时库）：3 封正式报告全部解析入库（voy + legs 字段完整）。
- 覆盖上传：DELETE 旧 asset → POST 新 asset，`state=uploaded`。
- 线上回读：271 航次 / 1090 航段 / last_parse=2026-11-23T15:50:01，数据完整。
