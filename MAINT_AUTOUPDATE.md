# Maintenance 数据跨机器自动更新 — 操作说明（culadmin）

## 背景

Maintenance 视图的数据源 `Vessel_Schedule_Maintain_Over_Time_Port_Log.xlsx`
原本只在 leahliu 本机的 `C:\CULINES\Claw Report\` 目录。

culadmin（自动跑 daily update 的机器）挂的是 **Z 盘，不是 C 盘**，没有该路径，
所以它在没有这个文件时生成的看板，Maintenance 会**整块空白**。

`gen_html.py`（自 commit `40319f3` 起）已改为：

- 本机有 `C:\CULINES` 副本 → 用本地副本（leahliu，免 VPN）
- 否则 → 从 **SFTP（10.5.4.2:6622，凭据同 `sftp_fetch.py`）** 拉取，存到 `.cache/`（已被 `.gitignore` 忽略）

本文档面向 **culadmin**，确保其自动更新也能带上 Maintenance 数据。

---

## 前置条件（一次性确认）

1. **VPN 已连接** —— `10.5.4.2` 是内网 IP，没连 VPN 会连接超时（WinError 10060）。
2. **python 已装 `paramiko`** —— `pip install paramiko`（Bapfile 的 SFTP 管线已在用，通常已具备）。
3. **`gen_html.py` 是最新** —— 已 `git pull`，版本 ≥ `40319f3`。

---

## 验证（一次性）

在 culadmin 上进入 Claw-Report 目录，跑预检脚本：

```bat
cd C:\path\to\Claw-Report
python check_maint_env.py
```

- 看到 **`RESULT: GO ✅ auto-update will populate the Maintenance view`** → 一切就绪。
- 看到 **`NO-GO` / 任何 `FIX:` 提示** → 按提示处理（连 VPN / 装 paramiko / 核对 `MAINT_SFTP_REMOTE`），再重跑直到 GO。

---

## 自动更新流程（daily update）

确保自动化在执行 `gen_html.py` 之前满足：

1. **已连 VPN**（SFTP 拉 MAINT 必需）
2. **已 `git pull`** 最新代码（含 SFTP 回退逻辑）
3. 正常跑 `gen_html.py` → 重新生成 `cul_daily_movement.html` → 推送

可选加固：在 `gen_html.py` **之前**先跑 `check_maint_env.py`，若返回非 0（NO-GO）
则**中止本次发布**，避免一个空白 Maintenance 的版本上线：

```bat
python check_maint_env.py || exit /b 1
python gen_html.py
```

---

## SFTP 路径变更

若 SFTP 上 MAINT 文件路径变化，**无需改代码**，用环境变量覆盖：

```bat
set MAINT_SFTP_REMOTE=/新/路径/Vessel_Schedule_Maintain_Over_Time_Port_Log.xlsx
```

```bash
export MAINT_SFTP_REMOTE=/新/路径/Vessel_Schedule_Maintain_Over_Time_Port_Log.xlsx
```

其他可调环境变量：`MAINT_SFTP_HOST` / `MAINT_SFTP_PORT` / `MAINT_SFTP_USER` / `MAINT_SFTP_PASS`。

---

## 故障速查

| 现象 | 原因 | 处理 |
|---|---|---|
| `UNREACHABLE ... 10060` | VPN 未连 / 内网不可达 | 连 VPN 后重跑 |
| `MISSING: No module named 'paramiko'` | 环境缺 paramiko | `pip install paramiko` |
| `NOT FOUND / ERROR` + `MAINT_SFTP_REMOTE` | SFTP 路径不对 | 用环境变量改成正确路径 |
| 生成后 Maintenance 仍空 | 走了本地 fallback 但本地也无文件 | 检查 SFTP 凭据/路径；跑预检脚本定位 |
