# CUL Bapfile 静态站 — 迁移到 culadmin 机器

> 目标：把每日「SFTP 下载 → 增量叠加 → 重生成 → 部署 GitHub Pages」的自动化
> 从 leahliu 机器（晚上关机）迁到 culadmin 机器（一直开机），保证每天稳定跑。

## 前提（culadmin 机器已具备）
- Windows，用户名 `culadmin`
- WorkBuddy（同账号，已装）
- Python 3.13 + pip
- git
- VPN（能连 10.5.4.2 内网）

## 步骤

### 1. 部署文件到 `C:\Users\culadmin\Claw-Report`

**方式 A（用打包好的 zip，最简单）**：
把 `migrate_culadmin.zip` 解压到 `C:\Users\culadmin\Claw-Report`
（含 5 个脚本 + bapfile.db + .gitignore + 本文档）

**方式 B（git clone + 单传 db）**：
```bat
git clone https://github.com/LeahLiuL/Claw-Report.git C:\Users\culadmin\Claw-Report
```
然后单独把 `bapfile.db`（373MB）从 leahliu 机器
`C:\Users\leahliu\Claw-Report\bapfile.db` 用 U 盘/共享拷到
`C:\Users\culadmin\Claw-Report\bapfile.db`。

> ⚠️ bapfile.db 是累计库（已删 2025-12 前旧数据，现 1,528,097 行 / 391MB），
> 必须传过去；不能让新机器从空库重建（SFTP 是滚动窗口，没有历史）。

### 2. 装 paramiko
```bat
pip install paramiko
```

### 3. 配 git 凭证（能 push cul-bapfile-site）
```bat
git clone https://github.com/LeahLiuL/cul-bapfile-site.git C:\Users\culadmin\cul-bapfile-site
```
若提示登录，用 GitHub 账号 LeahLiuL 授权。clone 成功即凭证 OK。

### 4. 连 VPN
连上公司 VPN（10.5.4.2 是内网 IP，不连会超时）。

### 5. 手动测试一次
```bat
cd C:\Users\culadmin\Claw-Report
python build_deploy.py
```
应依次：SFTP 下载 → 增量叠加 → 生成分片 → 部署 cul-bapfile-site/main。
耗时约 10-15 分钟。成功后 https://leahliul.github.io/cul-bapfile-site/ 数据更新。

### 6. 在 culadmin 的 WorkBuddy 里建自动化
在 culadmin 机器的 WorkBuddy 对话里，让 AI 创建一个自动化，参数如下：

| 字段 | 值 |
|---|---|
| name | CUL Bapfile 每日增量叠加部署 |
| scheduleType | recurring |
| rrule | `FREQ=DAILY;BYHOUR=10` |
| cwds | `C:\Users\culadmin\Claw-Report` |
| status | ACTIVE |

**prompt**（路径已改 culadmin，复制给 AI）：
```
每天把公司内网 SFTP 上最新的 Vessel Bapfile.xlsx【增量叠加】进本地累计库，并重新生成静态站点部署到 GitHub Pages（公开地址 https://leahliul.github.io/cul-bapfile-site/）。数据源是 SFTP 且【每天更新、滚动窗口会与已有数据重叠】，因此必须【增量叠加 + 整行去重】，绝不能全量替换。严格按下面顺序：1) 确认本机已连 VPN（SFTP 主机 10.5.4.2 是内网 IP，未连 VPN 会连接超时）；2) 在 Bash 中 cd 到 C:/Users/culadmin/Claw-Report，用 Python 运行 build_deploy.py（该脚本会依次：先 SFTP 下载最新 xlsx → 用 process_all.py --append 把 xlsx【增量叠加】进 bapfile.db，自动跳过所有「17 个数据列完全一致」的整行重复项 → 用 gen_static.py 重新生成按箱号前缀与月份的 gzip 静态分片到 site/ → 用 deploy_site.py 把站点部署到公开仓库 LeahLiuL/cul-bapfile-site 的 main 分支，该仓库已开启 GitHub Pages 且源设为 main）；3) 运行后确认站点已更新：https://leahliul.github.io/cul-bapfile-site/manifest.json 可访问（HTTP 200）。完成后用中文简短汇报结果（成功/失败，若输出了新增行数、去重删除行数、箱号数也一并说明）。注意：整个过程可能耗时 10-15 分钟，且第 2 步依赖 VPN，请耐心等待命令完成，勿中途中断；若 SFTP 下载失败提示 WinError 10060，请明确报告"VPN 未连接"并停止，不要继续后续步骤。
```

### 7. 停 leahliu 本机的自动化
culadmin 跑通后，在 leahliu 本机把自动化 `automation-1784020771490` 设为 **PAUSED**，
避免两台同时跑产生重复提交/冲突。

## 脚本说明（已改可移植，拷到 culadmin 不用改代码）
- `build_deploy.py` — 主入口：sftp_fetch → process_all --append → gen_static → deploy_site
- `sftp_fetch.py` — 从 SFTP 下载最新 xlsx（10.5.4.2:6622，用户 leah）
- `process_all.py` — 增量叠加 xlsx 进 db，整行去重（17 列唯一索引）
- `gen_static.py` — 从 db 生成 cont/month 分片 + manifest.json
- `deploy_site.py` — 部署 site/ 到 cul-bapfile-site/main

路径全部用「脚本所在目录」定位，不依赖用户名。环境变量可覆盖：
- `BAPFILE_LOCAL` — xlsx 本地路径（默认脚本目录下）
- `BAPFILE_DB` — db 路径（默认脚本目录下 bapfile.db）
- `DEPLOY_WT` — 部署用的 clone 目录（默认脚本目录的兄级 cul-bapfile-site）
