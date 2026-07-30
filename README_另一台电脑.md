# CUL Daily Movement 每日更新 — 另一台电脑操作说明

> 适用机器：culadmin（数据在 Z 盘）。本说明随仓库分发，每次 `git pull` 都会拿到最新版。

## 首次准备（只做一次）

1. 安装 **Git**（勾选 "Add to PATH"）和 **Python 3**（勾选 "Add to PATH"）。
   装完后打开命令行运行：
   ```bat
   pip install openpyxl
   ```
2. 在合适目录克隆仓库：
   ```bat
   git clone https://github.com/LeahLiuL/Claw-Report.git
   ```
3. 确认数据目录存在且原始文件在：
   ```
   Z:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新\
   ```
   里面需有原始的 `CUL DAILY MOVEMENT.xlsx`。

## 之后每天

- 双击 `Claw-Report` 文件夹里的 **`daily_update.bat`**
- 它会自动完成：拉最新脚本 → 刷新 PIC 表 → 重建大 Excel → 生成网页 → 推送
- 网页在 GitHub Pages 上自动更新，顶部显示 `Updated 年-月-日 时:分`（即本机当天跑脚本的时刻）

## 注意事项（建议贴显示器）

- **改 PIC**：打开 Z 盘 `更新\PIC汇总.xlsx`，只改 "PIC" 列，其余别动；保存后跑 bat 自动生效。
- **船名↔代码（vessel.csv）**：在 GitHub 网页上改最省事；要本地改就改仓库里的 `vessel.csv`，跑 bat 会自动推上去。
- **报错处理**：bat 中途若弹出红色报错并停住，别关窗口，截图发回。
- **不要手动改仓库里的 `.py` / `.bat` 脚本**，否则下次自动更新会卡住。

## 备注

- 若本机早已 clone 并跑过，跳过"首次准备"1–3 步，直接双击 `daily_update.bat` 即可（内置 `git pull` 会自动拉到最新脚本与网页更新时间）。
- `vessel.csv` 建议固定在一处维护（本机 / 另一台 / GitHub 网页），避免两边同时改同一行导致 pull 冲突。
- GitHub Pages 发布源为 **main 分支根目录**，历史遗留的 `gh-pages` 分支已清理，无需关心。
