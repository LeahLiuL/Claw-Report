# Container Explorer · 箱量智能查询台

一个**纯静态、零后端**的网页工具：把你的大体积 Excel 发货数据转换成轻量 `data.json`，
浏览器端即可按 **VVD / LANE / 箱号** 筛选，支持**去重**与**导出专业 Excel**。
可一键部署到 GitHub Pages，并通过 GitHub Actions 每天从 FTP 自动拉取更新。

## 功能

- 展示指定 16 列：`VVD, LANE, CONT NO., FULL/EMPTY, TYPE/SIZE, POL, POD, CONT_Weight, AWK, DG, RF, BB, SLOT_OPR, CONT_OPR, REVENUE MONTH, TARGET_PORT`
- **筛选**：VVD / LANE 子串匹配 + 粘贴箱号（换行 / 逗号 / 空格分隔，精确匹配）
- **去重（Dedup）**：同一箱号出现多条时
  1. 优先取 `TARGET_PORT == POL` 的那一行；
  2. 若仍有重复，取 `REVENUE MONTH` 最大的一行；
  3. 若没有任何一行满足 `TARGET_PORT == POL`，则在全部记录中取 `REVENUE MONTH` 最大的一行。
- **导出 Excel**：把当前筛选结果导出为带样式（加粗冻结表头、自动列宽、自动筛选）的 `.xlsx`
- 响应式 + 无障碍（WCAG AA，键盘可达、跳转链接、aria-live）

## 本地预览

```bash
# 生成演示数据（含用于演示去重的重复箱号）
python scripts/generate_sample.py

# 起一个本地静态服务器（file:// 直接打开会被浏览器拦截 fetch）
python -m http.server 8000
# 打开 http://localhost:8000
```

试试粘贴这几个箱号观察去重效果：
`DEMO000111`、`DEMO000222`、`DEMO000333`

## 接入你的真实 Excel

```bash
pip install openpyxl
# 普通文件（小）：
python scripts/convert.py 你的文件.xlsx --out data.json --meta meta.json
# 大文件（几十万行以上，强烈建议 --gzip）：输出 data.json.gz，体积约为明文 1/8
python scripts/convert.py 你的文件.xlsx --out data.json.gz --meta meta.json --gzip
# 调试可用 --limit 5000 先转一小段估算体积
```
表头匹配对大小写 / 空格 / 标点不敏感，常见别名（如 `Container No.` / `POL` / `Port of Loading`）都能识别。
本项目的真实数据源表头（`VVD, CONT_NR, FE_FLG, POL_CD, TAGERT_PORT`…）已内置兼容。

> 页面会自动读取 `meta.json` 的 `compressed` 字段：为 `true` 时加载 `data.json.gz`（浏览器端用 `DecompressionStream` 解压），否则加载明文 `data.json`。

## 部署到 GitHub Pages

本仓库用 **官方 Pages 部署流**（`upload-pages-artifact` + `deploy-pages`），
数据走 Artifact 而非 git 历史，因此**每天 13MB 刷新也不会撑大仓库**。

1. 把本仓库推到 GitHub（`main` 分支已含完整站点 + 当前 `data.json.gz`）。
2. **Settings → Pages → Build and deployment → Source** 选择 **GitHub Actions**。
3. 添加 SFTP 凭据到 **Settings → Secrets and variables → Actions → Secrets**：
   | Secret | 值 |
   |---|---|
   | `FTP_HOST` | `10.5.4.2` |
   | `FTP_PORT` | `6622` |
   | `FTP_USER` | `leah` |
   | `FTP_PASS` | `Fine@B!` |
   | `FTP_REMOTE` | 远端 xlsx 路径，例如 `/Master Data - Leah/近4月VesselBapfile.xlsx` |
4. 之后：
   - 推送 `main` 会立刻重新部署（用仓库内现有 `data.json.gz`，**无需 SFTP 即可上线**）。
   - `deploy.yml` 每天 **UTC 02:00（上海 10:00）** 尝试从 SFTP 拉新 Excel → 转换 → 重新部署；
     **SFTP 连不上时自动沿用上一次数据**，站点不会变空。
   - 也可在 **Actions → Build & Deploy Container Explorer → Run workflow** 手动立即刷新。

> 链接形态：`https://<用户名>.github.io/Claw-Report/`，任何人点开即可使用。
> 若要让「别的用户」无需登录就能访问，请把仓库设为 **Public**（Settings → General → Change repository visibility）。
> 私有仓库的 Pages 仅协作者可见。

## 大体积数据说明

- 真实数据源「近4月VesselBapfile.xlsx」约 **65.5 万行**。明文 `data.json` 约 100MB+，超限且加载慢；
  用 `--gzip` 压缩后 `data.json.gz` 约 **13MB**（实测 5000 行=0.1MB），远低于 GitHub Pages 100MB 单文件限制，下载快。
- 转换脚本已改为**流式写入**（不再一次性把全量读进内存），可在 CI runner 上稳定处理 65 万行。
- 浏览器端用 `DecompressionStream('gzip')` 解压后筛选：对 65 万行做 VVD/LANE 子串 + 箱号精确匹配均为毫秒级；
  表格分页渲染（每页 50 行），导出的是**全部**筛选结果。
- 若日后数据再涨一个数量级，可平滑升级为：FTP 侧按月拆分多文件 + 前端按需加载，或 IndexedDB + Web Worker 增量索引。

## 文件结构

```
index.html            # 页面
styles.css            # 样式
app.js                # 筛选 / 去重 / 导出逻辑
scripts/convert.py    # Excel -> data.json + meta.json
scripts/generate_sample.py  # 生成演示数据
data.json.gz          # 当前数据（gzip，仓库内已含完整 65 万行；CI 每日覆盖）
meta.json             # 数据元信息（更新时间、行数、是否压缩）
.nojekyll            # 禁用 Jekyll，确保 _site 直接以静态文件服务
.github/workflows/deploy.yml  # 构建 + 每日 SFTP 同步 + Pages 部署
```
