# MEMORY.md — 长期记忆

## Alphaliner Intelligence Digest 自动化

### 固定链接覆盖清单
每次生成周报后，必须**同时覆盖**以下三个 HTML 文件：

| 文件 | 用途 |
|------|------|
| `reports/alphaliner-intelligence-digest-wk{ISO_WEEK}-2026.html` | 按周归档（主文件） |
| `reports/alphaliner-intelligence-digest.html` | 固定链接版（始终最新） |
| `reports/alphaliner-digest.html` | 短链接版（始终最新） |

**推送顺序**：先 `copy` 三个文件，再一次性 `git add` + `git commit` + `git push`。

### 执行时间
- 自动化ID：`alphaliner-intelligence-digest-5`
- 每周四 08:30 执行
- 邮件脚本周数 = 当前 ISO 周（不是 `iso_week - 1`，因为周四执行覆盖当前周）

### 数据管线
1. Outlook 抓取（`fetch_alphaliner_email.py --last 2`）
2. PDF 文本提取（pdfplumber）
3. 公开数据搜索（SCFI / WCI / 燃油 / 新闻）
4. HTML 报告生成
5. GitHub Pages 推送
6. 邮件通知（`send_alphaliner_via_win32com.py`）

### 邮件来源
- 发件人：Mabel Ong / CUL-SG/NPD
- 标题：Internal circulation of Alphaliner weekly newsletter

---

## 项目约定

- 用户：Leah Liu，中远海运CULines租船团队
- 语言：中文交流，表格化输出
- 报告风格：Alphaliner 用浅色主题 + CUL金色logo，标注来源、日期、"AI解读仅供参考"
- PPT 忌用红色
- 文件路径使用正斜杠 `/`
- 开发环境：Node.js v22.12.0，Python 3.13.12

---

## Bapfile 站点每日自动化

### 数据管线
1. SFTP 下载 `Vessel Bapfile.xlsx`（`sftp_fetch.py`，端口 6622，需 VPN）
2. 增量叠加进 `bapfile.db`（`process_all.py --append`，17 列整行去重）
3. 生成 gzip 静态分片（`gen_static.py`：按箱号前缀、航线、月份）
4. 部署到 `LeahLiuL/cul-bapfile-site` main 分支（`deploy_site.py`）
5. 公开地址：https://leahliul.github.io/cul-bapfile-site/

### 故障应急
- 若 SFTP 文件为截断 ZIP（缺中央目录），使用 `repair_bapfile.py` 重建有效 xlsx：
  ```bash
  python repair_bapfile.py <corrupt.xlsx> <repaired.xlsx>
  ```
- 该脚本解析本地文件头、解压 deflate 流、验证 CRC，并在最后一个完整 `</row>` 处安全截断截断的 sheet XML。
