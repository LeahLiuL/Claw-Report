# Automation Memory — alphaliner-intelligence-digest-3

## Execution History

### 2026-05-25 (Week 22) — 08:57 CST
- **Status**: ✅ HTML生成 + Git推送成功；⚠ Outlook邮件通知跳过
- **HTML file**: `reports/alphaliner-digest.html` 已更新为 Week 22 (May 19–25, 2026)
- **Git commit**: `530ecd5` — "Update Alphaliner Digest - Week 22 2026"
- **GitHub Pages**: Push 成功 `c57264e..530ecd5 main -> main`
- **GitHub URL**: https://leahliul.github.io/Claw-Report/reports/alphaliner-digest.html
- **Data sources used**: SSE SCFI, Drewry WCI, Alphaliner Top100, Linerlytica, 财联社 (Cailian Press), 同花顺
- **Key data (2026-05-22)**:
  - SCFI: 2,218.15 (+3.6% WoW)
  - Drewry WCI: $2,712/FEU (+6%)
  - CCFI: 1,317.36 (+2.9%)
  - Europe: $1,905/TEU (+4.9%)
  - US West Coast: $3,154/FEU (+1.2%)
  - Persian Gulf: $4,306/TEU (+4.2%) [★ CUL core route]
- **Notable items**: MSC首破500万TEU里程碑; 船司集中空班压缩6月供给至26万TEU/周; 旺季提前引爆爆舱
- **Outlook status**: COM接口在自动化环境下超时（Outlook进程未在用户会话前台运行），邮件通知未发出
- **Email workaround**: 下次可考虑用 PowerShell Send-MailMessage 或 SMTP 方式替代 win32com 发送邮件

### 2026-05-19 (Week 21) — Prior run reference
- Week 21 report had been generated and deployed before this run
- This run successfully overwritten with Week 22 content

## Known Issues / Notes
- Outlook win32com: 在自动化执行上下文（无用户交互会话）中，COM接口调用通常会超时。不适合用于定时任务中的邮件发送。
- 建议改用 SMTP / PowerShell `Send-MailMessage` 或 Exchange Web Services 替代 Outlook COM
- Alphaliner官网需登录，公开页面数据有限；主要通过搜索引擎获取公开财经媒体数据重建内容
- Z盘共享文件夹在自动化环境下不可访问（需映射网络驱动器的用户会话）
