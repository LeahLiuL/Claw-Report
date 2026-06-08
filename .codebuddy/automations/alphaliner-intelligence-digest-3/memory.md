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

### 2026-06-01 (Week 23) — 09:00 CST
- **Status**: ✅ HTML生成 + Git推送成功；⚠ 邮件通知跳过
- **HTML file**: `reports/alphaliner-digest.html` 已更新为 Week 23 (May 26 – Jun 1, 2026)
- **Git commit**: `480112d` — "Update Alphaliner Digest - Week 23 2026"
- **GitHub Pages**: Push 成功 `9a2c63a..480112d main -> main`
- **GitHub URL**: https://leahliul.github.io/Claw-Report/reports/alphaliner-digest.html
- **Data sources used**: SSE SCFI, Drewry WCI, Alphaliner Top100, 财联社, 东方财富, 同花顺, 新华财经
- **Key data (2026-05-29)**:
  - SCFI: 2,571.73 (+353.58, +15.94% WoW) — 2026年最大单周涨幅
  - Drewry WCI: $2,800/FEU (+3.24% WoW, 月涨+26.35%)
  - CCFI: 1,366.76 (+3.75%)
  - Europe: $2,745/TEU (+29.93%)
  - Mediterranean: $3,750/TEU (+16.9%)
  - US West Coast: $4,419/FEU (+31.5%)
  - US East Coast: $5,333/FEU (+23.6%)
  - Persian Gulf: $4,462/TEU (小幅走高) [★ CUL core route]
- **Notable items**: 旺季全面提前引爆，欧美线爆舱甩柜常态化；四大船司6月1日集体GRI；霍尔木兹24艘/日护航通行；Drewry Wk23-27共47班空白航次
- **Email status**: 未配置SMTP，Outlook COM不可用，通知跳过

### 2026-06-08 (Week 24) — 09:00 CST
- **Status**: ✅ HTML生成 + Git推送成功 + Outlook邮件发送成功
- **HTML file**: `reports/alphaliner-intelligence-digest-wk24-2026.html` — Week 24 (Jun 2–8, 2026)
- **Git commit**: `77f5b8b` — "Add Alphaliner Intelligence Digest - Week 24 2026 (Jun 2-8)"
- **GitHub Pages**: Push 成功 `1afaee5..77f5b8b main -> main`
- **GitHub URL**: https://leahliul.github.io/Claw-Report/reports/alphaliner-intelligence-digest-wk24-2026.html
- **Data sources used**: Alphaliner Top100, Drewry WCI, SSE SCFI, Kuehne+Nagel Hormuz Update, HandyBulk, 财联社, 雪球
- **Key data (2026-06-05/07)**:
  - SCFI: 2,726.48 (+154.75, +6.0% WoW) — 五连涨，5月以来累计+23%
  - Drewry WCI: $3,433/FEU (+23% WoW) — 2026年最大单周涨幅
  - WCI Shanghai→LA: $4,565/FEU (+31%)
  - WCI Shanghai→NY: $5,505/FEU (+20%)
  - WCI Shanghai→Rotterdam: ~$3,050/FEU (+22%)
  - WCI Shanghai→Genoa: ~$4,100/FEU (+18%)
  - SCFIS Europe: 2,038.09 (+9.4%)
  - Alphaliner: 全球7,545艘/3,421万TEU
  - Hormuz: 日通行10艘(11%)，船公司仍绕航
- **Notable items**: 旺季全面爆发提前4-6周；Drewry WCI单周暴涨23%创年度纪录；6月GRI成功落地$1,450-1,500/FEU；MSC追加6月中PSS $600-800/FEU；跨太平洋舱位利用率98-100%；甩柜率上升；缺箱蔓延；霍尔木兹危机100天+无缓解
- **Outlook status**: COM接口本次成功发送邮件至leahliu@culines.com（含附件）

### 2026-05-19 (Week 21) — Prior run reference
- Week 21 report had been generated and deployed before this run
- This run successfully overwritten with Week 22 content

## Known Issues / Notes
- Outlook win32com: 在自动化执行上下文（无用户交互会话）中，COM接口调用通常会超时。不适合用于定时任务中的邮件发送。
- 建议改用 SMTP / PowerShell `Send-MailMessage` 或 Exchange Web Services 替代 Outlook COM
- Alphaliner官网需登录，公开页面数据有限；主要通过搜索引擎获取公开财经媒体数据重建内容
- Z盘共享文件夹在自动化环境下不可访问（需映射网络驱动器的用户会话）
