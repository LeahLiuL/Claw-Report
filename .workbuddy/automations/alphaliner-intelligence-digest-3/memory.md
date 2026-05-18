# Alphaliner Intelligence Digest — Automation Memory
**Automation ID**: alphaliner-intelligence-digest-3  
**Schedule**: Every Monday 09:00 (Asia/Shanghai)  
**Output**: `reports/alphaliner-digest.html` → GitHub Pages  
**URL**: https://leahliul.github.io/Claw-Report/reports/alphaliner-digest.html  

---

## Execution History

### 2026-05-18 | Week 21 ✅
- **Status**: SUCCESS
- **Data source**: Web search (Outlook/Maton API key not available; neodata 401)
- **Key data**:
  - Drewry WCI: $2,553/FEU (+12% WoW, 2026年最大单周涨幅)
  - SCFI: 2,140.66 (+186.45, 首次突破2,000)
  - CCFI: 1,280.46 (+0.1%)
  - TPEB估算: ~$3,800/FEU (Seavantage)
  - 霍尔木兹: 实质封锁（Operation Epic Fury后）
  - MSC: 1,000艘里程碑
  - 空白航次: 30班取消(Wk21-25), ~4.3%东西向运力
  - 关税缓和: 90天窗口(4/22–7/中旬), 均值~33%
- **Git commit**: `7abe103` — [Claw Auto] Alphaliner Digest Week 21 2026
- **Email**: leahliu@culines.com — EMAIL_SENT_SUCCESS
- **HTML size**: 72,606 bytes
- **Articles**: 10篇，涵盖运价/运力/航线/法务/贸易

### 2026-05-05 | Week 19 ✅
- **Status**: SUCCESS (previous run)
- **Report**: alphaliner-intelligence-digest-wk19-2026.html → also saved to alphaliner-digest.html
- **Key data**: WCI $2,286 (+3%), SCFI 1,954, CMA CGM San Antonio导弹袭击事件

---

## Known Issues / Notes
- `MATON_API_KEY` 环境变量在系统中未配置，Outlook API调用不可用
- neodata-financial-search skill 返回401，需要有效token
- 当两个数据源不可用时，使用 web search 作为数据来源
- Python脚本路径: `C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe`
- 生成脚本模式: 将HTML内容写入Python脚本变量后执行，规避token限制
