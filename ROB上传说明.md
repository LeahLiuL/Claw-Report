# ROB 每日更新 · 上传说明（双机协作版）

> 适用：本机（Leah）与另一台电脑同时维护 Claw-Report 仓库时的 ROB 更新 / 上传规范。
> 目的：保证「每天更新的 ROB 是**叠加（累计）数据**」，且我更新的数据能正确上网页、不被另一台电脑覆盖。

---

## 1. 数据流（必须同时上网页的两份产物）

```
Outlook 抓取
   └─> rob_data/rob_results.json        （最新快照，每船一条）
          └─> rob_oil_report.html       （网页：最新 ROB + 趋势 Tab）
          └─> combined_vessel_report.html（合并页内的 ROB Tab）

每日累计（叠加）：
   ├─> rob_data/history/rob_YYYY-MM-DD.json   （当天完整快照，按日自然累积）
   └─> rob_data/rob_history.csv              （逐船逐次，追加，绝不覆盖历史行）
```

**网页必须同时展示「最新快照」和「累计趋势」**，两份都要随更新发布。

---

## 2. 叠加（累计）数据规则（核心，不能改回）

- `rob_history.csv`：**只追加**，按 `(snapshot_time, vessel)` 幂等去重，**绝不重置 / 不覆盖历史行**。
- `rob_YYYY-MM-DD.json`：当天文件覆盖、跨天不同文件名 → 自然累积。
- `rob_results.json`（最新快照）：保留已抓到的旧数据，只更新本次成功重抓的船（`main()` 的 merge 逻辑已保证），**不能整文件清空重置**。
- 当前 `write_daily_history()` 已是「累加式」（见 `rob_refresh.py` 第 953–1008 行），**此行为必须保持**。

---

## 3. 上传（commit / push）流程

- **分支**：直接在 `main` 工作，不要各自开长期分叉。
- **拉取**：提交前先
  ```bash
  git pull --rebase --autostash
  ```
  ⚠️ 不要用 `--ff-only`（会中断）。
- **提交内容**（仅这些）：
  - 数据：`rob_data/rob_results.json`、`rob_data/history/rob_YYYY-MM-DD.json`、`rob_data/rob_history.csv`
  - 网页：`rob_oil_report.html`、`combined_vessel_report.html`（及依赖的 js/css）
  - 脚本（仅当改动）：`rob_refresh.py`、`rob_update.bat`
- **不要提交**（本地生成 / 临时）：`Vessel Bapfile.xlsx`、`bapfile.db`、`auto_bapfile.*`、`auto_bunker.*`、`site/`（若本地生成）、`*.txt` 临时文件、`_wk*.txt`。
- **推送**：`git push`（仓库已配 `store` credential helper，无需弹窗）。不要 force push，不要 reset 历史。

---

## 4. 双机协作硬规则（避免互相覆盖）

1. **同一时刻只让一台机器跑 `rob_refresh.py` 并 push**。另一台要么 `--no-outlook` 只读重建网页，要么不跑。
2. 若两台都会跑：每次运行前必须 `git pull --rebase --autostash`，跑完立即 push，缩短冲突窗口期。
3. **`rob_history.csv` 表头必须两台机器完全一致**（见第 5 点），否则追加会列错位、累计数据写坏。
4. 不要 force push，不要 reset 历史。

---

## 5. ⚠️ 必须修复的 CSV 表头不一致（当前真实存在的坑）

现状：
- 现有 `rob_history.csv` 表头 = **14 列**：
  ```
  date, vessel, code, lane, pic, lsfo, hsfo, mgo, ulsfo, bw, fw, refeer, found, report_time
  ```
- 当前 `rob_refresh.py` 的 `write_daily_history()` 的 `new_rows` **只写 13 列**（缺 `pic`/`refeer`，多 `sender`）。

后果：脚本把 13 个值塞进 14 列 → `sender` 落到 `report_time` 列、整列错位，累计 CSV 被写坏。

**统一方案（两台机器都必须同步改，推荐方案 A）**：

- **方案 A（推荐，保留已有 ~90 行历史）**：`new_rows` 改为写 14 列，顺序与表头完全一致：
  ```python
  new_rows.append([
      snap_time,
      r.get("vessel", ""), r.get("code", ""), r.get("lane", ""),
      r.get("pic", ""),                       # 补回 pic
      r.get("rob_lsfo"), r.get("rob_hsfo"), r.get("rob_mgo"),
      r.get("rob_ulsfo"), r.get("rob_bw"), r.get("rob_fw"),
      r.get("rob_refeer"),                    # 补回 refeer
      int(bool(r.get("found"))), (r.get("report_time") or "")[:19],
  ])
  ```
  `SNAP_FIELDS` 同步为上面的 14 列表头；`sender` 不入 CSV（保留在 `rob_results.json` 即可）。
- 方案 B：重生成 CSV（丢失已有累计，不推荐）。

---

## 6. 两个 ROB 抓取修复必须保留（无论哪台机器跑）

- `extract_rob()` 必须能抓 **BERTH / SAILING** 报告的 `POB ROB LSFO` 阶段前缀，**不能只认行首 `ROB `**（否则 SHENGTANG 等 BERTH 报告漏抓）。
- `build_folder_cache()` 必须**递归多层子文件夹**，不能只扫 `Vessel` 直接子文件夹。
- 已稳定、保留：X.500 发件人解析（`get_sender`）、主题 `norm` 过滤（`scan_for_rob`）、共享文件夹前缀匹配（`match_folder`）。

---

## 7. 每次上传前验证清单

- [ ] `rob_history.csv` 列数 = 14，最新日期行存在、`found` 正确（非 STALE）。
- [ ] `rob_results.json` 目标船（如 SHENGTANG）为最新报告，非 STALE。
- [ ] 网页 `rob_oil_report.html` 能打开，显示本次数据 + 趋势 Tab。
- [ ] 无 force push、无历史 reset、CSV 表头未被改动。
- [ ] 另一台电脑的 `auto_update.py` 本次未在同一窗口期覆盖本机推送。
