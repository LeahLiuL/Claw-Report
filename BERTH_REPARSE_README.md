# BERTH ROB 全量回刷操作说明（给 culadmin）

## 背景
2026-11-19 修复了 BERTH 报告 ROB 取值规则：旧规则「同一油种后值覆盖前值」取到了附件表格
最下方的 ANCHOR AWEIGH（起锚）/ DROP ANCHOR（抛锚）ROB，而航次期初应该用靠泊 POB ROB。
例：CUL XIAMEN 2625 的 09-13 BERTH，POB LSFO=635.57，却被起锚值 651.18 覆盖。

新解析规则已上线（`rob_refresh.extract_rob`，POB 优先），但只影响**今后抓取**。
历史 CSV 里 3600 条 BERTH 行仍是旧值（三重去重都会跳过已存在记录），需要跑一次本脚本强制回刷。

**脚本**：`backfill_berth_rob.py`（已在仓库，commit 65de095）
**注意**：本机（leahliu）没有这些邮件，无法实测邮件定位；请在 culadmin 那台跑。

## 前置条件
1. OUTLOOK.EXE 已启动并加载到主界面（否则 COM 会挂起）。
2. 同一时间不要跑其它访问 Outlook 的脚本（harvest/backfill/rob_refresh），避免并发 RPC 崩溃。
3. 建议避开 13:00 / 01:00 的自动批次。

## 步骤 1：先单船验证定位是否正确（必做）
在仓库目录用统一 Python 跑（只读，不写盘）：

```
C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe backfill_berth_rob.py --dry-run --vessel "CUL XIAMEN"
```

看输出：
- `目标 BERTH 行` 应为该船的历史条数（XIAMEN 约 223）。
- `命中 X / Y 封` —— X 应接近 Y。若 X 远小于 Y（比如命中个位数），说明邮件定位不到（可能邮件被移走或不在 these folders），**先不要 --apply**，把输出发回给我。
- `[diff] ... -> {...}` 列出会被改的行。XIAMEN 09-13 那条我们已手动改成 635.57，应该**不再出现**在 diff 里（说明脚本认为它已正确）；如果别的行有 diff，就是新规则找到的其它取值错误。

## 步骤 2：全量预演（看整体命中率和改动面）
```
...\python.exe backfill_berth_rob.py --dry-run
```
关注最后一行 `命中 X / 3600 封, 有差异 N 条`。把这几行输出发回给我确认。

## 步骤 3：先小批正式回刷（可选但推荐）
```
...\python.exe backfill_berth_rob.py --apply --limit 5
```
只改前 5 条差异。然后 `git diff rob_data/rob_history.csv` 检查：应只改 BERTH 行的 lsfo/hsfo/mgo/ulsfo 等 oil 列，不动其它行/列。

## 步骤 4：全量正式回刷
```
...\python.exe backfill_berth_rob.py --apply
```
脚本会自动备份 CSV 到 `rob_data/tpl_backup/berth_reparse_<时间戳>/` 再写。

## 步骤 5：提交推送 + 重算
回刷只改了 `rob_history.csv`，还需让页面重新计算航次油耗：
1. 检查 `git diff --stat`，确认只有 rob_history.csv 变化。
2. 正常跑一轮 `rob_update.bat`（会 commit+push 改后的 CSV，并用新数据重算页面）。
   ——或者手动提交后跑 `rob_refresh.py --no-outlook` 重生成页面再走 safe push。
3. 到页面核对几个已知航次，确认数字变化符合预期。

## 护栏（脚本内置，无需操作）
- 默认 dry-run，不显式 `--apply` 绝不写盘。
- 只改命中 BERTH 行的 oil 列（lsfo/hsfo/mgo/ulsfo/bw/fw），不新增/删除行，不碰 NOON/SAILING。
- 新值为空不覆盖（宁可留旧不写空）；解析出非 BERTH 类型跳过。
- 写前强制备份；全程串行访问 Outlook。
- 脚本本身不执行任何 git 操作，推送走正常流程。

## 万一出问题
备份在 `rob_data/tpl_backup/berth_reparse_<时间戳>/rob_history.csv`，直接拷回 `rob_data/rob_history.csv` 即可回滚。
