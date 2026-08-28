# culadmin 端修复：每日更新必须先 `git pull`（方案A）

> 适用机器：culadmin。本文件随仓库分发。
> 目标：彻底解决"每天自动更新用旧 `gen_html.py` 覆盖别人推上去的修复"这个反复复发的问题。

---

## 一、问题根因（先看懂再做）

culadmin 每天跑的 `auto_update.py` 有两个致命缺陷：

1. **它只提交产物，不更新代码**。看历史提交：
   ```
   deba575  daily update ... (auto_update.py schedule)
            只改了 cul_daily_movement.html + maint_snapshot.json
   a5492e9  daily update ... (auto_update.py maint)
            同样只改这两个文件
   ```
   `gen_html.py` 从不在提交范围内 → 本地代码长期停留在旧版本。

2. **它在生成 HTML 之前不同步代码**。`gen_html.py` 是生成 `cul_daily_movement.html` 的**源程序**。
   用旧版源程序重新生成 = 把 Agent/OP 列、Berth Rate 新口径、Per-Port Rate 等修复
   **全部抹掉**，再推送上线。这就是功能"第二天又没了"的真正原因。

3. 附带风险：`daily_update.bat`（另一条流水线）用的是 `git pull --ff-only`，
   一旦本地有任何本地提交就会失败并 `exit /b 1` 中断，同样跑不到生成步骤。

**结论：必须在"生成 HTML"这一步之前，强制把 `gen_html.py` 同步到 origin/main 最新版。**

---

## 二、修复方案（推荐：用封装脚本）

仓库里已新增 **`auto_update_sync.py`**。它不改变你原有的业务逻辑，只是在
`auto_update.py` **之前**插入"代码同步 + 校验"两步：

```
fetch origin
  ↓
检出/还原所有被版本管理的 .py/.bat（丢弃本地手改，以 GitHub 为准）
  ↓
git reset --hard origin/main      ← 强制对齐到远端最新
  ↓
校验 gen_html.py 含关键修复（AGENT_BY_PORT / berthedCalls），缺失就中止
  ↓
python auto_update.py             ← 你原来的逻辑，完全不变
  ↓
git add 产物 → rebase → push
```

### 部署步骤（一次性，约 2 分钟）

1. 打开命令行，进入 Claw-Report 仓库目录（就是 `auto_update.py` 所在的那个文件夹）：
   ```bat
   cd /d D:\path\to\Claw-Report
   ```

2. 把最新代码和本脚本拉下来：
   ```bat
   git fetch origin
   git reset --hard origin/main
   ```
   > 如果这里报 "local changes would be overwritten"，
   > 说明本地改过仓库里的脚本。先备份你改过的文件到仓库**外面**，再执行。

3. **先做一次自检**（不生成、不推送，只验证同步和修复是否到位）：
   ```bat
   python auto_update_sync.py --check
   ```
   期望输出：
   ```
   [sync] now at 7fc1869...
   [sync]   OK  Agent/OP column
   [sync]   OK  Berth Rate (threshold)
   [sync] check mode: sync + verify passed, skipping generation
   ```
   看到两个 `OK` 就说明代码已经是最新版。

4. **把每日任务改成跑这个脚本**：
   - 原来定时任务/计划任务里如果是 `python auto_update.py` → 改成 `python auto_update_sync.py`
   - 如果跑的是 `daily_update.bat` → 也改成 `python auto_update_sync.py`
   - 命令行手动跑同样是 `python auto_update_sync.py`

---

## 三、如果你不想用新脚本（最小改动方案）

在 `auto_update.py` 里，**在它读取/生成 HTML 的最开始**加这三行：

```python
import subprocess, sys
subprocess.run(["git", "fetch", "origin"], check=True)
subprocess.run(["git", "reset", "--hard", "origin/main"], check=True)
```

⚠️ 注意：
- 必须加在**生成 HTML 之前**，加在最后没用（那时已经用旧代码生成完了）。
- `git reset --hard` 会丢弃本地对仓库内脚本的手改。**不要手动改仓库里的 .py/.bat**，
  要改就在 GitHub 网页上改或在自己机器上改完推上去。
- 加了之后同样建议跑一次 `python auto_update_sync.py --check` 验证。

---

## 四、验证修复是否生效

跑完一次每日更新后，在 GitHub 上看最新的 commit：

- commit message 应含 `auto_update_sync`
- 该 commit **只改** `cul_daily_movement.html` / `maint_snapshot.json`（这是正常的，代码本来就是最新的）
- 打开 GitHub Pages 看板，确认这些功能都在：
  - Maintenance 未维护明细里有 **Agent / OP** 列
  - Port Wait 的 **Berth Rate 不是恒 100%**（如 CNSHA 约 5%、HKHKG 100%）
  - Monthly Trend 每根柱子有 `X/Y berthed` 且比例在 55%~75% 区间
  - Maintenance 里有 **Per-Port Rate** 表

只要这几项第二天还在，就说明修复生效了。

---

## 五、注意事项

1. **不要再手动改仓库里的 `.py` / `.bat`**。改了要么被 `reset --hard` 丢掉，
   要么会让同步失败。要改就在 GitHub 网页改，或改完立刻 push 到 main。
2. **数据/产物文件不受影响**：`cul_daily_movement.html`、`maint_snapshot.json`、
   `vessel.csv`、`.cache/` 不会被代码同步覆盖，它们由脚本自己重新生成。
3. **脚本失败会返回非 0 退出码**，不会静默上线一个用旧代码生成的坏版本。
   如果每天的任务有日志，留意 `[sync] FAILED:` 开头的行。
4. 若 SFTP 相关报错（连不上 10.5.4.2），那不是本修复的问题，
   参考 `MAINT_AUTOUPDATE.md`：先连 VPN，再跑 `python check_maint_env.py`。
