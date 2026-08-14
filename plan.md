# 实现计划：总任务标识（批次化）+ 断点恢复

> 目标文件 `content.js` · 基于已签核的 `design.md` · 单文件串行改动，不适用并行 subagent

## 任务拆解

### Task 1 — STORE 层批次化 + 批次工具
**位置**：`content.js` 66-90（STORE 对象 + K 常量）

- K 常量新增：`fileName: 'abw_file_name'`、`totalCount: 'abw_total_count'`
- 新增 `BATCH_BASE_KEYS`（logs/items/pending/results/xlHeaders/xlRows/batchNo/fileName/totalCount 对应的物理键）
- 新增模块级 `currentRunId`（启动时从 `abw_currentRunId` 读）+ `realKey(key)`（批次键加 `abw_run_<id>_` 前缀，全局键原样）
- STORE 的 `get/set/remove` 内部改用 `realKey(key)`
- 新增 `newRunId()`、`startNewBatch(fileName)`（清旧批次 → 切新 runId → 存 fileName）、`clearCurrentBatch()`
- **验证**：`node --check` 通过

### Task 2 — results 实时落盘 + 进度条分母
**位置**：`runAll`（825 附近）

- `results.push(result)` 后立即 `STORE.set(K.results, results)`
- 进度条改用全批次分母：`const total = STORE.get(K.totalCount, items.length)`，`updateProgress(results.length, total)`
- **验证**：`node --check`

### Task 3 — 上传文件生成新批次
**位置**：`handleFileUpload`（1380 附近）

- 文件选中即 `startNewBatch(file.name)`（解析前，保证「正在解析」等日志写新批次）
- 解析成功后 `STORE.set(K.totalCount, items.length)`
- 删除显式 `STORE.remove(K.logs)`（批次切换已自然清空旧日志）
- **验证**：`node --check`

### Task 4 — 面板批次恢复 + 批次显示
**位置**：`createPanel`（916 附近，恢复逻辑在 1104-1121）

- 恢复优先级：
  1. `pending` 非空 → 跳页中断 → 自动续跑（现有逻辑，results 一并恢复）
  2. `results` 非空且 `items` 非空 → 同页中断 → 恢复 results，算剩余行（`items.filter(it => !doneRows.has(it._row))`），renderTaskList 显示剩余，恢复进度条，**不自动跑**
  3. 都空 → 全新
- 面板 header 加批次显示元素，注入时若有 `currentRunId` + `fileName` 则显示 `文件名 · 已完成 N/总 M`
- **验证**：`node --check`

### Task 5 — 开始按钮区分「继续 / 全新」
**位置**：开始按钮事件（1077 附近）

- 有已完成 `results` → 继续模式：`toRun = items 中未完成行`，**不清空 results**；`toRun` 空则提示已全部完成
- 无 results → 全新模式：`results.length = 0`，跑全量
- **验证**：`node --check`

### Task 6 — 自测 + 版本号 + 日志
- node 模拟：`realKey` 前缀、批次切换清理、剩余行计算、继续/全新分支
- 版本号 v1.0.9 → v1.10.0
- 写工作日志、交付

## 验证方式（无测试框架的替代）

浏览器 content script 无测试框架，采用 **node 模拟纯函数**（realKey / 批次切换 / 剩余行计算 / 分支判定）+ `node --check` 语法校验，与 v1.0.9 做法一致。
