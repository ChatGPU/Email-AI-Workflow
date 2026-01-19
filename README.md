## Email-AI-Workflow（Outlook → Gmail → Apps Script → Gemini → Calendar/Tasks）

### 你要实现什么

把 PolyU（Outlook）邮件**转发到 Gmail**，由 Google Apps Script 触发 Gemini 分析邮件内容，自动：

- **分类 + 打 Gmail 标签**
- **从一封邮件中识别多个安排**（多个日期/多场 seminar/多个 deadline）
- **创建/更新多个 Google Calendar 事件**
- **创建/更新多个 Google Tasks 待办**
- **写入 Log（作为 AI 自己的备忘录）**，并在每次运行时只取**往前两个月**作为“记忆”输入，用于**避免重复创建**与**酌情更新已有事项**
- **每日总结邮件（可选）**：把当天执行过的动作与备忘录压缩成一封日报

---

### 推荐版本：V4

本仓库保留了 V1/V2/V3 的历史文件。**V4** 为当前推荐版本，文件：`V4 SmartEmailProcessor.js`。

> 注意：如果你把 **V1/V2/V3** 也一起复制进同一个 Apps Script 项目，会出现大量 **同名函数/同名变量冲突**（例如 `processEmails`）。推广/部署时建议只保留 V4 文件，或确保所有入口函数都已重命名。

V4 的核心改动：

- **配置更友好**：用户画像/风格预设/prompt 行为开关都放在文件开头（“SECTION A/B”）
- **多事项输出**：一次运行可对同一封邮件生成多个 Calendar/Task
- **Log 变成 AI 记忆**：每次仅带入最近两个月 Log，避免上下文过长
- **去重/更新**：重复提醒邮件会倾向于 SKIP 或 UPDATE，而不是重复 CREATE
- **输出风格自动切换**：TERSE / BALANCED / WARM 由 AI 根据邮件类型自动选择

---

### 一次性设置（必做）

#### 1) 准备 Gmail 标签与转发规则

- **Outlook → Gmail**：把 PolyU 学校邮箱转发到你的 Gmail
- 在 Gmail 建一个标签：`PolyU`（或你想要的名字）
- 用 Gmail 过滤器把这些转发邮件自动打上该标签（对应 `V4_CONFIG.GMAIL.SOURCE_LABEL`）

#### 2) Apps Script 启用 Tasks API（可选但推荐）

- Apps Script → **Services** → 启用 **Tasks API**
- 本仓库的 `appsscript.json` 已声明 Tasks 服务，但你仍需在脚本项目里手动启用

#### 3) 配置 Gemini API Key

在 Apps Script 里运行一次：

- `V4_setGeminiApiKey_("YOUR_API_KEY")`

（推荐写入 Script Properties，避免把 Key 写进代码）

#### 4) 配置你的“用户画像 + 偏好”

打开 `V4 SmartEmailProcessor.js`，只改文件顶部：

- **SECTION A**：`V4_PROFILE`（学院/年级/兴趣/研究主题/日程习惯等）
- **SECTION C**：`V4_CONFIG.DAILY_REPORT.RECIPIENT_EMAIL`（你的 Gmail）
- 可选：**SECTION B** 风格预设、**SECTION B2** prompt 行为开关

#### 5) 初始化（创建标签/日志表/触发器）

运行一次：

- `V4_setupSmartEmailProcessor_()`

这会：

- 创建 Gmail 标签体系（`AI/...`）
- 创建日志表：`SmartEmailProcessor Log V4`
- 创建触发器：每 5 分钟运行 `V4_processEmails_()`，每天 22 点运行 `V4_sendDailyReport_()`（可在配置里关闭）

---

### 日常使用

- 正常情况下你不需要手动运行，触发器会自动跑
- 想立刻测试一次：运行 `V4_testProcessOne_()`

---

### 去重/更新逻辑（V4 的关键）

- 每个 Calendar/Task 都会生成一个 **fingerprint → memoryId**（稳定哈希）
- Log 会记录 `memoryId`、`calendarEventId/taskId`、以及 AI 的 `assistantMemo/itemMemo`
- 下一次遇到“相同提醒邮件”，AI 会结合两个月 Log 决策：
  - **SKIP**：不重复建
  - **UPDATE**：已有事项需要补充/修改（时间/地点/链接/描述）

---

### 排错/维护

- **没有启用 Tasks API**：V4 会跳过 Tasks 写入（Log 会记录 `SKIP_TASKS_NOT_ENABLED`）
- **处理失败**：线程会被打上 `AI/❌ 处理失败` 与 `AI/⚠️ 待复核`
- **重置幂等状态（谨慎）**：运行 `V4_resetProcessorState_()`，会让旧线程可能被重新处理
