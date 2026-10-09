---
name: chat-notice-extractor
description: >-
  从微信 / QQ 聊天记录中提取「通知」，抽取出通知时间、发生时间、截止时间、地点、
  面向对象、是否需回应等字段，输出带时间列的 Excel 表格。采用「规则高召回 + AI 精判」
  两段式流程。当用户要求「从聊天记录里找通知 / 公告」「整理群里的通知」「把聊天里的
  安排、报名、截止时间做成表格」时使用。
---

# 聊天通知提取器

把散落在聊天记录里的**通知**（会议、报名、缴费、变更、提交材料、待办约定等）捞出来，
整理成一张带**通知时间**列的 Excel 台账。

## 核心设计：两段式

```
messages.json ──► [规则召回] ──► candidates.json ──► [AI 精判] ──► notices.json ──► 通知表.xlsx
                  高召回、可复现                       高准确、抽字段
```

**为什么分两段**：纯规则会漏变体表达，纯 AI 在长会话上又贵且会漏。规则先把候选压到
几十条，AI 只读这些候选，准确率和成本同时最优。规则分只用于**排序**，不作为最终结论。

> ⚠️ **铁律**：规则层只负责「不错过」，**不得**把规则命中直接当成最终通知交付。
> 除非用户明确要求「只要规则结果快速看看」，否则必须走完 AI 精判。

## 工作目录约定

- 临时产物放 `out/`，聊天数据放 `data/`，两者都不应提交到版本库。
- 下文 `<PYTHON>` 指当前环境可用的 Python 解释器（Python 3.8+，标准库即可运行；
  若已装 `openpyxl` 则表格格式更完整）。

---

## Step 0：确认范围

向用户确认三件事，缺一不可：

1. **数据来源**：微信本机导出 / QQ / 已有 `messages.json`？
2. **会话范围**：全部会话（群 + 单聊）/ 仅群聊 / 仅单聊？
3. **时间范围**：全量，还是只看最近 N 天？

若用户已在本轮对话中说明，直接采用，不要重复追问。

---

## Step 1：取得 messages.json

本技能**不自己实现聊天导入**，而是复用成熟的导入链路，只消费统一的 `messages.json`。

### 1A. 已有统一格式文件（最快）

目录结构：

```
data/contacts/
├── 2024级计科1班群__a1b2c3d4/messages.json
└── 张三__e5f6a7b8/messages.json
```

### 1B. 微信本机导出

读取并执行 `references/import-pipeline.md`。要点：微信本地库是加密 SQLite，
需要先从运行中的 `Weixin.exe` 取 64 位密钥，再解密、再按 `Msg_{md5(wxid)}` 表提取。
链路需要**微信登录 + 管理员终端**，这两步必须由用户完成。

### 1C. QQ 导出

QQ 走 NapCat + QQ Chat Exporter 的本地 REST API（默认 `127.0.0.1:40653`），
只需一个 Access Token，不需要管理员权限。步骤见 `references/import-pipeline.md`。

### 1D. 只想先跑通流程

用合成数据自测，不碰真实聊天记录：

```bash
<PYTHON> scripts/demo_data.py --out-dir demo/data/contacts
```

### 会话类型标注（重要）

`messages.json` 顶层可带 `chat_type` 字段（`group` / `private`）。
若导入链路没有提供，本技能会按内容启发式推断；**推断可能出错**，
建议用 `--meta` 显式指定：

```json
{ "2024级计科1班群__a1b2c3d4": { "chat_type": "group", "display_name": "计科1班群" } }
```

---

## Step 2：规则召回

```bash
<PYTHON> scripts/scan_notices.py \
  --contacts-dir data/contacts \
  --out out/candidates.json \
  --threshold 3 --context 2
```

- `--threshold`：默认 3。**调低 → 召回更全、噪声更多**；调高 → 更干净、可能漏。
- `--context`：候选窗口前后各取几条上下文，默认 2。通知常跨多条消息，不建议设为 0。
- 多会话时用 `--meta meta.json` 修正会话类型。
- 单会话时用 `--input <messages.json>`。

读完输出，向用户汇报：**扫了几个会话、多少条消息、召回多少候选**。
若候选数为 0，说明阈值过高或数据本身无通知，先降低阈值重试再判断。

---

## Step 3：AI 精判（核心）

读取 `out/candidates.json`，对**每一条候选**做判定与字段抽取，产出 `out/notices.json`。

判定与抽取口径详见 `references/notice-taxonomy.md`。三条执行铁律：

1. **看窗口不看单条** —— `window` 是上下文，`anchor_local_ids` 才是通知主体。若上下文里
   出现了更正/取消（如"改到晚上七点"），以**最后一条**为准。
2. **无证据不填** —— 时间、地点、截止等字段在原文里找不到就留空字符串 `""`，
   **不要**根据常识推测补全。宁可留空，不可编造。
3. **去重与合并** —— 同一件事被多次提及（先通知、后补充、再改期）时，**只保留最终态**，
   并在 `title` 中体现最终结果。

`notices.json` 结构：

```json
{
  "notices": [
    {
      "source_candidate": "c0002",
      "notice_time": "2026-03-02 09:28:00",
      "contact_display": "2024级计科1班群",
      "chat_type": "group",
      "sender": "班长",
      "notice_type": "会议",
      "title": "周五班会，确认毕业设计选题",
      "event_time": "2026-03-06 15:00",
      "deadline": "",
      "location": "实验楼B301",
      "audience": "全体成员",
      "action_required": "需参加",
      "confidence": "high",
      "evidence": "通知一下：本周五（3月6日）下午3点在实验楼B301开班会……"
    }
  ]
}
```

- `notice_time` **必须**填（消息本身的发送时间），这是表格的「通知时间」列。
- `confidence` 取 `high` / `medium` / `low`。
- 判为**不是通知**的候选直接不写入 `notices` 数组即可（等于丢弃）。
- 若整批候选都不是通知，输出 `{"notices": []}` 并如实告知用户，不要硬凑。

---

## Step 4：出表

```bash
<PYTHON> scripts/build_table.py \
  --candidates out/candidates.json \
  --notices out/notices.json \
  --out out/通知表.xlsx \
  --csv
```

产出两个工作表：

- **通知表**：14 列 —— 序号 / 通知时间 / 会话 / 会话类型 / 发送人 / 通知类型 / 事项 /
  发生时间 / 截止时间 / 地点 / 面向对象 / 需回应 / 置信度 / 原文
- **候选明细**：候选 ID、规则分、命中类别、命中关键词，便于回溯「为什么召回它」

若用户只想快速预览、不走 AI：

```bash
<PYTHON> scripts/build_table.py --candidates out/candidates.json --out out/通知表_规则版.xlsx
```

此时「通知类型」为规则类别或 `待判定`，「发生时间」只回填时间语义标签（如 `月日`）。

> 表格引擎优先用 `openpyxl`（格式更完整）；未安装时自动回退到内置的纯标准库写入器，
> 无需额外安装依赖。

---

## Step 5：交付

用 `present_files` 打开生成的 `.xlsx`，并在回复中给出：

- 扫描范围（会话数 / 消息数）
- 召回候选数 → 最终确认通知数
- 按通知类型、按会话的简要分布
- 置信度 `low` 或字段大面积留空的条目，单独列出让用户人工核对

---

## 错误处理

| 现象 | 处理 |
|------|------|
| `data/contacts` 下找不到 `*/messages.json` | 先完成 Step 1 导入；确认目录层级不是多套了一层 |
| 候选数为 0 | 降低 `--threshold`（试 2）后重跑；仍为 0 则数据可能确实无通知 |
| 候选数异常多（>300） | 提高 `--threshold`，或用 `--max-candidates` 限流后分批精判 |
| 会话类型显示「未知」 | 用 `--meta` 显式指定 `chat_type` |
| 表格打不开 / 中文乱码 | CSV 已用 UTF-8-BOM；若仍乱码请用 Excel 的「数据→从文本」并选 UTF-8 |
| 时间显示不对 | `messages.json` 的时间戳应为**秒**；毫秒会被自动降级，若差 8 小时请检查时区（本技能按 UTC+8 展示） |
| 语音/图片消息被跳过 | 属正常行为：非文字消息不参与通知抽取，除非数据源提供了 `transcript` |
| `weflow-cli init` 报「等待超时，未观察到数据库初始化事件」 | 见 `references/import-pipeline.md` 的 weflow-cli 实测踩坑。要点：需**管理员终端**，且必须**完全退出微信再重新登录** —— 该工具等的是「启动时首次打开数据库」事件，普通持续写库不算 |
| `weflow-cli sessions` 返回 `{"success":true,"sessions":[]}` | **不代表微信里没有会话**，而是密钥为未完成的占位值。用 `contacts --json` 交叉验证；两者都空即为密钥问题，重做密钥捕获 |
| `weflow-cli dbkey` 无输出且长时间挂起 | 其捕获流程是交互式的，非交互环境会静默卡在 `(y/N)` 确认。必须加 `--yes` |
