# 聊天通知提取器 · chat-notice-extractor

从微信 / QQ 聊天记录中提取**通知**，抽取出通知时间、发生时间、截止时间、地点、面向对象、
是否需回应等字段，输出一张带**时间列**的 Excel 台账。

采用「**规则高召回 + AI 精判**」两段式流程 —— 规则层负责不错过，AI 层负责判得准、抽得全。

```
messages.json ──►[规则召回]──► candidates.json ──►[AI 精判]──► notices.json ──► 通知表.xlsx
                  高召回、可复现                       高准确、抽字段
```

## 本仓库包含两个技能

| 目录 | 做什么 | 链接方式 |
|------|--------|----------|
| 仓库根（本 README 所述） | **通知提取**：从聊天里捞出「通知」，按 12 类打标并抽时间/地点/对象字段 | 消费统一的 `messages.json`（不自带导入） |
| [`wechat-plan-table/`](wechat-plan-table/SKILL.md) | **计划表**：从聊天里抽出「待办 / 约定 / 承诺」，按截止时间排成可执行的 Excel 计划表 | 自带实测可用的本机微信导入链路（`weflow-cli` 分页读取） |

两者互补：根目录的技能强在**分类与字段抽取**，`wechat-plan-table` 强在**端到端可跑**和
**原话逐字校验**。若已有 `messages.json`，用根目录的技能；若要从本机微信直接读起，用子目录的技能。

## 为什么分两段

纯规则会漏掉变体表达（「回头再说」这类没有关键词但实为约定的话）；
纯 AI 在长会话上既贵又容易漏。规则先把几千条消息压到几十条候选，AI 只读这些候选，
准确率与成本同时最优。

规则分**只用于排序**，不作为最终结论 —— 这是设计上的硬约束。

## 特性

- **通知时间与发生时间分列** —— 通知发出时间和事项发生时间常被混淆，本工具明确区分
- **会话类型感知** —— 群聊 / 单聊分开标注，群聊通知与私聊约定不混为一谈
- **零依赖出表** —— 优先用 `openpyxl`；未安装时自动回退到内置的纯标准库 xlsx 写入器
- **可回溯** —— 附「候选明细」工作表，记录每条为什么被召回（规则分、命中关键词）
- **不碰导入轮子** —— 只消费统一的 `messages.json`，导入层的复杂度被隔离在外
- **隐私优先** —— 密钥、Token、聊天正文一律不进对话，只回联系人名与产物路径

## 安装

作为 Agent Skill 使用，把本目录放到技能目录下：

```bash
git clone https://github.com/wwwwwwwwwwwwwwy/chat-notice-extractor.git \
  ~/.workbuddy-ai/skills/chat-notice-extractor
```

也可以直接当普通 Python 脚本用（Python 3.8+，标准库即可运行）。

## 快速开始

### 1. 准备聊天数据

目录结构：

```
data/contacts/
├── 2024级计科1班群__a1b2c3d4/messages.json
└── 张三__e5f6a7b8/messages.json
```

`messages.json` 的统一契约：

```json
{
  "source": "wechat",
  "chat_type": "group",
  "contact_username": "wxid_xxx",
  "contact_display": "2024级计科1班群",
  "total": 1234,
  "messages": [
    {
      "local_id": 1,
      "sender": "me",
      "sender_name": "我",
      "content": "通知一下：本周五下午3点开班会",
      "timestamp": 1772419680,
      "type": "text"
    }
  ]
}
```

- `timestamp` 为**秒**级时间戳，毫秒/微秒会自动降级
- `sender` 必须是 `me` 或 `them`
- `chat_type` 与 `sender_name` 为推荐字段，缺失时会降级推断

数据来源（微信 / QQ 导出）见 [`references/import-pipeline.md`](references/import-pipeline.md)。

### 2. 先跑一遍合成数据（不碰真实聊天记录）

```bash
python scripts/demo_data.py --out-dir demo/data/contacts
python scripts/scan_notices.py --contacts-dir demo/data/contacts --out out/candidates.json
python scripts/build_table.py --candidates out/candidates.json --out out/通知表.xlsx --csv
```

### 3. 用真实数据

```bash
# 规则召回
python scripts/scan_notices.py --contacts-dir data/contacts --out out/candidates.json

# 由 AI 读取 candidates.json，产出 out/notices.json（口径见 references/notice-taxonomy.md）

# 出表
python scripts/build_table.py \
  --candidates out/candidates.json \
  --notices out/notices.json \
  --out out/通知表.xlsx --csv
```

## 输出字段

| 列 | 说明 |
|----|------|
| 通知时间 | **消息本身的发送时间**（核心时间列） |
| 会话 / 会话类型 | 来自哪个会话，群聊还是单聊 |
| 发送人 | 谁发的 |
| 通知类型 | 会议 / 活动 / 报名 / 缴费 / 提交材料 / 变更 / 放假调课 / 考试 / 值班排班 / 待办 / 提醒 / 其他 |
| 事项 | 通知内容凝练（≤20 字） |
| 发生时间 | 通知指向的事件时间 |
| 截止时间 | 报名 / 提交 / 缴费的截止时间 |
| 地点 | 房间号、地址、店名 |
| 面向对象 | 全体成员 / 某小组 / 我 |
| 需回应 | 需参加 / 需报名 / 需缴费 / 需提交 / 需回复 / 无需回应 |
| 置信度 | high / medium / low |
| 原文 | 支持该判定的原话摘录 |

## 常用参数

| 参数 | 默认 | 说明 |
|------|------|------|
| `--threshold` | 3 | 规则打分阈值。调低召回更全、噪声更多 |
| `--context` | 2 | 候选窗口前后各取几条上下文 |
| `--max-candidates` | 300 | 单会话候选上限 |
| `--meta` | — | 会话元信息 JSON，用于修正 `chat_type` |

## 目录结构

```
├── SKILL.md                      # Agent Skill 主入口（5 步流程）
├── references/
│   ├── notice-taxonomy.md        # 通知 12 类分类 + 字段抽取口径 + 去重规则
│   └── import-pipeline.md        # 微信 / QQ 导入链路 + 统一契约 + 隐私红线
├── scripts/
│   ├── notice_rules.py           # 规则召回引擎
│   ├── scan_notices.py           # 遍历会话 → candidates.json
│   ├── build_table.py            # 候选(+精判) → xlsx
│   ├── xlsx_min.py               # 纯标准库 xlsx 写入器（无依赖回退）
│   └── demo_data.py              # 合成测试数据
├── tests/
│   └── test_pipeline.py          # 36 项单元 / 集成测试
└── wechat-plan-table/            # 第二个技能：聊天记录 → 待办/约定计划表
    ├── SKILL.md
    └── scripts/
        ├── read_messages.py      # weflow-cli 分页读取本机微信 → messages.json/txt
        ├── verify_quotes.py      # 原话逐字校验闸门（多会话）
        └── build_plan_xlsx.py    # 计划表 → 多 sheet xlsx
```

## 测试

```bash
python -m unittest discover -s tests -v
```

覆盖：时间戳归一化（秒/毫秒/微秒/时区）、噪声过滤、评分门控、会话类型推断、
候选窗口与锚点定位、AI/规则两种出表模式、xlsx 写入与非法字符转义、端到端链路。

## 设计要点

**双条件门控**：仅按分数召回会把「同学们早上好」（命中泛称呼 + 时段）、
「晚上一起吃饭吗」（命中时段）误判为通知。因此判定要求「分数达标 **且** 有实质信号」——
非泛称呼关键词、具体时间标签（`X点`/`星期`/`期限` 等）、或结构化标记（群公告 / @全体成员）。

**锚点定位**：候选窗口会向前后扩展以提供上下文，因此必须区分「上下文」与「通知主体」。
出表时按 `anchor_local_ids` 定位主体，否则会误取到邻近的另一条通知。

## 隐私

- 密钥、Token 一律不写入对话与文件
- 聊天正文不回显到对话，只展示联系人名、消息条数、产物路径
- 原始导出与转换结果应加入 `.gitignore`

## License

MIT
