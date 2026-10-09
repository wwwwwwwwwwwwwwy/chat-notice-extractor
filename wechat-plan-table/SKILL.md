---
name: wechat-plan-table
description: >-
  读取本机微信聊天记录，把里面的待办、约定、承诺、时间点抽成可执行的计划表，
  输出多 sheet 的 Excel (.xlsx)。Use when the user asks to read WeChat chats and
  turn them into a table/plan/schedule/Excel/to-do list, or asks 从聊天记录里整理
  待办、约定、承诺、排期、计划表、做表格。只读本机、不外发聊天内容。
---

# 微信聊天 → 计划表

把微信聊天里的「谁答应过什么、什么时候、做没做」抽出来，做成一张能直接拿来排事的 Excel 表。

**适用范围**：本机 Windows 微信 4.x 的本地数据库。全程离线，聊天正文不离开这台机器。

---

## 环境（本机已实测，照此执行）

| 项 | 值 |
|---|---|
| 读取工具 | `weflow-cli` 1.9.0，已 `npm i -g` 装好且已 `init` |
| 命令位置 | `C:\Users\21419\AppData\Roaming\npm\weflow-cli.CMD`（PATH 上直接叫 `weflow-cli`） |
| 消息库 | `D:\微信\xwechat_files\<你的微信账号目录>\db_storage\message\message_0.db` |
| Python | `C:\Users\21419\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe`（3.12.14，openpyxl 3.1.5） |
| 本技能目录 | `C:\Users\21419\.dsh\skills\wechat-plan-table` |
| 默认输出根 | `C:\Users\21419\.dsh\wechat-plan\<会话短名>\` |

**不要用 PATH 上的 `python`**（那是 3.15.0a8，且缺依赖）。跑脚本一律用上表的绝对路径。

---

## 三个本机实测的坑（先看，省得白跑）

### ⛔ 坑 1：`weflow-cli export` 是坏路径，别用

本机实测（2026-10-09，weflow-cli 1.9.0）：`export` 无论 `json` 还是 `excel`、带不带
`--contract` / `--json` / `--limit`，都会**空转 7 分钟以上、CPU 打满单核、不打印任何输出、
不产出任何文件**。已在多组参数下复现。

**只走 `weflow-cli messages` 分页读取。**实测该路径 2~3 秒返回，offset 分页无重叠、
无循环、越界返回空数组，可靠。

### ⛔ 坑 2：`Documents\` 下写不进文件

`C:\Users\21419\Documents\` 及其子目录对**子进程**不可写。实测：`os.makedirs` /
`open(...,'w')` 抛 `WinError 2`；更坏的是 PowerShell 的 `New-Item` 会**假装成功**，
但紧接的 `Test-Path` 立刻是 `False`——看起来建好了，其实没有。

所以输出一律放 `C:\Users\21419\.dsh\wechat-plan\` 下面（脚本已内置目录可写性探测，
被静默拦截时会直接报错，不会假装成功）。

### ⛔ 坑 3：微信 @ 分隔符是 `U+2005`，不是空格

群里 `@某人` 后面跟的是 U+2005（FOUR-PER-EM SPACE）。抄原话时把它写成普通空格，
「逐字原话」这条铁律就悄悄失效了。**Step 4.5 的 `verify_quotes.py` 就是为了兜这个。**

---

## 执行步骤（严格按顺序）

### Step 1 · 列出会话

```powershell
weflow-cli sessions -n 30 --json
```

只向用户展示**序号、显示名、会话 ID**。不要展示最后一条消息摘要。

### Step 2 · 用户选会话

会话选择必须由用户做。一次给全选项（序号 + 名字 + ID），让用户报一个。
`messages` 命令支持用 wxid / 昵称 / 备注 / 序号任意一种指定，但**优先用会话 ID**——
昵称可能重名，`--non-interactive` 下模糊匹配会直接报错。

### Step 3 · 读取全量消息

```powershell
<PYTHON> "C:\Users\21419\.dsh\skills\wechat-plan-table\scripts\read_messages.py" `
  --talker "<会话 ID>" `
  --since <可选，YYYY-MM-DD>
```

`--out` 不传时默认落在 `C:\Users\21419\.dsh\wechat-plan\<会话短名>\`。

产出：

- `<out>/messages.json` — 结构化全量 + 抓取元信息（条数、页数、时间范围、是否检测到分页循环）
- `<out>/messages.txt` — 按时间**正序**的 `[YYYY-MM-DD HH:MM:SS] 说话人: 内容`，
  消息内部换行压成 `⏎`，保证**一行一条**

stdout 只回计数和时间范围，**不含聊天正文**。

**先看 `meta.fetched`**：为 0 说明会话选错或时间范围过滤太窄；`loop_detected=true`
说明分页出现重复，结果可能不全，必须告诉用户，不要把残缺数据当全量。

### Step 4 · 通读 `messages.txt`，抽计划项

**这一步由你（Agent）亲自做，不要交给 weflow-cli 的 `todos extract`。**
`todos extract` 需要配置 `deepseekApiKey`，本机当前未配置，跑不通。

通读全文，写出 `<out>/commitments.json`：

```json
{
  "contact": "某班委群",
  "contact_id": "12345678901@chatroom",
  "source_range": "2026-10-07 18:31:15 ~ 2026-10-09 11:06:04",
  "message_count": 82,
  "generated_at": "2026-10-09 12:35:00",
  "items": [
    {
      "id": "C001",
      "what": "把导员的早读通知转发到本班班群，并提醒大家次日有早读",
      "owner": "各班班长",
      "due": "2026-10-07",
      "due_phrase": "明天有早读",
      "status": "不明",
      "confidence": "high",
      "category": "传达",
      "evidence": [
        {"time": "2026-10-07 18:31:15", "speaker": "导员 X", "quote": "各班班长转发班群 同时提醒大家明天有早读 各班收到请回复x班收到"}
      ],
      "note": "原话指的是「各班班长」，没有指明本人是否班长。"
    }
  ]
}
```

字段约束：

| 字段 | 取值 | 说明 |
|---|---|---|
| `what` | 字符串，必填 | 一句话说清要做什么 |
| `owner` | `我` / 对方名字 / `未指明` | 这件事**谁**要动 |
| `due` | `YYYY-MM-DD` 或 `null` | 只有原话能确定到具体日期才填；模糊时间填 `null` |
| `due_phrase` | 原话里的时间说法 | 如「这周日之前」「今晚」；没有就留空 |
| `status` | `待办`/`已承诺`/`已完成`/`已过期`/`不明` | 有完成证据才写 `已完成` |
| `confidence` | `high`/`medium`/`low` | 见下方判定 |
| `category` | 自由短词 | 传达 / 核查 / 提醒 / 交付 / 会议 / 变更 / 跟进 / 其他 |
| `evidence` | 数组，可为空 | 每条含 `time`、`speaker`、`quote`；`quote` 必须是**逐字**原话 |

**置信度判定**：

- `high` — 原话里明确说了「谁 / 做什么 / 什么时候」中的至少两个，且无歧义
- `medium` — 原话只能推断出意图，时间或责任人不明确
- `low` — 你倾向认为是个约定，但原话不足以支撑；让用户自己确认

**先确定本人身份，再定 `owner`。** 「本人是谁」通常没有原文直说，但这一步直接决定
「我的待办」那张表对不对，所以：

1. **先问用户**：在群里是什么身份（班长 / 组织委员 / 生委 / 普通同学）、哪个班。
   一次问清，不要猜。
2. 用户答了之后，凡原话指向**该身份**的（例：用户是班长，则「各班班长」「各位班长」
   这类都归本人），`owner` 写 `"我"`，`note` 里写清「原话指向 X，用户已确认本人是 X」。
3. 用户没答、或原话只是泛指时，`owner` 写原话里的角色（如「各班班长」）或 `未指明`，
   **不要**把泛指直接当成本人。
4. 本人**亲自发言认领**过的（例：本人回「某班全部进群」），即使没问过用户也可以
   归本人，但要注明依据哪一句原话。
5. 身份定下来后**回头修所有项**：原来写「未指明/各班班长」的要按规则 2 重判一遍，
   别只改新项、留下旧项不一致。
6. **同岗位搭档**：同一个岗位常有两个人（例：一个班两个班长、班长 + 团支书）。
   用户点名的人写进 `commitments.json` 顶层的 `team` 数组，例如 `"team": ["某同学"]`。
   该数组里的人名下的 `owner` 会被一并算进「我的待办」，并在表里多出一列「负责人」
   区分是谁的活。**不要自己推测谁是搭档，必须用户说了才写。**

### Step 4.5 · 原话逐字校验（**必过闸门，不许跳**）

```powershell
<PYTHON> "C:\Users\21419\.dsh\skills\wechat-plan-table\scripts\verify_quotes.py" `
  --plan "<out>\commitments.json" `
  --messages "<out>\a\messages.json" "<out>\b\messages.json" "<out>\c\messages.json" --fix
```

多会话合并时把每个 `messages.json` 都传进去（支持 `*` 通配符）。脚本按
`(会话, 时间戳)` 把每条 `quote` 对回库内原文逐字比对；`evidence` 里带 `source`
时按来源会话匹配，不带时按时间戳在所有会话里找。

`--fix` 会把对得上的 quote 校正成库里的原文；**对不上的不会被改动，只会报错**。

退出码 0 = 全部逐字一致，可以往下走；退出码 1 = 有对不上的，逐条查清楚再重跑。
**不许把 `mismatched > 0` 的结果拿去生成表。**

### Step 5 · 生成 Excel

```powershell
<PYTHON> "C:\Users\21419\.dsh\skills\wechat-plan-table\scripts\build_plan_xlsx.py" `
  --plan "<out>\commitments.json" `
  --messages "<out>\messages.json" `
  --contact "<会话显示名>" `
  --out "<out>\计划表_<会话短名>.xlsx"
```

输出工作簿含四张表：

- **说明** — 会话、消息条数、时间范围、证据规则、颜色的含义
- **我的待办** — 只筛 `owner` 是本人的项；未闭环的（待办/不明/已承诺）排在前，已完成的整行压暗
- **计划表** — 全部项，按截止日期升序（无日期的排最后）；已过期标红、≤3 天标黄、无日期标灰；带筛选与冻结首行
- **证据索引** — 每条计划对应的全部原话，带时间戳

脚本会校验 `items` 完整性，缺 `what` / 缺 `evidence` / 置信度取值非法都会在 stderr 报警告。

### Step 6 · 展示结论

在对话里给一份 Markdown 摘要：**已过期 / 今天本周到期 / 我该做的** 分三组列出，
每条带一句原话。然后给出 xlsx 的绝对路径，并用 `present` 把文件交给用户。

---

## 铁律（不可协商）

1. **无证据不成表。** 每一条 `what` 都必须能指回一条带时间戳的原话。指不到就不写进表，
   不写进表比写错强。
2. **不许编。** 不确定的日期填 `null`，不确定的约定标 `low`。不许用「通常」「应该」补全。
3. **`quote` 必须逐字**，且必须过 Step 4.5 的校验闸门。不许改写、润色、拼接。
4. **消息类型有限。** 语音、图片、表情在库里通常只有 `[表情]`、`[未识别的消息类型 N]`
   这类占位。看到占位符不要脑补内容；上下文因此断链时，明说「这段是语音，读不到内容」。
5. **身份不许猜。** 见 Step 4。
6. **隐私。** 聊天正文全程本机处理，未经用户明确同意不得发往任何外部服务（含联网检索）。
   stdout 只回计数，不要把大段聊天正文倒进对话。
7. **不要用 `weflow-cli export`**（坑 1）。不要 `config clear`、不要 `forget-keys`、
   不要动白名单/黑名单——那会破坏已配好的数据库密钥。

---

## 错误处理

| 现象 | 处理 |
|---|---|
| `找不到 weflow-cli` | `npm i -g weflow-cli`，再 `weflow-cli init` |
| `weflow-cli check` 显示「未初始化」 | 让用户启动并登录微信，然后 `weflow-cli init` |
| `messages` 报匹配不唯一 | 换成会话 ID（`xxx@chatroom` / `wxid_xxx`）重跑 |
| `meta.fetched = 0` | 会话选错，或 `--since` 过滤掉了全部消息 |
| `loop_detected: true` | 分页出现重复 localId，结果可能不全；必须告知用户，不要当作全量 |
| 输出目录报「创建后不存在」/ `WinError 2` | 撞上坑 2，把 `--out` 换到 `C:\Users\21419\.dsh\wechat-plan\` 下面 |
| `verify_quotes.py` 退出码 1 | 有 quote 对不上原文。看 `details` 里的 `time_hit_elsewhere` 找正确时间戳；多半是 @ 的 U+2005 被写成了普通空格 |
| 群里谁说的话认不出来 | `senderDisplay` 为空时脚本回落到 wxid；在报告里如实标为 wxid，别猜人名 |
| 时间戳看起来是 1970 年 | `createTime` 是秒级 epoch；若是毫秒级需除以 1000（脚本按秒处理） |
| 想加新的分析维度 | 改 `build_plan_xlsx.py` 的列定义，不要另起一套流程 |
