# 聊天记录导入链路

本文件说明如何把微信 / QQ 的聊天记录变成本技能可消费的 `messages.json`。

**职责边界**：本技能只消费 `messages.json`，不重复造导入轮子。下列链路改编自开源项目
`she-love-me`（`github.com/863401402/she-love-me`）的导入层，原理见文末。

---

## 一、统一数据契约

所有来源最终都必须产出这个结构，放在 `<会话目录>/messages.json`：

```json
{
  "source": "wechat | qq | ciphertalk | weflow | markdown",
  "chat_type": "group | private",
  "contact_username": "wxid_xxx 或 QQ号",
  "contact_display": "2024级计科1班群",
  "own_wxid": "自己的标识",
  "total": 1234,
  "messages": [
    {
      "local_id": 1,
      "sender": "me | them",
      "sender_name": "班长",
      "content": "通知一下：本周五下午3点开班会",
      "timestamp": 1772419680,
      "type": "text",
      "transcript": "语音转写（可选）"
    }
  ]
}
```

关键约定：

- `timestamp` 为**秒**级 Unix 时间戳。毫秒/微秒会被自动降级，无需手工转换。
- `sender` 必须是 `me` 或 `them`。群聊中 `them` 代表「非本人」，具体是谁看 `sender_name`。
- `sender_name` 为**可选增强字段**。缺失时表格的「发送人」列会退化为「我 / 会话名」。
  若导入链路能提供发送人显示名，务必带上 —— 通知往往需要追溯「谁发的」。
- `chat_type` 建议显式声明。缺失时本技能按内容启发式推断，可能出错。

> ⚠️ `she-love-me` 的原生契约**不含** `chat_type` 和 `sender_name`。若直接复用它产出的
> `messages.json`，请用 `scan_notices.py --meta meta.json` 补 `chat_type`，
> 并接受「发送人」列精度下降。

---

## 二、微信路径

微信本地数据是**加密的 SQLite**（WCDB 变体），不能直接读。链路必须两步：
**先从运行中的 `Weixin.exe` 进程内存取 64 位密钥 → 再解密 → 再按表结构提取**。

### 前置条件（必须由用户完成，Agent 无法代劳）

1. Windows 10/11，微信 **处于运行且已登录**状态
2. Node.js 18+（导出器是 npm 包）
3. **管理员权限终端**（读进程内存需要）
4. Python 3.9+

### 2A. 首选：weflow-cli

```powershell
npm install -g weflow-cli
weflow-cli init                          # 初始化，需要微信已登录
weflow-cli sessions -n 30                # 列出最近 30 个会话
weflow-cli export "<会话ID>" json --output "data/raw"
```

导出文件通常为 `data/raw/<会话ID>_messages.json`。
**群聊也要逐个会话导出** —— 想覆盖「全部会话」，需要对每个会话各导一次。

#### ⚠️ weflow-cli 实测踩坑（微信 4.x / 多进程，2026-10 实测）

以下均为实测结论，不是推测。复现时按此排查，可省大量时间。

**1. 配置目录是 `~/.weflow-cli/`，不是 `~/.weflow/`**

```bash
ls -la ~/.weflow-cli/config.json      # 正确
```

`config.json` 中存有 `decryptKey` / `ntKey` / `contactKey` / `snsKey` 等，
均为 `lock:` 前缀 —— 由 **DPAPI 保护**，外部无法取出明文密钥做独立校验。

**2. `dbkey` 在非交互环境会静默卡死**

它的捕获流程是交互式的，会停在确认提示上：

```
? 确认启动需要人工配合的数据库密钥捕获流程吗？ (y/N)
```

非交互环境下会一直等待直到超时，**输出里不会有任何错误提示**。
必须显式加 `--yes`：

```bash
weflow-cli dbkey --force --yes --timeout 180000
```

可用 `--dry-run --json` 先预览，它会明确告知是否需要读进程内存：

```json
{"action":"dbkey.capture","interactiveRequired":true,"scansProcessMemory":true}
```

**3. 它等的是「数据库初始化事件」，不是普通写库**

失败时的报错：

```
✗ 失败: 等待超时，未观察到数据库初始化事件
```

关键点：**微信 WAL 每秒都在更新，工具依然超时**。实测监测到
`message_0.db`、`contact.db`、`session.db` 及各自 `-wal` 持续被写入
（`-incremental.material` 也在刷新），但捕获仍报超时。

→ 它等待的是**微信启动时首次打开数据库**这一特定事件。
**光是让微信保持运行、持续收消息是不够的，必须完全退出微信再重新登录。**

**4. 需要管理员权限（高概率）**

`scansProcessMemory: true` 意味着要读另一个进程的内存。在非提权会话
（`IsAdmin=False`）下这通常被拒绝。推荐组合：

**管理员终端 + 重启微信**

```powershell
# 以管理员身份打开终端后
weflow-cli init
# 然后完全退出微信，重新登录，触发数据库初始化事件
```

**5. 如何判断密钥是否真的拿到了**

不要只看 `init` 报告成功。用这两条命令验证：

```bash
weflow-cli sessions -n 5 --json
weflow-cli contacts --json
```

若返回 `{"success":true,"sessions":[]}` / `{"success":true,"contacts":[]}` ——
**命令成功但数据为空，说明密钥是未完成的占位值**，不是「微信里没有会话」。

交叉验证手段：比对配置里的 salt 与库文件实际 salt。

```python
import json, os
cfg = json.load(open(os.path.expanduser("~/.weflow-cli/config.json"), encoding="utf-8"))
db = r"D:\微信\xwechat_files\<wxid>\db_storage\message\message_0.db"
print("库文件 salt:", open(db, "rb").read(16).hex())
print("配置 ntSalt :", cfg.get("ntSalt", ""))
```

两者**一致**只说明工具正确识别了目标库，**并不能**证明密钥有效。

**6. 微信 4.x 的库路径与 3.x 不同**

```
4.x:  db_storage/message/message_0.db          # 单文件
3.x:  message/message_N.db                     # 多文件，每联系人一张 Msg_{md5(wxid)} 表
```

4.x 走 WCDB/NT 连接（`ntDbPath` / `ntKey`），而 3.x 走 `dbPath3x` / `decryptKey3x`。
`config show --json` 会给出 `dataVersion` 字段，据此判断走哪条路径。

**7. WAL 可能远大于主库**

实测：`session.db` 主库仅 152 KB，但 `session.db-wal` 达 **4.1 MB**。
数据主要滞留在 WAL 中，未合并回主库。若自行解析库文件，
**必须一并处理 `-wal`**，否则会读到几乎空的库。

**8. 多进程微信**

实测同时存在 5 个 `Weixin.exe`（主窗口 1 个 + 辅助 4 个）外加多个 `WeChatAppEx.exe`。
这正是 CipherTalk「只 Hook 第一个进程」成为已知缺陷的原因，
也是 weflow-cli 需要「扫描进程内存」而非「附加单进程」的原因。

### 2B. 回退：CipherTalk CLI

当 weflow-cli 安装或初始化失败时使用。CipherTalk 通过官方
`wechat_key_tool.dll` + `wxKeyService.ts`，经 Node/koffi FFI 调用，
先做 `wkt_challenge` 签名挑战，再多进程扫描取密钥。

```powershell
npm install -g ciphertalk-cli
miyu --format=json --quiet key get --save     # 取密钥，密钥不落 stdout
miyu --format=json --quiet status             # 确认 configured=true
miyu --format=json --quiet -- --limit=30 sessions
miyu --format=json --quiet -- export "<会话ID>" --output "data/raw/ciphertalk-chat.json"
```

⚠️ CipherTalk CLI 只 Hook `tasklist` 返回的**第一个** `Weixin.exe`，
多进程微信下可能选错进程。若返回「等待密钥超时」，不要反复重登。

### 2C. 兜底：CipherTalk 桌面版 + MCP

CLI 无法验证数据库但桌面版能完成账号配置时使用。桌面版配置好后由其自带 MCP
完成列会话与完整导出。**不要**从桌面版单独抽取 `WCDB.dll` —— 官方二进制脱离应用
环境会拒绝初始化。

### 2D. 解密后的库结构（转换器需要知道的事）

**微信 4.x（WCDB / NT 连接）** —— 消息库为单文件：

```
<账号目录>/db_storage/
├── contact/contact.db          # contact 表：username / nick_name / remark
├── session/session.db          # 会话列表
└── message/message_0.db        # 消息主体（单文件）
```

⚠️ 同目录下每个库都附带 `-wal` / `-shm` / `-first.material` / `-incremental.material` /
`-last.material`。**`-wal` 可能远大于主库**（实测 `session.db` 152 KB 而
`session.db-wal` 4.1 MB），自行解析时必须一并处理，否则读到几乎空的库。

**微信 3.x（旧版）** —— 消息库按联系人分文件：

```
<账号目录>/db_storage/
├── contact/contact.db
└── message/message_N.db        # 每个联系人一张 Msg_{md5(username)} 表
```

`Msg_*` 表列（3.x）：

| 列 | 含义 |
|----|------|
| `local_id` | 会话内序号 |
| `local_type` | 1文本 3图片 34语音 43视频 47表情 49链接/文件 50通话 10000系统 10002撤回 |
| `create_time` | 秒级时间戳 |
| `real_sender_id` | 需经 `Name2Id` 表反查 `username`，才能判断 me/them |
| `message_content` | 正文；`WCDB_CT_message_content == 4` 时为 **zstd 压缩**，需解压 |

撤回消息的 `real_sender_id` 通常为 0，需回溯上一条的发送方来补。

---

## 三、QQ 路径（最轻）

QQ **不碰本地数据库**，走 NapCat（QQ 协议端）+ QQ Chat Exporter 暴露的本地 REST API。

### 前置

1. 下载 [NapCat-QCE Releases](https://github.com/shuakami/qq-chat-exporter/releases) 的
   `NapCat-QCE-Windows-x64-*.zip`
2. 解压后双击 `launcher-user.bat`，用手机 QQ 扫码登录
3. 控制台出现 `Token: xxxxx`，复制该 Token

### 调用流程（默认端口 40653）

```
GET  /api/system/info            → selfInfo.uid，用于判定 me/them
GET  /api/friends?limit=9999     → 好友列表（精确匹配 QQ号/uid 优先，备注/昵称模糊兜底）
POST /api/messages/export        → {"peer":{"chatType":1,"peerUid":"..."},"format":"JSON"} → taskId
GET  /api/tasks/{taskId}         → 轮询直到 status=completed → filePath
```

`chatType`：`1` = 私聊，`2` = 群聊。群聊要覆盖，需显式传 `2`。

请求头带 `Authorization: Bearer <token>`。轮询建议间隔 3 秒，超时上限 600 秒
（大量消息可能数分钟）。

---

## 四、隐私红线（务必遵守）

1. **密钥、Token、Access Token 一律不写入对话**，也不写入任何 tracked 文件。
2. **聊天正文不回显到对话**。只展示：联系人显示名、会话 ID、消息条数、产物路径。
3. 原始导出放 `data/raw/`，转换结果放 `data/contacts/`，两者都加入 `.gitignore`。
4. 每个会话独立目录，命名 `safe_slug(显示名)__md5(标识)[:8]`，避免互相覆盖。

---

## 五、导入原理速查

| 层 | 微信 | QQ |
|----|------|----|
| 采集 | 从 `Weixin.exe` 进程内存取密钥 → 解密 SQLite | NapCat 协议端暴露 REST API |
| 鉴权 | 管理员终端 + 进程内存读取 | Bearer Token |
| 取数 | 按 `Msg_{md5(wxid)}` 表 + `Name2Id` 反查 | `/api/messages/export` 异步任务 + 轮询 |
| 难点 | 密钥提取（多进程、Hook 选错进程） | 无（链路最轻） |
| 门槛 | 高：微信登录 + 管理员权限 + Node.js | 低：扫码 + 复制 Token |

两者最终都收敛到同一份 `messages.json`，这正是本技能只依赖该契约的原因 ——
**导入层的复杂度被隔离在技能之外**。
