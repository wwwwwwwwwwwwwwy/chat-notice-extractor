"""notice_rules.py - 通知规则召回层（纯标准库，无第三方依赖）

职责：把归一化后的聊天记录（messages.json）压缩成「可能包含通知」的候选窗口，
交由上层 AI 精判。本层只负责**高召回**，宁可多圈不可漏圈。

设计要点：
  1. 打分制而非硬过滤 —— 关键词类别分权重累加，避免单一词表遗漏变体表达。
  2. 窗口化而非单条 —— 通知常跨多条消息（先问"有人吗"再说"明天三点开会"），
     因此以命中消息为锚点，向前后各取 context 条组成窗口。
  3. 双通道 —— 规则打分 + 结构化标记（群公告、@全体成员、含链接）。
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone

# ── 时区：聊天记录按北京时间展示 ──────────────────────────────
TZ = timezone(timedelta(hours=8))

# ── 关键词库：(类别, 权重, 词表) ─────────────────────────────
KEYWORD_GROUPS = [
    ("指向性", 3, ["@全体成员", "全体成员", "所有人", "请大家", "通知大家", "各位同学", "各位同事"]),
    ("通知类", 3, ["通知", "公告", "通告", "须知", "周知", "请知悉", "特此", "提醒一下", "重要事项"]),
    ("安排类", 2, ["安排", "调整", "变更", "更改", "取消", "延期", "推迟", "提前", "改为", "敲定", "定了"]),
    ("报名类", 2, ["报名", "接龙", "统计", "填写", "问卷", "登记", "汇总", "报数", "举手", "回我一下", "回复我"]),
    ("会议类", 2, ["会议", "例会", "开会", "班会", "参会", "签到", "出席", "视频会", "线上会", "碰头", "复盘", "答辩"]),
    ("活动类", 2, ["活动", "聚餐", "团建", "出游", "聚会", "比赛", "培训", "讲座", "考试", "体检", "面试"]),
    ("事务类", 2, ["缴费", "交费", "交给", "交到", "提交", "上交", "截止", "deadline", "办理", "领取", "值班", "加班", "请假", "放假", "开学", "报到"]),
    ("泛称呼", 1, ["各位", "大家", "亲们", "同学们", "同事们"]),
]

# 泛称呼单独成词时不足以判定为通知（"同学们早上好" 不是通知）
WEAK_CATEGORY = "泛称呼"

# 仅凭「时段」不足以判定为通知（"晚上一起吃饭吗" 不是通知）
SPECIFIC_TIME_LABELS = {"绝对日期", "月日", "相对日", "星期", "时点", "期限"}

# ── 时间表达式：(正则, 语义标签) ────────────────────────────
TIME_PATTERNS = [
    (r"\d{4}\s*[-/年]\s*\d{1,2}\s*[-/月]\s*\d{1,2}\s*[日号]?", "绝对日期"),
    (r"(?<!\d)\d{1,2}\s*[月/-]\s*\d{1,2}\s*[日号]", "月日"),
    (r"(今天|今日|明天|明日|后天|大后天|今晚|明晚|当天|当日)", "相对日"),
    (r"((?:本|这|下|下个|上)\s*周\s*[一二三四五六日天末]|周[一二三四五六日天末]|星期[一二三四五六日天]|礼拜[一二三四五六日天])", "星期"),
    (r"(上午|早上|早晨|中午|下午|傍晚|晚上|夜里|凌晨)", "时段"),
    (r"(?<!\d)\d{1,2}\s*[点:：]\s*\d{0,2}\s*分?", "时点"),
    (r"(截止|最晚|务必于|之前|以前|前完成|前提交|前回复|前报名)", "期限"),
]

# ── 祈使/动作：(正则, 语义标签) ─────────────────────────────
ACTION_PATTERNS = [
    (r"(请|务必|记得|别忘|一定要|需要你|麻烦|尽快|抓紧|必须|要求)", "祈使"),
    (r"(回复|确认|接龙|报名|登记|填写|提交|签到|领取|交一下)", "动作要求"),
]

URL_RE = re.compile(r"https?://", re.I)

# ── 噪声：纯应答，不构成通知 ────────────────────────────────
NOISE_EXACT = {
    "嗯", "哦", "好", "好的", "行", "哈哈", "哈哈哈", "呵呵", "收到", "在", "在的",
    "嗯嗯", "哦哦", "谢谢", "谢谢啦", "ok", "okay", "好嘞", "好的好的", "？", "?", "。。",
}
NOISE_TYPES = {"emoji", "image", "voice", "video", "revoke", "call"}

# 群公告结构化标记
GROUP_NOTICE_MARKERS = ("群公告", "群通知", "公告栏")


def ts_to_str(ts, fmt: str = "%Y-%m-%d %H:%M:%S") -> str:
    """秒级时间戳 → 北京时间字符串。兼容毫秒。"""
    try:
        value = float(ts)
    except (TypeError, ValueError):
        return ""
    while value >= 100_000_000_000:
        value /= 1000.0
    try:
        return datetime.fromtimestamp(value, TZ).strftime(fmt)
    except (OverflowError, OSError, ValueError):
        return ""


def is_noise(content: str, msg_type: str) -> bool:
    if msg_type in NOISE_TYPES:
        return True
    text = (content or "").strip()
    if len(text) <= 1:
        return True
    return text.lower() in NOISE_EXACT


def score_message(msg: dict) -> dict:
    """对单条消息打分。返回 {'score': int, 'hits': {...}}。"""
    content = str(msg.get("content") or "")
    msg_type = str(msg.get("type") or "text")
    hits = {"keywords": [], "strong": [], "categories": [], "time": [], "actions": [], "markers": []}
    score = 0

    if is_noise(content, msg_type):
        return {"score": 0, "hits": hits}

    for category, weight, words in KEYWORD_GROUPS:
        matched = [w for w in words if w in content]
        if matched:
            score += weight
            hits["categories"].append(category)
            hits["keywords"].extend(matched)
            if category != WEAK_CATEGORY:
                hits["strong"].extend(matched)

    for pattern, label in TIME_PATTERNS:
        found = re.findall(pattern, content)
        if found:
            score += 2
            hits["time"].append(label)
            break  # 时间表达式只计一次，避免长句叠加虚高

    for pattern, label in ACTION_PATTERNS:
        if re.search(pattern, content):
            score += 1
            hits["actions"].append(label)

    if any(marker in content for marker in GROUP_NOTICE_MARKERS):
        score += 4
        hits["markers"].append("群公告")

    if msg_type == "system":
        score += 1
        hits["markers"].append("系统消息")

    if URL_RE.search(content):
        score += 1
        hits["markers"].append("含链接")

    # 长度适中更像通知正文；极短或极长降权
    length = len(content)
    if 12 <= length <= 400:
        score += 1
    elif length > 1200:
        score -= 1

    return {"score": score, "hits": hits}


def infer_chat_type(payload: dict, messages: list) -> str:
    """推断会话类型：group / private / unknown。

    优先使用 bundle 中显式声明的 chat_type；否则按内容特征启发式判断。
    """
    declared = str(payload.get("chat_type") or "").strip().lower()
    if declared in ("group", "private", "single", "friend"):
        return "group" if declared == "group" else "private"

    for msg in messages:
        content = str(msg.get("content") or "")
        if any(marker in content for marker in GROUP_NOTICE_MARKERS):
            return "group"
        if "@全体成员" in content:
            return "group"
    # 群聊常见：内容形如 "昵称:\n正文"
    prefixed = sum(1 for m in messages if ":\n" in str(m.get("content") or ""))
    if messages and prefixed / max(len(messages), 1) > 0.3:
        return "group"
    return "unknown"


def is_notice_like(score: int, hits: dict, threshold: int) -> bool:
    """双条件门控：既要分数够，又要有实质信号。

    仅靠「泛称呼」或「时段」凑分不算通知，避免 "同学们早上好"、"晚上一起吃饭吗"
    这类日常寒暄被误召。
    """
    if score < threshold:
        return False
    if hits.get("strong"):
        return True
    if any(label in SPECIFIC_TIME_LABELS for label in hits.get("time") or []):
        return True
    if hits.get("markers"):
        return True
    return False


def build_candidates(payload: dict, context: int = 2, threshold: int = 3,
                     max_candidates: int = 300) -> dict:
    """把单个 bundle 的 messages.json 转成候选窗口集合。"""
    messages = payload.get("messages") or []
    contact_display = payload.get("contact_display") or payload.get("contact_username") or "未知会话"
    chat_type = infer_chat_type(payload, messages)

    scored = []
    for index, msg in enumerate(messages):
        result = score_message(msg)
        scored.append((index, result["score"], result["hits"], msg))

    anchor_indices = [
        index for index, score, hits, _ in scored
        if is_notice_like(score, hits, threshold)
    ]
    if not anchor_indices:
        return {"contact_display": contact_display, "chat_type": chat_type,
                "messages": len(messages), "candidates": []}

    # 仅合并**相邻**消息（同一条通知被拆成两句），避免把不同通知并成一坨
    groups = []
    current = [anchor_indices[0]]
    for index in anchor_indices[1:]:
        if index - current[-1] <= 1:
            current.append(index)
        else:
            groups.append(current)
            current = [index]
    groups.append(current)

    candidates = []
    seen_spans = set()
    for group in groups:
        start = max(0, group[0] - context)
        end = min(len(messages) - 1, group[-1] + context)
        if (start, end) in seen_spans:
            continue
        seen_spans.add((start, end))

        window_msgs = []
        for offset in range(start, end + 1):
            msg = messages[offset]
            sender = msg.get("sender")
            sender_name = msg.get("sender_name")
            if not sender_name:
                sender_name = "我" if sender == "me" else contact_display
            window_msgs.append({
                "local_id": msg.get("local_id", offset + 1),
                "time": ts_to_str(msg.get("timestamp")),
                "sender": sender,
                "sender_name": sender_name,
                "type": msg.get("type", "text"),
                "content": str(msg.get("content") or ""),
            })

        best_score = max(score for index, score, _, _ in scored if index in group)
        merged_hits = {"keywords": [], "strong": [], "categories": [], "time": [], "actions": [], "markers": []}
        for index, score, hits, _ in scored:
            if index in group:
                for key in merged_hits:
                    merged_hits[key].extend(hits.get(key) or [])
        for key in merged_hits:
            merged_hits[key] = sorted(set(merged_hits[key]))

        anchor_msg = messages[group[0]]
        candidates.append({
            "candidate_id": "",  # 由调用方统一编号
            "bundle": payload.get("bundle_dir", ""),
            "contact_display": contact_display,
            "chat_type": chat_type,
            "score": best_score,
            "hits": merged_hits,
            "anchor_time": ts_to_str(anchor_msg.get("timestamp")),
            # 锚点消息 id：窗口会向两侧扩展，下游据此定位「通知主体」而非被卷入的上下文
            "anchor_local_ids": [messages[i].get("local_id", i + 1) for i in group],
            "window": window_msgs,
        })

    candidates.sort(key=lambda item: (-item["score"], item["anchor_time"]))
    candidates = candidates[:max_candidates]
    candidates.sort(key=lambda item: item["anchor_time"])

    for order, candidate in enumerate(candidates, start=1):
        candidate["candidate_id"] = f"c{order:04d}"

    return {"contact_display": contact_display, "chat_type": chat_type,
            "messages": len(messages), "candidates": candidates}
