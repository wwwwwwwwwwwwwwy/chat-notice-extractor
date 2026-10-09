"""demo_data.py - 生成合成聊天记录，用于端到端自测

生成两个会话 bundle：
  1. 2024级计科1班群  —— 群聊，含会议通知、报名接龙、缴费截止、放假调整等
  2. 张三              —— 单聊，含约时间、待办、改期等

用法：
    python scripts/demo_data.py --out-dir demo/data/contacts
"""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime, timedelta, timezone

TZ = timezone(timedelta(hours=8))
BASE = datetime(2026, 3, 2, 9, 0, 0, tzinfo=TZ)


def ts(minutes: int) -> int:
    return int((BASE + timedelta(minutes=minutes)).timestamp())


GROUP_MESSAGES = [
    ("them", "班长", "text", "同学们早上好"),
    ("them", "李四", "text", "早"),
    ("me", "我", "text", "早啊"),
    ("them", "班长", "text", "通知一下：本周五（3月6日）下午3点在实验楼B301开班会，请全体成员务必参加，会上要确认毕业设计选题。"),
    ("them", "王五", "text", "收到"),
    ("them", "李四", "text", "收到"),
    ("me", "我", "text", "收到"),
    ("them", "班长", "text", "另外接龙报名春游，下周六（3月14日）去青龙峡，费用120元，报名截止到3月10日晚上8点，回复我一下我统计人数。"),
    ("them", "王五", "text", "1 王五"),
    ("them", "李四", "text", "2 李四"),
    ("them", "赵六", "text", "哈哈哈这个不错"),
    ("them", "班长", "text", "@全体成员 补充：班费每人还要交50元，请在3月8日前交给生活委员，逾期不再受理。"),
    ("me", "我", "text", "3 我"),
    ("them", "李四", "text", "晚上一起吃饭吗"),
    ("them", "王五", "text", "不了，我减肥"),
    ("them", "辅导员", "text", "各位同学注意：接教务处通知，因设备检修，原定3月12日的英语四级模拟考推迟到3月19日上午9点，考场不变，仍为B201。请相互转告。"),
    ("them", "班长", "text", "好的老师"),
    ("them", "李四", "text", "谢谢老师"),
    ("them", "班长", "text", "还有个事，大家记得把学生证复印件交到办公室，最晚下周三之前。"),
    ("them", "王五", "text", "嗯"),
    ("them", "李四", "text", "好的"),
]

PRIVATE_MESSAGES = [
    ("them", "张三", "text", "在吗"),
    ("me", "我", "text", "在"),
    ("them", "张三", "text", "明天下午三点老地方见吧，我把合同带过去给你看。"),
    ("me", "我", "text", "行"),
    ("them", "张三", "text", "哦对了，周五之前得把报价单发我邮箱，客户催得比较紧。"),
    ("me", "我", "text", "收到，我这两天弄"),
    ("them", "张三", "text", "哈哈昨天那个视频太好笑了"),
    ("me", "我", "text", "确实"),
    ("them", "张三", "text", "改一下时间，明天三点我不行了，改到晚上七点，还是老地方，星巴克。"),
    ("me", "我", "text", "好的"),
]


def build_payload(display, username, chat_type, rows, source="wechat"):
    messages = []
    for index, (sender, sender_name, msg_type, content) in enumerate(rows, start=1):
        messages.append({
            "local_id": index,
            "sender": sender,
            "sender_name": sender_name,
            "content": content,
            "timestamp": ts(index * 7),
            "type": msg_type,
        })
    return {
        "source": source,
        "chat_type": chat_type,
        "contact_username": username,
        "contact_display": display,
        "own_wxid": "wxid_me_demo",
        "total": len(messages),
        "messages": messages,
    }


def write_bundle(root, slug, payload):
    bundle_dir = os.path.join(root, slug)
    os.makedirs(bundle_dir, exist_ok=True)
    payload["bundle_dir"] = bundle_dir
    with open(os.path.join(bundle_dir, "messages.json"), "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    return bundle_dir


def main():
    parser = argparse.ArgumentParser(description="生成合成聊天记录用于自测")
    parser.add_argument("--out-dir", default="demo/data/contacts")
    args = parser.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)
    write_bundle(args.out_dir, "2024级计科1班群__demo0001",
                 build_payload("2024级计科1班群", "group_demo", "group", GROUP_MESSAGES))
    write_bundle(args.out_dir, "张三__demo0002",
                 build_payload("张三", "wxid_zhangsan", "private", PRIVATE_MESSAGES))

    print(json.dumps({
        "status": "ok",
        "out_dir": os.path.abspath(args.out_dir),
        "bundles": 2,
        "messages": len(GROUP_MESSAGES) + len(PRIVATE_MESSAGES),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
