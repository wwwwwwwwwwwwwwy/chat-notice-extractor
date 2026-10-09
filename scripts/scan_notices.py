"""scan_notices.py - 遍历聊天记录 bundle，产出「通知候选窗口」

用法：
    # 扫描整个 data/contacts 目录下的所有会话
    python scripts/scan_notices.py --contacts-dir data/contacts --out out/candidates.json

    # 只扫单个 messages.json
    python scripts/scan_notices.py --input data/contacts/xx__abcd1234/messages.json --out out/candidates.json

可选：
    --meta meta.json    会话元信息覆盖（{"<bundle目录名>": {"chat_type": "group"}}）
    --threshold 5       规则打分阈值，越低召回越多、噪声越大
    --context 2         候选窗口前后各取几条上下文
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from notice_rules import TZ, build_candidates  # noqa: E402

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def load_json(path):
    # 兼容 Windows 编辑器写出的 UTF-8 BOM
    with open(path, encoding="utf-8-sig") as handle:
        return json.load(handle)


def discover_bundles(contacts_dir):
    """在目录下查找所有 */messages.json，返回 (bundle名, 路径) 列表。"""
    found = []
    for entry in sorted(os.listdir(contacts_dir)):
        candidate = os.path.join(contacts_dir, entry, "messages.json")
        if os.path.isfile(candidate):
            found.append((entry, candidate))
    if not found:
        direct = os.path.join(contacts_dir, "messages.json")
        if os.path.isfile(direct):
            found.append((os.path.basename(os.path.abspath(contacts_dir)), direct))
    return found


def main():
    parser = argparse.ArgumentParser(description="从聊天记录中召回通知候选窗口")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--contacts-dir", help="包含多个会话子目录的根目录（如 data/contacts）")
    source.add_argument("--input", help="单个 messages.json 路径")
    parser.add_argument("--out", required=True, help="候选结果输出路径 candidates.json")
    parser.add_argument("--meta", help="会话元信息 JSON，用于覆盖 chat_type")
    parser.add_argument("--threshold", type=int, default=3, help="规则打分阈值（默认 3）")
    parser.add_argument("--context", type=int, default=2, help="候选窗口上下文条数（默认 2）")
    parser.add_argument("--max-candidates", type=int, default=300, help="单会话候选上限")
    args = parser.parse_args()

    meta = load_json(args.meta) if args.meta and os.path.isfile(args.meta) else {}

    if args.input:
        bundles = [(os.path.basename(os.path.dirname(os.path.abspath(args.input))), args.input)]
    else:
        if not os.path.isdir(args.contacts_dir):
            print(json.dumps({"status": "error",
                              "error": f"目录不存在: {args.contacts_dir}"}, ensure_ascii=False),
                  file=sys.stderr)
            sys.exit(1)
        bundles = discover_bundles(args.contacts_dir)
        if not bundles:
            print(json.dumps({"status": "error",
                              "error": f"{args.contacts_dir} 下未找到任何 */messages.json"},
                             ensure_ascii=False), file=sys.stderr)
            sys.exit(1)

    all_candidates = []
    bundle_reports = []
    total_messages = 0

    for bundle_name, path in bundles:
        try:
            payload = load_json(path)
        except (OSError, ValueError) as exc:
            bundle_reports.append({"bundle": bundle_name, "error": str(exc)})
            continue

        override = meta.get(bundle_name) or {}
        if override.get("chat_type"):
            payload["chat_type"] = override["chat_type"]
        if override.get("display_name"):
            payload["contact_display"] = override["display_name"]

        result = build_candidates(payload, context=args.context, threshold=args.threshold,
                                  max_candidates=args.max_candidates)
        for candidate in result["candidates"]:
            candidate["bundle"] = bundle_name
        all_candidates.extend(result["candidates"])
        total_messages += result["messages"]
        bundle_reports.append({
            "bundle": bundle_name,
            "contact_display": result["contact_display"],
            "chat_type": result["chat_type"],
            "messages": result["messages"],
            "candidates": len(result["candidates"]),
        })

    # 全局重新编号，保证 candidate_id 在整批里唯一
    all_candidates.sort(key=lambda item: (item["anchor_time"], item["bundle"]))
    for order, candidate in enumerate(all_candidates, start=1):
        candidate["candidate_id"] = f"c{order:04d}"

    output = {
        "generated_at": datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "threshold": args.threshold,
        "context": args.context,
        "stats": {
            "bundles": len(bundle_reports),
            "messages": total_messages,
            "candidates": len(all_candidates),
        },
        "bundles": bundle_reports,
        "candidates": all_candidates,
    }

    out_dir = os.path.dirname(os.path.abspath(args.out))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as handle:
        json.dump(output, handle, ensure_ascii=False, indent=2)

    print(json.dumps({
        "status": "ok",
        "bundles": len(bundle_reports),
        "messages": total_messages,
        "candidates": len(all_candidates),
        "output": os.path.abspath(args.out),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
