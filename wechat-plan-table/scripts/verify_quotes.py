#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""校验 commitments.json 里的每条 evidence.quote 是否真的是聊天原文（支持多会话）。

为什么需要它：微信的 @ 分隔符是 U+2005（FOUR-PER-EM SPACE），不是普通空格；
[引用] 里的换行、消息内的换行也都容易被写歪。本脚本按 (会话, 时间戳) 把 quote
对回 messages.json 的真实消息逐字比对。

用法：
    python verify_quotes.py --plan plan.json --messages a/messages.json b/messages.json
    python verify_quotes.py --plan plan.json --messages */messages.json --fix

evidence 里可以带 "source"（会话名，与 --messages 文件 meta.talker_resolved/contact 对应）；
不带 source 时按时间戳在全部会话里找。

--fix：把对得上的 quote 直接替换为库里的原文（整条消息文本）。
       对不上的**不会**被改动，只报错——宁缺勿编。

退出码：0 = 全部逐字一致；1 = 存在对不上的 quote。
"""

import argparse
import glob
import json
import os
import re
import sys


def norm(s):
    # 所有空白（含 U+2005 / U+00A0 / 换行）折叠成单个 ASCII 空格。
    # U+23CE「⏎」是 messages.txt 里换行的可见写法，一并当分隔符。
    return re.sub(
        r"[\s\u00a0\u2000-\u200f\u2028\u2029\u23ce]+", " ", str(s or "")
    ).strip()


def load_sources(patterns):
    """返回 (by_key, by_time, sources)；key 为 (source, time)。"""
    files = []
    for p in patterns:
        hit = glob.glob(p)
        files.extend(hit if hit else [p])
    by_key, by_time, sources = {}, {}, []
    for fp in files:
        if not os.path.exists(fp):
            print("警告：找不到 %s" % fp, file=sys.stderr)
            continue
        with open(fp, "r", encoding="utf-8") as f:
            blob = json.load(f) or {}
        meta = blob.get("meta") or {}
        src = (meta.get("source_label") or meta.get("talker_resolved")
               or meta.get("talker_input") or os.path.basename(os.path.dirname(fp)))
        sources.append(src)
        for m in blob.get("messages") or []:
            by_key.setdefault((src, m.get("time")), []).append(m)
            by_time.setdefault(m.get("time"), []).append(m)
    return by_key, by_time, sources


def main():
    ap = argparse.ArgumentParser(description="校验 evidence.quote 是否为聊天原文")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--messages", required=True, nargs="+",
                    help="一个或多个 messages.json（可用通配符）")
    ap.add_argument("--fix", action="store_true")
    args = ap.parse_args()

    by_key, by_time, sources = load_sources(args.messages)
    print(json.dumps({"loaded_sources": sources, "indexed_times": len(by_time)},
                     ensure_ascii=False), file=sys.stderr)

    with open(args.plan, "r", encoding="utf-8") as f:
        plan = json.load(f)

    ok, bad, fixed = 0, [], 0
    multi_source = len(set(sources)) > 1

    for it in plan.get("items", []):
        iid = it.get("id") or it.get("what", "?")
        for ev in (it.get("evidence") or []):
            quote, t = ev.get("quote", ""), ev.get("time")
            src = ev.get("source") or it.get("source")
            cands = by_key.get((src, t)) if src else None
            if not cands:
                cands = by_time.get(t) or []
            nq = norm(quote)
            hit = None
            for m in cands:
                if nq and nq in norm(m.get("text", "")):
                    hit = m
                    break
            if hit is not None:
                ok += 1
                if args.fix and quote != hit.get("text"):
                    ev["quote"] = hit.get("text")
                    fixed += 1
                continue
            loose = [m for m in (by_time.get(t) or [])
                     if nq and nq in norm(m.get("text", ""))]
            bad.append({
                "id": iid,
                "source": src,
                "time": t,
                "quote_head": nq[:60],
                "reason": "该时间戳下找不到这句话" if cands else "没有任何会话有这个时间戳",
                "time_hit_elsewhere": [m.get("time") for m in loose[:3]],
            })

    if args.fix and fixed:
        with open(args.plan, "w", encoding="utf-8") as f:
            json.dump(plan, f, ensure_ascii=False, indent=2)

    print(json.dumps({"verified": ok, "mismatched": len(bad), "fixed": fixed,
                      "details": bad}, ensure_ascii=False, indent=2))
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
