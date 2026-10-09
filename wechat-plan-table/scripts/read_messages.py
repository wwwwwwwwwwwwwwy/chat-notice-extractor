#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""读取本机微信聊天记录 -> 归一化 messages.json + 人类可读 messages.txt

数据通路（本机 2026-10-09 实测，weflow-cli 1.9.0）：
    只使用 `weflow-cli messages <talker> -n <page> -o <offset> --json`
    分页读取。实测 offset 分页可靠：无重叠、无循环、越界返回空数组。

    不要改用 `weflow-cli export`：同一台机器上该子命令（json/excel、
    带不带 --contract、带不带 --json）都会空转 7 分钟以上、不产出任何文件、
    不打印任何输出，CPU 打满单核。属已知坏路径，别浪费时间重试。

用法：
    python read_messages.py --talker "某班委群"
    python read_messages.py --talker 12345678901@chatroom --out "C:\\Users\\21419\\.dsh\\wechat-plan\\bwq" --since 2026-09-01

输出（默认落在 C:\\Users\\21419\\.dsh\\wechat-plan\\<会话短名>\\）：
    <out>/messages.json  结构化全量（含抓取元信息与逐条消息）
    <out>/messages.txt   按时间正序的 [时间] 说话人: 内容，供阅读与抽取
    stdout               只打印计数与时间范围，不含聊天正文

注意：不要把 --out 指到 C:\\Users\\21419\\Documents\\ 下面——实测子进程写不进去。
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime

# 本机实测可用的 DSH 捆绑解释器（Python 3.12.14 + openpyxl 3.1.5）
BUNDLED_PYTHON = (
    r"C:\Users\21419\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
)

# 默认输出根目录。刻意不放在 Documents\ 下面：本机实测（2026-10-09）
# C:\Users\21419\Documents\ 及其子目录对子进程是写不进去的——os.makedirs /
# open(...,'w') 抛 WinError 2，PowerShell 的 New-Item 更坏，它假装成功但
# Test-Path 立刻是 False。所以默认落在这里。
DEFAULT_OUT_ROOT = r"C:\Users\21419\.dsh\wechat-plan"

# 空的 / 无信息量的正文，统一成占位符，避免污染抽取
EMPTY_PLACEHOLDERS = {"", None}


def slugify(name):
    s = re.sub(r"[\\/:*?\"<>|\s]+", "_", str(name)).strip("_")
    return s[:60] or "chat"


def ensure_dir(path):
    """建目录并确认它真的存在——被静默拦截时给一句人话，别丢 WinError 2。"""
    try:
        os.makedirs(path, exist_ok=True)
    except OSError as e:
        raise SystemExit(
            "无法创建目录 %s\n原因：%s\n"
            "提示：C:\\Users\\21419\\Documents\\ 下的目录对子进程不可写（实测），"
            "请把 --out 换到 C:\\Users\\21419\\.dsh\\wechat-plan\\ 之类的位置。" % (path, e)
        )
    if not os.path.isdir(path):
        raise SystemExit(
            "目录 %s 创建后不存在，写入被系统静默拦截了。"
            "请把 --out 换到 C:\\Users\\21419\\.dsh\\wechat-plan\\ 之类的位置。" % path
        )
    probe = os.path.join(path, ".write_probe")
    try:
        with open(probe, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(probe)
    except OSError as e:
        raise SystemExit("目录 %s 不可写：%s" % (path, e))
    return path


def find_cli():
    cli = shutil.which("weflow-cli")
    if cli:
        return cli
    fallback = os.path.join(
        os.environ.get("APPDATA", ""), "npm", "weflow-cli.CMD"
    )
    if os.path.exists(fallback):
        return fallback
    return None


def run_cli(cli, args, timeout=300):
    """.CMD 垫片在 Windows 上可直接 exec；失败时退回 cmd.exe /c。"""
    cmd = [cli] + list(args)
    try:
        p = subprocess.run(cmd, capture_output=True, timeout=timeout)
    except OSError:
        comspec = os.environ.get("ComSpec", "cmd.exe")
        p = subprocess.run(
            [comspec, "/c", cli] + list(args), capture_output=True, timeout=timeout
        )
    out = p.stdout.decode("utf-8", errors="replace")
    err = p.stderr.decode("utf-8", errors="replace")
    return p.returncode, out, err


def to_epoch(value, is_end=False):
    """接受 epoch 秒 / YYYY-MM-DD / ISO 时间串，返回 epoch 秒。"""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    if s.isdigit():
        return int(s)
    s = s.replace("Z", "+00:00")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(s, fmt)
            if fmt == "%Y-%m-%d" and is_end:
                dt = dt.replace(hour=23, minute=59, second=59)
            return int(dt.timestamp())
        except ValueError:
            continue
    try:
        return int(datetime.fromisoformat(s).timestamp())
    except ValueError:
        raise SystemExit("无法解析时间：%r（用 epoch 秒 / YYYY-MM-DD / ISO 串）" % value)


def norm_sender(msg):
    if msg.get("isSend") == 1:
        return "我"
    display = (msg.get("senderDisplay") or "").strip()
    if display:
        return display
    username = (msg.get("senderUsername") or "").strip()
    return username or "未知"


def pick_text(msg):
    for key in ("parsedContent", "content", "rawContent"):
        v = msg.get(key)
        if v not in EMPTY_PLACEHOLDERS and str(v).strip():
            return str(v).strip()
    return "[空消息]"


def normalize(msg):
    ts = int(msg.get("createTime") or 0)
    return {
        "localId": msg.get("localId"),
        "ts": ts,
        "time": datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S") if ts else "",
        "sender": norm_sender(msg),
        "is_send": msg.get("isSend"),
        "localType": msg.get("localType"),
        "text": pick_text(msg),
    }


def main():
    ap = argparse.ArgumentParser(description="分页读取本机微信聊天记录")
    ap.add_argument("--talker", required=True, help="会话 ID(wxid/xxx@chatroom)、昵称或备注")
    ap.add_argument(
        "--label",
        default=None,
        help="这个会话的人类可读名字（写进 meta.source_label，多群合并时用来标注来源）",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="输出目录；默认 %s\\<会话短名>" % DEFAULT_OUT_ROOT,
    )
    ap.add_argument("--page", type=int, default=200, help="每页条数（默认 200）")
    ap.add_argument("--max", type=int, default=0, help="最多抓取条数，0=全量（默认 0）")
    ap.add_argument("--since", default=None, help="只保留该时间之后，YYYY-MM-DD 或 ISO")
    ap.add_argument("--until", default=None, help="只保留该时间之前，YYYY-MM-DD 或 ISO")
    ap.add_argument("--timeout", type=int, default=300, help="单次调用超时秒数")
    args = ap.parse_args()

    cli = find_cli()
    if not cli:
        raise SystemExit("找不到 weflow-cli。先装：npm i -g weflow-cli，再 weflow-cli init")

    since_ts = to_epoch(args.since)
    until_ts = to_epoch(args.until, is_end=True)

    if not args.out:
        args.out = os.path.join(DEFAULT_OUT_ROOT, slugify(args.talker))
    ensure_dir(args.out)

    collected = []
    seen_ids = set()
    offset = 0
    pages = 0
    loop_detected = False
    talker_resolved = None
    truncated_by_since = False

    while True:
        rc, out, err = run_cli(
            cli,
            [
                "messages", args.talker,
                "-n", str(args.page),
                "-o", str(offset),
                "--json", "--non-interactive",
            ],
            timeout=args.timeout,
        )
        if rc != 0:
            raise SystemExit(
                "weflow-cli messages 失败 rc=%d\n%s\n%s" % (rc, out[:500], err[:500])
            )
        try:
            data = json.loads(out)
        except ValueError:
            raise SystemExit("weflow-cli 返回的不是 JSON：\n%s" % out[:800])
        if not data.get("success"):
            raise SystemExit("weflow-cli 报告失败：%s" % json.dumps(data, ensure_ascii=False)[:500])

        talker_resolved = data.get("talker") or talker_resolved
        batch = data.get("messages") or []
        if not batch:
            break

        pages += 1
        page_ids = [m.get("localId") for m in batch]
        if any(i in seen_ids for i in page_ids):
            loop_detected = True
            break
        seen_ids.update(page_ids)

        stop = False
        for m in batch:
            n = normalize(m)
            if since_ts is not None and n["ts"] < since_ts:
                # 消息按时间倒序返回，本页已进入早于 since 的区间，无需再翻
                truncated_by_since = True
                stop = True
                continue
            if until_ts is not None and n["ts"] > until_ts:
                continue
            collected.append(n)

        if stop:
            break
        if args.max and len(collected) >= args.max:
            collected = collected[: args.max]
            break

        offset += len(batch)

    collected.sort(key=lambda x: (x["ts"], x["localId"] or 0))

    meta = {
        "talker_input": args.talker,
        "talker_resolved": talker_resolved,
        "source_label": args.label or talker_resolved or args.talker,
        "fetched": len(collected),
        "pages": pages,
        "loop_detected": loop_detected,
        "stopped_at_since": truncated_by_since,
        "since": args.since,
        "until": args.until,
        "oldest": collected[0]["time"] if collected else None,
        "newest": collected[-1]["time"] if collected else None,
        "read_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source": "weflow-cli messages (paged)",
    }

    with open(os.path.join(args.out, "messages.json"), "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "messages": collected}, f, ensure_ascii=False, indent=2)

    with open(os.path.join(args.out, "messages.txt"), "w", encoding="utf-8") as f:
        for n in collected:
            # 消息正文里的换行要压成可见分隔符，否则一条多行消息在 txt 里
            # 看起来像好几条，抽取时会把同一句话拆成不同的"原话"
            flat = re.sub(r"\s*\r?\n\s*", " ⏎ ", n["text"])
            f.write("[%s] %s: %s\n" % (n["time"], n["sender"], flat))

    # stdout 只给计数，不倒聊天正文
    print(json.dumps(meta, ensure_ascii=False, indent=2))
    if loop_detected:
        print("警告：检测到分页重复 localId，已在重复处停止，结果可能不完整。", file=sys.stderr)


if __name__ == "__main__":
    main()
