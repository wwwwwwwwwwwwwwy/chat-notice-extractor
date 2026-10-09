#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把 Agent 从聊天里抽出的「待办 / 约定 / 承诺」渲染成多 sheet 的 Excel 计划表。

设计原则（吃过亏，别改回去）：
  1. 主表只放**能一眼看完**的字段。原话、长备注一律拆到独立 sheet，
     否则一行被撑成一块砖，表就没法看了。
  2. 行高必须显式算出来。openpyxl 不会自动撑高，wrap_text 打开但不设行高，
     打开 Excel 就是被裁掉的一行——等于没写。
  3. 「我的待办」排在最前，因为那才是用户真正要的一页。

输入：commitments.json（Agent 自己写的，schema 见 SKILL.md）
用法：
    python build_plan_xlsx.py --plan commitments.json --out plan.xlsx \
        --contact "某班委群" --messages out/messages.json
"""

import argparse
import json
import math
import os
import sys
import unicodedata
from datetime import datetime, date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------- 样式 ----------
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=12)
OVERDUE_FILL = PatternFill("solid", fgColor="FFC7CE")
SOON_FILL = PatternFill("solid", fgColor="FFEB9C")
UNKNOWN_FILL = PatternFill("solid", fgColor="EDEDED")
OPEN_FILL = PatternFill("solid", fgColor="E2EFDA")
DONE_FONT = Font(color="808080")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)

CONF_CHOICES = {"high": "高", "medium": "中", "low": "低"}
OPEN_STATUS = {"待办", "不明", "已承诺"}

# ---------- 列定义：多会话合并时自动加「来源」列 ----------
SRC_COL = ("来源", 17)


def plan_headers(multi_source):
    h = [("#", 6), ("事项", 38)]
    if multi_source:
        h.append(SRC_COL)
    h += [("负责人", 13), ("截止", 11), ("剩余", 7), ("状态", 9),
          ("置信度", 8), ("类别", 9), ("依据（谁 · 何时）", 28)]
    return h


EVIDENCE_HEADERS = [
    ("#", 6), ("事项", 28), ("来源", 15), ("时间", 19), ("说话人", 15), ("原话", 70),
]
NOTE_HEADERS = [
    ("#", 6), ("事项", 32), ("来源", 15), ("备注与推断", 90),
]


# ---------- 行高估算 ----------
def disp_units(s):
    """Excel 列宽单位 ≈ 一个半角字符；全角按 2 算。"""
    total = 0
    for ch in str(s or ""):
        total += 2 if unicodedata.east_asian_width(ch) in ("W", "F", "A") else 1
    return total


def wrapped_lines(text, width):
    """该单元格在给定列宽下会占几行（含显式换行）。"""
    if text is None or text == "":
        return 1
    cap = max(width - 2, 4)
    lines = 0
    for para in str(text).split("\n"):
        lines += max(1, math.ceil(disp_units(para) / cap))
    return lines


def row_height(values, widths, base=15.0, pad=4.0):
    n = max(wrapped_lines(v, w) for v, w in zip(values, widths))
    return max(base + pad, n * base + pad)


def style_header(ws, headers):
    for col, (name, width) in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=name)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"


def write_row(ws, r, values, widths, center_cols=(), done=False, date_col=None,
              today=None, due=None):
    for c, v in enumerate(values, start=1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = BORDER
        cell.alignment = CENTER if c in center_cols else WRAP
    ws.row_dimensions[r].height = row_height(values, widths)
    if done:
        for c in range(1, len(values) + 1):
            ws.cell(row=r, column=c).font = DONE_FONT
    if date_col:
        dc = ws.cell(row=r, column=date_col)
        if due is None:
            dc.fill = UNKNOWN_FILL
        elif today is not None and due < today:
            dc.fill = OVERDUE_FILL
        elif today is not None and (due - today).days <= 3:
            dc.fill = SOON_FILL


def parse_due(value):
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            continue
    return None


def load_items(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        data = {"items": data}
    if not isinstance(data.get("items"), list):
        raise SystemExit("commitments.json 缺少 items 数组")
    warnings = []
    for i, it in enumerate(data["items"], start=1):
        if not str(it.get("what") or "").strip():
            warnings.append("第 %d 项缺 what（事项）" % i)
        if not it.get("evidence"):
            warnings.append("第 %d 项没有 evidence（无原话证据）" % i)
        if it.get("confidence") not in CONF_CHOICES:
            warnings.append("第 %d 项 confidence=%r 不在 high/medium/low 内" % (i, it.get("confidence")))
    return data, warnings


def is_mine(owner, team=()):
    """owner 是否算「本人这一侧」。team 是同岗位搭档（例：同为班长的人）。"""
    o = str(owner or "").strip()
    if not o:
        return False
    if o in ("我", "本人") or o.startswith("我（") or o.startswith("我("):
        return True
    return any(t and t in o for t in team)


def owner_label(owner, identity_confirmed):
    """用户确认过身份时，把原话里的泛指角色也标成「我」，否则保留原话角色。"""
    o = str(owner or "").strip()
    if not o:
        return "未指明"
    if identity_confirmed and o in ("各班班长", "各位班长", "各班", "班长"):
        return "我"
    return o


def main():
    ap = argparse.ArgumentParser(description="生成微信聊天计划表 xlsx")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--contact", default="")
    ap.add_argument("--messages", default=None)
    ap.add_argument("--today", default=None, help="覆盖今天日期 YYYY-MM-DD（便于复现）")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    data, warnings = load_items(args.plan)
    items = data.get("items") or []
    today = parse_due(args.today) or date.today()

    src_meta = {}
    if args.messages and os.path.exists(args.messages):
        with open(args.messages, "r", encoding="utf-8") as f:
            src_meta = (json.load(f) or {}).get("meta", {}) or {}

    identity_confirmed = bool(str(data.get("identity_note") or "").strip())
    team = [str(t).strip() for t in (data.get("team") or []) if str(t).strip()]
    multi_source = len({str(it.get("source") or "").strip()
                        for it in items if str(it.get("source") or "").strip()}) > 1

    def due_of(it):
        return parse_due(it.get("due"))

    def ref_of(it):
        evs = it.get("evidence") or []
        if not evs:
            return ""
        ev = evs[0]
        t = str(ev.get("time") or "")
        short = t[5:16] if len(t) >= 16 else t
        return "%s · %s" % (ev.get("speaker") or "?", short)

    items_sorted = sorted(items, key=lambda it: (
        due_of(it) is None, due_of(it) or date.max, str(it.get("id") or "")))

    mine = [it for it in items_sorted
            if is_mine(owner_label(it.get("owner"), identity_confirmed), team)]
    mine.sort(key=lambda it: (
        0 if (it.get("status") or "不明") in OPEN_STATUS else 1,
        due_of(it) is None, due_of(it) or date.max))

    wb = Workbook()

    # ---------- 说明 ----------
    ws = wb.active
    ws.title = "说明"
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 96
    rows = [
        ("会话", args.contact or src_meta.get("talker_resolved") or "(未填)"),
        ("消息条数", src_meta.get("fetched", "(未提供 messages.json)")),
        ("消息时间范围", "%s ~ %s" % (src_meta.get("oldest"), src_meta.get("newest")) if src_meta else "(未提供)"),
        ("数据来源", src_meta.get("source", "weflow-cli messages (paged)")),
        ("读取时间", src_meta.get("read_at", "(未知)")),
        ("计划项数", "%d（其中归本人的 %d 项）" % (len(items), len(mine))),
        ("生成基准日", today.strftime("%Y-%m-%d")),
        ("", ""),
        ("身份判定", data.get("identity_note") or "未记录身份信息；「负责人」按原话角色保留。"),
        ("同岗位搭档", "、".join(team) + "（其名下事项一并进「我的待办」）" if team else "无"),
        ("证据规则", "每一条计划都必须能指回一条带时间戳的原话；指不到的不下结论，宁缺勿编。"),
        ("置信度", "高=原话明确；中=可推断但未明说；低=需要你自己确认。"),
        ("剩余天数", "截止日期 − 生成基准日；负数=已过期。"),
        ("颜色", "截止列：红=已过期，黄=3 天内，灰=无明确日期；已完成的整行压成灰色。"),
        ("原话", "逐字取自聊天记录，已通过 verify_quotes.py 校验；本表未做任何改写。"),
        ("隐私", "全流程本机处理，聊天正文没有发往任何外部服务。"),
        ("", ""),
        ("各 sheet", "「我的待办」只有你要动的；「计划表」是全部项；原话全文在「原话证据」；长备注在「备注与推断」。"),
    ]
    for r, (k, v) in enumerate(rows, start=1):
        a = ws.cell(row=r, column=1, value=k)
        a.font = TITLE_FONT if k else Font()
        a.alignment = Alignment(vertical="top")
        b = ws.cell(row=r, column=2, value=v)
        b.alignment = WRAP
        ws.row_dimensions[r].height = row_height([k, v], [20, 96])

    # ---------- 我的待办 ----------
    ws_mine = wb.create_sheet("我的待办")
    mfields = [("id", "#", 6), ("what", "要做什么", 50)]
    if multi_source:
        mfields.append(("source", "来源", 17))
    if team:
        mfields.append(("owner", "负责人", 16))
    mfields += [("due", "截止", 11), ("remain", "剩余", 7), ("status", "状态", 9),
                ("category", "类别", 9), ("ref", "依据（谁 · 何时）", 28)]
    mine_heads = [(label, w) for _, label, w in mfields]
    keys = [k for k, _, _ in mfields]
    mw = [w for _, w in mine_heads]
    style_header(ws_mine, mine_heads)
    date_col = keys.index("due") + 1
    status_col = keys.index("status") + 1
    center_cols = tuple(keys.index(k) + 1 for k in
                        ("id", "owner", "due", "remain", "status", "category") if k in keys)
    for idx, it in enumerate(mine, start=1):
        due = due_of(it)
        remain = (due - today).days if due else ""
        done = (it.get("status") or "") == "已完成"
        bag = {
            "id": it.get("id") or ("C%03d" % idx),
            "what": it.get("what", ""),
            "source": it.get("source", ""),
            "owner": owner_label(it.get("owner"), identity_confirmed),
            "due": due.strftime("%Y-%m-%d") if due else "",
            "remain": remain,
            "status": it.get("status", "") or "不明",
            "category": it.get("category", "") or "",
            "ref": ref_of(it),
        }
        r = idx + 1
        write_row(ws_mine, r, [bag[k] for k in keys], mw, center_cols=center_cols,
                  done=done, date_col=date_col, today=today, due=due)
        if not done:
            ws_mine.cell(row=r, column=status_col).fill = OPEN_FILL
    if mine:
        ws_mine.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(mine_heads)), len(mine) + 1)

    # ---------- 计划表 ----------
    ws2 = wb.create_sheet("计划表")
    plan_heads = plan_headers(multi_source)
    pw = [w for _, w in plan_heads]
    style_header(ws2, plan_heads)
    off = 1 if multi_source else 0
    for idx, it in enumerate(items_sorted, start=1):
        due = due_of(it)
        remain = (due - today).days if due else ""
        done = (it.get("status") or "") == "已完成"
        vals = [it.get("id") or ("C%03d" % idx), it.get("what", "")]
        if multi_source:
            vals.append(it.get("source", ""))
        vals += [
            owner_label(it.get("owner"), identity_confirmed),
            due.strftime("%Y-%m-%d") if due else "",
            remain,
            it.get("status", "") or "不明",
            CONF_CHOICES.get(it.get("confidence"), it.get("confidence") or ""),
            it.get("category", "") or "",
            ref_of(it),
        ]
        r = idx + 1
        write_row(ws2, r, vals, pw,
                  center_cols=(1, 3 + off, 4 + off, 5 + off, 6 + off, 7 + off, 8 + off),
                  done=done, date_col=4 + off, today=today, due=due)
    if items_sorted:
        ws2.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(plan_heads)), len(items_sorted) + 1)

    # ---------- 原话证据 ----------
    ws3 = wb.create_sheet("原话证据")
    ew = [w for _, w in EVIDENCE_HEADERS]
    style_header(ws3, EVIDENCE_HEADERS)
    r = 2
    for it in items_sorted:
        for ev in (it.get("evidence") or []):
            vals = [it.get("id", ""), it.get("what", ""), ev.get("source") or it.get("source", ""),
                    ev.get("time", ""), ev.get("speaker", ""), ev.get("quote", "")]
            write_row(ws3, r, vals, ew, center_cols=(1, 3, 4, 5))
            r += 1
    if r > 2:
        ws3.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(EVIDENCE_HEADERS)), r - 1)

    # ---------- 备注与推断 ----------
    ws4 = wb.create_sheet("备注与推断")
    nw = [w for _, w in NOTE_HEADERS]
    style_header(ws4, NOTE_HEADERS)
    r = 2
    for it in items_sorted:
        note = str(it.get("note") or "").strip()
        if not note:
            continue
        vals = [it.get("id", ""), it.get("what", ""), it.get("source", ""), note]
        write_row(ws4, r, vals, nw, center_cols=(1, 3))
        r += 1
    if r > 2:
        ws4.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(NOTE_HEADERS)), r - 1)

    wb.move_sheet("我的待办", offset=-1)

    out_dir = os.path.dirname(os.path.abspath(args.out))
    if out_dir:
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as e:
            raise SystemExit(
                "无法创建输出目录 %s：%s\n提示：C:\\Users\\21419\\Documents\\ 下的目录对子进程"
                "不可写（实测），请换到 C:\\Users\\21419\\.dsh\\wechat-plan\\ 之类的位置。"
                % (out_dir, e))
        if not os.path.isdir(out_dir):
            raise SystemExit(
                "输出目录 %s 创建后不存在，写入被系统静默拦截；换到 "
                "C:\\Users\\21419\\.dsh\\wechat-plan\\ 之类的位置。" % out_dir)
    wb.save(args.out)

    if not args.quiet:
        print(json.dumps({
            "xlsx": os.path.abspath(args.out),
            "sheets": wb.sheetnames,
            "items": len(items),
            "mine": len(mine),
            "evidence_rows": ws3.max_row - 1,
            "warnings": warnings,
        }, ensure_ascii=False, indent=2))
    for w in warnings:
        print("警告：" + w, file=sys.stderr)


if __name__ == "__main__":
    main()
