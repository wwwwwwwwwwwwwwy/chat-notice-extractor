"""xlsx_min.py - 极简 .xlsx 写入器（纯标准库）

为什么需要它：目标机器上未必装有 openpyxl。本模块用 zipfile + 手写 XML 生成
最小可用但格式完整的 xlsx（表头加粗填色、列宽、自动换行、冻结首行、筛选器），
保证技能在任何 Python 3.8+ 环境都能出表。

.xlsx 本质是一个 zip 包，内部结构：
    [Content_Types].xml
    _rels/.rels
    xl/workbook.xml
    xl/_rels/workbook.xml.rels
    xl/styles.xml
    xl/worksheets/sheet1.xml
文本一律用 inlineStr 写入，规避 sharedStrings 表带来的额外复杂度。
"""

from __future__ import annotations

import re
import zipfile
from xml.sax.saxutils import escape

_ILLEGAL_XML = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]"
)


def _clean(value) -> str:
    """转字符串并剔除 XML 非法控制字符（聊天正文里确实可能出现）。"""
    text = "" if value is None else str(value)
    return _ILLEGAL_XML.sub("", text)


def _col_letter(index: int) -> str:
    """0 → A, 25 → Z, 26 → AA"""
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _cell(ref: str, value, style: int) -> str:
    text = _clean(value)
    if text == "":
        return f'<c r="{ref}" s="{style}"/>'
    return (
        f'<c r="{ref}" s="{style}" t="inlineStr">'
        f"<is><t xml:space=\"preserve\">{escape(text)}</t></is></c>"
    )


def write_xlsx(path: str, sheet_name: str, headers: list, rows: list,
               col_widths: list | None = None, wrap_columns: set | None = None,
               autofilter: bool = True, freeze_header: bool = True) -> None:
    """写出一个单表 xlsx。

    headers     : 表头文本列表
    rows        : 二维列表，行 × 列
    col_widths  : 每列字符宽度（Excel 单位），缺省按表头长度估算
    wrap_columns: 需要自动换行的列下标集合（0 基）
    """
    wrap_columns = wrap_columns or set()
    if col_widths is None:
        col_widths = []
        for position, header in enumerate(headers):
            longest = len(_clean(header))
            for row in rows[:200]:
                if position < len(row):
                    longest = max(longest, len(_clean(row[position])))
            col_widths.append(min(max(longest + 2, 8), 60))

    sheet_name = _clean(sheet_name)[:31] or "Sheet1"
    total_cols = max(len(headers), 1)
    last_col = _col_letter(total_cols - 1)

    cols_xml = "".join(
        f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
        for i, w in enumerate(col_widths)
    )

    row_xml = []
    header_cells = "".join(
        _cell(f"{_col_letter(i)}1", header, 1) for i, header in enumerate(headers)
    )
    row_xml.append(f'<row r="1" ht="22" customHeight="1">{header_cells}</row>')

    for offset, row in enumerate(rows, start=2):
        cells = []
        for col, value in enumerate(row):
            style = 2 if col in wrap_columns else 0
            cells.append(_cell(f"{_col_letter(col)}{offset}", value, style))
        row_xml.append(f'<row r="{offset}">{"".join(cells)}</row>')

    last_row = len(rows) + 1
    pane = (
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        if freeze_header else ""
    )
    autofilter_xml = (
        f'<autoFilter ref="A1:{last_col}{last_row}"/>' if autofilter and rows else ""
    )

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:{last_col}{max(last_row, 1)}"/>'
        f'<sheetViews><sheetView workbookViewId="0">{pane}</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="16"/>'
        f"<cols>{cols_xml}</cols>"
        f"<sheetData>{''.join(row_xml)}</sheetData>"
        f"{autofilter_xml}"
        "</worksheet>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="\u5fae\u8f6f\u96c5\u9ed1"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/>'
        '<name val="\u5fae\u8f6f\u96c5\u9ed1"/></font>'
        "</fonts>"
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid">'
        '<fgColor rgb="FF2F5597"/><bgColor indexed="64"/></patternFill></fill>'
        "</fills>"
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" '
        'applyFill="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
        '<alignment vertical="top" wrapText="1"/></xf>'
        "</cellXfs>"
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="0"/>'
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
        "</styleSheet>"
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )

    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/styles.xml", styles_xml)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
