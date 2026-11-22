# -*- coding: utf-8 -*-
"""Markdown 转 Excel（两种导出模式）+ OCR 精确复制直写。

模式一「还原文档结构」：每个源文档写成一个工作表，按原文顺序输出标题、段落、
表格（含合并单元格）与公式，尽量保留版式。
模式二「规整表格数据」：抽取全部表格，纵向拼接到同一个工作表，便于后续筛选分析。
精确复制（convert_exact）：不走 Markdown，直接消费 table_rebuild 的几何重建
结果（TableBlock / TextBlock）——表格网格直写、全单元格细边框、首行表头加粗
底纹、行高按原图行距还原；文字块逐行合并、无边框，按原文档顺序排列。
"""
from __future__ import annotations

import html
import re
from collections import Counter
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

# ---------------------------------------------------------------- 数据模型


@dataclass
class Block:
    kind: str  # title / text / table / formula
    text: str = ""
    level: int = 0
    rows: list[list[str]] = field(default_factory=list)
    merges: list[tuple[int, int, int, int]] = field(default_factory=list)


@dataclass
class SourceDoc:
    name: str  # 显示用的源文件名（如发票1.png）
    markdown: str


@dataclass
class ExactPage:
    """精确复制模式的一页（对应一张源图片）的几何重建结果。"""

    name: str                                    # 显示名（如 运单1.png）
    blocks: list = field(default_factory=list)   # table_rebuild 的 blocks
    warnings: list[str] = field(default_factory=list)
    flags: list = field(default_factory=list)    # 交叉校验待复核项（结构化，写入 Excel 黄底+清单）
    source_name: str = ""                        # 原始文件名（PDF 页仍指向源 PDF）
    source_page: int = 0                         # 源文件页码，图片为 1
    image_path: str = ""                         # 源图片路径（PDF 为渲染页），原图定位用
    to_original: list | None = None              # 3x3 矩阵：表格坐标→原图坐标；None 表示同一坐标系
    ocr_boxes: list = field(default_factory=list)  # OCR 文本框（与表格同坐标系，用于原图高亮）


@dataclass
class ConvertResult:
    sheets: int = 0
    tables: int = 0
    warnings: list[str] = field(default_factory=list)
    flagged: int = 0
    waybills: int = 0                            # 单号候选条目数（同一单号多次出现会重复计数）
    unique_waybills: int = 0
    duplicate_waybills: int = 0
    suspicious_waybills: int = 0
    confirmed_waybills: int = 0                  # 格式有效且交叉校验无冲突
    pending_waybills: int = 0                    # 待核对：格式有效但有冲突/未校验/多表格
    format_invalid_waybills: int = 0             # 格式异常或非 YT
    index_issues: int = 0
    # 单号索引明细（应用内搜索用；不落地到 Excel 之外的文件）
    waybill_entries: list = field(default_factory=list)
    index_issue_items: list = field(default_factory=list)


# ---------------------------------------------------------------- 解析


class _TableParser(HTMLParser):
    """解析 MinerU 输出的 HTML 表格，保留 colspan / rowspan 信息。"""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._cells: list[tuple[int, int, int, str]] = []  # (行, rowspan, colspan, 文本)
        self._row_index = -1
        self._buf: list[str] | None = None
        self._rowspan = 1
        self._colspan = 1

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "tr":
            self._row_index += 1
        elif tag in ("td", "th") and self._row_index >= 0:
            attr = {k.lower(): (v or "") for k, v in attrs}

            def _int(name: str) -> int:
                try:
                    return max(1, int(attr.get(name, "1")))
                except (TypeError, ValueError):
                    return 1

            self._rowspan = _int("rowspan")
            self._colspan = _int("colspan")
            self._buf = []
        elif tag == "br" and self._buf is not None:
            self._buf.append("\n")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in ("td", "th") and self._buf is not None:
            text = _clean_text("".join(self._buf))
            self._cells.append((self._row_index, self._rowspan, self._colspan, text))
            self._buf = None

    def handle_data(self, data: str) -> None:
        if self._buf is not None:
            self._buf.append(data)

    def grid(self) -> tuple[list[list[str]], list[tuple[int, int, int, int]]]:
        n_rows = self._row_index + 1
        occupied: dict[tuple[int, int], bool] = {}
        placed: list[tuple[int, int, int, int, str]] = []
        n_cols = 0
        for row, rowspan, colspan, text in self._cells:
            col = 0
            while occupied.get((row, col)):
                col += 1
            for dr in range(rowspan):
                for dc in range(colspan):
                    occupied[(row + dr, col + dc)] = True
            n_cols = max(n_cols, col + colspan)
            placed.append((row, col, rowspan, colspan, text))

        n_rows = max(n_rows, max((r + rs for r, _c, rs, _cs, _t in placed), default=0))
        grid = [["" for _ in range(n_cols)] for _ in range(n_rows)]
        merges: list[tuple[int, int, int, int]] = []
        for row, col, rowspan, colspan, text in placed:
            grid[row][col] = text
            if rowspan > 1 or colspan > 1:
                merges.append((row, col, row + rowspan - 1, col + colspan - 1))
        return grid, merges


_INLINE_RULES: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"!\[[^\]]*\]\([^)]*\)"), ""),
    (re.compile(r"\[([^\]]*)\]\([^)]*\)"), r"\1"),
    (re.compile(r"\*\*([^*]+)\*\*"), r"\1"),
    (re.compile(r"(?<!\*)\*([^*\n]+)\*(?!\*)"), r"\1"),
    (re.compile(r"`([^`\n]+)`"), r"\1"),
]


def _clean_text(raw: str) -> str:
    """去掉行内 Markdown 标记与残留 HTML 标签。"""
    text = raw.replace("\u00a0", " ")
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"</?[a-zA-Z][^>]*>", "", text)
    text = html.unescape(text)
    for pattern, repl in _INLINE_RULES:
        text = pattern.sub(repl, text)
    lines = [ln.strip() for ln in text.split("\n")]
    return "\n".join(ln for ln in lines if ln).strip()


_HTML_TABLE_RE = re.compile(r"<table[^>]*>(.*?)</table>", re.S | re.I)
_PIPE_SEP_RE = re.compile(r"^\s*\|?[\s:|\-]+\|?\s*$")
_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_LIST_RE = re.compile(r"^\s*(?:[-*+]|\d+[.)、])\s+(.*)$")
_FORMULA_RE = re.compile(r"^\s*\$\$\s*$")


def _parse_pipe_table(lines: list[str], start: int) -> tuple[list[list[str]], int] | None:
    block: list[str] = []
    i = start
    while i < len(lines) and "|" in lines[i]:
        block.append(lines[i])
        i += 1
    rows: list[list[str]] = []
    for line in block:
        if _PIPE_SEP_RE.match(line) and "-" in line:
            continue
        body = line.strip()
        if body.startswith("|"):
            body = body[1:]
        if body.endswith("|"):
            body = body[:-1]
        cells = [_clean_text(cell) for cell in body.split("|")]
        if any(cells):
            rows.append(cells)
    if len(rows) < 2 or max(len(r) for r in rows) < 2:
        return None
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    return rows, i


def parse_markdown(markdown: str) -> list[Block]:
    """把 Markdown 拆成有序的内容块。"""
    blocks: list[Block] = []
    markdown = markdown.lstrip("\ufeff")
    pos = 0
    for match in _HTML_TABLE_RE.finditer(markdown):
        if match.start() > pos:
            blocks.extend(_parse_text_segment(markdown[pos : match.start()]))
        parser = _TableParser()
        parser.feed(match.group(1))
        parser.close()
        grid, merges = parser.grid()
        if grid and any(any(cell for cell in row) for row in grid):
            blocks.append(Block(kind="table", rows=grid, merges=merges))
        pos = match.end()
    if pos < len(markdown):
        blocks.extend(_parse_text_segment(markdown[pos:]))
    return blocks


def _parse_text_segment(segment: str) -> list[Block]:
    blocks: list[Block] = []
    lines = segment.split("\n")
    paragraph: list[str] = []
    i = 0

    def flush() -> None:
        if paragraph:
            text = _clean_text("\n".join(paragraph))
            if text:
                blocks.append(Block(kind="text", text=text))
            paragraph.clear()

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            flush()
            i += 1
            continue

        if _FORMULA_RE.match(line):
            flush()
            i += 1
            buf: list[str] = []
            while i < len(lines) and not _FORMULA_RE.match(lines[i]):
                buf.append(lines[i])
                i += 1
            i += 1
            formula = "\n".join(buf).strip()
            if formula:
                blocks.append(Block(kind="formula", text=formula))
            continue

        heading = _HEADING_RE.match(stripped)
        if heading:
            flush()
            blocks.append(
                Block(kind="title", level=len(heading.group(1)), text=_clean_text(heading.group(2)))
            )
            i += 1
            continue

        if "|" in stripped and stripped.startswith("|"):
            flush()
            parsed = _parse_pipe_table(lines, i)
            if parsed:
                rows, i = parsed
                blocks.append(Block(kind="table", rows=rows))
                continue

        item = _LIST_RE.match(line)
        if item:
            flush()
            blocks.append(Block(kind="text", text="• " + _clean_text(item.group(1))))
            i += 1
            continue

        paragraph.append(line)
        i += 1

    flush()
    return blocks


# ---------------------------------------------------------------- 写 Excel

HEADING_SIZES = {1: 16, 2: 14, 3: 12}
BODY_FONT = "微软雅黑"
HEADER_FILL = PatternFill("solid", fgColor="DCE6F1")
TITLE_FILL = PatternFill("solid", fgColor="F2F2F2")
FLAG_FILL = PatternFill("solid", fgColor="FFF2CC")   # 待复核单元格黄底
FLAG_FONT = Font(name=BODY_FONT, size=10.5, color="9C6500")
_THIN = Side(style="thin", color="BFBFBF")
TABLE_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _visual_width(text: str) -> int:
    """按视觉宽度估算字符数（中文按 2 个字符计）。"""
    width = 0
    for ch in text:
        width += 2 if ord(ch) > 127 else 1
    return width


def _width_for(texts: Iterable[str], minimum: int = 10, maximum: int = 46) -> float:
    longest = max((_visual_width(t.split("\n")[0]) for t in texts), default=0)
    return float(min(max(minimum, longest + 4), maximum))


def _sheet_title(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", Path(name).stem).strip() or "Sheet"
    base = base[:28]
    title = base
    index = 2
    while title in used:
        title = f"{base}_{index}"
        index += 1
    used.add(title)
    return title


def _merge_text(ws, row: int, col_count: int, font: Font, fill=None, value: str = "") -> None:
    ws.cell(row=row, column=1, value=value)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_count)
    cell = ws.cell(row=row, column=1)
    cell.font = font
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    if fill is not None:
        cell.fill = fill


def _write_table(ws, row: int, block: Block) -> int:
    rows = block.rows
    for r_index, cells in enumerate(rows, start=row):
        for c_index, value in enumerate(cells, start=1):
            cell = ws.cell(row=r_index, column=c_index, value=value)
            cell.border = TABLE_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font = Font(name=BODY_FONT, size=10.5)
            if r_index == row:
                cell.font = Font(name=BODY_FONT, size=10.5, bold=True)
                cell.fill = HEADER_FILL
    for r0, c0, r1, c1 in block.merges:
        try:
            ws.merge_cells(
                start_row=row + r0, start_column=c0 + 1, end_row=row + r1, end_column=c1 + 1
            )
        except (ValueError, TypeError):
            pass
    return row + len(rows)


def _write_structure_sheet(ws, blocks: list[Block], widths: list[float]) -> int:
    col_count = len(widths)
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width

    row = 1
    tables = 0
    for block in blocks:
        if block.kind == "title":
            size = HEADING_SIZES.get(block.level, 11)
            _merge_text(
                ws,
                row,
                col_count,
                Font(name=BODY_FONT, size=size, bold=True),
                TITLE_FILL if block.level <= 2 else None,
                block.text,
            )
            ws.row_dimensions[row].height = 24 if block.level <= 2 else 20
            row += 2 if block.level == 1 else 1
        elif block.kind == "text":
            _merge_text(ws, row, col_count, Font(name=BODY_FONT, size=11), None, block.text)
            row += 1
        elif block.kind == "formula":
            _merge_text(
                ws, row, col_count, Font(name="Cambria Math", size=11, italic=True), None, block.text
            )
            row += 1
        elif block.kind == "table":
            row = _write_table(ws, row, block)
            tables += 1
            row += 1
    return tables


def _write_table_sheet(ws, docs: list[SourceDoc]) -> tuple[int, int]:
    """把全部表格纵向拼接写入同一工作表，返回 (表格数, 是否有内容)。"""
    ws.column_dimensions["A"].width = 12
    row = 1
    tables = 0
    for doc in docs:
        doc_index = 0
        last_text = ""
        for block in parse_markdown(doc.markdown):
            if block.kind in ("text", "title") and block.text.strip():
                last_text = block.text.strip().replace("\n", " ")
                continue
            if block.kind != "table":
                continue
            doc_index += 1
            tables += 1
            cols = len(block.rows[0]) if block.rows else 0
            caption = f"{doc.name} · 表格{doc_index}（{len(block.rows)}行 × {cols}列）"
            if last_text and len(last_text) <= 40:
                caption += "  " + last_text
            last_text = ""
            _merge_text(
                ws, row, 8, Font(name=BODY_FONT, size=10.5, bold=True, color="1F4E79"),
                PatternFill("solid", fgColor="F2F7FF"), caption,
            )
            row += 1
            row = _write_table(ws, row, block)
            row += 2
    return tables, tables > 0


def _write_fallback(ws, docs: list[SourceDoc]) -> None:
    """未识别到表格时，把原文写入，避免出现空文件。"""
    row = 1
    _merge_text(
        ws, row, 8,
        Font(name=BODY_FONT, size=11, bold=True, color="C00000"),
        None,
        "未识别到表格结构，以下为原文内容，可改用「还原文档结构」模式查看排版。",
    )
    row += 2
    for doc in docs:
        _merge_text(ws, row, 8, Font(name=BODY_FONT, size=11, bold=True), TITLE_FILL, doc.name)
        row += 1
        for block in parse_markdown(doc.markdown):
            if block.kind == "table":
                row = _write_table(ws, row, block)
            elif block.text:
                _merge_text(ws, row, 8, Font(name=BODY_FONT, size=11), None, block.text)
                row += 1
            row += 1
    for index in range(1, 9):
        ws.column_dimensions[get_column_letter(index)].width = 14


def convert(docs: list[SourceDoc], out_path: Path, mode: str = "structure") -> ConvertResult:
    """把若干 Markdown 文档转换成 Excel 工作簿。

    mode: "structure" 还原文档结构 / "table" 只提取规整表格
    """
    result = ConvertResult()
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()

    if mode == "table":
        ws = wb.create_sheet("表格数据")
        tables, has_content = _write_table_sheet(ws, docs)
        if not has_content:
            result.warnings.append("未识别到表格，已改为输出原文内容。")
            _write_fallback(ws, docs)
        result.tables = tables
        result.sheets = 1
    else:
        for doc in docs:
            blocks = parse_markdown(doc.markdown)
            if not blocks:
                result.warnings.append(f"{doc.name}: 未解析到有效内容，已跳过。")
                continue
            ws = wb.create_sheet(_sheet_title(doc.name, used))
            max_cols = max(
                (len(b.rows[0]) for b in blocks if b.kind == "table" and b.rows), default=10
            )
            max_cols = max(max_cols, 10)
            widths: list[float] = []
            for col in range(max_cols):
                widths.append(
                    _width_for(
                        [
                            b.rows[r][col]
                            for b in blocks
                            if b.kind == "table" and b.rows
                            for r in range(len(b.rows))
                            if col < len(b.rows[r])
                        ],
                        minimum=11,
                        maximum=40,
                    )
                )
            tables = _write_structure_sheet(ws, blocks, widths)
            result.tables += tables
            result.sheets += 1

    if not wb.sheetnames:
        raise ValueError("没有可写入的内容")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return result


# ---------------------------------------------------------------- 精确复制（OCR 几何重建直写）

EXACT_TEXT_FONT = Font(name=BODY_FONT, size=11)


def _page_source_label(page) -> str:
    """一页的来源显示名：图片用文件名，PDF 用「文件名 第N页」。"""
    source = str(getattr(page, "source_name", "") or getattr(page, "name", "") or "")
    page_no = int(getattr(page, "source_page", 0) or 0)
    if source.lower().endswith(".pdf") and page_no:
        return f"{source} 第{page_no}页"
    return source


def _header_signature(row) -> tuple[str, ...]:
    """表头签名：去空白后逐格比较，用于判断多张图是否同一套表头。"""
    return tuple(re.sub(r"\s+", "", str(c or "")) for c in (row or []))


def _block_has_content(block) -> bool:
    """块内是否有非空内容（TableBlock 看 rows，TextBlock 看 lines）。"""
    if getattr(block, "kind", "") == "table":
        rows = getattr(block, "rows", None) or []
        return bool(rows) and any(any(str(c).strip() for c in r) for r in rows)
    lines = getattr(block, "lines", None) or []
    return any(str(ln).strip() for ln in lines)


def _row_height(pitch: float, minimum: float = 15.0, maximum: float = 60.0) -> float:
    """原图像素行距 -> Excel 行高（磅）：1px(96dpi) = 0.75pt。"""
    return min(max(minimum, pitch * 0.75), maximum)


def _exact_row_heights(block) -> list[float]:
    """按 row_y 各行 y 包络推算每行行高（磅），尽量还原原图行距。"""
    rows = getattr(block, "rows", None) or []
    row_y = getattr(block, "row_y", None) or []
    n = len(rows)
    if not row_y or len(row_y) != n:
        return [20.0] * n
    centers = [(y1 + y2) / 2.0 for y1, y2 in row_y]
    heights = []
    for i in range(n):
        if i + 1 < n:
            pitch = centers[i + 1] - centers[i]
        elif i > 0:
            pitch = centers[i] - centers[i - 1]
        else:
            pitch = (row_y[0][1] - row_y[0][0]) * 1.5
        if pitch <= 0:
            pitch = (row_y[i][1] - row_y[i][0]) * 1.5
        heights.append(_row_height(pitch if pitch > 0 else 24.0))
    return heights


def _exact_widths(blocks: list, n_cols: int) -> list[float]:
    """按全部表格的各列内容估算列宽。"""
    widths = []
    for col in range(n_cols):
        texts = [
            str(block.rows[r][col])
            for block in blocks
            if getattr(block, "kind", "") == "table"
            for r in range(len(block.rows))
            if col < len(block.rows[r])
        ]
        widths.append(_width_for(texts, minimum=10, maximum=46))
    return widths


def _write_exact_table(ws, row: int, block, heights: list[float],
                       flag_cells=None) -> int:
    """表格网格直写：全单元格细边框，首行表头加粗 + 底纹，行高按 heights。

    flag_cells：{(相对行, 相对列)} 集合，命中单元格黄底高亮（待人工复核）。
    """
    rows = block.rows
    n_cols = max((len(r) for r in rows), default=0)
    flag_cells = flag_cells or set()
    for r_index, cells in enumerate(rows, start=row):
        rel_row = r_index - row
        for c_index in range(n_cols):
            value = str(cells[c_index]) if c_index < len(cells) else ""
            cell = ws.cell(row=r_index, column=c_index + 1, value=value)
            cell.border = TABLE_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font = Font(name=BODY_FONT, size=10.5)
            if r_index == row:
                cell.font = Font(name=BODY_FONT, size=10.5, bold=True)
                cell.fill = HEADER_FILL
            if (rel_row, c_index) in flag_cells:
                cell.fill = FLAG_FILL
                cell.font = FLAG_FONT
        if r_index - row < len(heights):
            ws.row_dimensions[r_index].height = heights[r_index - row]
    return row + len(rows)


def _write_exact_sheet(ws, page: ExactPage, table_infos: list | None = None) -> tuple[int, bool]:
    """把一个 ExactPage 的 blocks 按原文档顺序写入工作表。

    返回 (表格数, 是否有内容)。表格后空一行与后续内容分隔；文字块逐行合并、
    无边框，保持原图阅读顺序。table_infos 用于记录表格在 Excel 中的起始行，
    供单号索引生成单元格地址。
    """
    blocks = [b for b in page.blocks if _block_has_content(b)]
    if not blocks:
        return 0, False
    n_cols = max(
        (max(len(r) for r in b.rows) for b in blocks if getattr(b, "kind", "") == "table"),
        default=8,
    )
    n_cols = max(n_cols, 4)
    for index, width in enumerate(_exact_widths(blocks, n_cols), start=1):
        ws.column_dimensions[get_column_letter(index)].width = width

    row = 1
    tables = 0
    # 交叉校验的 row/col 相对本页第一个表格块的 rows，只对该块高亮
    flag_cells = {(int(f.get("row", -1)), int(f.get("col", -1)))
                  for f in getattr(page, "flags", None) or []}
    flag_used = False
    for block in blocks:
        if getattr(block, "kind", "") == "table":
            cells = None if flag_used else flag_cells
            start_row = row
            row = _write_exact_table(ws, row, block, _exact_row_heights(block),
                                     flag_cells=cells)
            if table_infos is not None:
                table_infos.append({
                    "block": block,
                    "start_row": start_row,
                    "table_index": tables,
                    "flags": (getattr(page, "flags", None) or []) if not flag_used else [],
                })
            flag_used = True
            tables += 1
            row += 1                      # 表后空一行分隔
        else:
            for line in block.lines:
                if not str(line).strip():
                    continue
                _merge_text(ws, row, n_cols, EXACT_TEXT_FONT, None, str(line))
                ws.row_dimensions[row].height = 20
                row += 1
    return tables, True


def _write_merged_table(ws, row: int, block, heights: list[float], source_label: str,
                        skip_header: bool, first_table: bool,
                        flag_cells=None) -> int:
    """把一张表写入「合并数据」工作表：A 列来源图片，B 列起为表格内容。

    skip_header=True 表示与首张表表头相同，跳过表头行直接续数据；
    first_table=True 时 A 列表头写「来源图片」，其余表头行写具体文件名。
    """
    rows = getattr(block, "rows", None) or []
    n_cols = max((len(r) for r in rows), default=0)
    flag_cells = flag_cells or set()
    out_row = row
    for rel, cells in enumerate(rows):
        if skip_header and rel == 0:
            continue
        r = out_row
        source_cell = ws.cell(
            row=r, column=1,
            value="来源图片" if (first_table and rel == 0) else source_label)
        source_cell.border = TABLE_BORDER
        source_cell.alignment = Alignment(vertical="center", wrap_text=True)
        source_cell.font = Font(name=BODY_FONT, size=10.5)
        if rel == 0:
            source_cell.font = Font(name=BODY_FONT, size=10.5, bold=True)
            source_cell.fill = HEADER_FILL
        for c_index in range(n_cols):
            value = str(cells[c_index]) if c_index < len(cells) else ""
            cell = ws.cell(row=r, column=c_index + 2, value=value)
            cell.border = TABLE_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font = Font(name=BODY_FONT, size=10.5)
            if rel == 0:
                cell.font = Font(name=BODY_FONT, size=10.5, bold=True)
                cell.fill = HEADER_FILL
            if (rel, c_index) in flag_cells:
                cell.fill = FLAG_FILL
                cell.font = FLAG_FONT
        if rel < len(heights):
            ws.row_dimensions[r].height = heights[rel]
        out_row += 1
    return out_row


def _write_merged_page(ws, state: dict, page, table_infos: list | None,
                       n_cols: int) -> int:
    """把一页的表格与文字按原顺序写入「合并数据」工作表。

    state 跨页复用：{"row": 下一条可写行, "header": 首张表的表头签名}。
    结构相同的表格直接续行；结构不同的表格空一行并重写表头，避免串列。
    """
    blocks = [b for b in page.blocks if _block_has_content(b)]
    source = _page_source_label(page)
    flag_cells = {(int(f.get("row", -1)), int(f.get("col", -1)))
                  for f in getattr(page, "flags", None) or []}
    flag_used = False
    tables = 0
    for block in blocks:
        if getattr(block, "kind", "") == "table":
            rows = getattr(block, "rows", None) or []
            header = _header_signature(rows[0]) if rows else ()
            first_table = state["header"] is None
            if first_table:
                state["header"] = header
                skip_header = False
            else:
                skip_header = header == state["header"]
                if not skip_header:
                    state["row"] += 1
            start_row = state["row"] - (1 if skip_header else 0)
            cells = None if flag_used else flag_cells
            state["row"] = _write_merged_table(
                ws, state["row"], block, _exact_row_heights(block), source,
                skip_header=skip_header, first_table=first_table, flag_cells=cells)
            if table_infos is not None:
                table_infos.append({
                    "block": block,
                    "start_row": start_row,
                    "table_index": tables,
                    "col_offset": 1,
                    "flags": (getattr(page, "flags", None) or []) if not flag_used else [],
                })
            flag_used = True
            tables += 1
        else:
            for line in block.lines:
                if not str(line).strip():
                    continue
                source_cell = ws.cell(row=state["row"], column=1, value=source)
                source_cell.border = TABLE_BORDER
                source_cell.font = Font(name=BODY_FONT, size=10.5)
                source_cell.alignment = Alignment(vertical="center", wrap_text=True)
                ws.cell(row=state["row"], column=2, value=str(line))
                ws.merge_cells(start_row=state["row"], start_column=2,
                               end_row=state["row"], end_column=n_cols + 1)
                text_cell = ws.cell(row=state["row"], column=2)
                text_cell.font = EXACT_TEXT_FONT
                text_cell.alignment = Alignment(vertical="center", wrap_text=True)
                ws.row_dimensions[state["row"]].height = 20
                state["row"] += 1
    return tables


# ------------------------------------------------------------ 待复核清单分流
#
# 优先级依据 45 张真实图 1472 条台账的抽样实证（非主观划分）：
#   高 = 粘连拆分 / 已自动改写 / 放大差异中的条码·商品列粘连
#        —— 粘连拆分、已改写抽样全部有价值；条码列粘连直接影响商品对账；
#   中 = 文本不一致中两方给出不同具体值、双单号粘连、其他编码粘连、
#        实收列共识不足、单号前置噪声已剥离 —— 需人工裁定；
#   低 = 置信度不足 / 放大无结果 / 读邻格取框失败 / 单号·序号·状态词列粘连
#        —— 抽样基本全是噪音；单号列粘连由 YT 正则兜底提取，
#           不影响单号索引与纯单号清单。
# 低优先条目不删除：全部保留在「待复核-低优先」工作表，高低中合计与分流前一致。
PRIO_HIGH_FILL = PatternFill("solid", fgColor="C00000")
PRIO_HIGH_FONT = Font(name=BODY_FONT, size=10.5, bold=True, color="FFFFFF")
PRIO_MID_FILL = PatternFill("solid", fgColor="FCE4D6")   # 浅橙，区别于数据页黄底
PRIO_MID_FONT = Font(name=BODY_FONT, size=10.5, color="833C00")
FLAG_NOTE_FONT = Font(name=BODY_FONT, size=10, color="808080", italic=True)

_FLAG_YT_RE = re.compile(r"YT\d{6,}", re.I)
_FLAG_STATUS_RE = re.compile(r"签收|拒收|完好|破损")


def _flag_stick_evidence(rebuilt: str, ocr: str) -> bool:
    """重建值相对一致 OCR 值是否呈粘连证据（两段内容拼进一格）。"""
    rb, oc = rebuilt.strip(), ocr.strip()
    if not rb or not oc or rb == oc or len(rb) < len(oc) + 2:
        return False
    if oc in rb:
        return True
    same = 0
    for a, b in zip(rb, oc):
        if a != b:
            break
        same += 1
    return same >= 6


def _classify_flag(fl: dict) -> tuple[str, str]:
    """待复核条目分流：返回 (问题类型, 优先级)，优先级 ∈ {高, 中, 低}。"""
    reason = str(fl.get("reason", ""))
    rb = str(fl.get("rebuilt", ""))
    oc = str(fl.get("ocr", ""))
    z = str(fl.get("zoom", ""))
    if reason.startswith("已按放大复识自动改写"):
        return "已自动改写", "高"
    if "粘连" in reason:
        return "粘连拆分", "高"
    if reason.startswith("放大复识结果与原值差异过大"):
        if oc and oc == z and _flag_stick_evidence(rb, oc):
            if len(_FLAG_YT_RE.findall(rb)) >= 2:
                return "放大差异-双单号粘连", "中"
            if _FLAG_STATUS_RE.search(rb):
                return "放大差异-单号状态词粘连", "低"
            if oc.isdigit() and len(oc) >= 10:
                return "放大差异-条码/商品列粘连", "高"
            if _FLAG_YT_RE.search(rb) or _FLAG_YT_RE.search(oc):
                return "放大差异-单号列粘连", "低"
            if re.fullmatch(r"\d{1,6}", oc):
                return "放大差异-序号/短数字粘连", "低"
            return "放大差异-其他编码粘连", "中"
        return "放大差异-读邻格/取框失败", "低"
    if reason.startswith("文本字段两方不一致（整图 OCR 置信度仅"):
        return "置信度不足", "低"
    if reason.startswith("放大复识置信度"):
        return "置信度不足", "低"
    if reason.startswith("文本字段两方不一致"):
        if oc.strip() and z.strip() and oc != z:
            return "文本不一致-有具体差异", "中"
        return "文本不一致-无具体差异", "低"
    if reason.startswith("放大复识无结果"):
        return "放大复识无结果", "低"
    if reason.startswith("单号前置噪声已剥离"):
        return "单号前置噪声已剥离", "中"
    if reason.startswith("实收列"):
        return "实收列共识不足", "中"
    return "其他", "低"


def _triage_flags(entries: list[tuple[str, dict]]):
    """分流 + 排序：返回 (主清单, 低优先清单)。

    主清单按 高→中、文件、行、列 排序；低优先按问题类型聚块再按文件/行/列，
    方便用户整块跳过。两个清单合计条数 == 分流前条数。
    """
    tagged = [(name, fl, *_classify_flag(fl)) for name, fl in entries]
    hi_mid = [e for e in tagged if e[3] != "低"]
    low = [e for e in tagged if e[3] == "低"]
    hi_mid.sort(key=lambda e: (0 if e[3] == "高" else 1, e[0],
                               int(e[1].get("row", 0)), int(e[1].get("col", 0))))
    low.sort(key=lambda e: (e[2], e[0],
                            int(e[1].get("row", 0)), int(e[1].get("col", 0))))
    return hi_mid, low


def _flag_sheet_header(ws) -> None:
    headers = ("优先级", "问题类型", "文件", "行", "列",
               "重建值", "整图OCR", "放大OCR", "原因")
    widths = (7, 22, 30, 6, 6, 26, 26, 26, 44)
    for index, (title, width) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=index, value=title)
        cell.font = Font(name=BODY_FONT, size=10.5, bold=True)
        cell.fill = HEADER_FILL
        cell.border = TABLE_BORDER
        cell.alignment = Alignment(vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(index)].width = width


def _write_flag_rows(ws, start: int, entries: list[tuple[str, dict, str, str]]) -> int:
    """写待复核数据行，优先级单元格按 高红/中黄 着色。返回下一可用行。"""
    for r, (name, fl, ptype, prio) in enumerate(entries, start=start):
        # 引擎拆分阶段的粘连 flag 用 whole/local 键（原始粘连值/剥离段），
        # 回退到这两列，避免清单里出现空白的整图/放大列。
        values = (
            prio, ptype, name,
            int(fl.get("row", 0)) + 1,
            int(fl.get("col", 0)) + 1,
            str(fl.get("rebuilt", "") or "（无）"),
            str(fl.get("ocr", "") or fl.get("whole", "") or "（无）"),
            str(fl.get("zoom", "") or fl.get("local", "") or "（无）"),
            str(fl.get("reason", "")),
        )
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = TABLE_BORDER
            cell.font = Font(name=BODY_FONT, size=10.5)
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center" if c in (1, 4, 5) else "left",
                wrap_text=True,
            )
        pr = ws.cell(row=r, column=1)
        if prio == "高":
            pr.fill = PRIO_HIGH_FILL
            pr.font = PRIO_HIGH_FONT
        elif prio == "中":
            pr.fill = PRIO_MID_FILL
            pr.font = PRIO_MID_FONT
    return start + len(entries)


def _write_flag_sheet(ws, entries: list[tuple[str, dict, str, str]],
                      low_count: int = 0) -> None:
    """待复核清单 sheet（高+中）：优先级置顶，末尾附合计行便于核对总数。

    0 条时仍写出表头与「本次无待复核项」，让用户知道检查过，不做静默省略。
    """
    _flag_sheet_header(ws)
    high = sum(1 for e in entries if e[3] == "高")
    mid = len(entries) - high
    total = len(entries) + low_count
    if entries:
        row = _write_flag_rows(ws, 2, entries)
        note_text = (f"合计：高 {high} 条 + 中 {mid} 条 = {len(entries)} 条；"
                     f"另有低优先 {low_count} 条见「待复核-低优先」工作表；"
                     f"全部待复核共 {total} 条")
    else:
        row = 2
        note_text = (f"本次无待复核项：全部数据页未发现需要人工复核的单元格；"
                     f"高 0 条 + 中 0 条 = 0 条；另有低优先 {low_count} 条见"
                     f"「待复核-低优先」工作表；全部待复核共 {total} 条")
    note = ws.cell(row=row, column=1, value=note_text)
    note.font = FLAG_NOTE_FONT
    ws.freeze_panes = "A2"


def _write_flag_low_sheet(ws, entries: list[tuple[str, dict, str, str]],
                          high_count: int, mid_count: int) -> None:
    """待复核-低优先 sheet：全部保留、仅供参考，表头下放噪音说明行。

    0 条时写明「本次无低优先待复核项」，同样保留说明行结构。
    """
    _flag_sheet_header(ws)
    total = high_count + mid_count + len(entries)
    if entries:
        note_text = (f"本表 {len(entries)} 条多为噪音（置信度不足 / 放大复识无结果 / "
                     f"读邻格取框失败 / 单号·序号·状态词列粘连——单号已被 YT 格式规则"
                     f"兜底提取，不影响单号索引与纯单号清单），仅供参考，无需逐条复核。"
                     f"全部待复核共 {total} 条 = 高 {high_count} + 中 {mid_count} "
                     f"+ 低 {len(entries)}")
    else:
        note_text = (f"本次无低优先待复核项：全部数据页未发现低优先级别的待复核条目。"
                     f"全部待复核共 {total} 条 = 高 {high_count} + 中 {mid_count} + 低 0")
    note = ws.cell(row=2, column=1, value=note_text)
    note.font = FLAG_NOTE_FONT
    note.alignment = Alignment(vertical="center", wrap_text=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=9)
    ws.row_dimensions[2].height = 30
    _write_flag_rows(ws, 3, entries)
    ws.freeze_panes = "A3"


_WAYBILL_PRIMARY_HEADERS = (
    "运单号", "运单编号", "运单号码", "快递单号", "快递单号码", "物流单号", "物流单号码",
)
_WAYBILL_SECONDARY_HEADERS = (
    "退货单号", "退货单号码", "退仓单号", "配送单号", "发货单号",
)
# OCR 可能把“运单号”读成“包责号”，或模板本身用“包裹号”承载运单号。
# 这两个只作为弱别名，只有在没有运单号/快递单号等明确表头时才使用。
_WAYBILL_FALLBACK_HEADERS = ("包裹号", "包责号")
_WAYBILL_HEADER_EXCLUDES = (
    "序号", "商品号", "货号", "条码", "订单号", "批次号", "库位号",
)
# 常见运单号：YT + 13 位数字；R02TYT... 前缀也保留后再校验。
# 不满足此格式的候选一律不得标为“已确认”。
_STRICT_WAYBILL_RE = re.compile(r"^(?:R\d{2}T)?YT\d{13}$", re.I)


def _norm_header(value) -> str:
    """表头文本归一化：去空白和常见装饰字符，便于匹配运单号列。"""
    return re.sub(r"[\s\r\n\t:：*（）()【】\[\]]+", "", str(value or "")).lower()


def _waybill_header_score(value) -> int:
    text = _norm_header(value)
    if not text:
        return 0
    if any(word in text for word in _WAYBILL_HEADER_EXCLUDES):
        return 0
    if any(word in text for word in _WAYBILL_PRIMARY_HEADERS):
        return 120
    if any(word in text for word in _WAYBILL_SECONDARY_HEADERS):
        return 100
    if text == "单号":
        return 70
    if any(word in text for word in _WAYBILL_FALLBACK_HEADERS):
        return 55
    return 0


def _find_waybill_columns(block) -> list[tuple[int, int, str]]:
    """找出一张表里**所有**运单号列，返回 [(表头相对行, 列, 表头文本), ...]。

    双列版式（同一页水平并排两组「序号 + 运单号」，如退货清单的 1-50 /
    51-100）以前只取命中最多的一列，另一组单号整组进不了索引。这里把表头
    命中列与「按内容识别列」全部返回，按列序去重；表头读坏时仍靠内容兜底。
    """
    rows = getattr(block, "rows", None) or []
    if not rows:
        return []
    found: list[tuple[int, int, str]] = []
    seen: set[int] = set()

    best = None
    for row_index, row in enumerate(rows[:5]):
        for col_index, value in enumerate(row):
            score = _waybill_header_score(value)
            if score and (best is None or score > best[0]):
                best = (score, row_index, col_index, str(value))
    if best is not None:
        found.append((best[1], best[2], best[3]))
        seen.add(best[2])

    # 表头被 OCR 读坏（如"包责号"）、或第二组列根本没有表头时的兜底：
    # 找到多行都含 YT 开头运单号格式的列，按内容识别，不猜其他数字列。
    yt_pattern = re.compile(r"(?:R\d{2}T)?YT\d{8,}", re.I)
    max_cols = max((len(row) for row in rows[:30]), default=0)
    for col_index in range(max_cols):
        if col_index in seen:
            continue
        hits = 0
        for row in rows[:30]:
            if col_index < len(row) and yt_pattern.search(str(row[col_index] or "")):
                hits += 1
        if hits >= 2:
            found.append((-1, col_index, "按YT单号格式自动识别"))
            seen.add(col_index)
    found.sort(key=lambda item: item[1])
    return found


def _find_waybill_column(block) -> tuple[int, int, str] | None:
    """兼容旧调用：返回第一个运单号列。"""
    columns = _find_waybill_columns(block)
    return columns[0] if columns else None


def _looks_like_repeated_header(row: list, skip_col: int = -1) -> bool:
    """数据区里重复出现的表头行需要跳过，避免进入单号索引。"""
    score = 0
    for index, value in enumerate(row):
        if index == skip_col:
            continue
        text = _norm_header(value)
        if text and text in {
            "日期", "承运商", "运单号", "商品名称", "实收数", "短少数",
            "序号", "包裹号", "签收状态", "商品条码", "数量", "备注",
        }:
            score += 1
    return score >= 2


def _waybill_tokens(value) -> list[str]:
    """从单元格提取单号；支持一格里用换行/逗号/斜杠分隔多个单号。

    兼容两类真实版式：
      - ``45YT...``：序号/行号与运单号粘连，只取 YT 及后续号码；
      - ``R02TYT...``：R02T 是运单号前缀的一部分，必须保留。
    """
    raw = str(value or "").strip()
    if not raw:
        return []
    tokens: list[str] = []
    for part in re.split(r"[\r\n,，;；、|/]+", raw):
        compact = re.sub(r"\s+", "", part)
        if not compact:
            continue
        # 先处理 YT 开头运单号，避免 R02TYT 被截成 TYT
        found = re.findall(r"(?:R\d{2}T)?YT\d{8,}", compact, flags=re.I)
        if not found:
            found = re.findall(r"[A-Za-z]{1,8}\d{6,}[A-Za-z]?", compact)
        if not found:
            found = re.findall(r"\d{8,}[A-Za-z]?", compact)
        tokens.extend(found)
    # 同一格内去掉完全重复值，保留顺序
    return list(dict.fromkeys(tokens))


def _waybill_problem(token: str) -> str:
    digits = re.sub(r"\D", "", token)
    if len(digits) < 8:
        return "单号位数不足"
    if not re.fullmatch(r"[A-Za-z0-9-]+", token):
        return "单号包含异常字符"
    return ""


_MERGED_TAIL_RE = re.compile(r"^((?:R\d{2}T)?YT\d{13})(\d{1,4})$", re.I)


def _merged_tail_hint(token: str) -> str:
    """单号格疑似把相邻列数字粘进来时，给出保守提示（绝不自动截断）。

    典型：``YT123456789012380`` = 正确单号 ``YT1234567890123`` + 相邻列 ``80``。
    截断有拆错风险（R02TYT 前缀长度不同、OCR 也可能把 13 位读成 14 位），
    因此这里只生成"可能应为"的提示，由人工对照原图判断。
    """
    m = _MERGED_TAIL_RE.match(token.strip())
    if not m:
        return ""
    return f"疑似与相邻列数字粘连，若原图确认可能应为 {m.group(1)}（系统不会自动截断）"


def _waybill_status(token: str) -> tuple[str, str]:
    """严格校验候选单号。

    返回 (确认状态, 原因)。只有满足 YT + 13 位数字（或 R02TYT + 13 位）
    且后续没有交叉校验冲突的候选，才允许出现在“已确认”状态。
    格式不满足时绝不猜测补位，统一进入待核对/格式异常。
    """
    if _STRICT_WAYBILL_RE.fullmatch(token):
        return "待核对", ""
    if "YT" in token.upper():
        return "格式异常", "不符合 YT+13 位数字格式，禁止作为确认单号"
    return "非YT", "不是 YT 运单号，未纳入确认索引"


def _map_points(points: list[tuple[float, float]], matrix: list | None):
    """把表格坐标点映射到原图坐标；matrix 为空表示同一坐标系。"""
    if not matrix:
        return points
    m = matrix
    out = []
    for x, y in points:
        w = m[2][0] * x + m[2][1] * y + m[2][2]
        if abs(w) < 1e-9:
            w = 1.0
        out.append((
            (m[0][0] * x + m[0][1] * y + m[0][2]) / w,
            (m[1][0] * x + m[1][1] * y + m[1][2]) / w,
        ))
    return out


def _norm_box_text(text) -> str:
    return re.sub(r"[\s\-—_]+", "", str(text or "")).upper()


def _token_image_box(page: ExactPage, token: str) -> list | None:
    """按单号文字在 OCR 文本框里的位置，给出原图高亮框。

    表格单元格的边界受网格线漏检/误检影响（实测运单号列被一条假竖线切成
    两格，格子框会偏到号码后半段），而 OCR 文本框本身就是文字所在位置，
    优先用它；找不到时才回退到单元格框。
    """
    needle = _norm_box_text(token)
    if not needle:
        return None
    best: list | None = None
    best_area: float | None = None
    for b in getattr(page, "ocr_boxes", None) or []:
        try:
            if needle not in _norm_box_text(b.get("text", "")):
                continue
            x1, y1, x2, y2 = [float(v) for v in b["bbox"]]
        except (KeyError, TypeError, ValueError):
            continue
        mapped = _map_points(
            [(x1, y1), (x2, y1), (x2, y2), (x1, y2)],
            getattr(page, "to_original", None),
        )
        xs = [p[0] for p in mapped]
        ys = [p[1] for p in mapped]
        box = [int(round(min(xs))), int(round(min(ys))),
               int(round(max(xs))), int(round(max(ys)))]
        area = max(1, (box[2] - box[0]) * (box[3] - box[1]))
        if best is None or best_area is None or area < best_area:
            best, best_area = box, area
    return best


def _cell_image_box(page: ExactPage, block, rel_row: int, col: int) -> list | None:
    """单元格在原图中的像素框 [x1,y1,x2,y2]，用于搜索时在原图上高亮。"""
    try:
        if not getattr(block, "col_x", None) or not getattr(block, "row_y", None):
            return None
        if rel_row >= len(block.row_y) or col >= len(block.col_x):
            return None
        cx1, cx2 = block.col_x[col]
        ry1, ry2 = block.row_y[rel_row]
        mapped = _map_points(
            [(cx1, ry1), (cx2, ry1), (cx2, ry2), (cx1, ry2)],
            getattr(page, "to_original", None),
        )
        xs = [p[0] for p in mapped]
        ys = [p[1] for p in mapped]
        return [int(round(min(xs))), int(round(min(ys))),
                int(round(max(xs))), int(round(max(ys)))]
    except Exception:  # noqa: BLE001 - 高亮框失败不影响单号索引
        return None


def _box_from_warp(x1: float, y1: float, x2: float, y2: float,
                   to_original) -> list[int]:
    mapped = _map_points([(x1, y1), (x2, y1), (x2, y2), (x1, y2)], to_original)
    xs = [p[0] for p in mapped]
    ys = [p[1] for p in mapped]
    return [int(round(min(xs))), int(round(min(ys))),
            int(round(max(xs))), int(round(max(ys)))]


def _common_prefix_len(a: str, b: str) -> int:
    n = 0
    for ca, cb in zip(a, b):
        if ca != cb:
            break
        n += 1
    return n


def _text_match_score(target: str, candidate: str) -> int:
    """待复核单元格与 OCR 文本的相似度；0 表示不可信。

    待复核项本身就是“读错过”的格子，不能要求 OCR 文本与单元格完全一致：
    长编码允许少量数字误读，用公共前缀判定；文本类只接受完全包含。
    """
    if not target or not candidate:
        return 0
    if target == candidate:
        return 100 + len(target)
    if target in candidate or candidate in target:
        short = min(len(target), len(candidate))
        if short >= 4:
            return 60 + short
    if len(target) < 6 or len(candidate) < 6:
        return 0
    prefix = _common_prefix_len(target, candidate)
    if abs(len(target) - len(candidate)) <= 3 and prefix >= max(6, int(len(target) * 0.6)):
        return prefix
    return 0


def _aligned_row_y(block) -> list:
    """row_y 与数据行对齐：表头重定位只裁行首，几何对齐取尾部切片（同 xcheck）。"""
    row_y = list(getattr(block, "row_y", None) or [])
    rows = getattr(block, "rows", None) or []
    if rows and len(row_y) > len(rows):
        row_y = row_y[-len(rows):]
    return row_y


def flag_image_box(page: ExactPage, flag: dict) -> list | None:
    """待复核项在原图中的单元格框；定位不可信时返回 None。

    优先用单元格文字定位（长编码允许少量误读），其次用同一行其他列的
    可信文字锚定行高，最后才回退对齐后的 row_y。row_y 在表头重定位后可能
    比数据行多出一段（实测 52 vs 41），直接索引会把框打到标题区。
    """
    try:
        row = int(flag.get("row", -1))
        col = int(flag.get("col", -1))
    except (TypeError, ValueError):
        return None
    if row < 0 or col < 0:
        return None
    tables = [b for b in (getattr(page, "blocks", None) or [])
              if getattr(b, "kind", "") == "table"]
    candidates = []
    for block in tables:
        rows = getattr(block, "rows", None) or []
        if row >= len(rows) or col >= len(rows[row]):
            continue
        col_x = getattr(block, "col_x", None) or []
        if col >= len(col_x):
            continue
        cell = str(rows[row][col] or "")
        score = 0
        for key in ("whole", "rebuilt", "local"):
            value = str(flag.get(key) or "")
            if not value or not cell:
                continue
            if value == cell:
                score = max(score, 3)
            elif value in cell or cell in value:
                score = max(score, 2)
        candidates.append((score, block, cell))
    if not candidates:
        return None
    candidates.sort(key=lambda item: -item[0])
    block, cell = candidates[0][1], candidates[0][2]
    col_x = block.col_x
    cx1, cx2 = col_x[col]

    all_boxes = []
    for box in getattr(page, "ocr_boxes", None) or []:
        bbox = box.get("bbox")
        if not bbox or len(bbox) < 4:
            continue
        try:
            x1, y1, x2, y2 = (float(bbox[0]), float(bbox[1]),
                              float(bbox[2]), float(bbox[3]))
        except (TypeError, ValueError):
            continue
        all_boxes.append((_norm_box_text(box.get("text", "")),
                          x1, y1, x2, y2, (x1 + x2) / 2.0, (y1 + y2) / 2.0))

    def in_column(item, index: int) -> bool:
        if index >= len(col_x):
            return False
        a, b = col_x[index]
        width = max(1.0, float(b) - float(a))
        return float(a) - width * 0.35 <= item[5] <= float(b) + width * 0.35

    target = _norm_box_text(cell)
    scored = []
    if target:
        for item in all_boxes:
            if not in_column(item, col):
                continue
            score = _text_match_score(target, item[0])
            if score:
                scored.append((score, item))
    if scored:
        scored.sort(key=lambda pair: -pair[0])
        top_score, top = scored[0]
        same = sum(1 for score, _item in scored if score == top_score)
        is_code = len(target) >= 8 and bool(re.fullmatch(r"[A-Z0-9]+", target))
        # 唯一精确匹配，或长编码只有一处近似匹配（编码本身唯一）时才直接采信
        if (top_score >= 60 and same == 1) or (is_code and top_score >= 8 and same == 1):
            pad_x = (top[3] - top[1]) * 0.20 + 5.0
            pad_y = (top[4] - top[2]) * 0.25 + 2.0
            return _box_from_warp(top[1] - pad_x, top[2] - pad_y,
                                  top[3] + pad_x, top[4] + pad_y,
                                  getattr(page, "to_original", None))

    # 同行锚点：别的列读出可信文字时，用它的 y 中心确定本行位置
    row_values = block.rows[row] if row < len(block.rows) else []
    anchors = []
    for index, value in enumerate(row_values):
        if index == col:
            continue
        other = _norm_box_text(value)
        if len(other) < 6:
            continue
        best = (0, None)
        for item in all_boxes:
            if not in_column(item, index):
                continue
            score = _text_match_score(other, item[0])
            if score > best[0]:
                best = (score, item)
            if score >= 100:  # 完全一致，无需再找
                break
        if best[0] >= 60:
            anchors.append(best[1][6])
    aligned = _aligned_row_y(block)
    if anchors:
        height = 0.0
        if row < len(aligned):
            height = float(aligned[row][1]) - float(aligned[row][0])
        if height <= 0:
            heights = sorted(float(b) - float(a) for a, b in aligned)
            height = heights[len(heights) // 2] if heights else 38.0
        anchors.sort()
        center = anchors[len(anchors) // 2]
        return _box_from_warp(cx1, center - height / 2.0,
                              cx2, center + height / 2.0,
                              getattr(page, "to_original", None))

    if row < len(aligned):
        ry1, ry2 = aligned[row]
        return _box_from_warp(cx1, ry1, cx2, ry2,
                              getattr(page, "to_original", None))
    return None


def _row_summary(row: list, skip_col: int, limit: int = 90) -> str:
    parts = []
    for index, value in enumerate(row):
        if index == skip_col:
            continue
        text = re.sub(r"\s+", " ", str(value or "")).strip()
        if not text:
            continue
        parts.append(text)
        if len(" | ".join(parts)) >= limit:
            break
    return " | ".join(parts)[:limit]


def _collect_waybill_entries(page: ExactPage, sheet_title: str,
                             table_infos: list, page_issues: list) -> list[dict]:
    """从一页的所有表格块提取单号索引条目。"""
    source = getattr(page, "source_name", "") or page.name
    source_page = int(getattr(page, "source_page", 0) or 1)
    entries: list[dict] = []
    for info in table_infos:
        block = info["block"]
        rows = getattr(block, "rows", None) or []
        columns = _find_waybill_columns(block)
        if not columns:
            if rows:
                page_issues.append({
                    "source": source, "page": source_page, "sheet": sheet_title,
                    "kind": "未找到单号列",
                    "detail": f"表格{info['table_index'] + 1}前5行未找到运单号/单号表头",
                })
            continue
        flag_reasons = {}
        for fl in info.get("flags") or []:
            try:
                flag_reasons[(int(fl.get("row", -1)), int(fl.get("col", -1)))] = str(
                    fl.get("reason", "") or "交叉校验存疑"
                )
            except (TypeError, ValueError):
                continue

        added_total = 0
        for header_row, waybill_col, header_text in columns:
            added = 0
            for rel_row in range(header_row + 1, len(rows)):
                row = rows[rel_row]
                if waybill_col >= len(row):
                    continue
                raw = str(row[waybill_col] or "").strip()
                if not raw:
                    continue
                if _waybill_header_score(raw) or _looks_like_repeated_header(row, waybill_col):
                    continue
                tokens = _waybill_tokens(raw)
                if not tokens:
                    continue
                for token in tokens:
                    problem = _waybill_problem(token)
                    flag_reason = flag_reasons.get((rel_row, waybill_col), "")
                    format_status, format_reason = _waybill_status(token)
                    merged_hint = _merged_tail_hint(token)
                    notes = []
                    if problem:
                        notes.append(problem)
                    if flag_reason:
                        notes.append(flag_reason)
                    if merged_hint:
                        notes.append(merged_hint)

                    # 只有格式通过、当前页第一张表（交叉校验覆盖范围）且无待复核冲突，
                    # 才允许标为“已确认”。多表格页、格式异常、交叉校验未通过一律待核对。
                    if (format_status == "待核对" and not flag_reason
                            and info["table_index"] == 0):
                        status = "已确认"
                        status_reason = "格式有效，交叉校验无冲突"
                    else:
                        status = (format_status
                                  if format_status in ("格式异常", "非YT") else "待核对")
                        reasons = [r for r in (format_reason, flag_reason) if r]
                        if info["table_index"] > 0:
                            reasons.append("多表格页未完成单号交叉校验")
                        status_reason = "；".join(dict.fromkeys(reasons)) or "未完成确认"
                        notes.append(status_reason)

                    if len(tokens) > 1:
                        notes.append("同一单元格含多个单号")
                    excel_row = info["start_row"] + rel_row
                    col_offset = int(info.get("col_offset", 0) or 0)
                    entries.append({
                        "waybill": token,
                        "status": status,
                        "status_reason": status_reason,
                        "source": source,
                        "page": source_page,
                        "sheet": sheet_title,
                        "cell": f"{get_column_letter(waybill_col + 1 + col_offset)}{excel_row}",
                        "row": excel_row,
                        "header": header_text,
                        "summary": _row_summary(row, waybill_col),
                        "box": (_token_image_box(page, token)
                                or _cell_image_box(page, block, rel_row, waybill_col)),
                        "suspicious": bool(problem or flag_reason or merged_hint),
                        "note": "；".join(dict.fromkeys(notes)),
                    })
                    added += 1
            added_total += added
        if not added_total:
            page_issues.append({
                "source": source, "page": source_page, "sheet": sheet_title,
                "kind": "未提取到单号",
                "detail": (f"表格{info['table_index'] + 1}识别到表头"
                           f"「{'/'.join(c[2] for c in columns)}」，但数据行没有有效单号"),
            })
    return entries


def _write_index_sheet(ws, entries: list[dict]) -> None:
    """单号索引：单号 → 来源、工作表、单元格，可筛选并点击跳转。"""
    headers = (
        "单号", "确认状态", "来源文件", "页码", "工作表", "单元格", "行号",
        "列名", "行摘要", "重复", "可疑", "说明",
    )
    widths = (24, 11, 26, 7, 20, 10, 7, 12, 55, 8, 8, 44)
    for index, (title, width) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=index, value=title)
        cell.font = Font(name=BODY_FONT, size=10.5, bold=True)
        cell.fill = HEADER_FILL
        cell.border = TABLE_BORDER
        cell.alignment = Alignment(vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(index)].width = width

    for row_index, entry in enumerate(entries, start=2):
        values = (
            entry["waybill"], entry["status"], entry["source"], entry["page"],
            entry["sheet"], entry["cell"], entry["row"], entry["header"], entry["summary"],
            "是" if entry.get("duplicate") else "", "是" if entry["suspicious"] else "",
            entry["note"],
        )
        for column_index, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=column_index, value=value)
            cell.border = TABLE_BORDER
            cell.font = Font(name=BODY_FONT, size=10.5)
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center" if column_index in (2, 4, 6, 7, 10, 11) else "left",
                wrap_text=True,
            )
            if column_index == 1:
                cell.number_format = "@"       # 编码字段按文本保存，防止前导 0 丢失
            if column_index == 2:
                if entry["status"] == "已确认":
                    cell.fill = PatternFill("solid", fgColor="E7F6EC")
                    cell.font = Font(name=BODY_FONT, size=10.5, bold=True, color="166534")
                elif entry["status"] == "待核对":
                    cell.fill = PatternFill("solid", fgColor="FFF4D6")
                    cell.font = Font(name=BODY_FONT, size=10.5, bold=True, color="9A6700")
                else:
                    cell.fill = PatternFill("solid", fgColor="FDE2E1")
                    cell.font = Font(name=BODY_FONT, size=10.5, bold=True, color="B42318")

        target = ws.cell(row=row_index, column=6)
        sheet_ref = str(entry["sheet"]).replace("'", "''")
        target.hyperlink = Hyperlink(
            ref=target.coordinate,
            location=f"'{sheet_ref}'!{entry['cell']}",
            display=str(entry["cell"]),
            tooltip=f"打开 {entry['sheet']} 的 {entry['cell']}",
        )
        target.font = Font(name=BODY_FONT, size=10.5, color="0563C1", underline="single")

    last_row = max(1, len(entries) + 1)
    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{last_row}"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _write_index_issue_sheet(ws, issues: list[dict]) -> None:
    headers = ("来源文件", "页码", "工作表", "异常类型", "说明")
    widths = (30, 8, 22, 20, 64)
    for index, (title, width) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=index, value=title)
        cell.font = Font(name=BODY_FONT, size=10.5, bold=True)
        cell.fill = HEADER_FILL
        cell.border = TABLE_BORDER
        cell.alignment = Alignment(vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(index)].width = width
    for row_index, issue in enumerate(issues, start=2):
        values = (issue["source"], issue["page"], issue["sheet"],
                  issue["kind"], issue["detail"])
        for column_index, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=column_index, value=value)
            cell.border = TABLE_BORDER
            cell.font = Font(name=BODY_FONT, size=10.5)
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center" if column_index == 2 else "left",
                wrap_text=True,
            )
    ws.auto_filter.ref = f"A1:E{max(1, len(issues) + 1)}"
    ws.freeze_panes = "A2"


def _write_quality_sheet(wb, quality: dict) -> None:
    """拍摄质量说明 sheet：等级、原因、建议、指标（不改原数据表）。"""
    ws = wb.create_sheet("拍摄质量")
    rows = [
        ("拍摄质量", str(quality.get("level", ""))),
        ("评分", quality.get("score", "")),
        ("原因", "；".join(quality.get("reasons") or [])),
        ("拍法建议", "；".join(quality.get("suggestions") or [])),
        ("", ""),
        ("指标", "数值"),
    ]
    for k, v in (quality.get("metrics") or {}).items():
        rows.append((k, v))
    for row in rows:
        ws.append(list(row))
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 78
    ws.sheet_view.showGridLines = False


def convert_exact(pages: list[ExactPage], out_path: Path, *,
                  include_index: bool = True,
                  include_aux: bool = True,
                  include_flags: bool | None = None,
                  quality: dict | None = None,
                  layout: str = "per_image") -> ConvertResult:
    """全量识别：把每页图片里的表格与文字原样写入 Excel。

    与 convert()（Markdown 入口）并存。layout 控制数据页的输出方式：
      - "per_image"：每页一个工作表（工作表名即来源图片），逐一对照原图；
      - "merged"：全部页面合并到「合并数据」一个工作表，A 列「来源图片」
        标注每行数据来自哪张图；表头相同的表格直接续行，结构不同才重写表头。
    表格与文字块按原图文档顺序排列；行高/列宽尽量还原原图观感。
    交叉校验的待复核单元格在表中黄底高亮，并汇总到「待复核清单」工作表。
    数据页排在最前，打开工作簿即可对照原图；「单号索引」放在最后备查。
    include_index=False 时不生成「单号索引」附加页；
    include_aux=False 时也不生成「索引异常」附加页。
    include_flags 单独控制「待复核清单」「待复核-低优先」两张表（None 表示跟随
    include_aux）；两张表始终生成，0 条时写明「本次无待复核项」，不做静默省略。
    """
    result = ConvertResult()
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    flag_entries: list[tuple[str, dict]] = []
    index_entries: list[dict] = []
    index_issues: list[dict] = []

    merged = str(layout or "").strip().lower() == "merged"
    merged_ws = None
    merged_state = {"row": 1, "header": None}
    merged_ncols = 0
    if merged:
        merged_ws = wb.create_sheet(_sheet_title("合并数据", used))
        blocks_all = [b for page in pages for b in page.blocks if _block_has_content(b)]
        merged_ncols = max(
            (max(len(r) for r in b.rows) for b in blocks_all
             if getattr(b, "kind", "") == "table"),
            default=8,
        )
        merged_ncols = max(merged_ncols, 4)
        merged_ws.column_dimensions["A"].width = 26
        for index, width in enumerate(_exact_widths(blocks_all, merged_ncols), start=2):
            merged_ws.column_dimensions[get_column_letter(index)].width = width

    for page_index, page in enumerate(pages):
        for warning in page.warnings:
            result.warnings.append(f"{page.name}: {warning}")
        for fl in getattr(page, "flags", None) or []:
            flag_entries.append((page.name, fl))
        if not any(_block_has_content(b) for b in page.blocks):
            result.warnings.append(f"{page.name}: 未识别到内容，已跳过。")
            index_issues.append({
                "source": getattr(page, "source_name", "") or page.name,
                "page": int(getattr(page, "source_page", 0) or 1),
                "sheet": "",
                "kind": "未识别到内容",
                "detail": "该页没有可写入的表格或文字，已跳过数据工作表",
            })
            continue
        table_infos: list[dict] = []
        if merged_ws is not None:
            tables = _write_merged_page(
                merged_ws, merged_state, page, table_infos, merged_ncols)
            sheet_title = merged_ws.title
            result.sheets = 1
        else:
            ws = wb.create_sheet(_sheet_title(page.name, used))
            tables, _has = _write_exact_sheet(ws, page, table_infos=table_infos)
            sheet_title = ws.title
            result.sheets += 1
        result.tables += tables
        page_issues: list[dict] = []
        page_entries = _collect_waybill_entries(page, sheet_title, table_infos, page_issues)
        if page_entries:
            for entry in page_entries:
                entry["image_index"] = page_index
            index_entries.extend(page_entries)
        else:
            if not any(getattr(b, "kind", "") == "table" for b in page.blocks):
                page_issues = [{
                    "source": getattr(page, "source_name", "") or page.name,
                    "page": int(getattr(page, "source_page", 0) or 1),
                    "sheet": sheet_title,
                    "kind": "未识别到表格",
                    "detail": "该页只有文字或空内容，无法生成单号索引",
                }]
            elif not page_issues:
                page_issues = [{
                    "source": getattr(page, "source_name", "") or page.name,
                    "page": int(getattr(page, "source_page", 0) or 1),
                    "sheet": sheet_title,
                    "kind": "未提取到单号",
                    "detail": "表格中未提取到有效单号",
                }]
            index_issues.extend(page_issues)

    if result.sheets == 0:
        raise ValueError("没有可写入的内容")

    duplicates = Counter(entry["waybill"].upper() for entry in index_entries)
    for entry in index_entries:
        count = duplicates[entry["waybill"].upper()]
        entry["duplicate"] = count > 1
        if count > 1:
            note = f"出现 {count} 处"
            entry["note"] = f"{entry['note']}；{note}" if entry["note"] else note

    result.flagged = len(flag_entries)

    show_flags = include_aux if include_flags is None else include_flags
    if show_flags:
        hi_mid_flags, low_flags = _triage_flags(flag_entries)
        ws = wb.create_sheet(_sheet_title("待复核清单", used))
        _write_flag_sheet(ws, hi_mid_flags, low_count=len(low_flags))
        ws_low = wb.create_sheet(_sheet_title("待复核-低优先", used))
        _write_flag_low_sheet(
            ws_low, low_flags,
            high_count=sum(1 for e in hi_mid_flags if e[3] == "高"),
            mid_count=sum(1 for e in hi_mid_flags if e[3] == "中"))

    if include_aux and index_issues:
        ws = wb.create_sheet(_sheet_title("索引异常", used))
        _write_index_issue_sheet(ws, index_issues)

    # 附加页放最后：打开工作簿第一眼就是原图对应的数据页
    if include_index:
        index_ws = wb.create_sheet(_sheet_title("单号索引", used))
        _write_index_sheet(index_ws, index_entries)

    result.waybills = len(index_entries)
    result.unique_waybills = len(duplicates)
    result.duplicate_waybills = sum(1 for count in duplicates.values() if count > 1)
    result.suspicious_waybills = sum(1 for entry in index_entries if entry["suspicious"])
    result.confirmed_waybills = sum(1 for entry in index_entries if entry["status"] == "已确认")
    result.pending_waybills = sum(1 for entry in index_entries if entry["status"] == "待核对")
    result.format_invalid_waybills = sum(
        1 for entry in index_entries if entry["status"] in ("格式异常", "非YT")
    )
    result.index_issues = len(index_issues)
    # 供应用内单号搜索使用（Web 侧只在内存中保留最近几次任务）
    result.waybill_entries = index_entries
    result.index_issue_items = index_issues

    if quality:
        _write_quality_sheet(wb, quality)
    wb.active = 0
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return result


def convert_waybill(pages: list[ExactPage], out_path: Path) -> ConvertResult:
    """纯单号：整份 Excel 只有一列运单号，按出现顺序自动去重。

    识别仍走与全量识别相同的表格几何结果，只抽取「运单号/单号」列，
    不写数据工作表、不写单号索引等附加页。格式异常/非 YT 候选不写入清单
    （计入统计并在界面提示），避免把明显不是运单号的串混进来。
    """
    result = ConvertResult()
    candidates: list[dict] = []
    index_issues: list[dict] = []

    for page in pages:
        for warning in page.warnings:
            result.warnings.append(f"{page.name}: {warning}")
        source = getattr(page, "source_name", "") or page.name
        source_page = int(getattr(page, "source_page", 0) or 1)
        table_infos: list[dict] = []
        for block in page.blocks:
            if getattr(block, "kind", "") != "table" or not _block_has_content(block):
                continue
            table_infos.append({
                "block": block,
                "start_row": 1,          # 纯单号不落地单元格地址，仅用于抽取
                "table_index": len(table_infos),
                "flags": (getattr(page, "flags", None) or []) if not table_infos else [],
            })
        page_issues: list[dict] = []
        candidates.extend(_collect_waybill_entries(page, "单号", table_infos, page_issues))
        if not table_infos:
            page_issues.append({
                "source": source, "page": source_page, "sheet": "单号",
                "kind": "未识别到表格",
                "detail": "该页没有可写入的表格，无法提取运单号",
            })
        index_issues.extend(page_issues)

    valid = [c for c in candidates if c["status"] in ("已确认", "待核对")]
    counts = Counter(c["waybill"].upper() for c in valid)
    first_entries: dict[str, dict] = {}
    for entry in valid:
        first_entries.setdefault(entry["waybill"].upper(), entry)

    result.waybills = len(candidates)
    result.unique_waybills = len(first_entries)
    result.duplicate_waybills = sum(1 for count in counts.values() if count > 1)
    result.confirmed_waybills = sum(
        1 for entry in first_entries.values() if entry["status"] == "已确认"
    )
    result.pending_waybills = sum(
        1 for entry in first_entries.values() if entry["status"] == "待核对"
    )
    result.format_invalid_waybills = len(candidates) - len(valid)
    result.suspicious_waybills = sum(
        1 for entry in first_entries.values() if entry["suspicious"]
    )
    result.index_issues = len(index_issues)
    result.index_issue_items = index_issues

    if not first_entries:
        raise ValueError("没有识别到可导出的运单号（格式异常/非 YT 项不会写入纯单号清单）")

    wb = Workbook()
    ws = wb.active
    ws.title = "单号"
    header = ws.cell(row=1, column=1, value="单号")
    header.font = Font(name=BODY_FONT, size=10.5, bold=True)
    header.fill = HEADER_FILL
    header.border = TABLE_BORDER
    header.alignment = Alignment(vertical="center", horizontal="center")
    ws.column_dimensions["A"].width = 26

    for row_index, token in enumerate(first_entries, start=2):
        cell = ws.cell(row=row_index, column=1, value=token)
        cell.number_format = "@"          # 编码字段按文本保存，防止前导 0 丢失
        cell.border = TABLE_BORDER
        cell.font = Font(name=BODY_FONT, size=10.5)
        cell.alignment = Alignment(vertical="center", horizontal="left")

    ws.auto_filter.ref = f"A1:A{len(first_entries) + 1}"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return result


# ---------------------------------------------------------------- 命令行


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Markdown 转 Excel")
    parser.add_argument("inputs", nargs="+", help="Markdown 文件路径")
    parser.add_argument("-o", "--output", required=True, help="输出 xlsx 路径")
    parser.add_argument(
        "--mode", choices=["structure", "table"], default="structure", help="导出模式"
    )
    args = parser.parse_args()

    items = [
        SourceDoc(name=Path(p).stem, markdown=Path(p).read_text(encoding="utf-8", errors="replace"))
        for p in args.inputs
    ]
    res = convert(items, Path(args.output), args.mode)
    print(f"完成：{res.sheets} 个工作表，{res.tables} 个表格 -> {args.output}")
    for warning in res.warnings:
        print("提示：" + warning)
