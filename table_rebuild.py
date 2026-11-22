# -*- coding: utf-8 -*-
"""表格几何重建：OCR 文本框 -> 行聚类 -> 列聚类 -> 网格，标题/正文分离。

输入为 ocr_worker 的 box 列表（{"text","score","bbox":[x1,y1,x2,y2]}，绝对坐标）。
纯 stdlib，可同时被 venv worker 与系统 python app 进程导入。

输出块：
  TableBlock  行列表（rows[r][c] 为单元格文本，缺失补 ""），含首行表头
  TextBlock   标题/正文行（lines）
块按文档顺序（y）排列，表格与文字交替。
"""
import re
from collections import defaultdict
from dataclasses import dataclass, field

_CJK_RE = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]")

# 列聚类 gap：相邻框中心 x 差 <= 此值视为同列。弯曲纸张上同列框区间会漂移
# 搭接到邻列，但中心分得开；实测手机上 9 个列带全部正确分开。
_COL_GAP = 20.0


@dataclass
class Box:
    text: str
    score: float
    x1: float
    y1: float
    x2: float
    y2: float

    @property
    def xc(self):
        return (self.x1 + self.x2) / 2.0

    @property
    def yc(self):
        return (self.y1 + self.y2) / 2.0

    @property
    def w(self):
        return self.x2 - self.x1

    @property
    def h(self):
        return self.y2 - self.y1


@dataclass
class TableBlock:
    kind: str = "table"
    rows: list = field(default_factory=list)     # list[list[str]]
    bbox: tuple = (0, 0, 0, 0)
    col_x: list = field(default_factory=list)    # 各列 x 区间 [(x1,x2)...]
    row_y: list = field(default_factory=list)    # 各行 y 区间 [(y1,y2)...]，补行定位用

    @property
    def n_rows(self):
        return len(self.rows)

    @property
    def n_cols(self):
        return max((len(r) for r in self.rows), default=0)


@dataclass
class TextBlock:
    kind: str = "text"
    lines: list = field(default_factory=list)
    bbox: tuple = (0, 0, 0, 0)


@dataclass
class RebuildResult:
    blocks: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


def _median(xs):
    if not xs:
        return 0.0
    s = sorted(xs)
    n = len(s)
    return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2.0


def _to_boxes(boxes):
    out = []
    for b in boxes:
        if isinstance(b, Box):
            out.append(b)
            continue
        try:
            bb = b["bbox"]
            out.append(Box(text=str(b.get("text", "")), score=float(b.get("score", 0)),
                           x1=float(bb[0]), y1=float(bb[1]),
                           x2=float(bb[2]), y2=float(bb[3])))
        except Exception:
            continue
    return [b for b in out if b.text.strip() and b.x2 > b.x1 and b.y2 > b.y1]


def _cluster_rows(boxes):
    """按 y 中心聚成行。返回 list[list[Box]]（每行按 x 排序）。

    行锚 = 该行第一个（yc 最小）框的中心，固定不漂移：若用运行均值做行中心，
    行每吞入一个框中心就下移一点，倾斜/行距紧凑的文档会级联地把后续行全部
    吸进同一行（实测 256 框的照片雪崩成 6 行）。锚定后行带宽度有上界（±tol）。
    """
    boxes = sorted(boxes, key=lambda b: (b.yc, b.x1))
    med_h = _median([b.h for b in boxes]) or 1.0
    tol = 0.5 * med_h
    rows = []
    for b in boxes:
        placed = False
        for row in reversed(rows):          # 已按 yc 排序，正确行总在末尾几行
            if abs(b.yc - row[0].yc) > tol:
                continue
            r_y1 = min(x.y1 for x in row)
            r_y2 = max(x.y2 for x in row)
            ov = min(r_y2, b.y2) - max(r_y1, b.y1)
            if ov >= -0.3 * b.h:            # 允许字形框轻微不重叠
                row.append(b)
                placed = True
                break
        if not placed:
            rows.append([b])
    for row in rows:
        row.sort(key=lambda b: b.x1)
    rows.sort(key=lambda r: min(b.y1 for b in r))
    return rows


def _infer_columns(rows, boxes):
    """列推断：按框中心 x 聚类（gap 与 _column_bands 一致）。

    仅"多框行"的框参与聚类——单框行是标题/正文/脚注，其宽文本框横跨多列，
    若参与会把相邻列桥接成一个列簇。弯曲纸张上同一列的框 x 区间随 y 漂移，
    会与相邻列的区间互相搭接（区间并查集会按"重叠 >= 半宽"把两列并成一列），
    但各列框的中心分得开，故按中心聚类：相邻中心差 <= gap 视为同列。

    返回 (col_ranges, assign)：col_ranges[c]=(x1,x2) 为该列成员框的 x 包络，
    assign[id(box)]=col_id（未参与者不在 assign 中）。
    """
    cand = []
    for row in rows:
        if len(row) >= 2:
            cand.extend(row)
    cand.sort(key=lambda b: b.xc)
    groups, cur = [], []
    for b in cand:
        if cur and b.xc - cur[-1].xc > _COL_GAP:
            groups.append(cur)
            cur = []
        cur.append(b)
    if cur:
        groups.append(cur)
    assign = {}
    col_ranges = []
    for cid, g in enumerate(groups):
        col_ranges.append((min(b.x1 for b in g), max(b.x2 for b in g)))
        for b in g:
            assign[id(b)] = cid
    return col_ranges, assign


def _join_text(a, b):
    if not a:
        return b
    if not b:
        return a
    if _CJK_RE.search(a[-1]) or _CJK_RE.search(b[0]):
        return a + b
    return a + " " + b


def _cell_text(cell_boxes):
    """同一单元格内多个框：先去重（高重叠保留高分），再拼接。"""
    keep = []
    for b in sorted(cell_boxes, key=lambda b: -b.score):
        dup = False
        for k in keep:
            ov = min(k.x2, b.x2) - max(k.x1, b.x1)
            if ov > 0.6 * min(k.w, b.w):
                dup = True
                break
        if not dup:
            keep.append(b)
    keep.sort(key=lambda b: b.x1)
    text = ""
    for b in keep:
        text = _join_text(text, b.text.strip())
    return text


def _column_bands(boxes, gap=_COL_GAP):
    """按框中心 x 聚类成列带（升序）。弯曲/错列纸张上各列的框中心分得开，
    但框 x 区间可能互相搭接，故按中心而非区间聚类。"""
    cs = sorted(boxes, key=lambda b: b.xc)
    bands, cur = [], [cs[0]]
    for b in cs[1:]:
        if b.xc - cur[-1].xc <= gap:
            cur.append(b)
        else:
            bands.append(cur)
            cur = [b]
    bands.append(cur)
    return bands


def _rows_ycs(rows):
    return [sum(b.yc for b in r) / len(r) for r in rows]


def _serial_values(rows):
    """行的序号值列表（无序号的行记 None）：取行内最左侧的纯整数框。"""
    vals = []
    for row in rows:
        v = None
        for b in sorted(row, key=lambda x: x.x1):
            t = b.text.strip()
            if t.isdigit():
                v = int(t)
                break
        vals.append(v)
    return vals


def _serial_main_segment(vals):
    """序号值的最长密集连续段（抗 OCR 误读离群值，如 '111'→'1'）。

    误读值往往是孤立离群点（与真值段相差很远），而漏识别只会造成小断档，
    故相邻唯一值差 <=3 视为同一段，取最长段作为该栏序号的真实值域。
    """
    xs = sorted(set(v for v in vals if v is not None))
    if not xs:
        return []
    best, seg = [], [xs[0]]
    for a, b in zip(xs, xs[1:]):
        if b - a <= 3:              # 允许个别行漏识别造成的断档
            seg.append(b)
        else:
            if len(seg) > len(best):
                best = seg
            seg = [b]
    if len(seg) > len(best):
        best = seg
    return best


def _is_serial_band(rows, vals, ratio=0.6):
    """序号列带判定：多数行为纯整数，且主密集段值域覆盖率高（允许个别
    OCR 误读打乱局部顺序，如 '4'→'1'、'55'→'66'，故不要求严格递增；
    误读离群值只参与"多数行为整数"的计数，不参与值域跨度计算）。
    另要求值基本唯一：数量/页数等整数数据列重复率高（如 10/8/12/10…），
    借唯一性排除，避免把数据列当序号列触发错位拼接。"""
    xs = [v for v in vals if v is not None]
    if len(xs) < max(4, ratio * len(rows)):
        return False
    if len(set(xs)) < 0.8 * len(xs):
        return False
    seg = _serial_main_segment(vals)
    if len(seg) < 4 or seg[0] < 1:
        return False
    span = seg[-1] - seg[0] + 1
    return len(seg) >= 0.7 * span


def _align_rows_no_serial(bs, bands, infos, log):
    """无可靠序号带的弯曲纸张行重建（列间 y 错位随 y 缓变的透视/弯曲变形）。

    各带已独立聚行（带内错位小）；选最完整的带为参考，先以"显著最近邻"
    收集高置信行对，拟合带间线性 y 偏移 Δ(y)=a+b·(y-y0)（透视变形近似线性），
    校正后按最近邻（容差 0.5·pitch）对齐到参考带行号；未匹配行按投影单独成行。
    返回行列表；不符返回 None（调用方回退全局 y 聚类）。
    """
    if len(infos) < 2:
        return None
    ref = max(infos, key=lambda x: len(x["rows"]))
    rycs = ref["ycs"]
    y0 = rycs[0]
    pr = ref["pitch"]

    def nominal(yc):
        return int(round((yc - y0) / pr))

    g_of_ref = [nominal(yc) for yc in rycs]

    grid = defaultdict(list)
    for i, row in enumerate(ref["rows"]):
        grid[g_of_ref[i]].extend(row)

    shift_of = {id(ref): (0.0, 0.0)}
    matched = total = 0
    for x in infos:
        if x is ref:
            continue
        rows, ycs = x["rows"], x["ycs"]
        # 高置信行对：最近邻显著优于次优，用于拟合线性偏移
        pairs = []
        for yc in ycs:
            order = sorted(range(len(rycs)), key=lambda i: abs(rycs[i] - yc))
            i0, i1 = order[0], order[1]
            e0, e1 = abs(rycs[i0] - yc), abs(rycs[i1] - yc)
            if e0 < 0.35 * x["pitch"] and e1 - e0 > 0.15 * x["pitch"]:
                pairs.append((yc, rycs[i0] - yc))
        a, b = 0.0, 0.0
        if len(pairs) >= 2:
            n = float(len(pairs))
            sx = sum(yc - y0 for yc, _ in pairs)
            sy = sum(d for _, d in pairs)
            sxx = sum((yc - y0) ** 2 for yc, _ in pairs)
            sxy = sum((yc - y0) * d for yc, d in pairs)
            den = n * sxx - sx * sx
            if abs(den) > 1e-9:
                b = (n * sxy - sx * sy) / den
                a = (sy - b * sx) / n
            else:
                a = sy / n
        elif pairs:
            a = pairs[0][1]
        shift_of[id(x)] = (a, b)

        for row, yc in zip(rows, ycs):
            total += 1
            yc2 = yc - (a + b * (yc - y0))
            i0 = min(range(len(rycs)), key=lambda i: abs(rycs[i] - yc2))
            if abs(rycs[i0] - yc2) <= 0.5 * pr:
                k = g_of_ref[i0]
                matched += 1
            else:
                k = nominal(yc2)
            grid[k].extend(row)

    if total and matched < 0.5 * total:
        return None
    if len(grid) < 3:
        return None
    # 未成带的散框挂到最近带的偏移投影行
    banded = set()
    for band in bands:
        if len(band) >= 6:
            banded.update(id(b) for b in band)
    for bx in bs:
        if id(bx) not in banded:
            near = min(infos, key=lambda x: abs(x["xc"] - bx.xc))
            a2, b2 = shift_of[id(near)]
            grid[nominal(bx.yc - (a2 + b2 * (bx.yc - y0)))].append(bx)
    rows = [sorted(grid[k], key=lambda b: b.x1) for k in sorted(grid)]
    log(f"无序号列：以最完整列带为参考做跨列行对齐（{matched}/{total} 行匹配）")
    return rows


def _stitch_staggered_rows(bs, log):
    """列错位版式（弯曲/侧拍的多栏续表）的行重建。返回行列表；不符返回 None。

    弯曲纸张上，列与列之间的行 y 错位随 x 累积（可达半个行距以上），
    全局 y 聚类会把相邻行混在一起；单列带内部错位很小，可独立分行。
    以序号列带为骨架：序号值 v 的行号 k=(v-1) mod L（L=相邻序号带起始值之差，
    即每栏行数），对带内漏识别完全鲁棒；其余列带的行按 y 相对最近序号带
    k=0 行的投影定位：k=round((yc-y0)/p)。
    """
    bands = _column_bands(bs)
    if len(bands) < 2:
        return None
    infos = []
    for band in bands:
        if len(band) < 6:
            continue                    # 稀疏带（本就没几行）后面按 y 投影挂载
        rows = _cluster_rows(band)
        if len(rows) < 5:
            continue
        ycs = _rows_ycs(rows)
        diffs = [b - a for a, b in zip(ycs, ycs[1:]) if b > 0]
        if not diffs:
            continue
        p = _median(diffs)
        if p <= 6:
            continue
        vals = _serial_values(rows)
        infos.append({"rows": rows, "ycs": ycs, "pitch": p, "vals": vals,
                      "serial": _is_serial_band(rows, vals),
                      "xc": sum(b.xc for b in band) / len(band)})
    serials = [x for x in infos if x["serial"]]
    if not serials:
        return _align_rows_no_serial(bs, bands, infos, log)
    ps = [x["pitch"] for x in infos]
    p = _median(ps)
    grid = defaultdict(list)

    # 序号带：k = (v-1) mod L；L 优先取相邻序号带起始值差的中位数（=每栏行数），
    # 起始值取主密集段最小值，避免 '111'→'1' 这类误读把起始值拉到 1 污染 L；
    # 再用"过半序号带行数与 L 差距 <=2"交叉验证，失败则改按行数中位数。
    L = None
    if len(serials) >= 2:
        starts = [_serial_main_segment(x["vals"])[0] for x in serials]
        gaps = [b - a for a, b in zip(starts, starts[1:]) if b > a]
        if gaps:
            L = _median(gaps)
    if L is not None and L < 4:
        L = None
    if L is None:
        if len(serials) >= 2:
            # 多栏但起始值不可用（误读/倒序）：按行数中位数估 L。
            # 单序号带不走此路——漏识别会使行数 < 序号个数，(v-1)%L 会绕回碰撞，
            # 保持原来的 k=v-1（要求从 1 开始）语义。
            counts = [len(x["rows"]) for x in serials]
            if counts:
                L = _median(counts)
            if L is not None and L < 4:
                L = None
    else:
        counts = [len(x["rows"]) for x in serials]
        bad = sum(1 for c in counts if abs(c - L) > 2)
        if counts and bad * 2 > len(counts):
            L2 = _median(counts)
            if L2 >= 4:
                L = L2
    L = int(round(L)) if L is not None else None
    if L is None:
        # 单序号带：仅当序号从 1 开始才可按值定位（k=v-1）；否则退回
        # 无序号带对齐（弯曲纸张的透视变形下全局 y 聚类会劈裂行）
        if len(serials) > 1 or _serial_main_segment(
                [v for x in serials for v in x["vals"] if v is not None])[0] != 1:
            return _align_rows_no_serial(bs, bands, infos, log)
    y0_of = {}                          # 序号带 -> 该带 k=0 行的投影 y
    for x in serials:
        # k=0 行 y：优先取真实 k=0 行，否则按最小 k 行投影
        y0 = None
        for row, v in zip(x["rows"], x["vals"]):
            if v is None:
                continue
            k = (v - 1) % L if L is not None else v - 1
            if k == 0:
                y0 = sum(b.yc for b in row) / len(row)
                break
        if y0 is None:
            mined = [(v, row) for v, row in zip(x["vals"], x["rows"]) if v is not None]
            if not mined:
                continue
            # 仅按序号值取最小：重复值（如数量列 10/8/5/10）时若平级比较
            # 会退到比较 Box 列表，而 Box 无排序支持会 TypeError
            v, row = min(mined, key=lambda t: t[0])
            yc = sum(b.yc for b in row) / len(row)
            k = ((v - 1) % L) if L is not None else 0
            y0 = yc - k * x["pitch"]
        y0_of[id(x)] = y0
        # 挂载：有序号的行按值定位，无序号的行按 y 投影
        for row, v in zip(x["rows"], x["vals"]):
            yc = sum(b.yc for b in row) / len(row)
            if v is not None and L is not None:
                k = (v - 1) % L
            else:
                k = int(round((yc - y0) / p))
            grid[k].extend(row)
    if not y0_of:
        return None
    # 其余带（单号列/稀疏带/文字带）：按 y 相对最近序号带的 k=0 行投影
    for x in infos:
        if x["serial"]:
            continue
        nearest = min(serials, key=lambda s: abs(s["xc"] - x["xc"]))
        y0 = y0_of[id(nearest)]
        for row in x["rows"]:
            yc = sum(b.yc for b in row) / len(row)
            grid[int(round((yc - y0) / p))].extend(row)
    # 未成带的散框（极小列、游离标题等）也挂上，避免丢内容
    banded = set()
    for band in bands:
        if len(band) >= 6:
            banded.update(id(b) for b in band)
    for b in bs:
        if id(b) not in banded:
            nearest = min(serials, key=lambda s: abs(s["xc"] - b.xc))
            y0 = y0_of[id(nearest)]
            grid[int(round((b.yc - y0) / p))].append(b)
    if len(grid) < 3:
        return None
    rows = [sorted(grid[k], key=lambda b: b.x1) for k in sorted(grid)]
    log(f"检测到多栏错位版式（{len(serials)} 个序号列"
        f"{('，每栏 ' + str(L) + ' 行') if L else ''}），已按序号对齐拼接")
    return rows


def rebuild(boxes, on_log=None, min_table_rows=2):
    """主入口：OCR boxes -> blocks（表格/文字交替）。"""
    res = RebuildResult()
    bs = _to_boxes(boxes)
    if not bs:
        res.warnings.append("无有效文本框")
        return res

    def log(msg):
        if on_log:
            on_log(msg)

    rows = _stitch_staggered_rows(bs, log)
    if rows is None:
        rows = _cluster_rows(bs)
    col_ranges, assign = _infer_columns(rows, bs)
    n_cols_total = len(col_ranges)

    # 列使用统计：每列被多少行使用（单框行的框未参与列推断，不在 assign 中）
    col_rows = defaultdict(set)
    for ri, row in enumerate(rows):
        for b in row:
            cid = assign.get(id(b))
            if cid is not None:
                col_rows[cid].add(ri)
    n_rows_total = len(rows)
    major = {cid for cid, rs in col_rows.items() if len(rs) >= 2}

    # 判定每行是否为表格行：>=2 个框落在主要列，且主要列框占比 >= 0.6
    is_table_row = []
    for row in rows:
        majors = [b for b in row if assign.get(id(b)) in major]
        is_table_row.append(len(majors) >= 2 and len(majors) >= 0.6 * len(row))

    def log(msg):
        if on_log:
            on_log(msg)

    log(f"重建: {len(bs)} 框 -> {n_rows_total} 行 x {n_cols_total} 列簇, "
        f"主要列 {len(major)} 个")

    # 连续表格行构成一个表；其余行是文字
    i = 0
    while i < n_rows_total:
        if is_table_row[i]:
            j = i
            while j + 1 < n_rows_total and is_table_row[j + 1]:
                j += 1
            res.blocks.append(_build_table(rows[i:j + 1], assign, major, col_ranges, res))
            i = j + 1
        else:
            j = i
            lines = []
            while j < n_rows_total and not is_table_row[j]:
                lines.append("".join(b.text.strip() for b in rows[j]))
                j += 1
            if any(l.strip() for l in lines):
                blk = TextBlock(lines=lines)
                blk.bbox = _bbox_of(rows[i:j])
                res.blocks.append(blk)
            i = j

    tables = [b for b in res.blocks if b.kind == "table"]
    log(f"重建完成: {len(tables)} 个表格, {len(res.blocks) - len(tables)} 个文字块")
    return res


def _bbox_of(rows):
    x1 = min(b.x1 for r in rows for b in r)
    y1 = min(b.y1 for r in rows for b in r)
    x2 = max(b.x2 for r in rows for b in r)
    y2 = max(b.y2 for r in rows for b in r)
    return (x1, y1, x2, y2)


def _nearest_major(box, used, col_ranges):
    """表格行内游离框归列：并入与其 x 包络重叠 >= 自身宽度一半的主要列中重叠最大者。

    重叠不足一半的宽文本框（跨列段落/备注）不予归列，避免桥接真实列。
    返回 col_id 或 None。
    """
    best, best_ov = None, 0.0
    for cid in used:
        x1, x2 = col_ranges[cid]
        ov = min(box.x2, x2) - max(box.x1, x1)
        if ov >= 0.5 * box.w and ov > best_ov:
            best, best_ov = cid, ov
    return best


def _split_oversized_row(row, med_h):
    """把被行聚类并到一起的多行数据拆开（只对异常高行生效）。

    依据：行内出现 >=2 个不同 YT 单号，且这些单号的 y 中心能按空隙分成
    >=2 组。没有这个依据的超高行（单个单号 + 竖排备注/表格线）原样保留，
    避免把多行单元格硬切成假行。

    返回 list[list[Box]]；不满足条件时返回 [row]。
    """
    codes = []
    for b in row:
        m = re.search(r"YT\d{13}", b.text)
        if m:
            codes.append((b, m.group(0)))
    if len({c for _, c in codes}) < 2:
        return [row]
    gap = max(10.0, 0.6 * med_h)
    codes.sort(key=lambda t: t[0].yc)
    groups = []
    for b, code in codes:
        if groups and b.yc - groups[-1][-1][0].yc <= gap:
            groups[-1].append((b, code))
        else:
            groups.append([(b, code)])
    if len(groups) < 2:
        return [row]
    # 组间边界取两组 y 中心的中点；若相邻组贴得太近，不硬切
    bounds = []
    for g1, g2 in zip(groups, groups[1:]):
        y1 = max(b.yc for b, _ in g1)
        y2 = min(b.yc for b, _ in g2)
        if y2 - y1 < gap:
            return [row]
        bounds.append((y1 + y2) / 2.0)
    parts = [[] for _ in range(len(bounds) + 1)]
    for b in row:
        idx = 0
        while idx < len(bounds) and b.yc > bounds[idx]:
            idx += 1
        parts[idx].append(b)
    parts = [p for p in parts if p]
    if len(parts) < 2:
        return [row]
    # 每个切出的行都必须还带一个单号，否则视为普通高行不切
    if not all(any(re.search(r"YT\d{13}", b.text) for b in p) for p in parts):
        return [row]
    return parts


def _split_row_y(part, med_h):
    """拆分行的 y 边界：排除跨多行的超高框，避免 224px 巨行的边界残留。

    正常框（高度 <= 1.5 倍中位行高）取真实上下沿；若整组都是超高框，
    退回按框中心 ± 半行高，保证仍是一个可定位的行带。
    """
    normal = [b for b in part if b.h <= 1.5 * med_h]
    if normal:
        return (min(b.y1 for b in normal), max(b.y2 for b in normal))
    ycs = [b.yc for b in part]
    return (min(ycs) - med_h / 2.0, max(ycs) + med_h / 2.0)


def _build_table(tbl_rows, assign, major, col_ranges, res):
    # 表内实际用到的列（按 x 排序重编号），缺失列补 ""
    used = set()
    for row in tbl_rows:
        for b in row:
            cid = assign.get(id(b))
            if cid in major:
                used.add(cid)
    used = sorted(used)
    remap = {cid: k for k, cid in enumerate(used)}
    n_cols = len(used)
    dropped = 0
    # 行聚类可能把相邻 2-4 行并进一行（如 6320eb7a 第 4 行 224px，吞掉
    # 单号 96/97）。只对明显超过中位行高的行，按行内单号 y 分布拆分；
    # 普通行不受影响，拆出的每个框仍沿用原列归属，列对齐不变。
    row_heights = [max(b.y2 for b in row) - min(b.y1 for b in row) for row in tbl_rows]
    med_h = _median(row_heights) or 1.0
    split_count = 0
    split_rows = []
    for row in tbl_rows:
        h = max(b.y2 for b in row) - min(b.y1 for b in row)
        if h > 1.5 * med_h:
            parts = _split_oversized_row(row, med_h)
            if len(parts) > 1:
                split_count += 1
                split_rows.extend((p, True) for p in parts)
                continue
        split_rows.append((row, False))
    if split_count:
        res.warnings.append(f"拆分 {split_count} 个异常高行（多单号同行）")
    rows_out = []
    rows_y = []
    for row, was_split in split_rows:
        cells = [[] for _ in range(n_cols)]
        for b in row:
            cid = assign.get(id(b))
            if cid not in remap:
                # 游离框（非主要列/未推断）并入最近主要列，避免内容丢失
                cid = _nearest_major(b, used, col_ranges)
            if cid in remap:
                cells[remap[cid]].append(b)
            else:
                dropped += 1
        rows_out.append([_cell_text(c) for c in cells])
        if was_split:
            rows_y.append(_split_row_y(row, med_h))
        else:
            rows_y.append((min(b.y1 for b in row), max(b.y2 for b in row)))
    blk = TableBlock(rows=rows_out, bbox=_bbox_of(tbl_rows),
                     col_x=[col_ranges[c] for c in used], row_y=rows_y)
    # 空单元格统计（提示可能存在漏识别）
    empty = sum(1 for r in rows_out for c in r if not c.strip())
    if empty:
        res.warnings.append(f"表格 {len(rows_out)}x{n_cols} 有 {empty} 个空单元格")
    if dropped:
        res.warnings.append(f"表格有 {dropped} 个框无法归列被丢弃")
    return blk


# ---------------------------------------------------------------- 序号断档检测

def find_serial_gaps(table, col=0):
    """检测首列序号断档。返回 [(缺失值, 行下标)]；行下标为应插入位置（在该行之前）。"""
    vals = []
    for r, row in enumerate(table.rows):
        v = row[col].strip() if col < len(row) else ""
        vals.append(v)
    nums = []
    for r, v in enumerate(vals):
        if v.isdigit():
            nums.append((r, int(v)))
    if len(nums) < 2:
        return []
    gaps = []
    for (r1, v1), (r2, v2) in zip(nums, nums[1:]):
        if v2 > v1 + 1:
            for missing in range(v1 + 1, v2):
                gaps.append((missing, r2))
    return gaps


# ---------------------------------------------------------------- 序号断档补行

def _col_of(box, col_x):
    """框归列：与 col_x 包络重叠最大的列；全无重叠时取中心距最近列。"""
    if not col_x:
        return None
    best, best_ov = None, 0.0
    for cid, (x1, x2) in enumerate(col_x):
        ov = min(box.x2, x2) - max(box.x1, x1)
        if ov > best_ov:
            best, best_ov = cid, ov
    if best is not None:
        return best
    cx = box.xc
    return min(range(len(col_x)),
               key=lambda c: abs(cx - (col_x[c][0] + col_x[c][1]) / 2.0))


def fill_serial_gaps(table, ocr_fn, on_log=None, scale=3.0, serial_col=0):
    """序号断档补行：对缺失行区域 clip+scale 重识别，就地补入 table。

    ocr_fn(clip, scale) -> boxes：clip=[x1,y1,x2,y2] 为原图矩形，scale 为局部
    放大倍数，返回原图坐标的 box 列表（ocr_worker 协议）。序号列一律以断档
    推导值为准（几何槽位 + 上下文序列）；重识别无文本则不插行（可能原文档
    本身跳号）。返回补充行数。
    """
    def log(msg):
        if on_log:
            on_log(msg)

    if not table.row_y or len(table.row_y) != len(table.rows):
        log("补行: 无行定位信息，跳过")
        return 0
    gaps = find_serial_gaps(table, col=serial_col)
    if not gaps:
        return 0
    groups = defaultdict(list)          # 插入行下标 -> [缺失序号...]
    for missing, r in gaps:
        groups[r].append(missing)

    # 重识别 x 范围：列包络并表 bbox
    x1 = min([c[0] for c in table.col_x] + [table.bbox[0]])
    x2 = max([c[1] for c in table.col_x] + [table.bbox[2]])
    n_cols = len(table.col_x)

    filled = 0
    for r in sorted(groups, reverse=True):     # 从后往前插，避免行下标漂移
        missing_vals = sorted(groups[r])
        m = len(missing_vals)
        if r < 1:
            continue
        gy1, gy2 = table.row_y[r - 1][1], table.row_y[r][0]
        if gy2 - gy1 < 2:
            log(f"补行: 序号 {missing_vals[0]} 区域高度异常，跳过")
            continue
        pitch = (gy2 - gy1) / m
        try:
            new_boxes = _to_boxes(ocr_fn([x1, gy1, x2, gy2], scale))
        except Exception as e:
            log(f"补行: 序号 {missing_vals[0]} 区域重识别失败: {e}")
            continue

        def inside(b):
            # 只保留缺口内的框：yc 在 [gy1, gy2]，且与相邻行 y 重叠 < 60% 框高
            if not (gy1 <= b.yc <= gy2):
                return False
            for ry in (table.row_y[r - 1], table.row_y[r]):
                if min(ry[1], b.y2) - max(ry[0], b.y1) >= 0.6 * b.h:
                    return False
            return True

        inner = [b for b in new_boxes if inside(b)]
        if not inner:
            log(f"补行: 序号 {missing_vals[0]}-{missing_vals[-1]} 区域无文本，"
                f"可能原文档本身跳号，跳过")
            continue
        # 按 yc 映射到缺失槽位（同一槽位的多簇合并）
        slots = defaultdict(list)
        for row in _cluster_rows(inner):
            yc = sum(b.yc for b in row) / len(row)
            k = max(0, min(m - 1, int((yc - gy1) / pitch)))
            slots[k].extend(row)
        inserts = []                    # [(row_text, row_y)]，按槽位升序
        for k in sorted(slots):
            boxes = slots[k]
            cells = [[] for _ in range(n_cols)]
            ocr_serial = ""
            for b in boxes:
                cid = _col_of(b, table.col_x)
                if cid is None:
                    continue
                cells[cid].append(b)
                if cid == serial_col and not ocr_serial:
                    ocr_serial = b.text.strip()
            row_text = [_cell_text(c) for c in cells]
            expect = str(missing_vals[k])
            if ocr_serial and ocr_serial != expect:
                log(f"补行: OCR 序号 '{ocr_serial}' 与推导值 '{expect}' 不一致，取推导值")
            if serial_col < n_cols:
                row_text[serial_col] = expect
            inserts.append((row_text,
                            (min(b.y1 for b in boxes), max(b.y2 for b in boxes))))
        for row_text, ry in reversed(inserts):   # 组内按槽位降序插到 r 前
            table.rows.insert(r, row_text)
            table.row_y.insert(r, ry)
        filled += len(inserts)
        log(f"补行: 序号 {missing_vals[0]}-{missing_vals[-1]} 区域恢复 "
            f"{len(inserts)}/{m} 行")
    if filled:
        log(f"补行完成: 共补充 {filled} 行，表格 {len(table.rows)} 行")
    return filled
