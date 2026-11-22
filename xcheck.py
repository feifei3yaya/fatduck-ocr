# -*- coding: utf-8 -*-
"""交叉校验层：整图 OCR 重建值 vs 局部放大复识的两方核对。

背景：早期主链路用云端视觉模型做整表识别，视觉曾把运单号 YT1234567890123
误识为 YT1234567890120（编辑距离 3），而规则校正要求编辑距离恰为 1，够不着。
现主链路已改为纯本地 OCR，校验层随之改为两方核对：

  - 第一方（基准）：整图 OCR + 几何重建得到的单元格值；
  - 第二方（独立意见）：把该单元格区域裁剪放大后重新 OCR 的结果。

整图 OCR 是在缩放后的整幅图上识别，小字（运单号）最易误识；裁剪放大后像素
密度提升，识别更可靠。因此在纯 OCR 下，编码/数字类高风险字段一律做一次
放大复识核对（整图缩放的小字误识往往与重建值"自洽"，只有放大才看得出来）：

  - 编码/数字类：放大复识置信度够高时以放大结果自动改写，并入待复核清单；
  - 文本字段：只标记待复核，不自动改写；
  - 放大复识失败/置信度不足：保持原样，绝不猜测。

纯 stdlib + 复用 ai_calibrate 的编码判定，不 import engine（避免循环依赖）。
"""
from __future__ import annotations

import difflib
import re
from collections import Counter

from ai_calibrate import _canonize, _parse_code

ZOOM_SCALE = 3.0        # 局部放大倍数
ZOOM_PAD = 4.0          # 放大区域外扩像素
MATCH_MIN_SCORE = 0.85  # 整图 OCR 框最低置信分
SIM_MIN = 0.55          # 无几何对齐时，文本相似度匹配下限
TEXT_SAME_RATIO = 0.9   # 文本字段视为"基本一致"的相似度下限
ZOOM_MIN_SCORE = 0.85   # 放大复识置信度低于此值不自动改写，只标记待复核
# 自动改写还要求"显著更有信心"：放大复识 >= 0.95 且比整图 OCR 高出 >= 0.02。
# 实测：运单号末位错读的场景是 0.99 vs 0.96（该改）；而放大裁剪把字符边缘
# 切掉导致的误读只有 0.90 vs 0.89（不能改，改了反而把对的改错）。
ZOOM_STRONG_SCORE = 0.95
ZOOM_MARGIN = 0.02
DEFAULT_MAX_ZOOM = 120  # 单图放大复识次数上限（覆盖约 120 个编码字段）


def _norm(s) -> str:
    """归一化：全角转半角、删除所有空白。不转大写。"""
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    out = []
    for ch in s:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            ch = chr(code - 0xFEE0)
        elif code == 0x3000:
            ch = " "
        if not ch.isspace():
            out.append(ch)
    return "".join(out)


def _key(s) -> str:
    return _norm(s).upper()


def _ratio(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def _edit_distance(a: str, b: str) -> int:
    if a == b:
        return 0
    if abs(len(a) - len(b)) > 4:
        return max(len(a), len(b))
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1,
                           prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _log(on_log, msg: str) -> None:
    if on_log:
        try:
            on_log(msg)
        except Exception:
            pass


def _is_code_like(v: str, f: str) -> bool:
    """高风险字段：_parse_code(v) 或 _parse_code(f) 非 None，或 6 位以上纯数字串。"""
    if _parse_code(v) is not None or _parse_code(f) is not None:
        return True
    for t in (v, f):
        k = _key(t)
        if k and len(k) >= 6 and k.isdigit():
            return True
    return False


def _pick(winner_key: str, *cands) -> str:
    """取第一个键等于 winner_key 的候选原文（保持 OCR 原始大小写）。"""
    for c in cands:
        if c and _key(c) == winner_key:
            return _norm(c)
    return winner_key


def _aligned_local(rows, local_tables):
    """找与待校验结果维度完全一致的本地表格（行数相同且逐行列数相同）。"""
    for tb in local_tables or []:
        tb_rows = getattr(tb, "rows", None)
        if not tb_rows or len(tb_rows) != len(rows):
            continue
        if all(len(lr) == len(r) for r, lr in zip(rows, tb_rows)):
            return tb
    return None


def _geometry(tb, rows):
    """从 TableBlock 提取 (col_x, row_y) 切片；维度不足或元素非 2 浮点元组返回 None。"""
    if tb is None:
        return None
    col_x = list(getattr(tb, "col_x", None) or [])
    row_y = list(getattr(tb, "row_y", None) or [])
    ncol = max((len(r) for r in rows), default=0)
    if len(col_x) < ncol or len(row_y) < len(rows):
        return None
    # row_y 可能长于 rows：_realign_header 只从行首裁掉假行，数据行对应
    # row_y 的尾部切片。头部切片会把每行 region 整体上移 len(row_y)-len(rows)
    # 个行带（815b 错位 10 行），放大复识读到别行内容导致 flag 虚增。
    row_y = row_y[-len(rows):]
    for item in col_x[:ncol] + row_y:
        if not (isinstance(item, (tuple, list)) and len(item) == 2):
            return None
        try:
            float(item[0]); float(item[1])
        except (TypeError, ValueError):
            return None
    return col_x[:ncol], row_y


def _bbox_of(b):
    if not isinstance(b, dict):
        return None
    bb = b.get("bbox")
    if not bb:
        return None
    try:
        vals = [float(v) for v in list(bb)[:4]]
    except (TypeError, ValueError):
        return None
    return tuple(vals) if len(vals) == 4 else None


def _center_in(bbox, region, pad=ZOOM_PAD) -> bool:
    if not bbox or not region:
        return False
    try:
        x1, y1, x2, y2 = bbox
        rx1, ry1, rx2, ry2 = region
    except (TypeError, ValueError):
        return False
    cx, cy = (x1 + x2) / 2.0, (y1 + y2) / 2.0
    return (rx1 - pad) <= cx <= (rx2 + pad) and (ry1 - pad) <= cy <= (ry2 + pad)


def _cell_ocr(ri, ci, v, boxes, geom):
    """第一方来源框。返回 (text|None, score, bbox|None)。

    几何模式：region=(col_x[c][0], row_y[r][0], col_x[c][1], row_y[r][1])。
      重建值 v 本身就来自区域内某个框，所以这里取「与 v 最匹配的单个框」——
      若把整区域所有框拼接起来，多框单元格会被误判成差异。找不到单框达标时
      才退回整区域拼接（说明该单元格由多个框拼成）。区域内无框时返回区域
      本身，供放大复识定位。
    无几何时按文本相似度找最接近的框。
    """
    if geom is not None:
        col_x, row_y = geom
        region = (float(col_x[ci][0]), float(row_y[ri][0]),
                  float(col_x[ci][1]), float(row_y[ri][1]))
        hits = [(b, _bbox_of(b)) for b in boxes if _center_in(_bbox_of(b), region)]
        hits = [(b, bb) for b, bb in hits if bb is not None]
        if hits:
            vk = _key(v)
            best, best_r = None, 0.0
            for b, bb in hits:
                r = _ratio(vk, _key(b.get("text", "")))
                if r > best_r:
                    best, best_r = b, r
            if best is not None and best_r >= SIM_MIN:
                return (_norm(best.get("text", "")),
                        float(best.get("score", 0) or 0), _bbox_of(best))
            items = sorted(hits, key=lambda t: (t[1][1], t[1][0]))
            text = "".join(_norm(b.get("text", "")) for b, _ in items).strip()
            bbox = (min(bb[0] for _, bb in items), min(bb[1] for _, bb in items),
                    max(bb[2] for _, bb in items), max(bb[3] for _, bb in items))
            score = max((float(b.get("score", 0) or 0) for b, _ in items), default=0.0)
            return (text or None), score, bbox
        if region[2] > region[0] and region[3] > region[1]:
            return None, 0.0, region
        return None, 0.0, None
    vk = _key(v)
    best, best_r = None, 0.0
    for b in boxes:
        t = _norm(b.get("text", ""))
        if not t:
            continue
        r = _ratio(vk, _key(t))
        if r > best_r:
            best, best_r = b, r
    if best is not None and best_r >= SIM_MIN:
        return (_norm(best.get("text", "")), float(best.get("score", 0) or 0),
                _bbox_of(best))
    return None, 0.0, None


def _same_origin(v, z):
    """放大复识 z 与格子值 v 是否同源（取框自证）。

    长编码格（v 数字 >=8 位）专用：复识结果与 v 无公共前缀 8 位、无公共
    后缀 6 位、且互不含对方 >=4 位后缀段时，说明放大复识读到的是别行/别格
    内容（取框无效），不应据此产生待复核 flag。
    短值格（数字 <8 位）无法可靠判定，一律视为同源（保守保留 flag）。
    """
    vd = re.sub(r"\D", "", str(v))
    zd = re.sub(r"\D", "", str(z))
    if len(vd) < 8:
        return True
    if not zd:
        return False          # 长编码格复识没读到任何数字：取框无效
    # 公共前缀 >=8：同码误读差异（真差异，放行保留 flag）
    n = 0
    for a, b in zip(vd, zd):
        if a != b:
            break
        n += 1
    if n >= 8:
        return True
    # 公共后缀 >=6：同码尾部一致（如丢前缀/首位）
    n = 0
    for a, b in zip(reversed(vd), reversed(zd)):
        if a != b:
            break
        n += 1
    if n >= 6:
        return True
    # 互含后缀段 >=4
    for seg in re.findall(r"\d+", str(z)):
        if len(seg) >= 4 and vd.endswith(seg):
            return True
    for seg in re.findall(r"\d+", str(v)):
        if len(seg) >= 4 and zd.endswith(seg):
            return True
    return False


def _classify(v: str, f: str):
    """返回 (cls, detail)。cls ∈ same|minor|confusable|diff|unknown。"""
    if not v or not f:
        return "unknown", "OCR 无对应文本"
    vk, fk = _key(v), _key(f)
    if vk == fk:
        return "same", ""
    if _is_code_like(v, f):
        if _canonize(vk) == _canonize(fk):
            return "confusable", "混淆字符等价（如 O/0、l/1、S/5）"
        d = _edit_distance(vk, fk)
        tol = 3 if max(len(vk), len(fk)) >= 8 else 1
        if d <= tol:
            return "confusable", f"编码编辑距离 {d}"
        return "diff", f"编码编辑距离 {d}"
    if vk.isdigit() and fk.isdigit():
        tol = 2 if max(len(vk), len(fk)) >= 6 else 1
        d = _edit_distance(vk, fk)
        if d <= tol:
            return "confusable", f"数字编辑距离 {d}"
        return "diff", f"数字编辑距离 {d}"
    r = _ratio(vk, fk)
    if r >= TEXT_SAME_RATIO:
        return "minor", f"文本相似度 {r:.2f}"
    return "diff", f"文本相似度 {r:.2f}"


def _zoom_texts(worker, image_path, bbox, on_log=None,
                scale=ZOOM_SCALE, pad=ZOOM_PAD, v=""):
    """第二方：clip=[x1-pad,y1-pad,x2+pad,y2+pad]（下界钳 0）→ worker.ocr。

    返回 (text|None, score)。要点：放大后检测框会重新划分，同一格里可能
    出现多个框——此时**不做拼接**（拼接会把相邻文字粘成一个"更长"的串，
    是最常见的误改来源），而是取与原值最相似的单框，其次取置信度最高的
    单框。
    """
    if bbox is None:
        return None, 0.0
    try:
        x1, y1, x2, y2 = [float(v) for v in bbox]
        clip = [max(0.0, x1 - pad), max(0.0, y1 - pad), x2 + pad, y2 + pad]
        boxes = worker.ocr(image_path, clip=clip, scale=scale)
    except Exception as exc:
        _log(on_log, f"局部放大识别失败（{exc}）")
        return None, 0.0
    if not boxes:
        return None, 0.0
    items = [(b, _bbox_of(b)) for b in boxes]
    items = [(b, bb) for b, bb in items if bb is not None]
    if not items:
        return None, 0.0
    if len(items) > 1:
        vk = _key(v)
        best, best_r = None, 0.0
        for b, bb in items:
            r = _ratio(vk, _key(b.get("text", "")))
            if r > best_r:
                best, best_r = b, r
        if best is not None and best_r >= SIM_MIN:
            return (_norm(best.get("text", "")),
                    float(best.get("score", 0) or 0))
        top = max(items, key=lambda t: float(t[0].get("score", 0) or 0))[0]
        return (_norm(top.get("text", "")), float(top.get("score", 0) or 0))
    items.sort(key=lambda t: (t[1][1], t[1][0]))
    text = _norm(items[0][0].get("text", ""))
    score = float(items[0][0].get("score", 0) or 0)
    return (text or None), score


def _plausible_replacement(old: str, new: str) -> bool:
    """校验「用 new 替换 old」是否合理，防止把垃圾值写进表格。

    成立条件（任一）：
      - 混淆字符全映射后完全等价（O/0、l/1 之类的读法差异）；
      - 长度几乎相同（差 <= 1）且编辑距离不超过 max(2, len//4)——
        运单号末位错读正是这个量级（编辑距离 3）。
    长度差得多（多框拼接、跨格粘连）一律判为不合理，只标记不自动改写。
    """
    o, n = _key(old), _key(new)
    if not o or not n or o == n:
        return False
    if _canonize(o) == _canonize(n):
        return True
    if abs(len(o) - len(n)) > 1:
        return False
    return _edit_distance(o, n) <= max(2, len(o) // 4)


def _code_row_overrides(rows, boxes, geom):
    """单号列按行序对齐：返回 {(ri,ci): (code, score, bbox)}。

    仅当某列是纯 YT 单号列、且该列纯单号框数量与单元格数量相等时启用：
    按 y 排序的单号框对齐按行排序的单号格。用于避免 row_y 边界把上一行
    的单号框当成当前行（3c25 的假“行错位”）。不修改识别数据，只改校验取值。
    """
    col_x, row_y = geom
    if not col_x or not row_y:
        return {}
    overrides = {}
    for ci, (x1, x2) in enumerate(col_x):
        cells = []
        for ri, row in enumerate(rows):
            if ci >= len(row):
                continue
            m = re.fullmatch(r"YT\d{13}", _norm(row[ci]))
            if m:
                cells.append((ri, m.group(0)))
        if len(cells) < 2:
            continue
        hits = []
        for b in boxes:
            bb = _bbox_of(b)
            m = re.fullmatch(r"YT\d{13}", _norm(b.get("text", "")))
            if not bb or not m:
                continue
            cx = (bb[0] + bb[2]) / 2.0
            cy = (bb[1] + bb[3]) / 2.0
            if x1 <= cx <= x2 and row_y[0][0] <= cy <= row_y[-1][1]:
                hits.append((cy, m.group(0), float(b.get("score", 0) or 0), bb))
        if len(hits) != len(cells):
            continue          # 有漏识别/多框，无法按行序唯一对齐，保持原逻辑
        hits.sort(key=lambda t: t[0])
        cells.sort(key=lambda t: t[0])
        for (ri, _), (_, code, score, bb) in zip(cells, hits):
            overrides[(ri, ci)] = (code, score, bb)
    return overrides


def cross_check(rows, ocr_boxes, local_tables, worker, image_path,
                on_log=None, max_zoom=DEFAULT_MAX_ZOOM, check_text=True,
                skip_header=False):
    """两方核对：整图 OCR 重建值（第一方） vs 局部放大复识（第二方）。

    返回 (corrections, flagged)：
      corrections: [{"row","col","old","new","reason"}] 已自动改写的高风险字段；
      flagged:     [{"row","col","rebuilt","ocr","zoom","reason"}] 待人工复核。
    skip_header：表头已按标准列名还原（与整图 OCR 原文必然不同），跳过第 0 行。
    """
    corrections, flagged = [], []
    if not rows:
        return corrections, flagged
    # 低分框仍用于定位；置信度只影响自动改写，不应让最可疑的结果跳过复核。
    boxes = list(ocr_boxes or [])
    aligned = _aligned_local(rows, local_tables)
    geom = _geometry(aligned, rows) if aligned is not None else None
    overrides = _code_row_overrides(rows, boxes, geom) if geom is not None else {}
    if overrides:
        _log(on_log, f"交叉校验：单号列按行序对齐 {len(overrides)} 格")
    _log(on_log, "交叉校验：按单元格几何定位，与整图 OCR 结果逐格比对"
         if geom is not None else
         "交叉校验：无法几何对齐，按文本相似度比对")

    # 1) 收集需要放大复识的单元格：优先级 0=编码/数字，1=文本
    pending = []
    for ri, r in enumerate(rows):
        if skip_header and ri == 0:
            continue
        for ci in range(len(r)):
            v = _norm(r[ci])
            if not v:
                continue
            override = overrides.get((ri, ci))
            if override:
                f, fscore, bbox = override
            else:
                f, fscore, bbox = _cell_ocr(ri, ci, v, boxes, geom)
            if f is None and bbox is None:
                continue          # 无文本框且没有单元格几何，无法定位放大区域
            f = f or ""
            if _is_code_like(v, f):
                # 编码/数字类高风险字段：整图缩放后最易误识，一律放大复识核对
                pending.append((0, ri, ci, v, f, fscore, bbox, "编码字段复核"))
                continue
            if not check_text or not f:
                continue          # 普通文本没有第一方对应框时不猜测，避免无效复核
            if fscore < MATCH_MIN_SCORE:
                pending.append((1, ri, ci, v, f, fscore, bbox,
                                f"整图 OCR 置信度仅 {fscore:.2f}"))
                continue
            if _key(v) == _key(f):
                continue          # 高置信文本字段两方一致，无需核对
            cls, detail = _classify(v, f)
            if cls == "minor":
                continue          # 文本仅轻微差异，不做复识
            pending.append((1, ri, ci, v, f, fscore, bbox, detail))
    pending.sort(key=lambda t: (t[0], t[1], t[2]))
    if pending:
        _log(on_log, f"交叉校验：{len(pending)} 个单元格需要局部放大复识核对")

    # 2) 逐个放大复识并比对（放大复识次数受 max_zoom 预算控制）
    zoom_used = 0
    skipped = 0
    suppressed = 0
    for pri, ri, ci, v, f, fscore, bbox, detail in pending:
        if zoom_used >= max_zoom:
            skipped += 1
            continue
        z, zscore = _zoom_texts(worker, image_path, bbox, on_log=on_log, v=v)
        zoom_used += 1
        if not z:
            flagged.append({"row": ri, "col": ci, "rebuilt": v, "ocr": f,
                            "zoom": "", "reason": f"放大复识无结果（{detail}）"})
            continue
        vk, zk = _key(v), _key(z)
        if zk == vk:
            continue          # 放大复识确认重建值，通过
        # 取框自证：长编码格的复识结果与原值不同源，说明放大复识读到的是
        # 别行/别格内容（取框无效），据此改写或标待复核都是污染，直接丢弃。
        if not _same_origin(v, z):
            suppressed += 1
            continue
        if _is_code_like(v, f):
            # 编码类：放大复识更可靠，但必须"显著更有信心"才允许推翻整图 OCR。
            # 实测反例：0720457368625 整图读对，放大裁剪反而读成 9720457369625，
            # 两者置信度只差 0.01（0.90 vs 0.89）——这种必须保留原值。
            if zscore < ZOOM_MIN_SCORE:
                reason = f"放大复识置信度仅 {zscore:.2f}，未自动改写"
            elif not _plausible_replacement(v, z):
                reason = "放大复识结果与原值差异过大，判定不可靠，未自动改写"
            elif zscore < ZOOM_STRONG_SCORE or zscore - fscore < ZOOM_MARGIN:
                reason = (f"放大复识置信度 {zscore:.2f} 未显著高于整图 OCR "
                          f"{fscore:.2f}，未自动改写")
            else:
                new = _pick(zk, z, f)
                corrections.append({"row": ri, "col": ci, "old": v,
                                    "new": new, "reason": detail})
                flagged.append({"row": ri, "col": ci, "rebuilt": v, "ocr": f,
                                "zoom": z,
                                "reason": f"已按放大复识自动改写（{detail}），建议复核"})
                _log(on_log, f"交叉校验改写：第{ri+1}行第{ci+1}列 {v} -> {new}"
                             f"（放大复识 {zscore:.2f} > 整图 {fscore:.2f}）")
                continue
            flagged.append({"row": ri, "col": ci, "rebuilt": v, "ocr": f,
                            "zoom": z, "reason": reason})
            _log(on_log, f"交叉校验标记：第{ri+1}行第{ci+1}列 {v} 与放大复识"
                         f"「{z}」不一致，{reason}")
            continue
        # 文本字段：只标记待复核，绝不自动改写
        flagged.append({"row": ri, "col": ci, "rebuilt": v, "ocr": f, "zoom": z,
                        "reason": f"文本字段两方不一致（{detail}），未自动改写"})
        _log(on_log, f"交叉校验标记：第{ri+1}行第{ci+1}列 重建「{v}」"
                     f"与整图「{f}」/放大「{z}」不一致")
    if skipped:
        _log(on_log, f"交叉校验：{skipped} 个编码字段超出放大复识预算"
                     f"（上限 {max_zoom}）未核对")
    if suppressed:
        _log(on_log, f"交叉校验：{suppressed} 个复识结果与原值不同源"
                     f"（取框无效，已丢弃）")
    return corrections, flagged
