# -*- coding: utf-8 -*-
"""编码类文本的规则级校准（纯本地、零成本、无任何模型调用）。

早期版本有「AI 级」校准（把疑似编码批量送给大模型改写），现已整体移除：
本模块只剩规则级家族共识，全程离线，不联网、不调用任何云端服务。

规则级校准：同一家族（相同字母前缀 + 相同位数）的文本互为模板。
  - 家族真实前缀 = 成员原始前导字母串的多数派（可分辨 genuine 字母前缀
    如 SOS 与混淆读法 505）；
  - 修复三类：tail 内混淆字母->数字；成员前缀 = 家族前缀 + 混淆字母额外段
    （如 ZTO -> ZT+0）；成员前缀 == 家族前缀。

模块纯 stdlib（json/os/re/collections），可同时被 venv worker 与系统 python
app 进程导入。
"""
import json
import os
import re
import shutil
import sys
import threading
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

CONFIG_DIR_NAME = "肥鸭文档转换"
CONFIG_FILE_NAME = "config.json"
# 旧版本用过的配置目录名：更名后自动迁移设置，避免用户重新同意隐私声明、
# 丢失上次的导出模式与保存路径。
LEGACY_CONFIG_DIR_NAMES = ()

# OCR 常见混淆：字母 -> 数字
CONFUSABLE_TO_DIGIT = {
    "O": "0", "o": "0",
    "l": "1", "I": "1", "|": "1",
    "Z": "2", "z": "2",
    "S": "5", "s": "5",
    "B": "8",
    "g": "9", "q": "9",
}

# 编码形态：仅字母数字与竖线，总长 6-32，前导 1-8 字母 + 4-24 数字类尾段。
# 备用形态处理"前缀字母被误识成数字"（YT -> 1T），仅在前导段以数字开头时启用，
# 不影响正常编码的切分结果。
_RE_CODE = re.compile(r"^[A-Za-z0-9|]+$")
_RE_ALPHA_TAIL = re.compile(r"^([A-Za-z]{1,8})([A-Za-z0-9|]{4,24})$")
_RE_ALPHA_TAIL_DIGIT_HEAD = re.compile(r"^([0-9]{1,2}[A-Za-z]{1,6})([A-Za-z0-9|]{4,24})$")
DIGIT_RATIO_MIN = 0.5


def _split_code(t: str):
    """切分前导段与尾段，返回 (alpha, tail) 或 None。"""
    m = _RE_ALPHA_TAIL.match(t) or _RE_ALPHA_TAIL_DIGIT_HEAD.match(t)
    return (m.group(1), m.group(2)) if m else None


# 常见承运商代码（字母型短列的家族共识只认这些，避免把数字列 0/1/2 互相"纠正"）
CARRIER_CODES = ("SF", "ZTO", "YTO", "STO", "JD", "EMS", "BEST", "UC", "HT", "DB", "ZT", "YT")

# 单元格清洗：行号粘在日期前（原表序号列与日期列之间无竖线，OCR 并成一格）。
# 序号最大到 200，行号取 1-3 位；只有后面紧跟完整日期才剥，纯日期格不受影响
# （"2026.10.29" 剥任何前缀后剩下的都不成年，lookahead 保证不会误剥）。
_RE_ROWNO_BEFORE_DATE = re.compile(
    r"^\d{1,3}[./\-]?(?=(?:19|20)\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2})"
)
# 单元格清洗：运单号粘在商品名称前（运单号列文字与商品名文字相近，det 并成一个框）。
# 前缀允许字母，也允许 1-2 位误识数字（如 ZTO 被读成 2T/20），以免
# "10720377374599可口可乐…"这类漏剥。要求后面紧跟中文，纯运单号格子不受影响。
_RE_LEADING_TRACKING = re.compile(
    r"^((?:R\d{2}T)?[A-Za-z0-9]{1,3}[A-Za-z0-9]{9,15})(?=[\u4e00-\u9fff])"
)
# 日期解析（日期列增强 + 列家族共识共用）。先严格（两个分隔符齐全，月/日可能
# 只剩一位），再宽松（丢了一个分隔符或全丢成紧凑，允许前粘 1-3 位行号——序号列
# 与日期列之间没有竖线时 OCR 把两格并成一格，如 "220061029" = 行号 2 +
# 20061029）。从左到右取第一个月/日合法的匹配，"2026.13.01" 这类越界形态
# 不会被当成日期。
_DATE_STRICT_FULL = re.compile(r"((?:19|20)\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})")
_DATE_LOOSE_FULL = re.compile(r"((?:19|20)\d{2})[.\-/]?(\d{2})[.\-/]?(\d{2})(?!\d)")
# 一位字段里能唯一补零的数字：3-9（原值只可能是 03-09 或被去掉前导零的 3-9）。
# 0/1/2 有歧义：月 "1" 可能是 01，也可能是 10/11/12 丢了第 2 位；日 "2" 可能
# 是 02，也可能是 20-29 丢了末位。这类残缺只能按"不完整"处理。
_UNAMBIG_ONE_DIGIT = "3456789"


def _field_complete(digits: str) -> bool:
    """字段信息是否完整：两位必然完整；一位只有 3-9 能唯一推断。"""
    return len(digits) == 2 or (len(digits) == 1 and digits in _UNAMBIG_ONE_DIGIT)


def parse_date(text):
    """解析单元格文本中的日期，返回 (norm, complete, raw) 或 None。

    norm      - "YYYY.MM.DD" 规范串；
    complete  - True 表示信息完整（月/日两位，或一位但属于 3-9 可唯一补零）；
                False 表示残缺（如 "2026.10.2" 丢了日末位、月只剩 "1" 时
                原值可能是 10/11/12），值只是补零推断，调用方只能用于列内
                共识比对，不得直接当确定值输出（否则会把 29 臆造成 02）；
    raw       - 命中的原始日期子串（保留行号/粘连前缀时做原位置替换用）。
    覆盖的形态：2026.10.29 / 2026-10-29 / 20261029 / 12026.10.29 /
    9/2006.10.29 / 2002026.10.29 / 220061029（行号+紧凑）/ 202610.29、
    2026.1029（丢一个分隔符）/ 1S/202610.29 / 2026.9.05（单位数月可补零）。
    """
    if text is None:
        return None
    t = str(text).strip()
    if not t:
        return None
    for rx in (_DATE_STRICT_FULL, _DATE_LOOSE_FULL):
        for m in rx.finditer(t):
            y, mo, d = m.group(1), m.group(2), m.group(3)
            if not (1 <= int(mo) <= 12 and 1 <= int(d) <= 31):
                continue                      # 月/日越界，换下一个匹配
            norm = f"{y}.{int(mo):02d}.{int(d):02d}"
            return norm, _field_complete(mo) and _field_complete(d), m.group(0)
    return None


def _one_digit_lost(partial: str, full: str) -> bool:
    """partial 是否为 full 删掉恰好一个字符的结果（OCR 丢了一位）。"""
    if len(full) - len(partial) != 1 or not partial:
        return False
    return any(full[:i] + full[i + 1:] == partial for i in range(len(full)))


def _incomplete_date_digits(text: str) -> str:
    """残缺日期候选的数字核心（用于跟多数派日期比对丢位），不像日期返回 ""。

    覆盖 parse_date 判残缺的值（"2026.10.2"）与纯数字 7-9 位年内串
    （"2026102"、行号+残缺 "22006102"）。混入字母/中文的一律不猜。
    """
    t = (text or "").strip()
    if not t:
        return ""
    r = parse_date(t)
    if r is not None:
        # 残缺值的数字核心取原始命中串（不补零）："2026.10.2" -> "2026102"，
        # 补零成 "20261002" 就跟 majority 的 8 位等长、丢位比对必然失配。
        return "" if r[1] else "".join(c for c in r[2] if c.isdigit())
    if not t.replace(".", "").replace("/", "").replace("-", "").isdigit():
        return ""
    digits = t.replace(".", "").replace("/", "").replace("-", "")
    for cut in (0, 1, 2, 3):                  # 前 0-3 位视为行号
        core = digits[cut:]
        if len(core) not in (7, 8) or core[:2] not in ("19", "20"):
            continue
        if len(core) == 8 and not (
            1 <= int(core[4:6]) <= 12 and 1 <= int(core[6:8]) <= 31
        ):
            continue      # 8 位还月/日越界（"20261301"），不是残缺，是读错/非日期
        return core
    return ""


# ---------------------------------------------------------------- 数据结构

@dataclass(eq=False)
class Member:
    """一个疑似编码文本。"""
    index: int
    text: str
    alpha: str = ""   # 原始前导字母串
    tail: str = ""    # 尾段原始串
    canon: str = ""   # 混淆全映射后的规范化串
    key: tuple = ()   # 家族键 = (canon 前导字母, canon 尾段位数)


@dataclass
class CalibStats:
    total: int = 0            # 参与校验的疑似编码数
    rule_fixed: int = 0       # 规则级修复数
    families: int = 0         # 家族数
    log: list = field(default_factory=list)


# ---------------------------------------------------------------- 配置

def config_dir():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, CONFIG_DIR_NAME)


def config_path():
    return os.path.join(config_dir(), CONFIG_FILE_NAME)


def _migrate_legacy_config() -> None:
    """新目录还没有配置时，从旧版本目录复制一份 config.json。"""
    try:
        if os.path.exists(config_path()):
            return
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        for name in LEGACY_CONFIG_DIR_NAMES:
            old = os.path.join(base, name, CONFIG_FILE_NAME)
            if os.path.isfile(old):
                os.makedirs(config_dir(), exist_ok=True)
                shutil.copy2(old, config_path())
                return
    except Exception:
        return


def load_config():
    _migrate_legacy_config()
    p = config_path()
    try:
        with open(p, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_config(cfg):
    try:
        os.makedirs(config_dir(), exist_ok=True)
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------- 形态解析

def _canonize(s):
    return "".join(CONFUSABLE_TO_DIGIT.get(c, c) for c in s)


def _parse_code(text):
    """判定文本是否为编码类；返回 Member 或 None。"""
    if not text or not isinstance(text, str):
        return None
    t = text.strip()
    if len(t) < 6 or len(t) > 32:
        return None
    if not _RE_CODE.match(t):
        return None
    split = _split_code(t)
    if split is None:
        return None
    alpha, tail = split
    # 尾段不能是纯字母（纯字母串更可能是单词而非编码）
    if not any(c.isdigit() or c in CONFUSABLE_TO_DIGIT for c in tail):
        return None
    digits = sum(1 for c in t if c.isdigit())
    if digits / len(t) < DIGIT_RATIO_MIN:
        return None
    canon = _canonize(t)
    canon_split = _split_code(canon)
    if canon_split is None:
        return None
    ca, ct = canon_split
    return Member(index=-1, text=t, alpha=alpha, tail=tail, canon=canon,
                  key=(ca, len(ct)))


# ---------------------------------------------------------------- 规则级家族共识

def _family_fix(members):
    """对一个家族应用规则共识。返回 (fixes:dict[Member->str], prefix:str, tie:bool)。"""
    fixes = {}
    cnt = Counter(m.alpha for m in members)
    prefix, top_n = cnt.most_common(1)[0]
    tie = top_n > 1 and len(cnt) > 1 and list(cnt.values()).count(top_n) > 1
    for m in members:
        if m.alpha == prefix:
            continue  # 前缀无歧义，tail 的混淆不在规则职责内
        extra = m.alpha[len(prefix):] if m.alpha.startswith(prefix) else None
        if extra is not None and extra and all(c in CONFUSABLE_TO_DIGIT for c in extra):
            # 成员前缀 = 家族前缀 + 混淆字母额外段，如 ZTO -> ZT + 0
            fixed = prefix + _canonize(extra) + m.tail
            if fixed != m.text:
                fixes[m] = fixed
        elif tie:
            continue  # 并列前缀，规则不猜，保持原样
    return fixes, prefix, tie


def _log(on_log, msg):
    if on_log:
        try:
            on_log(msg)
        except Exception:
            pass


# 小家族并入：同位数、前缀只差一个字符、大家族样本远多于小家族时，
# 认为小家族是大家族前缀的形近误识，整体并入大家族。
# 覆盖的盲区：交叉校验的"两方都读错"（整图 OCR 与放大复识都把 YT 读成 XT/NT/VT，
# 两方一致 → 不会被标记待复核）。约束刻意保守：
#   大家族 >= 5 个样本、小家族 <= 2 个样本，前缀编辑距离 <= 1，尾段位数相同。
MERGE_MIN_BIG = 5
MERGE_MAX_SMALL = 2


def _merge_tiny_families(families: dict) -> list:
    fixes = []
    bigs = [(key, members) for key, members in families.items() if len(members) >= MERGE_MIN_BIG]
    if not bigs:
        return fixes
    for key, members in families.items():
        if len(members) > MERGE_MAX_SMALL:
            continue
        prefix, digits = key
        best = None
        for bkey, _bmembers in bigs:
            bprefix, bdigits = bkey
            if bdigits != digits:
                continue
            if _edit_distance(prefix, bprefix) <= 1:
                if best is None or _edit_distance(prefix, bprefix) < _edit_distance(prefix, best[0]):
                    best = (bprefix, bkey)
        if best is None:
            continue
        for m in members:
            fixes.append((m, best[0] + m.text[len(prefix):]))
    return fixes


# ---------------------------------------------------------------- 主入口

def calibrate_texts(texts, on_log=None):
    """对 texts 列表中的编码类文本做规则级家族共识校准。

    返回 (new_texts, stats)。new_texts 与输入等长，未识别为编码的项原样返回；
    规则无法确定的项保持原文（绝不猜测）。
    """
    stats = CalibStats()
    if not isinstance(texts, (list, tuple)):
        return texts, stats
    texts = list(texts)
    new_texts = list(texts)

    # 1) 形态识别 + 家族分组
    families = defaultdict(list)
    for i, t in enumerate(texts):
        m = _parse_code(t)
        if m is None:
            continue
        m.index = i
        families[m.key].append(m)
        stats.total += 1
    stats.families = len(families)
    if not families:
        return new_texts, stats

    # 2) 规则级修复
    rule_fixes = []   # (member, new_text)
    for key, members in families.items():
        fixes, prefix, tie = _family_fix(members)
        rule_fixes.extend(fixes.items())
    rule_fixes.extend(_merge_tiny_families(families))
    for m, new in rule_fixes:
        new_texts[m.index] = new
        stats.rule_fixed += 1
        _log(on_log, f"[规则] {m.text} -> {new}")
    return new_texts, stats


def calibrate_boxes(boxes, on_log=None):
    """便捷入口：直接校准 ocr_worker 返回的 box 列表（就地改 box['text']）。"""
    texts = [b.get("text", "") for b in boxes]
    new_texts, stats = calibrate_texts(texts, on_log=on_log)
    for b, t in zip(boxes, new_texts):
        b["text"] = t
    return stats


# ---------------------------------------------------------------- 单元格清洗（网格法）
#
# 网格法把表格切成规整的单元格后，仍有两类"格内粘连"要清洗：
#   a) 行号粘在日期前：原表序号列与日期列之间没有竖线，OCR 把两格读成一格；
#   b) 运单号粘在商品名称前：两列文字相近，det 把框并成一个，中心点落进商品名列。
# 另加字母型短列（承运商）的家族共识：被读成 TETo/hTo/ITo/TO 时按多数派纠正。
# 全部是确定性规则，只处理形态明确的场景，识别不了的一律保持原样（不猜测）。


def _edit_distance(a: str, b: str) -> int:
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def strip_row_number(text: str) -> str:
    """剥离开头粘在日期前的行号：'12026.10.29' -> '2026.10.29'。"""
    if not text:
        return text
    return _RE_ROWNO_BEFORE_DATE.sub("", text, count=1)


def strip_leading_tracking(text: str) -> str:
    """剥离开头粘住的运单号（仅当后面还有内容时）。"""
    if not text:
        return text
    m = _RE_LEADING_TRACKING.match(text)
    if not m:
        return text
    rest = text[m.end():]
    return rest if rest.strip() else text


def calibrate_alpha_column(texts: list, on_log=None) -> list:
    """字母型短列（承运商）的家族共识纠正。

    做法：列内每个字母型值向白名单投票（编辑距离最近且够近的码），得票最高的
    码作为该列真实值，再把距离够近的成员统一改写。误识值往往五花八门
    （承运商代码常被读成 TETo/hTo/ITo/TO/IMTO...），不能要求多数派本身正确，
    因此用「投票收敛」而不是「多数派即真理」。
    数字列/日期列的值不是字母，天然不投票，绝不会把 0 纠成 1。
    """
    if not texts:
        return texts
    nonempty = [t for t in texts if t and t.strip()]
    alpha = [t for t in nonempty if t.strip().upper().replace(" ", "").isalpha()]
    # 前提：该列以字母型值为主。数字列里个别数字被误识成字母（0->D/d）时，
    # 字母值占比很低，此时绝不能启用——否则会把整列"纠正"成某个承运商码。
    if len(nonempty) < 3 or len(alpha) < 3 or len(alpha) < 0.6 * len(nonempty):
        return list(texts)
    votes: Counter = Counter()
    for t in texts:
        if not t or not t.strip():
            continue
        u = t.strip().upper()
        if not u.replace(" ", "").isalpha():
            continue
        best, best_d = None, 99
        for c in CARRIER_CODES:
            d = _edit_distance(u, c)
            if d < best_d or (d == best_d and best is not None and len(c) > len(best)):
                best_d, best = d, c
        if best is not None and best_d <= max(1, len(u) // 2):
            votes[best] += 1
    if not votes:
        return list(texts)
    major, top = votes.most_common(1)[0]
    if top < 2:
        return list(texts)
    # 改写阈值比投票宽：投票要求"够近"，改写要求"不是别的码"。白名单码的精确
    # 匹配一律不动（其他真实承运商代码不能被合并）。
    # 注意改写阶段不要求值是纯字母——承运商误识常混入数字（ZTO -> 2TO/m0/1O/
    # Y7O），投票阶段靠"纯字母值"保证多数派干净即可。
    limit = max(2, len(major) + 1)
    # 长度下限：值太短（如空值被识别成 "0"）时宁可保留，不硬并进承运商码
    min_len = max(2, len(major) // 2)
    out = list(texts)
    for i, t in enumerate(texts):
        if not t or not t.strip():
            continue
        u = t.strip().upper()
        if u == major or u in CARRIER_CODES:
            continue
        if len(u) >= min_len and _edit_distance(u, major) <= limit:
            out[i] = major
            _log(on_log, f"[规则] 承运商 {t} -> {major}")
    return out


def calibrate_date_column(texts: list, on_log=None) -> list:
    """日期列家族共识：一列里 50% 以上是同一日期时，把形近误识改写成该日期。

    成员分三类：
      a) 完整日期（parse_date complete）-> 参与投票；与 majority 编辑距离
         <=2 的判为误识，规范成 majority；
      b) 残缺日期（complete=False，或纯数字年内残缺串如 "2026102"）-> 不投票；
         存在 majority 且其数字核心是 majority 丢掉一位的结果时，改写为
         majority（"2026.10.2" -> "2026.10.29"、"220061029" -> 2006.10.29）；
      c) 跨天混排/多数派不占优 -> 一律保持原样，绝不猜测。

    残缺值绝不单独补零输出（"2026.10.2" 补成 "2026.10.02" 是把 29 臆造成 02），
    只有列内多数派能证明丢的是哪一位时才改写。
    """
    if not texts:
        return texts
    nonempty = [t for t in texts if t and t.strip()]
    if len(nonempty) < 3:
        return list(texts)

    parsed = []        # (i, norm, complete)
    incomplete = {}    # i -> 残缺数字核心
    for i, t in enumerate(texts):
        r = parse_date(t)
        if r is not None:
            parsed.append((i, r[0], r[1]))
            if not r[1]:
                # 数字核心取原始命中串（不补零），供丢位比对用
                incomplete[i] = "".join(c for c in r[2] if c.isdigit())
            continue
        core = _incomplete_date_digits(t)
        if core:
            incomplete[i] = core
    hits = [p for p in parsed if p[2]]
    if len(hits) < 0.5 * len(nonempty):
        return list(texts)
    ranked = Counter(p[1] for p in hits).most_common(2)
    major, top = ranked[0]
    if len(ranked) > 1 and ranked[0][1] == ranked[1][1]:
        return list(texts)   # 两种日期并列（跨天混排），不猜
    if top < 0.5 * len(hits):
        return list(texts)
    major_digits = "".join(c for c in major if c.isdigit())

    out = list(texts)
    # 空格补齐：日期列里个别行被漏框（首行小字最常见，实测 14 行表 D1 行
    # 日期整格空）时，同列 50% 以上已是同一日期，就把空格补成 majority。
    # 这是"补残缺"不是"臆造"：该行真实值必属这一列当日登记表的多数派。
    # 跨天混排 major 为空、残缺占多数无完整多数派，均不触发。
    if major and top >= 2:
        for i, t in enumerate(texts):
            if t is not None and str(t).strip() == "":
                out[i] = major
                _log(on_log, f"[规则] 第{i + 1}行日期漏识别，按列多数派补 {major}")
    for i, norm, complete in parsed:
        if not complete:
            continue          # 残缺统一走下面的丢位比对，避免补零值参与距离判定
        if norm == major:
            # 日期本身认对了，但格式可能不规范（20261029 / 2002026.10.29），统一成规范格式
            if texts[i] != major:
                out[i] = major
                _log(on_log, f"[规则] 日期格式 {texts[i]} -> {major}")
            continue
        if _edit_distance(norm, major) <= 2:
            out[i] = major
            _log(on_log, f"[规则] 日期 {texts[i]} -> {major}")
    for i, core in sorted(incomplete.items()):
        if _one_digit_lost(core, major_digits):
            out[i] = major
            _log(on_log, f"[规则] 残缺日期 {texts[i]} -> {major}")
    return out


# ---------------------------------------------------------------- 领域词典纠错
#
# 面向商品名列：按 / 拆成语义段，把段纠到受控词表的 canonical 段上。
# 可选词表位于 assets/glossary_terms.json（按业务语料整理），
# 全程本地规则，无模型调用；找不到词表时自动降级为不处理。
# 防护（宁可不改，不可改错）：
#   - 段本身已是 canonical -> 不动；
#   - 规格段（PET450ML/CAN330ML）数字不同 -> 语义不同，绝不动；
#   - 最近与次近距离相同 -> 有歧义，不动；
#   - 距离 > 1（长段 > 长度/4）-> 超出 OCR 单字误识量级，不动。

_GLOSSARY_LOCK = threading.Lock()
_GLOSSARY_CACHE: dict | None = None
_SPEC_SEG = re.compile(r"^(PET|CAN|瓶|罐)?\d+(ML|ml|L|l|毫克|g|G)$")


def load_glossary() -> dict:
    """加载段级受控词表（找不到文件时返回空表，纠错自动降级为不处理）。"""
    global _GLOSSARY_CACHE
    with _GLOSSARY_LOCK:
        if _GLOSSARY_CACHE is not None:
            return _GLOSSARY_CACHE
        terms: dict = {"canonical": {}, "variants": {}}
        try:
            import json as _json
            for base in _glossary_paths():
                if base.is_file():
                    data = _json.loads(base.read_text(encoding="utf-8"))
                    terms["canonical"] = {
                        str(x["text"]): int(x.get("count", 1))
                        for x in data.get("segments", [])
                    }
                    terms["variants"] = {
                        str(k): str(v) for k, v in (data.get("variants") or {}).items()
                    }
                    break
        except Exception:
            terms = {"canonical": {}, "variants": {}}
        _GLOSSARY_CACHE = terms
        return terms


def _glossary_paths():
    """词表文件候选路径：环境变量 > exe 同级 > _MEIPASS > 源码目录。"""
    base = getattr(sys, "_MEIPASS", None)
    paths = []
    env = os.environ.get("IMG2XLSX_GLOSSARY")
    if env:
        paths.append(Path(env))
    try:
        paths.append(Path(sys.executable).resolve().parent / "assets" / "glossary_terms.json")
    except Exception:
        pass
    if base:
        paths.append(Path(base) / "assets" / "glossary_terms.json")
    paths.append(Path(__file__).resolve().parent / "assets" / "glossary_terms.json")
    return paths


_SPEC_NUM = re.compile(r"\d+")


def _is_spec_seg(seg: str) -> bool:
    return bool(_SPEC_SEG.match(seg.strip()))


# 宽松规格判断：以 PET/CAN 开头且含数字即视为规格段。
# "PET500MI"（末字母 I/L 误识）匹配不上精确规格正则，但仍应按"同容量"纠错。
_SPECISH_RE = re.compile(r"^(PET|CAN)", re.IGNORECASE)


def _looks_like_spec(seg: str) -> bool:
    s = seg.strip()
    return bool(_SPECISH_RE.match(s)) and bool(_SPEC_NUM.search(s))


def _spec_values(seg: str) -> tuple[str, str]:
    """规格段的容量数字候选：(原样, 易混字母归一后)。

    两条都要试：PET5OOML 归一后才等于 "500"（O 是 0 的误读）；
    PET500MI 归一后反而变成 "5001"（I 是 L 的误读），此时原样 "500" 才对。
    """
    raw = "".join(_SPEC_NUM.findall(seg))
    norm = "".join(CONFUSABLE_TO_DIGIT.get(c, c) for c in seg)
    fixed = "".join(_SPEC_NUM.findall(norm))
    return raw, fixed


def correct_segment(seg: str, terms: dict) -> str:
    """把单个语义段纠到受控词表；无法安全判定时原样返回。"""
    seg = seg.strip()
    if not seg:
        return seg
    canon = terms.get("canonical") or {}
    variants = terms.get("variants") or {}
    if seg in canon or seg in variants:
        return variants.get(seg, seg)
    if _looks_like_spec(seg):
        # 规格段（PET/CAN + 容量）：容量数字相同才是同一规格（PET500MI 与
        # PET500ML 只是末字母 I/L 误识），数字不同是不同产品（450 vs 500）
        # 一律不动。这样既修得住字符误识，也不会把 450ml 并成 500ml。
        value_raw, value_fixed = _spec_values(seg)
        if not (value_raw or value_fixed):
            return seg
        for c in canon:
            if not _looks_like_spec(c):
                continue
            c_raw, _c_fixed = _spec_values(c)
            # 原样或归一后任一与 canonical 的容量数字相同即视为同一规格；
            # 数字不同的规格（450 vs 500）归一后仍不同，不会被并到一起。
            if (value_raw and value_raw == c_raw) or \
                    (value_fixed and value_fixed == c_raw):
                return c
        return seg
    best, best_d, second_d = None, 10 ** 9, 10 ** 9
    for c in canon:
        if _is_spec_seg(c):
            continue
        d = _edit_distance(seg, c)
        if d < best_d:
            second_d, best_d, best = best_d, d, c
        elif d < second_d:
            second_d = d
    if best is None:
        return seg
    limit = 1 if len(seg) <= 4 else max(1, len(seg) // 4)
    if best_d > limit or second_d == best_d:
        return seg          # 超量程或有歧义 -> 不猜
    return best


def correct_product_name(text: str, terms: dict | None = None, on_log=None) -> str:
    """商品名整串纠错：按 / 拆段逐段纠错后原样拼回（保留原分隔符）。"""
    if not text or not text.strip():
        return text
    terms = terms if terms is not None else load_glossary()
    if not terms.get("canonical"):
        return text
    parts = re.split(r"(/|\|)", text)
    out = []
    for part in parts:
        if part in ("/", "|", ""):
            out.append(part)
            continue
        fixed = correct_segment(part, terms)
        if fixed != part.strip():
            _log(on_log, f"[词典] 商品名段 {part.strip()} -> {fixed}")
        out.append(fixed)
    return "".join(out)


def calibrate_numeric_columns(rows: list, on_log=None) -> list:
    """数量列数值共识（只在该列有足够多同位数数值时启用）。

    实拍表里常见三类数值错误：
      - 带汉字噪声（`的 1`）：去掉非数字，保留数字，不臆造位数；
      - 粘连（`242` 比列内常见位数多 1 位）：前缀是列内常见值（出现 ≥2 次）时取前缀；
      - 截断（`2` 比列内常见位数少 1 位）：列内以该数字开头的常见值恰好一个时补全。
    其余原样保留，避免把正常短值或其他列改坏。
    """
    if len(rows) < 8:
        return rows
    header = [str(h or "").strip() for h in (rows[0] if rows else [])]
    # 表头被截断成单字“数”时也要纳入；只认单字/数量词，避免把“包裹数”
    # 这类汇总行误当数据列。
    cols = [j for j, h in enumerate(header)
            if h in ("数", "数量", "件数")
            or any(k in h for k in ("数量", "件数", "短少", "实收"))]
    if not cols:
        return rows
    for ci in cols:
        header_text = header[ci]
        # 实收列允许 0/1/2 这类真实短值，不能按列共识补成 12/24；只去汉字噪声。
        allow_fill = not ("实收" in header_text
                          and not any(k in header_text for k in ("数量", "短少", "件数")))
        vals = [str(r[ci]).strip() for r in rows[1:] if ci < len(r)]
        digits = [v for v in vals if re.fullmatch(r"\d{1,3}", v)]
        # 表格列可能有一半单元格没识别出来，样本门槛不能按总行数设；
        # 5 个纯数字 + 70% 同位数已足够判断列的位数形态。
        if len(digits) < 5:
            continue
        mode_len, mode_count = Counter(len(v) for v in digits).most_common(1)[0]
        counts = Counter(digits)
        top_val, top_count = counts.most_common(1)[0]
        # 位数共识（70% 同位数）或值共识（最高频值占一半）任一成立即可。
        # 表格里常有 24 / 12 / 15 / 18 混合，位数共识可能只有 60% 多，
        # 但 24 这类最高频值仍足以判断 2 是 24 的截断。
        # 数量列允许补全时门槛放到 50%：末列常因一半单元格空着而只有 57% 同位；
        # 实收列 allow_fill=False，仍按 70% 只判断是否需要处理，不参与补全。
        by_len = mode_count / len(digits) >= (0.5 if allow_fill else 0.7)
        by_val = top_count / len(digits) >= 0.5
        if not (by_len or by_val):
            continue
        for ri in range(1, len(rows)):
            if ci >= len(rows[ri]):
                continue
            before = str(rows[ri][ci]).strip()
            if not before:
                continue
            stripped = re.sub(r"[^\d]", "", before)
            if stripped and stripped != before and len(stripped) <= mode_len:
                rows[ri][ci] = stripped
                _log(on_log, f"[规则] 第{ri}行第{ci + 1}列数量列去噪声：{before} -> {stripped}")
                before = stripped
            if not re.fullmatch(r"\d+", before):
                continue
            # 值共识优先：异常值只是最高频值的截断/多一位时，直接按该值修
            if allow_fill and by_val and len(before) == len(top_val) - 1 and top_val.startswith(before):
                rows[ri][ci] = top_val
                _log(on_log, f"[规则] 第{ri}行第{ci + 1}列数量列按列共识补全："
                             f"{before} -> {top_val}")
                continue
            if allow_fill and by_val and len(before) == len(top_val) + 1 and before.startswith(top_val):
                rows[ri][ci] = top_val
                _log(on_log, f"[规则] 第{ri}行第{ci + 1}列数量粘连拆分："
                             f"{before} -> {top_val}")
                continue
            if allow_fill and len(before) == mode_len + 1:
                prefix = before[:mode_len]
                if counts.get(prefix, 0) >= 2:
                    rows[ri][ci] = prefix
                    _log(on_log, f"[规则] 第{ri}行第{ci + 1}列数量粘连拆分："
                                 f"{before} -> {prefix}")
            elif allow_fill and len(before) == mode_len - 1:
                cands = [v for v in counts if len(v) == mode_len and v.startswith(before)]
                if len(cands) == 1:
                    rows[ri][ci] = cands[0]
                    _log(on_log, f"[规则] 第{ri}行第{ci + 1}列数量列按列共识补全："
                                 f"{before} -> {cands[0]}")
    return rows


_HEADER_KEYS = ("序号", "运单号", "包裹号", "日期", "承运商", "商品",
                "数量", "签收", "拒收", "破损", "实收", "单号")


def _dedupe_header_cell(text) -> str:
    """表头单元格去叠字：`状态 态` -> `状态`，`状态态` -> `状态`。

    只处理连续重复的单字/短词；数据行不调用，避免把真实的重复值改掉。
    长编码段（YT 单号/商品码等 >=8 位数字）先用占位符保护：连续重复数字是
    合法内容，去叠字会把 YT1234567890123 的 00/88 去掉变成残缺码（669d 实测）。
    """
    s = re.sub(r"\s+", " ", str(text or "")).strip()
    stash = []

    def _keep(m):
        stash.append(m.group(0))
        return f"\x00{chr(65 + len(stash) - 1)}\x00"

    s = re.sub(r"[A-Za-z]*\d{8,}[A-Za-z0-9]*", _keep, s)  # 长编码段占位
    s = re.sub(r"([^\x00])\s+\1", r"\1", s)  # 态 态 -> 态
    s = re.sub(r"([^\x00])\1", r"\1", s)     # 状态态 -> 状态（不碰占位符）
    s = re.sub(r"(?<!\S)(\S{2,4})\s+\1(?!\S)", r"\1", s)  # 商品 … 商品 -> 商品
    for i, v in enumerate(stash):
        s = s.replace(f"\x00{chr(65 + i)}\x00", v)
    return s


def clean_grid_rows(rows: list, on_log=None) -> list:
    """对网格法产出的二维表格做单元格清洗（就地返回新表）。"""
    if not rows:
        return rows
    n_cols = max(len(r) for r in rows)
    out = [list(r) + [""] * (n_cols - len(r)) for r in rows]

    # b) 商品名称列剥离前置运单号（除表头行外逐行处理）
    for ri in range(1, len(out)):
        for ci in range(n_cols):
            before = out[ri][ci]
            after = strip_leading_tracking(before)
            if after != before:
                out[ri][ci] = after
                _log(on_log, f"[规则] 第{ri}行第{ci + 1}列剥离前置运单号：{before} -> {after}")

    # b2) 商品名列做领域词典纠错（表头命中“商品名/名称/品名”才启用）
    header = out[0] if out else []
    product_cols = [j for j, h in enumerate(header)
                    if any(k in str(h) for k in ("商品名", "名称", "品名"))]
    if product_cols:
        terms = load_glossary()
        if terms.get("canonical"):
            for ri in range(1, len(out)):
                for ci in product_cols:
                    if ci >= len(out[ri]):
                        continue
                    fixed = correct_product_name(str(out[ri][ci]), terms, on_log=on_log)
                    if fixed != out[ri][ci]:
                        _log(on_log, f"[规则] 第{ri}行第{ci + 1}列商品名词典纠错："
                                     f"{out[ri][ci]} -> {fixed}")
                        out[ri][ci] = fixed

    # a) 日期列：先剥离行号，再按列做日期家族共识
    for ci in range(n_cols):
        col = [out[ri][ci] for ri in range(1, len(out))]
        if not col:
            continue
        # 识别口径放宽到 parse_date（完整/残缺/行号粘连/丢分隔符都算日期形态），
        # 让整列都是残缺日期的表也能进共识流程
        date_like = sum(1 for t in col if parse_date(t) is not None)
        if date_like < max(2, len([t for t in col if t]) // 2):
            continue
        for ri in range(1, len(out)):
            before = out[ri][ci]
            after = strip_row_number(before)
            if after != before:
                out[ri][ci] = after
        col = [out[ri][ci] for ri in range(1, len(out))]
        fixed = calibrate_date_column(col, on_log=on_log)
        for ri, val in enumerate(fixed, 1):
            out[ri][ci] = val
        # 共识也救不回来的残缺日期：保留原值（不臆造），但在日志里点名，
        # 避免无声地把错日期写进 Excel
        for ri in range(1, len(out)):
            v = str(out[ri][ci]).strip()
            if not v:
                continue
            r = parse_date(v)
            if (r is not None and not r[1]) or (r is None and _incomplete_date_digits(v)):
                _log(on_log, f"[规则] 第{ri}行第{ci + 1}列日期残缺未恢复，保留原值：{v}")

    # c) 承运商列共识
    for ci in range(n_cols):
        col = [out[ri][ci] for ri in range(1, len(out))]
        fixed = calibrate_alpha_column(col, on_log=on_log)
        for ri, val in enumerate(fixed, 1):
            out[ri][ci] = val

    # d0) 表头去叠字：只对得分最高的前 4 行（且确实像表头）生效，数据行不动
    header_idx, best = None, 0
    for i in range(min(4, len(out))):
        text = "".join(str(c or "") for c in out[i])
        score = sum(1 for k in _HEADER_KEYS if k in text)
        if score > best:
            header_idx, best = i, score
    if header_idx is not None and best >= 3:
        for ci in range(n_cols):
            out[header_idx][ci] = _dedupe_header_cell(out[header_idx][ci])

    # d) 数量列数值共识（样本足够才启用；不动文本列/其他列）
    calibrate_numeric_columns(out, on_log=on_log)

    # e) 去掉末尾整列全空的列（纸张候选常在右侧多切一个空列）。
    #    只对 <=8 列的小表生效，避免动 13 列大表已有的空列（第 1 张不得退化）。
    while 5 < len(out[0]) <= 8 and all(not str(r[-1]).strip() for r in out):
        for r in out:
            r.pop()
    return out
