# -*- coding: utf-8 -*-
"""可选的云端文本纠错（默认关闭，需用户在设置中手动开启）。

定位：**兜底**，不是主链路。只处理本地识别已标记为"待复核"的文本类单元格，
把单元格裁剪图送给 NVIDIA NIM 上的 step-5-preview 做二次确认。

设计约束（与"默认关闭的可选功能"这一合规选择对应）：
  1. 默认完全本地：开关不开时，本模块一个字节都不外发，也不发起任何请求。
  2. 密钥不落仓库：只从环境变量 IMG2DOC_NVIDIA_KEY 读取；用户也可填进
     config.json（界面会提示明文风险），但绝不写入任何会进分发包的文件。
  3. 最小外发：只发存疑单元格的裁剪图 + 当前 OCR 猜测，不发整张照片。
  4. 只改文本类单元格：含中日韩字符的格子（商品名等）。编码/数字/日期类
     字段（运单号、数量、日期）完全走本地规则与多尺度投票，不外发。
  5. 任何失败都静默降级：网络错误/超时/解析失败/置信度不足，一律保留本地
     结果并维持"待复核"标记，绝不因云端不可用而让转换失败。
"""
from __future__ import annotations

import json
import os
import re
import sys
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from ai_calibrate import load_config

# ---------------------------------------------------------------- 领域词表
#
# 本地整图 OCR 与几何重建同源，会犯一样的错（"芬达"两边都读成"开达"），
# 因此交叉校验标不出来的错例恰恰是最需要独立第三方确认的。
# 用从真实照片挖掘出的受控词表做"疑点检测"：某格的分段不在词表里，才送云端。
# 这样每张图只需外发少数几个真·存疑格，而不是全部文本格。

GLOSSARY_FILE = "glossary_terms.json"
_GLOSSARY_CACHE: dict | None = None


def _glossary_path() -> Path:
    """词表文件位置：打包后随 exe 解包目录的 assets/，开发期在项目 assets/。"""
    roots = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(Path(meipass) / "assets" / GLOSSARY_FILE)
    roots.append(Path(__file__).resolve().parent / "assets" / GLOSSARY_FILE)
    for r in roots:
        if r.is_file():
            return r
    return roots[-1]


def load_glossary() -> set:
    """加载受控词表的 canonical 段集合（加载失败返回空集=全部送，宁多发不漏发）。"""
    global _GLOSSARY_CACHE
    if _GLOSSARY_CACHE is not None:
        return _GLOSSARY_CACHE
    try:
        data = json.loads(_glossary_path().read_text(encoding="utf-8"))
        _GLOSSARY_CACHE = {s["text"] for s in data.get("segments", [])}
    except Exception:
        _GLOSSARY_CACHE = set()
    return _GLOSSARY_CACHE


def suspicious_cells(value: str) -> bool:
    """该文本格是否"可疑"。

    判据是"形近词"：某分段与受控词表里的 canonical 段编辑距离 <= NEAR_DIST，
    说明这段很可能就是词表里的那个词、只是被 OCR 读错了（PET50CM vs PET500ML、
    红提箱 vs 红提葡萄、FET24SML vs PET24SML）。与词表差得远的段视为新产品，
    不送云端——避免把整列都变成 API 调用。
    """
    if not value or not is_text_cell(value):
        return False
    terms = load_glossary()
    if not terms:
        return True
    segs = [s.strip() for s in re.split(r"[/|·•]+", value) if s.strip()]
    if not segs:
        return False
    for seg in segs:
        if seg in terms or len(seg) < 2:
            continue
        for term in terms:
            if abs(len(term) - len(seg)) > NEAR_DIST:
                continue
            if _edit_distance(seg, term) <= NEAR_DIST:
                return True
    return False


def _edit_distance(a: str, b: str) -> int:
    """Levenshtein 距离（短串，直接 DP）。"""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


# ---------------------------------------------------------------- 配置

DEFAULT_MODEL = "stepfun-ai/step-5-preview"
NVIDIA_BASE = "https://integrate.api.nvidia.com/v1/chat/completions"

CFG_ENABLED = "nv_fix_enabled"      # 开关（默认 False）
CFG_MODEL = "nv_fix_model"
CFG_KEY = "nvidia_api_key"
ENV_KEY = "IMG2DOC_NVIDIA_KEY"      # 首选密钥来源
ENV_MODEL = "IMG2DOC_NVIDIA_MODEL"

MIN_CONFIDENCE = 0.85   # 云端置信度下限；低于此值只保留本地结果
MAX_CELLS = 30          # 单张图最多外发多少个格子（成本/时延控制）
NEAR_DIST = 2           # 分段与词表 canonical 段的"形近"阈值
REQUEST_TIMEOUT = 150   # 秒；NIM 冷启动可达 1-2 分钟
MAX_TOKENS = 3000       # 推理模型思维链很长，2000 仍偶发被吃光返回空
CROP_SCALE = 2.0        # 小字单元格放大倍数
MAX_WORKERS = 4         # 并发请求数；NIM 可承受，串行 25 格要 4-8 分钟

_PROMPT = (
    "这是一张表格单元格的放大截图。请逐字准确识别其中的文字。"
    "严格只输出一行 JSON，不要思考过程、不要解释、不要代码块："
    '{"text":"识别到的原文","confidence":0.0}'
)

_JSON_RE = re.compile(r"\{[^{}]*\"text\"[^{}]*\}", re.S)
_CJK_RE = re.compile(r"[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7a3]")


def is_enabled() -> bool:
    """开关状态。未开启或没有可用密钥时都为 False（此时绝不发起请求）。"""
    cfg = load_config()
    if not cfg.get(CFG_ENABLED):
        return False
    return bool(get_key())


def get_key() -> str:
    """密钥来源：环境变量优先，其次 config.json（界面会提示明文风险）。"""
    return (os.environ.get(ENV_KEY, "").strip()
            or str(load_config().get(CFG_KEY, "") or "").strip())


def get_model() -> str:
    override = os.environ.get(ENV_MODEL, "").strip()
    if override:
        return override
    return (str(load_config().get(CFG_MODEL, "") or "").strip()
            or DEFAULT_MODEL)


def key_source() -> str:
    """给界面显示用：告诉用户当前密钥来自哪里（不返回密钥本身）。"""
    if os.environ.get(ENV_KEY, "").strip():
        return "environment"
    if str(load_config().get(CFG_KEY, "") or "").strip():
        return "config"
    return "missing"


def is_text_cell(value: str) -> bool:
    """是否属于"文本类"格子（含 CJK）。编码/数字/日期类一律不外发。"""
    return bool(value) and bool(_CJK_RE.search(value))


# ---------------------------------------------------------------- 调用


def _extract_json(text: str) -> dict | None:
    if not text:
        return None
    m = _JSON_RE.search(text.strip())
    if not m:
        return None
    try:
        d = json.loads(m.group(0))
    except Exception:
        return None
    return d if isinstance(d, dict) else None


def recognize_cell(image_b64: str, hint: str, log=None, timeout: int = REQUEST_TIMEOUT):
    """把单元格图送给云端模型。返回 (text, confidence)，失败返回 ("", 0.0)。

    任何异常都消化掉并降级，绝不向上抛——云端不可用不该让本地转换失败。
    """
    model = get_model()
    prompt = _PROMPT
    if hint:
        prompt += f'\n（本地 OCR 的猜测是 "{hint}"，如与图不符请以图为准）'
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url",
             "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}},
        ]}],
        "temperature": 0.05,
        "max_tokens": MAX_TOKENS,
    }
    req = urllib.request.Request(
        NVIDIA_BASE, data=json.dumps(payload).encode("utf-8"), method="POST",
        headers={"Authorization": f"Bearer {get_key()}",
                 "Content-Type": "application/json", "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8")[:120]
        except Exception:
            pass
        if log:
            log(f"云端纠错跳过：HTTP {exc.code} {detail}")
        return "", 0.0
    except Exception as exc:
        if log:
            log(f"云端纠错跳过：{type(exc).__name__}: {str(exc)[:80]}")
        return "", 0.0

    try:
        content = data["choices"][0]["message"].get("content") or ""
    except Exception:
        if log:
            log("云端纠错跳过：响应结构无法解析")
        return "", 0.0
    parsed = _extract_json(content)
    if parsed is None:
        if log:
            log(f"云端纠错跳过：返回内容不含可用 JSON（{str(content)[:60]}）")
        return "", 0.0
    text = str(parsed.get("text", "") or "").strip()
    try:
        conf = float(parsed.get("confidence", 0.0) or 0.0)
    except (TypeError, ValueError):
        conf = 0.0
    return text, conf


def fix_cells(items: list, worker, image_path, log=None) -> list:
    """对一批存疑单元格做云端二次确认。

    items: [{"row","col","rect","current"}]，rect 为 (x1,y1,x2,y2)。
    返回 [(row, col, new_text, confidence)]，仅包含确实要改写的条目。
    """
    if not is_enabled():
        return []
    picked = [it for it in items
              if is_text_cell(it.get("current", ""))
              and (it.get("suspect") or suspicious_cells(it.get("current", "")))]
    if not picked:
        return []
    if len(picked) > MAX_CELLS:
        if log:
            log(f"云端纠错：{len(picked)} 个存疑文本格超过单图上限 {MAX_CELLS}，"
                f"只处理前 {MAX_CELLS} 个")
        picked = picked[:MAX_CELLS]

    results = []
    # 先裁剪（本地子进程，快），再并发送云端；裁剪失败只跳过那一个格子
    prepared = []
    for it in picked:
        rect = it.get("rect")
        if not rect:
            continue
        try:
            image_b64 = worker.encode(
                image_path, clip=[float(v) for v in rect],
                scale=CROP_SCALE, quality=92, limit=2048,
            )
        except Exception as exc:
            if log:
                log(f"云端纠错跳过第{it['row'] + 1}行：裁剪失败（{exc}）")
            continue
        prepared.append((it, image_b64))
    if not prepared:
        return []

    def _one(pair):
        it, image_b64 = pair
        return it["row"], it["col"], it.get("current", ""), \
            recognize_cell(image_b64, it.get("current", ""), log=log)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        for row, col, current, (text, conf) in pool.map(_one, prepared):
            if not text or conf < MIN_CONFIDENCE:
                continue
            if text == current:
                continue      # 云端也认可本地结果，不必改
            if log:
                log(f"云端纠错：第{row + 1}行第{col + 1}列 "
                    f"「{current}」-> 「{text}」（置信度 {conf:.2f}）")
            results.append((row, col, text, conf))
    return results
