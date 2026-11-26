# -*- coding: utf-8 -*-
"""肥鸭文档转换 —— 桌面程序入口（本地服务 + 应用窗口，FastAPI + WebSocket）。

识别引擎（engine.py / ocr_worker.py / md2excel.py）与 Web 前端（web/ 目录）
配合工作：

  - 静态托管 web/ 与 assets/（favicon、品牌 logo、程序图标）
  - POST /api/convert      上传文件 + 模式 + 输出路径，启动转换任务
  - WebSocket /ws/{job_id} 日志 / 进度 / 文件状态 / 完成 / 错误事件流
  - POST /api/cancel       中断当前任务
  - GET|POST /api/config   设置记忆（last_mode / last_output / privacy_ack）
  - GET /api/privacy       版本与隐私声明全文
  - GET /api/download      下载生成的 Excel
  - GET /api/default-output 默认保存路径（桌面/肥鸭文档转换_时间戳.xlsx）

单实例：复用 core.ensure_single_instance（命名互斥体，只跑一个实例）。
仅监听 127.0.0.1，不对外暴露。
"""
from __future__ import annotations

import asyncio
import json
import math
import mimetypes
import os
import queue
import re
import secrets
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path
from urllib.parse import quote

import uvicorn
from fastapi import (FastAPI, File, Form, Request, UploadFile, WebSocket,
                     WebSocketDisconnect)
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles

try:
    # 读取系统剪贴板（Windows），供 /api/clipboard-paste 使用；写在这里是为了
    # 让 PyInstaller 静态收集 PIL.ImageGrab，避免打包后剪贴板功能失效。
    from PIL import ImageGrab
except Exception:  # noqa: BLE001 - 非 Windows 或缺少 Pillow 时接口返回不可用
    ImageGrab = None

# 共享常量与工具（core.py 不依赖 Tkinter，产品包不再携带 Tkinter 界面）
import core
import job_store
import nv_fix
from engine import EngineError, OcrWorker, build_exact_page
from ai_calibrate import parse_date
from md2excel import (_sheet_title, _triage_flags, convert_exact, convert_waybill,
                      flag_image_box)

ROOT = Path(__file__).resolve().parent
WEB_DIR = ROOT / "web"
ASSETS_DIR = ROOT / "assets"
HOST, DEFAULT_PORT = "127.0.0.1", 8765

APP_VERSION = core.APP_VERSION
# 版本检查：只查询 GitHub Releases 的最新版本号，不发送任何文件或个人信息。
# 仓库地址可通过环境变量覆盖（国内镜像/自建仓库同样适用）。
UPDATE_REPO = os.environ.get("FATDUCK_UPDATE_REPO", "feifei3yaya/fatduck-ocr").strip()
UPDATE_API_URL = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
UPDATE_CACHE = {"t": 0.0, "data": None}
UPDATE_CACHE_SECONDS = 6 * 3600


def _pick_port(preferred: int = DEFAULT_PORT, tries: int = 10) -> int:
    """优先用 8765；被其他程序占用时向后顺延，避免启动即失败。

    仍然只绑定 127.0.0.1；端口探测与真正监听之间存在极小竞态窗口，但
    单实例互斥体已保证本程序不会自己抢自己的端口。
    """
    for candidate in range(preferred, preferred + max(1, tries)):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            try:
                probe.bind((HOST, candidate))
            except OSError:
                continue
            return candidate
    return preferred


PORT = _pick_port()

# 前端每 5s 发一次心跳；浏览器在后台时定时器会被限流到约 1 次/分钟，
# 因此超时阈值取 180s，既不会误杀最小化的窗口，也能在关窗后自动退出。
IDLE_EXIT_SECONDS = int(os.environ.get("IMG2DOC_IDLE_EXIT", "180"))
FIRST_PING_TIMEOUT = int(os.environ.get("IMG2DOC_FIRST_PING", "300"))
# OCR 子进程处理大图会持续占用内存，长批次必须周期性重启，避免后段 OOM 被误判为空页。
WORKER_RESTART_EVERY = max(1, int(os.environ.get("IMG2DOC_WORKER_BATCH", "15")))
LAST_PING = {"t": 0.0, "seen": False}


def _is_worker_memory_error(exc: BaseException) -> bool:
    text = str(exc).lower()
    return any(token in text for token in ("unable to allocate", "out of memory", "memoryerror", "内存不足"))


def _is_worker_crash(exc: BaseException) -> bool:
    """OCR 子进程崩溃 / 管道断开：重启引擎后大概率能恢复，值得再试一次。"""
    text = str(exc)
    return any(token in text for token in ("OCR 工作进程无响应", "OCR 工作进程未运行", "通信中断"))


# 判据（供 done 事件的 failed_items[].retryable）：
#   只有“失败原因指向文件本身”才判为不可重试——路径不存在、无法打开/读取、
#   格式不支持、文件损坏/解码失败这类确定性错误，重传同一文件结果必然一样，
#   需要用户换文件；其余（OCR 进程无响应/崩溃、内存不足、超时、通信中断、
#   图片预处理或编码失败等）属于引擎/运行环境的一次性状态，重启引擎或重传
#   通常即可恢复，判为可重试。取消（已取消）不会进入 failed_items。
_FILE_FAULT_TOKENS = (
    "无法打开", "无法读取", "打不开", "读不出", "不存在", "找不到",
    "文件损坏", "已损坏", "不支持的格式", "不支持的图片", "不支持的文件", "格式不支持",
    "解码失败", "不是有效的", "不是图片", "空文件",
)


def _failure_retryable(reason: str) -> bool:
    """该失败能否通过重传/重试解决（判据见上方 _FILE_FAULT_TOKENS 注释）。"""
    text = str(reason or "")
    return not any(token in text for token in _FILE_FAULT_TOKENS)


# 表格输出方式：per_image = 一图一表；merged = 多图合并成一张表（带来源图片列）
LAYOUTS = ("per_image", "merged")


def _normalize_layout(value) -> str:
    text = str(value or "").strip().lower()
    return text if text in LAYOUTS else "per_image"



def _quality_denominator(tables, rows) -> tuple[int, str]:
    """单号率分母：优先“序号列跨度”，不可信时退回表行数。

    表行数会把表头块、商品条码续行、签字行也算进分母（6320eb7a 45 个数据行
    被算成 48 行）。序号列跨度 max-min+1 更接近真实数据行数。只有序号列
    基本连续、覆盖足够时才采用；否则退回表行，避免把数量列/条码列误当序号。

    返回 (denominator, source)，source 形如 "序号跨度 94–124" / "表行"。
    """
    if rows <= 0 or not tables:
        return rows, "表行"
    main = max(tables, key=lambda t: len(getattr(t, "rows", None) or []))
    tbl = getattr(main, "rows", None) or []
    if not tbl:
        return rows, "表行"
    ncols = max(len(r) for r in tbl)
    candidates = [0]
    # 表头不一定在第一行（前面常有标题/汇总行），前 6 行找“序号”列
    for r in tbl[:min(6, len(tbl))]:
        for ci in range(min(len(r), ncols)):
            header = str(r[ci])
            if "序号" in header and ci not in candidates:
                candidates.append(ci)

    best = None
    for ci in candidates:
        vals = []
        for r in tbl:
            cell = str(r[ci] if ci < len(r) else "").strip()
            m = re.match(r"^[^\d]{0,4}(\d{1,4})\b", cell)
            if m:
                v = int(m.group(1))
                if 1 <= v <= 5000:
                    vals.append(v)
        if len(vals) < 5:
            continue
        uniq = set(vals)
        span = max(uniq) - min(uniq) + 1
        missing = span - len(uniq)
        dup = len(vals) - len(uniq)
        # 序号列要基本连续；跨度不能远大于表行数（多栏序号会虚增分母）
        if span < 5 or span > len(tbl) * 1.2:
            continue
        if missing > 2 and missing / span >= 0.10:
            continue
        if dup > max(2, len(uniq) * 0.1):
            continue
        if len(vals) / len(tbl) < 0.5:
            continue
        cand_score = len(vals) - missing * 2 - dup * 2
        if best is None or cand_score > best[0]:
            best = (cand_score, span, min(uniq), max(uniq))
    if best is None:
        return rows, "表行"
    _, span, lo, hi = best
    return span, f"序号跨度 {lo}–{hi}"


def _quality_core(tables, flag_items, meta) -> dict:
    """拍摄质量评分核心：只读现有识别结果，不改识别逻辑。

    同一套口径同时服务任务级（全批次）与单图级（单页）：调用方传入对应的
    tables / flag_items / quality_meta 即可，阈值、分母规则与原因文案完全一致。

    2026-11 按 20 张真实图重校，三条改动：
      1) 文档类型分流：全表没有 YT+13 位单号、且内容含商品条码/数量/签收类
         特征时，判为“非运单清单/报表”，跳过单号召回率判据，只看结构、
         flags、模糊/高光——非运单表本来就不该按单号数量打分。
      2) 模糊度改为窗口化（ocr_worker.image_quality_metrics）：96x96 块、
         步长 48、跳过空白块，取方差最高 20% 块的中位数，避免大片留白把
         整图平均拉低（73d301e7 整图 113.4 -> 窗口化 1025.6）。
         新尺度下的两条硬规则：
           blur < 1000：不得判“良好”（扣 10 分，等级上限“需注意”）；
           blur < 900 且行数 ≥ 10 且存在风险信号（flags ≥ 25% 或单号率
           < 85%）：升“建议重拍”。
      3) 每张图都必须有模糊度与高光数值；任一缺失时不再判“良好”，
         并明确写出“未采集”，避免无指标却给高分。

    运单类保留原有主判据（单号召回率 70%、纸张透视、待复核 25%、纸张
    面积 60%、倾角 8°）；非运单类按要求跳过单号召回率与纸张几何判据。
    """
    rows = sum(len(getattr(t, "rows", None) or []) for t in tables)
    cols = max((len(r) for t in tables for r in (getattr(t, "rows", None) or [])), default=0)
    codes = set()
    nonempty = 0
    text_parts = []
    for t in tables:
        for row in (getattr(t, "rows", None) or []):
            for cell in row:
                text = str(cell or "")
                text_parts.append(text)
                if text.strip():
                    nonempty += 1
                codes.update(re.findall(r"YT\d{13}", text))
    text_blob = "\n".join(text_parts)
    flag_count = len(flag_items or [])
    flag_ratio = flag_count / (rows * cols) if rows and cols else 0.0
    fill_ratio = nonempty / (rows * cols) if rows and cols else 0.0
    paper = bool(meta.get("paper"))
    area = meta.get("paper_area_ratio")
    angle = meta.get("paper_angle")
    blur = meta.get("blur_var")
    highlight = meta.get("highlight_ratio")
    grid_rows = meta.get("grid_rows") or 0
    coverage = min(1.0, rows / max(rows, grid_rows)) if rows else 0.0

    # 文档类型：没有 YT 运单号，且表格内容以商品条码/数量/签收为主。
    barcode_hits = len(re.findall(r"(?<!\d)69\d{11}(?!\d)", text_blob))
    report_hints = sum(1 for k in ("商品条码", "条码", "商品", "品名", "数量",
                                   "签收", "完好", "签字", "实收", "短少")
                       if k in text_blob)
    is_report = (not codes) and (barcode_hits >= 2 or
                                 (barcode_hits >= 1 and report_hints >= 3))
    doc_type = "非运单清单/报表" if is_report else "运单明细"

    # 单号率分母：运单类优先序号跨度；非运单类不参与单号率评分。
    denominator, denom_source = rows, "表行"
    if not is_report and rows:
        denominator, denom_source = _quality_denominator(tables, rows)
    code_rate = min(1.0, len(codes) / denominator) if denominator else 0.0

    score = 100
    reasons, suggestions = [], []
    if is_report:
        reasons.append("文档类型：非运单清单/报表（全表无 YT 运单号），"
                       "已跳过单号召回率判据")
        if fill_ratio < 0.35:
            score -= 15
            reasons.append(f"表格结构稀疏（非空单元格 {fill_ratio:.0%}，低于 35%）")
            suggestions.append("正对单据拍摄，保持表格四边完整入镜")
    else:
        if code_rate < 0.70:
            score -= 40
            reasons.append(f"单号召回 {len(codes)}/{denominator}"
                           f"（{code_rate:.0%}，按{denom_source}），低于 70%")
            suggestions.append("正对单据拍摄，保持纸张四边完整入镜")
            suggestions.append("避免阴影和手指遮挡，单号列尽量占满画面")
            suggestions.append("可补拍单号列特写")
        if paper and code_rate < 0.75:
            score -= 10
            reasons.append("使用了纸张透视矫正，说明拍摄倾斜或背景复杂")
    if flag_ratio >= 0.25:
        score -= 25
        reasons.append(f"待复核 {flag_count} 条，占 {flag_ratio:.0%}（≥25%）")
        suggestions.append("保持光线均匀，避免反光")
        suggestions.append("逐条核对黄底单元格，数量列可局部补拍")
    if not is_report:
        if area is not None and area < 0.60:
            score -= 10
            reasons.append(f"纸张占画面 {area:.0%}，小于 60%")
        if angle is not None and abs(angle) > 8:
            score -= 5
            reasons.append(f"纸张倾斜 {angle:.1f}°，超过 8°")

    # 图像质量指标：窗口化模糊度 + 全图高光占比。
    # 缺失指标时不允许判“良好”，防止出现“无指标但给评分”。
    if blur is None:
        score -= 10
        reasons.append("模糊度未采集，不能判“良好”")
        suggestions.append("重新拍摄或检查 OCR 进程是否正常")
    if highlight is None:
        score -= 10
        reasons.append("高光占比未采集，不能判“良好”")
    risk = flag_ratio >= 0.25 or (not is_report and code_rate < 0.85)
    blur_severe = (blur is not None and blur < 900 and rows >= 10 and risk)
    blur_block_good = False
    if blur_severe:
        reasons.append(f"明显偏糊（窗口化拉普拉斯方差 {blur:.0f}，低于 900），"
                       f"且行数 {rows} ≥10、存在风险信号，建议重拍")
        suggestions.append("拍摄时保持手机稳定，点击单据对焦后再按快门")
    elif blur is not None and blur < 1000:
        blur_block_good = True
        score -= 10
        reasons.append(f"图像偏糊（窗口化拉普拉斯方差 {blur:.0f}，低于 1000），"
                       f"等级上限为“需注意”")
        suggestions.append("拍摄时保持手机稳定，点击单据对焦后再按快门")
    if highlight is not None and highlight > 0.02:
        score -= 10
        reasons.append(f"高光占比 {highlight:.1%}（>2%），反光可能影响识别，建议避开顶光")
        suggestions.append("避开顶灯/闪光灯，调整角度让光斑移出单据")

    score = max(0, score)
    level = "良好" if score >= 85 else ("需注意" if score >= 60 else "建议重拍")
    if blur_severe:
        score = min(score, 55)
        level = "建议重拍"
    elif blur_block_good and level == "良好":
        # 硬性下限优先于分数：90 分也不允许落回“良好”，分数一并压到 84，
        # 避免出现“评分 90 / 等级需注意”的自相矛盾。
        score = min(score, 84)
        level = "需注意"
    if not reasons:
        if is_report:
            reasons.append(f"非运单清单/报表：结构 {rows}×{cols}，非空 {fill_ratio:.0%}，"
                           f"待复核 {flag_ratio:.0%}，模糊度 {blur:.0f}，未触发重拍阈值")
        else:
            reasons.append(f"单号召回 {len(codes)}/{denominator}"
                           f"（{code_rate:.0%}，按{denom_source}），"
                           f"待复核 {flag_ratio:.0%}，模糊度 {blur:.0f}，未触发重拍阈值")
    if not suggestions:
        suggestions.append("保持当前拍法；如有表头噪声，建议补拍表头特写")

    return {
        "level": level,
        "score": score,
        "reasons": reasons,
        "suggestions": suggestions,
        "metrics": {
            "文档类型": doc_type,
            "数据行": rows,
            "列数": cols,
            "非空单元格占比": f"{fill_ratio:.0%}",
            "单号数": len(codes),
            "单号召回率": "不适用（非运单）" if is_report else f"{code_rate:.0%}",
            "待复核条数": flag_count,
            "待复核占比": f"{flag_ratio:.0%}",
            "使用纸张透视矫正": "是" if paper else "否",
            "纸张占画面": f"{area:.0%}" if area is not None else "未检出",
            "纸张倾角": f"{angle:.1f}°" if angle is not None else "未检出",
            "模糊度(窗口化拉普拉斯方差)": f"{blur:.1f}" if blur is not None else "未采集",
            "模糊度判据": "良好≥1000，重拍<900",
            "高光占比": f"{highlight:.1%}" if highlight is not None else "未采集",
            "网格行数": grid_rows or "未检出",
            "行覆盖率": f"{coverage:.0%}",
        },
    }


def _file_summary(pages, page_seconds) -> list[dict]:
    """结果卡「按文件」小表：每份文件一条，PDF 多页合并。

    只读取已识别页面的 tables/flags，不改识别与导出内容。
    单号数按文件内跨页去重；行数、待复核按页累加；列数取最大；耗时按页累加。
    """
    times: dict[tuple[str, int], float] = {}
    for item in page_seconds or []:
        try:
            times[(str(item.get("file", "")), int(item.get("page", 0)))] = float(
                item.get("seconds") or 0.0)
        except (TypeError, ValueError):
            continue
    groups: dict[str, dict] = {}
    for page in pages:
        name = getattr(page, "source_name", "") or getattr(page, "name", "")
        page_no = int(getattr(page, "source_page", 0) or 1)
        tables = [b for b in (getattr(page, "blocks", None) or [])
                  if getattr(b, "kind", "") == "table"]
        codes: set[str] = set()
        page_rows = 0
        for table in tables:
            table_rows = getattr(table, "rows", None) or []
            page_rows += len(table_rows)
            for row in table_rows:
                for cell in row:
                    codes.update(re.findall(r"YT\d{13}", str(cell or "")))
        entry = groups.setdefault(name or "（未知文件）", {
            "file": name, "pages": 0, "rows": 0,
            "waybills": 0, "flags": 0, "seconds": 0.0, "_codes": set(),
        })
        entry["pages"] += 1
        entry["rows"] += page_rows
        entry["_codes"].update(codes)
        entry["flags"] += len(getattr(page, "flags", None) or [])
        entry["seconds"] += times.get((name, page_no), 0.0)
    out = []
    for entry in groups.values():
        entry["waybills"] = len(entry.pop("_codes"))
        entry["seconds"] = round(entry["seconds"], 1)
        out.append(entry)
    return out


def _quality_tier(level: str) -> str:
    """对外档位：沿用现有三档语义，内部“良好”呈现为“可用”。"""
    return "可用" if level == "良好" else level


def _quality_page_entry(page) -> dict:
    """单张图评级：与任务级共用 _quality_core，阈值/分母/分流口径完全一致。"""
    tables = [b for b in (getattr(page, "blocks", None) or [])
              if getattr(b, "kind", "") == "table"]
    meta = getattr(page, "quality_meta", None) or {}
    flags = getattr(page, "flags", None) or []
    rep = _quality_core(tables, flags, meta)
    return {
        "file": getattr(page, "source_name", "") or getattr(page, "name", ""),
        "page": int(getattr(page, "source_page", 0) or 1),
        "sheet": getattr(page, "name", ""),
        "tier": _quality_tier(rep["level"]),
        "level": rep["level"],
        "score": rep["score"],
        "reasons": rep["reasons"],
        "suggestions": rep["suggestions"],
        "metrics": rep["metrics"],
    }


_TIER_SEVERITY = {"可用": 0, "需注意": 1, "建议重拍": 2}
_LEVEL_SEVERITY = {"良好": 0, "需注意": 1, "建议重拍": 2}


def _quality_label(entry: dict) -> str:
    """图上短标签（任务级 reasons/metrics 用它指明“是哪张图”）。"""
    text = entry.get("sheet") or Path(entry.get("file") or "").stem or "?"
    return str(text)[:8]


def _parse_pct(text) -> float | None:
    """解析 _quality_core 的百分比展示值（"68%" / "不适用（非运单）"）。"""
    if not isinstance(text, str) or not text.endswith("%"):
        return None
    try:
        return float(text[:-1]) / 100.0
    except ValueError:
        return None


def _parse_num(text, suffix: str = "") -> float | None:
    """解析 _quality_core 的数值展示值（"12.6°" / "5206.3" / "未采集"）。"""
    if not isinstance(text, str):
        return None
    body = text[:-len(suffix)] if suffix and text.endswith(suffix) else text
    try:
        return float(body)
    except ValueError:
        return None


def _quality_worst_page(entries) -> dict:
    """最差单图：先比档位，同档取分数最低（_quality_core 里档位与分数单调）。"""
    return max(entries, key=lambda e: (_LEVEL_SEVERITY.get(e["level"], 0), -e["score"]))


def _quality_task_codes(pages) -> int:
    """任务级单号数：跨页去重并集（与单图口径同源，分页重复不重复计数）。"""
    codes = set()
    for p in pages:
        for b in (getattr(p, "blocks", None) or []):
            if getattr(b, "kind", "") != "table":
                continue
            for row in (getattr(b, "rows", None) or []):
                for cell in row:
                    codes.update(re.findall(r"YT\d{13}", str(cell or "")))
    return len(codes)


def _quality_task_metrics(entries, pages=None) -> dict:
    """多页任务指标：逐页口径聚合，键名/类型与单页 metrics 完全一致。

    行/列/非空/单号数按全批汇总；单号召回率、待复核占比、纸张面积、倾角、
    模糊度、高光取最差页并在展示值里标注来源图；纸张透视任一页用过即“是”；
    行覆盖率按 Σ行数 / Σmax(行数, 网格行数) 汇总。
    """
    rows_total = sum(int(e["metrics"].get("数据行") or 0) for e in entries)
    cols = max((int(e["metrics"].get("列数") or 0) for e in entries), default=0)
    cells_total = 0
    nonempty_total = 0.0
    for e in entries:
        rows = int(e["metrics"].get("数据行") or 0)
        ncols = int(e["metrics"].get("列数") or 0)
        ratio = _parse_pct(e["metrics"].get("非空单元格占比"))
        cells_total += rows * ncols
        nonempty_total += (ratio or 0.0) * rows * ncols
    fill_ratio = nonempty_total / cells_total if cells_total else 0.0

    codes_total = (_quality_task_codes(pages) if pages
                   else sum(int(e["metrics"].get("单号数") or 0) for e in entries))
    is_waybill = any(e["metrics"].get("文档类型") == "运单明细" for e in entries)
    doc_type = "运单明细" if is_waybill else "非运单清单/报表"

    rates = []
    for e in entries:
        if e["metrics"].get("文档类型") != "运单明细":
            continue
        ratio = _parse_pct(e["metrics"].get("单号召回率"))
        if ratio is not None:
            rates.append((ratio, _quality_label(e)))
    rate_text = (f"{min(rates)[0]:.0%}（最差 {min(rates)[1]}）"
                 if rates else "不适用（非运单）")

    flag_ratios = [(_parse_pct(e["metrics"].get("待复核占比")) or 0.0,
                    _quality_label(e)) for e in entries]
    flag_ratio, flag_label = max(flag_ratios)
    flag_total = sum(int(e["metrics"].get("待复核条数") or 0) for e in entries)

    paper_labels = [_quality_label(e) for e in entries
                    if e["metrics"].get("使用纸张透视矫正") == "是"]
    paper_text = f"是（{'、'.join(paper_labels)}）" if paper_labels else "否"

    areas, angles, blurs, highlights = [], [], [], []
    grid_total, grid_seen = 0, False
    cover_rows, cover_base = 0, 0
    for e in entries:
        m = e["metrics"]
        label = _quality_label(e)
        area = _parse_pct(m.get("纸张占画面"))
        if area is not None:
            areas.append((area, label))
        angle = _parse_num(m.get("纸张倾角"), "°")
        if angle is not None:
            angles.append((abs(angle), label))
        blur = _parse_num(m.get("模糊度(窗口化拉普拉斯方差)"))
        if blur is not None:
            blurs.append((blur, label))
        highlight = _parse_pct(m.get("高光占比"))
        if highlight is not None:
            highlights.append((highlight, label))
        grid = m.get("网格行数")
        grid_val = int(grid) if isinstance(grid, int) else (
            int(grid) if isinstance(grid, str) and grid.isdigit() else None)
        rows = int(m.get("数据行") or 0)
        if grid_val is not None:
            grid_total += grid_val
            grid_seen = True
        cover_rows += rows
        cover_base += max(rows, grid_val or 0)
    coverage = min(1.0, cover_rows / cover_base) if cover_base else 0.0

    return {
        "文档类型": doc_type,
        "数据行": rows_total,
        "列数": cols,
        "非空单元格占比": f"{fill_ratio:.0%}",
        "单号数": codes_total,
        "单号召回率": rate_text,
        "待复核条数": flag_total,
        "待复核占比": f"{flag_ratio:.0%}（最差 {flag_label}）",
        "使用纸张透视矫正": paper_text,
        "纸张占画面": (f"{min(areas)[0]:.0%}（最差 {min(areas)[1]}）"
                       if areas else "未检出"),
        "纸张倾角": (f"{max(angles)[0]:.1f}°（最差 {max(angles)[1]}）"
                     if angles else "未检出"),
        "模糊度(窗口化拉普拉斯方差)": (f"{min(blurs)[0]:.1f}（最差 {min(blurs)[1]}）"
                                         if blurs else "未采集"),
        "模糊度判据": "良好≥1000，重拍<900",
        "高光占比": (f"{max(highlights)[0]:.1%}（最大 {max(highlights)[1]}）"
                     if highlights else "未采集"),
        "网格行数": grid_total if grid_seen else "未检出",
        "行覆盖率": f"{coverage:.0%}",
    }


def _quality_task_reasons(entries) -> list:
    """多页任务原因：首行汇总最差页与三档分布，随后逐张列出非“良好”页的原因。"""
    worst = _quality_worst_page(entries)
    counts = {}
    for e in entries:
        counts[e["tier"]] = counts.get(e["tier"], 0) + 1
    dist = " / ".join(f"{tier} {counts[tier]}"
                      for tier in ("可用", "需注意", "建议重拍") if counts.get(tier))
    lines = [f"共 {len(entries)} 张图，任务级按最差页 {_quality_label(worst)}"
             f"（{worst['tier']}，{worst['score']} 分）；单图分布：{dist}"]
    for e in entries:
        if e["level"] == "良好":
            continue
        prefix = f"{_quality_label(e)}："
        lines.extend(prefix + r for r in e["reasons"])
    return lines


def _quality_task_suggestions(entries) -> list:
    """多页任务建议：非“良好”页的建议去重保序；全良好时给通用建议。"""
    out, seen = [], set()
    for e in entries:
        if e["level"] == "良好":
            continue
        for s in e["suggestions"]:
            if s not in seen:
                seen.add(s)
                out.append(s)
    return out or ["保持当前拍法；如有表头噪声，建议补拍表头特写"]


def _quality_report(pages, flag_items) -> dict:
    """任务级拍摄质量自检：原有字段（level/score/reasons/suggestions/metrics）
    保持原样，新增 per_file 逐图评级（增量字段，前端按图定位重拍项）。

    任务级聚合（2026-11 修口径矛盾）：
      1) 分子分母同源：逐页调用 _quality_page_entry（每页用自己的
         _quality_denominator），任务级取最差页，不再用“全批并集单号数 /
         单页序号跨度”这种跨维度比值；
      2) flags 占比不被稀释：每页 flag_count/(该页 rows*cols)，任务级取最差页；
      3) 拍摄元数据取最差页：倾角最大、模糊最小、纸张占比最小、透视任一页
         用过即算用过，并在 metrics/reasons 里标注来源图；
      4) 不变式：任务级档位不优于任何一张单图。
    单页任务直接返回该页结果，保证逐图口径与 Excel 产出不变。
    """
    pages = list(pages)
    if not pages:
        report = _quality_core([], flag_items, {})
        report["per_file"] = []
        return report
    entries = [_quality_page_entry(p) for p in pages]
    if len(entries) == 1:
        only = entries[0]
        report = {
            "level": only["level"],
            "score": only["score"],
            "reasons": list(only["reasons"]),
            "suggestions": list(only["suggestions"]),
            "metrics": dict(only["metrics"]),
            "per_file": entries,
        }
        return report
    worst = _quality_worst_page(entries)
    report = {
        "level": worst["level"],
        "score": min(e["score"] for e in entries),
        "reasons": _quality_task_reasons(entries),
        "suggestions": _quality_task_suggestions(entries),
        "metrics": _quality_task_metrics(entries, pages),
        "per_file": entries,
    }
    task_severity = _TIER_SEVERITY[_quality_tier(report["level"])]
    worst_severity = max(_TIER_SEVERITY[e["tier"]] for e in entries)
    assert task_severity >= worst_severity, (
        f"任务级档位优于单图：task={report['level']} "
        f"worst={worst['tier']}")
    return report


app = FastAPI(title="肥鸭文档转换 Web")
app.mount("/web", StaticFiles(directory=str(WEB_DIR)), name="web")
app.mount("/assets", StaticFiles(directory=str(ASSETS_DIR)), name="assets")


@app.middleware("http")
async def _security_guard(request: Request, call_next):
    """统一安全边界：Host 白名单、同源校验、API token、首屏下发 cookie 与 CSP。

    静态资源（页面本身）放行，但页面响应会写入 token cookie；页面内的
    fetch / WebSocket 都是同源请求，会自动带上它。跨站页面既拿不到 cookie，
    也会因 Origin 校验被拒绝。
    """
    host = (request.headers.get("host") or "").lower()
    if host not in _ALLOWED_HOSTS:
        return JSONResponse({"detail": "非法的访问地址"}, status_code=421)
    origin = (request.headers.get("origin") or "").rstrip("/").lower()
    if origin and origin not in _ALLOWED_ORIGINS:
        return JSONResponse({"detail": "跨站请求被拒绝"}, status_code=403)
    if request.url.path.startswith("/api/"):
        token = (request.cookies.get(TOKEN_COOKIE)
                 or request.headers.get("x-img2doc-token", ""))
        if token != API_TOKEN:
            _startup_log(f"API 请求被拒绝（token 不匹配）：{request.url.path}")
            return JSONResponse({"detail": "未授权的本地请求"}, status_code=403)
    # 上传请求在解析 multipart 前先按 Content-Length 拒绝超限请求，
    # 避免几个 GB 的文件先写进临时目录再报错（磁盘被占满）。
    if request.url.path == "/api/convert" and request.method == "POST":
        try:
            declared = int(request.headers.get("content-length") or 0)
        except ValueError:
            declared = 0
        if declared > MAX_REQUEST_BYTES:
            return JSONResponse(
                {"detail": f"本次上传超过总量上限 {_MAX_TOTAL_TEXT}"},
                status_code=413,
            )
    response = await call_next(request)
    if response.headers.get("content-type", "").startswith("text/html"):
        response.set_cookie(TOKEN_COOKIE, API_TOKEN, httponly=True,
                            samesite="strict", path="/")
        response.headers.setdefault("Content-Security-Policy",
                                    _CONTENT_SECURITY_POLICY)
        # 首页必须每次回源：token 每次启动都会变，若命中 WebView2 缓存（或 304），
        # 上面这行 set_cookie 就不会执行，页面会带着上一次的旧 token 调 API 被 403，
        # 表现为首次运行弹窗显示 undefined、界面功能全失效。
        response.headers["Cache-Control"] = "no-store"
    return response



@app.get("/api/ping")
async def ping():
    """前端心跳：主窗口关闭后不再上报，本地服务据此自动退出。"""
    LAST_PING["t"] = time.time()
    LAST_PING["seen"] = True
    return {"ok": True}


# ---------------------------------------------------------------- 任务管理
class Job:
    def __init__(self) -> None:
        self.q: queue.Queue = queue.Queue()
        self.cancelled = False
        self.done = False
        self.finished_at = 0.0
        # 当前任务使用的 OCR 子进程：取消时直接 kill，避免阻塞在无超时的调用上
        self.worker: OcrWorker | None = None


JOBS: dict[str, Job] = {}
# 任务对象只在 WebSocket 消费完后清理；前端崩溃/刷新/从未连上时会漏掉。
# 这里再加一道 TTL 与总量上限，保证常驻托盘时 JOBS 不会无限增长。
JOB_TTL_SECONDS = 900
JOB_MAX_JOBS = 50


def _sweep_jobs() -> int:
    """清理已结束且超过 TTL 的任务；再按上限淘汰最旧的已结束任务。"""
    now = time.time()
    removed = 0
    for job_id, item in list(JOBS.items()):
        if item.done and now - (item.finished_at or now) > JOB_TTL_SECONDS:
            JOBS.pop(job_id, None)
            removed += 1
    if len(JOBS) > JOB_MAX_JOBS:
        finished = sorted(
            (job_id for job_id, item in JOBS.items() if item.done),
            key=lambda key: JOBS[key].finished_at,
        )
        for job_id in finished[:len(JOBS) - JOB_MAX_JOBS]:
            JOBS.pop(job_id, None)
            removed += 1
    return removed


def has_running_job() -> bool:
    """是否有未取消且未完成的任务（桌面端退出保护用）。

    JOBS 会随 WebSocket 断开被清掉，窗口收进托盘后前端若断连，
    这里还要能看到真正在跑的任务对象，避免退出确认被跳过。
    """
    return (any(not j.done and not j.cancelled for j in RUNNING_JOBS)
            or any(not j.done and not j.cancelled for j in JOBS.values()))


# 桌面托盘用的任务快照：不依赖 WebSocket 连接。主窗口收进托盘后前端可能被
# WebView2 降频，托盘仍要能显示「正在转换 x/y」、完成通知和取消入口。
JOB_STATE: dict[str, object] = {
    "running": False, "job_id": "", "mode": "", "files": 0,
    "processed": 0, "total": 0, "phase": "", "status": "idle",
    "message": "", "out_path": "", "waybills": 0, "total_pages": 0,
    "started_at": 0.0, "finished_at": 0.0,
}
_JOB_STATE_LOCK = threading.Lock()
# 正在跑的任务对象（不随 WebSocket 断开而丢失，供托盘「取消当前任务」使用）
RUNNING_JOBS: set[Job] = set()


def _state_update(**kw: object) -> None:
    with _JOB_STATE_LOCK:
        JOB_STATE.update(kw)


def job_snapshot() -> dict:
    """返回任务快照副本（托盘线程读取，避免与工作线程竞争）。"""
    with _JOB_STATE_LOCK:
        return dict(JOB_STATE)


def cancel_active_jobs() -> int:
    """托盘取消入口：给所有未完成的任务打取消标记，返回取消数量。"""
    count = 0
    for item in list(RUNNING_JOBS):
        if not item.done and not item.cancelled:
            item.cancelled = True
            _interrupt_job(item)
            count += 1
    return count


def _interrupt_job(job: Job) -> None:
    """取消时顺手结束正在阻塞的 OCR 调用，让取消在数秒内生效。"""
    worker = getattr(job, "worker", None)
    if worker is not None:
        try:
            worker.interrupt()
        except Exception:
            pass

# 最近完成任务的单号索引（应用内单号搜索用）。只在内存里保留最近几次任务，
# 不写额外文件；转换结果一旦被更晚的任务挤出，搜索会提示重新转换。
SEARCH_STORE: dict[str, dict] = {}
SEARCH_KEEP = 5

# 单文件输入上限：超过 100MB 的图片/PDF 直接拒绝（避免拖垮内存与识别时间）。
MAX_FILE_BYTES = 100 * 1024 * 1024
_MAX_FILE_TEXT = "100 MB"
# 单次任务配额：一次提交最多 200 个文件、请求体总量最多 2GB。
# 服务只监听本机，但托盘常驻期间仍要对误操作（整文件夹拖入）做兜底。
MAX_FILES = 200
MAX_TOTAL_BYTES = 2 * 1024 * 1024 * 1024
_MAX_TOTAL_TEXT = "2 GB"
# multipart 边界/表单字段会额外占少量字节，请求体上限留 32MB 余量。
MAX_REQUEST_BYTES = MAX_TOTAL_BYTES + 32 * 1024 * 1024
_COPY_CHUNK = 1024 * 1024
# 允许在本机文件系统上执行"预览/复制/删除"的文件白名单：
# 只有用户通过系统文件选择框选进来的文件才会进入，避免接口被拿来操作任意路径。
ALLOWED_FILES: set[str] = set()
SUPPORTED_EXTS = {"png", "jpg", "jpeg", "bmp", "tif", "tiff", "webp", "pdf"}

# ---------------------------------------------------------------- 安全边界
# 本地服务只监听 127.0.0.1，但任意网页仍可向其发请求（CSRF / DNS rebinding）。
# 因此启动时生成一次性 token，随首屏页面写入同源 cookie；API 与 WebSocket
# 必须携带该 token，同时校验 Host 与 Origin。
API_TOKEN = secrets.token_urlsafe(32)
TOKEN_COOKIE = "img2doc_token"
_ALLOWED_HOSTS = {f"{HOST}:{PORT}", f"localhost:{PORT}"}
_ALLOWED_ORIGINS = {f"http://{HOST}:{PORT}", f"http://localhost:{PORT}"}
_CONTENT_SECURITY_POLICY = (
    "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; "
    "img-src 'self' data: blob:; font-src 'self' data:; "
    f"connect-src 'self' ws://{HOST}:{PORT} ws://localhost:{PORT}; "
    "object-src 'none'; base-uri 'none'; frame-ancestors 'none'"
)
# 允许"用默认程序打开"的扩展名；可执行文件/脚本一律拒绝。
_OPENABLE_EXTS = {".xlsx", ".pdf", ".png", ".jpg", ".jpeg", ".bmp",
                  ".tif", ".tiff", ".webp"}
_FID_RE = re.compile(r"^[A-Za-z0-9_-]{1,64}$")
_UPLOAD_SUFFIX_RE = re.compile(r"^\.[A-Za-z0-9]{1,10}$")
# 本进程产出的 Excel 绝对路径白名单（下载/打开只认这里与任务历史）
ALLOWED_OUTPUTS: set[str] = set()


def _safe_fid(raw: object, index: int) -> str:
    """文件标识只允许字母、数字、下划线与短横线，禁止参与路径拼接。"""
    fid = str(raw) if raw not in (None, "") else f"i{index}"
    if not _FID_RE.fullmatch(fid):
        raise ValueError("文件标识不合法（仅允许字母、数字、下划线与短横线）")
    return fid


def _safe_upload_name(index: int, filename: str) -> str:
    """上传文件落盘名：序号 + 白名单后缀，原始文件名只用于界面展示。

    含路径分隔符 / 空字节的文件名一律拒绝（正常浏览器只会上传基名）。
    """
    raw = str(filename or "")
    if not raw or "\0" in raw or "/" in raw or "\\" in raw or raw != Path(raw).name:
        raise ValueError("上传文件名不合法")
    suffix = Path(raw).suffix.lower()
    if not _UPLOAD_SUFFIX_RE.fullmatch(suffix):
        suffix = ""
    return f"{index:04d}{suffix}"


def _contained_upload(temp_root: Path, filename: str) -> Path:
    """拼接后再次确认目标仍在临时目录内（双保险）。"""
    dst = temp_root / filename
    try:
        if not dst.resolve().is_relative_to(temp_root.resolve()):
            raise ValueError("上传路径非法")
    except OSError as exc:
        raise ValueError("上传路径非法") from exc
    return dst


def _save_upload_limited(upload: UploadFile, dst: Path,
                         limit: int = MAX_FILE_BYTES) -> int:
    """按块写盘并计数：超过单文件上限立即中止，不再等整份文件写完。

    返回写入字节数。异常时由调用方统一清理临时目录。
    """
    written = 0
    with dst.open("wb") as fh:
        while True:
            chunk = upload.file.read(_COPY_CHUNK)
            if not chunk:
                break
            written += len(chunk)
            if written > limit:
                fh.close()
                raise ValueError(f"{upload.filename} 超过单文件上限 {_MAX_FILE_TEXT}")
            fh.write(chunk)
    return written


def _expand_tiff_pages(source: Path, temp_root: Path, fid: str,
                       orig_name: str, log) -> list[Path] | None:
    """多页 TIFF 逐帧导出为临时 PNG；单页或失败时返回 None 走原逻辑。

    cv2.imdecode 只解首帧，扫描件常见的多页 TIFF 会静默漏页，这里先用
    Pillow 拆帧。Pillow 不可用/文件损坏时退回单页，不阻断任务。
    """
    if source.suffix.lower() not in (".tif", ".tiff"):
        return None
    try:
        from PIL import Image
    except Exception:
        return None
    try:
        with Image.open(source) as image:
            frames = int(getattr(image, "n_frames", 1) or 1)
            if frames <= 1:
                return None
            if frames > 200:
                log(f"{orig_name}：TIFF 共 {frames} 页，仅处理前 200 页")
                frames = 200
            page_dir = temp_root / f"tif_{fid}"
            page_dir.mkdir(parents=True, exist_ok=True)
            pages: list[Path] = []
            for index in range(frames):
                image.seek(index)
                dest = page_dir / f"p{index + 1:03d}.png"
                image.convert("RGB").save(dest)
                pages.append(dest)
        log(f"{orig_name}：多页 TIFF 已展开为 {len(pages)} 页")
        return pages
    except Exception as exc:  # noqa: BLE001 - 拆帧失败按单页继续
        log(f"{orig_name}：多页 TIFF 展开失败，按单页处理（{exc}）")
        return None


def cleanup_stale_temp_dirs(max_age_seconds: int = 24 * 3600) -> int:
    """清理上次进程崩溃遗留的上传/渲染临时目录（超过 24 小时才删）。"""
    root = Path(tempfile.gettempdir())
    now = time.time()
    removed = 0
    for prefix in ("img2xlsx_web_up_", "img2xlsx_web_", "img2xlsx_pdf_"):
        try:
            children = list(root.glob(prefix + "*"))
        except OSError:
            continue
        for child in children:
            try:
                if not child.is_dir():
                    continue
                if now - child.stat().st_mtime <= max_age_seconds:
                    continue
                shutil.rmtree(child, ignore_errors=True)
                removed += 1
            except OSError:
                continue
    return removed


def _resolve_path(raw: object) -> Path | None:
    try:
        return Path(str(raw)).resolve()
    except (OSError, RuntimeError, ValueError):
        return None


def _output_paths() -> set[Path]:
    """允许下载/打开的产物：本进程产出 + 本机任务历史里的 out_path。"""
    items = {_resolve_path(p) for p in ALLOWED_OUTPUTS}
    try:
        for meta in job_store.list_jobs(limit=job_store.MAX_KEEP):
            items.add(_resolve_path(meta.get("out_path") or ""))
    except Exception:
        pass
    return {p for p in items if p is not None}


def _default_output_dirs() -> set[Path]:
    """用户当前选择的保存目录：允许「打开目录」直接定位到它。"""
    dirs = {_resolve_path(_default_output().parent)}
    try:
        last = str(core.load_config().get("last_output") or "")
        if last:
            dirs.add(_resolve_path(Path(last).parent))
    except Exception:
        pass
    return {d for d in dirs if d is not None}


def _is_allowed_target(target: Path) -> bool:
    """路径白名单：产物本身、产物所在目录、保存目录、文件框选择的文件。"""
    resolved = _resolve_path(target)
    if resolved is None:
        return False
    allowed = {p for p in (_resolve_path(p) for p in ALLOWED_FILES) if p is not None}
    allowed |= _output_paths()
    if resolved in allowed or resolved in {p.parent for p in allowed}:
        return True
    return resolved in _default_output_dirs()


def _ws_authorized(websocket: WebSocket) -> bool:
    """WebSocket 握手校验：Host 合法、Origin 同源、cookie 带正确 token。"""
    host = (websocket.headers.get("host") or "").lower()
    if host not in _ALLOWED_HOSTS:
        return False
    origin = (websocket.headers.get("origin") or "").rstrip("/").lower()
    if origin and origin not in _ALLOWED_ORIGINS:
        return False
    return websocket.cookies.get(TOKEN_COOKIE) == API_TOKEN


_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def _run_powershell(script: str, target: str = "", timeout: int = 90,
                    interactive: bool = False) -> tuple[bool, str]:
    """跑一段 PowerShell；target 经环境变量传入，避免路径里的引号/空格拼接出错。"""
    if os.name != "nt":
        return False, "当前系统不支持"
    env = os.environ.copy()
    if target:
        env["IMG2DOC_FILE_TARGET"] = target
    args = ["powershell.exe", "-NoProfile", "-STA"]
    if not interactive:
        args.append("-NonInteractive")
    args += ["-Command", script]
    try:
        proc = subprocess.run(
            args, capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=timeout, env=env, creationflags=_NO_WINDOW,
        )
    except Exception as exc:  # noqa: BLE001 - 需要把原因反馈到界面
        return False, str(exc)
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()
        return False, detail[:300] or f"PowerShell 退出码 {proc.returncode}"
    return True, (proc.stdout or "").strip()


def _remember_search_index(job_id: str, result, out_path: Path,
                           images: dict[str, dict] | None = None,
                           mode: str = "full", target: str = "",
                           scanned: int = 0) -> None:
    entries = [dict(item) for item in (getattr(result, "waybill_entries", None) or [])]
    issues = [dict(item) for item in (getattr(result, "index_issue_items", None) or [])]
    SEARCH_STORE[job_id] = {
        "entries": entries,
        "issues": issues,
        "out_path": "" if mode == "find" else str(out_path),
        "images": images or {},
        "mode": mode,
        "target": _normalize_waybill(target) if target else "",
        "scanned": int(scanned or len(entries)),
        "time": time.time(),
    }
    while len(SEARCH_STORE) > SEARCH_KEEP:
        old_id = next(iter(SEARCH_STORE))
        SEARCH_STORE.pop(old_id)
        # 预览图生命周期跟随任务历史（job_keep），不跟内存索引（SEARCH_KEEP）：
        # 仍被历史保留的任务，重启或索引被挤出后还要靠磁盘缩略图定位原图。
        if job_store.load_job(old_id) is None:
            job_store.clear_previews(old_id)


def _make_search_previews(job_id: str, pages: list) -> dict[str, dict]:
    """为每页生成原图缩略图，供「单号查询」在原图上高亮定位。

    不直接暴露原始上传文件（临时目录会在任务结束/被更晚任务挤出时清理），
    而是落一份长边 3200 的 JPEG 缩略图；宽高记录原图尺寸，前端按百分比
    定位高亮框，缩略与原始比例一致，框的位置不受缩放影响。
    缩略图落在本机任务目录（job_store），历史回看与待复核定位共用。
    """
    out_dir = job_store.preview_dir(job_id)
    images: dict[str, dict] = {}
    for index, page in enumerate(pages):
        src = Path(str(getattr(page, "image_path", "") or ""))
        if not src.is_file():
            continue
        try:
            from PIL import Image

            with Image.open(src) as im:
                width, height = im.size
                thumb = im.convert("RGB")
            thumb.thumbnail((3200, 3200))
            out_dir.mkdir(parents=True, exist_ok=True)
            dst = out_dir / f"{index}.jpg"
            thumb.save(dst, "JPEG", quality=86)
            images[str(index)] = {
                "path": str(dst),
                "width": int(width),
                "height": int(height),
                "name": getattr(page, "source_name", "") or getattr(page, "name", ""),
                "page": int(getattr(page, "source_page", 0) or 1),
            }
        except Exception:  # noqa: BLE001 - 某页缩略失败不应影响转换结果
            continue
    return images



def _job_keep() -> int:
    """任务历史保留条数：设置里可调，越界回落到默认值。"""
    try:
        cfg = core.load_config()
        return max(1, min(job_store.MAX_KEEP,
                          int(cfg.get("job_keep") or job_store.DEFAULT_KEEP)))
    except (TypeError, ValueError):
        return job_store.DEFAULT_KEEP


def _default_output() -> Path:
    desktop = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"
    base = desktop if desktop.is_dir() else Path.cwd()
    return base / time.strftime("肥鸭文档转换_%Y%m%d_%H%M%S.xlsx")


def _unique_output(path: Path) -> Path:
    """目标已存在时自动另存为 "name (1).xlsx"、"name (2).xlsx"…

    与资源管理器重名文件的命名习惯一致：旧文件原样保留，新结果改名落盘，
    转换不需要用户介入。路径不存在时原样返回。
    """
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    for n in range(1, 1000):
        candidate = path.with_name(f"{stem} ({n}){suffix}")
        if not candidate.exists():
            return candidate
    # 兜底：同名前缀已占用 999 个时加时间戳，保证一定能写出
    return path.with_name(f"{stem} ({int(time.time())}){suffix}")


# ---------------------------------------------------------------- 图片命名
# 显示名（label）清洗复用 md2excel._sheet_title 同一套规则：非法字符 \ / ? * [ ] :
# 替换为下划线、去掉扩展名、截断到 28 字符（Excel 上限 31，预留重名后缀位）。
def _clean_label(name: str) -> str:
    return _sheet_title(str(name or ""), set())


# 磁盘文件名清洗：Windows 非法字符 \ / : * ? " < > | 与控制字符替换为下划线，
# 去首尾空白与结尾的点（资源管理器不允许），限长 120 字符。空名由调用方拒绝。
_DISK_ILLEGAL_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def _clean_disk_name(name: str) -> str:
    cleaned = _DISK_ILLEGAL_RE.sub("_", str(name or "")).strip()
    cleaned = cleaned.rstrip(". ").strip()
    return cleaned[:120].strip()


def _registered_file(target: Path) -> bool:
    """路径白名单复用：只认 pick-files / convert 登记过的文件（ALLOWED_FILES）。

    与 preview/file-op 同一机制；额外按 resolve 比对，兼容斜杠/大小写差异。
    """
    if str(target) in ALLOWED_FILES:
        return True
    resolved = _resolve_path(target)
    if resolved is None:
        return False
    return any(_resolve_path(p) == resolved for p in ALLOWED_FILES)


# 标题候选里的固定说明行/页脚/汇总词：这些出现在表格上方也不算表单标题
_TITLE_NOISE_RE = re.compile(
    r"备注|合计|小计|总计|序号|打印|页码|签收|制单|审核|第\s*\d+\s*页")
_TITLE_WAYBILL_RE = re.compile(r"(?:R\d{2}T)?YT\d{8,}", re.I)
_TITLE_PURE_NUM_RE = re.compile(r"[\d\s.\-/:：年月日]*")
_TITLE_MAX_LOOKBACK = 3  # 只向上回溯这么多个连续文字块，再远的文字不算标题


def _title_line_ok(line) -> bool:
    text = str(line or "").strip()
    if len(text) < 2 or text.endswith((":", "：")):
        return False
    if parse_date(text) is not None or _TITLE_PURE_NUM_RE.fullmatch(text):
        return False
    return not (_TITLE_WAYBILL_RE.search(text) or _TITLE_NOISE_RE.search(text))


def _page_title_text(page) -> str:
    """取该页的标题行：从首个表格块向上回溯最多 3 个连续文字块，返回第一条像标题的行。

    日期、纯数字、单号、备注/合计/打印页码等固定说明行都不是标题；回溯途中遇到
    非文字块（如另一张表）即停止。只读已有识别产物，不改识别逻辑。
    """
    blocks = getattr(page, "blocks", None) or []
    first_table = next((i for i, b in enumerate(blocks)
                        if getattr(b, "kind", "") == "table"), None)
    if first_table is None:
        return ""
    seen = 0
    for block in reversed(blocks[:first_table]):
        if getattr(block, "kind", "") != "text":
            break
        seen += 1
        if seen > _TITLE_MAX_LOOKBACK:
            break
        for line in getattr(block, "lines", None) or []:
            if _title_line_ok(line):
                return str(line).strip()

    # 几何回退路径的 blocks 只保留表格，表格外文字仍在 ocr_boxes 里：
    # 用首行表头同文 OCR 框的位置作为标题区上界（bbox 可能把标题区域也包进来）。
    # 找不到表头同文框时退回表格 bbox 顶边；不额外跑识别。
    for block in blocks:
        if getattr(block, "kind", "") != "table":
            continue
        boxes = []
        for box in getattr(page, "ocr_boxes", None) or []:
            if not isinstance(box, dict):
                continue
            b = box.get("bbox")
            text = str(box.get("text") or "").strip()
            if text and isinstance(b, (list, tuple)) and len(b) >= 4:
                cy = (float(b[1]) + float(b[3])) / 2
                boxes.append((cy, float(b[0]), text))
        rows = getattr(block, "rows", None) or []
        header_texts = {str(c).strip() for c in (rows[0] if rows else []) if str(c).strip()}
        header_ys = [cy for cy, _x, text in boxes if text in header_texts]
        bbox = getattr(block, "bbox", None)
        table_top = (min(header_ys) if header_ys else
                     (float(bbox[1]) if bbox and len(bbox) >= 4 else None))
        if table_top is not None:
            above = sorted((cy, x, text) for cy, x, text in boxes if cy < table_top)
            for _y, _x, text in above:
                if _title_line_ok(text):
                    return text
        break
    return ""


# 页脚「打印时间：2026-11-25 14:58」：只认带标签且日期完整的行，给文件日期/命名 {日期} 用
_PRINT_TIME_LABEL_RE = re.compile(r"打印\s*(?:时[间问]|日期)\s*[:：]?")
_CLOCK_RE = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}))?")
_PRINT_TIME_STRICT_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?$")


def _page_print_time(page) -> str:
    """返回 "YYYY-MM-DD HH:MM"（无时分则只有日期），找不到返回 ""。

    标签与日期需在同一行；日期走 parse_date 且只认 complete（残缺日期不当确定值）。
    先扫文字块（页脚通常是文字块），再扫表格单元格。
    """
    blocks = getattr(page, "blocks", None) or []
    lines: list[str] = []
    for block in blocks:
        if getattr(block, "kind", "") == "text":
            lines.extend(str(x) for x in (getattr(block, "lines", None) or []))
    for block in blocks:
        if getattr(block, "kind", "") == "table":
            for row in getattr(block, "rows", None) or []:
                for cell in row:
                    lines.append(str(cell or ""))
    # 几何回退路径 ocr_boxes 是全页文字；网格路径也包含矫正图边距内的页脚
    for box in getattr(page, "ocr_boxes", None) or []:
        if isinstance(box, dict) and box.get("text"):
            lines.append(str(box["text"]))
    for line in lines:
        m = _PRINT_TIME_LABEL_RE.search(line)
        if not m:
            continue
        rest = line[m.end():]
        parsed = parse_date(rest)
        if not parsed or not parsed[1]:
            continue
        y, mo, d = parsed[0].split(".")
        after = rest[rest.find(parsed[2]) + len(parsed[2]):]
        clock = _CLOCK_RE.search(after)
        if clock and int(clock.group(1)) <= 23 and int(clock.group(2)) <= 59:
            return f"{y}-{mo}-{d} {int(clock.group(1)):02d}:{clock.group(2)}"
        return f"{y}-{mo}-{d}"
    return ""


def _print_time_epoch(text) -> float | None:
    """前端回传的 "YYYY-MM-DD HH:MM" → 本地时间 epoch；格式/取值非法返回 None。"""
    m = _PRINT_TIME_STRICT_RE.match(str(text or "").strip())
    if not m:
        return None
    y, mo, d, hh, mm = m.groups()
    try:
        st = time.strptime(f"{y}-{mo}-{d} {hh or '00'}:{mm or '00'}", "%Y-%m-%d %H:%M")
        return time.mktime(st)
    except (ValueError, OverflowError):
        return None


def _page_waybills(page) -> set:
    """该页表格里的 YT 单号集合（与 _file_summary 同一提取口径）。"""
    codes: set[str] = set()
    for block in getattr(page, "blocks", None) or []:
        if getattr(block, "kind", "") != "table":
            continue
        for row in getattr(block, "rows", None) or []:
            for cell in row:
                codes.update(re.findall(r"YT\d{13}", str(cell or "")))
    return codes


def _attach_naming(pages, files, quality: dict, file_summary: list) -> None:
    """done 增量字段：per_file 每项加 label/suggested_name，file_summary 加 label。

    智能建议：标题首行清洗截断≤20 + _序号3位；否则该页单号恰好 1 个用该单号；
    否则回落默认 label。序号按任务内文件顺序 001 起。老字段一律不动。
    """
    order: dict[str, int] = {}
    label_by_orig: dict[str, str] = {}
    for idx, (_fid, _src, orig, label) in enumerate(files, 1):
        order.setdefault(orig, idx)
        label_by_orig.setdefault(orig, label)
    per_file = quality.get("per_file") or []
    print_by_orig: dict[str, str] = {}
    for i, page in enumerate(pages):
        orig = getattr(page, "source_name", "") or ""
        seq = order.get(orig, i + 1)
        default_label = getattr(page, "name", "") or label_by_orig.get(orig, "")
        title = _page_title_text(page)
        if title:
            suggested = f"{_clean_disk_name(title)[:20]}_{seq:03d}"
        else:
            codes = _page_waybills(page)
            suggested = next(iter(codes)) if len(codes) == 1 else default_label
        print_time = _page_print_time(page)
        if print_time and orig not in print_by_orig:
            print_by_orig[orig] = print_time
        if i < len(per_file):
            per_file[i]["label"] = default_label
            per_file[i]["suggested_name"] = suggested
            per_file[i]["print_time"] = print_time
    for entry in file_summary:
        orig = str(entry.get("file", "") or "")
        entry["label"] = label_by_orig.get(orig, Path(orig).stem)
        entry["print_time"] = print_by_orig.get(orig, "")


def _resolve_ocr_python(log):
    """定位 OCR 运行环境的 python（随包 runtime 优先）。"""
    return core.resolve_ocr_python(log)


_FAST_FIND_WORKER: OcrWorker | None = None
_FAST_FIND_WORKER_LOCK = threading.Lock()


def _get_fast_find_worker(python: Path, log) -> OcrWorker:
    """Reuse the warmed OCR process for targeted lookups between user jobs."""
    global _FAST_FIND_WORKER
    with _FAST_FIND_WORKER_LOCK:
        if (_FAST_FIND_WORKER is not None
                and not getattr(_FAST_FIND_WORKER, "_closed", False)
                and _FAST_FIND_WORKER.proc.poll() is None):
            return _FAST_FIND_WORKER
        if _FAST_FIND_WORKER is not None:
            try:
                _FAST_FIND_WORKER.close()
            except Exception:
                pass
        _FAST_FIND_WORKER = OcrWorker(python, log=log)
        return _FAST_FIND_WORKER


def _warm_fast_find_worker() -> None:
    """Load the lookup models while the user is adding files, not after Start."""
    try:
        python = _resolve_ocr_python(_startup_log)
        _get_fast_find_worker(python, _startup_log)
        _startup_log("单号查找引擎已预热")
    except Exception as exc:  # lazy job startup retries if warmup fails
        _startup_log(f"单号查找引擎预热失败，查找时会重试：{exc}")


# ---------------------------------------------------------------- 转换管线
def _run_job(job: Job, files: list[tuple[str, Path, str, str]], mode: str, out_path: Path,
             job_id: str = "", target: str = "", layout: str = "per_image",
             extra_temp_dirs: list[Path] | None = None) -> None:
    """工作线程：完整转换流程，事件经 job.q 推给 WebSocket。"""

    def emit(**kw) -> None:
        job.q.put(kw)
        # 同步维护托盘快照：进度、阶段、结果都不依赖前端 WebSocket 是否还在线
        kind = kw.get("type")
        if kind == "progress":
            _state_update(processed=int(kw.get("done") or 0),
                          total=int(kw.get("total") or 0))
        elif kind == "phase":
            _state_update(phase=str(kw.get("text") or ""))
        elif kind == "done":
            # 查找模式的 done 没有 waybills，用 matches 作为命中数
            _state_update(status="ok", out_path=str(kw.get("out_path") or ""),
                          waybills=int(kw.get("waybills") or kw.get("matches") or 0),
                          total_pages=int(kw.get("total_pages") or 0))
        elif kind == "error":
            _state_update(status="error", message=str(kw.get("message") or ""))
        elif kind == "cancelled":
            _state_update(status="cancelled")
        elif kind == "file_status":
            # 文件级状态推送频繁，仅用于计数兜底
            pass

    def log(text: str) -> None:
        emit(type="log", text=core.mask_sensitive(text))

    temp_root: Path | None = None
    worker: OcrWorker | None = None
    # PDF 渲染临时目录在拿到渲染结果后才登记；提前初始化，失败/取消时也能统一清理
    pdf_tmp_dirs: list[str] = []
    mode = core.normalize_mode(mode) or "full"
    layout = _normalize_layout(layout)
    task_started = time.time()  # done 事件的 elapsed 以整条任务（含建模/识别/写盘）计时
    RUNNING_JOBS.add(job)
    _state_update(running=True, job_id=job_id, mode=mode, files=len(files),
                  processed=0, total=0, phase="", status="running", message="",
                  out_path="", waybills=0, total_pages=0,
                  started_at=task_started, finished_at=0.0)
    try:
        if mode in core.MODE_NAMES:
            emit(type="progress", done=0, total=0, indeterminate=True)
            log("正在准备单号查找引擎…" if mode == "find"
                else "正在加载本地 OCR 模型（首次约 30 秒）…")
            python = _resolve_ocr_python(log)
            worker = (_get_fast_find_worker(python, log) if mode == "find"
                      else OcrWorker(python, log=log))
            job.worker = worker
            log("已启用本地 OCR 识别引擎")

            def restart_worker(reason: str) -> None:
                """长批次释放 OCR 子进程内存；页面结果已保存在 pages 中，不受影响。"""
                nonlocal worker
                emit(type="phase", text="正在释放识别内存并重新加载引擎…")
                log(f"重启 OCR 进程（{reason}）…")
                old = worker
                worker = None
                if old is not None:
                    try:
                        old.close()
                    except Exception:
                        pass
                worker = OcrWorker(python, log=log)
                job.worker = worker
                log("OCR 进程已重启")

            # 展开工作项：PDF -> 多页 PNG；多页 TIFF -> 逐帧 PNG；其余图片保持原样
            jobs: list[tuple[Path, str, str, str, int]] = []  # (图片路径, 文件id, 显示名, 源文件名, 页码)
            # 渲染/拆页失败按文件隔离：一个坏 PDF 不能让整批任务直接失败
            render_failures: list[tuple[str, str, str]] = []
            temp_root = Path(tempfile.mkdtemp(prefix="img2xlsx_web_"))
            for fid, source, orig_name, label in files:
                if source.suffix.lower() == ".pdf":
                    log(f"{orig_name}：正在渲染 PDF（300 DPI）…")
                    try:
                        rendered = worker.render_pdf(source, dpi=300)
                    except EngineError as exc:
                        message = str(exc)
                        reason = (message if message.startswith("PDF 渲染失败")
                                  else f"PDF 渲染失败：{message}")[:300]
                        render_failures.append((fid, orig_name, reason))
                        log(f"{orig_name}：{reason}")
                        # 渲染期间进程可能已崩溃，重启后再处理后续文件
                        if _is_worker_crash(str(exc)):
                            try:
                                restart_worker("PDF 渲染失败")
                            except Exception:
                                pass
                        continue
                    if not rendered:
                        reason = "PDF 渲染失败：未渲染出任何页面"
                        render_failures.append((fid, orig_name, reason))
                        log(f"{orig_name}：{reason}")
                        continue
                    first_dir = str(Path(rendered[0]["path"]).parent)
                    if first_dir not in pdf_tmp_dirs:
                        pdf_tmp_dirs.append(first_dir)
                    for page_no, page in enumerate(rendered, 1):
                        jobs.append((Path(page["path"]), fid,
                                     f"{label}-{Path(page['path']).stem}",
                                     orig_name, page_no))
                    log(f"{orig_name}：渲染出 {len(rendered)} 页")
                else:
                    tiff_pages = _expand_tiff_pages(source, temp_root, fid, orig_name, log)
                    if tiff_pages:
                        for page_no, page_path in enumerate(tiff_pages, 1):
                            jobs.append((page_path, fid,
                                         f"{label}-第{page_no}页",
                                         orig_name, page_no))
                    else:
                        jobs.append((source, fid, label, orig_name, 1))

            if not jobs:
                if render_failures:
                    detail = "；".join(f"{name}：{reason}" for _, name, reason in render_failures)
                    raise RuntimeError(f"所有文件都未能处理：{detail}"[:400])
                raise RuntimeError("没有可识别的图片，请检查文件或 PDF 是否有效。")

            if mode == "find":
                # Lookup mode skips table reconstruction and Excel writing. Images
                # are scanned in parallel; batches of 11 keep memory bounded and
                # allow progress/cancel updates between chunks.
                emit(type="progress", done=0, total=len(jobs), indeterminate=False)
                entries = []
                preview_pages = []
                failed = 0
                failed_map: dict[str, dict] = {}
                # PDF 渲染失败的文件不在 jobs 里，单独补记失败状态
                for fid, name, reason in render_failures:
                    failed += 1
                    failed_map.setdefault(fid, {"id": fid, "name": name, "reason": reason})
                    emit(type="file_status", id=fid, tag="fail", reason=reason)
                    log(f"{name}：{reason}")
                completed = 0
                find_batch_size = 11
                for start in range(0, len(jobs), find_batch_size):
                    if job.cancelled:
                        raise EngineError("已取消查找。")
                    chunk = jobs[start:start + find_batch_size]
                    for _, fid, _, _, _ in chunk:
                        emit(type="file_status", id=fid, tag="run")
                    emit(type="phase", text=f"正在快速查找 {start + 1}–{start + len(chunk)}/{len(jobs)}")
                    results = worker.fast_find([str(item[0]) for item in chunk], max_edge=1680, workers=8)
                    for (image_path, fid, label, source_name, page_no), result in zip(chunk, results):
                        completed += 1
                        if job.cancelled:
                            raise EngineError("已取消查找。")
                        if not result.get("ok"):
                            failed += 1
                            emit(type="file_status", id=fid, tag="fail")
                            reason = str(result.get("error") or "快速识别失败")[:200]
                            item = failed_map.setdefault(fid, {"id": fid, "name": source_name,
                                                               "reason": reason})
                            if reason not in item["reason"] and len(item["reason"]) < 160:
                                item["reason"] = f"{item['reason']}；{reason}"
                            log(f"{label} 查找失败：{reason}")
                        else:
                            emit(type="file_status", id=fid, tag="ok")
                            preview_pages.append(type("SearchPage", (), {
                                "image_path": str(image_path), "source_name": source_name,
                                "name": label, "source_page": page_no,
                            })())
                            page_hits = {}
                            for box in result.get("boxes") or []:
                                kind, matched = _match_waybill_text(target, box.get("text", ""))
                                if not kind:
                                    continue
                                if kind == "exact":
                                    status, note = "已找到", ""
                                elif kind == "tail":
                                    status = "待复核"
                                    note = (f"后四位 {_normalize_waybill(target)[-4:]} 一致，"
                                            "请核对前几位")
                                else:
                                    status, note = "疑似错读", ""
                                hit = {
                                    "waybill": matched or str(box.get("text", "")),
                                    "status": status,
                                    "note": note,
                                    "source": source_name, "page": page_no,
                                    "image_index": len(preview_pages) - 1,
                                    "box": box.get("bbox") or [],
                                    "score": box.get("score", 0),
                                    "suspicious": kind != "exact",
                                    "rotation": box.get("rotation", 0),
                                }
                                key = _normalize_waybill(target)
                                previous = page_hits.get(key)
                                # 同一页同一目标只保留最强的一条：已找到 > 待复核 > 疑似错读
                                rank = {"已找到": 0, "待复核": 1, "疑似错读": 2}
                                if previous is None or rank[status] < rank[previous["status"]]:
                                    page_hits[key] = hit
                            _upgrade_suspect_hits(worker, str(image_path),
                                                  list(page_hits.values()), target, log)
                            entries.extend(page_hits.values())
                        emit(type="progress", done=completed, total=len(jobs), indeterminate=False)
                if not preview_pages:
                    raise RuntimeError("图片没有成功识别，请换清晰图片后重试。")
                # Only matched pages need a thumbnail for the result view. This
                # avoids decoding/resizing every page after the OCR scan.
                matching_indexes = sorted({int(entry["image_index"]) for entry in entries})
                hit_pages = [preview_pages[i] for i in matching_indexes]
                index_map = {old: new for new, old in enumerate(matching_indexes)}
                for entry in entries:
                    entry["image_index"] = index_map[int(entry["image_index"])]
                images = _make_search_previews(job_id, hit_pages)
                _remember_search_index(job_id, type("SearchResult", (), {
                    "waybill_entries": entries, "index_issue_items": [],
                })(), Path(), images, mode="find", target=target, scanned=len(jobs))
                exact_count = sum(1 for item in entries if item["status"] == "已找到")
                review_count = sum(1 for item in entries if item["status"] == "待复核")
                suspect_count = sum(1 for item in entries if item["status"] == "疑似错读")
                if exact_count:
                    log(f"查找完成：完整命中 {exact_count} 处（无需复核）；"
                        f"后四位命中 {review_count} 处、疑似错读 {suspect_count} 处需复核。")
                elif review_count or suspect_count:
                    log(f"没有完整命中；后四位命中 {review_count} 处、疑似错读 {suspect_count} 处，"
                        "请对照原图核对前几位。")
                else:
                    log("快速识别没有找到匹配项；这不代表单号一定不存在，可用「图片转 Excel」进一步核对。")
                job_store.save_job(job_id, {
                    "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "mode": "find",
                    "out_path": "",
                    "query": target,
                    "files": [orig for _fid, _src, orig, _label in files],
                    "scanned": len(jobs),
                    "matches": len(entries),
                    "exact_matches": exact_count,
                    "review_matches": review_count,
                    "suspect_matches": suspect_count,
                    "failed_items": [dict(item, retryable=_failure_retryable(item.get("reason", "")))
                                     for item in failed_map.values()],
                }, keep=_job_keep())
                emit(type="done", mode="find", searchable=True, query=target,
                     matches=len(entries), exact_matches=exact_count,
                     review_matches=review_count, suspect_matches=suspect_count,
                     scanned=len(jobs),
                     failed=int(failed),
                     failed_items=[dict(item, retryable=_failure_retryable(item.get("reason", "")))
                                   for item in failed_map.values()],
                     elapsed=round(time.time() - task_started, 1))
                return

            emit(type="progress", done=0, total=len(jobs), indeterminate=False)
            pages = []
            page_seconds: list[dict] = []  # 每页识别耗时（结果卡「按文件」小表用）
            failed = 0
            failed_map: dict[str, dict] = {}  # 按文件去重，PDF 多页失败合并成一条
            processed_since_restart = 0

            def mark_failed(fid: str, source_name: str, reason: str) -> None:
                nonlocal failed
                failed += 1
                emit(type="file_status", id=fid, tag="fail")
                item = failed_map.get(fid)
                if item is None:
                    failed_map[fid] = {"id": fid, "name": source_name, "reason": reason}
                elif reason not in item["reason"] and len(item["reason"]) < 160:
                    item["reason"] = f"{item['reason']}；{reason}"

            # PDF 渲染失败的文件不在 jobs 里，单独补记失败状态
            for fid, name, reason in render_failures:
                mark_failed(fid, name, reason)
                log(f"{name}：{reason}")

            for index, (image_path, fid, label, source_name, page_no) in enumerate(jobs, 1):
                if job.cancelled:
                    raise EngineError("已取消转换。")
                if processed_since_restart >= WORKER_RESTART_EVERY:
                    restart_worker(f"已连续处理 {processed_since_restart} 页")
                    processed_since_restart = 0
                emit(type="file_status", id=fid, tag="run")
                emit(type="phase", text=f"正在识别 {index}/{len(jobs)}：{label}")
                started = time.time()
                try:
                    page = build_exact_page(
                        image_path, worker,
                        on_log=log, should_stop=lambda: job.cancelled,
                    )
                except EngineError as exc:
                    if "已取消" in str(exc):
                        raise
                    # 取消时我们会主动 kill 子进程，这里会表现为"无响应/通信中断"，
                    # 必须优先判定为取消，而不是走重启+重试。
                    if job.cancelled:
                        raise EngineError("已取消转换。")
                    # 内存不足、子进程崩溃、管道断开都属于"引擎状态坏了"，
                    # 重启后重试一次通常能恢复；其余错误直接记失败，避免拖慢整批。
                    if _is_worker_memory_error(exc) or _is_worker_crash(exc):
                        reason = "上一页内存不足" if _is_worker_memory_error(exc) else "OCR 进程异常退出"
                        log(f"{label} 识别中断（{reason}），重启 OCR 引擎后重试一次…")
                        restart_worker(reason)
                        processed_since_restart = 0
                        emit(type="phase", text=f"正在重试 {index}/{len(jobs)}：{label}")
                        try:
                            page = build_exact_page(
                                image_path, worker,
                                on_log=log, should_stop=lambda: job.cancelled,
                            )
                        except EngineError as retry_exc:
                            if "已取消" in str(retry_exc) or job.cancelled:
                                raise EngineError("已取消转换。")
                            mark_failed(fid, source_name, str(retry_exc)[:200])
                            log(f"{label} 重试仍失败：{retry_exc}")
                            continue
                    else:
                        mark_failed(fid, source_name, str(exc)[:200])
                        log(f"{label} 识别失败：{exc}")
                        continue
                # 本地管线对解码失败等确定性错误会返回空 ExactPage，而不是抛出
                # EngineError；这种页不能进入成功分支，否则会先发 tag=ok 和
                # “识别完成”，直到写 Excel 时才以任务级 error 收尾。
                if not getattr(page, "blocks", None):
                    warning = next(
                        (str(item) for item in (getattr(page, "warnings", None) or [])
                         if str(item).startswith("识别失败：")),
                        "",
                    )
                    reason = (warning.split("识别失败：", 1)[1]
                              if warning else "未识别出可写入的内容")
                    mark_failed(fid, source_name, reason[:200])
                    log(f"{label} 识别失败：{reason}")
                    continue
                page.source_name = source_name
                page.source_page = page_no
                # 工作表名用显示名 label（默认「图片N」或用户命名），清洗规则与
                # md2excel._sheet_title 一致；多页 PDF/TIFF 的「-第N页/-页码」后缀已在 label 里。
                # 来源列仍用 source_name（原名），不受命名影响。
                page.name = _clean_label(label)
                pages.append(page)
                processed_since_restart += 1
                emit(type="file_status", id=fid, tag="ok")
                spent = time.time() - started
                page_seconds.append({"file": source_name, "page": page_no,
                                     "seconds": round(spent, 1)})
                log(f"{label} 识别完成，用时 {spent:.0f} 秒")
                emit(type="progress", done=index, total=len(jobs), indeterminate=False)

            if not pages:
                raise RuntimeError("所有文件都未识别出内容，请换更清晰的图片重试。")

            # 兜底：识别要花数分钟，期间用户可能又跑了一次转换或别的程序写入了同名文件，
            # 写盘前再查一次，避免 wb.save 直接覆盖别人的成果
            if out_path.exists():
                safe_path = _unique_output(out_path)
                log(f"保存路径在转换期间被再次占用，已改存为：{safe_path}")
                out_path = safe_path
            ALLOWED_OUTPUTS.add(str(out_path.resolve()))

            # 待复核分流与 Excel 两张表必须同源：按 convert_exact 相同的顺序
            # （页顺序 × 页内 flags 顺序，页码即 page.name）重建 entries，再调用
            # md2excel._triage_flags 同一函数。同一输入 + 同一纯函数，保证
            # 高/中/低条数与「待复核清单」「待复核-低优先」写出行数逐条一致。
            flag_entries = []
            for page_index, page in enumerate(pages):
                for fl in (getattr(page, "flags", None) or []):
                    enriched = dict(fl)
                    enriched["page_index"] = page_index
                    enriched["page_no"] = int(getattr(page, "source_page", 0) or 1)
                    # 只有全量识别有「待复核」工作台与原图预览，其他模式不花这份算力
                    box = flag_image_box(page, fl) if mode == "full" else None
                    if box:
                        enriched["box"] = box
                    flag_entries.append((getattr(page, "name", ""), enriched))
            hi_mid_flags, low_flags = _triage_flags(flag_entries)
            flag_items = []
            for name, fl, kind, priority in list(hi_mid_flags) + list(low_flags):
                item = dict(fl)
                item["file"] = name
                item["priority"] = priority  # 高 / 中 / 低
                item["kind"] = kind          # 问题类型，如 粘连拆分
                flag_items.append(item)
            quality = _quality_report(pages, flag_items)
            file_summary = _file_summary(pages, page_seconds)
            if mode == "waybill":
                emit(type="phase", text="正在汇总运单号…")
                result = convert_waybill(pages, out_path)
            else:
                emit(type="phase", text="正在把每页内容原样写入 Excel…")
                log("输出方式：多图合并为一张表（首列标注来源图片）"
                    if layout == "merged" else "输出方式：一图一表（每张图片一个工作表）")
                result = convert_exact(pages, out_path, include_index=False,
                                       include_aux=False, include_flags=True,
                                       quality=quality, layout=layout)
            # 命名增量字段（label/suggested_name）在写盘后附加：不影响质量表写入，
            # 只随 done 事件下发，供前端「命名」弹窗与结果卡按文件展示。
            _attach_naming(pages, files, quality, file_summary)
            if mode == "full":
                images = _make_search_previews(job_id, pages)
                # 待复核原图定位：客户端按 page_index 取缩略图，按原图宽高换算红框
                for item in flag_items:
                    info = images.get(str(item.get("page_index")))
                    if info:
                        item["image_width"] = int(info["width"])
                        item["image_height"] = int(info["height"])
                _remember_search_index(job_id, result, out_path, images)
            if mode == "waybill":
                log(f"纯单号：共识别 {result.waybills} 条候选，去重后导出 "
                    f"{result.unique_waybills} 个（已确认 {result.confirmed_waybills}，"
                    f"待核对 {result.pending_waybills}，格式异常/非YT "
                    f"{result.format_invalid_waybills} 条未写入）")
                if result.pending_waybills:
                    log("注意：待核对单号已写入清单，请对照原图复核后再使用。")
            else:
                log(f"全量识别：{result.sheets} 个工作表、{result.tables} 个表格，按原图版式写入")
                if result.waybills:
                    log(f"单号：已确认 {result.confirmed_waybills} 处，"
                        f"待核对 {result.pending_waybills} 处，"
                        f"格式异常/非YT {result.format_invalid_waybills} 处")
            if result.index_issues:
                if mode == "waybill":
                    log(f"另有 {result.index_issues} 页未提取到单号，请对照原图核查。")
                else:
                    log(f"索引异常 {result.index_issues} 项，"
                        f"请对比原图核查（导出中不含「索引异常」工作表）。")
            job_store.save_job(job_id, {
                "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "mode": mode,
                "layout": layout,
                "out_path": str(out_path),
                "files": [orig for _fid, _src, orig, _label in files],
                "total_pages": len(pages),
                "sheets": int(getattr(result, "sheets", 0) or 0),
                "tables": int(getattr(result, "tables", 0) or 0),
                "waybills": int(getattr(result, "waybills", 0) or 0),
                "unique_waybills": int(getattr(result, "unique_waybills", 0) or 0),
                "confirmed_waybills": int(getattr(result, "confirmed_waybills", 0) or 0),
                "pending_waybills": int(getattr(result, "pending_waybills", 0) or 0),
                "format_invalid_waybills": int(getattr(result, "format_invalid_waybills", 0) or 0),
                "flags_total": len(flag_items),
                "quality_level": quality.get("level", ""),
                "quality_score": quality.get("score"),
                "file_summary": file_summary,
                "flags": flag_items,
            }, keep=_job_keep())
            emit(type="done", out_path=str(out_path), mode=mode, layout=layout,
                 searchable=(mode == "full"), sheets=result.sheets,
                 tables=result.tables, flagged=int(getattr(result, "flagged", 0) or 0),
                 waybills=int(getattr(result, "waybills", 0) or 0),
                 unique_waybills=int(getattr(result, "unique_waybills", 0) or 0),
                 duplicate_waybills=int(getattr(result, "duplicate_waybills", 0) or 0),
                 suspicious_waybills=int(getattr(result, "suspicious_waybills", 0) or 0),
                 confirmed_waybills=int(getattr(result, "confirmed_waybills", 0) or 0),
                 pending_waybills=int(getattr(result, "pending_waybills", 0) or 0),
                 format_invalid_waybills=int(getattr(result, "format_invalid_waybills", 0) or 0),
                 index_issues=int(getattr(result, "index_issues", 0) or 0),
                 warnings=result.warnings, failed=int(failed), flags=flag_items,
                 quality=quality, file_summary=file_summary,
                 failed_items=[dict(item,
                                    retryable=_failure_retryable(item.get("reason", "")))
                               for item in failed_map.values()],
                 total_pages=len(pages),
                 elapsed=round(time.time() - task_started, 1),
                 download_url=f"/api/download?path={quote(str(out_path))}")
        else:
            raise EngineError(f"不支持的导出模式：{mode}")
    except Exception as exc:  # noqa: BLE001 - 失败原因需要反馈到界面
        if job.cancelled or "已取消" in str(exc):
            emit(type="cancelled")
        else:
            emit(type="error", message=str(exc))
    finally:
        if worker is not None and mode != "find":
            try:
                worker.close()
            except Exception:
                pass
        # 无论成功、失败还是取消，都清掉这次上传的临时目录，避免反复使用时占盘；
        # 单号查询用的预览图落在本机任务目录（job_store），不受这里影响。
        if temp_root is not None:
            shutil.rmtree(temp_root, ignore_errors=True)
        # PDF 渲染临时目录：成功、失败、取消都在 finally 统一清理，避免失败时残留
        for tmp_dir in pdf_tmp_dirs:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        # 上传/旋转产生的临时副本（convert_start 建的目录）也一并清掉
        for extra in extra_temp_dirs or []:
            shutil.rmtree(extra, ignore_errors=True)
        RUNNING_JOBS.discard(job)
        _state_update(running=False, finished_at=time.time())
        job.worker = None
        job.done = True
        job.finished_at = time.time()
        job.q.put({"type": "_end"})


# ---------------------------------------------------------------- 路由
@app.get("/")
async def index() -> RedirectResponse:
    return RedirectResponse("/web/index.html")


@app.get("/api/privacy")
async def privacy() -> dict:
    return {"title": core.APP_TITLE, "version": core.APP_VERSION, "notice": core.PRIVACY_NOTICE}


def _version_tuple(text: str) -> tuple:
    """把 v1.2.3 / 1.2.3.0 之类解析成可比较的整数元组。"""
    nums = re.findall(r"\d+", str(text or ""))
    return tuple(int(n) for n in nums[:4]) or (0,)


@app.get("/api/update/check")
async def update_check(force: int = 0) -> dict:
    """查询最新版本号；只读公开 Release 信息，不发送本机文件。"""
    now = time.time()
    if (not force and UPDATE_CACHE["data"] is not None
            and now - UPDATE_CACHE["t"] < UPDATE_CACHE_SECONDS):
        return UPDATE_CACHE["data"]

    result = {
        "ok": False,
        "current": APP_VERSION,
        "latest": "",
        "has_update": False,
        "url": "",
        "notes": "",
        "published_at": "",
        "error": "",
        "repo": UPDATE_REPO,
    }
    try:
        req = urllib.request.Request(
            UPDATE_API_URL,
            headers={
                "User-Agent": f"FatDuckOCR/{APP_VERSION}",
                "Accept": "application/vnd.github+json",
            },
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        latest = str(data.get("tag_name") or data.get("name") or "").strip()
        latest = latest.lstrip("vV")
        result.update({
            "ok": True,
            "latest": latest,
            "has_update": _version_tuple(latest) > _version_tuple(APP_VERSION),
            "url": str(data.get("html_url") or ""),
            "notes": str(data.get("body") or "")[:800],
            "published_at": str(data.get("published_at") or ""),
        })
    except Exception as exc:  # noqa: BLE001 - 更新检查失败不应影响主流程
        result["error"] = f"{type(exc).__name__}: {exc}"
    UPDATE_CACHE["t"], UPDATE_CACHE["data"] = now, result
    return result


@app.get("/api/default-output")
async def default_output() -> dict:
    return {"path": str(_default_output())}


@app.get("/api/config")
async def get_config() -> dict:
    cfg = core.load_config()
    return {
        "last_mode": core.normalize_mode(cfg.get("last_mode")) or "full",
        "default_mode": (core.normalize_mode(cfg.get("default_mode"))
                         or core.normalize_mode(cfg.get("last_mode")) or "full"),
        "last_output": cfg.get("last_output", ""),
        "last_layout": _normalize_layout(cfg.get("last_layout")),
        "default_layout": _normalize_layout(
            cfg.get("default_layout") or cfg.get("last_layout")),
        "privacy_ack": bool(cfg.get("privacy_ack")),
        "auto_open": bool(cfg.get("auto_open")),
        "job_keep": _job_keep(),
        # 桌面版粘贴增强：鼠标移入自动激活 / 全局快捷键（缺省开启）
        "hover_activate": bool(cfg.get("hover_activate", True)),
        "paste_hotkey": bool(cfg.get("paste_hotkey", True)),
        "check_updates": bool(cfg.get("check_updates", True)),
        # 云端文本纠错（可选功能，默认关闭）：只回状态，不回密钥本身
        "nv_fix_enabled": bool(cfg.get("nv_fix_enabled")),
        "nv_fix_key_source": nv_fix.key_source(),
        "nv_fix_model": nv_fix.get_model(),
    }


@app.post("/api/config")
async def post_config(payload: dict) -> dict:
    cfg = core.load_config()
    for key in ("last_mode", "last_output", "privacy_ack", "nv_fix_enabled",
                "default_mode", "auto_open", "job_keep", "last_layout",
                "default_layout", "hover_activate", "paste_hotkey",
                "check_updates"):
        if key in payload:
            cfg[key] = payload[key]
    if "last_mode" in payload:
        cfg["last_mode"] = core.normalize_mode(payload["last_mode"]) or "full"
    if "default_mode" in payload:
        cfg["default_mode"] = core.normalize_mode(payload["default_mode"]) or "full"
    if "last_layout" in payload:
        cfg["last_layout"] = _normalize_layout(payload["last_layout"])
    if "default_layout" in payload:
        cfg["default_layout"] = _normalize_layout(payload["default_layout"])
    if "auto_open" in payload:
        cfg["auto_open"] = bool(payload["auto_open"])
    if "hover_activate" in payload:
        cfg["hover_activate"] = bool(payload["hover_activate"])
    if "paste_hotkey" in payload:
        cfg["paste_hotkey"] = bool(payload["paste_hotkey"])
    if "check_updates" in payload:
        cfg["check_updates"] = bool(payload["check_updates"])
    if "job_keep" in payload:
        try:
            cfg["job_keep"] = max(1, min(job_store.MAX_KEEP, int(payload["job_keep"])))
        except (TypeError, ValueError):
            cfg.pop("job_keep", None)
    core.save_config(cfg)
    if "job_keep" in cfg:
        job_store.prune(cfg["job_keep"])
    return {"ok": True}


_ROTATE_ANGLES = (90, 180, 270)


def _plan_rotation(plan: dict) -> int:
    """读取前端传来的旋转角度（顺时针），只接受 90 的整数倍。"""
    try:
        angle = int(plan.get("rotate") or 0)
    except (TypeError, ValueError):
        return 0
    return angle if angle in _ROTATE_ANGLES else 0


def _rotate_for_ocr(src: Path, display: str, angle: int,
                    temp_root: Path, fid: str) -> Path:
    """在临时目录生成转正后的副本供识别使用；原文件保持不动。"""
    from PIL import Image, ImageOps

    dst = temp_root / f"{fid}_rot{angle}.png"
    try:
        with Image.open(src) as im:
            # 手机照片常带 EXIF 方向：浏览器预览已按它转正，这里保持一致
            im = ImageOps.exif_transpose(im)
            im = im.rotate(-angle, expand=True)  # PIL 逆时针为正，取负即顺时针
            if im.mode in ("RGBA", "LA", "P"):
                rgba = im.convert("RGBA")
                canvas = Image.new("RGB", rgba.size, (255, 255, 255))
                canvas.paste(rgba, mask=rgba.split()[-1])
                im = canvas
            elif im.mode not in ("RGB", "L"):
                im = im.convert("RGB")
            im.save(dst, format="PNG")
    except Exception as exc:  # noqa: BLE001 - 具体文件失败原因要回传界面
        raise ValueError(f"{display} 旋转失败：{exc}") from exc
    return dst


@app.post("/api/convert")
async def convert_start(
    files: list[UploadFile] | None = File(default=None),
    ids: list[str] | None = Form(default=None),
    mode: str = Form("full"),
    output: str = Form(""),
    items: str = Form(""),
    target: str = Form(""),
    layout: str = Form("per_image"),
) -> JSONResponse:
    uploads = files or []
    legacy_ids = ids or []
    # 每次发起新任务时清掉过期/超量的已结束任务，避免常驻托盘时 JOBS 只增不减
    _sweep_jobs()
    mode = core.normalize_mode(mode)
    layout = _normalize_layout(layout)
    if not mode:
        return JSONResponse({"detail": "未知工作方式"}, status_code=400)
    if mode == "find":
        target = _normalize_waybill(target)
        if len(target) < 6:
            return JSONResponse({"detail": "请粘贴完整单号，至少输入 6 个字母或数字"}, status_code=400)

    out_path = Path() if mode == "find" else (Path(output).expanduser() if output else _default_output())
    if mode != "find" and out_path.suffix.lower() != ".xlsx":
        out_path = out_path.with_suffix(".xlsx")
    # 目标已存在不再直接失败：自动另存为 "name (n).xlsx"，
    # 旧文件保留、转换继续，前端会把改名结果告诉用户
    renamed_from = ""
    if mode != "find" and out_path.exists():
        renamed_from = str(out_path)
        out_path = _unique_output(out_path)
    if mode != "find":
        # 保存目录在任务开始前校验：识别要花数分钟，不能等全部结束才发现目录不存在
        parent = out_path.parent if str(out_path.parent) else Path(".")
        try:
            parent.mkdir(parents=True, exist_ok=True)
        except OSError:
            return JSONResponse(
                {"detail": f"保存位置不可用：{parent}"}, status_code=400)
        if not parent.is_dir():
            return JSONResponse(
                {"detail": f"保存位置不可用：{parent}"}, status_code=400)
        # 登记本次产物，供后续 /api/download 与 /api/reveal 白名单校验
        ALLOWED_OUTPUTS.add(str(out_path.resolve()))

    temp_root = Path(tempfile.mkdtemp(prefix="img2xlsx_web_up_"))
    # (fid, 识别用路径, 原始文件名, 显示名 label)
    saved: list[tuple[str, Path, str, str]] = []
    try:
        if items:
            # 新版前端：磁盘文件直接给路径（免上传、可预览/复制/删除），
            # 拖入的文件仍走上传（items 里用 upload 下标指向 multipart 数组）。
            try:
                plans = json.loads(items)
            except json.JSONDecodeError:
                raise ValueError("文件清单格式错误")
            if not isinstance(plans, list):
                raise ValueError("文件清单格式错误")
            if len(plans) > MAX_FILES:
                raise ValueError(f"一次最多处理 {MAX_FILES} 个文件，请分批转换")
            for index, plan in enumerate(plans):
                if not isinstance(plan, dict):
                    raise ValueError("文件清单格式错误")
                fid = _safe_fid(plan.get("id"), index)
                name = str(plan.get("name") or "")
                # 显示名 label（默认「图片N」/用户命名）；缺省回落原名去扩展名，
                # 与旧版工作表名行为一致。原名 name 仍单独保留供来源列使用。
                label_raw = str(plan.get("label") or "").strip()
                angle = _plan_rotation(plan)
                if plan.get("path"):
                    src = Path(str(plan["path"]))
                    orig_name = name or src.name
                    label = label_raw or Path(orig_name).stem
                    display = label
                    if not src.is_file():
                        raise ValueError(f"文件不存在或已被移动：{display}")
                    if src.stat().st_size > MAX_FILE_BYTES:
                        raise ValueError(f"{display} 超过单文件上限 {_MAX_FILE_TEXT}")
                    # 登记原始磁盘路径：预览/复制/删除/命名复用同一白名单（旋转副本是临时文件，不登记）
                    ALLOWED_FILES.add(str(src))
                    # 用户在预览里调正过方向：生成临时的转正副本再识别（PDF 页方向由渲染决定，不在此处理）
                    if angle and src.suffix.lower() != ".pdf":
                        src = _rotate_for_ocr(src, display, angle, temp_root, fid)
                    saved.append((fid, src, orig_name, label))
                    continue
                upload_index = plan.get("upload")
                if not isinstance(upload_index, int) or not (0 <= upload_index < len(uploads)):
                    continue
                upload = uploads[upload_index]
                if not upload.filename:
                    continue
                # 落盘名与展示名分离：原名只用于界面/工作表，磁盘名固定为序号+白名单后缀
                stored_name = _safe_upload_name(index, upload.filename)
                dst = _contained_upload(temp_root, stored_name)
                _save_upload_limited(upload, dst)
                if angle and Path(upload.filename).suffix.lower() != ".pdf":
                    dst = _rotate_for_ocr(dst, upload.filename, angle, temp_root, fid)
                # 同时保留界面上的原始文件名：来源列、搜索定位都用它，
                # 不能把上传临时目录的 id 前缀带进用户可见的结果里。
                orig_name = upload.filename
                label = label_raw or Path(orig_name).stem
                saved.append((fid, dst, orig_name, label))
        else:
            # 兼容旧调用（工具脚本/自检）：fid 与上传文件按顺序配对
            if len(uploads) > MAX_FILES:
                raise ValueError(f"一次最多处理 {MAX_FILES} 个文件，请分批转换")
            for index, (raw_fid, upload) in enumerate(zip(legacy_ids, uploads)):
                if not upload.filename:
                    continue
                fid = _safe_fid(raw_fid, index)
                stored_name = _safe_upload_name(index, upload.filename)
                dst = _contained_upload(temp_root, stored_name)
                _save_upload_limited(upload, dst)
                saved.append((fid, dst, upload.filename, Path(upload.filename).stem))
        if not saved:
            raise ValueError("没有收到可识别的文件")
    except ValueError as exc:
        shutil.rmtree(temp_root, ignore_errors=True)
        return JSONResponse({"detail": str(exc)}, status_code=400)

    job = Job()
    job_id = f"j{int(time.time() * 1000)}{os.getpid()}"
    JOBS[job_id] = job
    threading.Thread(
        target=_run_job,
        args=(job, saved, mode, out_path, job_id, target, layout, [temp_root]),
        daemon=True,
    ).start()
    return JSONResponse(
        {"job_id": job_id, "out_path": "" if mode == "find" else str(out_path),
         "renamed_from": renamed_from}
    )


@app.post("/api/pick-files")
# 同步 def（勿改 async）：文件框/PowerShell 是阻塞调用，必须走线程池，
# 否则对话框开着期间整个事件循环被冻住，所有请求无响应。
def pick_files() -> dict:
    """系统文件选择框：返回真实磁盘路径，供预览/复制/删除使用。

    浏览器自带的 file input 只给 File 对象、拿不到路径，因此"复制文件/删除文件"
    必须由本机进程选择文件；仅在 127.0.0.1 上可用。
    """
    script = r'''
Add-Type -AssemblyName System.Windows.Forms
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$d = New-Object System.Windows.Forms.OpenFileDialog
$d.Title = "选择图片或 PDF"
$d.Filter = "图片与 PDF|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp;*.pdf|所有文件|*.*"
$d.Multiselect = $true
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  $list = @()
  foreach ($f in $d.FileNames) {
    $list += [pscustomobject]@{ path = $f; name = [System.IO.Path]::GetFileName($f); size = (Get-Item -LiteralPath $f).Length }
  }
  $list | ConvertTo-Json -Compress
} else { "[]" }
'''
    ok, output = _run_powershell(script, timeout=120, interactive=True)
    if not ok:
        return {"ok": False, "error": output or "无法打开系统文件选择框"}
    line = next((item for item in reversed(output.splitlines()) if item.strip()[:1] in "[{"), "[]")
    try:
        data = json.loads(line)
    except json.JSONDecodeError:
        return {"ok": False, "error": "文件选择结果解析失败"}
    if isinstance(data, dict):
        data = [data]
    picked, skipped = [], []
    for entry in data:
        raw = str(entry.get("path") or "")
        if not raw:
            continue
        target = Path(raw)
        if target.suffix.lower().lstrip(".") not in SUPPORTED_EXTS:
            skipped.append(f"{target.name}（不支持的格式）")
            continue
        if not target.is_file():
            continue
        size = target.stat().st_size
        if size > MAX_FILE_BYTES:
            skipped.append(f"{target.name}（超过 {_MAX_FILE_TEXT}）")
            continue
        ALLOWED_FILES.add(str(target))
        picked.append({"path": str(target), "name": target.name, "size": size})
    return {"ok": True, "files": picked, "skipped": skipped,
            "max_mb": int(MAX_FILE_BYTES // (1024 * 1024))}


# 剪贴板图片落盘目录（系统临时目录下，程序退出后由系统自行清理）
CLIP_DIR = Path(tempfile.gettempdir()) / "img2doc_clipboard"
_clip_seq = 0

# 「复制原图到剪贴板」的临时 PNG 目录：写入剪贴板后立即删除，避免残留
COPY_IMAGE_DIR = Path(tempfile.gettempdir()) / "img2doc_copy_image"


@app.post("/api/clipboard-peek")
def clipboard_peek() -> dict:
    """只探测剪贴板里有没有可识别的图片/文件，不落盘。

    供「鼠标移入窗口时浮出粘贴提示」使用：必须足够轻，不写临时文件；
    读不到（被其他程序占用、非 Windows）时按“没有”处理，不打扰用户。
    """
    if ImageGrab is None:
        return {"ok": False, "has": False, "count": 0}
    try:
        data = ImageGrab.grabclipboard()
    except Exception:  # noqa: BLE001 - 剪贴板可能被其他程序短暂占用
        return {"ok": False, "has": False, "count": 0}
    if data is None:
        return {"ok": True, "has": False, "count": 0}
    if isinstance(data, list):
        count = 0
        for raw in data:
            target = Path(str(raw))
            if target.suffix.lower().lstrip(".") not in SUPPORTED_EXTS:
                continue
            try:
                if target.is_file() and target.stat().st_size <= MAX_FILE_BYTES:
                    count += 1
            except OSError:
                continue
        return {"ok": True, "has": count > 0, "count": count}
    # 位图（截图工具、网页/微信右键复制图片）
    return {"ok": True, "has": True, "count": 1}


@app.post("/api/clipboard-paste")
def clipboard_paste() -> dict:
    """读取系统剪贴板里的图片（或复制的文件），落盘后返回真实路径。

    前端把这里当作兜底：WebView2 在焦点不在输入框时不一定派发 paste 事件，
    从资源管理器复制的文件（CF_HDROP）也未必暴露给 JS。有了这个接口，鼠标在
    窗口任何位置按 Ctrl+V 都能把剪贴板里的图片加进列表。
    同步 def → 由线程池执行，读取剪贴板不会阻塞事件循环。
    """
    global _clip_seq
    if ImageGrab is None:
        return {"ok": False, "error": "当前环境无法读取剪贴板"}
    data = None
    for attempt in range(3):
        try:
            data = ImageGrab.grabclipboard()
            break
        except Exception as exc:  # noqa: BLE001 - 剪贴板可能被其他程序短暂占用
            if attempt == 2:
                return {"ok": False, "error": f"读取剪贴板失败：{exc}"}
            time.sleep(0.15)
    if data is None:
        return {"ok": True, "files": [], "skipped": []}

    paths: list[Path] = []
    skipped: list[str] = []
    if isinstance(data, list):
        # 剪贴板里是"复制的文件"（如从资源管理器 Ctrl+C 的图片）
        for raw in data:
            target = Path(str(raw))
            if target.suffix.lower().lstrip(".") not in SUPPORTED_EXTS:
                skipped.append(f"{target.name}（不支持的格式）")
                continue
            if not target.is_file():
                continue
            if target.stat().st_size > MAX_FILE_BYTES:
                skipped.append(f"{target.name}（超过 {_MAX_FILE_TEXT}）")
                continue
            paths.append(target)
    else:
        # 剪贴板里是位图（截图工具、网页右键复制图片等），统一存成 PNG
        try:
            CLIP_DIR.mkdir(parents=True, exist_ok=True)
            _clip_seq += 1
            stamp = time.strftime("%Y%m%d_%H%M%S")
            target = CLIP_DIR / f"剪贴板图片_{stamp}_{_clip_seq}.png"
            data.save(str(target), format="PNG")
            paths.append(target)
        except Exception as exc:  # noqa: BLE001 - 落盘失败要把原因反馈到界面
            return {"ok": False, "error": f"保存剪贴板图片失败：{exc}"}

    files = []
    for target in paths:
        ALLOWED_FILES.add(str(target))
        files.append({"path": str(target), "name": target.name,
                      "size": target.stat().st_size})
    return {"ok": True, "files": files, "skipped": skipped}


@app.get("/api/preview")
async def preview_file(path: str = "") -> Response:
    """预览用户选中的本地图片；TIFF 浏览器无法直接显示，转成 PNG 返回。"""
    target = Path(path)
    if str(target) not in ALLOWED_FILES or not target.is_file():
        return JSONResponse({"detail": "文件未授权或不存在"}, status_code=404)
    suffix = target.suffix.lower()
    if suffix in {".tif", ".tiff"}:
        try:
            import io

            from PIL import Image

            with Image.open(target) as im:
                im = im.convert("RGB")
                buf = io.BytesIO()
                im.save(buf, format="PNG")
            return Response(buf.getvalue(), media_type="image/png")
        except Exception as exc:  # noqa: BLE001
            return JSONResponse({"detail": f"预览转换失败：{exc}"}, status_code=500)
    media = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
    return FileResponse(str(target), media_type=media)


@app.post("/api/file-op")
def file_op(payload: dict) -> dict:
    """对白名单内的本地文件执行"复制到剪贴板 / 删除到回收站"。"""
    action = str(payload.get("action") or "")
    raw = str(payload.get("path") or "")
    target = Path(raw)
    if action not in ("copy", "delete"):
        return {"ok": False, "error": "未知操作"}
    if str(target) not in ALLOWED_FILES:
        return {"ok": False, "error": "该文件不是从程序内选择的，未授权操作"}
    if not target.is_file():
        return {"ok": False, "error": "文件不存在或已被移动"}
    if action == "copy":
        script = (
            "Add-Type -AssemblyName System.Windows.Forms;"
            "$col = New-Object System.Collections.Specialized.StringCollection;"
            "$col.Add($env:IMG2DOC_FILE_TARGET) | Out-Null;"
            "[System.Windows.Forms.Clipboard]::SetFileDropList($col)"
        )
        # 剪贴板可能被其他程序短暂占用（OpenClipboard 失败），重试几次再报错
        detail = ""
        for attempt in range(3):
            ok, detail = _run_powershell(script, target=str(target))
            if ok:
                return {"ok": True}
            if "剪贴板" not in detail and "clipboard" not in detail.lower():
                break
            time.sleep(0.4)
        if "剪贴板" in detail or "clipboard" in detail.lower():
            detail = "剪贴板正被其他程序占用，请稍后再试"
        return {"ok": False, "error": detail or "复制到剪贴板失败"}
    script = (
        "Add-Type -AssemblyName Microsoft.VisualBasic;"
        "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile("
        "$env:IMG2DOC_FILE_TARGET, 'OnlyErrorDialogs', 'SendToRecycleBin')"
    )
    ok, detail = _run_powershell(script, target=str(target))
    if not ok:
        return {"ok": False, "error": detail or "删除失败"}
    ALLOWED_FILES.discard(str(target))
    return {"ok": True}


# ---------------------------------------------------------------- 图片命名
# 最近一批 rename 的撤销映射：disk 记磁盘 old/new，sheets 记工作表 old/new。
# 随 rename_history.json 落盘，程序重启后仍可撤销（磁盘名与表名都改回）。
_RENAME_HISTORY: list[dict] = []
_RENAME_HISTORY_MAX = 20
# 辅助工作表标题：命名只改数据 sheet，这些固定附加页一律不动。
_AUX_SHEET_TITLES = {"合并数据", "待复核清单", "待复核-低优先",
                     "索引异常", "单号索引", "拍摄质量"}


def _sanitize_rename_batch(raw) -> dict | None:
    """校验从磁盘读回的一批撤销记录，只保留结构合法的条目。

    磁盘条目要求 old/new 同目录（rename 只在同一目录内改名），跨目录记录
    视为被篡改直接丢弃；工作表条目要求 path/old/new 均为非空字符串。
    没有任何有效条目的批次返回 None。
    """
    if not isinstance(raw, dict):
        return None
    batch = {"ts": raw.get("ts") if isinstance(raw.get("ts"), (int, float)) else 0,
             "job_id": str(raw.get("job_id") or "")[:64],
             "disk": [], "sheets": [], "labels": [], "dates": []}
    for rec in raw.get("labels") or []:
        if not isinstance(rec, dict):
            continue
        old, new = str(rec.get("old") or ""), str(rec.get("new") or "")
        if old and new:
            batch["labels"].append({"old": old, "new": new})
    for rec in raw.get("dates") or []:
        if not isinstance(rec, dict):
            continue
        path = str(rec.get("path") or "")
        old_m, new_m = rec.get("old_mtime"), rec.get("new_mtime")
        if path and all(isinstance(v, (int, float)) and math.isfinite(v) for v in (old_m, new_m)):
            batch["dates"].append({"path": path, "old_mtime": float(old_m),
                                   "new_mtime": float(new_m)})
    for rec in raw.get("disk") or []:
        if not isinstance(rec, dict):
            continue
        old, new = str(rec.get("old") or ""), str(rec.get("new") or "")
        if not old or not new:
            continue
        try:
            if Path(old).parent != Path(new).parent:
                continue
        except (OSError, ValueError):
            continue
        batch["disk"].append({"old": old, "new": new})
    for rec in raw.get("sheets") or []:
        if not isinstance(rec, dict):
            continue
        path = str(rec.get("path") or "")
        old, new = str(rec.get("old") or ""), str(rec.get("new") or "")
        if path and old and new:
            batch["sheets"].append({"path": path, "old": old, "new": new})
    has_meta = bool(batch["job_id"]) and bool(batch["labels"])
    return batch if (batch["disk"] or batch["sheets"] or batch["dates"] or has_meta) else None


def _load_rename_history() -> None:
    """启动时读回撤销历史；仍存在的「改名后路径」重新登记白名单。

    只有程序自己改过名的文件才会被登记，撤销时才能再把它们改回去，
    与 rename 的 fail-closed 白名单口径保持闭环。
    """
    _RENAME_HISTORY.clear()
    for raw in job_store.load_rename_history():
        batch = _sanitize_rename_batch(raw)
        if batch is None:
            continue
        _RENAME_HISTORY.append(batch)
        for rec in batch["disk"]:
            target = Path(rec["new"])
            try:
                if target.is_file():
                    ALLOWED_FILES.add(str(target))
            except OSError:
                continue
    del _RENAME_HISTORY[:-_RENAME_HISTORY_MAX]


def _save_rename_history() -> None:
    job_store.save_rename_history(_RENAME_HISTORY[-_RENAME_HISTORY_MAX:])


def rename_history_state() -> dict:
    last = _RENAME_HISTORY[-1] if _RENAME_HISTORY else None
    return {"ok": True, "can_undo": last is not None,
            "count": len(_RENAME_HISTORY),
            "last": ({"ts": last["ts"], "disk": len(last["disk"]),
                      "sheets": len(last["sheets"])} if last else None)}


@app.get("/api/rename/history")
async def get_rename_history() -> dict:
    """页面加载时查询：重启后是否还有可撤销的命名批次。"""
    return rename_history_state()


def _sync_job_labels(job_id: str, mapping: dict) -> bool:
    """把命名结果写回任务历史快照的 file_summary[].label；没有可改的就返回 False。

    老任务的条目没有 label 字段，自然不会被匹配到，不做任何补写。
    """
    if not job_id or not mapping:
        return False
    meta = job_store.load_job(job_id)
    if not meta:
        return False
    changed = False
    for entry in meta.get("file_summary") or []:
        if not isinstance(entry, dict):
            continue
        new = mapping.get(str(entry.get("label") or ""))
        if new and new != entry.get("label"):
            entry["label"] = new
            changed = True
    return job_store.save_job(job_id, meta, keep=_job_keep()) if changed else False


@app.post("/api/rename")
def rename_files(payload: dict):
    """批量命名：改显示名 label，可选改磁盘原文件名，并同步已导出 xlsx 的数据表名。

    安全：要求改磁盘名且给了路径的，路径必须在 ALLOWED_FILES 白名单内（pick-files /
    convert 登记过），否则整批拒绝 400；未给路径（粘贴/拖拽上传）只改显示名与表名。
    工作簿只认产物白名单内的 out_path，只改数据 sheet 标题，不动辅助页与单元格。
    """
    items = payload.get("items")
    if not isinstance(items, list) or not items:
        return JSONResponse({"detail": "没有要命名的文件"}, status_code=400)
    if len(items) > MAX_FILES:
        return JSONResponse({"detail": f"一次最多命名 {MAX_FILES} 个文件"}, status_code=400)
    job_id = str(payload.get("job_id") or "")
    out_path = str(payload.get("out_path") or "")

    # 安全前置校验：未登记路径一律 400（fail-closed），与 preview/file-op 同一白名单。
    # 改磁盘名与改文件日期都是动原图，同一口径。
    for item in items:
        if not isinstance(item, dict):
            return JSONResponse({"detail": "命名清单格式错误"}, status_code=400)
        if (bool(item.get("rename_disk")) or bool(item.get("set_date"))) \
                and str(item.get("path") or ""):
            if not _registered_file(Path(str(item["path"]))):
                return JSONResponse(
                    {"detail": "该文件不是从程序内选择的，未授权改名"}, status_code=400)

    # 仅当带 job_id 且 out_path 在产物白名单内才加载工作簿
    wb = None
    wb_path = None
    if job_id and out_path:
        resolved = _resolve_path(out_path)
        if resolved is not None and resolved.is_file() and resolved in _output_paths():
            try:
                from openpyxl import load_workbook
                wb = load_workbook(resolved)
                wb_path = resolved
            except Exception:  # noqa: BLE001 - 工作簿损坏时只改显示名/磁盘名
                wb = None
    used = set(wb.sheetnames) if wb is not None else set()

    results = []
    batch = {"ts": time.time(), "job_id": job_id, "disk": [], "sheets": [],
             "labels": [], "dates": []}
    sheet_changed = False
    for item in items:
        fid = str(item.get("fid") or "")
        raw_path = str(item.get("path") or "")
        old_label = str(item.get("old_label") or "")
        rename_disk = bool(item.get("rename_disk"))
        set_date = bool(item.get("set_date"))
        print_time = str(item.get("print_time") or "").strip()
        new_name = _clean_disk_name(item.get("new_name"))
        res = {"fid": fid, "old_path": raw_path, "new_path": raw_path,
               "old_label": old_label, "new_name": new_name,
               "disk_renamed": False, "sheet_renamed": False, "date_set": False}
        if not new_name:
            res["reason"] = "新名称为空或仅含非法字符"
            results.append(res)
            continue

        # 1) 磁盘改名：保留原后缀，重名走 _unique_output 的 (n) 规则
        if rename_disk and raw_path:
            target = Path(raw_path)
            if not target.is_file():
                res["reason"] = "原图不存在或已被移动"
            else:
                suffix = target.suffix
                stem = new_name
                if suffix and stem.lower().endswith(suffix.lower()):
                    stem = stem[: -len(suffix)]
                stem = stem.strip()
                if not stem:
                    res["reason"] = "新名称为空或仅含非法字符"
                else:
                    dest = target.with_name(stem + suffix)
                    if dest.resolve() == target.resolve():
                        res["new_path"] = str(target)  # 同名无需改磁盘
                    else:
                        if dest.exists():
                            dest = _unique_output(dest)
                        try:
                            os.rename(target, dest)
                            ALLOWED_FILES.discard(str(target))
                            ALLOWED_FILES.add(str(dest))
                            res["new_path"] = str(dest)
                            res["disk_renamed"] = True
                            batch["disk"].append({"old": str(target), "new": str(dest)})
                        except OSError as exc:
                            res["reason"] = f"磁盘改名失败：{exc}"
        elif rename_disk and not raw_path:
            res["reason"] = "原图位置未知，仅改表格名"

        # 1.5) 文件日期：把修改时间改成页脚识别到的打印时间；没识别到/格式无效的不动。
        # 放在磁盘改名之后，直接写最终路径；旧时间记进撤销批次。
        if set_date:
            epoch = _print_time_epoch(print_time)
            if not print_time:
                res.setdefault("reason", "未识别到打印时间，保持原样")
            elif epoch is None:
                res.setdefault("reason", "打印时间格式无效，保持原样")
            elif not raw_path:
                res.setdefault("reason", "原图位置未知，无法改文件日期")
            else:
                target = Path(res["new_path"])
                try:
                    st = os.stat(target)
                    os.utime(target, (st.st_atime, epoch))
                except OSError as exc:
                    res.setdefault("reason", f"文件日期修改失败：{exc}")
                else:
                    res["date_set"] = True
                    batch["dates"].append({"path": str(target),
                                           "old_mtime": st.st_mtime, "new_mtime": epoch})

        # 2) 工作表改名：旧 title=旧 label 清洗值，多页含「-第N页/-页码」后缀
        if wb is not None:
            old_base = _clean_label(old_label)
            for ws in list(wb.worksheets):
                title = ws.title
                if title in _AUX_SHEET_TITLES:
                    continue
                if title == old_base:
                    suffix_part = ""
                elif old_base and title.startswith(old_base + "-"):
                    suffix_part = title[len(old_base):]
                else:
                    continue
                used.discard(title)
                new_title = _sheet_title(new_name + suffix_part, used)
                if new_title != title:
                    ws.title = new_title
                    sheet_changed = True
                    res["sheet_renamed"] = True
                    batch["sheets"].append(
                        {"path": str(wb_path), "old": title, "new": new_title})
        if old_label and old_label != new_name:
            batch["labels"].append({"old": old_label, "new": new_name})
        results.append(res)

    if wb is not None and sheet_changed:
        try:
            wb.save(wb_path)
        except Exception as exc:  # noqa: BLE001
            return JSONResponse({"detail": f"工作表改名保存失败：{exc}"}, status_code=500)
    synced = _sync_job_labels(job_id, {lb["old"]: lb["new"] for lb in batch["labels"]})
    undoable = synced or bool(batch["disk"] or batch["sheets"] or batch["dates"])
    if undoable:
        _RENAME_HISTORY.append(batch)
        del _RENAME_HISTORY[:-_RENAME_HISTORY_MAX]
        _save_rename_history()
    return {"ok": True, "results": results, "undoable": undoable}


@app.post("/api/rename/undo")
def rename_undo() -> dict:
    """撤销上一批命名：磁盘改名逐条改回（目标被占用则跳过），工作表名也改回。"""
    if not _RENAME_HISTORY:
        return {"ok": False, "error": "没有可撤销的命名"}
    batch = _RENAME_HISTORY.pop()
    _save_rename_history()
    # 先还原文件日期再改回文件名：日期记录的是改名后的路径，改回后该路径就不存在了
    dates_restored = 0
    for rec in batch.get("dates") or []:
        target = Path(rec["path"])
        try:
            if target.is_file():
                st = os.stat(target)
                os.utime(target, (st.st_atime, rec["old_mtime"]))
                dates_restored += 1
        except OSError:
            continue
    disk_results = []
    for rec in reversed(batch["disk"]):
        old, new = rec["old"], rec["new"]
        src, dst = Path(new), Path(old)
        entry = {"old_path": new, "new_path": old, "restored": False}
        if not src.is_file():
            entry["reason"] = "改名后的文件不存在"
        elif dst.exists():
            entry["reason"] = "原文件名已被占用，跳过"
        else:
            try:
                os.rename(src, dst)
                ALLOWED_FILES.discard(str(src))
                ALLOWED_FILES.add(str(dst))
                entry["restored"] = True
            except OSError as exc:
                entry["reason"] = f"还原失败：{exc}"
        disk_results.append(entry)

    by_path: dict[str, list] = {}
    for rec in batch["sheets"]:
        by_path.setdefault(rec["path"], []).append(rec)
    sheet_restored = 0
    allowed_outputs = _output_paths()
    for p, recs in by_path.items():
        wp = Path(p)
        # 工作簿只认产物白名单：历史可能来自落盘文件，不能只凭记录里的路径就改表名
        if not wp.is_file() or wp not in allowed_outputs:
            continue
        try:
            from openpyxl import load_workbook
            wb = load_workbook(wp)
        except Exception:  # noqa: BLE001
            continue
        sheets = {ws.title: ws for ws in wb.worksheets}
        used = set(wb.sheetnames)
        changed = False
        for rec in recs:
            ws = sheets.get(rec["new"])
            if ws is None:
                continue
            used.discard(rec["new"])
            if rec["old"] in used:
                continue  # 目标表名被占用，跳过
            ws.title = rec["old"]
            used.add(rec["old"])
            sheets[rec["old"]] = ws
            changed = True
            sheet_restored += 1
        if changed:
            try:
                wb.save(wp)
            except Exception:  # noqa: BLE001
                pass
    _sync_job_labels(batch.get("job_id", ""),
                     {lb["new"]: lb["old"] for lb in batch.get("labels") or []})
    return {"ok": True, "disk": disk_results, "sheet_restored": sheet_restored,
            "dates_restored": dates_restored}


_load_rename_history()


@app.post("/api/cancel")
async def convert_cancel(payload: dict) -> dict:
    job = JOBS.get(payload.get("job_id", ""))
    if job:
        job.cancelled = True
        # 识别中可能正阻塞在 OCR 调用上，主动结束子进程让取消在数秒内生效
        _interrupt_job(job)
    return {"ok": True}


@app.websocket("/ws/{job_id}")
async def job_events(websocket: WebSocket, job_id: str) -> None:
    if not _ws_authorized(websocket):
        await websocket.close(code=1008)
        return
    await websocket.accept()
    job = JOBS.get(job_id)
    if job is None:
        await websocket.close()
        return
    try:
        while True:
            try:
                event = job.q.get_nowait()
            except queue.Empty:
                if job.done and job.q.empty():
                    break
                await asyncio.sleep(0.08)
                continue
            if event.get("type") == "_end":
                break
            await websocket.send_json(event)
    except WebSocketDisconnect:
        pass
    finally:
        JOBS.pop(job_id, None)


@app.get("/api/download")
async def download(path: str) -> FileResponse:
    target = Path(path)
    allowed = _output_paths()
    if (not target.is_file() or target.suffix.lower() != ".xlsx"
            or _resolve_path(target) not in allowed):
        return JSONResponse({"detail": "文件不存在或未授权下载"}, status_code=404)
    return FileResponse(
        str(target),
        filename=target.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# 附加工作表：不属于用户数据，预览与结构化导出时跳过
_AUX_SHEET_NAMES = ("待复核清单", "待复核-低优先", "拍摄质量", "单号索引", "索引异常")
_TABLE_ROWS_MAX = 500
_CSV_RISK_PREFIX = ("=", "+", "-", "@")


def _cell_text(value) -> str:
    """把 Excel 单元格值转成可展示的纯文本（None -> 空串）。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _authorized_xlsx(raw: str) -> Path | None:
    """校验产物路径：必须存在、是 xlsx，且在本进程产物/任务历史白名单内。"""
    target = Path(raw or "")
    if not target.is_file() or target.suffix.lower() != ".xlsx":
        return None
    if _resolve_path(target) not in _output_paths():
        return None
    return target


def _data_sheet_names(sheets: list[str]) -> list[str]:
    return [name for name in sheets if name not in _AUX_SHEET_NAMES]


@app.get("/api/result-table")
def result_table(path: str, sheet: str = "", offset: int = 0,
                 limit: int = 300) -> dict:
    """只读读取产物数据页，供结果卡内直接核对（不需要打开 Excel）。

    工作簿路径仍走产物白名单，避免被当作任意文件读取接口。
    """
    target = _authorized_xlsx(path)
    if target is None:
        return JSONResponse({"detail": "文件不存在或未授权读取"}, status_code=404)
    from openpyxl import load_workbook

    try:
        wb = load_workbook(target, read_only=True, data_only=False)
    except Exception as exc:  # 损坏/占用等
        return JSONResponse({"detail": f"无法读取工作簿：{exc}"}, status_code=400)
    try:
        sheets = list(wb.sheetnames)
        data_sheets = _data_sheet_names(sheets)
        name = sheet if sheet in sheets else (
            data_sheets[0] if data_sheets else (sheets[0] if sheets else ""))
        offset = max(0, int(offset or 0))
        limit = max(1, min(int(limit or 300), _TABLE_ROWS_MAX))
        rows: list[list[str]] = []
        total_rows = 0
        total_cols = 0
        if name:
            ws = wb[name]
            total_rows = int(ws.max_row or 0)
            total_cols = int(ws.max_column or 0)
            for row in ws.iter_rows(min_row=offset + 1,
                                    max_row=offset + limit, values_only=True):
                rows.append([_cell_text(v) for v in row])
        return {
            "ok": True,
            "sheets": sheets,
            "data_sheets": data_sheets,
            "sheet": name,
            "offset": offset,
            "limit": limit,
            "total_rows": total_rows,
            "total_cols": total_cols,
            "rows": rows,
            "truncated": offset + len(rows) < total_rows,
        }
    finally:
        wb.close()


@app.get("/api/export-structured")
def export_structured(path: str, fmt: str = "json", sheet: str = "") -> Response:
    """把识别结果导成结构化文件（JSON / CSV），便于二次处理。

    fmt=json 时导出全部数据页；fmt=csv 时只导出 sheet 指定（或首个）数据页。
    文件名与下载接口保持一致，中文名按 RFC 5987 编码。
    """
    target = _authorized_xlsx(path)
    if target is None:
        return JSONResponse({"detail": "文件不存在或未授权导出"}, status_code=404)
    fmt = (fmt or "json").lower()
    if fmt not in ("json", "csv"):
        return JSONResponse({"detail": "仅支持 json / csv"}, status_code=400)
    from openpyxl import load_workbook

    try:
        wb = load_workbook(target, read_only=True, data_only=False)
    except Exception as exc:
        return JSONResponse({"detail": f"无法读取工作簿：{exc}"}, status_code=400)
    try:
        sheets = list(wb.sheetnames)
        data_sheets = _data_sheet_names(sheets)
        payload: dict[str, list[list[str]]] = {}
        for name in data_sheets:
            ws = wb[name]
            payload[name] = [[_cell_text(v) for v in row]
                             for row in ws.iter_rows(values_only=True)]
        if not payload:
            return JSONResponse({"detail": "工作簿内没有数据页"}, status_code=400)
    finally:
        wb.close()

    stem = target.stem
    if fmt == "json":
        body = json.dumps({"source": target.name, "sheets": payload},
                          ensure_ascii=False, indent=2)
        media = "application/json; charset=utf-8"
        filename = f"{stem}.json"
    else:
        import csv
        import io

        name = sheet if sheet in payload else next(iter(payload))
        buf = io.StringIO()
        writer = csv.writer(buf, lineterminator="\r\n")
        for row in payload[name]:
            writer.writerow([
                "'" + cell if cell.startswith(_CSV_RISK_PREFIX) else cell
                for cell in row
            ])
        body = "\ufeff" + buf.getvalue()
        media = "text/csv; charset=utf-8"
        filename = f"{stem}_{name}.csv" if len(payload) > 1 else f"{stem}.csv"

    return Response(
        content=body,
        media_type=media,
        headers={
            "Content-Disposition":
                f"attachment; filename*=UTF-8''{quote(filename)}",
        },
    )


@app.post("/api/reveal")
def reveal(payload: dict) -> dict:
    """桌面场景下直接用默认程序打开产物，或在资源管理器里定位到它。

    浏览器形态下用户习惯点"下载"，但桌面壳里文件其实已经在本地，
    再走一次下载反而是绕路（还要多存一份、多找一次路径）。

    只允许打开本次/最近任务的产物目录与用户用系统文件框选择的文件；
    "用默认程序打开"进一步限制为 Excel / PDF / 图片，避免被当作启动器。
    """
    raw = str(payload.get("path") or "").strip()
    if not raw:
        return {"ok": False, "error": "缺少路径"}
    target = Path(raw)
    if not _is_allowed_target(target):
        return {"ok": False, "error": "该路径不在程序允许访问的范围内"}
    if not target.exists():
        return {"ok": False, "error": "文件不存在或已被移动"}
    mode = str(payload.get("mode") or "select")
    if mode == "open" and target.is_file() and target.suffix.lower() not in _OPENABLE_EXTS:
        return {"ok": False, "error": "仅支持打开 Excel、PDF、图片或目录"}
    try:
        if mode == "open" or target.is_dir():
            os.startfile(str(target))                      # noqa: S606
        elif os.name == "nt":
            subprocess.Popen(["explorer", "/select,", str(target)])
        else:
            return {"ok": False, "error": "当前系统不支持定位文件"}
    except Exception as exc:
        return {"ok": False, "error": str(exc)}
    return {"ok": True}


def _norm_search_text(text: str) -> str:
    """搜索归一化：忽略大小写、空格、连字符，便于粘贴带格式的单号。"""
    return re.sub(r"[\s\-—_]+", "", str(text or "")).upper()


def _normalize_waybill(text: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(text or "").upper())


def _edit_distance_at_most_one(a: str, b: str) -> bool:
    """Bounded edit distance for a one-character OCR slip."""
    if abs(len(a) - len(b)) > 1:
        return False
    if a == b:
        return True
    if len(a) > len(b):
        a, b = b, a
    i = j = edits = 0
    while i < len(a) and j < len(b):
        if a[i] == b[j]:
            i += 1
            j += 1
            continue
        edits += 1
        if edits > 1:
            return False
        if len(a) == len(b):
            i += 1
            j += 1
        else:
            j += 1
    if j < len(b) or i < len(a):
        edits += 1
    return edits <= 1


_WAYBILL_TOKEN_RE = re.compile(r"(?:R\d{2}T)?YT[A-Z0-9]{8,20}|[A-Z0-9]{11,23}", re.I)
_CONFUSABLE_DIGITS = str.maketrans({"O": "0", "Q": "0", "D": "0", "I": "1",
                                   "L": "1", "Z": "2", "S": "5", "B": "8", "G": "6"})


def _waybill_variants(value: str) -> set[str]:
    value = _normalize_waybill(value)
    variants = {value}
    match = re.match(r"^R\d{2}T?YT", value)
    if match:
        variants.add(value[match.end() - 2:])
    return {v for v in variants if v}


_TAIL_MATCH_MIN_LEN = 10      # 太短的号码只看后四位容易误命中
_TAIL_MATCH_LEN_TOLERANCE = 2  # 允许 OCR 多读/少读 1–2 位


def _match_waybill_text(target: str, text: str) -> tuple[str, str]:
    """Return (exact/tail/suspect/empty, OCR candidate) for a pasted waybill.

    匹配优先级：
      1. exact —— 整号对得上，直接命中，不需要复核；
      2. tail  —— 整号对不上但后四位一致，按“命中但需复核”处理；
      3. suspect —— 字形混淆/编辑距离 1 等旧逻辑，仍需复核。
    """
    query = _normalize_waybill(target)
    raw = str(text or "")
    compact = _normalize_waybill(raw)
    if not query or not compact:
        return "", ""
    tokens = _WAYBILL_TOKEN_RE.findall(compact)
    candidates = tokens or [compact]
    for candidate in candidates:
        if query in candidate:
            return "exact", candidate
    # 后四位命中：整号没对上时优先看尾号，命中后交给人工复核
    if len(query) >= _TAIL_MATCH_MIN_LEN:
        tail = query[-4:]
        query_has_alpha = bool(re.search(r"[A-Z]", query))
        for candidate in candidates:
            for value in _waybill_variants(candidate):
                if not value.endswith(tail):
                    continue
                if len(value) < _TAIL_MATCH_MIN_LEN:
                    continue
                if abs(len(value) - len(query)) > _TAIL_MATCH_LEN_TOLERANCE:
                    continue
                # 带字母的单号不应被纯数字商品条码蹭中
                if query_has_alpha and not re.search(r"[A-Z]", value):
                    continue
                return "tail", candidate
    q_variants = _waybill_variants(query)
    for candidate in candidates:
        c_variants = _waybill_variants(candidate)
        if any(q == c for q in q_variants for c in c_variants):
            return "suspect", candidate
        for q in q_variants:
            for c in c_variants:
                # Numeric OCR commonly confuses these glyphs; keep these results
                # separate from exact hits so they are never silently trusted.
                if q.translate(_CONFUSABLE_DIGITS) == c.translate(_CONFUSABLE_DIGITS):
                    return "suspect", candidate
                if (q.startswith("YT") and _edit_distance_at_most_one(q, c)
                        and (c.startswith(("YT", "TT")) or c == q[1:])):
                    # 含 Y->T 前缀错读、左余量切掉首 Y（c==q[1:]）两类：
                    # 编辑距离 1 但不满足双 YT 前缀，补进疑似交给复认
                    return "suspect", candidate
    return "", ""


def _reread_rotated_crop(worker, image_path: str, box, rot: int, t: float,
                         pad: float, scale: float):
    """竖排（整页旋转 90/270）命中：全分辨率下 DB 对竖排文本只返回字级碎框，
    必须先裁出疑似区域并转正再复读。返回 (复读框, 疑似框在转正裁剪图里的坐标)。"""
    from PIL import Image
    x1, y1, x2, y2 = box
    w, h = x2 - x1, y2 - y1
    cx1, cy1 = int(x1 - 0.2 * t), int(y1 - pad * t)
    cx2, cy2 = int(x2 + 0.2 * t), int(y2 + pad * t)
    tmp = None
    try:
        with Image.open(image_path) as im:
            crop = im.crop((max(0, cx1), max(0, cy1), cx2, cy2))
        crop = crop.transpose(Image.ROTATE_270 if rot == 90 else Image.ROTATE_90)
        cw, ch = crop.size
        fd, tmp = tempfile.mkstemp(suffix=".png", prefix="suspect_reread_")
        os.close(fd)
        crop.save(tmp)
        boxes = worker.ocr(tmp, scale=scale)
        # 原图裁剪框含边距；转正后疑似框位置按旋转方向换算
        pl, pt = x1 - max(0, cx1), y1 - max(0, cy1)
        if rot == 90:
            rect = (cw - pt - h, pl, cw - pt, pl + w)
        else:
            rect = (pt, ch - pl - w, pt + h, ch - pl)
        return boxes, rect
    except Exception:
        return [], None
    finally:
        if tmp:
            try:
                os.unlink(tmp)
            except OSError:
                pass


def _upgrade_suspect_hits(worker, image_path: str, hits: list, target: str, log=None) -> None:
    """疑似命中做高分辨率局部复认：扫描尺度错读一位的单号，原图局部 2/3 倍
    复读常能读对。仅当复读框与疑似框同位且精确匹配才升级「已找到」，
    否则保持疑似，绝不静默信任。"""
    for hit in hits:
        # 待复核（后四位一致）也复认一次：能读全就升级为“已找到”，免人工复核
        if hit.get("status") not in ("疑似错读", "待复核") or not hit.get("box"):
            continue
        x1, y1, x2, y2 = [float(v) for v in hit["box"]]
        t = max(1.0, min(x2 - x1, y2 - y1))
        rot = int(hit.get("rotation") or 0)
        for scale, pad in ((2.0, 0.6), (3.0, 1.0)):
            if rot in (90, 270):
                boxes, rect = _reread_rotated_crop(worker, image_path,
                                                   (x1, y1, x2, y2), rot, t, pad, scale)
                if rect is None:
                    break
            else:
                clip = (int(x1 - 0.2 * t), int(y1 - 0.2 * t),
                        int(x2 + pad * t), int(y2 + 0.2 * t))
                try:
                    boxes = worker.ocr(image_path, clip=clip, scale=scale)
                except Exception:
                    break
                rect = (x1, y1, x2, y2)
            sx1, sy1, sx2, sy2 = rect
            for b in boxes or []:
                bx1, by1, bx2, by2 = b.get("bbox") or [0, 0, 0, 0]
                cx, cy = (bx1 + bx2) / 2.0, (by1 + by2) / 2.0
                if not (sx1 - 0.2 * t <= cx <= sx2 + 0.2 * t
                        and sy1 - 0.2 * t <= cy <= sy2 + 0.2 * t):
                    continue
                kind, matched = _match_waybill_text(target, b.get("text", ""))
                if kind == "exact":
                    hit["status"] = "已找到"
                    hit["suspicious"] = False
                    hit["waybill"] = matched
                    hit["note"] = ""
                    hit["score"] = b.get("score", hit.get("score", 0))
                    if log:
                        log(f"复认后完整命中（局部 {scale:.0f} 倍复读）：{matched}")
                    break
            if hit["status"] == "已找到":
                break


def _search_hit(entry: dict, store_id: str, images: dict) -> dict:
    """把索引明细转成前端可渲染的一条命中（单号查询与批量查询共用）。"""
    index = entry.get("image_index")
    info = (images or {}).get(str(index)) if index is not None else None
    return {
        "waybill": entry.get("waybill", ""),
        "status": entry.get("status", ""),
        "source": entry.get("source", ""),
        "page": entry.get("page", 0),
        "sheet": entry.get("sheet", ""),
        "cell": entry.get("cell", ""),
        "row": entry.get("row", 0),
        "summary": entry.get("summary", ""),
        "note": entry.get("note", ""),
        "duplicate": bool(entry.get("duplicate")),
        "suspicious": bool(entry.get("suspicious")),
        # 原图定位：框为原图像素坐标，缩略图按同比例缩放
        "box": entry.get("box") or None,
        "image_width": int(info["width"]) if info else 0,
        "image_height": int(info["height"]) if info else 0,
        "image_url": (f"/api/job-image?job_id={quote(store_id)}&idx={int(index)}"
                      if info and entry.get("box") else ""),
    }


@app.get("/api/search")
async def search_waybills(q: str = "", job_id: str = "") -> dict:
    """应用内单号搜索：在当前（或最近一次）转换的单号索引里做包含匹配。

    只查内存中的索引明细，不读 Excel、不发网络请求；无匹配时如实返回
    "没有结果"，并附带索引异常数量，避免用户误以为"这个单号肯定不存在"。
    """
    item = SEARCH_STORE.get(job_id) if job_id else None
    if item is None and job_id:
        return {"ok": False, "error": "该任务的索引已过期（只保留最近 5 次转换结果）",
                "matches": [], "total": 0, "scanned": 0, "issues": 0}
    store_id = job_id
    if item is None and SEARCH_STORE:
        store_id = next(reversed(SEARCH_STORE))
        item = SEARCH_STORE[store_id]
    if item is None:
        return {"ok": False, "error": "还没有完成过转换，暂时没有可查询的单号索引",
                "matches": [], "total": 0, "scanned": 0, "issues": 0}
    entries = item["entries"]
    needle = _norm_search_text(q)
    if not needle:
        return {"ok": True, "mode": item.get("mode", "full"),
                "total": 0, "scanned": item.get("scanned", len(entries)),
                "issues": len(item["issues"]), "matches": [],
                "out_path": item["out_path"]}
    if item.get("mode") == "find":
        target = _normalize_waybill(item.get("target", ""))
        if needle and _normalize_waybill(needle) != target:
            return {"ok": True, "mode": "find",
                    "total": 0, "scanned": item.get("scanned", len(entries)),
                    "issues": 0, "matches": [], "out_path": ""}
    matches = []
    images = item.get("images") or {}
    for entry in entries:
        matched = (item.get("mode") == "find"
                   or needle in _norm_search_text(entry.get("waybill", "")))
        if matched:
            matches.append(_search_hit(entry, store_id, images))
    return {"ok": True, "mode": item.get("mode", "full"),
            "total": len(matches), "scanned": item.get("scanned", len(entries)),
            "issues": len(item["issues"]), "matches": matches[:300],
            "out_path": item["out_path"]}


_BATCH_SPLIT_RE = re.compile(r"[\r\n,，;；]+")


def _split_batch_queries(text: str) -> list[str]:
    """批量查询输入：按行/逗号/分号拆分，忽略大小写、空格与连字符并去重。"""
    seen: set[str] = set()
    queries: list[str] = []
    for raw in _BATCH_SPLIT_RE.split(str(text or "")):
        norm = _norm_search_text(raw)
        if not norm or norm in seen:
            continue
        seen.add(norm)
        queries.append(norm)
        if len(queries) >= 500:  # 防止整份日志误粘进来拖垮页面
            break
    return queries


@app.post("/api/search/batch")
async def search_waybills_batch(payload: dict) -> dict:
    """批量单号查询：一次返回「命中 / 未命中 / 格式异常」三组结果。

    与单个查询共用同一份内存索引与命中结构，不读 Excel、不发网络请求。
    """
    job_id = str(payload.get("job_id") or "")
    text = str(payload.get("text") or "")
    item = SEARCH_STORE.get(job_id) if job_id else None
    if item is None and job_id:
        return {"ok": False, "error": "该任务的索引已过期（只保留最近 5 次转换结果）",
                "groups": [], "exact": 0, "partial": 0, "none": 0, "invalid": 0}
    store_id = job_id
    if item is None and SEARCH_STORE:
        store_id = next(reversed(SEARCH_STORE))
        item = SEARCH_STORE[store_id]
    if item is None:
        return {"ok": False, "error": "还没有完成过转换，暂时没有可查询的单号索引",
                "groups": [], "exact": 0, "partial": 0, "none": 0, "invalid": 0}

    queries = _split_batch_queries(text)
    if not queries:
        return {"ok": False, "error": "请输入至少一个单号（每行一个，或用逗号分隔）",
                "groups": [], "exact": 0, "partial": 0, "none": 0, "invalid": 0}

    entries = item["entries"]
    images = item.get("images") or {}
    groups = []
    counts = {"exact": 0, "partial": 0, "none": 0, "invalid": 0}
    for query in queries:
        if len(query) < 6 or not re.search(r"[A-Z0-9]", query):
            counts["invalid"] += 1
            groups.append({"query": query, "status": "invalid",
                           "reason": "长度不足 6 位或不含字母数字", "total": 0, "hits": []})
            continue
        hits = []
        for entry in entries:
            waybill = _norm_search_text(entry.get("waybill", ""))
            if waybill and (query == waybill or query in waybill):
                hits.append(_search_hit(entry, store_id, images))
        exact = any(_norm_search_text(h.get("waybill", "")) == query for h in hits)
        status = "exact" if exact else ("partial" if hits else "none")
        counts[status] += 1
        groups.append({"query": query, "status": status, "total": len(hits),
                       "hits": hits[:20]})
    return {"ok": True, "mode": item.get("mode", "full"),
            "scanned": item.get("scanned", len(entries)),
            "issues": len(item["issues"]), "out_path": item["out_path"],
            "groups": groups, "query_count": len(queries), **counts}


@app.get("/api/jobs")
async def list_job_history(limit: int = 50) -> dict:
    """最近任务列表（M4）：只返回摘要字段，不带 flags 明细，明细走 /api/job。"""
    try:
        limit = max(1, min(job_store.MAX_KEEP, int(limit)))
    except (TypeError, ValueError):
        limit = job_store.MAX_KEEP
    jobs = []
    for meta in job_store.list_jobs(limit=limit):
        job_id = str(meta.get("job_id") or "")
        out_path = str(meta.get("out_path") or "")
        mode = str(meta.get("mode") or "")
        jobs.append({
            "job_id": job_id,
            "time": str(meta.get("time") or ""),
            "saved_at": meta.get("saved_at"),
            "mode": mode,
            "files": meta.get("files") or [],
            "total_pages": int(meta.get("total_pages") or meta.get("scanned") or 0),
            "waybills": int(meta.get("unique_waybills") or meta.get("waybills")
                            or meta.get("matches") or 0),
            "flags_total": int(meta.get("flags_total") or 0),
            "quality_level": str(meta.get("quality_level") or ""),
            "quality_score": meta.get("quality_score"),
            "out_path": out_path,
            "out_exists": bool(out_path) and Path(out_path).is_file(),
            "has_previews": job_store.preview_dir(job_id).is_dir(),
            # 只有全量识别有待复核工作台（纯单号/查找没有原图红框定位）
            "has_flags": mode == "full" and bool(meta.get("flags")),
        })
    return {"ok": True, "jobs": jobs}


@app.post("/api/jobs/clear")
async def clear_job_history() -> dict:
    """清除本机任务历史（设置页入口）；内存里的搜索索引一并失效。"""
    removed = job_store.clear_all()
    SEARCH_STORE.clear()
    return {"ok": True, "removed": removed}


@app.get("/api/job")
async def get_job(job_id: str = "") -> dict:
    """读取本机任务快照（M4 任务历史用；待复核红框也复用它）。"""
    meta = job_store.load_job(job_id)
    if not meta:
        return {"ok": False, "error": "任务不存在或已清理"}
    return {"ok": True, "job": meta}


@app.post("/api/job/flags")
async def save_job_flags(payload: dict) -> dict:
    """保存待复核处理状态（只写本机任务目录，不联网）。"""
    job_id = str(payload.get("job_id") or "")
    if not job_id:
        return {"ok": False, "error": "缺少任务 id"}
    ok = job_store.save_flag_states(job_id, payload.get("states") or {})
    return {"ok": ok}


@app.get("/api/job/flags")
async def load_job_flags(job_id: str = "") -> dict:
    return {"ok": True, "states": job_store.load_flag_states(job_id)}


def _locate_job_image(job_id: str, idx: int) -> Path | None:
    """定位某次任务某页的原图缩略图（不接受任意路径参数）。

    优先取内存索引；程序重启或索引被更晚任务挤出后，回退到本机任务目录
    （job_store）里的磁盘缩略图。job_id 经 job_store 白名单清洗，页下标
    只用来拼固定文件名。
    """
    item = SEARCH_STORE.get(job_id)
    if item is not None:
        info = (item.get("images") or {}).get(str(idx))
        if info:
            candidate = Path(info["path"])
            if candidate.is_file():
                return candidate
    if idx >= 0:
        candidate = job_store.preview_dir(job_id) / f"{idx}.jpg"
        if candidate.is_file():
            return candidate
    return None


def _finite_number(value) -> float | None:
    """把 JSON 数值转成有限浮点数；bool / 非数字 / NaN / Inf 均视为无效。"""
    if isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _mark_highlight_box(image, box, image_width, image_height) -> None:
    """按原图像素坐标在缩略图上描边高亮；参数无效时保持干净原图。"""
    if not isinstance(box, (list, tuple)) or len(box) != 4:
        return
    coords = [_finite_number(value) for value in box]
    if any(value is None for value in coords):
        return
    src_w = _finite_number(image_width)
    src_h = _finite_number(image_height)
    if not src_w or not src_h or src_w <= 0 or src_h <= 0:
        return
    from PIL import ImageDraw

    width, height = image.size
    # 缩略图与原始比例一致，按缩放比把原图坐标换算到缩略图坐标
    sx, sy = width / src_w, height / src_h
    left, right = sorted((coords[0] * sx, coords[2] * sx))
    top, bottom = sorted((coords[1] * sy, coords[3] * sy))
    left = max(0, int(round(left)))
    top = max(0, int(round(top)))
    right = min(width - 1, int(round(right)))
    bottom = min(height - 1, int(round(bottom)))
    if right <= left or bottom <= top:
        return
    # 线宽随图片尺寸自适应，小图也至少 2px，保证粘贴后仍能看清
    line = max(2, int(round(min(width, height) * 0.0025)))
    ImageDraw.Draw(image).rectangle((left, top, right, bottom),
                                    outline="#E6007E", width=line)


@app.get("/api/job-image")
async def job_image(job_id: str = "", idx: int = -1):
    """返回某次转换里某一页的原图缩略图（不接受任意路径）。"""
    path = _locate_job_image(job_id, idx)
    if path is None:
        return JSONResponse({"detail": "该任务的原图已清理，请重新转换后再定位"}, status_code=404)
    return FileResponse(path, media_type="image/jpeg")


@app.post("/api/copy-image")
def copy_image(payload: dict) -> dict:
    """把某页原图缩略图（可带高亮框）复制到系统剪贴板。

    同步 def → 由线程池执行：读图与 -STA 写剪贴板都会阻塞。
    """
    if os.name != "nt":
        return {"ok": False, "error": "当前系统不支持写入剪贴板（仅 Windows 可用）"}
    job_id = str(payload.get("job_id") or "")
    if not job_id:
        return {"ok": False, "error": "缺少任务 id"}
    try:
        idx = int(payload.get("idx"))
    except (TypeError, ValueError):
        return {"ok": False, "error": "缺少有效的页下标"}
    if idx < 0:
        return {"ok": False, "error": "页下标无效"}
    path = _locate_job_image(job_id, idx)
    if path is None:
        return {"ok": False, "error": "该任务的原图已清理，请重新转换后再复制"}

    tmp_path: Path | None = None
    try:
        from PIL import Image

        with Image.open(path) as src:
            image = src.convert("RGB")
        _mark_highlight_box(image, payload.get("box"),
                            payload.get("image_width"),
                            payload.get("image_height"))
        COPY_IMAGE_DIR.mkdir(parents=True, exist_ok=True)
        tmp_path = COPY_IMAGE_DIR / f"copy_{secrets.token_hex(16)}.png"
        image.save(tmp_path, format="PNG")

        script = (
            "Add-Type -AssemblyName System.Windows.Forms;"
            "Add-Type -AssemblyName System.Drawing;"
            "$img = [System.Drawing.Image]::FromFile($env:IMG2DOC_FILE_TARGET);"
            "try { [System.Windows.Forms.Clipboard]::SetImage($img) }"
            " finally { $img.Dispose() }"
        )
        # 剪贴板可能被其他程序短暂占用（OpenClipboard 失败 / 0x800401D0），
        # 与 /api/file-op 一致重试 3 次，间隔 0.4s。
        detail = ""
        for attempt in range(3):
            ok, detail = _run_powershell(script, target=str(tmp_path))
            if ok:
                return {"ok": True}
            lowered = detail.lower()
            if ("剪贴板" not in detail and "clipboard" not in lowered
                    and "0x800401d0" not in lowered):
                break
            if attempt < 2:
                time.sleep(0.4)
        lowered = detail.lower()
        if ("剪贴板" in detail or "clipboard" in lowered
                or "0x800401d0" in lowered):
            detail = "剪贴板正被其他程序占用，请稍后再试"
        return {"ok": False, "error": detail or "复制到剪贴板失败"}
    except Exception as exc:  # noqa: BLE001 - 接口返回中文原因，避免 500
        return {"ok": False, "error": f"生成剪贴板图片失败：{exc}"}
    finally:
        # SetImage 已把位图写入剪贴板，临时 PNG 可以立即删除
        if tmp_path is not None:
            try:
                tmp_path.unlink(missing_ok=True)
            except OSError:
                pass


# ---------------------------------------------------------------- 启动
def _open_browser(port: int) -> subprocess.Popen | None:
    """用 Edge/Chrome 的 app 模式打开主界面（无地址栏，等同桌面应用）。

    返回浏览器进程句柄：调用方可据此在窗口关闭后退出本地服务。
    使用独立的 user-data-dir，避免浏览器已在运行时 --app 被现有进程接管
    （否则 Popen 会立即退出，导致误判窗口已关闭）。
    """
    url = f"http://{HOST}:{port}/web/index.html"
    candidates = [
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
    ]
    profile = Path(os.environ.get("LOCALAPPDATA", str(ROOT))) / core.APP_TITLE / "webview-profile"
    for exe in candidates:
        if not exe.is_file():
            continue
        try:
            profile.mkdir(parents=True, exist_ok=True)
        except OSError:
            profile = ROOT / "webview-profile"
            try:
                profile.mkdir(parents=True, exist_ok=True)
            except OSError:
                pass
        try:
            return subprocess.Popen([
                str(exe),
                f"--app={url}",
                f"--user-data-dir={profile}",
                "--window-size=1120,860",
                "--no-first-run",
                "--no-default-browser-check",
            ])
        except OSError:
            continue
    webbrowser.open(url)
    return None


def _ensure_utf8_console() -> None:
    """Windows 控制台切到 UTF-8，保证 Python 打印的中文不乱码。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    except Exception:
        pass
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def _startup_log(text: str) -> None:
    """启动期日志：委托 core（windowed 程序无控制台）。"""
    core.startup_log(text)


def _ensure_streams() -> None:
    """stdout/stderr 兜底：委托 core。"""
    core.ensure_streams()


def main() -> None:
    _ensure_streams()
    _ensure_utf8_console()
    _startup_log("=" * 40 + " 启动")
    if not core.ensure_single_instance():
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(
                0,
                "程序已在运行中。\n请在任务栏或浏览器窗口中找到已打开的「肥鸭文档转换」。",
                core.APP_TITLE,
                0x40,
            )
        except Exception:
            pass
        return

    try:
        cleaned = cleanup_stale_temp_dirs()
        if cleaned:
            _startup_log(f"已清理过期临时目录 {cleaned} 个")
    except Exception:
        pass

    config = uvicorn.Config(app, host=HOST, port=PORT, log_level="warning")
    server = uvicorn.Server(config)

    def _run_server() -> None:
        try:
            server.run()
        except Exception:
            import traceback

            _startup_log("本地服务启动失败：\n" + traceback.format_exc())

    threading.Thread(target=_run_server, daemon=True).start()
    threading.Thread(target=_warm_fast_find_worker, name="find-ocr-warmup",
                     daemon=True).start()

    url = f"http://{HOST}:{PORT}/web/index.html"
    ready = False
    for _ in range(60):
        try:
            urllib.request.urlopen(url, timeout=1)
            ready = True
            break
        except Exception:
            time.sleep(0.25)
    _startup_log(f"服务就绪={ready} 地址={url}")

    print("=" * 52)
    print(" 肥鸭文档转换 已启动")
    print(f" 地址：{url}")
    print(" 操作界面将以独立窗口打开，关闭该窗口即退出程序。")
    print("=" * 52, flush=True)

    _open_browser(PORT)
    _startup_log("界面已打开，等待前端心跳…")

    # 关窗即退出：Edge 的 --app 进程常把窗口交接给已有浏览器进程后立刻退出，
    # 不能用进程存活判断。改由前端定时心跳（/api/ping）判断窗口是否还开着。
    first_wait = time.time()
    while True:
        time.sleep(2)
        if os.environ.get("IMG2DOC_DEBUG"):
            _startup_log(
                f"heartbeat seen={LAST_PING['seen']} idle={time.time() - LAST_PING['t']:.1f}s"
                f" threshold={IDLE_EXIT_SECONDS}s"
            )
        # 兜底：界面始终没连上（例如机器上找不到可用浏览器），不留下僵尸进程
        if not LAST_PING["seen"] and time.time() - first_wait > FIRST_PING_TIMEOUT:
            _startup_log(f"等待界面连接超过 {FIRST_PING_TIMEOUT}s，退出")
            break
        if LAST_PING["seen"] and time.time() - LAST_PING["t"] > IDLE_EXIT_SECONDS:
            _startup_log(f"前端心跳超过 {IDLE_EXIT_SECONDS}s 未更新，判定窗口已关闭，退出")
            break
    try:
        server.should_exit = True
    except Exception:
        pass
    global _FAST_FIND_WORKER
    if _FAST_FIND_WORKER is not None:
        try:
            _FAST_FIND_WORKER.close()
        except Exception:
            pass
        _FAST_FIND_WORKER = None


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback

        _startup_log("未捕获异常：\n" + traceback.format_exc())
        raise
