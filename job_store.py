# -*- coding: utf-8 -*-
"""本机任务历史持久化（M2 底座）。

只写本机 AppData，不联网、不上传：

    %LOCALAPPDATA%\\肥鸭文档转换\\jobs\\<job_id>\\
        meta.json      # 任务摘要 + 待复核快照 + 按文件汇总
        previews\\      # 原图缩略图（供待复核定位与历史回看）

保留最近 N 个任务（默认 10，可在设置里调整），超出的整目录删除。
所有写入失败只返回 False/None，不向转换主流程抛异常。
"""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from pathlib import Path

from core import APP_TITLE

DEFAULT_KEEP = 10
MAX_KEEP = 50


def jobs_root() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / APP_TITLE / "jobs"


def _safe_id(job_id: str) -> str:
    """任务 id 只允许字母数字与下划线，避免任何路径拼接风险。"""
    return re.sub(r"[^0-9A-Za-z_-]", "", str(job_id or ""))[:64]


def job_dir(job_id: str) -> Path:
    return jobs_root() / _safe_id(job_id)


def preview_dir(job_id: str) -> Path:
    return job_dir(job_id) / "previews"


def save_job(job_id: str, meta: dict, keep: int = DEFAULT_KEEP) -> bool:
    """原子写入 meta.json，并按 keep 淘汰最旧任务。失败不抛异常。"""
    if not _safe_id(job_id):
        return False
    target = job_dir(job_id)
    try:
        target.mkdir(parents=True, exist_ok=True)
        payload = dict(meta or {})
        payload.setdefault("job_id", job_id)
        payload.setdefault("saved_at", time.time())
        tmp = target / "meta.json.tmp"
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(str(tmp), str(target / "meta.json"))
    except Exception:
        return False
    prune(keep)
    return True


def prune(keep: int = DEFAULT_KEEP) -> int:
    """只保留最近 keep 个带 meta.json 的任务目录，返回删除数量。"""
    try:
        keep = max(1, min(MAX_KEEP, int(keep)))
    except (TypeError, ValueError):
        keep = DEFAULT_KEEP
    try:
        root = jobs_root()
        if not root.is_dir():
            return 0
        entries = []
        for child in root.iterdir():
            meta = child / "meta.json"
            if child.is_dir() and meta.is_file():
                entries.append((meta.stat().st_mtime, child))
        entries.sort(key=lambda item: item[0], reverse=True)
    except Exception:
        return 0
    removed = 0
    for _, child in entries[keep:]:
        try:
            shutil.rmtree(child, ignore_errors=True)
            removed += 1
        except Exception:
            continue
    return removed


def clear_all() -> int:
    """删除全部本机任务历史（含预览图与待复核核对标记），返回删除的目录数。"""
    removed = 0
    try:
        root = jobs_root()
        if not root.is_dir():
            return 0
        for child in root.iterdir():
            if not child.is_dir():
                continue
            try:
                shutil.rmtree(child, ignore_errors=True)
                removed += 1
            except Exception:
                continue
    except Exception:
        return removed
    return removed


def load_job(job_id: str) -> dict | None:
    try:
        path = job_dir(job_id) / "meta.json"
        if not path.is_file():
            return None
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else None
    except Exception:
        return None


def list_jobs(limit: int = 20) -> list[dict]:
    """按保存时间倒序返回任务摘要，最多 limit 条。"""
    try:
        limit = max(1, int(limit))
    except (TypeError, ValueError):
        limit = 20
    try:
        root = jobs_root()
        if not root.is_dir():
            return []
        items = []
        for child in root.iterdir():
            meta_path = child / "meta.json"
            if not (child.is_dir() and meta_path.is_file()):
                continue
            try:
                data = json.loads(meta_path.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    items.append(data)
            except Exception:
                continue
        items.sort(key=lambda d: float(d.get("saved_at") or 0), reverse=True)
        return items[:limit]
    except Exception:
        return []


def clear_previews(job_id: str) -> None:
    """只清预览图，保留 meta.json（历史摘要仍可回看）。"""
    try:
        shutil.rmtree(preview_dir(job_id), ignore_errors=True)
    except Exception:
        pass


FLAG_STATES = {"checked", "retake"}


def _flag_states_path(job_id: str) -> Path:
    return job_dir(job_id) / "flag_states.json"


def save_flag_states(job_id: str, states: dict) -> bool:
    """保存待复核处理状态：{flag_index: "checked" | "retake"}。

    只接受白名单取值；写入失败不影响转换主流程。
    """
    if not _safe_id(job_id):
        return False
    cleaned = {}
    for key, value in (states or {}).items():
        if str(value) in FLAG_STATES:
            cleaned[str(key)] = str(value)
    try:
        target = job_dir(job_id)
        target.mkdir(parents=True, exist_ok=True)
        tmp = target / "flag_states.json.tmp"
        tmp.write_text(json.dumps(cleaned, ensure_ascii=False), encoding="utf-8")
        os.replace(str(tmp), str(_flag_states_path(job_id)))
        return True
    except Exception:
        return False


def load_flag_states(job_id: str) -> dict:
    try:
        path = _flag_states_path(job_id)
        if not path.is_file():
            return {}
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {}
        return {str(k): str(v) for k, v in data.items() if str(v) in FLAG_STATES}
    except Exception:
        return {}
