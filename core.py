# -*- coding: utf-8 -*-
"""肥鸭文档转换 —— 产品级共享常量与工具函数（不依赖任何 GUI 框架）。

桌面程序（desktop.py）与本地服务（web_app.py）都从这里取程序名、隐私文案、
配置读写、日志脱敏、OCR 运行环境定位等能力。

之所以独立成模块：打包 exe 时不需要再导入 GUI 框架，从而剔除 tkinter/tcl，
exe 体积更小、启动更快。
"""
from __future__ import annotations

import ctypes
import os
import re
import sys
import time
from pathlib import Path

from ai_calibrate import config_dir, load_config, save_config
from engine import EngineError, MineruEngine

# ---------------------------------------------------------------- 常量

APP_TITLE = "肥鸭文档转换"


def _read_version() -> str:
    """版本号唯一来源：项目根目录 VERSION 文件（打包后随包携带）。

    注意：此函数在模块导入期即被调用，不能依赖本文件后面才定义的函数。
    """
    candidates = []
    base = getattr(sys, "_MEIPASS", None)
    if base:
        candidates.append(Path(base) / "VERSION")
    candidates.append(Path(__file__).resolve().parent / "VERSION")
    try:
        candidates.append(Path(sys.executable).resolve().parent / "VERSION")
    except Exception:
        pass
    for p in candidates:
        try:
            if p.is_file():
                text = p.read_text(encoding="utf-8").strip()
                if text:
                    return text
        except Exception:
            continue
    return "1.1.2"


APP_VERSION = _read_version()
SUPPORTED_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".pdf"}
MODE_NAMES = {"waybill": "纯单号", "full": "图片转 Excel", "find": "查找单号"}
# 兼容旧配置/旧客户端：旧版模式统一归并为「全量识别」
LEGACY_MODE_ALIASES = {"exact": "full", "structure": "full", "table": "full"}


def normalize_mode(mode: str | None) -> str:
    """把模式名归一化到当前三种工作方式；无法识别时返回空串。"""
    value = str(mode or "").strip().lower()
    value = LEGACY_MODE_ALIASES.get(value, value)
    return value if value in MODE_NAMES else ""

# 隐私声明全文：关于弹窗与首次运行弹窗共用，与 privacy_notice.txt 保持一致
PRIVACY_NOTICE = (
    "· 本工具在您的电脑上本地运行，全部识别过程在本机完成，不依赖在线解析服务。\n"
    "· 「图片转 Excel」模式使用本地 OCR 引擎（PP-OCRv6，缺失时回退 v5）识别图片与 PDF 中的"
    "文字和表格，并把每页内容原样写成一个工作表。\n"
    "· 「单号清单导出」模式同样使用本机 OCR 引擎与表格网格重建，只导出运单号一列；"
    "「查找单号」模式用轻量本机识别，在当前选中的图片中查找用户输入的单号并定位原图。\n"
    "· 本工具不会将您的图片、PDF 或识别结果发送到任何外部服务，"
    "也不会上传或留存识别结果。\n"
    "· 例外（默认关闭，需您主动开启）：「设置 → 云端文本纠错」一旦开启，"
    "对识别存疑的单元格，程序会把该单元格的局部裁剪图与当前识别文字"
    "发送到 NVIDIA 托管模型接口（integrate.api.nvidia.com）做二次确认，"
    "用以校正商品名称等中文文本。运单号、日期、数量等编码字段不会外发。"
    "密钥由您自行提供，仅从环境变量或本机配置读取。云端纠错关闭时不会发送任何图片或文字。\n"
    "· 版本检查（默认开启，可在设置中关闭）：启动时只向 GitHub 公开的 Releases 接口查询"
    "最新版本号，发送程序版本与固定 User-Agent，不发送图片、文件名、识别结果或任何个人信息。\n"
    "· 转换结果与导出的 Excel 文件仅保存在您自行选择的位置。\n"
    "· 任务历史（默认保留最近 10 个，可在设置中调整）：每次转换完成后，程序会在本机 "
    "%LOCALAPPDATA%\\肥鸭文档转换\\jobs\\ 保存任务摘要、待复核条目快照、您的核对/重拍标记"
    "与每页原图缩略图，用于重启后继续核对与原图定位；只写本机，不联网。\n"
    "· 运行日志保存在 %LOCALAPPDATA%\\肥鸭文档转换\\logs，"
    "其中的运单号等编码类信息已自动中段打码。\n"
    "· 设置保存于 %APPDATA%\\肥鸭文档转换\\config.json。\n"
    "· 本软件以 MIT 许可证开源，未购买代码签名证书。首次启动如遇 Windows SmartScreen "
    "蓝色警告，请点击“更多信息 → 仍要运行”，或先运行随包附带的“解除锁定.bat”。"
)


# ---------------------------------------------------------------- 路径与文案

def resource_path(rel: str) -> Path:
    """开发环境相对项目根；打包后相对 _MEIPASS 解包目录。"""
    base = getattr(sys, "_MEIPASS", None)
    root = Path(base) if base else Path(__file__).resolve().parent
    return root / rel


# 日志脱敏：单号/编码类长串中段打码，保留前 4 后 2 便于定位
_MASK_PATTERNS = (
    re.compile(r"\b((?:R\d{2}T)?YT)\d{9,}(\d{2})\b"),        # 运单号
    re.compile(r"\b([A-Za-z]{1,8}\d{6})\d{4,}(\d{2})\b"),     # 其他字母数字混合编码
)


def mask_sensitive(text: str) -> str:
    """日志写入前脱敏：编码中段替换为 *，保留前后缀。"""
    if not text:
        return text
    out = text
    for pattern in _MASK_PATTERNS:
        out = pattern.sub(lambda m: f"{m.group(1)}{'*' * 6}{m.group(2)}", out)
    return out


def bundled_skill_dirs() -> list[Path]:
    """随包携带的 local-mineru（便携包/安装包放在 exe 同级）。"""
    candidates = []
    try:
        candidates.append(Path(sys.executable).resolve().parent / "local-mineru")
    except Exception:
        pass
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "local-mineru")
    candidates.append(resource_path("local-mineru"))
    return candidates


def find_skill_dir() -> Path | None:
    """定位本机已安装的旧版离线解析环境（仅作兼容回退）。

    查找顺序：显式环境变量 → 随包携带 → 本机公共安装目录。
    """
    bases = []
    if os.environ.get("MINERU_SKILL_DIR"):
        bases.append(Path(os.environ["MINERU_SKILL_DIR"]))
    bases += bundled_skill_dirs()
    bases += [Path(r"C:\ProgramData\local-mineru")]
    for base in bases:
        if (base / "scripts" / "server.py").is_file():
            return base
    return None


# ---------------------------------------------------------------- 日志

def log_dir() -> Path:
    """配置与日志目录（%APPDATA%\\肥鸭文档转换）。"""
    return Path(config_dir())


def local_log_dir() -> Path:
    """本机（非漫游）日志目录，用于启动期排障。"""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / APP_TITLE / "logs"


def startup_log(text: str) -> None:
    """启动期日志写到文件：打包成 windowed 程序后没有控制台可看。"""
    try:
        target = local_log_dir()
        target.mkdir(parents=True, exist_ok=True)
        with open(target / "web_start.log", "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {text}\n")
    except Exception:
        pass


def ensure_streams() -> None:
    """windowed 程序没有控制台，sys.stdout/stderr 是 None。

    uvicorn 初始化日志时会调用 sys.stderr.isatty()，为 None 会直接抛异常，
    因此这里兜底重定向到日志文件（失败则丢弃）。
    """
    if sys.stdout is not None and sys.stderr is not None:
        return
    stream = None
    try:
        target = local_log_dir()
        target.mkdir(parents=True, exist_ok=True)
        stream = open(target / "web_server.log", "a", encoding="utf-8", buffering=1)
    except Exception:
        try:
            stream = open(os.devnull, "w", encoding="utf-8")
        except Exception:
            return
    if sys.stdout is None:
        sys.stdout = stream
    if sys.stderr is None:
        sys.stderr = stream


def ensure_utf8_console() -> None:
    """Windows 控制台切到 UTF-8，保证打印的中文不乱码。"""
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    except Exception:
        pass
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ---------------------------------------------------------------- 单实例

def ensure_single_instance() -> bool:
    """Windows 命名互斥体单实例检测；已是第二个实例返回 False。

    避免多开时多个实例并发写同一个输出 Excel / 日志目录。
    互斥体句柄故意不关闭，随进程退出自动释放。
    """
    if sys.platform != "win32":
        return True
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW(None, False, "肥鸭文档转换_SingleInstance")
    return kernel32.GetLastError() != 183  # ERROR_ALREADY_EXISTS


def show_message(text: str, title: str = APP_TITLE, flags: int = 0x40) -> None:
    """无主窗口时的系统消息框（0x40 信息 / 0x10 错误）。"""
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, flags)
    except Exception:
        pass


# ---------------------------------------------------------------- OCR 运行环境

def bundled_runtime() -> Path | None:
    """便携包随附的 OCR runtime：优先 exe 同级目录，其次解包目录。"""
    candidates = []
    try:
        candidates.append(Path(sys.executable).resolve().parent / "runtime")
    except Exception:
        pass
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "runtime")
    candidates.append(resource_path("runtime"))
    for base in candidates:
        # 自包含 runtime（新版）：解释器直接摊平在 runtime\python.exe
        for rel in (("python.exe",), ("Scripts", "python.exe")):
            python = base.joinpath(*rel)
            if python.is_file():
                return python
    return None


def resolve_ocr_python(log=None) -> Path:
    """定位 OCR 运行环境的 python 解释器。

    查找顺序：
      1. 随包 runtime\\python.exe（自包含 CPU 推理环境）
      2. local-mineru 技能的 venv（还原文档结构模式安装的环境）
      3. ~/.openvino/venv/mineru 的默认路径
    """
    bundled = bundled_runtime()
    if bundled is not None:
        return bundled
    log = log or (lambda _t: None)
    skill_dir = find_skill_dir()
    if skill_dir is not None:
        try:
            return MineruEngine(skill_dir).ensure_env(log)
        except EngineError:
            pass
    home = Path(os.environ.get("USERPROFILE", str(Path.home())))
    default = home / ".openvino" / "venv" / "mineru" / "Scripts" / "python.exe"
    if default.is_file():
        return default
    raise EngineError(
        "未找到 OCR 运行环境。\n"
        "请确认安装目录下 runtime 目录完整，"
        "或按 README 的「从源码运行」准备运行时后重试。"
    )


__all__ = [
    "APP_TITLE", "APP_VERSION", "SUPPORTED_EXTS", "MODE_NAMES", "PRIVACY_NOTICE",
    "resource_path", "mask_sensitive", "find_skill_dir", "bundled_skill_dirs",
    "config_dir", "load_config", "save_config", "log_dir", "local_log_dir",
    "startup_log", "ensure_streams", "ensure_utf8_console",
    "ensure_single_instance", "show_message",
    "bundled_runtime", "resolve_ocr_python",
]
