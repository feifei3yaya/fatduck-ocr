# -*- coding: utf-8 -*-
"""打包便携包：exe + runtime + OCR 模型 + 字库 -> 一个 zip。

用法：python build_package.py
产物：dist\\肥鸭文档转换_便携包.zip

前提：
  - dist\\肥鸭文档转换.exe 已由 build_exe.py 生成
  - runtime\\ 已由 build_runtime.py 生成（含可重定位的 CPU 推理环境）
  - OCR 模型与字库来自本机（用户机器首次启动时也可由工具自行准备）

体积构成（约）：
  exe 27MB + runtime ~350MB + OCR 模型 45MB + 字库 1MB
  + local-mineru 脚本 0.4MB + MinerU 模型 1.7GB ≈ 2.1GB，zip 后约 1.9GB
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
STAGE_NAME = "肥鸭文档转换"
# OCR 模型：优先项目内 PP-OCRv6（识别更准），回退本机 venv 的 PP-OCRv5
MODELS_V6 = ROOT / "models" / "PP-OCRv6"
MODELS_V5 = (
    Path(os.environ["USERPROFILE"]) / ".openvino" / "models"
    / "mineru-ov-models" / "PP-OCRv5"
)
# 结构/表格模式已改用本地 PP-OCRv6 + 网格重建，不再随包携带 MinerU 模型

# zip 内需要额外带的运行期文件：(源路径, zip 内相对路径)
# ocr_worker.py 必须随包：OCR 子进程由 runtime python 以脚本方式直接拉起，
# 而 onefile 的 _MEIPASS 是 exe 私有的解包目录，子进程访问不到，
# 因此便携包里要放一份原文（exe 内那份由 add-data 打包，两者一致）。
# 使用说明与解除锁定脚本面向终端用户（未购买代码签名证书的 SmartScreen 应对）。
EXTRA_FILES = (
    (ROOT / "ocr_worker.py", "ocr_worker.py"),
    (ROOT / "VERSION", "VERSION"),
    (ROOT / "assets" / "ppocrv6_dict.txt", "assets/ppocrv6_dict.txt"),
    (ROOT / "assets" / "ppocrv5_dict.txt", "assets/ppocrv5_dict.txt"),
    (ROOT / "LICENSE", "LICENSE"),
    (ROOT / "privacy_notice.txt", "privacy_notice.txt"),
    (ROOT / "使用说明_SmartScreen.txt", "使用说明_SmartScreen.txt"),
    (ROOT / "tools" / "unblock_dist.bat", "解除锁定.bat"),
    (ROOT / "tools" / "unblock_files.ps1", "unblock_files.ps1"),
)


def read_version() -> str:
    """版本号唯一来源：项目根 VERSION 文件。"""
    try:
        return (ROOT / "VERSION").read_text(encoding="utf-8").strip() or "1.1.2"
    except OSError:
        return "1.1.2"


def human(mb: float) -> str:
    return f"{mb:.1f} MB"


def copy_ir(src: Path, dst: Path) -> None:
    """只拷 OpenVINO IR（*.xml/*.bin），跳过 onnx 源文件与配置（便携包用不到）。"""
    for part in ("det", "rec"):
        out = dst / part
        out.mkdir(parents=True, exist_ok=True)
        for f in sorted((src / part).glob("*")):
            if f.suffix in (".xml", ".bin"):
                shutil.copy2(f, out / f.name)


def dir_size_mb(path: Path) -> float:
    total = 0
    for item in path.rglob("*"):
        try:
            if item.is_file():
                total += item.stat().st_size
        except OSError:
            continue
    return total / 1024 / 1024


def ignore_junk(_dir: str, names: list[str]) -> set[str]:
    """打包时跳过 __pycache__ 等无用文件。"""
    return {n for n in names if n == "__pycache__" or n.endswith((".pyc", ".pyo"))}


def stage_package(stage: Path) -> float:
    """把 exe / runtime / 模型 / 字库 组装到暂存目录，返回体积。"""
    target = stage / STAGE_NAME
    target.mkdir(parents=True, exist_ok=True)

    exe = DIST / "肥鸭文档转换.exe"
    if not exe.is_file():
        raise FileNotFoundError(f"缺少 exe，请先运行 build_exe.py：{exe}")
    shutil.copy2(exe, target / exe.name)
    print(f"  exe      {human(exe.stat().st_size / 1024 / 1024)}")

    runtime = ROOT / "runtime"
    # 新版 runtime 是自包含的普通 Python 安装目录（解释器在 runtime\python.exe）；
    # 兼容旧版 venv 布局（runtime\Scripts\python.exe）。
    has_python = (runtime / "python.exe").is_file() or (runtime / "Scripts" / "python.exe").is_file()
    if not has_python:
        raise FileNotFoundError("缺少 runtime，请先运行 build_runtime.py")
    print(f"  runtime  {human(dir_size_mb(runtime))}")
    shutil.copytree(runtime, target / "runtime", symlinks=True)

    if not MODELS_V6.is_dir():
        raise FileNotFoundError(f"缺少 OCR 模型目录：{MODELS_V6}")
    print(f"  模型     {human(dir_size_mb(MODELS_V6))}")
    copy_ir(MODELS_V6, target / "models" / "PP-OCRv6")
    # 不再打包 PP-OCRv5 回退模型：MinerU 缓存目录不随包，主模型 PP-OCRv6 已够用

    for src, rel in EXTRA_FILES:
        dst = target / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        print(f"  附带文件 已放入 {rel}")

    return dir_size_mb(target)


def make_zip(stage: Path, zip_path: Path) -> float:
    """把暂存目录压成 zip，返回 zip 体积。"""
    base = stage / STAGE_NAME
    if zip_path.exists():
        zip_path.unlink()
    t0 = time.time()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for item in base.rglob("*"):
            if item.is_file():
                zf.write(item, item.relative_to(stage).as_posix())
    print(f"压缩完成，用时 {time.time() - t0:.0f}s")
    return zip_path.stat().st_size / 1024 / 1024


def main() -> int:
    DIST.mkdir(exist_ok=True)
    version = read_version()
    zip_path = DIST / f"{STAGE_NAME}_便携包_v{version}.zip"
    print(f"版本：{version}")
    stage = Path(tempfile.mkdtemp(prefix="img2xlsx_pkg_"))
    try:
        print("组装便携包…")
        size = stage_package(stage)
        print(f"组装后共 {human(size)}")
        print("压缩为 zip…")
        zip_size = make_zip(stage, zip_path)
        print(f"产物：{zip_path}（{human(zip_size)}）")
        print("\n使用方式：解压后双击 肥鸭文档转换.exe，无需安装 Python。")
        return 0
    finally:
        shutil.rmtree(stage, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
