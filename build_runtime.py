# -*- coding: utf-8 -*-
"""构建 OCR 运行时（runtime）：把 openvino venv 裁剪成便携运行环境。

用法：python build_runtime.py
产物：runtime\\python.exe + runtime\\Lib\\site-packages

背景：
  三种模式都由本地 venv 的 python 以子进程方式拉起 ocr_worker.py（PP-OCRv6）。
  完整 venv 约 512 MB，其中 ocr_worker 只用到 cv2 / numpy / openvino / pypdfium2。
  MinerU 相关依赖（openvino_genai / tokenizers / modelscope / pydantic / PIL）
  已随功能移除一并裁掉。

策略（不破坏现有 venv）：
  1. 复制 venv 到 runtime/（原 venv 保持不动）
  2. 删除运行期不需要的包（见 REMOVE_DIRS）
  3. 删除 CPU 推理不需要的 openvino 组件与 cv2 视频/级联数据
  4. 清掉 __pycache__ / tests / *.pyi 等非运行文件
  5. 校验体积上限，并用 runtime python 试跑一次 OCR 冒烟

体积红线（构建后自动校验）：
  - runtime 目录 <= 280 MB
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().home() / ".openvino" / "venv" / "mineru"
if not ROOT.is_dir():
    ROOT = Path(os.environ["USERPROFILE"]) / ".openvino" / "venv" / "mineru"
DEST = Path(__file__).resolve().parent / "runtime"
MAX_RUNTIME_MB = 280

# 运行期不需要的包（已逐包核对 import 关系）
# 注意：pypdfium2 / pypdfium2_raw 是 PDF 渲染引擎，必须保留；
#       cv2 / numpy / openvino 是 OCR 必需，不能删。
REMOVE_DIRS = (
    "cryptography",     # 安装/加密用，运行期不用
    "pip",              # 包管理器
    # MinerU 相关依赖：结构/表格模式已改用本地 OCR + 网格重建，不再需要
    "modelscope",
    "modelscope_hub",
    "openvino_genai",
    "openvino_tokenizers",
    "mineru_vl_utils",
    "pydantic",
    "pydantic_core",
    "loguru",
    "PIL",
    "pillow",
)

# openvino/libs 下运行期不需要的二进制（OCR 固定 CPU）
# 注意：openvino_ir_frontend.dll 是读取 .xml IR 模型的前端，必须保留，
#       删掉会报 "Available frontends: jax pytorch"。
#       cache.json 是前端注册表缓存，删掉后 openvino 会自动重建（省 8.4MB）。
REMOVE_LIB_FILES = (
    "openvino_intel_npu_compiler.dll",
    "openvino_intel_npu_plugin.dll",
    "openvino_onnx_frontend.dll",
    "openvino_tensorflow_frontend.dll",
    "openvino_paddle_frontend.dll",
    "openvino_tensorflow_lite_frontend.dll",
    "openvino_intel_gpu_plugin.dll",
    "openvino_pytorch_frontend.dll",
    "openvino_gguf_frontend.dll",
    "openvino_jax_frontend.dll",
    "openvino_intel_npu_vm_runtime.dll",
    "openvino_intel_npu_compiler_loader.dll",
    "cache.json",
)

# 清理规则：目录名或后缀命中即删
CLEAN_DIR_NAMES = ("__pycache__", "tests", "test")
CLEAN_SUFFIXES = (".pyi", ".pyc", ".h", ".cpp", ".cu", ".cuh", ".lib")


def human(mb: float) -> str:
    return f"{mb:.1f} MB"


def dir_size_mb(path: Path) -> float:
    total = 0
    for item in path.rglob("*"):
        try:
            if item.is_file():
                total += item.stat().st_size
        except OSError:
            continue
    return total / 1024 / 1024


def remove_packages(site: Path) -> tuple[int, float]:
    """删除冗余包，返回 (删除条数, 释放 MB)。"""
    freed = 0.0
    removed = 0
    for name in REMOVE_DIRS:
        for target in (site / name, site / f"{name}.dist-info", site / f"{name}-*.dist-info"):
            if any(ch in target.name for ch in "*?["):
                for hit in site.glob(target.name):
                    freed += dir_size_mb(hit)
                    shutil.rmtree(hit, ignore_errors=True)
                    removed += 1
                continue
            if target.is_dir():
                freed += dir_size_mb(target)
                shutil.rmtree(target, ignore_errors=True)
                removed += 1
            elif target.is_file():
                freed += target.stat().st_size / 1024 / 1024
                target.unlink(missing_ok=True)
                removed += 1
    return removed, freed


def remove_openvino_extras(site: Path) -> tuple[int, float]:
    """删除 openvino/libs 下 CPU 推理不需要的二进制，返回 (删除条数, 释放 MB)。"""
    libs = site / "openvino" / "libs"
    if not libs.is_dir():
        return 0, 0.0
    freed = 0.0
    removed = 0
    for name in REMOVE_LIB_FILES:
        target = libs / name
        if target.is_file():
            freed += target.stat().st_size / 1024 / 1024
            target.unlink(missing_ok=True)
            removed += 1
    return removed, freed


def clean_junk(site: Path) -> tuple[int, float]:
    """清 __pycache__ / 测试 / 源码与中间文件，返回 (删除条数, 释放 MB)。"""
    freed = 0.0
    removed = 0
    for item in list(site.rglob("*")):
        try:
            if not item.exists():
                continue
            if item.is_dir():
                if item.name in CLEAN_DIR_NAMES:
                    freed += dir_size_mb(item)
                    shutil.rmtree(item, ignore_errors=True)
                    removed += 1
                continue
            if item.suffix in CLEAN_SUFFIXES:
                freed += item.stat().st_size / 1024 / 1024
                item.unlink(missing_ok=True)
                removed += 1
        except OSError:
            continue
    return removed, freed


def trim_extras(dest: Path) -> tuple[int, float]:
    """删除 OCR 子进程用不到的文件：cv2 视频组件与级联数据、tk 残留、多余标准库。"""
    removed = 0
    freed = 0.0

    def drop(path: Path) -> None:
        nonlocal removed, freed
        if path.is_dir():
            freed += dir_size_mb(path)
            shutil.rmtree(path, ignore_errors=True)
            removed += 1
        elif path.is_file():
            freed += path.stat().st_size / 1024 / 1024
            path.unlink(missing_ok=True)
            removed += 1

    cv2_dir = dest / "Lib" / "site-packages" / "cv2"
    drop(cv2_dir / "opencv_videoio_ffmpeg4110_64.dll")
    drop(cv2_dir / "data")
    lib = dest / "Lib"
    for name in ("ensurepip", "idlelib", "pydoc_data", "tkinter", "distutils",
                 "lib2to3", "turtledemo", "turtle.py", "venv"):
        drop(lib / name)
    dlls = dest / "DLLs"
    for name in ("tcl86t.dll", "tk86t.dll"):
        drop(dlls / name)
    return removed, freed


def make_relocatable() -> None:
    """改写 pyvenv.cfg 的 home 为相对路径，使 runtime 可在任意目录运行。

    原 venv 的 pyvenv.cfg 指向 uv 缓存里的基础解释器，换机后该路径不存在，
    venv 会直接启动失败。实测 home = . 即可让解释器就地启动（Python 3.11）。
    """
    cfg = DEST / "pyvenv.cfg"
    if not cfg.is_file():
        return
    lines = []
    for raw in cfg.read_text(encoding="utf-8", errors="replace").splitlines():
        if raw.strip().lower().startswith("home"):
            lines.append("home = .")
        else:
            lines.append(raw)
    cfg.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("pyvenv.cfg 已改为相对 home，runtime 可脱离原路径运行")


# VC++ 2015-2022 运行库全套（约 1.4 MB）。依赖扫描结论：runtime 内 16 个
# .pyd/.dll（含 openvino 核心 _pyopenvino.pyd）导入 msvcp140.dll，而基础
# 解释器只自带 vcruntime140*。打包机装了系统级运行库所以一切正常；换到
# 未装 VC++ 运行库的电脑就 DLL load failed，OCR 引擎起不来，界面还会
# 误报「模型缺失」（v1.5.2 他机部署事故的根因）。
VC_REDIST_DLLS = (
    "vcruntime140.dll", "vcruntime140_1.dll", "vcruntime140_threads.dll",
    "msvcp140.dll", "msvcp140_1.dll", "msvcp140_2.dll",
    "msvcp140_atomic_wait.dll", "msvcp140_codecvt_ids.dll", "concrt140.dll",
)


def bundle_vc_redist(dest: Path) -> None:
    """把 VC++ 运行库 DLL 补进 runtime 根目录（与 python.exe 同级）。

    DLL 放在解释器同级目录即可被 Windows 加载器优先找到，目标机无需
    另装 VC++ 运行库。从打包机 System32 取（即微软官方 redist 内容）。
    """
    sys32 = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32"
    copied, missing = [], []
    for name in VC_REDIST_DLLS:
        src = sys32 / name
        if src.is_file():
            shutil.copy2(src, dest / name)
            copied.append(name)
        elif not (dest / name).is_file():
            missing.append(name)
    print(f"VC++ 运行库：已补齐 {len(copied)} 个 DLL 到 runtime 根目录")
    if missing:
        print(f"警告：打包机系统缺少 {missing}；目标机若也未装 VC++ 运行库，OCR 将无法启动")


def embed_base_interpreter(dest: Path | None = None) -> None:
    """把基础 CPython 解释器摊平嵌入 runtime，使其完全自包含、可整体搬迁。

    venv 里的 Scripts\\python.exe 只是重定向器，真正的解释器与 python311.dll
    位于本机绝对路径（如 uv 缓存目录 C:\\Users\\...\\uv\\python\\cpython-3.11-...）。
    换一台电脑该路径不存在，子进程会秒退且没有任何输出，表现为
    「OCR 工作进程无响应（可能已崩溃）」。

    这里把基础解释器直接摊平到 runtime/ 根目录，并删除 pyvenv.cfg，
    让 runtime 变成普通可重定位的 Python 安装目录（解释器 + 标准库 + 依赖
    全在包内，不再依赖打包机的任何路径）。
    """
    dest = dest or DEST
    launcher = dest / "Scripts" / "python.exe"
    if not launcher.is_file():
        raise FileNotFoundError(f"未找到 venv 启动器：{launcher}")

    proc = subprocess.run(
        [str(launcher), "-c", "import sys; print(sys.base_prefix)"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120,
    )
    base = Path(proc.stdout.strip())
    if not (base / "python.exe").is_file():
        raise RuntimeError(f"无法定位基础解释器（base_prefix={base}）")
    print(f"基础解释器：{base}")

    # 1) 解释器与运行库文件
    for name in (
        "python.exe", "pythonw.exe", "python3.dll", "python311.dll",
        "vcruntime140.dll", "vcruntime140_1.dll", "python311.zip",
    ):
        src = base / name
        if src.is_file():
            shutil.copy2(src, dest / name)

    # 1b) 补齐 VC++ 运行库（msvcp140.dll 等），否则未装系统运行库的电脑起不来
    bundle_vc_redist(dest)

    # 2) 扩展模块（DLLs）
    if (base / "DLLs").is_dir():
        shutil.copytree(base / "DLLs", dest / "DLLs", dirs_exist_ok=True)

    # 3) 标准库（Lib 下除 site-packages 外全部拷入，site-packages 保持依赖原样）
    lib = dest / "Lib"
    lib.mkdir(parents=True, exist_ok=True)
    for item in (base / "Lib").iterdir():
        if item.name == "site-packages":
            continue
        target = lib / item.name
        if item.is_dir():
            shutil.copytree(item, target, dirs_exist_ok=True)
        else:
            shutil.copy2(item, target)

    # 4) 去掉 venv 痕迹：存在 pyvenv.cfg 时，Python 会把该目录当成 venv，
    #    从而按 home 去找外部解释器——那正是换机失效的根因。
    (dest / "pyvenv.cfg").unlink(missing_ok=True)
    shutil.rmtree(dest / "Scripts", ignore_errors=True)

    print(f"已嵌入基础解释器，runtime 体积 {human(dir_size_mb(dest))}")


def smoke_test() -> bool:
    """用 runtime python 试跑 ocr_worker 的 ping，确认环境可用。"""
    python = DEST / "python.exe"
    if not python.is_file():
        python = DEST / "Scripts" / "python.exe"
    worker = Path(__file__).resolve().parent / "ocr_worker.py"
    models = Path(__file__).resolve().parent / "models" / "PP-OCRv6"
    script = (
        "import sys, json;"
        f"sys.argv=['x','--models',r'{models}'];"
        "import importlib.util;"
        f"spec=importlib.util.spec_from_file_location('w', r'{worker}');"
        "m=importlib.util.module_from_spec(spec);"
        "spec.loader.exec_module(m);"
        "e=m.OcrEngine(m.default_models_dir(), m.default_dict_path(m.default_models_dir()), 'CPU');"
        "print('OCR_ENGINE_OK')"
    )
    proc = subprocess.run(
        [str(python), "-c", script],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300,
    )
    print(proc.stdout.strip()[-500:])
    if proc.returncode != 0:
        print(proc.stderr.strip()[-1500:])
    return proc.returncode == 0 and "OCR_ENGINE_OK" in proc.stdout


def main() -> int:
    src_venv = ROOT
    if not (src_venv / "Scripts" / "python.exe").is_file():
        print(f"未找到源 venv：{src_venv}")
        return 1

    if DEST.exists():
        shutil.rmtree(DEST, ignore_errors=True)
    print(f"复制 venv -> {DEST}（源 venv {human(dir_size_mb(src_venv))}）")
    t0 = time.time()
    shutil.copytree(src_venv, DEST, symlinks=True)
    print(f"复制完成，用时 {time.time() - t0:.0f}s，当前 {human(dir_size_mb(DEST))}")

    site = next((DEST / "Lib" / "site-packages").glob("*"), None)
    site = DEST / "Lib" / "site-packages"
    if not site.is_dir():
        print("未找到 site-packages")
        return 1

    n, freed = remove_packages(site)
    print(f"删除冗余包 {n} 项，释放 {human(freed)}，当前 {human(dir_size_mb(DEST))}")

    n, freed = remove_openvino_extras(site)
    print(f"删除 openvino 非 CPU 组件 {n} 项，释放 {human(freed)}，当前 {human(dir_size_mb(DEST))}")

    n, freed = clean_junk(site)
    print(f"清理缓存/测试/中间文件 {n} 项，释放 {human(freed)}，当前 {human(dir_size_mb(DEST))}")

    # Scripts 下同样清缓存
    n, freed = clean_junk(DEST / "Scripts")
    if n:
        print(f"清理 Scripts {n} 项，释放 {human(freed)}，当前 {human(dir_size_mb(DEST))}")

    n, freed = trim_extras(DEST)
    print(f"删除 cv2 视频/级联数据与 tk 残留 {n} 项，释放 {human(freed)}，"
          f"当前 {human(dir_size_mb(DEST))}")

    size = dir_size_mb(DEST)
    print(f"最终 runtime 体积：{human(size)}（上限 {MAX_RUNTIME_MB} MB）")
    if size > MAX_RUNTIME_MB:
        print("runtime 超重，构建判定失败。")
        return 1

    make_relocatable()
    print("嵌入基础解释器（保证换机可用）…")
    embed_base_interpreter(DEST)
    print("冒烟测试：用 runtime python 加载 PP-OCRv6 引擎…")
    if not smoke_test():
        print("runtime 冒烟测试失败，构建判定失败。")
        return 1
    print("runtime 冒烟测试通过。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
