# -*- coding: utf-8 -*-
"""打包 肥鸭文档转换.exe（PyInstaller onefile + windowed）。

用法：python build_exe.py
产物：dist\\肥鸭文档转换.exe（体积上限 MAX_EXE_MB，超出即判失败）

程序形态（2026-10-31 起）：
  入口是 desktop.py —— 启动本地 FastAPI 服务，再用 pywebview + WebView2 打开
  原生窗口（无浏览器外壳、无地址栏），关窗即退出；WebView2 缺失时回退到
  Edge/Chrome 的 app 模式窗口。EXE 内打包 web/ 静态资源与 assets/，运行时
  解包到 _MEIPASS 由 FastAPI 直接提供。

注意：ocr_worker.py 与 assets\\ppocrv6_dict.txt 是运行时数据文件——
exe 内的 OCR 工作进程由本地 runtime 的 python 直接以脚本方式拉起，必须拿到
.py 原文与字典文件，因此用 --add-data 打包，不能用 hidden-imports 代替。

体积控制：
  - 代码未引用 Pillow，通过 --exclude-module 排除（openpyxl 的图片支持可选）。
  - 构建后强制校验 exe 体积，超过上限直接非零退出，防止体积悄悄变大。
"""
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
EXE_NAME = "肥鸭文档转换"
EXE_PATH = ROOT / "dist" / f"{EXE_NAME}.exe"
MAX_EXE_MB = 35


def read_version() -> str:
    """版本号唯一来源：项目根 VERSION 文件。"""
    p = ROOT / "VERSION"
    try:
        text = p.read_text(encoding="utf-8").strip()
        return text or "1.1.0"
    except OSError:
        return "1.1.0"


def write_version_files(version: str) -> None:
    """按 VERSION 生成 exe 版本资源与 Inno 版本定义，避免多处手写不一致。"""
    nums = [int(x) for x in (version.split(".") + ["0", "0", "0", "0"])[:4]]
    v4 = ".".join(str(n) for n in nums)
    (ROOT / "version.iss").write_text(
        f'#define MyAppVersion "{version}"\n#define MyAppVersion4 "{v4}"\n',
        encoding="utf-8",
    )
    (ROOT / "version_info.txt").write_text(
        f'''# -*- coding: utf-8 -*-
# 由 build_exe.py 依据 VERSION 生成，请勿手工修改。
VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=({nums[0]}, {nums[1]}, {nums[2]}, {nums[3]}),
        prodvers=({nums[0]}, {nums[1]}, {nums[2]}, {nums[3]}),
        mask=0x3F,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0),
    ),
    kids=[
        StringFileInfo(
            [
                StringTable(
                    "040904B0",
                    [
                        StringStruct("CompanyName", "Fat Duck OCR Contributors"),
                        StringStruct("FileDescription", "肥鸭文档转换 —— 本地 OCR 图片/PDF 转 Excel 工具"),
                        StringStruct("FileVersion", "{v4}"),
                        StringStruct("InternalName", "肥鸭文档转换"),
                        StringStruct("LegalCopyright", "Copyright (c) 2026 Fat Duck OCR Contributors"),
                        StringStruct("OriginalFilename", "肥鸭文档转换.exe"),
                        StringStruct("ProductName", "肥鸭文档转换"),
                        StringStruct("ProductVersion", "{v4}"),
                    ],
                )
            ]
        ),
        VarFileInfo([VarStruct("Translation", [0x0409, 0x04B0])]),
    ],
)
''',
        encoding="utf-8",
    )
    print(f"版本：{version}（版本资源 {v4}）", flush=True)

# 主程序未使用的模块，排除以控制体积。
# 注意：PIL 不能排除——托盘图标用 Pillow 加载；tkinter 未被任何入口导入，
# PyInstaller 会自动剔除，无需显式 exclude。
# numpy 仅由 ocr_worker.py 使用，而该文件是运行时数据、由随包 runtime 的 python
# 执行，不属于 exe 进程；PyInstaller 通过依赖 hook 误收集会多占约 15 MB。
EXCLUDES = ("pytest", "setuptools", "pip", "numpy")


def kill_running_exe() -> None:
    """打包前终止占用 dist exe 的残留进程，避免 WinError 5。"""
    try:
        subprocess.run(
            ["taskkill", "/IM", f"{EXE_NAME}.exe", "/F"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
        )
    except OSError:
        return
    time.sleep(0.6)


def build() -> int:
    write_version_files(read_version())
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name", EXE_NAME,
        "--icon", str(ROOT / "assets" / "app.ico"),
        "--add-data", f"{ROOT / 'ocr_worker.py'};.",
        "--add-data", f"{ROOT / 'VERSION'};.",
        "--add-data", f"{ROOT / 'web'};web",
        "--add-data", f"{ROOT / 'assets'};assets",
        "--collect-submodules", "uvicorn",
        "--collect-all", "webview",
        "--collect-all", "clr_loader",
        "--collect-all", "pystray",
        "--hidden-import", "clr",
        "--hidden-import", "pystray._win32",
        "--version-file", str(ROOT / "version_info.txt"),
    ]
    for name in EXCLUDES:
        cmd += ["--exclude-module", name]
    cmd.append(str(ROOT / "desktop.py"))

    print(" ".join(cmd), flush=True)
    return subprocess.call(cmd, cwd=str(ROOT))


def check_size() -> int:
    """体积门禁：超限返回 1，正常返回 0。"""
    if not EXE_PATH.is_file():
        print(f"未找到产物：{EXE_PATH}", flush=True)
        return 1
    size_mb = EXE_PATH.stat().st_size / 1024 / 1024
    print(f"产物体积：{size_mb:.1f} MB（上限 {MAX_EXE_MB} MB）", flush=True)
    if size_mb > MAX_EXE_MB:
        print("体积超出上限，构建判定失败。", flush=True)
        return 1
    return 0


def main() -> int:
    kill_running_exe()
    if build() != 0:
        return 1
    return check_size()


if __name__ == "__main__":
    raise SystemExit(main())
