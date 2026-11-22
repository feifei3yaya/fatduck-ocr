# -*- coding: utf-8 -*-
"""本地解析引擎的进程管理与调用。

两条独立管线：

1) MinerU 管线（结构/表格两种 Markdown 模式）：
   技能自带的 run.ps1 依赖共享的 server-dog 守护进程，守护进程会周期性检查
   宿主程序是否存活；在本机环境中该检查会误判为已退出，从而在解析途中回收
   解析服务。桌面工具需要可控的引擎生命周期，因此这里直接拉起 server.py，
   并通过其命名管道（\\\\.\\pipe\\local-mineru）收发请求。

2) 精确复制管线（纯本地 OCR，默认全程不联网）：
   ocr_worker.py（PP-OCRv6，JSON-lines 协议）-> 规则级校准 -> 几何重建
   -> 序号断档补行 -> xcheck 交叉校验（整图 OCR vs 局部放大 OCR 两方对比：
   编码类字段不一致时以放大复识结果自动改写，文本字段只标记待复核）
   -> 可选的云端文本纠错（默认关闭；仅当用户在设置里开启后，
      才把"待复核"的文本类单元格裁剪图外发，见 nv_fix.py），
   最后 convert_exact 直写 Excel。
"""
from __future__ import annotations

import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
from collections import Counter
from multiprocessing.connection import Client
from pathlib import Path

import nv_fix

from ai_calibrate import (_incomplete_date_digits, _one_digit_lost,
                          calibrate_boxes, clean_grid_rows, parse_date)
from md2excel import ConvertResult, ExactPage, convert_exact
from table_rebuild import TableBlock, TextBlock, fill_serial_gaps, rebuild
from xcheck import cross_check

PIPE_ADDRESS = r"\\.\pipe\local-mineru"
AUTHKEY = b"local-mineru"
CREATE_NO_WINDOW = 0x08000000
PARSE_TIMEOUT_S = 900
READY_STATE = "running"


class EngineError(RuntimeError):
    """引擎启动或调用失败。"""


def _candidate_bases() -> list[Path]:
    """查找随包资源的根目录：exe 同级 → _MEIPASS → 源码目录。"""
    bases = []
    try:
        bases.append(Path(sys.executable).resolve().parent)
    except Exception:
        pass
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bases.append(Path(meipass))
    bases.append(Path(__file__).resolve().parent)
    return bases


def bundled_runtime_python() -> Path | None:
    """随包 runtime\\python.exe（便携包/安装包放在 exe 同级）。"""
    for base in _candidate_bases():
        python = base / "runtime" / "python.exe"
        if python.is_file():
            return python
    return None


class MineruEngine:
    def __init__(self, skill_dir: Path) -> None:
        self.skill_dir = Path(skill_dir)
        self.proc: subprocess.Popen | None = None
        self._spawn_logged = 0
        self._info: dict | None = None

    @property
    def info(self) -> dict:
        if self._info is None:
            self._info = json.loads((self.skill_dir / "info.json").read_text(encoding="utf-8"))
        return self._info

    def partial_model_dir(self) -> Path | None:
        models = self.info.get("models") or []
        if not models:
            return None
        root = Path(os.environ.get("USERPROFILE", str(Path.home()))) / ".openvino" / "models"
        return root / (str(models[0].get("dir_name", "")) + ".partial")

    def _downloaded_mb(self) -> int:
        partial = self.partial_model_dir()
        if partial is None or not partial.is_dir():
            return 0
        total = 0
        for item in partial.rglob("*"):
            if item.is_file():
                try:
                    total += item.stat().st_size
                except OSError:
                    continue
        return int(total / 1024 / 1024)

    # ------------------------------------------------------------ 运行环境

    def venv_python(self) -> Path:
        # 优先用随包 runtime（已含 MinerU 依赖），干净机器无需再装环境
        bundled = bundled_runtime_python()
        if bundled is not None:
            return bundled
        venv = (
            Path(os.environ.get("USERPROFILE", str(Path.home())))
            / ".openvino" / "venv" / str(self.info.get("venv_name", "mineru"))
        )
        return venv / "Scripts" / "python.exe"

    def bundled_models_root(self) -> Path | None:
        """随包携带的 MinerU 模型根目录（models\\mineru-ov-models 的父目录）。"""
        try:
            model = (self.info.get("models") or [{}])[0]
            dir_name = str(model.get("dir_name") or "")
            required = list(model.get("required_files") or [])
        except Exception:
            return None
        if not dir_name:
            return None
        for base in _candidate_bases():
            root = base / "models"
            model_dir = root / dir_name
            if not model_dir.is_dir():
                continue
            if all((model_dir / rel).is_file() for rel in required) or not required:
                return root
        return None

    def ensure_env(self, log) -> Path:
        """确保离线运行环境已安装，返回其 python 解释器路径。"""
        python = self.venv_python()
        if python.is_file():
            return python
        installer = self.skill_dir / "scripts" / "install-env.ps1"
        if not installer.is_file():
            raise EngineError(f"解析引擎缺少安装脚本：{installer}")
        log("首次使用：正在准备离线解析环境，请稍候（约 1-2 分钟）…")
        code = subprocess.call(
            [
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                "-File", str(installer), "-SkillRoot", str(self.skill_dir),
            ],
            creationflags=CREATE_NO_WINDOW,
        )
        if code != 0:
            raise EngineError("离线解析环境安装失败，请检查网络后重试。")
        if not python.is_file():
            raise EngineError("离线解析环境安装不完整，未找到 python 解释器。")
        log("离线解析环境准备完成。")
        return python

    # ------------------------------------------------------------ 管道通信

    def _connect(self):
        try:
            return Client(PIPE_ADDRESS, authkey=AUTHKEY)
        except (ConnectionRefusedError, FileNotFoundError, OSError):
            return None

    def status(self) -> dict | None:
        conn = self._connect()
        if conn is None:
            return None
        try:
            conn.send({"op": "status"})
            if not conn.poll(5):
                return None
            reply = conn.recv()
            return reply if isinstance(reply, dict) else None
        except (EOFError, OSError):
            return None
        finally:
            conn.close()

    def _spawn(self, python: Path, log) -> None:
        if self.proc is not None and self.proc.poll() is None:
            return
        env = os.environ.copy()
        env["LOCAL_MINERU_OV_PATH"] = str(self.skill_dir / "scripts" / "mineru_ov")
        env["PYTHONIOENCODING"] = "utf-8"
        models_root = self.bundled_models_root()
        if models_root is not None:
            env["LOCAL_MINERU_MODELS_ROOT"] = str(models_root)
        if self._spawn_logged < 2:
            log("正在启动本地识别引擎…")
        self._spawn_logged += 1
        self.proc = subprocess.Popen(
            [str(python), str(self.skill_dir / "scripts" / "server.py")],
            cwd=str(self.skill_dir),
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW,
        )

    def ensure_running(self, log, on_phase=None, should_stop=None) -> None:
        """等待引擎进入可用状态。首次运行会下载约 2-4GB 模型。"""
        python = self.ensure_env(log)
        deadline = time.time() + 3600
        announced_download = False
        while time.time() < deadline:
            if should_stop is not None and should_stop():
                raise EngineError("已取消转换。")
            state = self.status()
            if state is None:
                self._spawn(python, log)
                time.sleep(2)
                continue

            name = str(state.get("state", ""))
            if name == READY_STATE:
                if on_phase is not None:
                    on_phase("识别引擎就绪")
                return
            if name == "error":
                detail = (state.get("error") or "").strip().splitlines()[-1:]
                raise EngineError("识别引擎初始化失败：" + (" ".join(detail) or "未知原因"))

            progress = state.get("download_progress")
            if name == "downloading":
                downloaded = self._downloaded_mb()
                text = "正在下载识别模型（首次使用，约 2-4GB）"
                text += f"：已下载 {downloaded} MB" if downloaded else f"：{progress or '准备中'}"
                if not announced_download:
                    log("首次使用需要下载识别模型，请保持网络畅通，此过程可能需要 10 分钟以上。")
                    announced_download = True
            else:
                text = f"正在加载识别模型：{progress or '准备中'}"
            if on_phase is not None:
                on_phase(text)
            time.sleep(2)

        raise EngineError("识别引擎启动超时，请稍后重试。")

    def parse(self, input_path: Path, output_dir: Path) -> dict:
        conn = self._connect()
        if conn is None:
            raise EngineError("识别引擎未响应，请重试。")
        try:
            conn.send(
                {
                    "op": "parse",
                    "input_path": str(input_path),
                    "output_dir": str(output_dir),
                }
            )
            if not conn.poll(PARSE_TIMEOUT_S):
                raise EngineError("解析超时，请尝试分辨率更低的图片。")
            reply = conn.recv()
            return reply if isinstance(reply, dict) else {"ok": False, "error": "无效的引擎响应"}
        except (EOFError, OSError) as exc:
            raise EngineError(f"与识别引擎通信中断：{exc}") from exc
        finally:
            conn.close()

    def shutdown(self) -> None:
        """主动关闭引擎（默认不调用，引擎空闲后会自动退出）。"""
        conn = self._connect()
        if conn is not None:
            try:
                conn.send({"op": "shutdown", "timeout": 10})
                conn.poll(12) and conn.recv()
            except (EOFError, OSError):
                pass
            finally:
                conn.close()
        if self.proc is not None and self.proc.poll() is None:
            try:
                self.proc.wait(15)
            except subprocess.TimeoutExpired:
                self.proc.terminate()


# ---------------------------------------------------------------- OCR 工作进程（精确复制管线）

# Windows 常见崩溃退出码 -> 中文原因（用于把「无响应」翻译成可执行的处理建议）
_CRASH_REASONS = {
    -1073741515: "缺少运行库（Microsoft Visual C++ 2015-2022 可再发行组件未安装）",
    -1073741795: "当前 CPU 不支持 AVX2 指令集（OpenVINO 加速的硬性要求）",
    -1073740791: "运行库堆损坏（通常是 VC++ 运行库版本冲突）",
    -1073741819: "非法访问（指令集或内存不兼容）",
    -1073741811: "找不到依赖的 DLL（安装目录文件不完整或被安全软件拦截）",
}


def _describe_exit(code: int | None) -> str:
    """把子进程退出码翻译成中文原因。"""
    if code is None:
        return "进程仍在运行但无输出"
    reason = _CRASH_REASONS.get(code)
    text = f"退出码 {code}（0x{code & 0xFFFFFFFF:08X}）"
    return f"{text}：{reason}" if reason else text


def _probe_runtime(python: Path) -> str:
    """用同一个解释器做一次依赖自检，给出真实的失败原因。"""
    try:
        proc = subprocess.run(
            [str(python), "-c", "import openvino, cv2, numpy; print('RUNTIME_OK')"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
        )
    except Exception as exc:
        return f"运行环境自检无法执行：{exc}"
    if proc.returncode == 0 and "RUNTIME_OK" in (proc.stdout or ""):
        return "运行环境自检通过（openvino / cv2 / numpy 均可导入）"
    tail = (proc.stderr or proc.stdout or "").strip()[-800:]
    return f"运行环境自检失败：{tail or '（无输出）'}"


# 子进程响应超时（秒）：模型卡死/大图异常时必须有界，否则任务与取消都会永久挂起。
# 数值按"正常最慢 + 充分余量"取：整页 OCR 300s、批量查找 420s、PDF 渲染 900s。
OCR_STARTUP_TIMEOUT = 300.0
OCR_DEFAULT_TIMEOUT = 300.0
OCR_READ_TIMEOUTS = {
    "ocr": 300.0,
    "fast_find": 420.0,
    "encode": 180.0,
    "render_pdf": 900.0,
    "ping": 10.0,
}


class OcrWorker:
    """ocr_worker.py 子进程客户端（stdin/stdout，一行一个 JSON）。

    与 MineruEngine 相互独立：worker 常驻内存加载 PP-OCRv6 det/rec 模型，
    支持 clip + scale 局部放大重识别（供序号断档补行使用）。
    """

    @staticmethod
    def _locate_script() -> Path:
        """定位 ocr_worker.py：便携包根目录优先，其次 _MEIPASS，最后源码目录。"""
        candidates = []
        try:
            candidates.append(Path(sys.executable).resolve().parent / "ocr_worker.py")
        except Exception:
            pass
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(Path(meipass) / "ocr_worker.py")
        candidates.append(Path(__file__).resolve().parent / "ocr_worker.py")
        for path in candidates:
            if path.is_file():
                return path
        return candidates[-1]

    def __init__(self, python: Path, script: Path | None = None, log=None) -> None:
        # 脚本定位优先级：便携包根目录（exe 同级）> _MEIPASS 内打包副本 > 源码目录。
        # 便携包里 ocr_worker.py 与 models/ 同級，子进程用 runtime python 直接
        # 以脚本方式拉起，必须能拿到原文；_MEIPASS 是 exe 私有解包目录，
        # 对子进程不可见，因此不能只依赖它。
        if script is None:
            script = self._locate_script()
        if not script.is_file():
            raise EngineError(f"缺少 OCR 工作进程脚本：{script}")
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        try:
            self.proc = subprocess.Popen(
                [str(python), str(script)],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
                bufsize=1,
                env=env,
                creationflags=CREATE_NO_WINDOW,
            )
        except (FileNotFoundError, OSError) as exc:
            raise EngineError(f"OCR 引擎启动失败（无法执行 {python}）：{exc}") from exc
        self._id = 0
        self._closed = False
        self._lock = threading.Lock()
        self.engine_name = "PP-OCR"
        # stderr 持续排泄：子进程日志（补行/几何归一化等）在长任务下会超过
        # 管道缓冲（64KB），无人读取会把子进程堵死在 stderr 写上、父进程
        # 随后全部挂起。后台线程读走并保留尾部，崩溃时取最后片段报错。
        self._stderr_tail: list[str] = []
        self.last_quality: dict = {}
        # stdout 由独立线程读入队列，主线程用带超时的 queue.get 取响应：
        # readline 阻塞时既无法超时、也无法被取消，一旦子进程卡死整批任务就永久停住。
        self._lines: queue.Queue = queue.Queue()
        self._aborted = False

        def _drain_stderr() -> None:
            try:
                for line in self.proc.stderr:
                    self._stderr_tail.append(line)
                    if len(self._stderr_tail) > 400:
                        del self._stderr_tail[:200]
            except Exception:
                pass

        def _drain_stdout() -> None:
            try:
                for line in self.proc.stdout:
                    self._lines.put(line)
            except Exception:
                pass
            finally:
                self._lines.put(None)

        threading.Thread(target=_drain_stderr, daemon=True).start()
        threading.Thread(target=_drain_stdout, daemon=True).start()
        try:
            evt = self._read_event(timeout=OCR_STARTUP_TIMEOUT, op="启动")
        except Exception as exc:
            self.close()
            raise EngineError(
                f"OCR 引擎启动失败：{exc}\n{_probe_runtime(python)}"
            ) from exc
        if not (evt.get("ok") and evt.get("event") == "ready"):
            detail = evt.get("error") or "未知原因"
            self.close()
            raise EngineError(f"OCR 引擎未就绪：{detail}")
        self.engine_name = evt.get("engine") or "PP-OCR"
        if log:
            log(f"OCR 引擎就绪（{self.engine_name} 本地模型，全离线运行）")

    # ------------------------------------------------------------ 协议收发

    def _read_event(self, timeout: float | None = None,
                    op: str = "") -> dict:
        try:
            line = self._lines.get(timeout=timeout) if timeout else self._lines.get()
        except queue.Empty:
            self._abort()
            raise EngineError(
                f"OCR 工作进程无响应（{op or '处理'}超过 {int(timeout or 0)} 秒），"
                "已终止该进程，可重试")
        if line is None:
            time.sleep(0.3)  # 给 stderr 排水线程一点收尾时间
            err = ""
            try:
                err = "".join(self._stderr_tail)[-2000:].strip()
            except Exception:
                pass
            raise EngineError(
                f"OCR 工作进程无响应（可能已崩溃）：{_describe_exit(self.proc.poll())}"
                + (f"\n子进程输出：{err}" if err else "")
            )
        try:
            return json.loads(line)
        except json.JSONDecodeError as exc:
            raise EngineError(f"OCR 工作进程返回了无效响应：{exc}") from exc

    def _abort(self) -> None:
        """强制终止卡死的子进程（随后调用方会重启 worker）。"""
        self._aborted = True
        try:
            self.proc.kill()
        except Exception:
            pass

    def interrupt(self) -> None:
        """取消任务时由外部调用：立刻结束正在阻塞的调用。"""
        self._abort()

    def call(self, req: dict, timeout: float | None = None) -> dict:
        if self._closed or self.proc.poll() is not None:
            detail = ""
            try:
                detail = "".join(self._stderr_tail)[-800:]
            except Exception:
                pass
            raise EngineError(
                f"OCR 工作进程未运行：{_describe_exit(self.proc.returncode)}{detail}")
        self._id += 1
        req = dict(req)
        req["id"] = self._id
        op = str(req.get("op") or "")
        if timeout is None:
            timeout = OCR_READ_TIMEOUTS.get(op, OCR_DEFAULT_TIMEOUT)
        try:
            self.proc.stdin.write(json.dumps(req, ensure_ascii=False) + "\n")
            self.proc.stdin.flush()
        except (BrokenPipeError, OSError) as exc:
            raise EngineError(f"与 OCR 工作进程通信中断：{exc}") from exc
        return self._read_event(timeout=timeout, op=op)

    def call_locked(self, req: dict, timeout: float | None = None) -> dict:
        """带锁的 call：供多线程环境使用（本地管线线程与主线程不会同时进入）。

        stdin/stdout 协议是同步请求-响应，两个线程同时写会让 JSON 行交错、
        响应对不上请求。正常流程已保证 worker 同一时刻只被一个线程使用，
        这里是防御性兜底。
        """
        with self._lock:
            return self.call(req, timeout=timeout)

    def ocr(self, path, clip=None, scale: float = 1.0) -> list[dict]:
        """识别整图或局部区域，返回原图坐标的 box 列表。"""
        req: dict = {"op": "ocr", "path": str(path)}
        if clip is not None:
            req["clip"] = [float(v) for v in clip]
        req["scale"] = float(scale or 1.0)
        reply = self.call_locked(req)
        if not reply.get("ok"):
            raise EngineError(f"OCR 识别失败：{reply.get('error') or reply}")
        if clip is None and float(scale or 1.0) == 1.0:
            self.last_quality = dict(reply.get("quality") or {})
        return reply.get("boxes") or []

    def fast_find(self, paths, max_edge: int = 1680, workers: int = 8) -> list[dict]:
        """Targeted OCR for single-number lookup across one or more images."""
        reply = self.call_locked({
            "op": "fast_find", "paths": [str(path) for path in paths],
            "max_edge": int(max_edge), "workers": int(workers),
        })
        if not reply.get("ok"):
            raise EngineError(f"单号查找失败：{reply.get('error') or reply}")
        return reply.get("results") or []

    def encode(self, path, limit: int = 2048, quality: int = 90,
               rectify: bool = False, clip=None, scale: float = 1.0) -> str:
        """把图片编码为 JPEG base64（供外部模型识别）。

        预处理（透视矫正/旋转/降尺度/裁剪）都在 OCR 子进程内完成，主进程无需 cv2。
        与 ocr 共用 _prepared 缓存：先 encode 再 ocr 不会重复做几何归一化。
        rectify=True 时输出强制透视矫正后的矩形 sheet（去背景拉正），
        视觉接口对该输入的响应速度远优于带背景的原图。
        clip=[x1,y1,x2,y2] 时只编码该矩形（坐标与 path 同坐标系）——外部模型
        只需要单元格的图，绝不外发整张照片；scale 在裁剪后放大，小字建议 2.0。
        """
        req: dict = {"op": "encode", "path": str(path),
                     "limit": int(limit), "quality": int(quality),
                     "rectify": bool(rectify)}
        if clip is not None:
            req["clip"] = [float(v) for v in clip]
        req["scale"] = float(scale or 1.0)
        reply = self.call_locked(req)
        if not reply.get("ok"):
            raise EngineError(f"图片预处理失败：{reply.get('error') or reply}")
        image = reply.get("image")
        if not image:
            raise EngineError("图片编码结果为空")
        return image

    def render_pdf(self, path, dpi: int = 300) -> list[dict]:
        """把 PDF 每页渲染为临时 PNG，返回 [{path,width,height}]。

        渲染在 OCR 子进程内完成（pypdfium2 + cv2），主进程不依赖图像库。
        返回的临时文件由调用方负责清理。
        """
        reply = self.call_locked({"op": "render_pdf", "path": str(path), "dpi": int(dpi)})
        if not reply.get("ok"):
            raise EngineError(f"PDF 渲染失败：{reply.get('error') or reply}")
        return reply.get("pages") or []

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            with self._lock:
                self.proc.stdin.write(json.dumps({"op": "shutdown"}) + "\n")
                self.proc.stdin.flush()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=5)
        except Exception:
            try:
                self.proc.kill()
            except Exception:
                pass
        # openvino / multiprocessing 会再拉工作进程，只杀父进程会留孤儿；
        # 批量长跑时孤儿持续吃内存，最终把整批拖成 OOM（实测残留 6 对进程、
        # 可用内存掉到 1.7G）。这里按进程树清理，失败静默。
        if os.name == "nt":
            try:
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(self.proc.pid)],
                    capture_output=True, timeout=10, check=False)
            except Exception:
                pass


# ---------------------------------------------------------------- 精确复制管线


# 纯本地 OCR 管线：整图 OCR -> 规则级校准 -> 几何重建 -> 序号断档补行
# -> 交叉校验（整图 OCR vs 局部放大 OCR 两方对比）。全程不联网。


# 运单号等编码的前导字母段（如 YT、SF、JD）是固定品牌/类型标识，
# 不是可替换字符——规则校正若按"形状相同、距离 1"去改，会把 YT 误改成
# TT/NT，因此校正前后缀不一致的候选一律排除。


def _prefix_of(text: str) -> str:
    """取文本的前导非数字段（运单号前缀）。

    支持 R02TYT... 这类「字母+数字+字母」前缀：从开头取到连续纯数字段
    之前。普通 YT... 前缀则取全部前导字母。
    """
    s = text.strip()
    if not s:
        return ""
    i = 0
    # 先吃掉前导字母
    while i < len(s) and s[i].isalpha():
        i += 1
    # 若是「字母+短数字+字母」形态（如 R02TYT），继续吃掉中段数字与后续字母
    if i < len(s) and s[i].isdigit():
        j = i
        while j < len(s) and s[j].isdigit():
            j += 1
        if j < len(s) and s[j].isalpha() and (j - i) <= 3:
            k = j
            while k < len(s) and s[k].isalpha():
                k += 1
            return s[:k]
    return s[:i]


def _prefix_compatible(cell: str, candidate: str) -> bool:
    """前缀是否兼容：两侧都无前缀，或前缀相同（大小写不敏感）。

    防止把固定前缀改掉（YT -> TT/NT 这类误纠正）。
    """
    a, b = _prefix_of(cell).upper(), _prefix_of(candidate).upper()
    if not a or not b:
        return True          # 任一侧无前缀，不拦
    return a == b


def _fill_local_tables(worker: "OcrWorker", image_path, blocks, name, log) -> int:
    """对本地重建的表格块做序号断档补行（就地修改 block.rows/row_y）。

    补行需要逐断档做 clip 放大重识别，成本高；仅在本地结果会被实际使用时
    才执行（视觉失败回退本地、或视觉成功但维度不一致需要几何对齐）。
    """
    filled_total = 0
    for block in blocks:
        if getattr(block, "kind", "") != "table":
            continue
        ocr_fn = lambda clip, scale: worker.ocr(image_path, clip=clip, scale=scale)
        try:
            filled_total += fill_serial_gaps(block, ocr_fn, on_log=log)
        except Exception as exc:
            log(f"{name}：序号补行失败（{exc}）")
    if filled_total:
        log(f"{name}：序号断档补行共补充 {filled_total} 行")
    return filled_total


# ---------------------------------------------------------------- 网格法（实拍表格）
#
# 为什么需要它：手机俯拍的表格上，PP-OCR 的文本框常把「日期 + 承运商 + 运单号」
# 粘成一个跨列长框，行/列聚类无从下手（实测 25 行表被重建成 26 行 x 3 列，
# 三列内容挤进一列）。表格自带网格线，列边界由竖线给出，比任何聚类都可靠。

GRID_MIN_FILL_RATIO = 0.35   # 非空单元格占比低于此值认为网格法失败，回退几何重建
GRID_SKIP_FILL_RATIO = 0.10  # 切格后非空率低于此值：网格不可用，跳过补识直接回退
GRID_FILL_BUDGET = 80        # 空单元格放大补识的次数上限（每次一次 clip OCR）

# 编码类单元格（运单号等）的多尺度投票识别。
# 单次放大复识会因裁剪切掉字符边缘而读错（实测 0720457368625 -> 9720457369625）；
# 同一格用多个放大倍数各识一次，取多数共识，可把这类偶发错误压下去。
CODE_SCALES = (2.0, 3.0, 4.0)
CODE_MIN_VOTES = 2           # 至少这个票数才算共识（防单尺度孤证）

# 单号补录（_supplement_grid_codes）总开关。
# v1.4.3 验证结论（_verify_out/v143_conflict_align.md）：透视图上 raw_y→行中心的
# 线性映射有系统性漂移，121 码中 45 码落错行；15 条「补录实得」逐条核对 0 条真实
# 新增、全部写错行（重复+错行污染），比丢码更危险。默认关闭，回到纯网格基线；
# 仅在映射修复实验时临时打开。
GRID_CODE_SUPPLEMENT = False

# 单号局部放大修复（_zoom_repair_codes）总开关。
# v1.4.4 验证（_verify_out/v144_zoom_5.json）：与补录不同，不做线性猜行——
# 单应逆映射定位 + 格子同源尾号自证 + 局部放大双读一致（13 位数字全同）三重
# 判据，5 张目标图 65/69 候选确认且零错行零重复；复用整图 OCR 框不新增整图
# OCR，单张增量 <0.5s（zoom 合计 0.14-0.32s）。
GRID_ZOOM_REPAIR = True


# 网格表头行关键词：warp 顶部 pad（15%）带入的标题区横线是假线（与真实行高
# 等距时 _keep_dense_run 剔不掉），表头行会被抬成空假行；det 框也可能把表头
# 字的中心挤过格线、掉到第一数据行。按关键词重定位表头并把掉字提回。
_HEADER_KEYS = ("日期", "单号", "序号", "承运商", "商品", "名称", "品名",
                "数量", "签收", "状态", "实收", "短少", "收数量")


def _realign_header(rows: list) -> list:
    """表头行重定位，返回新的 rows（表头行 + 数据行）。

    原表头行（rows[0]）命中 >=2 个关键词 -> 原样返回，不干预；
    否则取第一行命中 >=2 个关键词的行作真实表头（其之前的多为标题区假行）；
    表头行的空列上，若下一行同列是 <=4 个纯汉字且命中某个关键词，视为跨线
    掉下来的表头字，提回表头行。
    """
    def hits(row):
        text = " ".join(str(c) for c in row)
        return sum(1 for k in _HEADER_KEYS if k in text)

    if not rows or hits(rows[0]) >= 2:
        return rows
    hi = next((i for i, r in enumerate(rows) if hits(r) >= 2), None)
    if hi is None or hi == 0:
        return rows
    header = list(rows[hi])
    body = [list(r) for r in rows[hi + 1:]]
    if body:
        first = body[0]
        for ci in range(min(len(header), len(first))):
            if str(header[ci]).strip():
                continue
            v = str(first[ci]).strip()
            if v and len(v) <= 4 and re.fullmatch(r"[\u4e00-\u9fff]+", v) \
                    and any(k in v for k in _HEADER_KEYS):
                header[ci] = v      # 跨线掉下来的表头字，提回表头行
                first[ci] = ""
    return [header] + body


def _snap_vlines_to_content(vlines: list, boxes: list, log) -> list:
    """把穿过窄文字框的列线左移到框外（只用于纸张候选小表）。

    第 2 张单号列右线 x=479.5 穿过右侧数字框（bbox 449-491，中心 470 仍在
    单号列），导致 `26 YT… 76` 被并进单号格。这里只对“宽度 ≤2.5 倍行高的
    窄框”（数字/序号）生效，长文本框（单号、商品名）不参与，避免把正常列
    线推走；且只左移，不右移，保持其他列宽不变。
    """
    out = list(vlines)
    if len(out) <= 3:
        return out
    # 只处理第一条“两侧列宽大致均衡（0.6–1.7 倍）且穿过窄数字框”的列线。
    # 第 2 张是单号列右线（370/331）；序号线（109/370）、101 窄列线
    # （331/61）都不均衡，151 列线在目标线之后，break 后不会被动。
    for i in range(1, len(out) - 1):
        left = float(out[i]) - float(out[i - 1])
        right = float(out[i + 1]) - float(out[i])
        if left < 60 or right < 60 or not (0.6 <= left / right <= 1.7):
            continue
        x = float(out[i])
        narrow = []
        for b in boxes:
            try:
                bx1, by1, bx2, by2 = b["bbox"]
            except (KeyError, TypeError, ValueError):
                continue
            h = max(1.0, float(by2) - float(by1))
            if float(bx1) < x < float(bx2) and (float(bx2) - float(bx1)) <= 2.5 * h:
                narrow.append(b)
        if len(narrow) < 2:
            continue
        new_x = min(float(b["bbox"][0]) for b in narrow) - 6.0
        if out[i - 1] + 30 < new_x < x:
            log(f"列线吸附：x {x:.0f} -> {new_x:.0f}"
                f"（跨线窄框 {len(narrow)} 个）")
            out[i] = new_x
            break
    return out


def _grid_cell_grid(hlines: list, vlines: list):
    """由网格线构造二维单元格：返回 (表头行 cells, 数据行 cells)。"""
    header = [(c, hlines[0], d, hlines[1]) for c, d in zip(vlines, vlines[1:])]
    body = [[(c, a, d, b) for c, d in zip(vlines, vlines[1:])]
            for a, b in zip(hlines[1:-1], hlines[2:])]
    return header, body


def _join_boxes(items: list) -> str:
    """同一格内多个 OCR 框按 x 序拼接；水平间隙大于半个字高时补空格。"""
    parts: list[str] = []
    prev_x2 = None
    for b in items:
        text = str(b.get("text", "")).strip()
        if not text:
            continue
        x1, _y1, x2, _y2 = b["bbox"]
        if prev_x2 is not None and x1 - prev_x2 > 0.5 * max(1.0, _y2 - _y1):
            parts.append(" ")
        parts.append(text)
        prev_x2 = x2
    return "".join(parts)


def _place_boxes(boxes: list, cells: list) -> list:
    """把 OCR 框按中心点落进单元格，返回与 cells 同形的二维文本。"""
    out = []
    for row in cells:
        texts = []
        for (x1, y1, x2, y2) in row:
            inside = []
            for b in boxes:
                bx1, by1, bx2, by2 = b["bbox"]
                cx, cy = (bx1 + bx2) / 2.0, (by1 + by2) / 2.0
                if x1 <= cx <= x2 and y1 <= cy <= y2:
                    inside.append(b)
            inside.sort(key=lambda b: b["bbox"][0])
            texts.append(_join_boxes(inside))
        out.append(texts)
    return out


def _fill_empty_cells(cells: list, texts: list, worker, warp_path, log, budget: list) -> int:
    """对空格子做 clip 放大补识（就地写入 texts），返回填充数。"""
    filled = 0
    pad = 4
    for ri, row in enumerate(cells):
        for ci, (x1, y1, x2, y2) in enumerate(row):
            if texts[ri][ci] or budget[0] <= 0:
                continue
            budget[0] -= 1
            try:
                boxes = worker.ocr(warp_path, clip=(x1 - pad, y1 - pad, x2 + pad, y2 + pad),
                                   scale=2.0)
            except Exception:
                continue
            text = _join_boxes(boxes)
            if text:
                texts[ri][ci] = text
                filled += 1
    if filled:
        log(f"网格补识：填充 {filled} 个空单元格")
    return filled


def _recognize_code_cell(worker, warp_path, cell, pad: int = 4) -> str:
    """编码类单元格多尺度投票识别：2x/3x/4x 各识一次，取多数共识。

    返回票数最高的识别结果；无共识（各尺度互不相同）时返回空串，由调用方
    保留原落格值。投票按"规范化文本"计（忽略空格），输出取共识票中最长的
    一个（尺度越大字符越全，偶发的边缘截断会让结果偏短）。
    """
    x1, y1, x2, y2 = cell
    clip = (x1 - pad, y1 - pad, x2 + pad, y2 + pad)
    votes: Counter = Counter()
    for scale in CODE_SCALES:
        try:
            boxes = worker.ocr(warp_path, clip=clip, scale=scale)
        except Exception:
            continue
        text = _join_boxes(boxes).strip()
        if not text:
            continue
        votes[text.replace(" ", "")] += 1
    if not votes:
        return ""
    top = votes.most_common()
    best_n = top[0][1]
    if best_n < CODE_MIN_VOTES:
        return ""          # 各尺度说法不一，不猜
    winners = [t for t, n in top if n == best_n]
    return max(winners, key=len)


def _text_lines(boxes: list) -> list:
    """网格外的 OCR 框按行聚成文字行（y 相邻的合并，行内按 x 序拼接）。"""
    if not boxes:
        return []
    items = sorted(boxes, key=lambda b: (b["bbox"][1], b["bbox"][0]))
    rows: list[list[dict]] = []
    for b in items:
        cy = (b["bbox"][1] + b["bbox"][3]) / 2.0
        if rows and abs(cy - rows[-1][0]["_cy"]) <= max(6.0, (b["bbox"][3] - b["bbox"][1]) * 0.6):
            rows[-1].append(b)
        else:
            b["_cy"] = cy
            rows.append([b])
    out = []
    for row in rows:
        row.sort(key=lambda b: b["bbox"][0])
        out.append(_join_boxes(row))
    return [t for t in out if t]


_WAYBILL_LIKE_RE = re.compile(r"(?:R\d{2}T)?YT\d{10,}", re.I)


def _guess_code_columns(body_texts, min_hits: int = 5,
                        min_ratio: float = 0.3) -> list[int]:
    """表头缺失或为空时，按内容猜哪些列是单号列。

    双列版式的实测表格（c2c7b955）表头行被网格法识别成空，靠表头文字
    定位单号列会全部落空，右半边的单号就没人补识。这里退化为看内容：
    一列里「含 YT 串的非空格」占比够高，就按单号列处理。
    """
    if not body_texts:
        return []
    ncol = max(len(r) for r in body_texts)
    out = []
    for ci in range(ncol):
        texts = [str(r[ci]).strip() for r in body_texts if ci < len(r)]
        texts = [t for t in texts if t]
        if len(texts) < min_hits:
            continue
        hits = sum(1 for t in texts if _WAYBILL_LIKE_RE.search(t))
        if hits >= min_hits and hits / len(texts) >= min_ratio:
            out.append(ci)
    return out


def _fill_code_column_full(worker, warp_path, cells, texts, log,
                           chunk: int = 10, scale: float = 1.8) -> int:
    """【已停用，勿接入主流程】单号列按段放大识别并按行补空。

    实测负收益：空白格补出来的常常是表头文字或相邻列内容，会让真实的单号行
    被 `_looks_like_repeated_header` 当成重复表头跳过（be680b97 丢 6 条、
    fc03c6d3 丢 6 条、890f24fd 丢 2 条），或与相邻列数字粘连
    （815bfb1c B2 由 13 位变成 15 位）。全量 45 张对比：净 -5 条且 18 条质量退化。

    保留说明：技术本身可行（按 chunk 行分段、worker.ocr 返回原图坐标不要除 scale），
    但必须先解决"补出来的字属于哪一列"的归属判断，才能重新启用。
    """
    usable = [(i, c[0]) for i, c in enumerate(cells) if c]
    if not usable or not texts:
        return 0
    x1 = min(b[0] for _i, b in usable)
    x2 = max(b[2] for _i, b in usable)
    pad = 3
    filled = 0
    for start in range(0, len(usable), chunk):
        part = usable[start:start + chunk]
        clip = [int(x1 - pad), int(min(b[1] for _i, b in part) - pad),
                int(x2 + pad), int(max(b[3] for _i, b in part) + pad)]
        try:
            boxes = worker.ocr(warp_path, clip=clip, scale=scale)
        except Exception as exc:
            log(f"单号列分段补识失败（{exc}）")
            continue
        if not boxes:
            continue
        centers = {i: (b[1] + b[3]) / 2.0 for i, b in part}
        heights = {i: max(8.0, b[3] - b[1]) for i, b in part}
        for box in boxes:
            text = str(box.get("text", "")).strip()
            if not text:
                continue
            bbox = box.get("bbox") or [0, 0, 0, 0]
            cy = (float(bbox[1]) + float(bbox[3])) / 2.0
            ri = min(centers, key=lambda i: abs(centers[i] - cy))
            if abs(centers[ri] - cy) > heights[ri] / 1.5:
                continue
            if texts[ri][0]:
                continue
            texts[ri][0] = text
            filled += 1
    if filled:
        log(f"单号列分段补识：补上 {filled} 格")
    return filled


_MEMORY_ERROR_HINTS = ("unable to allocate", "out of memory", "memoryerror",
                       "cannot allocate", "bad_alloc", "not enough memory")


def _is_memory_error(exc) -> bool:
    """判断异常是否内存不足。

    批量长跑时 OCR worker 常驻，内存会累积；这类错误必须让上层看到，
    好让调用方重启 worker 后重试当前页，不能当普通失败吞掉。
    """
    return any(h in str(exc).lower() for h in _MEMORY_ERROR_HINTS)


def _code_prefix_ok(a: str, b: str) -> bool:
    """两串「字母部分」是否一致，用来挡住 YT→XT/NT 这类换字母的误识。

    数字差异是允许的（末位错读是主要错误来源），字母部分必须完全一样；
    至少 2 个字母才算有前缀可比（YT、R02TYT 都满足）。
    """
    aa = re.sub(r"[^0-9A-Za-z]", "", a).upper()
    bb = re.sub(r"[^0-9A-Za-z]", "", b).upper()
    la = "".join(ch for ch in aa if ch.isalpha())
    lb = "".join(ch for ch in bb if ch.isalpha())
    return len(la) >= 2 and la == lb


def _enhance_code_column(worker, warp_path, cells, texts, log,
                         max_votes: int = 24) -> int:
    """【已修复但未接入，勿直接启用】运单号列多尺度投票纠错。

    历史问题：依赖 `xcheck._prefix_compatible`，该函数并不存在，导入即抛，
    被上层 except 吞掉 —— 这个函数从上线起从未真正执行过。
    现已改为本地 `_code_prefix_ok`（比字母部分，防 YT→XT）。

    未接入原因：它会**改写**落格值（条件：位数相同 + 字母部分一致），
    从未在真实数据上验证过准确性；启用前需要先建"改对/改错"的对照样本。
    max_votes 限制单列投票次数（每格 3 次 clip OCR，长列全量跑会拖慢整页）。
    """
    fixed = 0
    votes = 0
    for ri, row in enumerate(cells):
        for ci, (x1, y1, x2, y2) in enumerate(row):
            cur = str(texts[ri][ci]).strip()
            if not cur:
                continue
            if votes >= max_votes:
                log(f"[多尺度] 本列投票已达上限 {max_votes} 次，其余保持落格值")
                return fixed
            votes += 1
            voted = _recognize_code_cell(worker, warp_path, (x1, y1, x2, y2))
            if not voted or voted == cur:
                continue
            if len(voted) != len(cur) or not _code_prefix_ok(cur, voted):
                log(f"[多尺度] 第{ri + 1}行第{ci + 1}列 保留 {cur}"
                    f"（投票 {voted} 位形/前缀不兼容）")
                continue
            log(f"[多尺度] 第{ri + 1}行第{ci + 1}列 {cur} -> {voted}（多尺度共识）")
            texts[ri][ci] = voted
            fixed += 1
    return fixed


def _normalize_date_cell(text: str) -> str:
    """从单元格文本提取规范日期 YYYY.MM.DD；解析不出或信息残缺返回空串。

    残缺（如 "2026.10.2" 丢了日最后一位、月只剩 "1" 时原值可能是 10/11/12）
    返回空串、由调用方保留原值：补零臆造会把 29 变成 02，残缺值只能靠列内
    家族共识（与多数派日期比对丢位）来恢复。共享 ai_calibrate.parse_date，
    行号粘连（"12026.10.29"）、丢分隔符（"202610.29"）、行号+紧凑
    （"220061029"）都由它一次性处理。
    """
    r = parse_date(text)
    return r[0] if r and r[1] else ""


def _enhance_date_column(worker, warp_path, cells, texts, log) -> int:
    """日期列增强：整列一次 clip 识别 → 按行分配 → 解析日期 → 列内共识兜底。

    整列一次识别优于逐格：行高仅 40px，逐格 clip 只有几个字符宽，det 容易
    漏框（实测逐格全空，整列一次 25 行都能读到）。
    cells/texts 均为「每行一个子列表、子列表内只有该行日期格」的同形结构。
    残缺行（如 "2026.10.2" 丢了日末位）解析不出规范值，保持原值，交给后续
    日期列家族共识按"丢位"比对恢复；届时仍救不回的就在日志里点名保留，
    绝不补零臆造。
    """
    rows = len(cells)
    if rows == 0:
        return 0
    try:
        flat = [c[0] for c in cells if c]
        if not flat:
            return 0
        x1 = min(c[0] for c in flat)
        x2 = max(c[2] for c in flat)
        y1 = min(c[1] for c in flat)
        y2 = max(c[3] for c in flat)
        boxes = worker.ocr(warp_path, clip=(x1 - 4, y1 - 4, x2 + 4, y2 + 4),
                           scale=2.0)
    except Exception as exc:
        log(f"[整列] 日期列识别失败（{exc}）")
        return 0
    if not boxes:
        return 0

    parsed: list[str] = []
    for ri in range(rows):
        cell = cells[ri][0]
        cy1, cy2 = cell[1], cell[3]
        inside = [b for b in boxes
                  if cy1 - 2 <= (b["bbox"][1] + b["bbox"][3]) / 2 <= cy2 + 2]
        parsed.append(_normalize_date_cell(_join_boxes(inside)))

    hit = [p for p in parsed if p]
    major, major_n = Counter(hit).most_common(1)[0] if hit else ("", 0)
    if major and major_n < 0.4 * max(1, len(hit)):
        major = ""      # 共识不足，不用共识兜底（可能是跨天混排）

    changed = 0
    major_digits = "".join(c for c in major if c.isdigit())
    for ri in range(rows):
        old = str(texts[ri][0]).strip()
        new = parsed[ri]
        if not new and major:
            # 只有原值能证明是多数日期恰好漏了一位时才补齐；空白、跨天或
            # 无法解析的单元格保持原样，不能静默篡改成列内的多数日期。
            core = _incomplete_date_digits(old)
            if core and _one_digit_lost(core, major_digits):
                new = major
        if new and new != old:
            log(f"[整列] 第{ri + 1}行日期 {old!r} -> {new}")
            texts[ri][0] = new
            changed += 1
    return changed


_TITLE_PROTECT = ("序号", "运单号", "包裹号", "日期", "承运商", "商品", "数量",
                  "接收", "交接", "拒收", "破损", "件数", "单号", "实收")


def _trim_title_rows(rows: list, row_y: list) -> tuple[list, list, list]:
    """剥掉表格顶部混入的标题行（只认无单号、无列名、短文本/含“退回”的行）。

    第 2 张“退货清单”没有表头，网格把纸张顶部的标题/日期当成前几行，
    Excel 数据区第一行成了“一/百井/退货清单…”。这里只从最前面连续剥离，
    一遇到单号或列名关键词就停；标题原文转成 TextBlock 保留，不丢内容。
    """
    drop = 0
    for i in range(min(3, len(rows) - 1)):
        text = "".join(str(c or "") for c in rows[i]).strip()
        if not text:
            drop += 1
            continue
        if re.search(r"YT\d{9,}", text):
            break
        if any(k in text for k in _TITLE_PROTECT):
            break
        if any(k in text for k in ("退回", "清单", "标题")) or len(text) <= 8:
            drop += 1
            continue
        break
    if drop == 0:
        return rows, row_y, []
    title = [" ".join(str(c).strip() for c in row if str(c).strip())
             for row in rows[:drop]]
    return rows[drop:], row_y[drop:], [t for t in title if t]


def _normalize_code_rows(rows: list, log) -> tuple[list, list]:
    """单号列清理：剥离前置序号；双单号安全拆分，第二个只记待复核。

    例：`3 YT1234567890123YT1234567890456`
      -> 单元格保留 `YT1234567890123`，第二个单号进 flags，不硬塞进别的行。
    只处理合法 `YT` + 13 位；`YT…签收` 这类粘连不动，交给现有交叉校验。
    """
    if not rows:
        return rows, []
    n_cols = max(len(r) for r in rows)
    if n_cols > 8:
        # 第 1 张这类大表里序号与单号同格，剥离会动数据行；本轮只处理
        # “退货清单”这种 6-7 列小表，保证数据行不变。
        return rows, []
    counts = []
    for ci in range(n_cols):
        counts.append(sum(len(re.findall(r"YT\d{13}", str(r[ci])))
                         for r in rows if ci < len(r)))
    if not counts:
        return rows, []
    ci = max(range(len(counts)), key=lambda j: counts[j])
    if counts[ci] < 2:
        return rows, []

    flags = []
    for ri, row in enumerate(rows):
        if ci >= len(row):
            continue
        original = str(row[ci] or "").strip()
        if not original:
            continue
        rest = original
        m = re.match(r"^(\d{1,3})\s*(YT\d{13}.*)$", original)
        if m:
            rest = m.group(2).strip()
        codes = re.findall(r"YT\d{13}", rest)
        if len(codes) >= 2:
            first, second = codes[0], codes[1]
            row[ci] = first
            flags.append({
                "row": ri, "col": ci,
                "rebuilt": first, "whole": original, "local": second,
                "reason": "双单号粘连，第二个单号待复核",
            })
            log(f"第{ri + 1}行第{ci + 1}列：双单号分离，保留 {first}，"
                f"{second} 转入待复核")
        elif len(codes) == 1:
            code = codes[0]
            if rest == code:
                row[ci] = code
                if m:
                    log(f"第{ri + 1}行第{ci + 1}列：剥离单号前置序号 {original!r} -> {code}")
            else:
                idx = rest.find(code)
                prefix = rest[:idx].strip()
                suffix = rest[idx + len(code):].strip()
                # 只剥“纯噪声前缀”（数字/汉字/空白，最多 8 字），且后缀为空；
                # 前缀含签收状态词的不在这里动（第 1 张由 _split_code_status 处理）。
                noisy = bool(prefix) and bool(
                    re.fullmatch(r"[0-9\u4e00-\u9fff\s\-_.]{1,8}", prefix))
                if (not suffix and noisy
                        and not any(s in prefix for s in ("签收", "拒收", "完好", "破损"))):
                    row[ci] = code
                    flags.append({
                        "row": ri, "col": ci,
                        "rebuilt": code, "whole": original, "local": prefix,
                        "reason": "单号前置噪声已剥离，原值待复核",
                    })
                    log(f"第{ri + 1}行第{ci + 1}列：剥离单号前置噪声 "
                        f"{original!r} -> {code}")
    return rows, flags


# ---------------------------------------------------------------- 宽表序号列修复

# 仓配交接清单的固定 16 列列名（照片经网格切分后表头常被并成 2-4 个长格，
# 序号列/包裹号列也会被并成一列：'序号+包裹号' -> '717652014667960'）。
_HANDOFF_HEADER = (
    "序号", "包裹号", "签收状态", "包装状态", "包裹破损原因", "拒收原因",
    "商品条码", "签收数量", "完好数量", "破损数量", "拒收数量", "LT数量",
    "DE-XS", "DE-HS", "商品拒收原因", "商品破损原因",
)
_LONG_DIGITS_RE = re.compile(r"\d{9,}")
_SERIAL_TEXT_RE = re.compile(r"^\d{1,3}$")
# "序号+包裹号"粘连：1-3 位数字前缀 + 仍像包裹号的后缀
_SERIAL_GLUE_RE = re.compile(r"^(\d{1,3})\s*([A-Za-z]{0,4}\s*\d{9,})$")
# 字母/数字短前缀 + 空格 + 纯数字包裹号（'N 1765…' / 'S4 1765…'）
_LETTER_GLUE_RE = re.compile(r"^[A-Za-z0-9]{1,3}\s+(\d{9,})$")


def _locate_serial_pkg_cols(rows: list) -> tuple:
    """定位序号列与包裹号列：先按表头关键词，再按数据形状回退。

    返回 (serial_idx, pkg_idx)，无法定位时返回 (None, None)。
    表头被并格后关键词可能全丢（8a2b 表头是"弹导 权票号"），此时用前 4 列
    里"长数字串"最多的列当包裹号列，其左邻列当序号列。
    """
    n_cols = max((len(r) for r in rows), default=0)
    for row in rows[:2]:
        cells = [str(c or "") for c in row]
        sidx = next((i for i, h in enumerate(cells) if "序号" in h), None)
        pidx = next((i for i, h in enumerate(cells)
                     if i != sidx and any(k in h for k in ("包裹号", "运单号", "快递号"))), None)
        if pidx is not None:
            if sidx is None and pidx > 0:
                sidx = pidx - 1
            return sidx, pidx
    counts = []
    for ci in range(min(n_cols, 4)):
        counts.append(sum(1 for r in rows[1:] if ci < len(r)
                          and _LONG_DIGITS_RE.search(str(r[ci] or ""))))
    if not counts or max(counts) < 5:
        return None, None
    pidx = counts.index(max(counts))
    if pidx == 0:
        return None, None
    return pidx - 1, pidx


def _apply_handoff_header(rows: list, sidx: int | None) -> bool:
    """把仓配交接清单的表头还原成 16 个标准列名（仅该表型生效）。

    触发条件：表头含"商品条码"且含"LT数量/DE-XS/DE-HS"之一，网格列数
    16-19，且从序号列起放得下 16 个列名。列名从序号列开始写，左侧细条
    列（纸张边缘）与右侧多余列保持原样。
    """
    if not rows or sidx is None:
        return False
    n_cols = max((len(r) for r in rows), default=0)
    if not 16 <= n_cols <= 19 or sidx + len(_HANDOFF_HEADER) > n_cols:
        return False
    head = " ".join(str(c or "") for c in rows[0])
    if "商品条码" not in head:
        return False
    if not any(k in head for k in ("LT数量", "DE-XS", "DE-HS")):
        return False
    row = rows[0]
    while len(row) < n_cols:
        row.append("")
    for k, name in enumerate(_HANDOFF_HEADER):
        row[sidx + k] = name
    return True


def _fix_wide_table_serials(rows: list, log) -> tuple[list, list]:
    """宽表（>8 列）序号列修复：拆"序号+包裹号"粘连 + 按连续序列补齐。

    网格列线在左缘缺失/照片歪斜时，序号与包裹号会被并进同一格
    （'717652014667960' = 71 + 17652014667960，'72T7652478898830'），
    组首行的序号也可能整格为空（fc03 缺 1/4/5/6/8/9）。这里只改写
    序号列与粘连的包裹号格：用包裹号行拟合 start..start+N-1 的连续序号，
    命中率 >=70%（且 >=6 行）才应用，序列不成立时原样返回。
    """
    if not rows or max((len(r) for r in rows), default=0) <= 8:
        return rows, []
    sidx, pidx = _locate_serial_pkg_cols(rows)
    if sidx is None or pidx is None:
        return rows, []
    groups, originals, hints, glued = [], {}, {}, {}
    # 交接清单的"签收状态"列固定在序号列 +2；组首行可能包裹号整格糊掉
    # （8a2b 第 88 行只剩 '88396 4R'），此时靠签收状态仍能认出这是组首行。
    status_idx = sidx + 2
    status_ok = (status_idx < len(rows[0]) and
                 "签收" in str(rows[0][status_idx]) and
                 "状态" in str(rows[0][status_idx]))
    for ri in range(1, len(rows)):
        row = rows[ri]
        if pidx >= len(row):
            continue
        pkg = str(row[pidx] or "").strip()
        status = str(row[status_idx] or "").strip() if status_ok and status_idx < len(row) else ""
        has_status = bool(re.search(r"签收|拒收", status) and len(status) <= 6)
        if not _LONG_DIGITS_RE.search(pkg) and not has_status:
            continue
        groups.append(ri)
        originals[ri] = pkg
        serial = str(row[sidx] or "").strip() if sidx < len(row) else ""
        if _SERIAL_TEXT_RE.match(serial):
            hints[ri] = {int(serial)}
            continue
        # 粘连格的切分点有 1-3 位三种可能（'8017651845870541' 的序号可能是
        # 8/80/801），全部作为候选提示，交给连续序列投票选真值。
        cands = set()
        for cut in (1, 2, 3):
            if len(pkg) <= cut or not pkg[:cut].isdigit():
                continue
            if _LONG_DIGITS_RE.search(pkg[cut:]):
                cands.add(int(pkg[:cut]))
        if cands:
            hints[ri] = cands
            glued[ri] = True
    n = len(groups)
    if n < 8:
        return rows, []
    best_s0, best_hits = None, 0
    for k, ri in enumerate(groups):
        for hint in hints.get(ri, ()):
            s0 = hint - k
            hits = sum(1 for j, rj in enumerate(groups) if s0 + j in hints.get(rj, ()))
            if hits > best_hits or (hits == best_hits and best_s0 is not None and s0 < best_s0):
                best_s0, best_hits = s0, hits
    if best_s0 is None or best_hits < max(6, (7 * n + 9) // 10):
        if os.environ.get("IMG2DOC_DEBUG_SERIAL"):
            log(f"宽表序号调试：sidx={sidx} pidx={pidx} n={n} best={best_s0} "
                f"hits={best_hits} need={(7 * n + 9) // 10}")
            for k, ri in enumerate(groups):
                log(f"  组{k} 行{ri}: 序号={rows[ri][sidx]!r} "
                    f"提示={sorted(hints.get(ri, ()))} 包裹={originals[ri]!r}")
        return rows, []
    # 组首行已能明确读到连续序号（如 1,2,3）时，以眼前的序列为准：
    # 个别行被合并/丢失时，全表多数投票会把起点整体后移（bb7bfdf0 有 44 行
    # 支持起点 2，但前 3 行明明是 1,2,3），不修正就会把正确序号整体 +1。
    head_fix, run = {}, []
    for k, ri in enumerate(groups):
        hs = hints.get(ri)
        if glued.get(ri) or not hs or len(hs) != 1:
            break
        v = next(iter(hs))
        if run and v != run[-1][1] + 1:
            break
        run.append((k, v))
    if (len(run) >= 2 and run[0][1] < best_s0
            and run[-1][1] < best_s0 + run[-1][0]):
        head_fix = dict(run)
    flags, changed = [], 0
    for k, ri in enumerate(groups):
        want = head_fix.get(k, best_s0 + k)
        row = rows[ri]
        while len(row) <= max(sidx, pidx):
            row.append("")
        old = str(row[sidx] or "").strip()
        if old != str(want):
            if old:
                flags.append({
                    "row": ri, "col": sidx,
                    "rebuilt": str(want), "whole": old, "local": "",
                    "reason": "序号按连续序列校正",
                })
            row[sidx] = str(want)
            changed += 1
        pkg = originals[ri]
        if glued.get(ri):
            # 按已定序号断开粘连：序号位数决定切点（90|1765… 而不是 90|1|765…）
            cut = len(str(want))
            rest = pkg[cut:].lstrip() if pkg[:cut].isdigit() else ""
            if not rest or not _LONG_DIGITS_RE.search(rest):
                m = _SERIAL_GLUE_RE.match(pkg)
                rest = m.group(2).strip() if m else ""
            if rest and _LONG_DIGITS_RE.search(rest):
                row[pidx] = rest
        elif not old:
            m = _LETTER_GLUE_RE.match(pkg)
            if m:
                row[pidx] = m.group(1)
    # 首组之前的空壳行（8a2b 首行左半是空的，补识把表头残字读成 '1'/'Li1'）：
    # 序号小于序列起点且包裹号不像包裹号时清空，避免凭空多出一个“1 号”。
    first = groups[0]
    for ri in range(1, first):
        row = rows[ri]
        if sidx >= len(row) or pidx >= len(row):
            continue
        serial = str(row[sidx] or "").strip()
        pkg = str(row[pidx] or "").strip()
        if (_SERIAL_TEXT_RE.match(serial) and int(serial) < best_s0
                and pkg and not _LONG_DIGITS_RE.search(pkg)):
            row[sidx] = ""
            row[pidx] = ""
    if changed:
        log(f"宽表序号列：按连续序列 {best_s0}..{best_s0 + n - 1} 校正 {changed} 行"
            f"（{best_hits}/{n} 行命中）")
    return rows, flags


# ---------------------------------------------------------------- 包裹号前缀还原
#
# 交接清单的包裹号固定为 `YT` + 13 位数字。网格透视矫正后包裹号列左侧的竖线
# 常缺失、文字整体左漂，OCR 检测框会从 `YT` 之后起框，前缀被整段吞掉
# （8a2b 31 条里 27 条如此）：`7652014660000`（YT 丢失）、`T7652478890000`
# （Y 丢失）、`17652867900000`（YT 读成 1）、`TYT7652686400000`（多读一个 T）。
# 单靠放大倍数无效（框仍从数字起），只能在文本层按列内格式还原前缀；
# 剩下的位数/乱码异常再做行带左扩复识。

_PKG_YT_RE = re.compile(r"^YT\d{13}$")
# 13-14 位数字，前面最多 3 个字母/竖线/1 的误读前缀（用于识别"这是包裹号列"）
_PKG_SHAPE_RE = re.compile(r"^[A-Za-z|1]{0,3}\d{13,14}$")
_PKG_RULES = (
    re.compile(r"^TYT\d{13}$"),      # 多读一个 T
    re.compile(r"^\d{13}$"),         # YT 整段丢失
    re.compile(r"^T\d{13}$"),        # Y 丢失
    re.compile(r"^[1NIL|7]\d{13}$"),  # YT 被读成 1/N/I/L/|/7（70086422100000）
    re.compile(r"^[1NIY]\d{14}$"),   # 前缀多占一位（N1 785…）
)

# 包裹号左扩复识总开关（只对整列大面积破损的表生效，见 _locate_pkg_col）。
# 诊断/回滚：IMG2DOC_PKG_REPAIR=0 关闭整套包裹号修复。
GRID_PKG_REPAIR = os.environ.get("IMG2DOC_PKG_REPAIR", "1") != "0"


def _locate_pkg_col(rows: list) -> tuple:
    """定位交接清单的包裹号列并统计列内形态。

    返回 (col_idx, stat)；stat = {"total": 非空条数, "yt": YT+13 条数}。
    只有"确实是 YT 包裹号列"时才返回列号：至少 2 条已是 `YT`+13 位，
    且 >=60% 的非空格是 13-14 位数字形态（挡住商品条码列、单号列）。
    """
    if not rows:
        return None, {}
    n_cols = max((len(r) for r in rows), default=0)
    if n_cols < 6:
        return None, {}
    ci = None
    for row in rows[:2]:
        for j, cell in enumerate(row):
            if any(k in str(cell or "") for k in ("包裹号", "运单号", "快递号")):
                ci = j
                break
        if ci is not None:
            break
    if ci is None:
        # 表头被并格读坏（8a2b 曾读成"弹导 权票号"）：按值形态兜底
        best, best_n = None, 0
        for j in range(n_cols):
            n = sum(1 for r in rows[1:] if j < len(r)
                    and _PKG_SHAPE_RE.match(str(r[j] or "").strip()))
            if n > best_n:
                best, best_n = j, n
        if best is None or best_n < 6:
            return None, {}
        ci = best
    vals = [str(r[ci] or "").strip() for r in rows[1:] if ci < len(r)]
    vals = [v for v in vals if v]
    if len(vals) < 6:
        return None, {}
    shape = sum(1 for v in vals if _PKG_SHAPE_RE.match(v))
    yt = sum(1 for v in vals if _PKG_YT_RE.match(v))
    if yt < 2 or shape < 0.6 * len(vals):
        return None, {}
    return ci, {"total": len(vals), "yt": yt}


def _pkg_rule_fix(cell: str) -> str | None:
    """按列内 `YT`+13 位格式可确定性还原时返回新值，否则 None。"""
    for pat in _PKG_RULES:
        if pat.match(cell):
            cand = "YT" + cell[-13:]
            return cand if _PKG_YT_RE.match(cand) else None
    return None


def _normalize_pkg_prefix(rows: list, ci: int, log, flagged=None,
                          warnings=None) -> int:
    """按列内 `YT`+13 位格式还原被吞掉/读错的前缀（就地改写）。

    只做"补回 YT、数字只取末 13 位"的确定性还原；14 位纯数字、位数异常、
    带签收状态词的格子不动，留给 `_reread_invalid_pkg` 用像素证据处理。
    改写后清掉该格在交叉校验阶段留下的旧待复核项（整图/放大对比的是改写前
    的值，继续黄底只会指向一个已经修好的格子）。
    """
    fixed = 0
    for ri in range(1, len(rows)):
        row = rows[ri]
        if ci >= len(row):
            continue
        cell = str(row[ci] or "").strip()
        if not cell or _PKG_YT_RE.match(cell):
            continue
        new = _pkg_rule_fix(cell)
        if not new:
            continue
        row[ci] = new
        fixed += 1
        log(f"第{ri + 1}行第{ci + 1}列：包裹号前缀还原 {cell} -> {new}")
        if isinstance(flagged, list):
            flagged[:] = [f for f in flagged
                          if not (int(f.get("row", -1)) == ri
                                  and int(f.get("col", -1)) == ci)]
        if isinstance(warnings, list):
            head = f"待复核：第{ri + 1}行第{ci + 1}列"
            warnings[:] = [w for w in warnings if not str(w).startswith(head)]
    if fixed:
        log(f"包裹号前缀还原 {fixed} 处（列内格式为 YT+13 位）")
    return fixed


def _reread_invalid_pkg(block, ci: int, worker, ref_path, name, log,
                        warnings: list, flagged: list) -> int:
    """对仍不合规的包裹号做"行带 × 包裹号列左扩 0.15 列宽"的放大复识。

    前缀被吞是检测框从 YT 之后起框，向左扩一点才能把前缀圈进来；
    3x/4x 双尺度都读出 `YT`+13 且得分 >=0.90 才改写；单尺度 >=0.95 且
    末 13 位与格内原值一致（只是补前缀）也改写；其余保持原值并标待复核。
    """
    rows = getattr(block, "rows", None) or []
    col_x = list(getattr(block, "col_x", []) or [])
    row_y = list(getattr(block, "row_y", []) or [])
    if not rows or ci >= len(col_x) or not row_y:
        return 0
    if len(row_y) > len(rows):
        # _trim_title_rows 只从行首裁假行；数据行对应 row_y 尾部切片
        row_y = row_y[-len(rows):]
    if len(row_y) != len(rows):
        return 0
    x0, x1 = float(col_x[ci][0]), float(col_x[ci][1])
    pad = 0.15 * max(1.0, x1 - x0)
    fixed = review = 0
    for ri in range(1, min(len(rows), 41)):
        row = rows[ri]
        if ci >= len(row):
            continue
        cell = str(row[ci] or "").strip()
        if _PKG_YT_RE.match(cell):
            continue
        a, b = row_y[ri]
        clip = [max(0.0, x0 - pad), float(a), x1, float(b)]
        found: dict[str, dict] = {}
        for scale in (3.0, 4.0):
            try:
                boxes = worker.ocr(str(ref_path), clip=clip, scale=scale) or []
            except Exception:
                boxes = []
            for box in boxes:
                txt = re.sub(r"\s+", "", str(box.get("text", "")))
                for code in re.findall(r"YT\d{13}", txt, re.I):
                    code = code.upper()
                    rec = found.setdefault(code, {"score": 0.0, "scales": set()})
                    rec["score"] = max(rec["score"], float(box.get("score") or 0.0))
                    rec["scales"].add(scale)
        ok = next((c for c, r in found.items()
                   if len(r["scales"]) == 2 and r["score"] >= 0.90), None)
        if ok is None:
            old_d = re.sub(r"\D", "", cell)
            ok = next((c for c, r in found.items()
                       if r["score"] >= 0.95
                       and (not old_d or old_d[-13:] == c[-13:])), None)
        if ok is None:
            # 列内格式能确定性还原的格子由 _normalize_pkg_prefix 负责，
            # 这里不重复记待复核（否则整列 20+ 条黄底噪声）。
            if cell and _pkg_rule_fix(cell) is None:
                review += 1
                flagged.append({
                    "row": ri, "col": ci, "rebuilt": cell, "whole": cell,
                    "local": "", "reason": "包裹号前缀/位数异常，左扩复识未确认",
                })
                warnings.append(
                    f"待复核：第{ri + 1}行第{ci + 1}列「{cell}」"
                    f"（包裹号前缀/位数异常，左扩复识未确认）"
                )
                log(f"{name}：包裹号复识未确认，保留原值：{cell}")
            continue
        if ok != cell:
            row[ci] = ok
            fixed += 1
            log(f"{name}：第{ri + 1}行第{ci + 1}列包裹号左扩复识 {cell} -> {ok}")
    if fixed or review:
        log(f"{name}：包裹号左扩复识改写 {fixed} 处，待复核 {review} 处")
    return fixed


_RECEIVED_KEYS = ("实收",)


_STATUS_WORD = ("签收", "拒收", "完好", "破损")


def _split_code_status(rows: list, flags: list, log) -> int:
    """大表单号与签收状态粘连：`YT…签收` -> 单号，状态词进待复核。

    只处理单号列、单号后紧跟状态词且后缀只有状态词/分隔符的单元格；
    后缀带其他内容的（如 `签收数5`）不硬拆，只标待复核说明。
    """
    if not rows:
        return 0
    n_cols = max(len(r) for r in rows)
    if n_cols <= 8:
        return 0
    counts = [sum(len(re.findall(r"YT\d{13}", str(r[ci])))
                  for r in rows if ci < len(r)) for ci in range(n_cols)]
    ci = max(range(n_cols), key=lambda j: counts[j])
    if counts[ci] < 2:
        return 0
    added = 0
    for ri, row in enumerate(rows):
        if ci >= len(row):
            continue
        cell = str(row[ci] or "").strip()
        codes = re.findall(r"YT\d{13}", cell)
        if len(codes) != 1:
            continue
        code = codes[0]
        idx = cell.find(code)
        suffix = cell[idx + len(code):].strip()
        if not suffix:
            continue
        pattern = r"^(?:%s)((?:[\s，,、]*(?:%s))*)$" % (
            "|".join(_STATUS_WORD), "|".join(_STATUS_WORD))
        if re.fullmatch(pattern, suffix):
            row[ci] = code
            flags.append({
                "row": ri, "col": ci, "rebuilt": code, "whole": cell, "local": suffix,
                "reason": "单号与签收状态粘连，已拆分待复核",
            })
            added += 1
            log(f"第{ri + 1}行第{ci + 1}列：单号与状态词拆分 {cell!r} -> {code}，"
                f"状态 {suffix!r} 转入待复核")
        elif any(st in suffix for st in _STATUS_WORD):
            flags.append({
                "row": ri, "col": ci, "rebuilt": code, "whole": cell, "local": suffix,
                "reason": "单号后粘连状态词，无法安全拆分，待复核",
            })
            added += 1
            log(f"第{ri + 1}行第{ci + 1}列：单号后粘连状态词 {cell!r} 未拆分，标待复核")
    return added


def _mark_uncertain_received(rows: list, flags: list, log) -> int:
    """实收列的 1 位数值：有可靠同列共识才补，否则保留原值并标待复核。

    第 3 张实收列识别出 `2`、`1`，同列常见值是 12/24，`2/1` 不是任何常见值
    的截断，无法唯一确定；这里不补（绝不臆造），只追加黄底待复核标记。
    """
    if len(rows) < 6:
        return 0
    header_idx = None
    for i in range(min(4, len(rows))):
        text = "".join(str(c or "") for c in rows[i])
        if "实收" in text:
            header_idx = i
            break
    if header_idx is None:
        return 0
    header = [str(c or "") for c in rows[header_idx]]
    ci = next((j for j, h in enumerate(header)
               if "实收" in h and not any(k in h for k in ("数量", "短少", "件数"))), None)
    if ci is None:
        return 0

    vals = [str(r[ci]).strip() for r in rows[header_idx + 1:] if ci < len(r)]
    digits = [v for v in vals if re.fullmatch(r"\d{1,3}", v)]
    if len(digits) < 5:
        return 0
    top_val, top_count = Counter(digits).most_common(1)[0]
    existing = set()
    for f in flags:
        try:
            existing.add((int(f["row"]), int(f["col"])))
        except (KeyError, TypeError, ValueError):
            continue

    added = 0
    for ri in range(header_idx + 1, len(rows)):
        if ci >= len(rows[ri]):
            continue
        value = str(rows[ri][ci]).strip()
        if not re.fullmatch(r"\d", value) or (ri, ci) in existing:
            continue
        if len(top_val) > 1 and top_val.startswith(value) and top_count / len(digits) >= 0.6:
            rows[ri][ci] = top_val
            log(f"[规则] 第{ri + 1}行第{ci + 1}列实收按列共识补全：{value} -> {top_val}")
            continue
        flags.append({
            "row": ri, "col": ci, "rebuilt": value, "whole": value, "local": "",
            "reason": "实收列 1 位值，同列共识不足，待复核",
        })
        added += 1
        log(f"[规则] 第{ri + 1}行第{ci + 1}列实收 {value!r} 无可靠同列共识，保留并标待复核")
    return added


def _mark_numeric_glue(rows: list, flags: list, log) -> int:
    """小表数字列粘连：只在列序列唯一时拆，否则整格标待复核。

    例 `103102`（前后 101/104）缺 102/103 两个值，两段都可能是本行，
    无法唯一确定，不拆；`5554`（前后 53/56）同理。只有前后只缺一个值
    （prev+1 == next-1）且该值正好在两段里时，才保留该值、另一段进待复核。
    """
    if not rows:
        return 0
    n_cols = max(len(r) for r in rows)
    if n_cols > 8:          # 只处理第 2 张这类小表，大表不动
        return 0

    added = 0
    for ci in range(n_cols):
        vals = [str(r[ci]).strip() for r in rows if ci < len(r)]
        if any("YT" in v for v in vals):
            continue
        digits = [v for v in vals if re.fullmatch(r"\d{1,8}", v)]
        if len(digits) < 5:
            continue
        mode_len = Counter(len(v) for v in digits).most_common(1)[0][0]
        for ri, row in enumerate(rows):
            if ci >= len(row):
                continue
            value = str(row[ci]).strip()
            if not re.fullmatch(r"\d{1,8}", value) or len(value) <= mode_len:
                continue
            parts = None
            for seg in (2, 3):
                if len(value) == 2 * seg:
                    parts = [value[:seg], value[seg:]]
                    break
            if not parts:
                continue
            prev = nxt = None
            for rj in range(ri - 1, -1, -1):
                x = str(rows[rj][ci]).strip() if ci < len(rows[rj]) else ""
                if re.fullmatch(r"\d{1,4}", x):
                    prev = int(x)
                    break
            for rj in range(ri + 1, len(rows)):
                x = str(rows[rj][ci]).strip() if ci < len(rows[rj]) else ""
                if re.fullmatch(r"\d{1,4}", x):
                    nxt = int(x)
                    break
            expected = None
            if prev is not None and nxt is not None and prev + 1 == nxt - 1:
                expected = str(prev + 1)
            if expected and expected in parts:
                other = parts[1] if parts[0] == expected else parts[0]
                row[ci] = expected
                flags.append({
                    "row": ri, "col": ci, "rebuilt": expected, "whole": value,
                    "local": other,
                    "reason": "数字列粘连，已按列序列拆分，另一段待复核",
                })
                log(f"第{ri + 1}行第{ci + 1}列数字粘连拆分：{value} -> {expected}，"
                    f"{other} 待复核")
            else:
                flags.append({
                    "row": ri, "col": ci, "rebuilt": "", "whole": value, "local": "",
                    "reason": "数字列粘连，无法唯一拆分，待复核",
                })
                log(f"第{ri + 1}行第{ci + 1}列数字粘连 {value!r} 无法唯一拆分，标待复核")
            added += 1
    return added


def _grid_pipeline(image_path, worker: OcrWorker, name: str, log, should_stop=None):
    """表格网格法识别。成功返回 (ExactPage, ocr_boxes, warp_path)，失败返回 None。"""
    reply = worker.call_locked({"op": "table_grid", "path": str(image_path)})
    if not reply.get("ok") or not reply.get("grid"):
        log(f"{name}：未检测到可用表格网格"
            f"（{reply.get('reason') or reply.get('error')}），改用几何重建")
        return None
    warp_path = Path(reply["warp_path"])
    try:
        hlines = [float(v) for v in reply["hlines"]]
        vlines = [float(v) for v in reply["vlines"]]
        log(f"{name}：检测到表格网格 {reply.get('n_rows')} 行 x {reply.get('n_cols')} 列，"
            f"已透视矫正（{reply.get('warp_w')}x{reply.get('warp_h')}）")
        if should_stop is not None and should_stop():
            raise EngineError("已取消转换。")

        boxes = worker.ocr(warp_path)
        try:
            calibrate_boxes(boxes, on_log=log)
        except Exception as exc:
            log(f"{name}：规则级校准跳过（{exc}）")

        # 纸张候选（倾斜实拍小表）列线常被窄数字框“吸偏”，把右列数字并进
        # 单号列；按跨线窄框的实际左边界做一次列线吸附。非纸张候选不启用，
        # 第 1/3 张既有列线逻辑完全不变。
        if reply.get("paper"):
            vlines = _snap_vlines_to_content(vlines, boxes, log)

        header_cells, body_cells = _grid_cell_grid(hlines, vlines)
        header_texts = _place_boxes(boxes, [header_cells])[0]
        body_texts = _place_boxes(boxes, body_cells)

        # 网格外推会在表格下方按延伸行距补出无内容的假行（底框/阴影有少量
        # 墨迹就继续外推）。这些行本身没有数据，留着会让后面“日期列按多数派
        # 补空”把日期填进空行（实测 0683 多出 2 行假数据），也会白耗补识预算；
        # 这里只裁尾部、只裁全空行，中间漏识别的行不动。row_y 要同步裁掉同数
        # 的尾带，否则 realign 后按“尾部切片”取行带会整体上移两行（xcheck 把
        # 表头日期格对到第一数据行）。
        trim_tail = 0
        while body_texts and not any(str(c).strip() for c in body_texts[-1]):
            body_texts.pop()
            body_cells.pop()
            trim_tail += 1

        # 网格过空的前置判定：切格后如果非空率已经极低（例如 2/44），
        # 说明这张纸的网格检测不可用，再花预算做几十次空格放大补识也救不回来；
        # 直接回退几何重建，避免“补识白跑 + 仍然回退”的双管线耗时。
        # 判据取“表头重定位后”的行集合：网格常把标题区/空白区切成大量假行，
        # 原始切格非空率会被这些假行摊薄；而真正决定成败的是重定位后剩下的
        # 数据行（6320eb7a 原始网格 25x11，重定位后只剩 4 行 / 44 格，非空 2）。
        # 阈值必须低于补识后的 GRID_MIN_FILL_RATIO：补识本来就能把一部分空格
        # 填起来，只有“低到补满也救不回”的网格才值得提前放弃。
        probe_rows = [header_texts] + body_texts
        aligned_rows = _realign_header([list(r) for r in probe_rows])
        probe_total = sum(1 for r in aligned_rows for c in r if c is not None)
        probe_nonempty = sum(1 for r in aligned_rows for c in r if str(c).strip())
        if probe_total and probe_nonempty < GRID_SKIP_FILL_RATIO * probe_total:
            log(f"{name}：网格切格后过空（{probe_nonempty}/{probe_total} 非空），"
                f"跳过空格补识，直接回退几何重建")
            return None

        budget = [GRID_FILL_BUDGET]
        _fill_empty_cells([header_cells], [header_texts], worker, warp_path, log, budget)
        _fill_empty_cells(body_cells, body_texts, worker, warp_path, log, budget)

        # 表头重定位：warp 顶部 pad 带入的标题区横线是假线（与真实行高等距时
        # _keep_dense_run 剔不掉），表头行被抬成空假行；det 框也可能把表头字
        # 挤过格线掉到第一数据行。赶在列级增强之前做：date_cols/code_cols 靠
        # 表头文字定位，用空假表头判断时日期列增强永远不触发，首行日期残缺
        # 就没人救。重定位裁掉的是表头上方的假行，body_cells 尾部与保留的
        # 数据行逐行对应（裁剪只发生在行首，几何对齐取尾部切片）。
        try:
            rows = _realign_header([header_texts] + body_texts)
        except Exception as exc:
            log(f"{name}：表头重定位跳过（{exc}）")
            rows = [header_texts] + body_texts
        try:
            sidx_h, _pidx_h = _locate_serial_pkg_cols(rows)
            header_rebuilt = _apply_handoff_header(rows, sidx_h)
            if header_rebuilt:
                log(f"{name}：表头按交接清单 16 列还原（序号列 {sidx_h}）")
        except Exception as exc:
            header_rebuilt = False
            log(f"{name}：表头还原跳过（{exc}）")
        header_texts = rows[0]
        body_texts = rows[1:]
        if len(body_cells) > len(body_texts):
            body_cells = body_cells[len(body_cells) - len(body_texts):]

        # 列级增强（按表头定位列，只对命中列生效）：
        #   运单号列 -> 多尺度投票纠错；日期列 -> 整列放大识别补残缺值。
        try:
            # 列级增强整体暂不接入主流程（保持与历史结果一致），原因：
            #
            # 1) `_enhance_code_column` 依赖的 `xcheck._prefix_compatible` 并不存在，
            #    导入即抛 ImportError，被下面的 except 吞掉 —— 于是**整个 try 块
            #    从来没执行过**，日期列增强也被一起跳过了。
            # 2) 修好导入后直接启用会改写落格值（条件：位数相同 + 字母部分一致），
            #    该逻辑从未在真实数据上验证过；
            # 3) 启用 `_enhance_date_column` 会把「日期+承运商+单号」这种跨列长框
            #    整格覆盖成纯日期（实测 be680b97 A2/A4-A8 因此丢 6 条单号）；
            # 4) `_fill_code_column_full` 会把表头/相邻列内容填进空白格。
            #
            # 全量 45 张对照：启用后 4 张图丢 16 条真实单号（be680b97、fc03c6d3、
            # 815bfb1c、890f24fd），净收益为负。三者都要先解决「格子内容归属」
            # 才能重新启用，见各自 docstring。
            pass
        except Exception as exc:
            if _is_memory_error(exc):
                raise
            log(f"{name}：列级增强跳过（{exc}）")

        rows = [header_texts] + body_texts
        total = sum(1 for r in rows for c in r if c is not None)
        nonempty = sum(1 for r in rows for c in r if str(c).strip())
        if len(rows) < 2 or len(rows[0]) < 2:
            log(f"{name}：网格结果行列不足（{len(rows)}x{len(rows[0])}），回退几何重建")
            return None
        if nonempty < GRID_MIN_FILL_RATIO * total:
            log(f"{name}：网格结果过空（{nonempty}/{total} 非空），回退几何重建")
            return None

        row_y_all = [(a, b) for a, b in zip(hlines, hlines[1:])]
        if trim_tail:
            row_y_all = row_y_all[:-trim_tail]
        block = TableBlock(
            kind="table",
            rows=rows,
            bbox=(int(vlines[0]), int(hlines[0]), int(vlines[-1]), int(hlines[-1])),
            col_x=[(c, d) for c, d in zip(vlines, vlines[1:])],
            row_y=row_y_all,
        )
        block.header_rebuilt = header_rebuilt

        # 标题行剥离 + 单号列清理：只动表格顶部/单号列，不改其他数据列。
        rows, row_y_used, title_lines = _trim_title_rows(rows, block.row_y)
        rows, code_flags = _normalize_code_rows(rows, log)
        # 宽表序号列：拆“序号+包裹号”粘连，按连续序列补齐缺号。
        rows, serial_flags = _fix_wide_table_serials(rows, log)
        code_flags += serial_flags
        block.rows = rows
        block.row_y = row_y_used

        # 网格外的文字（表格上方的标题、下方的脚注/签名）不能丢：
        # warp 区域比表格大一圈，这些框会落在网格之外，收集为文字块按位置插入。
        cells_flat = [c for row in body_cells for c in row] + list(header_cells)

        def _outside(b):
            cx = (b["bbox"][0] + b["bbox"][2]) / 2.0
            cy = (b["bbox"][1] + b["bbox"][3]) / 2.0
            return not any(x1 <= cx <= x2 and y1 <= cy <= y2 for x1, y1, x2, y2 in cells_flat)

        stray = [b for b in boxes if _outside(b)]
        above = [b for b in stray if (b["bbox"][1] + b["bbox"][3]) / 2.0 < hlines[0]]
        below = [b for b in stray if (b["bbox"][1] + b["bbox"][3]) / 2.0 > hlines[-1]]
        blocks: list = []
        if above:
            blocks.append(TextBlock(kind="text", lines=_text_lines(above),
                                    bbox=(0, 0, 0, 0)))
        if title_lines:
            # 从表格顶部剥离的标题不丢，按文字块写回 Excel 上方
            blocks.append(TextBlock(kind="text", lines=title_lines,
                                    bbox=(0, 0, 0, 0)))
        blocks.append(block)
        if below:
            blocks.append(TextBlock(kind="text", lines=_text_lines(below),
                                    bbox=(0, 0, 0, 0)))
        page = ExactPage(name=name, blocks=blocks, warnings=[])
        page.flags = code_flags
        page.image_path = str(image_path)
        page.to_original = reply.get("to_original")
        page.ocr_boxes = boxes
        q = reply.get("quality") or {}
        page.quality_meta = {
            "paper": bool(reply.get("paper")),
            "paper_area_ratio": reply.get("paper_area_ratio"),
            "paper_angle": reply.get("paper_angle"),
            "grid_rows": reply.get("n_rows"),
            "grid_cols": reply.get("n_cols"),
            "hlines": len(reply.get("hlines") or []),
            "blur_var": q.get("blur_var"),
            "highlight_ratio": q.get("highlight_ratio"),
        }
        return page, boxes, warp_path
    except Exception as exc:
        if _is_memory_error(exc):
            # 内存不足不是这张图的问题，回退几何重建也救不回来，且会继续吃内存；
            # 直接抛给上层，由批量/Web 侧重启 worker 后重试本页。
            log(f"{name}：内存不足（{exc}），交由上层重启识别进程")
            try:
                warp_path.unlink(missing_ok=True)
            except OSError:
                pass
            raise EngineError(f"内存不足，请重启识别进程后重试：{exc}") from exc
        log(f"{name}：网格管线失败（{exc}），回退几何重建")
        try:
            warp_path.unlink(missing_ok=True)
        except OSError:
            pass
        return None


def _run_local_pipeline(image_path, worker, name, log, do_fill: bool = True) -> dict:
    """本地 OCR 管线：整图 OCR -> 规则级校准 -> 几何重建（-> 序号断档补行）。

    序号补行成本高（逐断档 clip 放大重识别），由调用方按需触发：
    纯 OCR 管线中它紧接几何重建执行（缺失的序号行需要在写入 Excel 前补齐）。
    返回结果字典；失败时 error 字段非空，其余字段为空。
    """
    out = {
        "local_page": None, "local_rows": 0, "local_table_rows": [],
        "ocr_texts": [], "ocr_boxes": [], "local_tables": [], "error": "",
    }
    try:
        boxes = worker.ocr(image_path)
        quality = dict(getattr(worker, "last_quality", {}) or {})
        out["ocr_boxes"] = boxes
        log(f"{name}：本地识别到 {len(boxes)} 个文本框")
        if boxes:
            try:
                calibrate_boxes(boxes, on_log=log)
            except Exception as exc:
                log(f"{name}：规则级校准跳过（{exc}）")
            out["ocr_texts"] = [(str(b.get("text", "")), float(b.get("score", 0) or 0))
                                for b in boxes if float(b.get("score", 0) or 0) >= 0.85]
            res = rebuild(boxes, on_log=log)
            if do_fill:
                _fill_local_tables(worker, image_path, res.blocks, name, log)
            out["local_page"] = ExactPage(name=name, blocks=res.blocks, warnings=res.warnings)
            out["local_page"].quality_meta = quality
            for block in res.blocks:
                if getattr(block, "kind", "") != "table":
                    continue
                out["local_tables"].append(block)
                for row in getattr(block, "rows", []):
                    out["local_rows"] += 1
                    if row:
                        out["local_table_rows"].append([str(c).strip() for c in row])
    except Exception as exc:
        if _is_memory_error(exc):
            log(f"{name}：本地识别内存不足（{exc}），交由上层重启识别进程")
            raise EngineError(f"内存不足，请重启识别进程后重试：{exc}") from exc
        out["error"] = str(exc)
        log(f"{name}：本地识别失败（{exc}）")
    return out


def _cloud_text_pass(page, worker, image_path, flagged: list, log) -> None:
    """可选的云端文本兜底（默认关闭，见 nv_fix.py）。

    候选池 = 所有含中日韩字符的文本格；nv_fix 再用受控词表筛出"可疑"的
    （不在词表里的分段，如把 芬达 读成 开达），并把交叉校验标记过的格一并纳入。
    本地整图 OCR 与几何重建同源会犯同样的错，两方一致就不进待复核清单，
    所以这里必须主动看全部文本格，而不是只盯待复核格。
    被云端高置信度纠正的格子会从 flagged 中移除——它已经不需要人工复核了。
    """
    if not nv_fix.is_enabled():
        return
    tables = [b for b in (getattr(page, "blocks", []) or [])
              if getattr(b, "kind", "") == "table" and getattr(b, "rows", None)]
    if not tables:
        return
    flagged_keys = {(f.get("row"), f.get("col")) for f in flagged}
    items, owner = [], {}
    for bi, block in enumerate(tables):
        rows = block.rows
        col_x = list(getattr(block, "col_x", []) or [])
        row_y = list(getattr(block, "row_y", []) or [])
        for ri, row in enumerate(rows):
            for ci, cell in enumerate(row):
                value = str(cell)
                if not nv_fix.is_text_cell(value):
                    continue
                if ci >= len(col_x) or ri >= len(row_y):
                    continue
                rect = (float(col_x[ci][0]), float(row_y[ri][0]),
                        float(col_x[ci][1]), float(row_y[ri][1]))
                items.append({"row": ri, "col": ci, "rect": rect, "current": value,
                              "suspect": (ri, ci) in flagged_keys})
                owner[(ri, ci)] = bi
    if not items:
        return
    try:
        fixes = nv_fix.fix_cells(items, worker, image_path, log=log)
    except Exception as exc:
        log(f"云端纠错跳过：{exc}")
        return
    for ri, ci, text, _conf in fixes:
        block = tables[owner.get((ri, ci), 0)]
        if 0 <= ri < len(block.rows) and 0 <= ci < len(block.rows[ri]):
            block.rows[ri][ci] = text
    flagged[:] = [f for f in flagged
                  if (f.get("row"), f.get("col")) not in
                  {(r, c) for r, c, _t, _s in fixes}]


def _table_code_set(page) -> set[str]:
    """页面表格里的唯一 YT 单号集合（只读）。"""
    codes = set()
    for block in getattr(page, "blocks", []) or []:
        if getattr(block, "kind", "") != "table":
            continue
        for row in getattr(block, "rows", None) or []:
            for cell in row:
                codes.update(re.findall(r"YT\d{13}", str(cell)))
    return codes


def _supplement_grid_codes(page, image_path, worker, name, log, warnings) -> dict:
    """只补单号：网格主表不动，仅把原图 OCR 缺码写进单号列的空单元格。

    坐标：用“已有单号”的 raw y 与网格行中心拟合线性映射；残差可信才按 y
    定位。不可信时退化为按行序补到末尾空行，并标低置信。已有非空单号格
    一律不改写；已有 1 个不同码/多个码的行不补，记为未定位。
    """
    tables = [b for b in getattr(page, "blocks", []) or []
              if getattr(b, "kind", "") == "table"]
    if not tables:
        return {"added": 0, "unlocated": 0}
    main = max(tables, key=lambda b: (sum(len(set(re.findall(r"YT\d{13}", str(cell))))
                                       for r in (getattr(b, "rows", None) or [])
                                       for cell in r),
                                       len(getattr(b, "rows", None) or [])))
    rows = getattr(main, "rows", None) or []
    if not rows:
        return {"added": 0, "unlocated": 0}
    col_counts = {}
    for r in rows:
        for ci, cell in enumerate(r):
            n = len(set(re.findall(r"YT\d{13}", str(cell))))
            if n:
                col_counts[ci] = col_counts.get(ci, 0) + n
    if not col_counts:
        return {"added": 0, "unlocated": 0}
    code_ci = max(col_counts, key=col_counts.get)
    row_y = getattr(main, "row_y", None) or []
    if len(row_y) > len(rows):
        # _realign_header 只从行首裁掉假行；保留的数据行对应 row_y 尾部切片。
        row_y = row_y[-len(rows):]
    elif len(row_y) < len(rows):
        row_y = []
    page_codes = {c for r in rows for cell in r
                  for c in re.findall(r"YT\d{13}", str(cell))}
    raw_boxes = worker.ocr(str(image_path))
    raw_occ = {}
    for box in raw_boxes:
        for code in re.findall(r"YT\d{13}", str(box.get("text", ""))):
            y = (float(box["bbox"][1]) + float(box["bbox"][3])) / 2.0
            raw_occ.setdefault(code, y)
    raw_codes = set(raw_occ)
    log(f"{name}：单号补录检查：原图 OCR {len(raw_codes)}，页面 {len(page_codes)}，"
        f"缺 {len(raw_codes - page_codes)}")
    if len(raw_codes) < 10 or len(raw_codes) <= len(page_codes):
        return {"added": 0, "unlocated": 0}
    missing = sorted(raw_codes - page_codes, key=lambda c: raw_occ[c])
    # 已有单号作锚点，拟合 raw_y -> page_y
    anchors = []
    if row_y:
        for ri, r in enumerate(rows):
            if code_ci >= len(r):
                continue
            for code in re.findall(r"YT\d{13}", str(r[code_ci])):
                if code in raw_occ:
                    anchors.append((raw_occ[code], (row_y[ri][0] + row_y[ri][1]) / 2.0))
    slope = None
    if len(anchors) >= 3:
        mx = sum(x for x, _ in anchors) / len(anchors)
        my = sum(y for _, y in anchors) / len(anchors)
        den = sum((x - mx) ** 2 for x, _ in anchors)
        if den > 0:
            slope = sum((x - mx) * (y - my) for x, y in anchors) / den
            inter = my - slope * mx
            rms = (sum((y - (slope * x + inter)) ** 2 for x, y in anchors)
                   / len(anchors)) ** 0.5
            heights = [b - a for a, b in row_y]
            heights.sort()
            row_h = heights[len(heights) // 2] if heights else 1.0
            trusted = rms <= 0.75 * row_h
        else:
            inter = 0.0; rms = 1e9; trusted = False; row_h = 1.0
    else:
        inter = 0.0; rms = 1e9; trusted = False; row_h = 1.0
    log(f"{name}：单号补录映射：锚点 {len(anchors)}，RMS "
        f"{'∞' if rms >= 1e9 else round(rms, 2)}，可信={trusted}，行高={row_h:.1f}")

    added, unlocated = [], []
    occupied = set()
    if trusted:
        for code in missing:
            py = slope * raw_occ[code] + inter
            hit = [ri for ri, ry in enumerate(row_y) if ry[0] <= py <= ry[1]]
            if not hit:
                hit = [min(range(len(row_y)),
                           key=lambda i: abs((row_y[i][0] + row_y[i][1]) / 2.0 - py))]
                if abs((row_y[hit[0]][0] + row_y[hit[0]][1]) / 2.0 - py) > 0.8 * row_h:
                    unlocated.append({"code": code, "y": raw_occ[code],
                                      "reason": "y 映射无匹配行带"})
                    continue
            ri = min(hit, key=lambda i: abs((row_y[i][0] + row_y[i][1]) / 2.0 - py))
            cell = str(rows[ri][code_ci]).strip() if code_ci < len(rows[ri]) else ""
            if cell:
                unlocated.append({"code": code, "y": raw_occ[code],
                                  "row": ri + 1,
                                  "reason": "目标行单号格非空，不改写"})
                continue
            if (ri, code_ci) in occupied:
                unlocated.append({"code": code, "y": raw_occ[code],
                                  "row": ri + 1, "reason": "同行已补入其他码"})
                continue
            rows[ri][code_ci] = code
            occupied.add((ri, code_ci))
            added.append({"code": code, "y": raw_occ[code], "row": ri + 1,
                          "page_y": round(py, 1)})
            log(f"{name}：单号补录 第{ri+1}行 {code}（原图 y={raw_occ[code]:.1f}）")
    else:
        empties = [(ri, code_ci) for ri in range(len(rows))
                   if code_ci < len(rows[ri]) and not str(rows[ri][code_ci]).strip()]
        for (ri, ci), code in zip(empties[-len(missing):], missing):
            rows[ri][ci] = code
            added.append({"code": code, "y": raw_occ[code], "row": ri + 1,
                          "page_y": None, "low_confidence": True})
            log(f"{name}：单号低置信补录 第{ri+1}行 {code}（原图 y={raw_occ[code]:.1f}，"
                f"无可信坐标映射）")
        if len(added) < len(missing):
            unlocated.extend({"code": c, "y": raw_occ[c], "reason": "低置信补录无空行"}
                             for c in missing[len(added):])
    if added or unlocated:
        for u in unlocated:
            log(f"{name}：单号未定位 {u['code']}（原图 y={u['y']}，"
                f"{u.get('row', '无行')}，{u['reason']}）")
        warnings.append(f"单号补录：补 {len(added)}，未定位 {len(unlocated)}")
    return {"added": added, "unlocated": unlocated,
            "trusted": trusted, "rms": round(rms, 2) if rms < 1e9 else None,
            "anchors": len(anchors), "row_h": round(row_h, 1)}


def _inv3x3(m):
    """3x3 单应矩阵求逆（伴随矩阵法）。"""
    a, b, c = m[0]
    d, e, f = m[1]
    g, h, i = m[2]
    det = a * (e * i - f * h) - b * (d * i - f * g) + c * (d * h - e * g)
    if abs(det) < 1e-12:
        return None
    return [[(e * i - f * h) / det, (c * h - b * i) / det, (b * f - c * e) / det],
            [(f * g - d * i) / det, (a * i - c * g) / det, (c * d - a * f) / det],
            [(d * h - e * g) / det, (b * g - a * h) / det, (a * e - b * d) / det]]


def _apply_h(m, x, y):
    w = m[2][0] * x + m[2][1] * y + m[2][2]
    if abs(w) < 1e-12:
        return None
    return ((m[0][0] * x + m[0][1] * y + m[0][2]) / w,
            (m[1][0] * x + m[1][1] * y + m[1][2]) / w)


def _code_suffix_score(cell, code_d):
    """格子数字段与码数字部分的最长公共后缀长度（同源性强度的量化）。"""
    best = 0
    for seg in re.findall(r"\d+", str(cell)):
        n = 0
        for x, y in zip(reversed(seg), reversed(code_d)):
            if x != y:
                break
            n += 1
        if n > best:
            best = n
    return best


def _cell_matches_code(cell, code_d):
    """格子与码同源：某数字段(>=4位)是码后缀，或公共后缀 >=6 位。

    覆盖网格 OCR 丢 YT 前缀/首位的残缺形态（`5030 签收`、`552868031007`），
    同时挡住粘入无关数字的格子（段整体不是码后缀且公共后缀不足）。
    """
    for seg in re.findall(r"\d+", str(cell)):
        if len(seg) >= 4 and code_d.endswith(seg):
            return True
    return _code_suffix_score(cell, code_d) >= 6


def _zoom_repair_codes(page, image_path, worker, name, log,
                       warnings, force: bool = False) -> dict:
    """局部放大修复网格漏读单号。

    与被关闭的 _supplement_grid_codes 的本质区别：
      - 行定位用 page.to_original 的逆单应把原图框投到 warp 行带（几何精确），
        不做线性拟合猜行；
      - 空格不直接写入：±1 行有同源残缺格时改升级那行（格子尾号自证优先于
        单应落点；v1.4.4 实测 2 条纯空格写入全是单应偏半行造成的错行）；
      - 每个候选必须局部放大（检测框外扩 15%、2x）双读一致（13 位数字全同）
        才写入；异码格/双码格/行带外 >0.5 行/双读不一致一律只记待复核；
      - 写入前全表查重，同码已存在则不写。
    成本：一次原图整图 OCR（候选来源；warp 图的 ocr_boxes 因矫正后小字变形
    完整码更少，不可用）+ 每候选一次 clip 放大（约 20-30ms）。
    """
    empty = {"added": 0, "upgraded": 0, "review": 0}
    tables = [b for b in getattr(page, "blocks", []) or []
              if getattr(b, "kind", "") == "table"]
    if not tables:
        return empty
    main = max(tables, key=lambda b: (sum(len(set(re.findall(r"YT\d{13}", str(c))))
                                      for r in (getattr(b, "rows", None) or [])
                                      for c in r),
                                      len(getattr(b, "rows", None) or [])))
    rows = getattr(main, "rows", None) or []
    if not rows:
        return empty
    col_counts = {}
    for r in rows:
        for ci, cell in enumerate(r):
            n = len(set(re.findall(r"YT\d{13}", str(cell))))
            if n:
                col_counts[ci] = col_counts.get(ci, 0) + n
    if not col_counts:
        return empty
    code_ci = max(col_counts, key=col_counts.get)
    row_y = getattr(main, "row_y", None) or []
    if len(row_y) > len(rows):
        # _realign_header 只从行首裁假行；数据行对应 row_y 尾部切片
        row_y = row_y[-len(rows):]
    elif len(row_y) < len(rows):
        row_y = []
    to_original = getattr(page, "to_original", None)
    if not row_y or not to_original:
        return empty
    Hi = _inv3x3(to_original)
    if Hi is None:
        return empty
    grid_codes = {c for r in rows for cell in r
                  for c in re.findall(r"YT\d{13}", str(cell))}
    # 前置早退（避免每张图都白付一次原图整图 OCR）：
    # 单号密度 >=0.5 说明网格没怎么漏读；<3 个完整码说明不是运单表。
    # force：行带二次切行命中的图（669d 型）——切行后密度达标但单号列
    # 仍有大量 T 开头残缺格，必须跟进修复。
    if not force and (len(grid_codes) < 3 or len(grid_codes) >= 0.5 * len(rows)):
        return empty
    # 原图整图 OCR：warp 图经透视矫正后小字变形，完整码明显更少，必须用原图
    raw_boxes = worker.ocr(str(image_path))
    raw_full = {}
    for box in raw_boxes or []:
        for code in re.findall(r"YT\d{13}", str(box.get("text", ""))):
            raw_full.setdefault(code, box.get("bbox"))
    if len(raw_full) <= len(grid_codes):
        return empty
    row_h = (row_y[1][0] - row_y[0][0]) if len(row_y) > 1 else 40.0
    if row_h <= 0:
        row_h = 40.0
    added, upgraded, review = 0, 0, 0
    new_codes = set(grid_codes)
    cands = [c for c in raw_full if c not in grid_codes]
    log(f"{name}：单号修复候选 {len(cands)}："
        f"{' '.join(c[-6:] for c in sorted(cands))}")
    for code in sorted(raw_full, key=lambda c: raw_full[c][1] if raw_full[c] else 0):
        if code in grid_codes:
            continue
        bb = raw_full[code]
        if not bb or len(bb) < 4:
            continue
        pt = _apply_h(Hi, (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0)
        if pt is None:
            continue
        wy = pt[1]
        hit = [i for i, (a, b) in enumerate(row_y) if a <= wy <= b]
        if not hit:
            i0 = min(range(len(row_y)),
                     key=lambda j: abs((row_y[j][0] + row_y[j][1]) / 2 - wy))
            if abs((row_y[i0][0] + row_y[i0][1]) / 2 - wy) > 0.5 * row_h:
                review += 1
                log(f"{name}：单号修复待复核 {code}（行带外>0.5行，wy={wy:.0f}）")
                continue
            hit = [i0]
        ri = hit[0]
        cell = str(rows[ri][code_ci]) if code_ci < len(rows[ri]) else ""
        cell_codes = re.findall(r"YT\d{13}", cell)
        if code in cell_codes:
            continue
        code_d = re.sub(r"\D", "", code)
        if cell.strip() and not cell_codes:
            if not _cell_matches_code(cell, code_d):
                # 格子不同源：单应 wy 与 warp 框 y 有半框高偏差（约 17px），
                # 切行后段高仅 ~47px 时易跨段；查 ±1 行同源格再判。
                best_j, best_s = None, 0
                for j in (ri - 1, ri + 1):
                    if 0 <= j < len(rows) and code_ci < len(rows[j]):
                        cj = str(rows[j][code_ci])
                        if (cj.strip() and not re.findall(r"YT\d{13}", cj)
                                and _cell_matches_code(cj, code_d)):
                            s = _code_suffix_score(cj, code_d)
                            if s > best_s:
                                best_j, best_s = j, s
                if best_j is None:
                    review += 1
                    log(f"{name}：单号修复待复核 {code}（行{ri+1}格内容不同源："
                        f"「{cell[:24]}」）")
                    continue
                ri = best_j
                cell = str(rows[ri][code_ci])
            # 同源 -> 升级本行
        elif not cell.strip():
            # 空格：先查 ±1 行有无同源残缺格，有则升级那行（格子自证优先）
            best_j, best_s = None, 0
            for j in (ri - 1, ri + 1):
                if 0 <= j < len(rows) and code_ci < len(rows[j]):
                    cj = str(rows[j][code_ci])
                    if (cj.strip() and not re.findall(r"YT\d{13}", cj)
                            and _cell_matches_code(cj, code_d)):
                        s = _code_suffix_score(cj, code_d)
                        if s > best_s:
                            best_j, best_s = j, s
            if best_j is None:
                review += 1   # 空格且无邻行自证：不写，防单应错行
                log(f"{name}：单号修复待复核 {code}（行{ri+1}空格且邻行无同源格）")
                continue
            ri = best_j
            cell = str(rows[ri][code_ci])
        else:
            review += 1   # 格内已有异码（含双码格）
            log(f"{name}：单号修复待复核 {code}（行{ri+1}格内已有异码："
                f"「{cell[:24]}」）")
            continue
        # 局部放大双读确认：13 位数字全同才算一致（Y/1/2 前缀误读放行）
        w, h = bb[2] - bb[0], bb[3] - bb[1]
        clip = [bb[0] - w * 0.15, bb[1] - h * 0.15,
                bb[2] + w * 0.15, bb[3] + h * 0.15]
        # 2.0 不一致才补 3.0：be68 实测 2.0/2.5 对粘连小字读错、3.0 读对；
        # 与整图 13 位完全一致才确认，补读不放松判据。
        zdigits = set()
        ok_zoom = False
        for scale in (2.0, 3.0):
            try:
                zboxes = worker.ocr(str(image_path), clip=clip, scale=scale)
            except Exception:
                continue
            zdigits = {re.sub(r"\D", "", str(b.get("text", ""))) for b in zboxes}
            if any(code_d in d for d in zdigits if len(d) >= 13):
                ok_zoom = True
                break
        if not ok_zoom:
            review += 1
            log(f"{name}：单号修复待复核 {code}（行{ri+1}放大双读不一致："
                f"「{sorted(zdigits)[:2]}」）")
            continue
        if code in new_codes:
            review += 1   # 全表查重
            log(f"{name}：单号修复待复核 {code}（全表已有该码）")
            continue
        # 写入：把格子中与码公共后缀最长的数字段替换为完整码，保留其余文本
        best_span, best_s = None, 0
        for m in re.finditer(r"\d+", cell):
            s = 0
            for x, y in zip(reversed(m.group(0)), reversed(code_d)):
                if x != y:
                    break
                s += 1
            if len(m.group(0)) >= 4 and code_d.endswith(m.group(0)):
                s = max(s, len(m.group(0)))
            if s > best_s:
                best_span, best_s = m.span(), s
        if best_span is None or best_s < 4:
            review += 1
            continue
        rows[ri][code_ci] = cell[:best_span[0]] + code + cell[best_span[1]:]
        new_codes.add(code)
        if cell.strip():
            upgraded += 1
            log(f"{name}：单号修复 第{ri+1}行 {code}（残缺格「{cell[:24]}」"
                f"经放大双读确认）")
        else:
            added += 1
            log(f"{name}：单号补全 第{ri+1}行 {code}（放大双读确认）")
    if added or upgraded or review:
        log(f"{name}：单号局部放大修复：补全 {added}，修复残缺 {upgraded}，"
            f"待复核 {review}")
        warnings.append(f"单号修复：补全 {added}，修复残缺 {upgraded}，"
                        f"待复核 {review}")
    return {"added": added, "upgraded": upgraded, "review": review}


def _rebuild_grid_rows_from_codes(ocr, col_x, code_ci, code_ys, table_top):
    """按单号码框重建整表行（0683 型弯纸：网格线在弯折处漏检）。

    网格线漏检时，一个行带会并进 2 个物理行；只切行带仍会把序号/日期
    错配到相邻行（这些列的文字框中心比单号框高约 20-28px）。这里改用
    单号框 y 作为行锚点，并为每列估计统一的垂直偏移后重新落格：

      - 某列框数恰好等于单号框数时，按 y 顺序与锚点一一配对，取中位差
        作为该列偏移（序号列约 -25px、日期列约 -22px）；
      - 其余列按原始中心就近落行。

    返回 (rows, row_y)；锚点不足或间距不稳定时返回 None，交回原有流程。
    """
    n = len(code_ys)
    if n < 3:
        return None
    gaps = [code_ys[i + 1] - code_ys[i] for i in range(n - 1)]
    pitch = sorted(gaps)[len(gaps) // 2]
    if pitch <= 0 or max(gaps) > 1.5 * pitch:
        return None
    per_col: list[list] = [[] for _ in col_x]
    for box in ocr:
        bb = box.get("bbox")
        text = str(box.get("text", "")).strip()
        if not bb or len(bb) < 4 or not text:
            continue
        cx, cy = (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0
        for ci, (x0, x1) in enumerate(col_x):
            if x0 <= cx <= x1:
                per_col[ci].append((cy, cx, text))
                break
    offsets = [0.0] * len(col_x)
    for ci, items in enumerate(per_col):
        if ci == code_ci or len(items) != n:
            continue
        items.sort(key=lambda t: t[0])
        diffs = sorted(items[i][0] - code_ys[i] for i in range(n))
        offsets[ci] = diffs[n // 2]
    top = code_ys[0] - pitch / 2.0
    bottom = code_ys[-1] + pitch / 2.0
    bounds = [top] + [(code_ys[i] + code_ys[i + 1]) / 2.0
                      for i in range(n - 1)] + [bottom]
    header: list = [[] for _ in col_x]
    data: list = [[[] for _ in col_x] for _ in range(n)]
    for ci, items in enumerate(per_col):
        for cy, cx, text in items:
            yy = cy - offsets[ci]
            if yy < top:
                if cy >= table_top - pitch:
                    header[ci].append(text)
                continue
            if yy > bottom:
                continue
            si = 0 if yy <= bounds[1] else min(n - 1, int((yy - top) // pitch))
            si = max(0, min(n - 1, si))
            data[si][ci].append(text)
    new_rows = []
    if any(header):
        new_rows.append([" ".join(part for part in col if part).strip()
                         for col in header])
    for row in data:
        new_rows.append([" ".join(part for part in col if part).strip()
                         for col in row])
    new_row_y = []
    y_start = table_top if any(header) else top
    if any(header):
        new_row_y.append((y_start, top))
    for i in range(n):
        new_row_y.append((bounds[i], bounds[i + 1]))
    return new_rows, new_row_y


def _split_dense_bands(page, name, log) -> int:
    """行带内挤多个物理行的二次切行（669d 型：296px 行带内堆叠 4-7 个物理行）。

    触发（刻意收窄，v1.4.4 实测 45 张仅 669d 命中）：
      行带内、单号列 x 区间内、含 >=8 位数字的框 >=2 个，
      且带高 > 2.5x 这些框的中位 y 间距。
    切法：段界用 warp 码框自身的 y（相邻框 y 中点），与内容框同源——若改用
    原图码单应逆映射的 wy，两套 y 差半个框高会把内容框错归到相邻段；
    带内全部 warp 框按 y 归段、按 col_x 分列聚成新行（各列跟着各自段走）。
    返回新增行数（未触发返回 0）。
    """
    tables = [b for b in getattr(page, "blocks", []) or []
              if getattr(b, "kind", "") == "table"]
    if not tables:
        return 0
    main = max(tables, key=lambda b: len(getattr(b, "rows", None) or []))
    rows = getattr(main, "rows", None) or []
    row_y_full = list(getattr(main, "row_y", None) or [])
    col_x = list(getattr(main, "col_x", None) or [])
    if not rows or not row_y_full or not col_x:
        return 0
    row_y = row_y_full[-len(rows):] if len(row_y_full) > len(rows) else row_y_full
    if len(row_y) != len(rows):
        return 0
    # 单号列：完整码最多的列
    col_counts = {}
    for r in rows:
        for ci, cell in enumerate(r):
            n = len(set(re.findall(r"YT\d{13}", str(cell))))
            if n:
                col_counts[ci] = col_counts.get(ci, 0) + n
    if not col_counts:
        return 0
    code_ci = max(col_counts, key=col_counts.get)
    if code_ci >= len(col_x):
        return 0
    cx1, cx2 = col_x[code_ci]
    ocr = list(getattr(page, "ocr_boxes", None) or [])
    if not ocr:
        return 0
    # 1.5) 弯纸整表重建（0683 型）：网格行数比“表头 + 单号数”少 2 行以上，
    #      说明弯折处成段水平线漏检，行带切分无法修正序号/日期列的垂直偏移。
    #      仅当单号集中在唯一列、码框 y 间距稳定时才接管，其余仍走下面的切行。
    code_ys = []
    other_code_cols = set()
    for box in ocr:
        bb = box.get("bbox")
        text = str(box.get("text", "")).strip()
        if not bb or len(bb) < 4 or not re.search(r"YT\d{13}", text):
            continue
        cx, cy = (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0
        for ci, (x0, x1) in enumerate(col_x):
            if x0 <= cx <= x1:
                if ci == code_ci:
                    code_ys.append(cy)
                else:
                    other_code_cols.add(ci)
                break
    code_ys.sort()
    if (len(code_ys) >= 6 and not other_code_cols
            and len(code_ys) + 1 - len(rows) >= 2):
        rebuilt = _rebuild_grid_rows_from_codes(
            ocr, col_x, code_ci, code_ys, row_y[0][0])
        if rebuilt is not None:
            new_rows, new_row_y = rebuilt
            main.rows = new_rows
            main.row_y = new_row_y
            log(f"{name}：弯纸行锚重建 {len(rows)} 行 -> {len(new_rows)} 行"
                f"（单号锚点 {len(code_ys)} 个）")
            return len(new_rows) - len(rows)
    # 1) 判定哪些行带触发（全部用 warp 坐标，零额外 OCR）
    trigger = {}
    for ri, (a, b) in enumerate(row_y):
        h = b - a
        frames = []
        code_cols_hit = set()
        for box in ocr:
            bb = box.get("bbox")
            if not bb or len(bb) < 4:
                continue
            cx, cy = (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0
            if not (a <= cy <= b):
                continue
            if len(re.sub(r"\D", "", str(box.get("text", "")))) < 8:
                continue
            is_code = bool(re.match(r"^(?:YT|T)\d{10,}$",
                                    str(box.get("text", "")).strip()))
            for ci, (x1, x2) in enumerate(col_x):
                if x1 <= cx <= x2:
                    if is_code:
                        code_cols_hit.add(ci)
                    if ci == code_ci:
                        frames.append(cy)
                    break
        if len(code_cols_hit) > 1:
            # 多栏：单号形态框分布在多列（7fc3 左栏+中栏），按 y 切会把
            # 同逻辑行的两栏拆到不同段 -> 不切（列方向是多栏识别的事）
            continue
        if len(frames) >= 2:
            frames.sort()
            gaps = [frames[i + 1] - frames[i] for i in range(len(frames) - 1)]
            med_gap = sorted(gaps)[len(gaps) // 2]
            # 带高绝对阈值 + 多栏拦截：正常行带 36-60px，669d 型粗行带 ~296px；
            # 0683 弯纸上半段水平线漏检一半，行带高约 1.5 倍行距（74px vs 48px），
            # 带内确实挤了 2 个完整单号——只按 >200px 会漏掉这种"半粗"行带。
            # 因此再补一条窄条件：带高 ≥1.4 倍中位码框间距，且单号列内确有
            # ≥2 个完整码框（下面的 frames >= 2）才切，普通行带不受影响。
            gap_ratio = h / med_gap if med_gap > 0 else 99.0
            if (h > 200 and h > 2.5 * med_gap) or gap_ratio >= 1.4:
                trigger[ri] = frames
    if not trigger:
        return 0
    # 2) 逐触发带切分。段界用 warp 码框自身的 y（与内容框同源；若用原图码
    #    单应逆映射的 wy，两套 y 差半个框高会把内容框错归到相邻段）。
    new_rows = []
    new_row_y = []
    split_count = 0
    for ri, (a, b) in enumerate(row_y):
        ys = trigger.get(ri)
        if not ys or len(ys) < 2:
            new_rows.append(rows[ri])
            new_row_y.append((a, b))
            continue
        bounds = [a] + [(ys[i] + ys[i + 1]) / 2.0 for i in range(len(ys) - 1)] + [b]
        # 带内全部 warp 框（不限列）
        band_frames = []
        for box in ocr:
            bb = box.get("bbox")
            if not bb or len(bb) < 4:
                continue
            cx, cy = (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0
            if a <= cy <= b:
                band_frames.append((cy, cx, str(box.get("text", ""))))
        n_seg = len(bounds) - 1
        # 段界用单号框中点，内容框按自身中心落段。这里刻意保持“逐框就近”，
        # 不做按列序号配对：669d 型宽表格的签收/数量列文字框高度不一，按列
        # 顺序配对会把整列内容整体挪到下一行（实测 669d 下段序号 47-49 错
        # 行、末段单号重复）。0683 型的列垂直偏移由上面的弯纸整表重建处理。
        for si in range(n_seg):
            y0, y1 = bounds[si], bounds[si + 1]
            cells = []
            for ci in range(len(col_x)):
                # 仅保留属于本列的文本，列序与 col_x 一致
                col_texts = [t for cy, cx, t in sorted(band_frames)
                             if y0 <= cy <= y1
                             and col_x[ci][0] <= cx <= col_x[ci][1] and t]
                cells.append(" ".join(col_texts).strip())
            new_rows.append(cells)
            new_row_y.append((y0, y1))
        split_count += n_seg - 1
        log(f"{name}：行带二次切行 第{ri+1}带 {a:.0f}-{b:.0f}（高{b-a:.0f}px）"
            f"切出 {n_seg} 个物理行")
    if split_count <= 0:
        return 0
    main.rows = new_rows
    main.row_y = new_row_y
    log(f"{name}：行带二次切行完成：{len(rows)} 行 -> {len(new_rows)} 行")
    return split_count


def build_exact_page(image_path, worker: OcrWorker, on_log=None,
                     do_fill: bool = True, should_stop=None) -> ExactPage:
    """单图精确复制管线（默认纯本地，可选的云端文本纠错默认关闭）。

    两条识别路径，优先表格网格法：
      A. 网格法（手机实拍表格）：检测网格线 -> 透视矫正拉正 -> 按线切格 ->
         框按中心点落格 -> 空格放大补识。列由竖线决定，合并单元格天然正确，
         不会出现「日期+承运商+运单号粘成一列」的塌陷。
      B. 几何重建法（扫描件/截图/无线表格）：整图 OCR -> 规则级校准 ->
         行/列聚类 -> 序号断档补行。
    A 检测不到网格或结果质量不足时自动回退 B。
    之后统一做交叉校验（整图 OCR vs 局部放大 OCR 两方对比）：
       编码/数字类字段以放大复识结果自动改写，文本字段只标记待复核。
    放大复识是唯一的独立第二意见：整图 OCR 在缩放后的小字易误识，
    裁剪放大重识别可纠正这类错误（运单号末几位最典型）。
    should_stop 提供取消检查（各阶段均可中断）。
    """
    image_path = Path(image_path)
    name = image_path.name

    def log(msg: str) -> None:
        if on_log:
            on_log(msg)

    def stopped() -> bool:
        return should_stop is not None and should_stop()

    if stopped():
        raise EngineError("已取消转换。")

    page = None
    ocr_boxes: list = []
    ref_path = image_path
    warp_tmp: Path | None = None
    try:
        if not stopped():
            try:
                grid = _grid_pipeline(image_path, worker, name, log, should_stop=should_stop)
            except Exception as exc:  # 网格法任何异常都不阻塞主流程（内存不足除外）
                if _is_memory_error(exc):
                    raise
                log(f"{name}：网格管线异常，回退几何重建（{exc}）")
                grid = None
            if grid is not None:
                page, ocr_boxes, warp_tmp = grid
                ref_path = warp_tmp

        if page is None:
            result = _run_local_pipeline(image_path, worker, name, log, do_fill=do_fill)
            if result["error"]:
                return ExactPage(name=name, blocks=[], warnings=[f"识别失败：{result['error']}"])
            page = result["local_page"]
            ocr_boxes = result["ocr_boxes"]

        # ---- 行带内多物理行二次切行（669d 型：296px 行带挤 4-7 个物理行）----
        # 放在交叉校验之前：让 xcheck 校验切分后的正确结构；触发极窄（45 张仅 669d）。
        _split_n = 0
        if warp_tmp is not None:
            try:
                _split_n = _split_dense_bands(page, name, log)
            except Exception as exc:
                log(f"{name}：行带二次切行跳过（{exc}）")

        # ---- 交叉校验：整图 OCR vs 局部放大 OCR 两方对比 ----
        warnings: list[str] = list(getattr(page, "warnings", []) or [])
        flagged: list[dict] = list(getattr(page, "flags", []) or [])
        for block in getattr(page, "blocks", []) or []:
            if getattr(block, "kind", "") != "table":
                continue
            rows = getattr(block, "rows", [])
            if not rows:
                continue
            if stopped():
                raise EngineError("已取消转换。")
            try:
                corrections, flags = cross_check(
                    rows, ocr_boxes, [block], worker, ref_path,
                    on_log=log, check_text=True,
                    skip_header=bool(getattr(block, "header_rebuilt", False)),
                )
            except Exception as exc:
                log(f"{name}：交叉校验跳过（{exc}）")
                continue
            for c in corrections:
                ri, ci = c["row"], c["col"]
                if not (0 <= ri < len(rows) and 0 <= ci < len(rows[ri])):
                    continue
                # 前缀保护：运单号前缀（YT 等）是固定标识，
                # 放大复识若动到前缀（YT->TT/NT）一律拒绝，保持原值
                if not _prefix_compatible(rows[ri][ci], c["new"]):
                    log(f"{name}：交叉校验拒绝改写第{ri+1}行第{ci+1}列"
                        f"（前缀不一致：{rows[ri][ci]} -> {c['new']}）")
                    continue
                rows[ri][ci] = c["new"]
            for item in flags:
                flagged.append(item)
                warnings.append(
                    f"待复核：第{item['row'] + 1}行第{item['col'] + 1}列"
                    f"「{item.get('rebuilt', '')}」（{item['reason']}）"
                )
            if corrections or flags:
                log(f"{name}：交叉校验完成，自动改写 {len(corrections)} 处，"
                    f"待复核 {len(flags)} 处")

        # ---- 单号局部放大修复：网格漏读的码经单应定位+格子自证+双读确认 ----
        # 放在单元格清洗之前：此时格子仍含状态粘合文本，同源尾号判定信息最全；
        # 写入的完整码格子随后由 _split_code_status 正常拆分，格式与既有一致。
        if warp_tmp is not None and GRID_ZOOM_REPAIR:
            try:
                _zoom_repair_codes(page, image_path, worker,
                                   name, log, warnings, force=_split_n > 0)
            except Exception as exc:
                log(f"{name}：单号局部放大修复跳过（{exc}）")

        # 网格法的单元格清洗放在交叉校验之后：校验要拿落格原始值与 OCR 原始框
        # 对比，清洗（行号剥离/前置运单号剥离/承运商共识）在最后做，避免把
        # "清洗后的正确值 vs 原始脏值"误报成待复核。
        if warp_tmp is not None:
            for block in getattr(page, "blocks", []) or []:
                if getattr(block, "kind", "") != "table":
                    continue
                try:
                    block.rows = clean_grid_rows(block.rows, on_log=log)
                except Exception as exc:
                    log(f"{name}：单元格清洗跳过（{exc}）")
                # 包裹号修复：整列大面积丢失前缀时先做行带左扩像素复识，
                # 再按列内 `YT`+13 位格式补回被吞掉的前缀（见函数注释）。
                try:
                    pkg_ci, pkg_stat = _locate_pkg_col(block.rows)
                except Exception as exc:
                    pkg_ci, pkg_stat = None, {}
                    log(f"{name}：包裹号列定位跳过（{exc}）")
                if pkg_ci is not None and GRID_PKG_REPAIR:
                    total = max(int(pkg_stat.get("total", 0)), 1)
                    broken = total - int(pkg_stat.get("yt", 0))
                    if broken >= 3 and broken / total >= 0.6:
                        try:
                            _reread_invalid_pkg(block, pkg_ci, worker, ref_path,
                                                name, log, warnings, flagged)
                        except Exception as exc:
                            log(f"{name}：包裹号左扩复识跳过（{exc}）")
                    try:
                        _normalize_pkg_prefix(block.rows, pkg_ci, log,
                                              flagged=flagged, warnings=warnings)
                    except Exception as exc:
                        log(f"{name}：包裹号前缀还原跳过（{exc}）")
                try:
                    _split_code_status(block.rows, flagged, log)
                except Exception as exc:
                    log(f"{name}：单号/状态拆分跳过（{exc}）")
                try:
                    _mark_uncertain_received(block.rows, flagged, log)
                except Exception as exc:
                    log(f"{name}：实收列待复核标记跳过（{exc}）")
                try:
                    _mark_numeric_glue(block.rows, flagged, log)
                except Exception as exc:
                    log(f"{name}：数字列粘连检查跳过（{exc}）")

        # ---- 单号补录：默认关闭（GRID_CODE_SUPPLEMENT，见常量区注释）----
        if warp_tmp is not None and GRID_CODE_SUPPLEMENT:
            grid_codes = _table_code_set(page)
            grid_rows = sum(len(getattr(b, "rows", None) or [])
                            for b in getattr(page, "blocks", []) or []
                            if getattr(b, "kind", "") == "table")
            log(f"{name}：补录前置：warp={warp_tmp is not None}，表行={grid_rows}，"
                f"网格单号={len(grid_codes)}")
            if (grid_rows >= 5 and grid_codes
                    and len(grid_codes) / max(grid_rows, 1) < 0.5):
                _supplement_grid_codes(page, image_path, worker, name, log, warnings)

        # 可选的云端文本兜底：默认关闭；开启后只外发仍待复核的文本类单元格图
        try:
            _cloud_text_pass(page, worker, ref_path, flagged, log)
        except Exception as exc:
            log(f"{name}：云端纠错整体跳过（{exc}）")

        final = ExactPage(name=name, blocks=page.blocks, warnings=warnings, flags=flagged)
        final.image_path = str(image_path)
        final.to_original = getattr(page, "to_original", None)
        final.ocr_boxes = ocr_boxes
        final.quality_meta = getattr(page, "quality_meta", {})
        return final
    finally:
        if warp_tmp is not None:
            try:
                warp_tmp.unlink(missing_ok=True)
            except OSError:
                pass


def convert_images_exact(image_paths, out_path, worker: OcrWorker,
                         on_log=None) -> ConvertResult:
    """多图精确复制：逐图建 ExactPage，合并写入一个 Excel（每图一个工作表）。"""
    pages = []
    for path in image_paths:
        page = build_exact_page(path, worker, on_log=on_log)
        pages.append(page)
        if on_log:
            tables = sum(1 for b in page.blocks if getattr(b, "kind", "") == "table")
            on_log(f"{page.name}：重建完成，{tables} 个表格、"
                   f"{len(page.blocks) - tables} 个文字块")
    return convert_exact(pages, Path(out_path))
