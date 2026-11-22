# -*- coding: utf-8 -*-
"""肥鸭文档转换 —— 桌面程序入口（本地服务 + WebView2 原生窗口）。

启动流程：
  1. 单实例检查（命名互斥体，第二次启动通过命名事件唤起已有窗口）
  2. 起本地 FastAPI 服务（仅 127.0.0.1），就绪后创建无边框主窗口
  3. 主窗口记住上次的大小/位置；启动即在系统托盘常驻图标，托盘菜单显示
     任务进度并可取消当前任务；最小化保持任务栏行为；关闭窗口（X）收进
     托盘继续运行
  4. 只有托盘菜单「退出」才真正结束进程：转换中先弹确认框（可取消退出）
     → 保存窗口几何 → 停本地服务 → 退出进程；任务结束在托盘中弹气泡通知

兼容性：若机器没有 WebView2 运行时，回退到 Edge/Chrome 的 app 模式窗口
（此时靠前端心跳判断关窗），保证程序始终可用。
"""
from __future__ import annotations

import ctypes
import os
import secrets
import sys
import threading
import time
import traceback
import urllib.request
from pathlib import Path

import uvicorn

import core
import web_app as backend

# 首屏 URL 带一次性 nonce：确保每次启动都回源取最新页面并刷新 token cookie。
# 若命中 WebView2 缓存，服务端的安全中间件不会执行，页面会带着旧 token 调用 API
# 被 403，表现为「首次使用」弹窗显示 undefined、界面功能全部失效。
URL = (f"http://{backend.HOST}:{backend.PORT}/web/index.html"
       f"?v={core.APP_VERSION}&n={secrets.token_hex(4)}")

DEFAULT_W, DEFAULT_H = 1120, 860
MIN_W, MIN_H = 960, 680
GEO_KEYS = ("win_w", "win_h", "win_x", "win_y", "win_maximized")
_NATIVE_DRAG_LOGGED = False


def _log(text: str) -> None:
    core.startup_log(text)


# ---------------------------------------------------------------- 资源

def _asset(*names: str) -> Path | None:
    """按 exe 同级 → _MEIPASS → 源码目录的顺序找资源。"""
    bases = []
    try:
        bases.append(Path(sys.executable).resolve().parent / "assets")
    except Exception:
        pass
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bases.append(Path(meipass) / "assets")
    bases.append(Path(__file__).resolve().parent / "assets")
    for base in bases:
        target = base.joinpath(*names)
        if target.is_file():
            return target
    return None


# ---------------------------------------------------------------- 窗口几何记忆

def _screen_size() -> tuple[int, int]:
    try:
        user32 = ctypes.windll.user32
        return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
    except Exception:
        return 1920, 1080


def _clamp_geometry(x: int, y: int, w: int, h: int) -> tuple[int, int, int, int]:
    """把保存的窗口位置夹到当前屏幕内，避免换显示器后窗口跑到屏幕外。"""
    sw, sh = _screen_size()
    w = max(MIN_W, min(int(w), sw))
    h = max(MIN_H, min(int(h), sh))
    if x < -w + 120 or y < 0 or x > sw - 120 or y > sh - 120:
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
    return int(x), int(y), int(w), int(h)


def _load_geometry() -> dict:
    cfg = core.load_config()
    sw, sh = _screen_size()
    try:
        w = int(cfg.get("win_w") or DEFAULT_W)
        h = int(cfg.get("win_h") or DEFAULT_H)
        x = int(cfg.get("win_x"))
        y = int(cfg.get("win_y"))
    except (TypeError, ValueError):
        w, h, x, y = DEFAULT_W, DEFAULT_H, -1, -1
    # 过滤掉最小化时误存的几何（-32000 / 极小尺寸）
    if w < MIN_W or h < MIN_H or x <= -30000 or y <= -30000:
        w, h = DEFAULT_W, DEFAULT_H
        x, y = max(0, (sw - w) // 2), max(0, (sh - h) // 2)
    x, y, w, h = _clamp_geometry(x, y, w, h)
    return {"x": x, "y": y, "width": w, "height": h, "maximized": bool(cfg.get("win_maximized"))}


def _save_geometry(geo: dict, maximized: bool) -> None:
    """保存窗口几何；拒绝明显无效的值（例如最小化时的 160x28 / -32000）。"""
    try:
        w = int(geo.get("width") or 0)
        h = int(geo.get("height") or 0)
        x = int(geo.get("x") or 0)
        y = int(geo.get("y") or 0)
        if w < MIN_W or h < MIN_H or x <= -30000 or y <= -30000:
            _log(f"忽略无效窗口几何：{w}x{h} @ {x},{y}")
            return
        cfg = core.load_config()
        cfg["win_w"] = w
        cfg["win_h"] = h
        cfg["win_x"] = x
        cfg["win_y"] = y
        cfg["win_maximized"] = bool(maximized)
        core.save_config(cfg)
    except Exception:
        _log("保存窗口几何失败：\n" + traceback.format_exc())


# ---------------------------------------------------------------- 托盘

def _confirm_exit_during_job() -> bool:
    """转换中请求退出时弹系统确认框；返回 True 表示可以退出。

    pywebview 6 的 closing 事件同步执行，处理器返回 False 即取消关窗
    （winforms.on_closing → args.Cancel，已实证）；托盘退出走同一确认。
    """
    if not backend.has_running_job():
        return True
    _log("转换中请求退出，弹出确认框")
    try:
        result = ctypes.windll.user32.MessageBoxW(
            0,
            "正在转换中，退出将中断本次转换（不会生成半成品文件）。\n\n确定要退出吗？",
            core.APP_TITLE,
            0x1 | 0x30 | 0x100,  # OK/Cancel + 警告图标 + 默认焦点在「取消」
        )
    except Exception:
        _log("退出确认框弹出失败，放行退出：\n" + traceback.format_exc())
        return True
    if result == 1:
        _log("用户确认中断转换并退出")
        return True
    _log("用户取消退出，继续转换")
    return False


def _system_shutting_down() -> bool:
    """系统注销/关机时要放行关窗，否则 Windows 会卡在关机界面。"""
    try:
        return bool(ctypes.windll.user32.GetSystemMetrics(0x2000))  # SM_SHUTTINGDOWN
    except Exception:
        return False


# ---------------------------------------------------------- 第二实例唤起窗口

ACTIVATE_EVENT = "肥鸭文档转换_ActivateWindow"
_activate_handle = {"h": None}


def _create_activate_event() -> None:
    """首个实例创建命名事件；第二个实例 SetEvent 即可把已有窗口唤到前台。"""
    try:
        _activate_handle["h"] = ctypes.windll.kernel32.CreateEventW(
            None, False, False, ACTIVATE_EVENT)
    except Exception:
        _log("创建激活事件失败：\n" + traceback.format_exc())


def _signal_existing_instance() -> bool:
    try:
        k32 = ctypes.windll.kernel32
        h = k32.OpenEventW(0x0002, False, ACTIVATE_EVENT)  # EVENT_MODIFY_STATE
        if not h:
            return False
        try:
            return bool(k32.SetEvent(h))
        finally:
            k32.CloseHandle(h)
    except Exception:
        return False


def _wait_activate(tray: "Tray") -> None:
    handle = _activate_handle["h"]
    if not handle:
        return
    while True:
        try:
            # 超时轮询而非无限等待：窗口销毁后线程也能自然退出循环
            if ctypes.windll.kernel32.WaitForSingleObject(handle, 1000) == 0:
                _log("收到第二实例唤起请求，显示主窗口")
                tray.show()
        except Exception:
            _log("等待唤醒事件失败：\n" + traceback.format_exc())
            return


class Tray:
    """常驻系统托盘：显示主界面 / 取消任务 / 退出，并转发任务结果通知。

    启动即创建图标，保证托盘区（右下角「小任务栏」）始终能看到本程序；
    关闭窗口只是隐藏主窗口，图标不销毁。Windows 11 默认把新图标
    收进「隐藏的图标」浮层，可在 设置 → 个性化 → 任务栏 → 其他系统托盘图标
    中把它设为「显示」，或从浮层直接拖到任务栏上。
    """

    def __init__(self, window, on_show=None, is_hidden=None, on_quit=None) -> None:
        self.window = window
        self.on_show = on_show
        self.is_hidden = is_hidden or (lambda: False)
        self.on_quit = on_quit
        self.icon = None
        self.need_restore = False          # 从最小化态恢复时才需要 restore()
        self._thread = None
        self._watching = False
        self._status_key = None
        self._notified_job = ""

    def start(self) -> None:
        if self.icon is not None:
            return
        try:
            import pystray
            from PIL import Image
        except Exception:
            _log("托盘依赖不可用：关闭窗口将直接退出：\n" + traceback.format_exc())
            return
        icon_path = _asset("app.ico") or _asset("app.png")
        if icon_path is None:
            return
        try:
            image = Image.open(icon_path)
            image = image.convert("RGBA").resize((64, 64))
            menu = pystray.Menu(
                pystray.MenuItem(self._status_text, None, enabled=False),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("显示主界面", lambda *_: self.show(), default=True),
                pystray.MenuItem("取消当前任务", lambda *_: self.cancel_job(),
                                 enabled=lambda *_: bool(backend.job_snapshot().get("running"))),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("退出", lambda *_: self.quit()),
            )
            self.icon = pystray.Icon("肥鸭文档转换", image,
                                     "肥鸭文档转换 · 空闲", menu)
            self._thread = threading.Thread(target=self.icon.run, daemon=True)
            self._thread.start()
            if not self._watching:
                self._watching = True
                threading.Thread(target=self._watch, daemon=True).start()
            _log("托盘图标已创建（常驻）")
        except Exception:
            _log("创建托盘图标失败：\n" + traceback.format_exc())

    # ------------------------------------------------------------ 外部操作

    def show(self) -> None:
        # 顺序很重要：先 restore 再 show。若先 show，窗口会以「最小化」状态显示，
        # 点托盘图标后窗口一闪就没了，无法正常拉起。
        if self.need_restore:
            try:
                self.window.restore()
            except Exception:
                pass
            self.need_restore = False
        try:
            self.window.show()
        except Exception:
            pass
        if self.on_show is not None:
            try:
                self.on_show()
            except Exception:
                pass

    def stop(self) -> None:
        """退出时移除托盘图标，避免留下点不动的幽灵图标。"""
        self._stop_icon()

    def cancel_job(self) -> None:
        count = backend.cancel_active_jobs()
        _log(f"托盘请求取消当前任务（{count} 个）")

    def quit(self) -> None:
        # 真正退出只走托盘菜单；窗口 X 是收进托盘继续运行
        if not _confirm_exit_during_job():
            return
        if self.on_quit is not None:
            try:
                self.on_quit()
            except Exception:
                pass
        try:
            self.window.destroy()
        except Exception:
            pass
        self._stop_icon()

    # ------------------------------------------------------------ 状态与通知

    def _status_text(self, *_args) -> str:
        snap = backend.job_snapshot()
        if snap.get("running"):
            verb = "查找" if snap.get("mode") == "find" else "转换"
            total = int(snap.get("total") or 0)
            done = int(snap.get("processed") or 0)
            return f"正在{verb} {done}/{total}" if total > 0 else f"正在{verb}…"
        return {"ok": "上次任务：已完成", "error": "上次任务：失败",
                "cancelled": "上次任务：已取消"}.get(str(snap.get("status")), "空闲")

    def _watch(self) -> None:
        """每秒同步一次任务状态：更新托盘提示/菜单，完成后弹气泡。"""
        while self.icon is not None:
            try:
                snap = backend.job_snapshot()
                key = (snap.get("running"), snap.get("processed"),
                       snap.get("total"), snap.get("status"), snap.get("mode"))
                if key != self._status_key:
                    self._status_key = key
                    try:
                        self.icon.title = f"肥鸭文档转换 · {self._status_text()}"
                    except Exception:
                        pass
                    try:
                        self.icon.update_menu()
                    except Exception:
                        pass
                self._notify_finished(snap)
            except Exception:
                pass
            time.sleep(1.0)

    def _notify_finished(self, snap: dict) -> None:
        if snap.get("running"):
            return
        status = str(snap.get("status") or "")
        if status not in ("ok", "error", "cancelled"):
            return
        job_key = str(snap.get("job_id") or snap.get("started_at") or "")
        if not job_key or job_key == self._notified_job:
            return
        self._notified_job = job_key
        if not self.is_hidden():
            return  # 窗口可见时界面自己有结果卡片，不重复打扰
        try:
            if status == "ok":
                verb = "单号查找" if snap.get("mode") == "find" else "转换"
                title = f"{verb}完成"
                files = int(snap.get("files") or 0)
                hits = int(snap.get("waybills") or 0)
                if snap.get("mode") == "find":
                    msg = f"{files} 个文件，找到 {hits} 个单号"
                else:
                    msg = (f"{files} 个文件 / {int(snap.get('total_pages') or 0)} 页，"
                           f"{hits} 个单号")
                name = Path(str(snap.get("out_path") or "")).name
                if name:
                    msg += f"\n已保存：{name}"
            elif status == "error":
                title = "转换失败"
                msg = str(snap.get("message") or "请打开主界面查看原因")[:150]
            else:
                title = "转换已取消"
                msg = "任务已停止，未生成半成品文件"
            self.icon.notify(msg, title)
        except Exception:
            _log("托盘通知发送失败：\n" + traceback.format_exc())

    def _stop_icon(self) -> None:
        try:
            if self.icon is not None:
                self.icon.stop()
        except Exception:
            pass
        self.icon = None


# ---------------------------------------------------------- 无边框窗口控制

class WindowApi:
    """暴露给前端标题栏的窗口控制接口（window.pywebview.api.*）。

    注意：属性必须全部以下划线开头。pywebview 启动时会递归枚举 js_api 对象的
    公开属性来生成 JS 接口，实例上的 window 属性会让它去遍历整个 WinForms
    对象图（window.native...），递归爆栈并长时间占着 GIL，界面直接卡死。
    """

    def __init__(self) -> None:
        self._window = None
        self._state = None
        self._restore_bounds = None   # 最大化前的窗口几何（还原时恢复用）

    def _maximized(self) -> bool:
        try:
            native = getattr(self._window, "native", None)
            if native is not None:
                return str(native.WindowState) == "Maximized"
        except Exception:
            pass
        return bool(self._state and self._state.get("maximized"))

    def win_state(self) -> bool:
        """当前是否最大化；页面启动后用于同步按钮图标。"""
        return self._maximized()

    def win_activate(self) -> bool:
        """鼠标移入窗口时自动激活，让随后的 Ctrl+V 直接落到本程序。

        可在设置里关闭（hover_activate=false），避免在其它程序里打字时
        鼠标划过本窗口把焦点抢走。
        """
        if self._window is None or not _config_bool("hover_activate", True):
            return False
        try:
            return _bring_to_front(self._window)
        except Exception:
            _log("鼠标移入激活窗口失败：\n" + traceback.format_exc())
            return False

    def win_minimize(self) -> bool:
        if self._window is not None:
            self._window.minimize()
        return True

    def win_toggle_maximize(self) -> bool:
        if self._window is None:
            return False
        target = not self._maximized()
        try:
            if target:
                # 记下最大化前的几何：无边框窗口没有原生 thick frame，
                # WinForms 自己记的 RestoreBounds 与实际窗口对不上，直接还原
                # 会让窗口小一圈，反复还原越缩越小。
                self._restore_bounds = _window_rect(self._window)
                self._window.maximize()
            else:
                self._window.restore()
                if self._restore_bounds:
                    _set_window_bounds(self._window, *self._restore_bounds)
        except Exception:
            _log("窗口最大化/还原失败：\n" + traceback.format_exc())
            return self._maximized()
        return target

    def win_close(self) -> bool:
        """走 window.destroy()，仍会触发 closing 事件里的转换中退出确认。"""
        if self._window is not None:
            self._window.destroy()
        return True

    def win_bounds(self):
        """当前窗口的位置与尺寸（逻辑像素），供前端缩放热区计算。"""
        if self._window is None:
            return None
        try:
            x, y, w, h = _window_rect(self._window)
            return {"x": x, "y": y, "width": w, "height": h,
                    "maximized": self._maximized()}
        except Exception:
            _log("读取窗口几何失败：\n" + traceback.format_exc())
            return None

    def win_resize(self, x: int, y: int, width: int, height: int) -> bool:
        """按逻辑像素改窗口位置与尺寸（前端八方向缩放热区调用）。"""
        if self._window is None or self._maximized():
            return False
        try:
            w = max(MIN_W, int(width))
            h = max(MIN_H, int(height))
            _set_window_bounds(self._window, int(x), int(y), w, h)
            return True
        except Exception:
            _log("调整窗口尺寸失败：\n" + traceback.format_exc())
            return False

    def win_drag_start(self, cursor_x: float = 0):
        """开始拖动标题栏，返回窗口当前逻辑坐标。

        最大化时先还原窗口，再按鼠标在标题栏上的横向比例摆好位置，避免一拖就跳。
        """
        if self._window is None:
            return None
        try:
            restored = False
            if self._maximized():
                x0, _, w0, _ = _window_rect(self._window)
                bounds = self._restore_bounds
                if not bounds and self._state:
                    last = self._state.get("last") or {}
                    if last.get("width") and last.get("height"):
                        bounds = (last.get("x", 0), last.get("y", 0),
                                  last.get("width"), last.get("height"))
                try:
                    self._window.restore()
                except Exception:
                    _log("窗口还原失败：\n" + traceback.format_exc())
                deadline = time.time() + 0.4
                while time.time() < deadline:
                    try:
                        if str(self._window.native.WindowState) != "Maximized":
                            break
                    except Exception:
                        break
                    time.sleep(0.02)
                restored = True
                if bounds:
                    try:
                        bx, by, bw, bh = bounds
                        sw, _ = _screen_size()
                        if w0 > 0:
                            # 按鼠标在标题栏上的横向比例摆好，避免一拖就跳
                            bx = int(round(cursor_x - (cursor_x - x0) / w0 * bw))
                        bx = max(-bw + 120, min(bx, sw - 120))
                        _set_window_bounds(self._window, bx, by, bw, bh)
                    except Exception:
                        _log("最大化窗口拖动还原定位失败：\n" + traceback.format_exc())
            x, y, _, _ = _window_rect(self._window)
            return {"x": x, "y": y, "restored": restored}
        except Exception:
            _log("开始拖动窗口失败：\n" + traceback.format_exc())
            return None

    def win_move(self, x: int, y: int) -> bool:
        """拖动过程中按逻辑像素移动窗口（直接 SetWindowPos，绕开 WebView 层的坐标换算）。"""
        if self._window is None:
            return False
        try:
            _move_window(self._window, int(x), int(y))
            return True
        except Exception:
            return False

    def win_drag_native(self) -> bool:
        """把标题栏拖动交给 Windows 原生移动循环（标准无边框窗口做法）。

        等价于以 HTCAPTION 命中标题栏：ReleaseCapture + WM_NCLBUTTONDOWN。
        之后鼠标捕获、贴边分屏、跨显示器 DPI、最大化拖动还原都由系统处理，
        不需要前端逐帧算坐标再 SetWindowPos（那条路径会掉帧、跟手差、跨屏跳变）。

        必须在 UI 线程执行，否则 ReleaseCapture 释放的不是 WebView 所在线程的
        捕获；这里用 WinForms 的 BeginInvoke 投递，避免阻塞 JS 桥接线程。
        """
        if self._window is None:
            return False
        native = getattr(self._window, "native", None)
        if native is None:
            return False
        try:
            hwnd = _window_hwnd(self._window)
        except Exception:
            _log("获取窗口句柄失败：\n" + traceback.format_exc())
            return False

        def _start_move() -> None:
            global _NATIVE_DRAG_LOGGED
            try:
                if not _NATIVE_DRAG_LOGGED:
                    _NATIVE_DRAG_LOGGED = True
                    _log("原生窗口拖动已启用（ReleaseCapture + WM_NCLBUTTONDOWN）")
                user32 = ctypes.windll.user32
                user32.ReleaseCapture()
                WM_NCLBUTTONDOWN, HTCAPTION = 0x00A1, 2
                user32.SendMessageW(ctypes.c_void_p(hwnd), WM_NCLBUTTONDOWN,
                                    ctypes.c_void_p(HTCAPTION), ctypes.c_void_p(0))
            except Exception:
                _log("原生窗口拖动失败：\n" + traceback.format_exc())

        try:
            from System import Action  # pythonnet 已随 pywebview 加载
            native.BeginInvoke(Action(_start_move))
            return True
        except Exception:
            _log("原生拖动不可用，回退到逐帧移动：\n" + traceback.format_exc())
            return False


def _window_hwnd(window) -> int:
    """取 WinForms 窗口句柄（pythonnet 的 Handle 是 IntPtr，不能直接 int()）。"""
    handle = window.native.Handle
    return int(handle.ToInt64()) if hasattr(handle, "ToInt64") else int(handle)


def _style_frameless_window(window) -> None:
    """去掉窗口自带的边框色，只保留 Windows 11 圆角。

    无边框窗口一旦补 WS_THICKFRAME，系统会在网页内容上方额外画一条
    非客户区色带；它与网页顶栏的紫色渐变不可能完全一致，窗口失焦时还会
    变成浅色，看起来就是“两层”。所以缩放改由前端热区 + SetWindowPos
    完成，这里只把最外 1px 的 DWM 边框隐藏掉，让内容真正铺满整窗。
    """
    try:
        hwnd = _window_hwnd(window)
    except Exception:
        return
    try:
        dwm = ctypes.windll.dwmapi
        corner = ctypes.c_int(2)          # DWMWCP_ROUND
        dwm.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(corner), 4)
        border = ctypes.c_uint(0xFFFFFFFE)  # DWMWA_COLOR_NONE：不画边框线
        dwm.DwmSetWindowAttribute(hwnd, 34, ctypes.byref(border), 4)
    except Exception:
        _log("设置无边框窗口样式失败：\n" + traceback.format_exc())


def _enable_taskbar_minimize(window) -> None:
    """给无边框窗口补上 WS_MINIMIZEBOX/WS_SYSMENU。

    FormBorderStyle.None 的窗口没有最小化框样式位，Windows 认为它「不可最小化」：
    窗口在前台时点任务栏图标毫无反应，只能点自绘的最小化按钮。
    这里只加样式位，不加 WS_CAPTION/WS_THICKFRAME，窗口外观与边框保持不变；
    补上后任务栏图标即可像普通程序一样「点一次收回、再点一次还原」。
    """
    try:
        hwnd = _window_hwnd(window)
        user32 = ctypes.windll.user32
        gwl_style = -16
        ws_minimizebox, ws_maximizebox, ws_sysmenu = 0x00020000, 0x00010000, 0x00080000
        style = user32.GetWindowLongW(hwnd, gwl_style) & 0xFFFFFFFF
        user32.SetWindowLongW(hwnd, gwl_style,
                              style | ws_minimizebox | ws_maximizebox | ws_sysmenu)
        # NOSIZE|NOMOVE|NOZORDER|FRAMECHANGED：立即生效，但不移动/重排窗口
        user32.SetWindowPos(ctypes.c_void_p(hwnd), ctypes.c_void_p(0), 0, 0, 0, 0,
                            0x0001 | 0x0002 | 0x0004 | 0x0020)
        _log("已启用任务栏图标最小化/还原")
    except Exception:
        _log("启用任务栏最小化失败：\n" + traceback.format_exc())


def _config_bool(key: str, default: bool = True) -> bool:
    """读取布尔配置；缺失或读取失败时用默认值。"""
    try:
        value = core.load_config().get(key)
    except Exception:
        return default
    return default if value is None else bool(value)


def _bring_to_front(window) -> bool:
    """把主窗口带回前台；用于鼠标移入自动激活与全局快捷键。"""
    try:
        hwnd = _window_hwnd(window)
    except Exception:
        return False
    u32 = ctypes.windll.user32
    handle = ctypes.c_void_p(hwnd)
    try:
        if u32.IsIconic(handle):
            u32.ShowWindow(handle, 9)          # SW_RESTORE
        elif not u32.IsWindowVisible(handle):
            u32.ShowWindow(handle, 5)          # SW_SHOW
    except Exception:
        pass
    if u32.GetForegroundWindow() == hwnd:
        return True
    try:
        u32.SetForegroundWindow(handle)
        u32.BringWindowToTop(handle)
    except Exception:
        pass
    if u32.GetForegroundWindow() == hwnd:
        return True
    # 前台锁定（Windows 限制）：把本线程输入队列挂到当前前台线程上再试一次
    try:
        foreground = u32.GetForegroundWindow()
        if foreground:
            fg_thread = u32.GetWindowThreadProcessId(foreground, None)
            cur_thread = ctypes.windll.kernel32.GetCurrentThreadId()
            if fg_thread and fg_thread != cur_thread:
                u32.AttachThreadInput(cur_thread, fg_thread, True)
                try:
                    u32.SetForegroundWindow(handle)
                    u32.BringWindowToTop(handle)
                finally:
                    u32.AttachThreadInput(cur_thread, fg_thread, False)
    except Exception:
        pass
    if u32.GetForegroundWindow() == hwnd:
        return True
    # 仍失败时交给 WinForms 在 UI 线程里 Activate 一次
    try:
        native = getattr(window, "native", None)
        if native is not None:
            from System import Action

            def _activate_native() -> None:
                try:
                    native.Activate()
                except Exception:
                    pass

            native.BeginInvoke(Action(_activate_native))
    except Exception:
        pass
    return False


# ---------------------------------------------------------- 全局粘贴快捷键
_HOTKEY_ID = 0x1D2C
PASTE_HOTKEY_LABEL = "Ctrl+Shift+Alt+V"


def _hotkey_paste(window, tray) -> None:
    """全局快捷键触发：窗口带回前台，并让页面读取剪贴板图片。"""
    try:
        tray.show()
    except Exception:
        pass
    _bring_to_front(window)
    time.sleep(0.15)   # 等窗口从托盘/最小化恢复后再调用页面
    try:
        window.evaluate_js(
            "window.__pasteFromHotkey && window.__pasteFromHotkey();")
    except Exception:
        _log("全局粘贴调用失败：\n" + traceback.format_exc())


def _start_paste_hotkey(window, tray) -> None:
    """注册全局粘贴快捷键（Ctrl+Shift+Alt+V）。

    在微信/浏览器/资源管理器复制图片后，不用切换窗口，按快捷键即可把
    剪贴板图片送进本程序。占用冲突或用户在设置里关闭时跳过。
    """
    if not _config_bool("paste_hotkey", True):
        _log("全局粘贴快捷键已在设置中关闭")
        return

    def loop() -> None:
        u32 = ctypes.windll.user32
        # ALT | CONTROL | SHIFT | MOD_NOREPEAT
        mod = 0x0001 | 0x0002 | 0x0004 | 0x4000
        if not u32.RegisterHotKey(None, _HOTKEY_ID, mod, 0x56):   # VK_V
            _log(f"全局粘贴快捷键 {PASTE_HOTKEY_LABEL} 注册失败（可能被其他程序占用）")
            return
        _log(f"全局粘贴快捷键已注册：{PASTE_HOTKEY_LABEL}")
        from ctypes import wintypes

        msg = wintypes.MSG()
        try:
            while u32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                if msg.message == 0x0312 and int(msg.wParam) == _HOTKEY_ID:
                    threading.Thread(target=_hotkey_paste, args=(window, tray),
                                     daemon=True).start()
                u32.TranslateMessage(ctypes.byref(msg))
                u32.DispatchMessageW(ctypes.byref(msg))
        except Exception:
            _log("全局快捷键监听异常：\n" + traceback.format_exc())
        finally:
            u32.UnregisterHotKey(None, _HOTKEY_ID)

    threading.Thread(target=loop, daemon=True).start()


def _dpi_scale(window) -> float:
    """窗口所在显示器的缩放比例（96 DPI = 1.0）。"""
    try:
        hwnd = _window_hwnd(window)
        dpi = int(ctypes.windll.user32.GetDpiForWindow(ctypes.c_void_p(hwnd)) or 96)
        return dpi / 96.0
    except Exception:
        return 1.0


def _window_rect(window) -> tuple[int, int, int, int]:
    """窗口位置与尺寸，换算成和网页 screenX/screenY 一致的逻辑像素。"""
    hwnd = ctypes.c_void_p(_window_hwnd(window))
    rect = (ctypes.c_long * 4)()
    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
    scale = _dpi_scale(window)
    x = int(round(rect[0] / scale))
    y = int(round(rect[1] / scale))
    w = int(round((rect[2] - rect[0]) / scale))
    h = int(round((rect[3] - rect[1]) / scale))
    return x, y, w, h


def _move_window(window, x: int, y: int) -> None:
    """按逻辑像素移动窗口；SetWindowPos 跨线程调用是安全的。"""
    hwnd = ctypes.c_void_p(_window_hwnd(window))
    scale = _dpi_scale(window)
    SWP_NOSIZE, SWP_NOZORDER, SWP_NOACTIVATE = 0x0001, 0x0004, 0x0010
    ctypes.windll.user32.SetWindowPos(
        hwnd, None, int(round(x * scale)), int(round(y * scale)), 0, 0,
        SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE)


def _set_window_bounds(window, x: int, y: int, w: int, h: int) -> None:
    """按逻辑像素设置窗口位置与尺寸。"""
    hwnd = ctypes.c_void_p(_window_hwnd(window))
    scale = _dpi_scale(window)
    SWP_NOZORDER, SWP_NOACTIVATE = 0x0004, 0x0010
    ctypes.windll.user32.SetWindowPos(
        hwnd, None, int(round(x * scale)), int(round(y * scale)),
        int(round(w * scale)), int(round(h * scale)),
        SWP_NOZORDER | SWP_NOACTIVATE)


# ---------------------------------------------------------------- 主窗口

def _run_webview() -> bool:
    """创建 WebView2 原生窗口；失败返回 False 以便回退。"""
    try:
        import webview
    except Exception:
        _log("pywebview 不可用：\n" + traceback.format_exc())
        return False

    storage = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / core.APP_TITLE / "webview2"
    try:
        storage.mkdir(parents=True, exist_ok=True)
    except OSError:
        storage = Path.home() / ".肥鸭文档转换" / "webview2"
        storage.mkdir(parents=True, exist_ok=True)

    geo = _load_geometry()
    # last：最近一次“正常显示”时的几何；最小化时窗口坐标会变成 -32000，
    # 必须靠这个快照保存，否则下次启动会拿到无效位置。
    state = {
        "maximized": geo["maximized"],
        "minimized": False,
        "hidden": False,      # 窗口是否在托盘里（点关闭后）
        "exiting": False,     # 是否走托盘「退出」/系统关机，允许真正关窗
        "last": {"x": geo["x"], "y": geo["y"],
                 "width": geo["width"], "height": geo["height"]},
    }

    try:
        # 先把本地服务起好，再创建主窗口，窗口直接加载应用页面。
        # 不再使用“闪屏窗口 + 主窗口”两个 WebView2 控制器：打包后先后初始化
        # 容易触发 E_ABORT，导致主窗口刚创建就退出。
        server_thread_start()
        if not _wait_ready():
            _log("服务未就绪，直接打开界面（页面会提示连接失败）")
        api = WindowApi()
        window = webview.create_window(
            core.APP_TITLE, URL,
            x=geo["x"], y=geo["y"], width=geo["width"], height=geo["height"],
            min_size=(MIN_W, MIN_H), background_color="#FFF8E7", text_select=True,
            js_api=api, frameless=True, easy_drag=False, shadow=True,
        )
        api._window = window
        api._state = state

        def _apply_window_chrome():
            # 页面加载完成时原生窗口句柄已就绪；此时再去 DWM 边框与任务栏样式位。
            try:
                _style_frameless_window(window)
                _enable_taskbar_minimize(window)
            except Exception:
                _log("窗口样式应用失败：\n" + traceback.format_exc())

        try:
            window.events.loaded += _apply_window_chrome
        except Exception:
            _log("窗口加载事件订阅失败：\n" + traceback.format_exc())

        def boot():
            try:

                def _push_win_state(maximized: bool) -> None:
                    """把窗口最大化状态同步给标题栏按钮图标。"""
                    try:
                        window.evaluate_js(
                            "window.__winState && window.__winState(%s)"
                            % ("true" if maximized else "false"))
                    except Exception:
                        pass

                def _reset_flags():
                    state["minimized"] = False
                    state["maximized"] = False
                    state["hidden"] = False

                def _begin_exit():
                    state["exiting"] = True

                tray = Tray(window, on_show=_reset_flags,
                            is_hidden=lambda: state["hidden"],
                            on_quit=_begin_exit)
                # 启动即常驻托盘：窗口开着时右下角也能看到本程序
                tray.start()
                # 已有实例再次被双击启动时，唤起本窗口（隐藏到托盘也能唤到）
                threading.Thread(target=_wait_activate, args=(tray,), daemon=True).start()
                # 全局粘贴快捷键：复制图片后不用切窗口，按 Ctrl+Shift+Alt+V 直接贴入
                _start_paste_hotkey(window, tray)

                def _snapshot():
                    if state["minimized"] or state["maximized"]:
                        return
                    try:
                        # 最大化时窗口尺寸是屏幕工作区，不能当成普通尺寸落盘
                        if str(window.native.WindowState) == "Maximized":
                            return
                    except Exception:
                        pass
                    try:
                        x, y = int(window.x), int(window.y)
                        w, h = int(window.width), int(window.height)
                    except Exception:
                        return
                    # 最小化后坐标会变成 -32000、尺寸变成极小值，直接丢弃
                    if w < MIN_W or h < MIN_H or x <= -30000 or y <= -30000:
                        return
                    state["last"] = {"x": x, "y": y, "width": w, "height": h}

                def _on_minimized():
                    # 保持任务栏最小化，方便点任务栏图标或双击恢复；不再收进托盘。
                    # 注意：事件触发时窗口已经最小化，此时不能再取几何
                    state["minimized"] = True
                    # 窗口最小化后仍留在任务栏；标记需要 restore，托盘菜单
                    # 「显示主界面」才能把最小化的窗口正常拉起。
                    if tray.icon is not None:
                        tray.need_restore = True

                def _on_maximized():
                    state["maximized"] = True
                    state["minimized"] = False
                    state["hidden"] = False
                    tray.need_restore = False
                    _push_win_state(True)

                def _on_restored():
                    state["maximized"] = False
                    state["minimized"] = False
                    state["hidden"] = False
                    tray.need_restore = False
                    _push_win_state(False)

                def _hide_to_tray():
                    state["hidden"] = True
                    try:
                        window.hide()
                    except Exception:
                        _log("收进托盘失败：\n" + traceback.format_exc())

                def _on_closing():
                    # 托盘「退出」或系统关机：允许真正关窗，几何与图标收尾
                    if state["exiting"] or _system_shutting_down():
                        state["exiting"] = True
                        if not state["minimized"] and not state["maximized"]:
                            _snapshot()
                        _save_geometry(state["last"], state["maximized"])
                        tray.stop()
                        return True
                    # 用户点 X：收进托盘继续运行，服务与任务都不停
                    if tray.icon is None:
                        # 没有托盘图标就无法找回窗口，退化为真正退出
                        if not _confirm_exit_during_job():
                            return False
                        state["exiting"] = True
                        if not state["minimized"] and not state["maximized"]:
                            _snapshot()
                        _save_geometry(state["last"], state["maximized"])
                        return True
                    if not state["minimized"] and not state["maximized"]:
                        _snapshot()
                    _save_geometry(state["last"], state["maximized"])
                    tray.need_restore = False
                    _log("关闭窗口：收进托盘继续运行（托盘菜单可退出）")
                    # closing 回调里直接 hide 可能嵌套派发，延后一帧更稳
                    timer = threading.Timer(0.05, _hide_to_tray)
                    timer.daemon = True
                    timer.start()
                    return False  # 取消本次关闭

                try:
                    window.events.minimized += lambda *_: _on_minimized()
                    window.events.maximized += lambda *_: _on_maximized()
                    window.events.restored += lambda *_: _on_restored()
                    window.events.resized += lambda *_: _snapshot()
                    window.events.moved += lambda *_: _snapshot()
                    window.events.closing += lambda *_: _on_closing()
                except Exception:
                    _log("窗口事件注册不完整：\n" + traceback.format_exc())

                _log("主窗口已创建")
            except Exception:
                _log("主窗口初始化失败：\n" + traceback.format_exc())

        _log("WebView2 界面已启动")
        webview.start(boot, icon=str(_asset("app.ico") or ""), private_mode=False,
                      storage_path=str(storage))
        _log("窗口已关闭")
        return True
    except Exception:
        _log("WebView2 启动失败，准备回退：\n" + traceback.format_exc())
        return False


# ---------------------------------------------------------------- 服务与兜底

_server_holder = {"server": None}


def server_start() -> None:
    server = uvicorn.Server(uvicorn.Config(backend.app, host=backend.HOST,
                                           port=backend.PORT, log_level="warning"))
    _server_holder["server"] = server

    def _run() -> None:
        try:
            server.run()
        except Exception:
            _log("本地服务启动失败：\n" + traceback.format_exc())

    threading.Thread(target=_run, daemon=True).start()


def server_thread_start() -> None:
    """供闪屏 boot 线程调用（避免重复启动）。"""
    if _server_holder["server"] is None:
        server_start()


def server_stop() -> None:
    server = _server_holder["server"]
    if server is not None:
        try:
            server.should_exit = True
        except Exception:
            pass


def _wait_ready(timeout: float = 40.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(URL, timeout=1)
            return True
        except Exception:
            time.sleep(0.2)
    return False


def _fallback_edge_window() -> None:
    """回退：Edge/Chrome app 模式窗口 + 前端心跳判断关窗。"""
    _log("回退到浏览器 app 模式窗口")
    backend._open_browser(backend.PORT)
    first_wait = time.time()
    defer_logged = False
    while True:
        time.sleep(2)
        if not backend.LAST_PING["seen"] and time.time() - first_wait > backend.FIRST_PING_TIMEOUT:
            _log(f"等待界面连接超过 {backend.FIRST_PING_TIMEOUT}s，退出")
            break
        if backend.LAST_PING["seen"] and time.time() - backend.LAST_PING["t"] > backend.IDLE_EXIT_SECONDS:
            if backend.has_running_job():
                # 浏览器外壳无法拦截关窗；转换中暂缓退出，等结果落盘后再走空闲退出
                if not defer_logged:
                    _log("界面已关闭但转换仍在进行，暂缓退出直到任务完成")
                    defer_logged = True
                continue
            _log(f"前端心跳超过 {backend.IDLE_EXIT_SECONDS}s 未更新，判定窗口已关闭，退出")
            break


def main() -> None:
    core.ensure_streams()
    core.ensure_utf8_console()
    _log("=" * 40 + " 启动（WebView2 桌面窗口）")

    if not core.ensure_single_instance():
        # 已有实例：优先唤起它的主窗口（可能正收在托盘里）
        if _signal_existing_instance():
            _log("检测到已有实例，已请求唤起其主窗口")
        else:
            core.show_message("程序已在运行中。\n请在任务栏或托盘区找到已打开的「肥鸭文档转换」。")
        return
    _create_activate_event()

    try:
        cleaned = backend.cleanup_stale_temp_dirs()
        if cleaned:
            _log(f"已清理过期临时目录 {cleaned} 个")
    except Exception:
        pass

    try:
        import webview  # noqa: F401

        has_webview = True
    except Exception:
        has_webview = False
        _log("未安装 pywebview：" + traceback.format_exc())

    try:
        if has_webview:
            opened = _run_webview()
        else:
            server_start()
            opened = False
    except Exception:
        _log("启动流程异常：\n" + traceback.format_exc())
        opened = False

    if not opened:
        if _server_holder["server"] is None:
            server_start()
        if not _wait_ready():
            core.show_message(
                "本地服务启动失败，请重试或重启电脑。\n\n排障日志：\n"
                f"{core.local_log_dir() / 'web_start.log'}",
                flags=0x10,
            )
            return
        _fallback_edge_window()

    server_stop()
    _log("程序退出")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        _log("未捕获异常：\n" + traceback.format_exc())
        raise
