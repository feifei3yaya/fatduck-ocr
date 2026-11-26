# -*- coding: utf-8 -*-
"""批次 0 安全回归：路径穿越、XSS 转义、跨站 Host/Origin、token、路径白名单。

直接跑：python tests/security_guard_test.py（不依赖 pytest）。
"""
from __future__ import annotations

import asyncio
import json
import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from fastapi.responses import FileResponse, JSONResponse
from fastapi.testclient import TestClient

import job_store
import web_app

BASE = f"http://{web_app.HOST}:{web_app.PORT}"
INDEX = "/web/index.html"


def _client(with_token: bool = True) -> TestClient:
    client = TestClient(web_app.app, base_url=BASE)
    if with_token:
        resp = client.get(INDEX)
        assert resp.status_code == 200, resp.status_code
        assert web_app.TOKEN_COOKIE in client.cookies, "首屏未下发 token cookie"
    return client


def test_index_sets_token_and_csp():
    client = TestClient(web_app.app, base_url=BASE)
    resp = client.get(INDEX)
    assert resp.status_code == 200
    cookie = resp.headers.get("set-cookie", "")
    assert web_app.TOKEN_COOKIE in cookie, cookie
    assert "HttpOnly" in cookie and "SameSite=strict" in cookie.replace("samesite", "SameSite"), cookie
    csp = resp.headers.get("content-security-policy", "")
    assert "script-src 'self'" in csp and "object-src 'none'" in csp, csp
    print("PASS: 首屏下发 HttpOnly/SameSite token cookie 与 CSP")


def test_api_requires_token():
    anon = TestClient(web_app.app, base_url=BASE)
    for path in ("/api/config", "/api/jobs", "/api/default-output"):
        resp = anon.get(path)
        assert resp.status_code == 403, (path, resp.status_code)
    resp = anon.post("/api/cancel", json={"job_id": "x"})
    assert resp.status_code == 403, resp.status_code

    ok = _client()
    assert ok.get("/api/config").status_code == 200
    # Header 方式同样可用（工具脚本/自检）
    resp = anon.get("/api/config", headers={"X-Img2Doc-Token": web_app.API_TOKEN})
    assert resp.status_code == 200, resp.status_code
    print("PASS: 无 token 的 API 请求全部 403，带 token 正常")


def test_host_and_origin_guard():
    client = _client()
    resp = client.get("/api/config", headers={"Host": "evil.example"})
    assert resp.status_code == 421, resp.status_code

    resp = client.post("/api/cancel", json={"job_id": "x"},
                       headers={"Origin": "http://evil.example"})
    assert resp.status_code == 403, resp.status_code

    resp = client.get("/api/config", headers={"Origin": BASE})
    assert resp.status_code == 200, resp.status_code
    print("PASS: 非法 Host 421、跨站 Origin 403、同源 Origin 放行")


def test_fid_and_filename_traversal_rejected():
    # 纯函数：文件标识白名单
    assert web_app._safe_fid("abc_1-2", 0) == "abc_1-2"
    assert web_app._safe_fid(None, 3) == "i3"
    for bad in ("../evil", "..\\..\\evil", "C:\\Windows\\Temp\\x", "a/b", "a" * 65,
                "x y"):
        try:
            web_app._safe_fid(bad, 0)
        except ValueError:
            pass
        else:
            raise AssertionError(f"非法 fid 未被拒绝：{bad!r}")

    # 纯函数：上传文件名含路径分隔符/空字节一律拒绝，正常中文名可用
    for bad in ("..\\..\\evil.exe", "../../evil.png", "C:\\Windows\\a.png", "\0.png", ""):
        try:
            web_app._safe_upload_name(0, bad)
        except ValueError:
            pass
        else:
            raise AssertionError(f"非法上传文件名未被拒绝：{bad!r}")
    assert web_app._safe_upload_name(0, "中文 名称(1).PNG") == "0000.png"
    assert web_app._safe_upload_name(1, "noext") == "0001"

    # 接口层：恶意 id / 文件名返回 400，临时目录外无写入
    client = _client()
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp) / "up"
        root.mkdir()
        out = Path(tmp) / "out.xlsx"
        marker = Path(tmp) / "pwned.txt"
        with patch.object(web_app.tempfile, "mkdtemp", return_value=str(root)):
            for payload in (
                {"id": "../..\\pwned", "upload": 0, "name": "a.png"},
                {"id": "C:\\Windows\\Temp\\pwned", "upload": 0, "name": "a.png"},
            ):
                resp = client.post(
                    "/api/convert",
                    data={"items": json.dumps([payload]), "mode": "full",
                          "output": str(out)},
                    files={"files": ("a.png", b"not-an-image", "image/png")},
                )
                assert resp.status_code == 400, (payload, resp.status_code, resp.text)
            resp = client.post(
                "/api/convert",
                data={"items": json.dumps([{"id": "ok1", "upload": 0,
                                            "name": "a.png"}]),
                      "mode": "full", "output": str(out)},
                files={"files": ("..\\..\\pwned.exe", b"x", "application/octet-stream")},
            )
            assert resp.status_code == 400, (resp.status_code, resp.text)
        assert not marker.exists() and not out.exists(), "穿越请求产生了目录外写入"
    print("PASS: fid / 文件名穿越全部 400，且无目录外写入")


def test_reveal_download_whitelist():
    saved = set(web_app.ALLOWED_OUTPUTS)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        allowed = tmp_path / "result.xlsx"
        allowed.write_bytes(b"fake")
        evil = tmp_path / "calc.exe"
        evil.write_bytes(b"MZ")
        other = tmp_path / "other.xlsx"
        other.write_bytes(b"fake")

        try:
            # 未登记 → 下载 404
            resp = asyncio.run(web_app.download(str(other)))
            assert isinstance(resp, JSONResponse) and resp.status_code == 404

            web_app.ALLOWED_OUTPUTS.add(str(allowed.resolve()))
            resp = asyncio.run(web_app.download(str(allowed)))
            assert isinstance(resp, FileResponse), resp
            resp = asyncio.run(web_app.download(str(other)))
            assert isinstance(resp, JSONResponse) and resp.status_code == 404

            # reveal：任意可执行文件拒绝；产物目录/文件允许；open 限制扩展名
            assert web_app.reveal({"path": str(evil), "mode": "open"})["ok"] is False
            calls: list[str] = []
            with patch.object(web_app.os, "startfile", side_effect=calls.append):
                assert web_app.reveal({"path": str(allowed), "mode": "open"})["ok"] is True
                assert calls == [str(allowed)]
                calls.clear()
                assert web_app.reveal({"path": str(tmp_path), "mode": "open"})["ok"] is True
                assert calls == [str(tmp_path)]
            # 已登记产物但扩展名不可打开（伪装成白名单路径的脚本）
            script = tmp_path / "note.ps1"
            script.write_text("echo hi", encoding="utf-8")
            web_app.ALLOWED_OUTPUTS.add(str(script.resolve()))
            assert web_app.reveal({"path": str(script), "mode": "open"})["ok"] is False
            assert web_app.reveal({"path": str(script), "mode": "select"})["ok"] is True
        finally:
            web_app.ALLOWED_OUTPUTS.clear()
            web_app.ALLOWED_OUTPUTS.update(saved)
    print("PASS: download/reveal 仅白名单路径可用，可执行文件不会被打开")


def test_ws_token_required():
    class FakeWS:
        def __init__(self, host: str, origin: str = "", token: str = ""):
            self.headers = {"host": host, "origin": origin}
            self.cookies = {web_app.TOKEN_COOKIE: token} if token else {}

    good_host = f"{web_app.HOST}:{web_app.PORT}"
    good_origin = f"http://{web_app.HOST}:{web_app.PORT}"
    assert web_app._ws_authorized(FakeWS(good_host, good_origin, web_app.API_TOKEN))
    # 无 token / 跨站 Origin / 非法 Host 一律拒绝
    assert not web_app._ws_authorized(FakeWS(good_host, good_origin))
    assert not web_app._ws_authorized(
        FakeWS(good_host, "http://evil.example", web_app.API_TOKEN))
    assert not web_app._ws_authorized(
        FakeWS("evil.example:8765", good_origin, web_app.API_TOKEN))
    assert not web_app._ws_authorized(FakeWS(good_host, good_origin, "wrong"))
    print("PASS: WebSocket 无 token / 跨站 / 非法 Host 全部拒绝")


if __name__ == "__main__":
    script = Path(__file__).resolve().parent.parent / "web" / "app.js"
    if not script.is_file():
        raise SystemExit("找不到 web/app.js")
    if "${f.name}" in script.read_text(encoding="utf-8"):
        raise SystemExit("FAIL: web/app.js 仍存在未转义的文件名插值")
    test_index_sets_token_and_csp()
    test_api_requires_token()
    test_host_and_origin_guard()
    test_fid_and_filename_traversal_rejected()
    test_reveal_download_whitelist()
    test_ws_token_required()
    print("\n全部安全用例通过")

