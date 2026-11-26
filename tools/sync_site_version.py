# -*- coding: utf-8 -*-
"""把官网 index.html 的下载直链与版本标注同步为 VERSION 中的版本号。

背景：官网下载按钮用直链（点击即下载），版本号写死在链接里，发版时容易漏改。
用法：
    python tools/sync_site_version.py           # 按 VERSION 改写官网链接
    python tools/sync_site_version.py --check   # 只检查，不一致时退出码 1（供 CI/发版前校验）
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / "index.html"
VERSION = (ROOT / "VERSION").read_text(encoding="utf-8").strip()

LINK_RE = re.compile(
    r"releases/download/v[\d.]+/fatduck-ocr_(installer|portable)_v?[\d.]+\.(exe|zip)")
CAPTION_RE = re.compile(r'(<span>Windows 桌面版 · 图片转 Excel</span><span>)v[\d.]+(</span>)')


def desired(html: str) -> str:
    html = LINK_RE.sub(
        lambda m: f"releases/download/v{VERSION}/fatduck-ocr_{m.group(1)}_v{VERSION}.{m.group(2)}",
        html)
    return CAPTION_RE.sub(lambda m: f"{m.group(1)}v{VERSION}{m.group(2)}", html)


def main() -> int:
    check_only = "--check" in sys.argv[1:]
    src = SITE.read_text(encoding="utf-8")
    out = desired(src)
    if out == src:
        print(f"官网版本一致：v{VERSION}")
        return 0
    if check_only:
        old = sorted(set(LINK_RE.findall(src)))
        print(f"官网版本不一致，VERSION={VERSION}，当前链接：{old}")
        print("请运行 python tools/sync_site_version.py 后提交。")
        return 1
    SITE.write_text(out, encoding="utf-8")
    print(f"已把官网下载链接与版本标注同步为 v{VERSION}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
