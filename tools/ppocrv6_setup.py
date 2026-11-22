# -*- coding: utf-8 -*-
"""PP-OCRv6 ONNX -> OpenVINO IR 转换 + 字库导出。

需用带 ONNX 前端的 python 运行（runtime 里该前端已被裁掉）：
    & "$env:USERPROFILE\.openvino\venv\mineru\Scripts\python.exe" tools\ppocrv6_setup.py

产物：models\PP-OCRv6\det\rec 的 xml+bin、assets\ppocrv6_dict.txt
"""
import sys
from pathlib import Path

PROJ = Path(__file__).resolve().parent.parent
BASE = PROJ / "models" / "PP-OCRv6"

import openvino as ov


def convert(onnx, xml):
    if xml.is_file() and xml.with_suffix(".bin").is_file():
        print(f"已存在，跳过: {xml}")
        return
    if not onnx.is_file():
        print(f"缺少 ONNX 权重: {onnx}")
        sys.exit(1)
    xml.parent.mkdir(parents=True, exist_ok=True)
    model = ov.convert_model(str(onnx))
    ov.save_model(model, str(xml))
    print(f"转换完成: {xml.name} "
          f"{xml.stat().st_size / 1024:.0f}KB + "
          f"{xml.with_suffix('.bin').stat().st_size / 1048576:.1f}MB")


def dump_dict():
    yml = BASE / "rec" / "inference.yml"
    chars = []
    try:
        import yaml
        cfg = yaml.safe_load(yml.read_text(encoding="utf-8"))
        chars = list(cfg["PostProcess"]["character_dict"])
    except Exception as exc:
        print(f"yaml 解析失败({exc})，改用行解析")
    if not chars:
        started = False
        for line in yml.read_text(encoding="utf-8").splitlines():
            s = line.strip()
            if s.startswith("character_dict:"):
                started = True
                continue
            if not started:
                continue
            if not s.startswith("- "):
                break
            v = s[2:].strip()
            if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"":
                v = v[1:-1].replace("''", "'")
            chars.append(v)
    out = PROJ / "assets" / "ppocrv6_dict.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(chars), encoding="utf-8")
    print(f"字库导出: {out} 共 {len(chars)} 项")


if __name__ == "__main__":
    convert(BASE / "det" / "inference.onnx", BASE / "det" / "det.xml")
    convert(BASE / "rec" / "inference.onnx", BASE / "rec" / "rec.xml")
    dump_dict()
