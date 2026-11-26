# -*- coding: utf-8 -*-
"""v1.8.0 新增接口测试：结果表读取 / 结构化导出 / 路径白名单。

不依赖 OCR 模型与 runtime，直接构造一份产物工作簿验证接口行为。
"""
from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, r"D:\fatduck-ocr")

from openpyxl import Workbook

import web_app


def _make_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "清单页"
    ws.append(["序号", "单号", "商品", "数量"])
    ws.append([1, "YT0720065031287", "冰红茶", 12])
    ws.append([2, "YT0720065031288", "=危险公式", 6])

    aux = wb.create_sheet("待复核清单")
    aux.append(["优先级", "问题类型"])
    aux.append(["高", "粘连拆分"])

    quality = wb.create_sheet("拍摄质量")
    quality.append(["拍摄质量", "良好"])
    wb.save(path)


def test_result_table():
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "结果.xlsx"
        _make_workbook(out)
        web_app.ALLOWED_OUTPUTS.add(str(out.resolve()))

        data = web_app.result_table(path=str(out), sheet="", offset=0, limit=50)
        assert data["ok"] is True, data
        assert data["sheets"] == ["清单页", "待复核清单", "拍摄质量"], data["sheets"]
        assert data["data_sheets"] == ["清单页"], data["data_sheets"]
        assert data["sheet"] == "清单页", data
        assert data["total_rows"] == 3 and data["total_cols"] == 4, data
        assert data["rows"][0] == ["序号", "单号", "商品", "数量"], data["rows"][0]
        assert data["rows"][1][1] == "YT0720065031287", data["rows"][1]
        assert data["rows"][2][3] == "6", data["rows"][2]
        assert data["truncated"] is False, data

        aux = web_app.result_table(path=str(out), sheet="待复核清单")
        assert aux["sheet"] == "待复核清单", aux
        assert aux["rows"][1][0] == "高", aux["rows"]

        paged = web_app.result_table(path=str(out), sheet="清单页", offset=1, limit=2)
        assert paged["offset"] == 1 and len(paged["rows"]) == 2, paged
        assert paged["truncated"] is False, paged

    print("PASS: 结果表读取（数据页/附表/分页）")


def test_export_structured():
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "结果.xlsx"
        _make_workbook(out)
        web_app.ALLOWED_OUTPUTS.add(str(out.resolve()))

        res = web_app.export_structured(path=str(out), fmt="json")
        payload = json.loads(bytes(res.body).decode("utf-8"))
        assert list(payload["sheets"].keys()) == ["清单页"], payload.keys()
        assert payload["sheets"]["清单页"][1][1] == "YT0720065031287", payload
        assert "待复核清单" not in payload["sheets"], payload["sheets"]
        disposition = res.headers.get("content-disposition", "")
        assert "filename*=UTF-8''" in disposition, disposition

        csv_res = web_app.export_structured(path=str(out), fmt="csv")
        text = bytes(csv_res.body).decode("utf-8")
        assert text.startswith("\ufeff"), "CSV 应带 BOM"
        rows = text.lstrip("\ufeff").splitlines()
        assert rows[0].startswith("序号,单号,商品,数量"), rows[0]
        # 以 = 开头的单元格必须被转义，避免被 Excel 当公式执行
        assert "'=危险公式" in text, text

        bad = web_app.export_structured(path=str(out), fmt="xml")
        assert getattr(bad, "status_code", 200) == 400, "不支持的格式应拒绝"

    print("PASS: 结构化导出（JSON/CSV、公式转义、格式校验）")


def test_path_guard():
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "外部.xlsx"
        _make_workbook(out)
        # 未登记到白名单的路径一律拒绝
        res = web_app.result_table(path=str(out))
        assert getattr(res, "status_code", 200) == 404, "白名单外路径必须拒绝"

        not_xlsx = Path(tmp) / "说明.txt"
        not_xlsx.write_text("x", encoding="utf-8")
        web_app.ALLOWED_OUTPUTS.add(str(not_xlsx.resolve()))
        res2 = web_app.result_table(path=str(not_xlsx))
        assert getattr(res2, "status_code", 200) == 404, "非 xlsx 必须拒绝"

    print("PASS: 结果表路径白名单")


if __name__ == "__main__":
    test_result_table()
    test_export_structured()
    test_path_guard()
    print("ALL PASS")

