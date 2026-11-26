# -*- coding: utf-8 -*-
"""Web 任务集成测试：模拟 OCR，验证「全量识别」「纯单号」两种导出模式。

不依赖模型和 runtime，只验证任务编排、来源元数据、导出与完成事件。
"""
from __future__ import annotations

import asyncio
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, r"D:\fatduck-ocr")

from openpyxl import load_workbook

import web_app
import job_store
from md2excel import ExactPage
from table_rebuild import TableBlock


def _fake_page(image_path, worker, on_log=None, should_stop=None):
    table = TableBlock(
        rows=[
            ["日期", "承运商", "运单号", "商品名称", "实收数", "短少数"],
            ["2026.10.28", "YTO", "YT0720065031287", "康师傅冰红茶", "12", "0"],
            ["2026.10.28", "YTO", "YT0720065031288", "统一绿茶", "6", "1"],
        ],
        bbox=(0, 0, 100, 100),
    )
    return ExactPage(name=Path(image_path).stem, blocks=[table])


def _fake_failed_page(image_path, worker, on_log=None, should_stop=None):
    return ExactPage(
        name=Path(image_path).stem,
        blocks=[],
        warnings=[f"识别失败：无法解码图片: {image_path}"],
    )


class _FakeWorker:
    def __init__(self, *args, **kwargs):
        pass

    def close(self):
        pass


def _drain(job):
    events = []
    while True:
        event = job.q.get_nowait()
        if event.get("type") == "_end":
            break
        events.append(event)
    return events


def test_full_job():
    """全量识别：图片里的表格原样落地，第一页就是数据页，单号索引放在最后。"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "全量结果.xlsx"
        job = web_app.Job()
        files = [("f1", Path(tmp) / "f1_退仓表_001.jpg", "退仓表_001.jpg", "退仓表_001")]
        with patch.object(web_app, "_resolve_ocr_python", return_value=Path("python")), \
             patch.object(web_app, "OcrWorker", _FakeWorker), \
             patch.object(web_app, "build_exact_page", _fake_page):
            web_app._run_job(job, files, "full", out, "j-full")

        done = [e for e in _drain(job) if e.get("type") == "done"][-1]
        assert done["mode"] == "full", done
        assert done["searchable"] is True, done
        assert done["waybills"] == 2, done
        assert done["confirmed_waybills"] == 2, done
        assert out.is_file(), "结果文件没有生成"

        wb = load_workbook(out)
        assert wb.sheetnames[0] == "退仓表_001", wb.sheetnames
        assert "待复核清单" in wb.sheetnames, wb.sheetnames
        assert "待复核-低优先" in wb.sheetnames, wb.sheetnames
        assert wb.sheetnames[-1] == "拍摄质量", wb.sheetnames
        assert "单号索引" not in wb.sheetnames, wb.sheetnames
        assert wb.active.title == "退仓表_001", "打开时应先看到数据页"
        ws = wb["退仓表_001"]
        assert ws.cell(1, 1).value == "日期"
        assert ws.cell(2, 3).value == "YT0720065031287", ws.cell(2, 3).value
        assert ws.cell(3, 3).value == "YT0720065031288", ws.cell(3, 3).value

        # 全量识别仍保留应用内单号查询（数据页位置可定位）
        search = asyncio.run(web_app.search_waybills(q="YT0720065031288", job_id="j-full"))
        assert search["ok"] and search["total"] == 1, search
        assert search["matches"][0]["sheet"] == "退仓表_001", search["matches"][0]
        assert search["matches"][0]["cell"] == "C3", search["matches"][0]

    print("PASS: 全量识别原样导出数据页、done 事件与单号查询正确")


def test_waybill_job():
    """纯单号：整份 Excel 只有一个「单号」工作表，一列且自动去重。"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "纯单号.xlsx"
        job = web_app.Job()
        files = [("f1", Path(tmp) / "f1_退仓表_001.jpg", "退仓表_001.jpg", "退仓表_001")]
        with patch.object(web_app, "_resolve_ocr_python", return_value=Path("python")), \
             patch.object(web_app, "OcrWorker", _FakeWorker), \
             patch.object(web_app, "build_exact_page", _fake_page):
            web_app._run_job(job, files, "waybill", out, "j-waybill")

        done = [e for e in _drain(job) if e.get("type") == "done"][-1]
        assert done["mode"] == "waybill", done
        assert done["searchable"] is False, done
        assert done["unique_waybills"] == 2, done
        assert out.is_file(), "结果文件没有生成"

        wb = load_workbook(out)
        assert wb.sheetnames == ["单号"], wb.sheetnames
        ws = wb["单号"]
        assert ws.cell(1, 1).value == "单号"
        assert ws.cell(2, 1).value == "YT0720065031287", ws.cell(2, 1).value
        assert ws.cell(3, 1).value == "YT0720065031288", ws.cell(3, 1).value
        assert ws.max_column == 1, ws.max_column

        expired = asyncio.run(web_app.search_waybills(q="YT", job_id="j-waybill"))
        assert expired["ok"] is False, expired

    print("PASS: 纯单号只输出一列去重单号，不写索引、不提供查询")


def test_duplicate_and_invalid():
    """纯单号去重：重复单号只保留第一个；格式异常项不写入。"""
    from md2excel import convert_waybill
    from table_rebuild import TableBlock as TB

    page = ExactPage(
        name="重复表",
        blocks=[TB(rows=[
            ["日期", "承运商", "运单号", "商品名称", "实收数"],
            ["2026.10.28", "YTO", "YT0720065031287", "红茶", "12"],
            ["2026.10.28", "YTO", "YT0720065031287", "红茶", "6"],
            ["2026.10.28", "YTO", "YT07200650312870", "绿茶", "3"],
        ], bbox=(0, 0, 100, 100))],
        source_name="重复表.jpg",
        source_page=1,
    )
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "去重.xlsx"
        result = convert_waybill([page], out)
        assert result.unique_waybills == 1, result.unique_waybills
        assert result.duplicate_waybills == 1, result.duplicate_waybills
        assert result.format_invalid_waybills == 1, result.format_invalid_waybills
        wb = load_workbook(out)
        ws = wb["单号"]
        assert ws.max_row == 2, ws.max_row
        assert ws.cell(2, 1).value == "YT0720065031287", ws.cell(2, 1).value

    print("PASS: 纯单号去重、异常单号排除正确")


def test_failed_page_status():
    """解码失败的空页不得先报成功，再仅以任务级 error 收尾。"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "失败结果.xlsx"
        job = web_app.Job()
        files = [("bad1", Path(tmp) / "bad1_坏文件.txt", "坏文件.txt", "图片1")]
        with patch.object(web_app, "_resolve_ocr_python", return_value=Path("python")), \
             patch.object(web_app, "OcrWorker", _FakeWorker), \
             patch.object(web_app, "build_exact_page", _fake_failed_page):
            web_app._run_job(job, files, "full", out, "j-failed")

        events = _drain(job)
        statuses = [e for e in events if e.get("type") == "file_status"]
        assert statuses[-1]["tag"] == "fail", statuses
        assert all(e.get("tag") != "ok" for e in statuses), statuses
        logs = [e.get("text", "") for e in events if e.get("type") == "log"]
        assert not any("识别完成" in text for text in logs), logs
        assert events[-1]["type"] == "error", events
        assert not out.exists(), "失败任务不应生成 Excel"

    print("PASS: 解码失败空页只报失败，不发送识别完成或成功状态")


if __name__ == "__main__":
    # 任务快照落临时目录，测试不污染真实 AppData 的任务历史
    with tempfile.TemporaryDirectory() as tmp_jobs:
        with patch.object(job_store, "jobs_root", return_value=Path(tmp_jobs)):
            test_full_job()
            test_waybill_job()
            test_duplicate_and_invalid()
            test_failed_page_status()

