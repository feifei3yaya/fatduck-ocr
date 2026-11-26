# -*- coding: utf-8 -*-
"""图片命名功能轻量验证（函数级直调，不起服务、不跑 OCR 全量回归）。

复用 file_ops_test 的方式：import web_app，临时目录造假文件，直接登记
ALLOWED_FILES（最简合法路径），asyncio.run 调接口函数。覆盖：
  1) 默认 label 进 convert（plan -> saved -> _run_job）与 done 增量字段
  2) rename 磁盘改名 + 重名 (n) + undo 还原
  3) 未登记路径 400
  4) sheet 重写：合成 page 走 md2excel.convert_exact 生成 xlsx，再 rename 验证表名
"""
from __future__ import annotations

import asyncio
import json
import sys
import tempfile
import time
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, r"D:\fatduck-ocr")

import web_app  # noqa: E402
import job_store  # noqa: E402
from md2excel import ExactPage, convert_exact  # noqa: E402
from table_rebuild import TableBlock, TextBlock  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402

PASS, FAIL = 0, 0


def check(name, cond, extra=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  [PASS] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name} {extra}")


def _table(rows):
    return TableBlock(kind="table", rows=rows)


def test_convert_label_flow(tmp):
    print("[1] 默认 label 进 convert")
    src = tmp / "a.jpg"
    src.write_bytes(b"\xff\xd8\xff\xe0fake-jpeg")
    captured = {}

    def fake_run_job(job, files, mode, out_path, job_id="", target="",
                     layout="per_image", extra_temp_dirs=None):
        captured["files"] = files
        job.done = True
        job.q.put({"type": "_end"})

    plans = [
        {"id": "f1", "name": "a.jpg", "label": "图片1", "path": str(src)},
        {"id": "f2", "name": "a.jpg", "path": str(src)},  # 无 label -> 回落原名去扩展名
    ]
    out = tmp / "flow.xlsx"
    with patch.object(web_app, "_run_job", fake_run_job):
        resp = asyncio.run(web_app.convert_start(
            files=None, ids=None, mode="full", output=str(out),
            items=json.dumps(plans), target="", layout="per_image"))
    for _ in range(200):
        if "files" in captured:
            break
        time.sleep(0.02)
    check("convert 启动成功", getattr(resp, "status_code", 200) == 200, repr(resp)[:120])
    files = captured.get("files") or []
    check("saved 携带 label", len(files) == 2 and files[0][3] == "图片1", repr(files))
    check("saved 保留原名", len(files) == 2 and files[0][2] == "a.jpg", repr(files))
    check("无 label 回落 stem", len(files) == 2 and files[1][3] == "a", repr(files))
    check("convert 登记磁盘路径", str(src) in web_app.ALLOWED_FILES)


def test_done_naming_fields():
    print("[2] done 增量字段 label / suggested_name")
    files = [
        ("f1", Path("a.jpg"), "a.jpg", "图片1"),
        ("f2", Path("b.jpg"), "b.jpg", "图片2"),
        ("f3", Path("c.jpg"), "c.jpg", "图片3"),
    ]
    p1 = ExactPage(name="图片1", blocks=[
        TextBlock(kind="text", lines=["某某快递运单清单"]),
        _table([["单号"], ["YT0718286485990"]])], source_name="a.jpg", source_page=1)
    p2 = ExactPage(name="图片2", blocks=[
        _table([["单号"], ["YT0718286485991"]])], source_name="b.jpg", source_page=1)
    p3 = ExactPage(name="图片3", blocks=[
        _table([["单号"], ["YT0718286485992"], ["YT0718286485993"]])],
        source_name="c.jpg", source_page=1)
    quality = {"per_file": [{"file": "a.jpg"}, {"file": "b.jpg"}, {"file": "c.jpg"}]}
    summary = [{"file": "a.jpg"}, {"file": "b.jpg"}, {"file": "c.jpg"}]
    web_app._attach_naming([p1, p2, p3], files, quality, summary)
    pf = quality["per_file"]
    check("标题建议=首行清洗+_001", pf[0]["suggested_name"] == "某某快递运单清单_001", pf[0])
    check("单号唯一建议=该单号", pf[1]["suggested_name"] == "YT0718286485991", pf[1])
    check("多单号回落默认 label", pf[2]["suggested_name"] == "图片3", pf[2])
    check("per_file 带 label", pf[0]["label"] == "图片1", pf[0])
    check("file_summary 带 label", summary[0]["label"] == "图片1", summary[0])


def test_rename_disk_and_undo(tmp):
    print("[3] 磁盘改名 + 重名 (n) + undo")
    a = tmp / "orig_a.jpg"
    b = tmp / "orig_b.jpg"
    a.write_bytes(b"aaa")
    b.write_bytes(b"bbb")
    web_app.ALLOWED_FILES.add(str(a))
    web_app.ALLOWED_FILES.add(str(b))
    items = [
        {"fid": "f1", "path": str(a), "old_label": "图片1", "new_name": "运单A", "rename_disk": True},
        {"fid": "f2", "path": str(b), "old_label": "图片2", "new_name": "运单A", "rename_disk": True},
    ]
    resp = web_app.rename_files({"items": items})
    results = resp.get("results", [])
    check("rename 返回 ok", resp.get("ok") is True, repr(resp)[:160])
    na = Path(results[0]["new_path"])
    nb = Path(results[1]["new_path"])
    check("第一个改名为 运单A.jpg", na.name == "运单A.jpg" and na.is_file(), results[0])
    check("重名自动 (1)", nb.name == "运单A (1).jpg" and nb.is_file(), results[1])
    check("原文件已不在", not a.exists() and not b.exists())
    check("新路径已登记", str(na) in web_app.ALLOWED_FILES and str(nb) in web_app.ALLOWED_FILES)
    check("可撤销", resp.get("undoable") is True)

    undo = web_app.rename_undo()
    check("undo ok", undo.get("ok") is True, repr(undo)[:160])
    check("磁盘还原 a", a.is_file() and not na.exists(), undo.get("disk"))
    check("磁盘还原 b", b.is_file() and not nb.exists(), undo.get("disk"))
    check("还原后重新登记", str(a) in web_app.ALLOWED_FILES and str(b) in web_app.ALLOWED_FILES)


def test_unregistered_400(tmp):
    print("[4] 未登记路径 400")
    c = tmp / "unregistered.jpg"
    c.write_bytes(b"ccc")
    resp = web_app.rename_files({"items": [
        {"fid": "f9", "path": str(c), "old_label": "图片9", "new_name": "黑客", "rename_disk": True}]})
    check("未登记路径拒绝 400", getattr(resp, "status_code", 200) == 400, repr(resp)[:160])
    check("文件未被改动", c.is_file() and c.name == "unregistered.jpg")
    # 无路径（拖入/粘贴）只改显示名，不报 400
    resp2 = web_app.rename_files({"items": [
        {"fid": "f8", "path": "", "old_label": "图片8", "new_name": "拖入图", "rename_disk": True}]})
    r2 = resp2.get("results", [{}])[0]
    check("无路径仅改显示名", resp2.get("ok") is True and r2.get("disk_renamed") is False, repr(resp2)[:160])
    check("无路径返回原因", "原图位置未知" in (r2.get("reason") or ""), r2)


def test_sheet_rewrite(tmp):
    print("[5] sheet 重写（合成 page -> convert_exact -> rename）")
    pages = [
        ExactPage(name="图片1", blocks=[_table([["单号", "数量"], ["YT0718286485990", "1"]])],
                  source_name="a.jpg", source_page=1),
        ExactPage(name="图片2", blocks=[_table([["单号"], ["YT0718286485991"]])],
                  source_name="b.jpg", source_page=1),
        ExactPage(name="图片3-第1页", blocks=[_table([["单号"], ["YT0718286485992"]])],
                  source_name="c.tif", source_page=1),
        ExactPage(name="图片3-第2页", blocks=[_table([["单号"], ["YT0718286485993"]])],
                  source_name="c.tif", source_page=2),
    ]
    out = tmp / "sheets.xlsx"
    convert_exact(pages, out, include_index=False, include_aux=False, include_flags=True)
    web_app.ALLOWED_OUTPUTS.add(str(out.resolve()))
    before = load_workbook(out).sheetnames
    check("初始数据表名用 label", "图片1" in before and "图片3-第2页" in before, before)

    resp = web_app.rename_files({
        "job_id": "jtest", "out_path": str(out),
        "items": [
            {"fid": "f1", "path": "", "old_label": "图片1", "new_name": "运单A", "rename_disk": False},
            {"fid": "f3", "path": "", "old_label": "图片3", "new_name": "多页图", "rename_disk": False},
        ]})
    check("sheet rename ok", resp.get("ok") is True, repr(resp)[:160])
    after = load_workbook(out).sheetnames
    check("单页表名已改", "运单A" in after and "图片1" not in after, after)
    check("多页后缀保留", "多页图-第1页" in after and "多页图-第2页" in after, after)
    check("未命名表不动", "图片2" in after, after)
    check("辅助表不动", "待复核清单" in after and "待复核-低优先" in after, after)
    check("单元格不动", load_workbook(out)["运单A"].cell(row=1, column=1).value is not None)

    undo = web_app.rename_undo()
    restored = load_workbook(out).sheetnames
    check("undo 还原表名", "图片1" in restored and "图片3-第1页" in restored and "运单A" not in restored, restored)


def test_undo_persist(tmp):
    print("[6] 撤销历史持久化（模拟重启）")
    hist = tmp / "rename_history.json"
    web_app._RENAME_HISTORY.clear()
    with patch.object(job_store, "rename_history_path", return_value=hist):
        a = tmp / "p_a.jpg"
        a.write_bytes(b"aaa")
        web_app.ALLOWED_FILES.add(str(a))
        resp = web_app.rename_files({"items": [
            {"fid": "f1", "path": str(a), "old_label": "图片1", "new_name": "改名后", "rename_disk": True}]})
        new_path = Path(resp["results"][0]["new_path"])
        check("改名后历史落盘", hist.is_file())
        data = json.loads(hist.read_text(encoding="utf-8"))
        check("落盘含 disk 记录", bool(data) and data[0]["disk"][0]["new"] == str(new_path), data)

        # 模拟程序重启：内存历史与白名单清空，只剩磁盘上的历史文件
        web_app._RENAME_HISTORY.clear()
        web_app.ALLOWED_FILES.discard(str(new_path))
        web_app._load_rename_history()
        state = web_app.rename_history_state()
        check("重启后仍可撤销", state.get("can_undo") is True, state)
        check("重启后改名路径重新登记", str(new_path) in web_app.ALLOWED_FILES)
        undo = web_app.rename_undo()
        check("重启后 undo 还原磁盘名", a.is_file() and not new_path.exists(), undo)
        check("撤销后历史清空落盘", json.loads(hist.read_text(encoding="utf-8")) == [])
        check("无历史时 can_undo False", web_app.rename_history_state().get("can_undo") is False)


def test_undo_persist_guard(tmp):
    print("[7] 持久化历史的防篡改校验")
    hist = tmp / "rename_history.json"
    web_app._RENAME_HISTORY.clear()
    other_dir = tmp / "other"
    other_dir.mkdir()
    victim = other_dir / "victim.jpg"
    victim.write_bytes(b"vvv")
    # 跨目录记录 + 结构残缺记录 + 正常记录混在一起
    hist.write_text(json.dumps([
        {"ts": 1, "disk": [{"old": str(tmp / "x.jpg"), "new": str(victim)}], "sheets": []},
        {"ts": 2, "disk": "not-a-list", "sheets": []},
        {"ts": 3, "disk": [], "sheets": [{"path": "", "old": "a", "new": "b"}]},
    ], ensure_ascii=False), encoding="utf-8")
    web_app.ALLOWED_FILES.discard(str(victim))
    with patch.object(job_store, "rename_history_path", return_value=hist):
        web_app._load_rename_history()
    check("跨目录记录被丢弃", all(r.get("disk") != [{"old": str(tmp / "x.jpg"), "new": str(victim)}]
                                 for r in web_app._RENAME_HISTORY), web_app._RENAME_HISTORY)
    check("篡改路径不会被登记", str(victim) not in web_app.ALLOWED_FILES)
    check("无有效磁盘/表记录则不可撤销", web_app.rename_history_state().get("can_undo") is False)

    hist.write_text("{not json", encoding="utf-8")
    with patch.object(job_store, "rename_history_path", return_value=hist):
        web_app._RENAME_HISTORY.clear()
        web_app._load_rename_history()
    check("损坏文件不抛异常且为空", web_app._RENAME_HISTORY == [])


def test_undo_sheet_whitelist(tmp):
    print("[8] 撤销只改产物白名单内的工作簿")
    hist = tmp / "rename_history.json"
    other = tmp / "其他" / "book.xlsx"
    other.parent.mkdir()
    wb = Workbook()
    wb.active.title = "被改名"
    wb.save(other)
    hist.write_text(json.dumps(
        [{"ts": 1, "disk": [],
          "sheets": [{"path": str(other), "old": "原名", "new": "被改名"}]}],
        ensure_ascii=False), encoding="utf-8")
    with patch.object(job_store, "rename_history_path", return_value=hist):
        web_app._RENAME_HISTORY.clear()
        web_app._load_rename_history()
        check("结构合法则可撤销", web_app.rename_history_state()["can_undo"] is True)
        undo = web_app.rename_undo()
    check("白名单外工作表未被改回", load_workbook(other).sheetnames == ["被改名"],
          load_workbook(other).sheetnames)
    check("sheet_restored 为 0", undo.get("sheet_restored") == 0, undo)


def test_rename_syncs_job_meta(tmp):
    print("[9] 命名结果同步任务历史 file_summary.label")
    web_app._RENAME_HISTORY.clear()
    with tempfile.TemporaryDirectory() as tj, \
            patch.object(job_store, "jobs_root", return_value=Path(tj)), \
            patch.object(job_store, "rename_history_path", return_value=tmp / "rh9.json"):
        job_store.save_job("jsync", {
            "time": "t", "out_path": "", "files": ["a.jpg", "b.jpg"],
            "file_summary": [{"file": "a.jpg", "label": "图片1"},
                             {"file": "b.jpg", "label": "图片2"}]})
        resp = web_app.rename_files({"job_id": "jsync", "items": [
            {"fid": "f1", "path": "", "old_label": "图片1", "new_name": "运单A", "rename_disk": False}]})
        meta = job_store.load_job("jsync")
        labels = [e["label"] for e in meta["file_summary"]]
        check("历史 label 已同步", labels == ["运单A", "图片2"], labels)
        check("历史同步后可撤销", resp.get("undoable") is True, resp)
        check("原保存时间不被改写", meta.get("saved_at") == job_store.load_job("jsync").get("saved_at"))
        undo = web_app.rename_undo()
        check("undo 还原历史 label",
              [e["label"] for e in job_store.load_job("jsync")["file_summary"]] == ["图片1", "图片2"],
              undo)

        job_store.save_job("jold", {"file_summary": [{"file": "a.jpg"}]})
        web_app.rename_files({"job_id": "jold", "items": [
            {"fid": "f1", "old_label": "图片1", "new_name": "X", "rename_disk": False}]})
        check("老任务无 label 字段不被改写",
              job_store.load_job("jold")["file_summary"] == [{"file": "a.jpg"}])
        check("无有效改动不产生撤销批次", web_app.rename_history_state().get("can_undo") is False)


def test_title_suggestion_filters():
    print("[10] 智能建议标题过滤（日期/噪音行/回溯深度）")
    names = ["a.jpg", "b.jpg", "c.jpg", "d.jpg", "e.jpg"]
    files = [(f"f{i}", Path(n), n, f"图片{i}") for i, n in enumerate(names, 1)]

    def page(idx, blocks):
        return ExactPage(name=f"图片{idx}", blocks=blocks, source_name=names[idx - 1], source_page=1)

    pages = [
        page(1, [TextBlock(kind="text", lines=["圆通速递退仓表"]),
                 TextBlock(kind="text", lines=["2026.11.25"]),
                 _table([["单号"], ["YT0718286485990", "YT0718286485991"]])]),
        page(2, [TextBlock(kind="text", lines=["备注：请核对"]),
                 TextBlock(kind="text", lines=["打印页码：1/2"]),
                 _table([["单号"], ["YT0718286485993"]])]),
        page(3, [TextBlock(kind="text", lines=["很远处的标题"]),
                 TextBlock(kind="text", lines=["2026.11.25"]),
                 TextBlock(kind="text", lines=["2026.11.25"]),
                 TextBlock(kind="text", lines=["2026.11.25"]),
                 TextBlock(kind="text", lines=["2026.11.25"]),
                 _table([["单号"], ["YT0718286485994"]])]),
        page(4, [TextBlock(kind="text", lines=["打印时间：2026-11-25 14:58", "圆通退仓明细表"]),
                 _table([["单号"], ["YT0718286485995", "YT0718286485996"]])]),
        page(5, [_table([["单号"], ["YT0718286485997", "YT0718286485998"]])]),
    ]
    quality = {"per_file": [{"file": n} for n in names]}
    summary = [{"file": n} for n in names]
    web_app._attach_naming(pages, files, quality, summary)
    pf = quality["per_file"]
    check("日期行不当标题，向上取真标题", pf[0]["suggested_name"] == "圆通速递退仓表_001", pf[0])
    check("噪音行不当标题，回落唯一单号", pf[1]["suggested_name"] == "YT0718286485993", pf[1])
    check("标题超过 3 个文字块不回溯", pf[2]["suggested_name"] == "YT0718286485994", pf[2])
    check("多行块跳过打印时间取标题行", pf[3]["suggested_name"] == "圆通退仓明细表_004", pf[3])
    check("表格为首块无标题回落默认名", pf[4]["suggested_name"] == "图片5", pf[4])


def test_print_time_and_file_date(tmp):
    print("[11] 打印时间识别 + 文件日期写入 + 撤销还原")
    import datetime as _dt
    import os as _os
    names = ["a.jpg", "b.jpg", "c.pdf"]
    files = [("f1", Path("a.jpg"), "a.jpg", "图片1"),
             ("f2", Path("b.jpg"), "b.jpg", "图片2"),
             ("f3", Path("c.pdf"), "c.pdf", "图片3")]
    pages = [
        ExactPage(name="图片1", blocks=[
            _table([["单号"], ["YT0718286485990", "YT0718286485991"]]),
            TextBlock(kind="text", lines=["打印页码：1/2"]),
            TextBlock(kind="text", lines=["打印时间：2026-11-25 14:58"])], source_name="a.jpg", source_page=1),
        ExactPage(name="图片2", blocks=[
            _table([["单号"], ["YT0718286485992", "YT0718286485993"]])], source_name="b.jpg", source_page=1),
        ExactPage(name="图片3-第1页", blocks=[
            TextBlock(kind="text", lines=["打印时间: 2026.10.28 09:05:30"])], source_name="c.pdf", source_page=1),
        ExactPage(name="图片3-第2页", blocks=[
            TextBlock(kind="text", lines=["打印时间 2026-10-28"]),
            _table([["单号"], ["YT0718286485994", "YT0718286485995"]])], source_name="c.pdf", source_page=2),
    ]
    quality = {"per_file": [{"file": "a.jpg"}, {"file": "b.jpg"},
                            {"file": "c.pdf"}, {"file": "c.pdf"}]}
    summary = [{"file": n} for n in names]
    web_app._attach_naming(pages, files, quality, summary)
    pf = quality["per_file"]
    check("页脚打印时间被识别", pf[0].get("print_time") == "2026-11-25 14:58", pf[0])
    check("无打印时间为空串", pf[1].get("print_time") == "", pf[1])
    check("兼容点分隔/冒号/秒（归一到分）", pf[2].get("print_time") == "2026-10-28 09:05", pf[2])
    check("只有日期无时分", pf[3].get("print_time") == "2026-10-28", pf[3])
    check("file_summary 取首页打印时间", summary[2].get("print_time") == "2026-10-28 09:05", summary[2])
    check("file_summary 无打印时间为空串", summary[1].get("print_time") == "", summary[1])

    # 文件日期写入：os.utime 改修改时间，改名后再写，未识别到的不动
    a = tmp / "pt_a.jpg"
    a.write_bytes(b"aaa")
    old_epoch = 1_700_000_000
    _os.utime(a, (old_epoch, old_epoch))
    web_app.ALLOWED_FILES.add(str(a))
    web_app._RENAME_HISTORY.clear()
    with patch.object(job_store, "rename_history_path", return_value=tmp / "rh11.json"):
        resp = web_app.rename_files({"items": [
            {"fid": "f1", "path": str(a), "old_label": "图片1", "new_name": "运单A",
             "rename_disk": True, "set_date": True, "print_time": "2026-11-25 14:58"},
            {"fid": "f2", "path": "", "old_label": "图片2", "new_name": "无日期图",
             "rename_disk": False, "set_date": True, "print_time": ""},
            {"fid": "f3", "path": "", "old_label": "图片3", "new_name": "坏日期图",
             "rename_disk": False, "set_date": True, "print_time": "不是日期"},
        ]})
        r = resp["results"]
        new_path = Path(r[0]["new_path"])
        check("改名后在新路径上写日期", new_path.name == "运单A.jpg" and r[0].get("date_set") is True, r[0])
        check("文件修改时间 = 打印时间",
              _dt.datetime.fromtimestamp(new_path.stat().st_mtime) == _dt.datetime(2026, 11, 25, 14, 58),
              _dt.datetime.fromtimestamp(new_path.stat().st_mtime))
        check("未识别到打印时间如实提示", r[1].get("date_set") is False and "未识别到打印时间" in (r[1].get("reason") or ""), r[1])
        check("无效打印时间如实提示", r[2].get("date_set") is False and "格式无效" in (r[2].get("reason") or ""), r[2])
        hist = json.loads((tmp / "rh11.json").read_text(encoding="utf-8"))
        check("日期记录随历史落盘", bool(hist) and hist[0]["dates"][0]["old_mtime"] == old_epoch, hist)

        # 模拟重启后撤销：旧日期与旧文件名都要还原
        web_app._RENAME_HISTORY.clear()
        web_app.ALLOWED_FILES.discard(str(new_path))
        web_app._load_rename_history()
        undo = web_app.rename_undo()
    check("undo 返回日期还原数", undo.get("dates_restored") == 1, undo)
    check("undo 还原文件修改时间", a.is_file() and abs(a.stat().st_mtime - old_epoch) < 1, a.stat().st_mtime if a.is_file() else None)

    # 未登记路径 + set_date：与改名同一 fail-closed 口径
    victim = tmp / "victim.jpg"
    victim.write_bytes(b"vvv")
    _os.utime(victim, (old_epoch, old_epoch))
    resp = web_app.rename_files({"items": [
        {"fid": "f9", "path": str(victim), "old_label": "图片9", "new_name": "黑客",
         "rename_disk": False, "set_date": True, "print_time": "2026-11-25 14:58"}]})
    check("未登记路径改日期拒绝 400", getattr(resp, "status_code", 200) == 400, repr(resp)[:120])
    check("未登记文件日期未被改动", abs(victim.stat().st_mtime - old_epoch) < 1)


def test_ocr_boxes_fallback():
    print("[12] 几何回退路径：从 ocr_boxes 读标题与打印时间")
    page = ExactPage(name="图片1", blocks=[
        TableBlock(kind="table", rows=[["单号"], ["YT0718286485990", "YT0718286485991"]],
                   bbox=(0, 0, 1000, 600))
    ], ocr_boxes=[
        {"text": "圆通速递退仓表", "bbox": [10, 10, 300, 60]},
        {"text": "2026.11.25", "bbox": [10, 70, 200, 100]},
        {"text": "单号", "bbox": [10, 200, 100, 240]},
        {"text": "打印页码：1/2", "bbox": [10, 500, 200, 530]},
        {"text": "打印时间：2026-11-25 14:58", "bbox": [10, 550, 400, 590]},
    ])
    check("ocr_boxes 上方文字可作标题", web_app._page_title_text(page) == "圆通速递退仓表",
          web_app._page_title_text(page))
    check("ocr_boxes 可读页脚打印时间", web_app._page_print_time(page) == "2026-11-25 14:58",
          web_app._page_print_time(page))


def main():
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        test_convert_label_flow(tmp)
        test_done_naming_fields()
        test_rename_disk_and_undo(tmp)
        test_unregistered_400(tmp)
        test_sheet_rewrite(tmp)
        test_undo_persist(tmp)
        test_undo_persist_guard(tmp)
        test_undo_sheet_whitelist(tmp)
        test_rename_syncs_job_meta(tmp)
        test_title_suggestion_filters()
        test_print_time_and_file_date(tmp)
        test_ocr_boxes_fallback()
    print(f"\n结果：PASS={PASS} FAIL={FAIL}")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())

