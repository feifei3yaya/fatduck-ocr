# -*- coding: utf-8 -*-
"""excel-flag 测试：黄底高亮 + 待复核清单 sheet + result.flagged。"""
import sys
import tempfile
from pathlib import Path

PROJ = Path(r"D:\fatduck-ocr")
sys.path.insert(0, str(PROJ))

from md2excel import Block, ExactPage, convert_exact

PASS, FAIL = 0, 0


def check(name, cond, extra=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  [PASS] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name} {extra}")


tmp = Path(tempfile.mkdtemp(prefix="flag_test_"))
out = tmp / "r.xlsx"

print("== 用例 1：有 flags -> 黄底 + 分层清单 sheet ==")
flags = [
    {"row": 0, "col": 3, "rebuilt": "YT0718623650560签收", "ocr": "YT0718623650560",
     "zoom": "YT0718623650560", "reason": "单号与签收状态粘连，已拆分待复核"},
    {"row": 5, "col": 6, "rebuilt": "0", "ocr": "0", "zoom": "",
     "reason": "放大复识置信度仅 0.54，未自动改写"},
]
page = ExactPage(
    name="运单1.png",
    blocks=[Block(kind="table", rows=[
        ["1", "2026.10.26", "YTO", "YT0718623650601", "货名", "10", "2"],
        ["2", "2026.10.26", "YTO", "YT0718048735987", "货名", "8", "4"],
        ["3", "2026.10.26", "YTO", "YT0720059354356", "货名", "12", "0"],
        ["4", "2026.10.26", "YTO", "YT0718880632729", "货名", "10", "2"],
        ["5", "2026.10.26", "YTO", "YT0718931927851", "货名", "9", "3"],
        ["6", "2026.10.26", "YTO", "YT0719277689599", "货名", "12", "0"],
    ])],
    warnings=["待复核：第1行第4列「YT0718623650560」"],
    flags=flags,
)
res = convert_exact([page], out)
check("result.flagged == 2", res.flagged == 2, res.flagged)

from openpyxl import load_workbook
wb = load_workbook(out)
check("生成 4 个 sheet（数据+清单+低优先+索引）",
      wb.sheetnames == ["运单1", "待复核清单", "待复核-低优先", "单号索引"],
      wb.sheetnames)
ws = wb["运单1"]   # 数据页按名字取（单号索引在最后）
# flag row=0 col=3 -> sheet 第1行第4列；row=5 col=6 -> 第6行第7列
c1 = ws.cell(row=1, column=4)
c2 = ws.cell(row=6, column=7)
c3 = ws.cell(row=2, column=4)   # 未标记单元格
check("标记单元格黄底 FFF2CC",
      c1.fill.fgColor.rgb in ("00FFF2CC", "FFFFF2CC") and
      c2.fill.fgColor.rgb in ("00FFF2CC", "FFFFF2CC"),
      f"{c1.fill.fgColor.rgb} {c2.fill.fgColor.rgb}")
check("标记单元格字体色 9C6500",
      c1.font.color.rgb in ("009C6500", "FF9C6500"), c1.font.color.rgb)
check("未标记单元格无黄底",
      c3.fill.fgColor.rgb not in ("00FFF2CC", "FFFFF2CC"), c3.fill.fgColor.rgb)
check("标记单元格值保持重建结果", c1.value == "YT0718623650601", c1.value)

fl_ws = wb["待复核清单"]
check("清单表头 9 列",
      [fl_ws.cell(row=1, column=i).value for i in range(1, 10)] ==
      ["优先级", "问题类型", "文件", "行", "列", "重建值", "整图OCR", "放大OCR", "原因"])
check("主清单第 1 条=高优先粘连",
      fl_ws.cell(row=2, column=1).value == "高"
      and fl_ws.cell(row=2, column=2).value == "粘连拆分"
      and fl_ws.cell(row=2, column=3).value == "运单1.png"
      and fl_ws.cell(row=2, column=4).value == 1
      and fl_ws.cell(row=2, column=6).value == "YT0718623650560签收")
check("高优先单元格红底", fl_ws.cell(row=2, column=1).fill.fgColor.rgb
      in ("00C00000", "FFC00000"), fl_ws.cell(row=2, column=1).fill.fgColor.rgb)
check("主清单合计行核对总数",
      fl_ws.cell(row=3, column=1).value
      == "合计：高 1 条 + 中 0 条 = 1 条；另有低优先 1 条见「待复核-低优先」工作表；"
         "全部待复核共 2 条",
      fl_ws.cell(row=3, column=1).value)
check("清单冻结首行", fl_ws.freeze_panes == "A2")

low_ws = wb["待复核-低优先"]
check("低优先表头同为 9 列",
      low_ws.cell(row=1, column=1).value == "优先级"
      and low_ws.cell(row=1, column=9).value == "原因")
check("低优先说明行含总数核对",
      isinstance(low_ws.cell(row=2, column=1).value, str)
      and "共 2 条 = 高 1 + 中 0 + 低 1" in low_ws.cell(row=2, column=1).value,
      low_ws.cell(row=2, column=1).value)
check("低优先第 1 条数据",
      low_ws.cell(row=3, column=1).value == "低"
      and low_ws.cell(row=3, column=2).value == "置信度不足"
      and low_ws.cell(row=3, column=4).value == 6
      and low_ws.cell(row=3, column=6).value == "0")
check("空 zoom 显示（无）", low_ws.cell(row=3, column=8).value == "（无）")
check("低优先冻结说明行", low_ws.freeze_panes == "A3")

print("== 用例 2：无 flags -> 无清单 sheet，flagged=0 ==")
out2 = tmp / "r2.xlsx"
page2 = ExactPage(name="运单2.png",
                  blocks=[Block(kind="table", rows=[["1", "a"], ["2", "b"]])])
res2 = convert_exact([page2], out2)
wb2 = load_workbook(out2)
check("flagged == 0", res2.flagged == 0)
# 0 条时清单表仍写出（表头 + 「本次无待复核项」说明），不做静默省略
check("0 条仍写出待复核清单并注明",
      "待复核清单" in wb2.sheetnames
      and isinstance(wb2["待复核清单"].cell(row=2, column=1).value, str)
      and "本次无待复核项" in wb2["待复核清单"].cell(row=2, column=1).value,
      wb2.sheetnames)
check("数据页 + 索引异常 + 单号索引",
      wb2.sheetnames[0] == "运单2" and wb2.sheetnames[-1] == "单号索引", wb2.sheetnames)

print("== 用例 3：多页 flags 汇总到一个清单 ==")
out3 = tmp / "r3.xlsx"
pa = ExactPage(name="A.png",
               blocks=[Block(kind="table", rows=[["1", "x"]])],
               flags=[{"row": 0, "col": 1, "rebuilt": "", "whole": "1011",
                       "local": "", "reason": "数字列粘连，无法唯一拆分，待复核"}])
pb = ExactPage(name="B.png",
               blocks=[Block(kind="table", rows=[["2", "z"]])],
               flags=[{"row": 0, "col": 0, "rebuilt": "2", "ocr": "3",
                       "zoom": "3", "reason": "放大复识无结果（整图 OCR 置信度仅 0.20）"}])
res3 = convert_exact([pa, pb], out3)
wb3 = load_workbook(out3)
check("flagged == 2", res3.flagged == 2)
check("主清单收高优先(A.png)",
      wb3["待复核清单"].cell(row=2, column=3).value == "A.png"
      and wb3["待复核清单"].cell(row=2, column=1).value == "高")
check("粘连 flag 回退显示 whole/local",
      wb3["待复核清单"].cell(row=2, column=6).value == "（无）"
      and wb3["待复核清单"].cell(row=2, column=7).value == "1011")
check("低优先 sheet 收 B.png",
      wb3["待复核-低优先"].cell(row=3, column=3).value == "B.png"
      and wb3["待复核-低优先"].cell(row=3, column=1).value == "低")
check("两 sheet 合计 == flagged",
      (wb3["待复核清单"].max_row - 2)          # 主清单：表头 + 条目 + 合计行
      + (wb3["待复核-低优先"].max_row - 2)     # 低优先：表头+说明行
      == res3.flagged,
      f"{wb3['待复核清单'].max_row} {wb3['待复核-低优先'].max_row}")
check("sheet 顺序：数据页在首、清单在后、索引最后",
      wb3.sheetnames[:2] == ["A", "B"]
      and wb3.sheetnames.index("待复核清单") > wb3.sheetnames.index("B")
      and wb3.sheetnames[-1] == "单号索引",
      wb3.sheetnames)

print("== 用例 4：空内容页不影响清单 ==")
out4 = tmp / "r4.xlsx"
pc = ExactPage(name="空.png", blocks=[], warnings=["未识别到内容"])
pd = ExactPage(name="D.png",
               blocks=[Block(kind="table", rows=[["1", "q"]])],
               flags=[{"row": 0, "col": 1, "rebuilt": "q", "ocr": "w",
                       "zoom": "", "reason": "t"}])
res4 = convert_exact([pc, pd], out4)
check("空页跳过但有清单", res4.sheets == 1 and res4.flagged == 1)

print(f"\n结果：{PASS} 通过 / {FAIL} 失败")
sys.exit(1 if FAIL else 0)

