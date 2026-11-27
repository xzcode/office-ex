"""Independent readers; not production dependencies or real Excel acceptance."""
from pathlib import Path
import subprocess
import tempfile
import openpyxl
import xlrd

with tempfile.TemporaryDirectory() as tmp:
    subprocess.run(["node", "--import", "tsx", "tests/sheets-format-fixture.mjs", tmp], check=True)
    p = Path(tmp)
    wb = openpyxl.load_workbook(p / "workbook.xlsx", data_only=False)
    assert wb.sheetnames == ["Sheet1", "Other"]
    ws = wb["Sheet1"]
    assert ws["B1"].value == "中文 Excel"
    assert ws["A2"].value == "=SUM(A1,3)"
    assert ws["A4"].font.bold and ws["A4"].font.italic and ws["A4"].font.underline
    assert ws["A4"].fill.fgColor.rgb == "FFFFC864"
    assert ws["A4"].border.left.style == "thin"
    assert ws["A4"].number_format == "0.00"
    assert ws.freeze_panes == "B2" and ws.sheet_view.rightToLeft
    assert "A5:B5" in [str(m) for m in ws.merged_cells.ranges]
    assert wb.epoch.year == 1904
    data = openpyxl.load_workbook(p / "workbook.xlsx", data_only=True)
    assert data["Sheet1"]["A2"].value == 15
    assert data["Sheet1"]["A3"].value == "你好"
    assert data["Sheet1"]["A7"].value == "#DIV/0!"
    assert data["Sheet1"]["A8"].value is True
    assert data["Other"]["A1"].value == 14 and data["Other"]["A2"].value == 27

    x = xlrd.open_workbook(p / "workbook.xls", formatting_info=True)
    s = x.sheet_by_index(0)
    assert x.sheet_names() == ["Sheet1", "Other"] and x.datemode == 1
    assert s.cell_value(0, 1) == "中文 Excel"
    assert s.cell_value(1, 0) == 15 and s.cell_value(2, 0) == "你好"
    assert s.cell(5, 0).ctype == xlrd.XL_CELL_ERROR and s.cell(6, 0).value == 7
    assert s.cell(7, 0).ctype == xlrd.XL_CELL_BOOLEAN and s.cell(7, 0).value == 1
    assert x.sheet_by_index(1).cell_value(0, 0) == 14
    assert x.sheet_by_index(1).cell_value(1, 0) == 27
    xf = x.xf_list[s.cell_xf_index(3, 0)]
    font = x.font_list[xf.font_index]
    assert font.bold and font.italic and font.underline_type == 1
    assert x.colour_map[xf.background.pattern_colour_index] == (255, 200, 100)
    assert xf.border.left_line_style == 1
    assert s.merged_cells == [(4, 5, 0, 2)] and s.horz_split_pos == s.vert_split_pos == 1
print("Independent XLSX/openpyxl + BIFF8/xlrd interoperability passed (values, caches, styles, dates and panes).")
