#!/usr/bin/env python3
"""Generate synthetic OOXML fixtures and check round-trip semantic invariants."""
import argparse
from pathlib import Path
from zipfile import ZipFile
from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from pptx import Presentation
from pptx.util import Inches

TITLE = "Office EX 兼容性"
FILES = ("writer.docx", "sheets.xlsx", "slides.pptx")

def generate(root):
    root.mkdir(parents=True, exist_ok=True)
    for name in FILES:
        if (root / name).exists():
            raise FileExistsError(root / name)
    doc = Document()
    doc.add_heading(TITLE, 1)
    doc.add_paragraph("中文 English 123").runs[0].bold = True
    doc.add_table(rows=1, cols=2).rows[0].cells[0].text = "保留表格"
    doc.save(root / FILES[0])

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Item", "Value"])
    ws.append(["A", 12])
    ws.append(["B", 30])
    ws["B4"] = "=SUM(B2:B3)"
    ws["B2"].number_format = "0.00"
    ws.merge_cells("A6:B6")
    ws["A6"] = TITLE
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True)
    ws.add_chart(chart, "D2")
    wb.create_sheet("第二页")["A1"] = "keep"
    wb.save(root / FILES[1])

    ppt = Presentation()
    for text in (TITLE, "第二张 Slide 2"):
        slide = ppt.slides.add_slide(ppt.slide_layouts[6])
        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text_frame.text = text
    ppt.save(root / FILES[2])

def require(ok, message):
    if not ok:
        raise ValueError(message)

def check(root):
    for name in FILES:
        with ZipFile(root / name) as archive:
            require(archive.testzip() is None, name + ": damaged ZIP")
            require("[Content_Types].xml" in archive.namelist(), name + ": no content types")
    doc = Document(root / FILES[0])
    require(any(p.text == TITLE and p.style.name == "Heading 1" for p in doc.paragraphs), "DOCX heading")
    require(any(p.text == "中文 English 123" and all(r.bold for r in p.runs if r.text) for p in doc.paragraphs), "DOCX text/bold")
    require(doc.tables[0].cell(0, 0).text == "保留表格", "DOCX table")

    wb = load_workbook(root / FILES[1])
    require(wb.sheetnames == ["Data", "第二页"], "XLSX sheet order")
    ws = wb["Data"]
    require((ws["B2"].value, ws["B3"].value, ws["B4"].value) == (12, 30, "=SUM(B2:B3)"), "XLSX values/formula")
    require(ws["B2"].number_format == "0.00", "XLSX number format")
    require("A6:B6" in ws.merged_cells and ws["A6"].value == TITLE, "XLSX merge")
    require(len(ws._charts) == 1, "XLSX chart")
    require(wb["第二页"]["A1"].value == "keep", "XLSX second sheet")
    wb.close()

    ppt = Presentation(root / FILES[2])
    require(len(ppt.slides) == 2, "PPTX slide count")
    for slide, expected in zip(ppt.slides, (TITLE, "第二张 Slide 2")):
        require(any(s.has_text_frame and s.text == expected for s in slide.shapes), "PPTX text/order")
    print("PASS: DOCX, XLSX, PPTX starter semantic checks (not visual or engine verification).")

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("generate", "check"))
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    if args.action == "generate":
        generate(args.directory)
        print("Generated:", args.directory)
    else:
        check(args.directory)

if __name__ == "__main__":
    main()
