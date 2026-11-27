# SPDX-License-Identifier: MIT
"""Independent office documents that deliberately need a bounded fallback."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import subprocess,sys

def fixtures(p):
    p.mkdir(parents=True,exist_ok=True)
    Image.new('RGB',(32,16),(51,112,255)).save(p/'preview.png',format='JPEG')
    wb=Workbook();s=wb.active;s.title='Report';s['A1']=20;s['B1']='Original';s['A2']=True;s['A1'].border=Border(diagonal=Side(style='thin',color='FF224466'),diagonalUp=True)
    hidden=wb.create_sheet('Hidden');hidden.sheet_state='hidden';hidden['A1']='Private hidden cell'
    s.row_dimensions[4].hidden=True;s['A4']='Private hidden row'
    wb.save(p/'compat.xlsx')
    wb=load_workbook(p/'compat.xlsx');wb['Hidden']['B1']='=Report!A1';wb.save(p/'formula-locked.xlsx')
    deck=Presentation();sl=deck.slides.add_slide(deck.slide_layouts[6])
    box=sl.shapes.add_textbox(Inches(1),Inches(.5),Inches(8),Inches(.8));box.text='Original slide'
    table=sl.shapes.add_table(2,2,Inches(1),Inches(2),Inches(8),Inches(2)).table
    paragraph=table.cell(0,0).text_frame.paragraphs[0];r=paragraph.add_run();r.text='Bold part';r.font.bold=True;r=paragraph.add_run();r.text=' Plain part'
    table.cell(1,1).text='Independent table cell'
    sl=deck.slides.add_slide(deck.slide_layouts[6]);sl.shapes.add_textbox(Inches(1),Inches(1),Inches(8),Inches(1)).text='Preserve this slide'
    deck.save(p/'compat.pptx')
    subprocess.run(['node','--import','tsx','tests/office-m76-compat.mjs',str(p)],check=True)

def verify(p):
    wb=load_workbook(p/'compat-edited.xlsx');s=wb['Report'];assert s['A1'].value==99.25 and s['B1'].value=='Updated cell' and s['A2'].value is False
    assert s['A1'].border.diagonalUp and s['A1'].border.diagonal.style=='thin'
    assert wb['Hidden'].sheet_state=='hidden' and wb['Hidden']['A1'].value=='Private hidden cell'
    assert s['A4'].value=='Private hidden row' and s.row_dimensions[4].hidden
    deck=Presentation(p/'compat-edited.pptx');assert deck.slides[0].shapes[0].text=='Updated slide'
    t=deck.slides[0].shapes[1].table
    assert t.cell(0,0).text=='Bold part Plain part' and t.cell(0,0).text_frame.paragraphs[0].runs[0].font.bold
    assert t.cell(1,1).text=='Updated table cell'
    assert deck.slides[1].shapes[0].text=='Preserve this slide'
    d=Document(p/'compat-edited.docx');assert d.paragraphs[0].text=='Updated Word'
    assert d.paragraphs[0].runs[0].font.strike is True
    with ZipFile(p/'compat.docx') as a,ZipFile(p/'compat-edited.docx') as b:
        for name in a.namelist():
            if name!='word/document.xml':assert a.read(name)==b.read(name),name
        xml=b.read('word/document.xml')
        assert b'bookmarkStart' in xml and b'OLEObject' in xml
    print('Independent openpyxl/python-pptx/python-docx readback: values, direct styles, hidden data, mixed fonts, untouched slide and embedded workbook preserved')

def run(p):fixtures(p);verify(p)
if len(sys.argv)>1:run(Path(sys.argv[1]).resolve())
else:
    with TemporaryDirectory() as p:run(Path(p))
