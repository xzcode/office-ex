# SPDX-License-Identifier: MIT
"""Independent office document producers; no exact Microsoft Office layout claim."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import os,subprocess,sys
import xlsxwriter
from docx import Document
from docx.shared import Inches
from pptx import Presentation
from pptx.util import Inches as I
from pptx.enum.text import MSO_AUTO_SIZE
from PIL import Image
from openpyxl import load_workbook
ROOT=Path(__file__).resolve().parents[1]
def generate(p):
    p.mkdir(parents=True,exist_ok=True)
    Image.new('RGB',(100,70),(40,140,190)).save(p/'sample.png')
    w=xlsxwriter.Workbook(p/'standard.xlsx')
    s=w.add_worksheet('销售');s.write('A1','数据');s.write('B1',123);s.write_formula('B2','=B1*2',None,246)
    w.add_worksheet('汇总').write('A1','汇总表');w.close()
    w=xlsxwriter.Workbook(p/'extended.xlsx')
    s=w.add_worksheet('业务数据')
    s.write_rich_string('A1',w.add_format({'bold':True}),'业务','数据');s.write('B1',125);s.write_formula('C1','=SUM(B1,25)',None,150)
    s.write('A80','第八十行');s.write('Z80','远端')
    s.set_column('E:E',None,None,{'hidden':True});s.write('E1','隐藏列')
    s.set_row(2,None,None,{'hidden':True});s.write('A3','隐藏行')
    s.conditional_format('B1:B20',{'type':'cell','criteria':'>','value':100,'format':w.add_format({'bg_color':'#FFDDAA'})})
    w.add_worksheet('汇总').write('A1','第二张表')
    hidden=w.add_worksheet('隐藏');hidden.hide();hidden.write('A1','隐藏数据');w.close()
    d=Document();d.add_heading('标题',0);d.add_paragraph('中文正文');d.add_paragraph('First item',style='List Number');d.save(p/'ordinary-text.docx')
    d=Document();t=d.add_table(rows=2,cols=2);t.style='Table Grid'
    for i,row in enumerate([['产品','数量'],['办公软件','12']]):
        for j,value in enumerate(row):t.cell(i,j).text=value
    d.save(p/'ordinary-table.docx')
    d=Document();t=d.add_table(rows=2,cols=2)
    t.cell(0,0).text='Normal table';t.cell(0,1).text='ordinary'
    d.save(p/'normal-table.docx')
    d=Document();t=d.add_table(rows=2,cols=2);t.style='Table Grid'
    t.cell(0,0).merge(t.cell(0,1)).text='Merged cell';d.save(p/'merged-table.docx')
    d.sections[0].header.paragraphs[0].text='Report Header';d.sections[0].footer.paragraphs[0].text='Footer';d.add_paragraph('Read-only content');d.save(p/'advanced.docx')
    d=Document();d.add_picture(str(p/'sample.png'),width=Inches(1));d.save(p/'ordinary-image.docx')
    r=Presentation();s=r.slides.add_slide(r.slide_layouts[0]);s.shapes.title.text='项目报告';s.placeholders[1].text='Office EX'
    s=r.slides.add_slide(r.slide_layouts[1]);s.shapes.title.text='工作内容';s.placeholders[1].text='第一条\n第二条';r.save(p/'ordinary.pptx')
    r=Presentation();s=r.slides.add_slide(r.slide_layouts[6]);b=s.shapes.add_textbox(I(1),I(1),I(4),I(1));b.text='常规文本框';s.shapes.add_picture(str(p/'sample.png'),I(1),I(2),I(2));r.save(p/'ordinary-shapes.pptx')
    r=Presentation();s=r.slides.add_slide(r.slide_layouts[6]);s.shapes.add_textbox(I(1),I(1),I(4),I(1)).text='第一张幻灯片'
    t=s.shapes.add_table(2,2,I(1),I(2),I(5),I(1)).table
    for i,row in enumerate([['类别','数量'],['文件','12']]):
        for j,value in enumerate(row):t.cell(i,j).text=value
    s=r.slides.add_slide(r.slide_layouts[5]);s.shapes.title.text='第二张幻灯片';r.save(p/'advanced.pptx')
def verify(p):
    subprocess.run([os.environ.get('NODE','node'),'--import','tsx','tests/office-external-compat.mjs',str(p)],cwd=ROOT,check=True)
    a=load_workbook(p/'standard-edited.xlsx',data_only=True);assert a['销售']['B1'].value==300 and a['销售']['B2'].value==600
    d=Document(p/'image-edited.docx');assert len(d.inline_shapes)==1 and d.inline_shapes[0].width==457200
    with ZipFile(p/'ordinary-image.docx') as before,ZipFile(p/'image-edited.docx') as after:
        for name in before.namelist():
            if name!='word/document.xml':assert before.read(name)==after.read(name),name
    from docx.oxml.ns import qn
    d=Document(p/'table-edited.docx');t=d.tables[0]
    assert t.style.style_id=='TableGrid' and t.autofit and len(t.rows)==3 and len(t.columns)==2
    assert t.cell(1,0).text=='办公软件已编辑' and t.cell(2,0).text=='新增行'
    assert t._tbl.tblPr.find(qn('w:tblW')).get(qn('w:type'))=='auto'
    assert t._tbl.tblPr.find(qn('w:tblLook')).get(qn('w:val'))=='04A0'
    with ZipFile(p/'ordinary-table.docx') as before,ZipFile(p/'table-edited.docx') as after:
        for name in before.namelist():
            if name!='word/document.xml':assert before.read(name)==after.read(name),name
    d=Document(p/'normal-table-edited.docx');assert d.tables[0].cell(0,0).text=='Normal edited' and d.tables[0].autofit
    r=Presentation(p/'shapes-edited.pptx');assert r.slides[0].shapes[0].text=='编辑普通文本框'
    assert r.slides[0].shapes[0].text_frame.auto_size==MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT and len(r.slides[0].shapes)==2
    print('Independent Excel, Word and PowerPoint readback of edited files passed')
if len(sys.argv)>1:
    p=Path(sys.argv[1]).resolve();generate(p);verify(p)
else:
    with TemporaryDirectory() as tmp:
        p=Path(tmp);generate(p);verify(p)
