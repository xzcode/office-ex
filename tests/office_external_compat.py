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
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, PatternFill
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
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
    s.conditional_format('B1:B20',{'type':'2_color_scale','min_color':'#FFDDAA','max_color':'#AACCEE'})
    w.add_worksheet('汇总').write('A1','第二张表')
    hidden=w.add_worksheet('隐藏');hidden.hide();hidden.write('A1','隐藏数据');w.close()
    b=Workbook();s=b.active;s.title='Report'
    s.sheet_format.defaultRowHeight=26;s.sheet_format.defaultColWidth=18;s.sheet_format.customHeight=True
    s['A1']='业务报表';s['A1'].border=Border(bottom=Side(style='double',color='FF118844'),left=Side(style='mediumDashed',color='FFCC2233'))
    s['B1']=12;s['B2']='=IFERROR(PRODUCT(B1,2),0)';s['B3']='=COUNTBLANK(A1:A5)';s['B4']='=LEFT("Office EX",6)';s['B5']='=MOD(-5,3)'
    s['A2']='保留格式';s['A2'].border=Border(right=Side(style='dashDot',color='FF0000FF'))
    s.row_dimensions[3].height=31;s.column_dimensions['C'].width=25
    b.create_sheet('Other')['A1']='=Report!B1+3'
    b.save(p/'m6-sheet.xlsx')
    s['A1'].border=Border(diagonal=Side(style='thin',color='FF000000'),diagonalUp=True);b.save(p/'m6-unsupported-diagonal.xlsx')
    b=Workbook();report=b.active;report.title='Report';report['A1']='Conditional report'
    data=b.create_sheet('Data')
    for i,row in enumerate([['Apple','East',10],['apple','West',20],['Pear','East',30],['A*','East',40]],1):
        for j,value in enumerate(row,1):data.cell(i,j).value=value
    for i,row in enumerate([(1,100),(5,500),(10,1000)],1):
        data.cell(i,5).value=row[0];data.cell(i,6).value=row[1]
    for j,value in enumerate([1,5,10],5):
        data.cell(5,j).value=value;data.cell(6,j).value=value*10
    expressions=['SUMIFS(Data!C1:C4,Data!A1:A4,"Apple",Data!B1:B4,"East")',
        'COUNTIFS(Data!A1:A4,"A*",Data!B1:B4,"East")','AVERAGEIF(Data!B1:B4,"East",Data!C1:C4)',
        'VLOOKUP("Apple",Data!A1:C4,3,FALSE)','INDEX(Data!C1:C4,MATCH("Pear",Data!A1:A4,0))',
        'VLOOKUP(7,Data!E1:F3,2,TRUE)','HLOOKUP(5,Data!E5:G6,2,FALSE)','COUNTIF(Data!A:A,"")',
        'SUMIF(Data!A1:A4,"apple",Data!C1:C4)','AVERAGEIFS(Data!C1:C4,Data!A1:A4,"A*",Data!B1:B4,"East")',
        'COUNTIF(Data!A1:A4,"A~*")']
    for i,expression in enumerate(expressions,1):report.cell(i,2).value='='+expression
    b.save(p/'m6-lookups.xlsx')
    b=Workbook();s=b.active;s.title='Chart report'
    for row in [['Name','Value'],['A',2],['B',4],['C',6]]:s.append(row)
    s.conditional_formatting.add('B2:B4',CellIsRule(operator='greaterThan',formula=['3'],fill=PatternFill('solid',fgColor='FFFFAA00')))
    for cls,pos in [(BarChart,'D2'),(LineChart,'D20'),(PieChart,'D38')]:
        chart=cls();chart.add_data(Reference(s,min_col=2,min_row=1,max_row=4),titles_from_data=True)
        chart.set_categories(Reference(s,min_col=1,min_row=2,max_row=4));s.add_chart(chart,pos)
    b.save(p/'m6-charts.xlsx')
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
    d=Document();t=d.add_table(rows=4,cols=4);t.style='Table Grid'
    t.cell(0,0).merge(t.cell(1,1)).text='矩形合并'
    t.cell(0,2).merge(t.cell(2,2)).text='纵向合并'
    t.cell(3,0).merge(t.cell(3,2)).text='横向合并'
    t.cell(2,0).text='独立单元格'
    d.save(p/'complex-merged-table.docx')
    # Unsupported continuation content must never disappear during save.
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    for tr in t._tbl.findall(qn('w:tr')):
        for tc in tr.findall(qn('w:tc')):
            vm=tc.find(qn('w:tcPr')).find(qn('w:vMerge'))
            if vm is not None and vm.get(qn('w:val')) in (None,'continue'):
                run=OxmlElement('w:r');text=OxmlElement('w:t');text.text='Protected continuation';run.append(text);tc.find(qn('w:p')).append(run)
                break
        else:continue
        break
    d.save(p/'invalid-merged-table.docx')
    d=Document();d.add_table(rows=2,cols=2).cell(0,0).text='办公软件'
    d.sections[0].header.paragraphs[0].text='Report Header';d.sections[0].footer.paragraphs[0].text='Footer';d.add_paragraph('Read-only content')
    d.sections[0].header.add_paragraph('Unsupported second header paragraph');d.save(p/'advanced.docx')
    from docx.enum.section import WD_SECTION
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    d=Document();d.add_paragraph('Section one')
    d.sections[0].header.paragraphs[0].text='Independent running header'
    footer=d.sections[0].footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER;footer.add_run('Page ')
    for marker in ['begin','instruction','separate','result','end']:
        r=OxmlElement('w:r')
        if marker=='instruction':
            e=OxmlElement('w:instrText');e.set(qn('xml:space'),'preserve');e.text=' PAGE '
        elif marker=='result':
            e=OxmlElement('w:t');e.text='999'
        else:
            e=OxmlElement('w:fldChar');e.set(qn('w:fldCharType'),marker)
        r.append(e);footer._p.append(r)
    footer.add_run(' / report')
    num=OxmlElement('w:pgNumType');num.set(qn('w:start'),'5');num.set(qn('w:fmt'),'decimal')
    d.sections[0]._sectPr.insert_element_before(num,'w:cols')
    d.add_page_break();d.add_paragraph('Second page')
    d.add_section(WD_SECTION.NEW_PAGE)
    last=d.sections[-1]._sectPr.find(qn('w:pgNumType'))
    if last is not None:d.sections[-1]._sectPr.remove(last)
    d.add_paragraph('Linked section')
    d.save(p/'running-stories.docx')
    # A non-PAGE instruction is never evaluated or rewritten.
    for e in d.sections[0].footer.paragraphs[0]._p.iter(qn('w:instrText')):e.text=' DATE '
    d.save(p/'unsupported-story-field.docx')
    d=Document();d.add_picture(str(p/'sample.png'),width=Inches(1));d.save(p/'ordinary-image.docx')
    r=Presentation();s=r.slides.add_slide(r.slide_layouts[0]);s.shapes.title.text='项目报告';s.placeholders[1].text='Office EX'
    s=r.slides.add_slide(r.slide_layouts[1]);s.shapes.title.text='工作内容';s.placeholders[1].text='第一条\n第二条';r.save(p/'ordinary.pptx')
    r=Presentation();s=r.slides.add_slide(r.slide_layouts[6]);b=s.shapes.add_textbox(I(1),I(1),I(4),I(1));b.text='常规文本框';s.shapes.add_picture(str(p/'sample.png'),I(1),I(2),I(2));r.save(p/'ordinary-shapes.pptx')
    r=Presentation();s=r.slides.add_slide(r.slide_layouts[6]);s.shapes.add_textbox(I(1),I(1),I(4),I(1)).text='第一张幻灯片'
    t=s.shapes.add_table(2,2,I(1),I(2),I(5),I(1)).table
    for i,row in enumerate([['类别','数量'],['文件','12']]):
        for j,value in enumerate(row):t.cell(i,j).text=value
    s=r.slides.add_slide(r.slide_layouts[5]);s.shapes.title.text='第二张幻灯片';r.save(p/'m6-slide-table.pptx')
    t.cell(0,0).merge(t.cell(0,1));r.save(p/'advanced.pptx')
def verify(p):
    subprocess.run([os.environ.get('NODE','node'),'--import','tsx','tests/office-external-compat.mjs',str(p)],cwd=ROOT,check=True)
    a=load_workbook(p/'standard-edited.xlsx',data_only=True);assert a['销售']['B1'].value==300 and a['销售']['B2'].value==600
    b=load_workbook(p/'m6-sheet-edited.xlsx');s=b['Report'];cache=load_workbook(p/'m6-sheet-edited.xlsx',data_only=True)
    assert s['B1'].value==21 and cache['Report']['B2'].value==42 and cache['Report']['B3'].value==3 and cache['Report']['B4'].value=='Office' and cache['Report']['B5'].value==1
    assert cache['Other']['A1'].value==24
    assert s['A1'].border.bottom.style=='double' and s['A1'].border.bottom.color.rgb=='FF118844'
    assert s['A1'].border.left.style=='mediumDashed' and s['A1'].border.left.color.rgb=='FFCC2233'
    assert s['A2'].border.right.style=='dashDot' and s['A2'].border.right.color.rgb=='FF0000FF'
    assert s.sheet_format.defaultRowHeight==26 and s.sheet_format.defaultColWidth==18 and s.row_dimensions[3].height==31 and s.column_dimensions['C'].width==25
    with ZipFile(p/'m6-sheet.xlsx') as before,ZipFile(p/'m6-sheet-edited.xlsx') as after:
        for name in ['xl/styles.xml','xl/theme/theme1.xml','docProps/core.xml','docProps/app.xml']:
            assert before.read(name)==after.read(name),name
    b=load_workbook(p/'m6-lookups-edited.xlsx');cached=load_workbook(p/'m6-lookups-edited.xlsx',data_only=True)
    original=load_workbook(p/'m6-lookups.xlsx')
    assert b['Data']['C1'].value==20
    expected=[20,2,30,20,30,500,50,1048572,40,30,1]
    for i,v in enumerate(expected,1):
        assert b['Report'].cell(i,2).value==original['Report'].cell(i,2).value
        assert cached['Report'].cell(i,2).value==v,(i,cached['Report'].cell(i,2).value,v)
    with ZipFile(p/'m6-lookups.xlsx') as before,ZipFile(p/'m6-lookups-edited.xlsx') as after:
        for name in ['xl/styles.xml','xl/theme/theme1.xml','docProps/core.xml','docProps/app.xml']:
            assert before.read(name)==after.read(name),name
    for name,count in [('m6-charts-edited.xlsx',3),('m6-charts-rebuilt.xlsx',2),('m6-charts-new.xlsx',3)]:
        chartbook=load_workbook(p/name);cs=chartbook.active;assert cs['B2'].value==17 and len(cs._charts)==count
        assert [type(c).__name__ for c in cs._charts]==(['BarChart','PieChart'] if count==2 else ['BarChart','LineChart','PieChart'])
        if count==2:assert cs._charts[0].series[0].val.numRef.f=="'Renamed chart'!$B$2:$B$4"
        assert len(cs.conditional_formatting)==1
    with ZipFile(p/'m6-charts.xlsx') as before,ZipFile(p/'m6-charts-edited.xlsx') as after:
        for n in before.namelist():
            if n.startswith(('xl/drawings/','xl/charts/')) or n=='xl/styles.xml':assert before.read(n)==after.read(n),n
    tabledeck=Presentation(p/'m6-slide-table-edited.pptx');table=tabledeck.slides[0].shapes[1].table
    assert table.cell(1,0).text=='已编辑表格' and table.cell(1,1).text=='12'
    assert len(table.rows)==2 and len(table.columns)==2 and table.first_row and table.horz_banding
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
    d=Document(p/'merged-table-edited.docx');t=d.tables[0]
    assert t.cell(0,0).text==t.cell(0,1).text=='Merged edited'
    assert t.cell(0,0)._tc is t.cell(0,1)._tc
    d=Document(p/'complex-merged-table-edited.docx');t=d.tables[0]
    assert t.style.style_id=='TableGrid' and len(t.rows)==4 and len(t.columns)==4
    assert t.cell(0,0).text==t.cell(1,1).text=='矩形已编辑'
    assert t.cell(0,2).text==t.cell(2,2).text=='纵向已编辑'
    assert t.cell(3,0).text==t.cell(3,2).text=='横向已编辑'
    assert t.cell(2,0).text=='独立单元格'
    with ZipFile(p/'complex-merged-table.docx') as before,ZipFile(p/'complex-merged-table-edited.docx') as after:
        for name in before.namelist():
            if name!='word/document.xml':assert before.read(name)==after.read(name),name
    d=Document(p/'running-stories-edited.docx')
    assert len(d.sections)==2
    assert d.sections[0].header.paragraphs[0].text=='Independent running header'
    assert d.sections[1].header.paragraphs[0].text=='Detached section header'
    assert not d.sections[1].header.is_linked_to_previous
    assert d.sections[1].footer.is_linked_to_previous
    assert d.sections[0]._sectPr.find(qn('w:pgNumType')).get(qn('w:start'))=='5'
    assert d.sections[1]._sectPr.find(qn('w:pgNumType')).get(qn('w:start'))=='20'
    fld=d.sections[0].footer.paragraphs[0]._p.find(qn('w:fldSimple'))
    assert fld is not None and fld.get(qn('w:instr')).strip()=='PAGE'
    assert d.sections[0].footer.paragraphs[0].text=='Updated page  / report'
    with ZipFile(p/'running-stories.docx') as before,ZipFile(p/'running-stories-edited.docx') as after:
        for name in before.namelist():
            if name not in ['word/document.xml','word/footer1.xml','word/_rels/document.xml.rels','[Content_Types].xml']:
                assert before.read(name)==after.read(name),name
    r=Presentation(p/'shapes-edited.pptx');assert r.slides[0].shapes[0].text=='编辑普通文本框'
    assert r.slides[0].shapes[0].text_frame.auto_size==MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT and len(r.slides[0].shapes)==2
    print('Independent Excel, Word and PowerPoint readback of edited files passed')
if len(sys.argv)>1:
    p=Path(sys.argv[1]).resolve();generate(p);verify(p)
else:
    with TemporaryDirectory() as tmp:
        p=Path(tmp);generate(p);verify(p)
