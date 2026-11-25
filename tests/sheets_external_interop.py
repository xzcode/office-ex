"""Independent standard-template XLSX corpus (no reference Excel claim)."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as E
import os
import subprocess
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Color
from openpyxl.formatting.rule import CellIsRule
from openpyxl.workbook.defined_name import DefinedName

ROOT=Path(__file__).resolve().parents[1]
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006'
def n(ns,key):return '{'+ns+'}'+key
def shared_strings(path):
    with ZipFile(path) as z: parts={name:z.read(name) for name in z.namelist()}
    top=E.fromstring(parts['xl/worksheets/sheet1.xml'])
    cell=top.find(".//"+n(S,'c')+"[@r='A1']")
    assert cell is not None
    cell.clear();cell.set('r','A1');cell.set('s','1');cell.set('t','s');E.SubElement(cell,n(S,'v')).text='0'
    parts['xl/worksheets/sheet1.xml']=E.tostring(top,encoding='utf-8')
    table=E.Element(n(S,'sst'),{'count':'1','uniqueCount':'1'});E.SubElement(E.SubElement(table,n(S,'si')),n(S,'t')).text='中文 Shared'
    parts['xl/sharedStrings.xml']=E.tostring(table,encoding='utf-8')
    rels=E.fromstring(parts['xl/_rels/workbook.xml.rels'])
    E.SubElement(rels,n(PKG+'/relationships','Relationship'),{'Id':'rIdOexShared','Type':REL+'/sharedStrings','Target':'sharedStrings.xml'})
    parts['xl/_rels/workbook.xml.rels']=E.tostring(rels,encoding='utf-8')
    types=E.fromstring(parts['[Content_Types].xml']);E.SubElement(types,n(PKG+'/content-types','Override'),{'PartName':'/xl/sharedStrings.xml','ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})
    parts['[Content_Types].xml']=E.tostring(types,encoding='utf-8')
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)

with TemporaryDirectory() as tmp:
    p=Path(tmp);book=Workbook();sales=book.active;sales.title='Sales'
    sales['A1']='中文 Shared';sales['A1'].font=Font(name='Calibri',size=12,bold=True,color='FF008800')
    sales['A1'].fill=PatternFill('solid',fgColor=Color(theme=4,tint=.25))
    sales['B1']=12;sales['B1'].number_format='0.00';sales['B2']='=SUM(B1,Other!A1)'
    sales['A2']='Notes';sales['A2'].alignment=Alignment(wrap_text=True)
    sales.freeze_panes='B2';sales.column_dimensions['A'].width=22;sales.row_dimensions[1].height=20
    other=book.create_sheet('Other');other['A1']=7;other['B1']='=Sales!B1*2'
    book.create_sheet('Keep')['A1']='Unchanged sheet';book.create_sheet('Remove')['A1']='Deleted by explicit command'
    book.properties.title='External sample';book.properties.creator='Office EX test fixture'
    source=p/'external.xlsx';book.save(source);shared_strings(source)
    test=load_workbook(source);test['Sales'].protection.sheet=True;test.save(p/'external-protected.xlsx')
    test=load_workbook(source);test['Sales'].conditional_formatting.add('B1',CellIsRule(operator='greaterThan',formula=['0'],fill=PatternFill('solid',fgColor='FFFF0000')));test.save(p/'external-conditional.xlsx')
    test=load_workbook(source);test.defined_names.add(DefinedName('Revenue',attr_text='Sales!$B$1'));test.save(p/'external-custom-name.xlsx')
    subprocess.run([os.environ.get('NODE','node'),'--import','tsx','tests/sheets-external-template.mjs',str(p)],cwd=ROOT,check=True)
    for name,amount in [('external-edited.xlsx',25),('external-second.xlsx',40)]:
        out=load_workbook(p/name);cache=load_workbook(p/name,data_only=True)
        assert out.sheetnames==['Keep','Other','Sales','New Sheet']
        s=out['Sales'];assert s['A1'].value=='Edited 中文' and s['A1'].font.bold and s['A1'].font.color.rgb=='FF008800'
        assert s['A1'].fill.fgColor.type=='theme' and s['A1'].fill.fgColor.theme==4 and s['A1'].fill.fgColor.tint==.25
        assert s['B1'].value==amount and s['B1'].number_format=='0.00'
        assert s['B2'].value=='=SUM(B1,Other!A1)' and cache['Sales']['B2'].value==amount+7
        assert cache['Other']['B1'].value==amount*2
        assert s['A3'].fill.fgColor.rgb=='FFFFDC78' and s['A3'].number_format=='0.0000'
        assert s.freeze_panes=='B2' and s.column_dimensions['A'].width==22 and s.row_dimensions[1].height==20
        assert out.properties.title=='External sample' and out.properties.creator=='Office EX test fixture'
        with ZipFile(source) as before,ZipFile(p/name) as after:
            for part in ['docProps/core.xml','docProps/app.xml','xl/theme/theme1.xml','xl/worksheets/sheet3.xml','_rels/.rels']:
                assert before.read(part)==after.read(part),part
            assert 'xl/worksheets/sheet4.xml' not in after.namelist()
            orig=E.fromstring(before.read('xl/styles.xml'));new=E.fromstring(after.read('xl/styles.xml'))
            # Original style indices and theme declarations remain intact,
            # even when new directly formatted styles are appended.
            for table in ['fonts','fills','borders','cellXfs']:
                a=list(orig.find(n(S,table)));b=list(new.find(n(S,table)))
                assert [E.tostring(x) for x in a]==[E.tostring(x) for x in b[:len(a)]],table
    print('Independent openpyxl readback: 2 native saves, source style/theme/core metadata preservation, formula caches, structural edits and protected fixtures passed.')
