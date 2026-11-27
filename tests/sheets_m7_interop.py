# SPDX-License-Identifier: MIT
"""Independent rich string, named style and chart-range verification."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import NamedStyle, Font, PatternFill
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.chart import BarChart, Reference
import xlsxwriter, subprocess, sys

def run(directory):
    p=Path(directory).resolve();p.mkdir(parents=True,exist_ok=True)
    w=xlsxwriter.Workbook(p/'shared.xlsx');s=w.add_worksheet('Report')
    first=w.add_format({'bold':True,'font_color':'#AA2244'})
    second=w.add_format({'italic':True})
    s.write_rich_string('A1',first,'Title',second,' Tail')
    s.write_rich_string('A2',first,'Keep',second,' Original')
    s.write_formula('B1','=LEN(A1)',None,10)
    w.close()
    w=Workbook();s=w.active;s.title='Report'
    report=NamedStyle(name='Report Format')
    report.font=Font(name='Calibri',size=12,bold=True,color='FF224466')
    report.fill=PatternFill('solid',fgColor='FFEEEEDD');report.number_format='0.00'
    w.add_named_style(report)
    s['A1']=CellRichText(TextBlock(InlineFont(rFont='Calibri',b=True,color='FFAA2244'),'Title'),TextBlock(InlineFont(rFont='Calibri',i=True),' Tail'))
    s['B1']=20;s['B1'].style=report.name
    s['B2']=21;s['B2'].style=report.name;s['B2'].font=Font(name='Calibri',size=12,bold=False)
    s['A3']='Category';s['B3']='Value';s['A4']='East';s['B4']=12;s['A5']='West';s['B5']=18
    chart=BarChart();chart.add_data(Reference(s,min_col=2,min_row=3,max_row=5),titles_from_data=True);chart.set_categories(Reference(s,min_col=1,min_row=4,max_row=5));s.add_chart(chart,'D4')
    w.save(p/'inline.xlsx')
    subprocess.run(['node','--import','tsx','tests/sheets-m7-interop.mjs',str(p)],check=True)
    r=load_workbook(p/'shared-edited.xlsx',rich_text=True);a=r.active['A1'].value
    assert str(a)=='Title New Tail' and a[0].font.b and a[0].font.color.rgb=='FFAA2244' and a[1].font.i
    assert str(r.active['A2'].value)=='Keep Original'
    assert load_workbook(p/'shared-edited.xlsx',data_only=True).active['B1'].value==14
    with ZipFile(p/'shared.xlsx') as a,ZipFile(p/'shared-edited.xlsx') as b:
        assert a.read('xl/styles.xml')==b.read('xl/styles.xml')
        # The original rich shared string nodes remain available. Only the
        # optional derived reference count may change after an inline edit.
        from xml.etree import ElementTree as ET
        aroot=ET.fromstring(a.read('xl/sharedStrings.xml'));broot=ET.fromstring(b.read('xl/sharedStrings.xml'))
        assert [ET.tostring(c) for c in aroot]==[ET.tostring(c) for c in broot]
    r=load_workbook(p/'inline-edited.xlsx',rich_text=True);s=r.active
    assert s['B1'].style=='Report Format' and s['B1'].font.italic and s['B2'].style=='Report Format'
    assert s['C1'].style=='Report Format' and s['C1'].number_format=='0.00'
    assert 'Report Format' in r.named_styles
    assert str(s['A1'].value)=='Title Tail' and s['A1'].value[0].font.b
    assert s._charts[0].series[0].val.numRef.f=="'Report'!$B$5:$B$6"
    assert s._charts[0].anchor._from.row==4
    print('Independent rich shared/inline text, named style association, formula cache and native chart shift readback passed')

if len(sys.argv)>1:run(sys.argv[1])
else:
    with TemporaryDirectory() as p:run(p)
