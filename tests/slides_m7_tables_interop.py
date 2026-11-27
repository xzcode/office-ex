# SPDX-License-Identifier: MIT
"""Independent native DrawingML merged-table and cell-format interop."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN,MSO_ANCHOR
from pptx.oxml.xmlchemy import OxmlElement
import subprocess,sys

def run(directory):
    p=Path(directory).resolve();p.mkdir(parents=True,exist_ok=True)
    deck=Presentation();slide=deck.slides.add_slide(deck.slide_layouts[6])
    t=slide.shapes.add_table(4,4,Inches(1),Inches(1),Inches(8),Inches(3)).table
    for r in range(4):
        for c in range(4):t.cell(r,c).text=f'{r},{c}'
    t.cell(0,0).merge(t.cell(1,2));t.cell(2,1).merge(t.cell(3,1))
    cell=t.cell(2,2);cell.fill.solid();cell.fill.fore_color.rgb=RGBColor(20,40,120)
    cell.vertical_anchor=MSO_ANCHOR.MIDDLE;cell.margin_left=Inches(.2)
    paragraph=cell.text_frame.paragraphs[0];paragraph.alignment=PP_ALIGN.CENTER
    run=paragraph.runs[0];run.font.bold=True;run.font.name='Arial';run.font.size=Pt(16);run.font.color.rgb=RGBColor(255,255,255)
    tcpr=cell._tc.get_or_add_tcPr()
    edge=OxmlElement('a:lnL');edge.set('w','19050');fill=OxmlElement('a:solidFill');rgb=OxmlElement('a:srgbClr');rgb.set('val','334455');fill.append(rgb);edge.append(fill)
    dash=OxmlElement('a:prstDash');dash.set('val','dash');edge.append(dash);tcpr.insert(0,edge)
    deck.slides.add_slide(deck.slide_layouts[6]).shapes.add_textbox(Inches(1),Inches(1),Inches(5),Inches(1)).text='Unchanged slide'
    deck.save(p/'merged.pptx')
    subprocess.run(['node','--import','tsx','tests/slides-m7-tables-interop.mjs',str(p)],check=True)
    edited=Presentation(p/'merged-edited.pptx');t=edited.slides[0].shapes[0].table
    assert len(t.rows)==5 and len(t.columns)==5
    assert t.cell(0,0).is_merge_origin and t.cell(0,0).span_height==3 and t.cell(0,0).span_width==3
    assert t.cell(0,0).text=='0,0\n0,1\n0,2\n1,0\n1,1\n1,2'
    assert t.cell(1,1).is_spanned
    assert not t.cell(3,1).is_merge_origin and t.cell(3,1).text=='2,1\n3,1'
    assert t.cell(4,1).text==''
    c=t.cell(3,2);assert c.text=='Edited';assert c.fill.fore_color.rgb==RGBColor(20,40,120)
    assert c.text_frame.paragraphs[0].runs[0].font.bold and c.text_frame.paragraphs[0].runs[0].font.italic
    assert c.text_frame.paragraphs[0].runs[0].font.size==Pt(16)
    assert c.text_frame.paragraphs[0].alignment==PP_ALIGN.CENTER and c.vertical_anchor==MSO_ANCHOR.MIDDLE
    assert c.margin_left==Inches(.2)
    assert c._tc.tcPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}lnL').get('w')=='19050'
    assert t.cell(3,4).fill.fore_color.rgb==RGBColor(100,200,40)
    with ZipFile(p/'merged.pptx') as a,ZipFile(p/'merged-edited.pptx') as b:
        for n in a.namelist():
            if n not in ['ppt/slides/slide1.xml','[Content_Types].xml']:assert a.read(n)==b.read(n),n
    print('Independent python-pptx verifies merged rectangles, text, formats, new dimensions and unchanged source parts')

if len(sys.argv)>1:run(sys.argv[1])
else:
    with TemporaryDirectory() as p:run(p)
