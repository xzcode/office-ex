# SPDX-License-Identifier: MIT
"""Independent Word variant/multi-paragraph running-story source and readback."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import subprocess,sys
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def page_start(section,start):
    old=section._sectPr.find(qn('w:pgNumType'))
    if old is not None:section._sectPr.remove(old)
    field=OxmlElement('w:pgNumType');field.set(qn('w:start'),str(start));field.set(qn('w:fmt'),'decimal')
    section._sectPr.insert_element_before(field,'w:cols','w:titlePg','w:docGrid')

def field(p,label):
    p.add_run(label)
    f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),' PAGE ')
    r=OxmlElement('w:r');t=OxmlElement('w:t');t.text='999';r.append(t);f.append(r);p._p.append(f)
    p.add_run(' / report')

def run(directory):
    p=Path(directory).resolve();p.mkdir(parents=True,exist_ok=True)
    d=Document();d.settings.odd_and_even_pages_header_footer=True
    s=d.sections[0];s.different_first_page_header_footer=True
    s.top_margin=Inches(1.5);s.bottom_margin=Inches(1.5);s.header_distance=Inches(.25);s.footer_distance=Inches(.25);page_start(s,4)
    s.header.paragraphs[0].text='Odd header';s.header.add_paragraph('Odd second',style='Header').alignment=WD_ALIGN_PARAGRAPH.RIGHT
    s.first_page_header.paragraphs[0].text='First header';s.first_page_header.add_paragraph('First subtitle',style='Header').alignment=WD_ALIGN_PARAGRAPH.CENTER
    s.even_page_header.paragraphs[0].text='Even header'
    field(s.footer.paragraphs[0],'Odd ');field(s.first_page_footer.paragraphs[0],'First ');field(s.even_page_footer.paragraphs[0],'Even ')
    s.footer.add_paragraph('Footer second line',style='Footer')
    d.add_paragraph('Body A');d.add_page_break();d.add_paragraph('Body B');d.add_page_break();d.add_paragraph('Body C')
    s=d.add_section(WD_SECTION.NEW_PAGE);s.different_first_page_header_footer=False;page_start(s,10)
    d.add_paragraph('Body D');d.add_page_break();d.add_paragraph('Body E')
    d.save(p/'variants.docx')
    subprocess.run(['node','--import','tsx','tests/writer-m7-stories-interop.mjs',str(p)],check=True)
    result=Document(p/'variants-edited.docx')
    assert result.settings.odd_and_even_pages_header_footer
    assert result.sections[0].different_first_page_header_footer and not result.sections[1].different_first_page_header_footer
    assert [x.text for x in result.sections[0].first_page_header.paragraphs]==['First header','Edited subtitle']
    assert result.sections[0].even_page_header.paragraphs[0].text=='Even header'
    assert not result.sections[1].even_page_header.is_linked_to_previous
    assert result.sections[1].even_page_header.paragraphs[0].text=='Detached even header'
    assert result.sections[1].header.is_linked_to_previous
    assert result.sections[0]._sectPr.find(qn('w:pgNumType')).get(qn('w:start'))=='4'
    assert result.sections[1]._sectPr.find(qn('w:pgNumType')).get(qn('w:start'))=='10'
    assert result.sections[0].even_page_footer.paragraphs[0]._p.find(qn('w:fldSimple')).get(qn('w:instr')).strip()=='PAGE'
    off=Document(p/'variants-disabled.docx');assert not off.settings.odd_and_even_pages_header_footer
    assert not off.sections[0].different_first_page_header_footer
    # Inactive variants remain native package parts and can be restored.
    assert off.sections[0].first_page_header.paragraphs[0].text=='First header'
    with ZipFile(p/'variants.docx') as a,ZipFile(p/'variants-disabled.docx') as b:
        old=ET.fromstring(a.read('word/settings.xml'));new=ET.fromstring(b.read('word/settings.xml'))
        assert [ET.tostring(c) for c in old if c.tag!=W+'evenAndOddHeaders']==[ET.tostring(c) for c in new]
    print('Independent python-docx readback: variants, multi-paragraph text, section inheritance, PAGE and settings toggle passed')

if len(sys.argv)>1:run(sys.argv[1])
else:
    with TemporaryDirectory() as p:run(p)
