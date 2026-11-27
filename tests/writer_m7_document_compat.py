# SPDX-License-Identifier: MIT
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.dml import MSO_THEME_COLOR_INDEX
import subprocess,sys
def run(directory):
    p=Path(directory).resolve();p.mkdir(parents=True,exist_ok=True)
    d=Document();d.add_heading('需求文档',0);d.add_heading('一、功能概述',1)
    paragraph=d.add_paragraph();r=paragraph.add_run('普通文档');r.bold=True;r.font.size=Pt(12);r.font.color.theme_color=MSO_THEME_COLOR_INDEX.ACCENT_1
    r.font.name='Calibri';fonts=r._r.rPr.rFonts;fonts.set(qn('w:asciiTheme'),'minorHAnsi');fonts.set(qn('w:hAnsiTheme'),'minorHAnsi');fonts.set(qn('w:hint'),'default')
    proof=OxmlElement('w:noProof');r._r.rPr.append(proof)
    marker=OxmlElement('w:lastRenderedPageBreak');r._r.insert(1,marker)
    start=OxmlElement('w:proofErr');start.set(qn('w:type'),'spellStart');r._r.addprevious(start)
    end=OxmlElement('w:proofErr');end.set(qn('w:type'),'spellEnd');r._r.addnext(end)
    paragraph._p.set(qn('w14:paraId'),'11112222');paragraph._p.set(qn('w14:textId'),'33334444')
    paragraph.add_run(' 的正文及样式。')
    d.add_heading('二、用户角色',1);d.add_paragraph('用户可以编辑和保存普通文档。')
    d.save(p/'ordinary-word.docx')
    subprocess.run(['node','--import','tsx','tests/writer-m7-document-compat.mjs',str(p)],check=True)
    edited=Document(p/'ordinary-word-edited.docx')
    assert edited.paragraphs[2].text=='已编辑文档 的正文及样式。'
    assert edited.paragraphs[2].runs[0].bold and edited.paragraphs[2].runs[0].font.size==Pt(12)
    with ZipFile(p/'ordinary-word.docx') as before,ZipFile(p/'ordinary-word-edited.docx') as after:
        for n in before.namelist():
            if n!='word/document.xml':assert before.read(n)==after.read(n),n
    print('Independent Chinese DOCX theme/proofing/paragraph metadata edit and preserved parts passed')
if len(sys.argv)>1:run(sys.argv[1])
else:
    with TemporaryDirectory() as p:run(p)
