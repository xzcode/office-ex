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
    r.font.strike=True
    shade=OxmlElement('w:shd');shade.set(qn('w:val'),'clear');shade.set(qn('w:fill'),'EFF0F1');r._r.rPr.append(shade)
    proof=OxmlElement('w:noProof');r._r.rPr.append(proof)
    marker=OxmlElement('w:lastRenderedPageBreak');r._r.insert(1,marker)
    start=OxmlElement('w:proofErr');start.set(qn('w:type'),'spellStart');r._r.addprevious(start)
    end=OxmlElement('w:proofErr');end.set(qn('w:type'),'spellEnd');r._r.addnext(end)
    paragraph._p.set(qn('w14:paraId'),'11112222');paragraph._p.set(qn('w14:textId'),'33334444')
    paragraph.add_run(' 的正文及样式。')
    pstyle=OxmlElement('w:pStyle');pstyle.set(qn('w:val'),'9999');paragraph._p.get_or_add_pPr().insert(0,pstyle)
    sect=d.sections[0]._sectPr;sect.remove(sect.find(qn('w:pgMar')))
    d.add_heading('二、用户角色',1);listed=d.add_paragraph('用户可以编辑和保存普通文档。')
    numbering=d.part.numbering_part.element
    abstract=OxmlElement('w:abstractNum');abstract.set(qn('w:abstractNumId'),'9876')
    level=OxmlElement('w:lvl');abstract.append(level)
    for key,value in [('numFmt','bullet'),('suff','tab'),('lvlText','•')]:
        child=OxmlElement('w:'+key);child.set(qn('w:val'),value);level.append(child)
    rp=OxmlElement('w:rPr');level.append(rp)
    for key,value in [('color','3370FF'),('sz','20')]:
        child=OxmlElement('w:'+key);child.set(qn('w:val'),value);rp.append(child)
    first_num=numbering.find(qn('w:num'));numbering.insert(list(numbering).index(first_num),abstract)
    instance=OxmlElement('w:num');instance.set(qn('w:numId'),'9876');aid=OxmlElement('w:abstractNumId');aid.set(qn('w:val'),'9876');instance.append(aid);numbering.append(instance)
    np=listed._p.get_or_add_pPr().get_or_add_numPr();np.get_or_add_numId().val=9876
    d.save(p/'ordinary-word.docx')
    subprocess.run(['node','--import','tsx','tests/writer-m7-document-compat.mjs',str(p)],check=True)
    edited=Document(p/'ordinary-word-edited.docx')
    assert edited.paragraphs[2].text=='已编辑文档 的正文及样式。'
    assert edited.paragraphs[2].runs[0].bold and edited.paragraphs[2].runs[0].font.size==Pt(12)
    assert edited.paragraphs[2].runs[0].font.strike is True
    assert edited.paragraphs[2].runs[0]._r.rPr.find(qn('w:shd')).get(qn('w:fill'))=='EFF0F1'
    assert edited.paragraphs[2]._p.pPr.pStyle.val=='9999'
    assert edited.sections[0]._sectPr.find(qn('w:pgMar')) is None
    assert edited.paragraphs[4]._p.pPr.numPr.numId.val==9876
    with ZipFile(p/'ordinary-word.docx') as before,ZipFile(p/'ordinary-word-edited.docx') as after:
        for n in before.namelist():
            if n!='word/document.xml':assert before.read(n)==after.read(n),n
    print('Independent Chinese DOCX decoration, missing margin/style and minimal numbering preservation passed')
if len(sys.argv)>1:run(sys.argv[1])
else:
    with TemporaryDirectory() as p:run(p)
