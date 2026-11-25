"""Independent default-template DOCX and Word97 reader evidence, not Office visual certification."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import os, shutil, subprocess
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import olefile
ROOT=Path(__file__).resolve().parents[1]
def run(*args):
    subprocess.run([os.environ.get('NODE','node'),'--import','tsx','tests/writer-m2-completion-fixture.mjs',*map(str,args)],cwd=ROOT,check=True)
def main():
    with TemporaryDirectory() as tmp:
        ordinary,edited,legacy,legacy_edit,modern=(Path(tmp)/n for n in ['ordinary.docx','edited.docx','legacy.doc','legacy-edit.doc','modern.docx'])
        doc=Document();doc.add_paragraph('Independent 中文');doc.add_heading('A real heading',level=1);doc.add_heading('A second-level heading',level=2);doc.save(ordinary)
        run('ordinary',ordinary,edited)
        out=Document(edited);assert out.paragraphs[0].text=='Office EX 中文';assert out.paragraphs[1].style.name=='Heading 1';assert out.paragraphs[2].style.name=='Heading 2' and out.paragraphs[2].text=='A second-level heading'
        with ZipFile(ordinary) as a, ZipFile(edited) as b:
            assert set(a.namelist())==set(b.namelist())
            for name in a.namelist():
                if name!='word/document.xml':assert a.read(name)==b.read(name),name
        print('PASS: independent default-template DOCX with heading, styles, theme, settings, numbering and ancillary parts; edited reopen and byte preservation.')
        title_source,title_edited,protected=(Path(tmp)/n for n in ['title.docx','title-edited.docx','title-protected.docx'])
        sample=Document();sample.add_heading('Original Title',0);sample.add_heading('Second Title',0)
        paragraph=sample.add_paragraph('Hanging indentation example');paragraph.paragraph_format.left_indent=Pt(36);paragraph.paragraph_format.right_indent=Pt(18);paragraph.paragraph_format.first_line_indent=Pt(-18)
        run_style=paragraph.runs[0]._r.get_or_add_rPr();spacing=OxmlElement('w:spacing');spacing.set(qn('w:val'),'5');run_style.append(spacing)
        kern=OxmlElement('w:kern');kern.set(qn('w:val'),'28');run_style.append(kern)
        sample.save(title_source)
        invalid=Document(title_source);ppr=invalid.paragraphs[2]._p.get_or_add_pPr();borders=OxmlElement('w:pBdr');left=OxmlElement('w:left');left.set(qn('w:val'),'single');left.set(qn('w:sz'),'8');left.set(qn('w:color'),'000000');borders.append(left);ppr.append(borders);invalid.save(protected)
        run('title',title_source,title_edited,protected)
        checked=Document(title_edited);assert checked.paragraphs[0].text=='Modified Title' and checked.paragraphs[0].style.name=='Title';assert checked.paragraphs[1].style.name=='Title'
        assert checked.paragraphs[2].paragraph_format.first_line_indent==Pt(-18) and checked.paragraphs[2].paragraph_format.right_indent==Pt(18)
        assert checked.paragraphs[2].runs[0]._r.rPr.find(qn('w:spacing')).get(qn('w:val'))=='5'
        with ZipFile(title_source) as a, ZipFile(title_edited) as b:
            assert set(a.namelist())==set(b.namelist())
            for name in a.namelist():
                if name!='word/document.xml':assert a.read(name)==b.read(name),name
        print('PASS: independent Word Title, hanging indentation, character spacing and kerning, source styles unchanged, unsupported side rule protected.')
        run('legacy',legacy,legacy_edit,modern)
        antiword=os.environ.get('ANTIWORD') or shutil.which('antiword')
        if not antiword:
            raise RuntimeError('Antiword is required for independent DOC reader verification (test-only)')
        for file,expected in [(legacy,'Office EX 中文 direct style'),(legacy_edit,'Writer EX 中文 direct style')]:
            with olefile.OleFileIO(str(file)) as ole:
                assert sorted(ole.listdir())==[['1Table'],['WordDocument']]
                assert ole.openstream('WordDocument').read(2)==b'\xec\xa5'
            output=subprocess.check_output([antiword,'-m','UTF-8.txt',str(file)],text=True,encoding='utf-8')
            assert expected in output,repr(output)
            assert 'Second paragraph' in output
        out=Document(modern);assert out.paragraphs[0].text=='Writer EX 中文 direct style'
        print('PASS: independent Antiword reads generated/edited DOC Chinese and English; olefile validates streams; python-docx reads DOC to DOCX output.')
if __name__=='__main__':main()
