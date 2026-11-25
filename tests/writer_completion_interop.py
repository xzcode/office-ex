"""Independent default-template DOCX and Word97 reader evidence, not Office visual certification."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import os, shutil, subprocess
from docx import Document
import olefile
ROOT=Path(__file__).resolve().parents[1]
def run(*args):
    subprocess.run(['node','--import','tsx','tests/writer-m2-completion-fixture.mjs',*map(str,args)],cwd=ROOT,check=True)
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
