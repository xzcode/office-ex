"""Independent reader/writer coverage for the closed mixed-image and section profile."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from docx import Document
from docx.oxml.ns import qn
import subprocess
ROOT=Path(__file__).resolve().parents[1]
def run(*args):
    subprocess.run(['node','--import','tsx','tests/writer-integrated-fixture.mjs',*map(str,args)],cwd=ROOT,check=True)
def main():
    with TemporaryDirectory() as tmp:
        original,authored,edited=(Path(tmp)/n for n in ['mixed.docx','python-authored.docx','edited.docx'])
        run('generate',original)
        doc=Document(original)
        assert len(doc.sections)==2
        assert doc.sections[0].page_width < doc.sections[0].page_height
        assert doc.sections[1].page_width > doc.sections[1].page_height
        assert doc.paragraphs[0].text=='Before after'
        assert doc.paragraphs[0].runs[-1].bold is True
        assert doc.paragraphs[0].runs[-1]._r.rPr.find(qn('w:lang')).get(qn('w:val'))=='zh-TW'
        assert len(doc.inline_shapes)==2
        assert doc.tables[0].cell(0,0).text=='Cell content'
        doc.add_paragraph('Independent edit')
        doc.save(authored)
        run('edit',authored,edited)
        reopened=Document(edited)
        assert len(reopened.sections)==2
        assert len(reopened.inline_shapes)==2
        assert reopened.paragraphs[-1].text=='Office EX edit'
        with ZipFile(authored) as a, ZipFile(edited) as b:
            assert set(a.namelist())==set(b.namelist())
            for path in a.namelist():
                if path!='word/document.xml': assert a.read(path)==b.read(path),path
        print('PASS: mixed text/cell pictures, section geometry and language; independent python-docx edit imported and resaved without changing other parts.')
if __name__=='__main__': main()
