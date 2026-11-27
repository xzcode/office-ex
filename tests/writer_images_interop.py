"""Independent python-docx checks of embedded PNG/JPEG geometry and original media."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from hashlib import sha256
import subprocess
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
def main():
    with TemporaryDirectory() as tmp:
        a,b=Path(tmp)/'images.docx',Path(tmp)/'resized.docx'
        subprocess.run(['node','--import','tsx','tests/writer-image-fixture.mjs',str(a),str(b)],cwd=ROOT,check=True)
        original,edited=Document(a),Document(b)
        assert original.paragraphs[0].text=='Picture example'
        assert len(original.inline_shapes)==2 and len(edited.inline_shapes)==2
        assert edited.inline_shapes[0].width==1270000 and edited.inline_shapes[0].height==952500
        with ZipFile(a) as za, ZipFile(b) as zb:
            paths=[n for n in za.namelist() if '/media/' in n]
            assert len(paths)==2
            for path in za.namelist():
                if path!='word/document.xml': assert za.read(path)==zb.read(path),path
            for path in paths:
                assert sha256(za.read(path)).hexdigest() in path
        print('PASS: embedded PNG/JPEG and resized geometry independently read by python-docx; source media bytes unchanged.')
if __name__=='__main__': main()
