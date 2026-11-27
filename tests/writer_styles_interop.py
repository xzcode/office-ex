"""Independent python-docx reading of controlled named-style/theme fixtures."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import subprocess
from docx import Document
from docx.shared import RGBColor

ROOT = Path(__file__).resolve().parents[1]

def main():
    with TemporaryDirectory() as tmp:
        a, b = Path(tmp) / 'styled.docx', Path(tmp) / 'edited.docx'
        subprocess.run(['node','--import','tsx','tests/writer-styled-fixture.mjs',str(a),str(b)],cwd=ROOT,check=True)
        original, edited = Document(a), Document(b)
        assert original.paragraphs[0].text == 'Styled 中文 😀'
        assert edited.paragraphs[0].text == 'Edited 中文 😀'
        assert edited.paragraphs[0].style.style_id == 'Heading'
        assert edited.paragraphs[1].style.style_id == 'Normal'
        assert edited.styles['Heading'].base_style.style_id == 'Normal'
        assert edited.paragraphs[0].runs[0].style.style_id == 'Accent'
        assert edited.paragraphs[0].runs[0].font.name == 'Times New Roman'
        assert edited.paragraphs[0].runs[0].font.size.pt == 14
        assert edited.paragraphs[0].runs[0].font.color.rgb == RGBColor(170,34,17)
        assert edited.sections[0].header_distance.twips == 720
        with ZipFile(a) as za, ZipFile(b) as zb:
            assert set(za.namelist()) == set(zb.namelist())
            for name in za.namelist():
                if name != 'word/document.xml':
                    assert za.read(name) == zb.read(name), name
        print('PASS: named styles, direct overrides, theme/source parts and page offsets independently read by python-docx.')
if __name__ == '__main__':
    main()
