"""Independent python-docx check of M2 semantic DOCX output, not an Office GUI test."""
from pathlib import Path
from tempfile import TemporaryDirectory
import subprocess
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = Path(__file__).resolve().parents[1]

def main():
    with TemporaryDirectory() as tmp:
        a, b = Path(tmp) / 'original.docx', Path(tmp) / 'edited.docx'
        subprocess.run(['node', '--import', 'tsx', 'tests/writer-docx-fixture.mjs', str(a), str(b)], cwd=ROOT, check=True)
        for path, prefix in [(a, ' 中文'), (b, ' 繁體')]:
            doc = Document(path)
            assert len(doc.paragraphs) == 3
            assert doc.paragraphs[0].text == prefix + ' & <Office> 😀 日本語\n한국어\tالعربية'
            assert doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
            assert doc.paragraphs[0].runs[0].bold is True
            assert doc.paragraphs[0].runs[-1].italic is True
            assert doc.paragraphs[0].runs[-1].underline is True
            assert doc.paragraphs[0].runs[-1].font.size.pt == 14
            assert doc.paragraphs[2].text == 'e\u0301 हिन्दी русский'
            assert abs(doc.sections[0].page_width - 7560000) <= 318
        print('PASS: M2 Writer DOCX semantic export and edited copy independently read by python-docx.')
if __name__ == '__main__':
    main()
