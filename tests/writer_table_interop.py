"""Independent python-docx checks of fixed rectangular table documents."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import subprocess
from docx import Document
ROOT = Path(__file__).resolve().parents[1]
def main():
    with TemporaryDirectory() as tmp:
        a, b = Path(tmp) / 'table.docx', Path(tmp) / 'edited.docx'
        subprocess.run(['node','--import','tsx','tests/writer-table-fixture.mjs',str(a),str(b)],cwd=ROOT,check=True)
        original, edited = Document(a), Document(b)
        assert original.paragraphs[0].text == 'Table example'
        for doc, rows in [(original,2),(edited,3)]:
            t = doc.tables[0]
            assert len(t.rows) == rows and len(t.columns) == 2
            assert t.autofit is False
            assert len(t.cell(0,0).paragraphs) == 2
            assert t.cell(0,0).paragraphs[0].text == '中文'
            assert t.cell(0,0).paragraphs[0].runs[0].bold is True
            assert t.cell(0,0).paragraphs[1].text == ' 😀'
            assert t.cell(0,1).text == '日本語'
            assert t.cell(1,0).text == 'Second row'
            assert t.cell(0,0).width == t.columns[0].width
            assert t._tbl.tblPr.xpath('./w:tblBorders/w:top')[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz') == '4'
            assert [child.tag.rsplit('}',1)[-1] for child in t._tbl.tblPr.xpath('./w:tblCellMar')[0]] == ['top','left','bottom','right']
        assert edited.tables[0].cell(2,1).text == 'Added row 中文'
        with ZipFile(a) as za, ZipFile(b) as zb:
            assert set(za.namelist()) == set(zb.namelist())
            for name in za.namelist():
                if name != 'word/document.xml':
                    assert za.read(name) == zb.read(name), name
        print('PASS: fixed rectangular DOCX table, cell paragraphs, widths, border and edited row independently read by python-docx.')
if __name__ == '__main__':
    main()
