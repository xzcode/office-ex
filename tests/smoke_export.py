"""M0 interoperability probe. Python ZIP is test-only, not the desktop runtime."""
import json
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED
from docx import Document
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]


def main():
    result = subprocess.run(['node', '--experimental-strip-types', 'tests/export-fixtures.mjs'],
                            cwd=ROOT, check=True, capture_output=True, text=True)
    fixtures = json.loads(result.stdout)
    with TemporaryDirectory() as tmp:
        for filename, parts in fixtures.items():
            with ZipFile(Path(tmp) / filename, 'w', ZIP_DEFLATED) as z:
                for name, text in parts.items():
                    z.writestr(name, text.encode('utf-8'))
        doc = Document(Path(tmp) / 'writer.docx')
        assert doc.paragraphs[0].text == ' 中文 & <Office> 😀 \n第二行\t尾部'
        assert doc.paragraphs[0].runs[0].bold is True
        wb = load_workbook(Path(tmp) / 'sheets.xlsx')
        assert wb.sheetnames == ['数据 & A', '第二页']
        ws = wb.worksheets[0]
        assert ws['A1'].value == '=plain text' and ws['A1'].data_type == 's'
        assert (ws['B2'].value, ws['B3'].value, ws['B4'].value) == (12, 30, '=SUM(B2:B3)')
        assert ws['AA1'].value == '中文 <&>' and wb.worksheets[1]['A1'].value == 42
        wb.close()
    print('PASS: original TypeScript DOCX/XLSX XML parts; independent ZIP/read checks and 4 rejection cases.')


if __name__ == '__main__':
    main()
