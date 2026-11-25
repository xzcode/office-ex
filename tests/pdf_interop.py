# SPDX-License-Identifier: MIT
"""Test-only PDF fixture generation and independent structural/Poppler checks."""
import argparse
import json
from pathlib import Path
import re
import subprocess
import unicodedata

from pypdf import PdfReader, PdfWriter
from pypdf.annotations import Link
from pypdf.actions import JavaScript
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject


def generate(output: Path):
    output.mkdir(parents=True, exist_ok=True)
    writer = PdfWriter()
    page = writer.add_blank_page(width=612, height=792)
    page[NameObject('/Resources')] = DictionaryObject({NameObject('/Font'): DictionaryObject({NameObject('/F1'): DictionaryObject({NameObject('/Type'): NameObject('/Font'), NameObject('/Subtype'): NameObject('/Type1'), NameObject('/BaseFont'): NameObject('/Helvetica')})})})
    stream = DecodedStreamObject()
    stream.set_data(b'BT /F1 12 Tf 50 700 Td (Security fixture external actions must not execute) Tj ET')
    page.replace_contents(stream)
    writer.add_annotation(0, Link(rect=(50, 675, 350, 710), url='https://example.invalid/office-ex-test'))
    writer.add_attachment('untrusted.txt', b'Not executed')
    writer.add_open_action(JavaScript('app.alert("unsafe fixture")'))
    writer.write(output/'actions.pdf')
    for name, password in [('encrypted.pdf', 'secret'), ('encrypted-empty.pdf', '')]:
        encrypted = PdfWriter(clone_from=output/'actions.pdf')
        # Test fixtures only: legacy encryption is sufficient to exercise the refusal policy.
        encrypted.encrypt(password, owner_password='test-owner', algorithm='RC4-128')
        encrypted.write(output/name)
    assert len(PdfReader(output/'actions.pdf').pages) == 1
    assert all(PdfReader(output/name).is_encrypted for name in ('encrypted.pdf', 'encrypted-empty.pdf'))


def normalized(text: str) -> str:
    return ''.join(ch for ch in unicodedata.normalize('NFKC', text) if unicodedata.category(ch) != 'Cf')


def verify(output: Path):
    expected = json.loads((output/'projection.json').read_text(encoding='utf-8'))
    evidence = {}
    for kind in ('writer', 'slides'):
        source = output/(kind+'.pdf')
        reader = PdfReader(source)
        sizes = expected[kind]['pageSizes']
        assert len(reader.pages) == len(sizes), (kind, len(reader.pages), len(sizes))
        for page, size in zip(reader.pages, sizes):
            assert abs(float(page.mediabox.width)-size['width']*.75) < 1, (kind, page.mediabox, size)
            assert abs(float(page.mediabox.height)-size['height']*.75) < 1, (kind, page.mediabox, size)
        text = normalized(subprocess.check_output(['pdftotext','-layout',str(source),'-']).decode('utf-8'))
        fonts = subprocess.check_output(['pdffonts',str(source)]).decode('utf-8')
        rows = [line for line in fonts.splitlines()[2:] if line.strip()]
        assert rows and all(re.search(r'\s+yes\s+(?:yes|no)\s+yes\s+\d+\s+\d+\s*$', row) for row in rows), fonts
        if kind == 'writer':
            assert 'Writer PDF' in text and '中文' in text, text[:400]
            numbers = re.findall(r'UniqueLine(\d{3})',text)
            assert numbers == [f'{i:03d}' for i in range(1,81)], numbers
        else:
            assert len(reader.pages)==3
            for i in range(1,4): assert f'Slide {i}' in text, text[:400]
            assert '文本' in text, text[:400]
        assert any(len(page.images) for page in reader.pages), f'{kind}: embedded images absent'
        evidence[kind] = {'pages': len(reader.pages), 'pageSizesPt': [[float(page.mediabox.width),float(page.mediabox.height)] for page in reader.pages], 'fontCount': len(rows), 'allFontsEmbeddedWithUnicode': True, 'pagesWithImages': sum(bool(len(page.images)) for page in reader.pages)}
    (output/'independent-result.json').write_text(json.dumps(evidence,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(evidence,ensure_ascii=False,indent=2))


if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('operation',choices=['generate','verify'])
    parser.add_argument('output',type=Path)
    args=parser.parse_args()
    (generate if args.operation=='generate' else verify)(args.output)
