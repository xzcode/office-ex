# SPDX-License-Identifier: MIT
"""Independent python-pptx + olefile interoperability, not reference Office rendering."""
import hashlib
import json
import os
from pathlib import Path
import struct
import subprocess
import tempfile
import olefile
from pptx import Presentation
from pptx.util import Inches
import zipfile

ROOT = Path(__file__).resolve().parents[1]

def check_pptx(path, expected):
    presentation = Presentation(path)
    assert len(presentation.slides) == expected['slides']
    assert (presentation.slide_width, presentation.slide_height) == (expected['size']['width'], expected['size']['height'])
    for index, slide in enumerate(presentation.slides):
        assert len(slide.shapes) == 5
        rect, ellipse, line, textbox, picture = slide.shapes
        assert (rect.left, rect.top, rect.width, rect.height) == (914400,914400,2743200,1371600)
        assert abs(rect.rotation-15) < 0.01
        assert rect.fill.fore_color.rgb == (40+50*index,100,180)
        assert (ellipse.left, ellipse.width) == (3657600,1828800)
        assert (line.left, line.top) == (914400,2743200)
        assert textbox.text.replace('\v','\n') == expected['text'][index]
        assert (picture.left,picture.top,picture.width,picture.height) == (7315200,1828800,2743200,2057400)
        assert hashlib.sha256(picture.image.blob).hexdigest() == expected['images'][int(index==1)]['sha256']

def check_ppt(path, expected):
    # Independent OLE/CFB reader; validate PowerPoint stream names and the UserEdit/persist entry chain.
    with olefile.OleFileIO(path) as ole:
        current = ole.openstream('Current User').read()
        document = ole.openstream('PowerPoint Document').read()
        pictures = ole.openstream('Pictures').read()
    assert len(current) >= 20 and len(document) > 200 and len(pictures) > 0
    def record(offset, expected_type=None):
        assert 0 <= offset <= len(document)-8
        options, kind, size = struct.unpack_from('<HHI', document, offset)
        assert offset+8+size <= len(document)
        if expected_type is not None: assert kind == expected_type, (kind,expected_type)
        return offset+8, offset+8+size
    assert struct.unpack_from('<H',current,2)[0] == 4086
    edit = struct.unpack_from('<I',current,16)[0]
    seen=set(); edits=0
    while edit:
        assert edit not in seen;seen.add(edit);edits+=1
        start,end=record(edit,4085)
        assert end-start>=28
        previous,persist=struct.unpack_from('<II',document,start+8)
        p,endp=record(persist,6002)
        assert p<endp
        while p<endp:
            entry,=struct.unpack_from('<I',document,p);p+=4
            count=entry>>20;assert count>0 and p+count*4<=endp
            for i in range(count):record(struct.unpack_from('<I',document,p+4*i)[0])
            p+=4*count
        edit=previous
    assert edits>=1
    # Images in the independent Pictures stream must be original bytes, not scene screenshots.
    for image in expected['images']:
        # Read the original bytes from independently decoded PPTX and look for them in Pictures.
        original=Presentation(path.parent/'scene.pptx')
        image_bytes=next(s.image.blob for slide in original.slides for s in slide.shapes if s.shape_type==13 and hashlib.sha256(s.image.blob).hexdigest()==image['sha256'])
        assert image_bytes in pictures
    return edits

with tempfile.TemporaryDirectory(prefix='office-ex-slides-interop-') as temporary:
    target=Path(temporary)
    subprocess.run([os.environ.get('NODE','node'),'--import','tsx','tests/slides-format-fixture.mjs',str(target)],cwd=ROOT,check=True)
    expected=json.loads((target/'expected.json').read_text())
    for name in ('scene.pptx','ppt-to-pptx.pptx'):
        check_pptx(target/name,expected)
    for name in ('scene.ppt','pptx-to-ppt.ppt','edited.ppt'):
        edits=check_ppt(target/name,expected)
        if name=='edited.ppt': assert edits>=2
    # An independent authoring tool changes the package and adds core properties.
    independent = Presentation(target/'scene.pptx')
    independent.core_properties.title = 'External presentation 中文'
    independent.core_properties.author = 'Independent fixture author'
    independent.slides[0].shapes[0].left = Inches(2)
    independent.save(target/'external-python.pptx')
    subprocess.run([os.environ.get('NODE','node'),'--import','tsx','tests/slides-external-edit.mjs',str(target)],cwd=ROOT,check=True)
    result = Presentation(target/'external-edited.pptx')
    assert result.core_properties.title == independent.core_properties.title
    assert result.core_properties.author == independent.core_properties.author
    assert result.slides[0].shapes[0].left == Inches(3)
    assert len(result.slides) == len(independent.slides) == 3
    with zipfile.ZipFile(target/'external-python.pptx') as before, zipfile.ZipFile(target/'external-edited.pptx') as after:
        assert before.read('docProps/core.xml') == after.read('docProps/core.xml')
        assert before.read('_rels/.rels') == after.read('_rels/.rels')
    print('Independent PPTX semantic read, PPT OLE records and externally edited PPTX metadata-preserving round trip passed (7 fixtures).')
