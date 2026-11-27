"""Independent python-docx numbering source and output checks."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import os,subprocess
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
ROOT=Path(__file__).resolve().parents[1]
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def el(name,attrs=None,*children):
    n=OxmlElement('w:'+name)
    for k,v in (attrs or {}).items():n.set(qn('w:'+k),str(v))
    for c in children:n.append(c)
    return n
def prop(name,value):return el(name,{'val':value})
def add_list(root,abstract_id,num_id,fmt='decimal'):
    levels=[]
    for i in range(2):
        left=720+i*720;hanging=720
        levels.append(el('lvl',{'ilvl':i},prop('start',1),prop('numFmt',fmt),prop('lvlText','%1.' if i==0 else '%1.%2.'),prop('lvlJc','left'),el('pPr',{},el('tabs',{},el('tab',{'val':'num','pos':left})),el('ind',{'left':left,'hanging':hanging}))))
    abstract=el('abstractNum',{'abstractNumId':abstract_id},prop('multiLevelType','multilevel'),*levels)
    first=root.find(qn('w:num'));root.insert(root.index(first),abstract)
    root.append(el('num',{'numId':num_id},prop('abstractNumId',abstract_id),el('lvlOverride',{'ilvl':0},prop('startOverride',4))))
def number(p,num_id,level):
    p._p.get_or_add_pPr().append(el('numPr',{},prop('ilvl',level),prop('numId',num_id)))
with TemporaryDirectory() as tmp:
    p=Path(tmp);d=Document();d.add_paragraph('First',style='List Number');d.add_paragraph('Second',style='List Number');d.add_paragraph('Bullet',style='List Bullet')
    root=d.part.numbering_part.element
    add_list(root,20,20);add_list(root,21,21,'upperRoman')
    for text,level in [('Main',0),('Nested one',1),('Nested two',1),('Next main',0)]:number(d.add_paragraph(text),20,level)
    d.save(p/'lists.docx')
    bad=Document(p/'lists.docx');number(bad.add_paragraph('Unsupported active roman list'),21,0);bad.save(p/'protected.docx')
    subprocess.run([os.environ.get('NODE','node'),'--import','tsx','tests/writer-numbering-interop.mjs',str(p)],cwd=ROOT,check=True)
    edited=Document(p/'lists-edited.docx');assert edited.paragraphs[0].text=='Edited' and edited.paragraphs[0].style.name=='List Number'
    restarted=Document(p/'lists-restarted.docx')
    ids=[x._p.pPr.numPr.numId.val for x in restarted.paragraphs[3:]]
    assert len(set(ids))==1 and ids[0]==22
    assert [x._p.pPr.numPr.ilvl.val for x in restarted.paragraphs[3:]]==[0,1,1,0]
    with ZipFile(p/'lists.docx') as a,ZipFile(p/'lists-restarted.docx') as b:
        for name in a.namelist():
            if name not in ['word/document.xml','word/numbering.xml']:assert a.read(name)==b.read(name),name
        before,after=[ET.fromstring(z.read('word/numbering.xml')) for z in [a,b]]
        for name,key in [('abstractNum','abstractNumId'),('num','numId')]:
            current={x.get(W+key):ET.tostring(x) for x in after.findall(W+name)}
            for x in before.findall(W+name):assert current[x.get(W+key)]==ET.tostring(x)
        assert len(after.findall(W+'num'))==len(before.findall(W+'num'))+1
print('Independent python-docx verifies source numbering XML preservation, native list identity, new restart and edited paragraph text.')
