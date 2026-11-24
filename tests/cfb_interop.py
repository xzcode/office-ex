#!/usr/bin/env python3
"""LEG-01 independent container verification (olefile 0.47; test-only).

The original TypeScript writer generates synthetic streams and storage metadata.
Python's independent reader verifies every stream hash and the hierarchy, then
changes streams in place; our reader must observe the independent writer's bytes.
No real Office application or DOC/XLS/PPT semantics are claimed by this check.
"""
from __future__ import annotations
import hashlib
import json
from pathlib import Path
import subprocess
import tempfile
import olefile

ROOT = Path(__file__).resolve().parent.parent
CREATE = r"""
import { writeCfb } from '@office-ex/cfb';
import { writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';
const pattern=(n, seed)=>Uint8Array.from({length:n}, (_,i)=>(i*seed+31)&255);
const entries=[
 {type:'storage',path:['Unknown'],metadata:{stateBits:8,creationTime:'01bab44b13921e80',modifiedTime:'01bab44b13921e81'}},
 {type:'storage',path:['Unknown','Empty storage']},
 {type:'stream',path:['Unknown','Small'],data:pattern(65,17)},
 {type:'stream',path:['Unknown','Boundary'],data:pattern(4095,19)},
 {type:'stream',path:['Large'],data:pattern(8*1024*1024+7,23)},
 {type:'stream',path:['Exact'],data:pattern(4096,29)},
 {type:'stream',path:['Empty'],data:new Uint8Array()},
 {type:'stream',path:['\u0005SummaryInformation'],data:pattern(57,31)},
];
const manifest=[];
for(const majorVersion of [3,4]) {
 const path=join(process.argv[1], `v${majorVersion}.cfb`);
 writeFileSync(path,writeCfb({entries,root:{clsid:'0102030405060708090a0b0c0d0e0f10',stateBits:0x12345678}},{majorVersion}));
 manifest.push({majorVersion,path,streams:entries.filter(e=>e.type==='stream').map(e=>({path:e.path,size:e.data.length,sha256:createHash('sha256').update(e.data).digest('hex')}))});
}
console.log(JSON.stringify(manifest));
"""
READ = r"""
import { readCfb } from '@office-ex/cfb';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
const doc=readCfb(new Uint8Array(readFileSync(process.argv[1])));
console.log(JSON.stringify(doc.entries.filter(e=>e.type==='stream').map(e=>({path:e.path,size:e.data.length,sha256:createHash('sha256').update(e.data).digest('hex')}))));
"""


def node(script: str, path: str) -> object:
    completed = subprocess.run(["node", "--import", "tsx", "--input-type=module", "-e", script, path], cwd=ROOT,
                               text=True, capture_output=True, check=True)
    return json.loads(completed.stdout)


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="office-ex-cfb-") as temp:
        manifest = node(CREATE, temp)
        for generated in manifest:
            with olefile.OleFileIO(generated["path"], raise_defects=olefile.DEFECT_INCORRECT) as reference:
                assert reference.sectorsize == (512 if generated["majorVersion"] == 3 else 4096)
                assert reference.root.dwUserFlags == 0x12345678
                assert len(reference.listdir(streams=True, storages=False)) == len(generated["streams"])
                assert sorted(reference.listdir(streams=False, storages=True)) == [["Unknown"], ["Unknown", "Empty storage"]]
                for expected in generated["streams"]:
                    data = reference.openstream(expected["path"]).read()
                    assert len(data) == expected["size"]
                    assert hashlib.sha256(data).hexdigest() == expected["sha256"]
                assert reference.parsing_issues == [], reference.parsing_issues
            replacements = {("Unknown", "Small"): b"x" * 65, ("Exact",): b"y" * 4096}
            with olefile.OleFileIO(generated["path"], write_mode=True) as reference:
                for path, data in replacements.items():
                    reference.write_stream(list(path), data)
            observed = {tuple(item["path"]): item for item in node(READ, generated["path"])}
            for expected in generated["streams"]:
                path = tuple(expected["path"])
                wanted = hashlib.sha256(replacements[path]).hexdigest() if path in replacements else expected["sha256"]
                assert observed[path]["sha256"] == wanted
        print(f"CFB v3/v4 independent interop passed: olefile {olefile.__version__}; nested, empty, mini, 4096 boundary, 8MiB DIFAT; independent small/large stream edits")


if __name__ == "__main__":
    main()
