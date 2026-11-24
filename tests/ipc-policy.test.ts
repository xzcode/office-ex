// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { OfficeError } from '@office-ex/contracts';
import {
  TRUSTED_APP_ORIGIN, isLocalAppUrl, assertTrustedSender, assertIpcArity, ipcId,
  parseDroppedPath, parseTabStates, parseDraftSave, parseDocxSave, parsePreservedCopy, parseCheckpoint,
  parseRelease, parseSettingsPatch, parseRecoveryId, commandLineFiles,
} from '../apps/desktop/src/ipc-policy.js';
const fails = (code: string) => (error: unknown): boolean => error instanceof OfficeError && error.code === code;
const snapshot = { documentId: 'tab-1', contentId: 'document-1:content:42', revision: 42, bytes: new Uint8Array([1,2]) };
const save = { ...snapshot, saveAs: false, title: '中文草稿 😀' };
test('SEC-01/FS-01: only the live main frame of the expected WebContents may invoke IPC', () => {
  const evidence = { windowAlive: true, expectedWebContents: true, mainFrame: true, frameOrigin: TRUSTED_APP_ORIGIN };
  assert.doesNotThrow(() => assertTrustedSender(evidence));
  for (const patch of [{ windowAlive: false }, { expectedWebContents: false }, { mainFrame: false }, { frameOrigin: null }, { frameOrigin: 'office-ex://app.evil' }, { frameOrigin: 'https://app' }, { frameOrigin: 'null' }]) assert.throws(() => assertTrustedSender({ ...evidence, ...patch }), fails('unsafe'));
});
test('SEC-01: custom-scheme allowlist inspects scheme, hostname, port and credentials', () => {
  assert.equal(isLocalAppUrl('office-ex://app/index.html'), true);
  assert.equal(isLocalAppUrl('office-ex://app/format-worker.js'), true);
  for (const url of ['office-ex://app.evil/index.html','office-ex://user@app/index.html','office-ex://app:443/index.html','https://app/index.html','file:///etc/passwd','data:text/html,hello','office-ex://other/index.html','malformed']) assert.equal(isLocalAppUrl(url), false);
});
test('IPC argument counts reject unknown channels and ignored trailing payloads', () => {
  assert.doesNotThrow(() => assertIpcArity('office:release', ['tab-1', false]));
  for (const [channel, args] of [['office:initialize', [1]], ['office:save-draft', []], ['office:save-draft', [save, 'unexpected']], ['office:arbitrary-fs', ['/tmp/x']]] as const) assert.throws(() => assertIpcArity(channel, args), fails('invalid'));
});
test('IPC identifiers and native path strings cannot inject paths or shell arguments', () => {
  ipcId('doc-0123'); assert.equal(parseDroppedPath(resolve('example.docx')), resolve('example.docx'));
  for (const value of ['../other','a/b','',null,'x'.repeat(129)]) assert.throws(() => ipcId(value), fails('invalid'));
  for (const value of ['relative.docx', resolve('bad')+'\0', [], null]) assert.throws(() => parseDroppedPath(value), fails('invalid'));
});
test('IPC tab schema bounds counts and rejects duplicate IDs, coercion, accessors and unknown fields', () => {
  const tab = { id: 'tab-1', title: 'Draft', dirty: true, revision: 1 };
  assert.deepEqual(parseTabStates([tab]), [tab]);
  for (const invalid of [[tab,tab], [{...tab,dirty:'true'}], [{...tab,revision:NaN}], [{...tab,revision:-1}], [{...tab,extra:true}], [{...tab,get title(){throw new Error('must not execute');}}]]) assert.throws(() => parseTabStates(invalid), fails('invalid'));
  assert.throws(() => parseTabStates(Array.from({length:33},(_,i)=>({...tab,id:`doc-${i}`}))), fails('limit'));
});
test('IPC save/checkpoint requests carry real content IDs and bounded immutable byte messages', () => {
  assert.deepEqual(parseDraftSave(save), save);
  assert.deepEqual(parseCheckpoint(snapshot), snapshot);
  for (const invalid of [{...save,extra:1},{...save,saveAs:[]},{...save,contentId:''},{...save,revision:'42'},{...save,sourceToken:'../token'}]) assert.throws(() => parseDraftSave(invalid), fails('invalid'));
  for (const data of [new Uint8Array(),new Uint8Array(32*1024*1024+1),[1,2]]) assert.throws(() => parseDraftSave({...save,bytes:data}), fails('limit'));
  assert.throws(() => parseCheckpoint({...snapshot,bytes:new Uint8Array(new SharedArrayBuffer(2))}), fails('unsafe'));
});
test('IPC preserves separate save-copy, close and recovery capability shapes', () => {
  assert.deepEqual(parsePreservedCopy({documentId:'tab-1',sourceToken:'token-1'}), {documentId:'tab-1',sourceToken:'token-1'});
  assert.throws(() => parsePreservedCopy({documentId:'tab-1',sourceToken:'token-1',bytes:new Uint8Array([1])}), fails('invalid'));
  assert.deepEqual(parseRelease('tab-1',true), {documentId:'tab-1',discard:true});
  assert.throws(() => parseRelease('tab-1','true'), fails('invalid'));
  const recovery='document-1.42.00000000-0000-0000-0000-000000000000.oex-recovery';
  assert.equal(parseRecoveryId(recovery), recovery);
  for (const id of ['../'+recovery,'/tmp/'+recovery,'','other.json']) assert.throws(() => parseRecoveryId(id), fails('invalid'));
});
test('IPC settings validation is atomic and rejects arrays masquerading as enum strings', () => {
  const original={theme:'dark',language:'en-US',clearRecent:true};
  assert.deepEqual(parseSettingsPatch(original), original);
  for (const value of [{theme:['dark']},{language:['en']},{theme:'dark',language:'invalid'},{clearRecent:'true'},{path:'/tmp/settings'},{theme:null},{region:['de-DE']},{documentLanguage:'bad'},{language:'ar'}]) assert.throws(() => parseSettingsPatch(value), fails('invalid'));
  assert.deepEqual(original,{theme:'dark',language:'en-US',clearRecent:true});
});
test('OS-open argument selection ignores executable/app/switch/URL arguments and bounds queue size', () => {
  const cwd=resolve('test-documents');
  assert.deepEqual(commandLineFiles(['electron','/app/path','--inspect=9222','hello.docx','https://evil.example/a.docx','note.txt','hello.docx'],true,cwd), [resolve(cwd,'hello.docx')]);
  assert.deepEqual(commandLineFiles(['OfficeEX','deck.ppt'],false,cwd), [resolve(cwd,'deck.ppt')]);
  assert.equal(commandLineFiles(['OfficeEX',...Array.from({length:40},(_,i)=>`${i}.xls`)],false,cwd).length,32);
});

test('DOCX save IPC accepts only bounded snapshot requests bound to the registered document identity',()=>{
  const request={...save,contentId:'tab-1:content:42',sourceToken:'source-1'};
  assert.deepEqual(parseDocxSave(request),request);
  assert.doesNotThrow(()=>assertIpcArity('office:save-docx',[request]));
  assert.throws(()=>assertIpcArity('office:save-docx',[request,'extra']),fails('invalid'));
  for(const invalid of [
    {...request,contentId:'different:content:42'},
    {...request,contentId:'tab-1:content:'},
    {...request,contentId:'tab-1:content:1x'},
    {...request,sourceBytes:new Uint8Array([1])},
    {...request,targetPath:'/tmp/file.docx'}
  ])assert.throws(()=>parseDocxSave(invalid),fails('invalid'));
  assert.throws(()=>parseDocxSave({...request,bytes:new Uint8Array(new SharedArrayBuffer(4))}),fails('unsafe'));
});
