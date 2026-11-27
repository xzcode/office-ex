import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { OfficeError, assertCellAddress, graphemeBoundaries, replaceGraphemeRange } from '@office-ex/contracts';
import { DocumentSession, DerivedTaskGate, ResourceStore, createDocument, migrateDocument, patchCommand, resolveTextStyle, sha256, validateDocument, type DocumentCommand, type OfficeDocument } from '@office-ex/core';
import { replaceText } from '@office-ex/writer';
import { renameSheet, setCells } from '@office-ex/sheets';
import { insertShapes, setShapeGeometry } from '@office-ex/slides';

const code = (expected: string) => (error: unknown): boolean => error instanceof OfficeError && error.code === expected;
const rename = (session: DocumentSession, title: string): DocumentCommand => ({ documentId: session.documentId, expectedRevision: session.revision, type: 'document.rename', payload: { title } });

test('MOD-01: three canonical serializable models validate and have stable IDs', () => {
  for (const kind of ['writer', 'sheets', 'slides'] as const) {
    const original = createDocument(kind, `document-${kind}`, '中文 Office EX');
    assert.deepEqual(validateDocument(JSON.parse(JSON.stringify(original))), original);
    assert.equal(original.schemaVersion, 1);
  }
  const document = createDocument('writer', 'writer-1');
  const session = new DocumentSession(document, { saved: true });
  document.title = 'Mutated caller copy';
  assert.equal(session.snapshot().document.title, 'Untitled');
  assert.throws(() => { (session.snapshot().document as OfficeDocument).title = 'Mutable snapshot'; }, TypeError);
  assert.equal(session.dirty, false);
});

test('MOD-01/ARC-02: reject unknown schema, duplicate IDs, accessors, cycles and unsafe keys', () => {
  const future = { ...createDocument('writer', 'future-1'), schemaVersion: 2 };
  assert.throws(() => validateDocument(future), code('unsupported'));
  const duplicate = createDocument('writer', 'duplicate-1');
  duplicate.sections[0]!.blocks[0]!.id = duplicate.sections[0]!.id;
  assert.throws(() => validateDocument(duplicate), code('invalid'));
  const unsafe = JSON.parse(JSON.stringify(createDocument('writer', 'unsafe-1'))) as Record<string, unknown>;
  unsafe.resources = JSON.parse('{"__proto__":{"polluted":true}}');
  assert.throws(() => validateDocument(unsafe), code('unsafe'));
  const accessor = createDocument('writer', 'accessor-1');
  Object.defineProperty(accessor, 'title', { get: () => { throw new Error('Accessor executed'); } });
  assert.throws(() => validateDocument(accessor), code('invalid'));
  const cyclic: Record<string, unknown> = { schemaVersion: 1 }; cyclic.self = cyclic;
  assert.throws(() => validateDocument(cyclic), code('invalid'));
});

test('MOD-03: a failing multi-operation transaction has no model, history or dirty side effects', () => {
  const session = new DocumentSession(createDocument('sheets', 'atomic-1'), { saved: true });
  const before = session.snapshot();
  const command = patchCommand(before, 'SHT-02', [
    { op: 'set', path: ['sheets', 0, 'cells', '1:1'], value: { value: { kind: 'string', value: 'must roll back' } } },
    { op: 'set', path: ['sheets', 0, 'cells', '1048577:1'], value: { value: { kind: 'number', value: 1 } } },
  ]);
  assert.throws(() => session.execute(command), code('invalid'));
  assert.deepEqual(session.snapshot(), before); assert.equal(session.canUndo, false); assert.equal(session.dirty, false);
  assert.deepEqual(session.historyUsage, { items: 0, bytes: 0 });
  assert.throws(() => session.execute({ ...rename(session, 'new'), expectedRevision: 999 }), code('conflict'));
  assert.throws(() => session.execute({ ...rename(session, 'new'), documentId: 'other-doc' }), code('conflict'));
  assert.equal(session.revision, 0);
});

test('MOD-03: undo/redo revisions advance while saved-content identity restores clean state', () => {
  const session = new DocumentSession(createDocument('writer', 'history-1'), { saved: true });
  const originalContentId = session.contentId;
  session.execute(rename(session, 'Edited')); assert.equal(session.dirty, true);
  const editedContentId = session.contentId;
  session.undo(); assert.equal(session.revision, 2); assert.equal(session.contentId, originalContentId); assert.equal(session.dirty, false);
  session.redo(); assert.equal(session.revision, 3); assert.equal(session.contentId, editedContentId); assert.equal(session.dirty, true);
  session.undo(); session.execute(rename(session, 'Different edit')); assert.equal(session.canRedo, false); assert.equal(session.revision, 5);
  const noOp = session.execute(rename(session, 'Different edit')); assert.equal(noOp.changed, false); assert.equal(session.revision, 5);
});

test('MOD-03: incremental history is bounded and snapshots share unchanged branches', () => {
  const document = createDocument('writer', 'sharing-1');
  document.sections.push({ id: 'section-2', page: structuredClone(document.sections[0]!.page), blocks: [{ id: 'paragraph-2', type: 'paragraph', runs: [{ id: 'run-2', type: 'text', text: 'Unchanged' }] }] });
  const session = new DocumentSession(document, { saved: true, historyItems: 2 });
  const before = session.snapshot().document; assert.equal(before.kind, 'writer');
  session.execute(replaceText(session.snapshot(), 'run-1', 0, 0, 'hello'));
  const after = session.snapshot().document; assert.equal(after.kind, 'writer');
  if (before.kind === 'writer' && after.kind === 'writer') assert.equal(before.sections[1], after.sections[1]);
  session.execute(rename(session, 'one')); session.execute(rename(session, 'two'));
  assert.equal(session.historyUsage.items, 2); session.undo(); session.undo(); assert.equal(session.canUndo, false);
  const bounded = new DocumentSession(createDocument('writer', 'small-history'), { saved: true, historyBytes: 80 });
  assert.throws(() => bounded.execute(rename(bounded, 'x'.repeat(100))), code('limit')); assert.equal(bounded.revision, 0); assert.equal(bounded.dirty, false);
});

test('MOD-04: save tickets keep edits during asynchronous saves dirty and reject forged completion', () => {
  const session = new DocumentSession(createDocument('writer', 'save-1'));
  const ticket = session.beginSave(); assert.equal(session.saveState, 'saving');
  session.execute(rename(session, 'Edited during save'));
  assert.throws(() => session.beginSave(), code('conflict'));
  assert.throws(() => session.completeSave({ ...ticket }), code('conflict'));
  session.completeSave(ticket); assert.equal(session.dirty, true); assert.equal(session.saveState, 'idle');
  session.undo(); assert.equal(session.dirty, false);
  const next = session.beginSave(); session.cancelSave(next); assert.equal(session.dirty, false);
  assert.throws(() => session.completeSave(next), code('conflict'));
});

test('MOD-03: document sessions, histories, and feature permissions remain isolated', () => {
  const first = new DocumentSession(createDocument('writer', 'first-1'), { saved: true });
  const second = new DocumentSession(createDocument('writer', 'second-1'), { saved: true });
  first.execute(rename(first, 'First only')); assert.equal(second.revision, 0); assert.equal(second.canUndo, false);
  first.setCapabilities({ canEdit: false, allowedFeatures: [], diagnostics: [] });
  assert.throws(() => first.undo(), code('unsupported')); assert.throws(() => first.execute(rename(first, 'blocked')), code('unsupported'));
  second.setCapabilities({ canEdit: true, allowedFeatures: ['WRT-03'], diagnostics: [] });
  assert.throws(() => second.execute(patchCommand(second.snapshot(), 'WRT-03', [{ op: 'set', path: ['title'], value: 'wrong feature' }])), code('unsupported'));
  first.close(); assert.throws(() => first.snapshot(), code('conflict')); assert.equal(second.documentId, 'second-1');
});

test('MOD-02/WRT-03: IME-sized text edits preserve emoji, combining sequences, UTF-16 and undo', () => {
  const text = 'A👩🏽‍💻é中';
  assert.deepEqual(graphemeBoundaries(text), [0, 1, 8, 10, 11]);
  assert.throws(() => replaceGraphemeRange(text, 2, 3, ''), code('invalid'));
  assert.throws(() => replaceGraphemeRange(text, 8, 9, ''), code('invalid'));
  assert.equal(replaceGraphemeRange(text, 1, 8, '🙂'), 'A🙂é中');
  const session = new DocumentSession(createDocument('writer', 'ime-1'), { saved: true });
  session.execute(replaceText(session.snapshot(), 'run-1', 0, 0, text));
  assert.equal(session.historyUsage.items, 1); session.undo(); assert.equal(session.dirty, false);
});

test('SHT-01/MOD-02: sparse typed values retain empty string, zero and formula source, reject XLS overflow', () => {
  const session = new DocumentSession(createDocument('sheets', 'cells-1'), { saved: true });
  session.execute(setCells(session.snapshot(), 'sheet-1', [
    { row: 1, column: 1, cell: { value: { kind: 'string', value: '' } } },
    { row: 2, column: 1, cell: { value: { kind: 'number', value: 0 } } },
    { row: 1_048_576, column: 16_384, cell: { value: { kind: 'formula', source: '=SUM(A1:A2)' } } },
  ]));
  const snapshot = session.snapshot(); assert.equal(snapshot.document.kind, 'sheets');
  if (snapshot.document.kind === 'sheets') assert.equal(Object.keys(snapshot.document.sheets[0]!.cells).length, 3);
  assert.throws(() => assertCellAddress(65_537, 1, 'xls'), code('invalid'));
  assert.throws(() => assertCellAddress(1, 257, 'xls'), code('invalid'));
  session.execute(renameSheet(session.snapshot(), 'sheet-1', 'renamed')); assert.equal((session.snapshot().document as typeof snapshot.document).kind, 'sheets'); session.undo();
  session.execute(setCells(session.snapshot(), 'sheet-1', [{ row: 2, column: 1, cell: null }]));
  session.undo(); session.undo(); assert.equal(session.dirty, false);
});

test('SLD-01/MOD-03: shape order and stable IDs survive a single drag transaction and undo', () => {
  const session = new DocumentSession(createDocument('slides', 'slides-1'), { saved: true });
  const original = { x: 0, y: 0, width: 914400, height: 914400, rotation: 0 };
  session.execute(insertShapes(session.snapshot(), 'slide-1', 0, [{ id: 'shape-1', type: 'rect', geometry: original }]));
  session.execute(setShapeGeometry(session.snapshot(), 'slide-1', 'shape-1', { ...original, x: 100, rotation: 90 * 60_000 }));
  session.undo(); const snapshot = session.snapshot(); assert.equal(snapshot.document.kind, 'slides');
  if (snapshot.document.kind === 'slides') { assert.equal(snapshot.document.slides[0]!.shapes[0]!.id, 'shape-1'); assert.deepEqual(snapshot.document.slides[0]!.shapes[0]!.geometry, original); }
});

test('MOD-04: derived work cancels superseded tasks and discards stale revisions', () => {
  const session = new DocumentSession(createDocument('writer', 'derive-1')); const gate = new DerivedTaskGate<string>();
  const first = gate.start(session.snapshot()); const second = gate.start(session.snapshot()); assert.equal(first.cancellation.cancelled, true);
  assert.throws(() => first.cancellation.throwIfCancelled(), code('cancelled'));
  assert.equal(gate.accept(first, session.snapshot(), 'old'), false);
  session.execute(rename(session, 'New revision')); assert.equal(gate.accept(second, session.snapshot(), 'stale'), false);
  const current = gate.start(session.snapshot()); assert.equal(gate.accept(current, session.snapshot(), 'current'), true);
  assert.equal(gate.resultFor(session.revision), 'current'); assert.equal(gate.resultFor(session.revision - 1), undefined);
  gate.clear(); assert.equal(gate.resultFor(session.revision), undefined);
});

test('MOD-01: schema migrations preserve their input on failure and refuse unknown future versions', () => {
  const source = { ...createDocument('writer', 'migration-1'), schemaVersion: 0 }; const original = structuredClone(source);
  assert.throws(() => migrateDocument(source, [{ from: 0, to: 1, migrate: (copy) => { (copy as { title: string }).title = 'damaged'; throw new Error('failed migration'); } }]), /failed migration/);
  assert.deepEqual(source, original);
  const migrated = migrateDocument(source, [{ from: 0, to: 1, migrate: (copy) => ({ ...(copy as Record<string, unknown>), schemaVersion: 1 }) }]);
  assert.equal(migrated.schemaVersion, 1); assert.equal(source.schemaVersion, 0);
  assert.throws(() => migrateDocument({ ...source, schemaVersion: 2 }), code('unsupported'));
});

test('MOD-01: SHA-256 agrees with independent Node crypto across padding/block boundaries', () => {
  for (const length of [0, 3, 55, 56, 63, 64, 65, 127, 128, 1024]) {
    const bytes = Uint8Array.from({ length }, (_, index) => index * 17 % 256);
    assert.equal(sha256(bytes), createHash('sha256').update(bytes).digest('hex'));
  }
  assert.equal(sha256(new TextEncoder().encode('abc')), 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
});

test('MOD-01: resource ownership validates hashes, budgets, identity and complete cleanup proof', () => {
  const resources = new ResourceStore(10); const bytes = new Uint8Array([1, 2, 3]);
  const metadata = resources.put('image-1', bytes, 'image/png'); bytes[0] = 99;
  assert.deepEqual(resources.get('image-1'), new Uint8Array([1, 2, 3]));
  const copy = resources.get('image-1'); copy[0] = 77; assert.equal(resources.get('image-1')[0], 1);
  assert.equal(resources.put('image-1', new Uint8Array([1, 2, 3]), 'image/png'), metadata);
  assert.throws(() => resources.put('image-1', new Uint8Array([4]), 'image/png'), code('conflict'));
  assert.throws(() => resources.put('image-2', new Uint8Array(8), 'image/png'), code('limit'));
  assert.throws(() => resources.put('image-2', new Uint8Array([2]), 'image/png', '0'.repeat(64)), code('invalid'));
  assert.throws(() => resources.removeUnreferenced(['image-1'], { complete: false, referencedIds: [] }), code('unsupported'));
  assert.throws(() => resources.removeUnreferenced(['image-1'], { complete: true, referencedIds: ['image-1'] }), code('conflict'));
  resources.removeUnreferenced(['image-1'], { complete: true, referencedIds: [] }); assert.equal(resources.byteLength, 0);
});

test('WRT-02: style cascade resolves inheritance, direct formatting, cycles and missing references', () => {
  const styles = { normal: { id: 'normal', name: 'Normal', properties: { fontFamily: 'Example', bold: true } }, heading: { id: 'heading', name: 'Heading', basedOn: 'normal', properties: { fontSize: 254000 } } };
  const result = resolveTextStyle({ fontSize: 152400 }, styles, ['heading'], { bold: false });
  assert.deepEqual(result.style, { fontSize: 254000, fontFamily: 'Example', bold: false });
  assert.equal(resolveTextStyle({}, styles, ['missing']).diagnostics[0]?.code, 'style.missing');
  assert.throws(() => resolveTextStyle({}, { loop: { id: 'loop', name: 'Loop', basedOn: 'loop', properties: {} } }, ['loop']), code('invalid'));
});
