import { CancellationToken, assert, assertId, isRecord, type Diagnostic } from '@office-ex/contracts';
import { assertPlainData, cloneData, freezeData, validateDocument, type OfficeDocument } from './model.js';
import { validateImageResource } from './images.js';

export type DeepReadonly<T> = T extends readonly (infer U)[] ? readonly DeepReadonly<U>[] : T extends object ? { readonly [K in keyof T]: DeepReadonly<T[K]> } : T;
export type ModelPath = readonly (string | number)[];
export type PatchOperation =
  | { readonly op: 'set'; readonly path: ModelPath; readonly value: unknown }
  | { readonly op: 'delete'; readonly path: ModelPath }
  | { readonly op: 'splice'; readonly path: ModelPath; readonly index: number; readonly deleteCount: number; readonly items: readonly unknown[] };
export interface CommandEnvelope { readonly documentId: string; readonly expectedRevision: number }
export type DocumentCommand = CommandEnvelope & (
  | { readonly type: 'document.rename'; readonly payload: { readonly title: string } }
  | { readonly type: 'model.patch'; readonly payload: { readonly featureId: string; readonly operations: readonly PatchOperation[] } }
);
export interface SessionCapabilities {
  readonly canEdit: boolean;
  /** Empty means no model-changing feature is enabled, never "allow everything". */
  readonly allowedFeatures: readonly string[];
  readonly diagnostics: readonly Diagnostic[];
}
export const EDITABLE_CAPABILITIES: SessionCapabilities = Object.freeze({ canEdit: true, allowedFeatures: Object.freeze(['MOD-01', 'MOD-03', 'WRT-02', 'WRT-03', 'SHT-02', 'SLD-02']), diagnostics: Object.freeze([]) });
export interface DocumentSnapshot { readonly document: DeepReadonly<OfficeDocument>; readonly revision: number; readonly contentId: string }
export interface SaveTicket extends DocumentSnapshot { readonly id: string; readonly documentId: string }
export interface TransactionResult {
  readonly changed: boolean; readonly revision: number; readonly contentId: string; readonly dirty: boolean;
  readonly changedPaths: readonly ModelPath[]; readonly inverse: readonly PatchOperation[];
}
interface HistoryEntry { readonly forward: readonly PatchOperation[]; readonly inverse: readonly PatchOperation[]; readonly before: string; readonly after: string; readonly bytes: number; readonly featureId: string }
export interface SessionOptions { readonly saved?: boolean; readonly capabilities?: SessionCapabilities; readonly historyItems?: number; readonly historyBytes?: number }

function sameData(a: unknown, b: unknown): boolean {
  if (Object.is(a, b)) return true;
  if (Array.isArray(a) && Array.isArray(b)) return a.length === b.length && a.every((item, index) => sameData(item, b[index]));
  if (isRecord(a) && isRecord(b)) { const keys = Object.keys(a); return keys.length === Object.keys(b).length && keys.every((key) => Object.hasOwn(b, key) && sameData(a[key], b[key])); }
  return false;
}
function validatePath(path: unknown): asserts path is ModelPath {
  assert(Array.isArray(path) && path.length > 0 && path.length <= 64, 'invalid', 'Invalid model path');
  for (const key of path) assert((typeof key === 'string' && key.length > 0 && key.length <= 128 && !['__proto__', 'constructor', 'prototype'].includes(key)) || (typeof key === 'number' && Number.isSafeInteger(key) && key >= 0), 'unsafe', 'Unsafe model path');
  assert(!['schemaVersion', 'documentId', 'kind'].includes(String(path[0])), 'invalid', 'Document identity cannot be changed by a command');
}
function assertOperation(value: unknown): asserts value is PatchOperation {
  assert(isRecord(value), 'invalid', 'Invalid patch operation'); validatePath(value.path);
  if (value.op === 'set') { assert(Object.keys(value).length === 3 && Object.hasOwn(value, 'value'), 'invalid', 'Invalid set patch'); assertPlainData(value.value); }
  else if (value.op === 'delete') assert(Object.keys(value).length === 2, 'invalid', 'Invalid delete patch');
  else {
    assert(value.op === 'splice' && Object.keys(value).length === 5 && Number.isSafeInteger(value.index) && Number(value.index) >= 0 && Number.isSafeInteger(value.deleteCount) && Number(value.deleteCount) >= 0 && Array.isArray(value.items), 'invalid', 'Invalid splice patch');
    assertPlainData(value.items);
  }
}
function child(container: unknown, key: string | number): unknown {
  assert(Array.isArray(container) || isRecord(container), 'invalid', 'Model path does not name a container');
  if (Array.isArray(container)) { assert(typeof key === 'number' && key < container.length, 'invalid', 'Array path is out of bounds'); return container[key]; }
  assert(typeof key === 'string' && Object.hasOwn(container, key), 'invalid', 'Model path does not exist'); return container[key];
}
function atPath(root: unknown, path: ModelPath): unknown { let current = root; for (const key of path) current = child(current, key); return current; }
function replaceAt(root: unknown, path: ModelPath, value: unknown, remove = false): unknown {
  const [key, ...rest] = path;
  assert(key !== undefined && (Array.isArray(root) || isRecord(root)), 'invalid', 'Invalid patch container');
  if (Array.isArray(root)) {
    assert(typeof key === 'number' && key < root.length && (!remove || rest.length > 0), 'invalid', 'Array entries must be removed with splice');
    const copy: unknown[] = root.slice(); copy[key] = rest.length ? replaceAt(child(root, key), rest, value, remove) : value; return copy;
  }
  assert(typeof key === 'string', 'invalid', 'Object paths require string keys');
  const copy: Record<string, unknown> = { ...root };
  if (rest.length) copy[key] = replaceAt(child(root, key), rest, value, remove);
  else if (remove) delete copy[key]; else copy[key] = value;
  return copy;
}
function applyPatch(root: OfficeDocument, patch: PatchOperation): { document: OfficeDocument; inverse?: PatchOperation } {
  if (patch.op === 'splice') {
    const target: unknown = atPath(root, patch.path); assert(Array.isArray(target), 'invalid', 'Splice target is not an array');
    assert(patch.index <= target.length && patch.deleteCount <= target.length - patch.index, 'invalid', 'Splice range is out of bounds');
    const removed = target.slice(patch.index, patch.index + patch.deleteCount);
    if (sameData(removed, patch.items)) return { document: root };
    const copy = target.slice(0, patch.index).concat(cloneData(patch.items), target.slice(patch.index + patch.deleteCount));
    return { document: replaceAt(root, patch.path, copy) as OfficeDocument, inverse: { op: 'splice', path: patch.path, index: patch.index, deleteCount: patch.items.length, items: removed } };
  }
  const parentPath = patch.path.slice(0, -1); const key = patch.path[patch.path.length - 1]!;
  const parent: unknown = parentPath.length ? atPath(root, parentPath) : root;
  assert(isRecord(parent) || Array.isArray(parent), 'invalid', 'Invalid patch parent');
  if (Array.isArray(parent)) assert(typeof key === 'number' && key < parent.length && patch.op !== 'delete', 'invalid', 'Invalid array patch');
  else assert(typeof key === 'string', 'invalid', 'Invalid object patch');
  const exists = Object.hasOwn(parent, key);
  const before = exists ? child(parent, key) : undefined;
  if (patch.op === 'delete') {
    if (!exists) return { document: root };
    return { document: replaceAt(root, patch.path, undefined, true) as OfficeDocument, inverse: { op: 'set', path: patch.path, value: before } };
  }
  if (exists && sameData(before, patch.value)) return { document: root };
  return { document: replaceAt(root, patch.path, cloneData(patch.value)) as OfficeDocument, inverse: exists ? { op: 'set', path: patch.path, value: before } : { op: 'delete', path: patch.path } };
}
function applyBatch(document: OfficeDocument, operations: readonly PatchOperation[]): { document: OfficeDocument; forward: PatchOperation[]; inverse: PatchOperation[] } {
  let next = document;
  const forward: PatchOperation[] = []; const inverse: PatchOperation[] = [];
  // Coalesce consecutive direct edits to one sheet's cell map. Clone the
  // map once, preserve unaffected immutable cells and keep the history
  // at individual-cell granularity. Every patch is still checked and the
  // complete candidate is validated before committing.
  const directCell=(operation:PatchOperation)=>operation.op!=='splice'&&operation.path.length===4&&operation.path[0]==='sheets'&&typeof operation.path[1]==='number'&&operation.path[2]==='cells'&&typeof operation.path[3]==='string';
  for(let i=0;i<operations.length;){
    const operation=operations[i]!;assertOperation(operation);
    if(next.kind==='sheets'&&directCell(operation)){
      const parentPath=operation.path.slice(0,3),parent=atPath(next,parentPath);
      assert(isRecord(parent),'invalid','Invalid cell patch parent');
      const copy={...parent};let changed=false,j=i;
      for(;j<operations.length;j++){
        const patch=operations[j]!;assertOperation(patch);
        if(!directCell(patch)||patch.path[1]!==operation.path[1])break;
        const key=patch.path[3] as string,exists=Object.hasOwn(copy,key),before=copy[key];
        if(patch.op==='delete'){
          if(!exists)continue;
          delete copy[key];inverse.push({op:'set',path:patch.path,value:before});
        }else if(patch.op==='set'){
          if(exists&&sameData(before,patch.value))continue;
          copy[key]=cloneData(patch.value);
          inverse.push(exists?{op:'set',path:patch.path,value:before}:{op:'delete',path:patch.path});
        }else throw new Error('Unexpected cell splice');
        forward.push(cloneData(patch));changed=true;
      }
      if(changed)next=replaceAt(next,parentPath,copy) as OfficeDocument;
      i=j;
    }else{
      const result=applyPatch(next,operation);
      if(result.inverse){forward.push(cloneData(operation));inverse.push(result.inverse);next=result.document;}
      i++;
    }
  }
  inverse.reverse();
  if (sameData(document, next)) return { document, forward: [], inverse: [] };
  // Candidate state is detached by structural sharing. Validation cannot alter live state.
  validateDocument(next);
  return { document: next, forward, inverse };
}

/** Preview a validated candidate for a synchronous layout-dependent command.
 * This is not a mutation authority; final execution still checks capabilities. */
export function previewModelPatch(snapshot:DocumentSnapshot,command:DocumentCommand):DocumentSnapshot{
  assert(command.type==='model.patch'&&command.documentId===snapshot.document.documentId&&command.expectedRevision===snapshot.revision,'conflict','Invalid preview command');
  assert(command.payload.operations.length<=10000,'limit','Preview operation budget exceeded');
  const candidate=applyBatch(snapshot.document as OfficeDocument,command.payload.operations);
  return {...snapshot,document:candidate.document};
}

/** A single document authority; history retains deltas, not whole-document snapshots. */
export class DocumentSession {
  #document: OfficeDocument; #revision = 0; #contentCounter = 0; #contentId: string;
  #lastSavedContentId: string | null; #past: HistoryEntry[] = []; #future: HistoryEntry[] = [];
  #historyItems: number; #historyBytes: number; #capabilities: SessionCapabilities;
  #saveCounter = 0; #activeSave: SaveTicket | undefined;
  #closed = false;
  constructor(document: unknown, options: SessionOptions = {}) {
    this.#document = freezeData(cloneData(validateDocument(document)));
    this.#contentId = `${this.#document.documentId}:content:0`;
    this.#lastSavedContentId = options.saved ? this.#contentId : null;
    this.#historyItems = options.historyItems ?? 200; this.#historyBytes = options.historyBytes ?? 64 * 1024 * 1024;
    assert(Number.isSafeInteger(this.#historyItems) && this.#historyItems >= 1 && this.#historyItems <= 100_000 && Number.isSafeInteger(this.#historyBytes) && this.#historyBytes > 0, 'invalid', 'Invalid history budget');
    this.#capabilities = freezeData(cloneData(options.capabilities ?? EDITABLE_CAPABILITIES));
    this.#validateCapabilities(this.#capabilities);
  }
  #validateCapabilities(capabilities: SessionCapabilities): void {
    assertPlainData(capabilities); assert(typeof capabilities.canEdit === 'boolean' && Array.isArray(capabilities.allowedFeatures) && capabilities.allowedFeatures.every((id) => typeof id === 'string' && /^[A-Z]+-\d+$/.test(id)) && Array.isArray(capabilities.diagnostics), 'invalid', 'Invalid session capabilities');
  }
  #assertOpen(): void { assert(!this.#closed, 'conflict', 'Document session is closed'); }
  #assertEditable(featureId: string): void { this.#assertOpen(); assert(this.#capabilities.canEdit && this.#capabilities.allowedFeatures.includes(featureId), 'unsupported', 'This document does not allow that edit'); }
  get documentId(): string { return this.#document.documentId; }
  get revision(): number { return this.#revision; }
  get contentId(): string { return this.#contentId; }
  get lastSavedContentId(): string | null { return this.#lastSavedContentId; }
  get dirty(): boolean { return this.#contentId !== this.#lastSavedContentId; }
  get canUndo(): boolean { return this.#past.length > 0 && this.#capabilities.canEdit && !this.#closed; }
  get canRedo(): boolean { return this.#future.length > 0 && this.#capabilities.canEdit && !this.#closed; }
  get saveState(): 'idle' | 'saving' { return this.#activeSave ? 'saving' : 'idle'; }
  get capabilities(): SessionCapabilities { return this.#capabilities; }
  get historyUsage(): { items: number; bytes: number } { const all = [...this.#past, ...this.#future]; return { items: all.length, bytes: all.reduce((total, entry) => total + entry.bytes, 0) }; }
  snapshot(): DocumentSnapshot { this.#assertOpen(); return Object.freeze({ document: this.#document, revision: this.#revision, contentId: this.#contentId }); }
  setCapabilities(capabilities: SessionCapabilities): void { this.#assertOpen(); this.#validateCapabilities(capabilities); this.#capabilities = freezeData(cloneData(capabilities)); }
  #result(changed: boolean, operations: readonly PatchOperation[] = [], inverse: readonly PatchOperation[] = []): TransactionResult {
    return freezeData({ changed, revision: this.#revision, contentId: this.#contentId, dirty: this.dirty, changedPaths: operations.map((operation) => [...operation.path]), inverse: cloneData(inverse) });
  }
  execute(command: DocumentCommand): TransactionResult {
    this.#assertOpen(); assertPlainData(command);
    assert(isRecord(command) && Object.keys(command).length === 4 && command.documentId === this.documentId && command.expectedRevision === this.#revision, 'conflict', 'Command targets a stale or different document');
    assert(isRecord(command.payload), 'invalid', 'Invalid command payload');
    let featureId: string; let operations: readonly PatchOperation[];
    if (command.type === 'document.rename') {
      assert(Object.keys(command.payload).length === 1 && typeof command.payload.title === 'string', 'invalid', 'Invalid rename command');
      featureId = 'MOD-03'; operations = [{ op: 'set', path: ['title'], value: command.payload.title }];
    } else {
      assert(command.type === 'model.patch' && Object.keys(command.payload).length === 2 && typeof command.payload.featureId === 'string' && Array.isArray(command.payload.operations), 'invalid', 'Unknown command or invalid patch payload');
      featureId = command.payload.featureId; operations = command.payload.operations;
      assert(operations.length <= 10_000, 'limit', 'Command operation budget exceeded');
    }
    this.#assertEditable(featureId);
    const permittedRoots: Record<string, readonly string[]> = { 'MOD-01': ['resources', 'styles', 'defaultTextStyle'], 'MOD-03': ['title'], 'WRT-02': ['sections'], 'WRT-03': ['sections', 'resources','numbering','pageStories','differentEvenOddPages'], 'SHT-02': ['sheets', 'dateSystem'], 'SLD-02': ['slides', 'size', 'theme', 'resources'] };
    const requiredKind: Record<string, string> = { 'WRT-02': 'writer', 'WRT-03': 'writer', 'SHT-02': 'sheets', 'SLD-02': 'slides' };
    assert(!requiredKind[featureId] || requiredKind[featureId] === this.#document.kind, 'invalid', 'Command feature does not match document kind');
    const newImages=new Set<string>();let numberingAdded=false;
    for (const operation of operations) {
      assertOperation(operation);
      assert(permittedRoots[featureId]?.includes(String(operation.path[0])), 'unsupported', 'Feature cannot modify that model region');
      if(featureId==='WRT-02'){
        const path=operation.path,index=path[1];
        const pageUpdate=operation.op==='set'&&path[2]==='page';
        const materialize=operation.op==='delete'&&path[2]==='nativeImplicitMargins'&&this.#document.kind==='writer'&&typeof index==='number'&&this.#document.sections[index]?.nativeImplicitMargins===true&&operations.some(p=>p.op==='set'&&p.path.length===3&&p.path[0]==='sections'&&p.path[1]===index&&p.path[2]==='page');
        assert(path.length===3&&path[0]==='sections'&&typeof index==='number'&&(pageUpdate||materialize),'unsupported','Page command can only replace section page geometry and materialize its implicit margins');
      }
      if(featureId==='WRT-03'&&operation.path[0]==='numbering'){
        assert(!numberingAdded&&operation.op==='set'&&operation.path.length===1&&this.#document.kind==='writer','unsupported','Numbering commands may append definitions only');
        const old=this.#document.numbering,registry=operation.value as import('./model.js').WriterNumbering;
        assert(registry&&typeof registry==='object'&&registry.lists&&typeof registry.lists==='object'&&Number.isSafeInteger(registry.nextId),'invalid','Invalid numbering registry');
        const first=old?.nextId??1;assert(registry.nextId>=first&&registry.nextId-first<=2048,'limit','Invalid numbering allocation');
        for(const [id,def]of Object.entries(old?.lists??{}))assert(JSON.stringify(registry.lists[id])===JSON.stringify(def),'conflict','Numbering definitions are immutable');
        const fresh=Object.keys(registry.lists).filter(id=>!Object.hasOwn(old?.lists??{},id));
        assert(fresh.length===registry.nextId-first&&fresh.every(id=>Number(id)>=first&&Number(id)<registry.nextId),'invalid','Noncontiguous numbering allocation');
        numberingAdded=true;
      }
      if((featureId==='WRT-03'||featureId==='SLD-02')&&operation.path[0]==='resources'){
        assert(operation.op==='set'&&operation.path.length===2&&typeof operation.path[1]==='string','unsupported','Image commands may only add an immutable image resource');
        validateImageResource(operation.value);const id=operation.path[1];
        assert(id===operation.value.id&&!Object.hasOwn(this.#document.resources,id)&&!newImages.has(id),'conflict','Cannot replace an existing resource');
        newImages.add(id);
      }
    }
    const batch = applyBatch(this.#document, operations);
    if (batch.forward.length === 0) return this.#result(false);
    const bytes = new TextEncoder().encode(JSON.stringify([batch.forward, batch.inverse])).byteLength;
    assert(bytes <= this.#historyBytes, 'limit', 'Edit exceeds the undo history budget');
    const after = `${this.documentId}:content:${this.#contentCounter + 1}`;
    const entry: HistoryEntry = freezeData({ forward: batch.forward, inverse: batch.inverse, before: this.#contentId, after, bytes, featureId });
    // Commit only after every operation, schema, permission, and history check succeeds.
    this.#document = freezeData(batch.document); this.#revision++; this.#contentCounter++; this.#contentId = after;
    this.#future = []; this.#past.push(entry);
    while (this.#past.length > this.#historyItems || this.historyUsage.bytes > this.#historyBytes) this.#past.shift();
    return this.#result(true, batch.forward, batch.inverse);
  }
  undo(expectedRevision = this.#revision): TransactionResult {
    this.#assertOpen(); assert(this.#capabilities.canEdit, 'unsupported', 'Read-only document'); assert(expectedRevision === this.#revision, 'conflict', 'Stale undo request');
    const entry = this.#past[this.#past.length - 1]; if (!entry) return this.#result(false);
    this.#assertEditable(entry.featureId);
    const batch = applyBatch(this.#document, entry.inverse);
    this.#document = freezeData(batch.document); this.#past.pop(); this.#future.push(entry); this.#revision++; this.#contentId = entry.before;
    return this.#result(true, entry.inverse, entry.forward);
  }
  redo(expectedRevision = this.#revision): TransactionResult {
    this.#assertOpen(); assert(this.#capabilities.canEdit, 'unsupported', 'Read-only document'); assert(expectedRevision === this.#revision, 'conflict', 'Stale redo request');
    const entry = this.#future[this.#future.length - 1]; if (!entry) return this.#result(false);
    this.#assertEditable(entry.featureId);
    const batch = applyBatch(this.#document, entry.forward);
    this.#document = freezeData(batch.document); this.#future.pop(); this.#past.push(entry); this.#revision++; this.#contentId = entry.after;
    return this.#result(true, entry.forward, entry.inverse);
  }
  beginSave(): SaveTicket {
    this.#assertOpen(); assert(!this.#activeSave, 'conflict', 'A save is already in progress');
    const ticket = Object.freeze({ ...this.snapshot(), id: `${this.documentId}:save:${++this.#saveCounter}`, documentId: this.documentId });
    this.#activeSave = ticket; return ticket;
  }
  completeSave(ticket: SaveTicket): void {
    this.#assertOpen(); assert(this.#activeSave === ticket, 'conflict', 'Save ticket is stale or belongs to another session');
    this.#lastSavedContentId = ticket.contentId; this.#activeSave = undefined;
  }
  cancelSave(ticket: SaveTicket): void { this.#assertOpen(); assert(this.#activeSave === ticket, 'conflict', 'Save ticket is stale'); this.#activeSave = undefined; }
  close(): void { this.#closed = true; this.#activeSave = undefined; this.#past = []; this.#future = []; }
}

export interface DerivedTask { readonly id: string; readonly revision: number; readonly documentId: string; readonly cancellation: CancellationToken }
/** One replaceable task per computation slot, with explicit cancellation and stale-result rejection. */
export class DerivedTaskGate<T> {
  #counter = 0; #active: DerivedTask | undefined; #result: { revision: number; value: T } | undefined;
  start(snapshot: DocumentSnapshot): DerivedTask {
    this.cancel(); const task = Object.freeze({ id: `${snapshot.document.documentId}:task:${++this.#counter}`, revision: snapshot.revision, documentId: snapshot.document.documentId, cancellation: new CancellationToken() });
    this.#active = task; return task;
  }
  accept(task: DerivedTask, current: DocumentSnapshot, value: T): boolean {
    if (task !== this.#active || task.cancellation.cancelled || task.documentId !== current.document.documentId || task.revision !== current.revision) return false;
    this.#result = { revision: task.revision, value }; this.#active = undefined; return true;
  }
  resultFor(revision: number): T | undefined { return this.#result?.revision === revision ? this.#result.value : undefined; }
  cancel(): void { this.#active?.cancellation.cancel(); this.#active = undefined; }
  clear(): void { this.cancel(); this.#result = undefined; }
}

export function patchCommand(snapshot: DocumentSnapshot, featureId: string, operations: readonly PatchOperation[]): DocumentCommand {
  assertId(snapshot.document.documentId);
  return { documentId: snapshot.document.documentId, expectedRevision: snapshot.revision, type: 'model.patch', payload: { featureId, operations } };
}
