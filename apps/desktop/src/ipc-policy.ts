// SPDX-License-Identifier: MIT
/** Pure request/origin validation. Does not parse document contents or grant filesystem access. */
import { isAbsolute, resolve } from 'node:path';
import { OfficeError, isRecord } from '@office-ex/contracts';
import { isLanguagePreference, isRegionalPreference, isDocumentLanguagePreference } from '@office-ex/i18n';
import type { TabState, PublicSettings } from '@office-ex/ui';

export const TRUSTED_APP_ORIGIN = 'office-ex://app';
export const IPC_MAX_DRAFT_BYTES = 32 * 1024 * 1024;
export const IPC_MAX_TABS = 32;
const ID = /^[A-Za-z0-9_-]{1,128}$/;
const CONTENT_ID = /^[A-Za-z0-9_.:-]{1,256}$/;
const RECOVERY_ID = /^[A-Za-z0-9_-]{1,128}\.[0-9]+\.[a-f0-9-]{36}\.oex-recovery$/;
const ARITY: Readonly<Record<string, number>> = Object.freeze({
  'office:initialize': 0, 'office:open': 0, 'office:recent': 1, 'office:drop': 1,
  'office:tabs': 1, 'office:save-draft': 1, 'office:save-docx': 1, 'office:save-doc': 1, 'office:save-xlsx': 1, 'office:save-xls': 1, 'office:save-pptx': 1, 'office:save-ppt': 1, 'office:save-copy': 1,
  'office:confirm-close': 1, 'office:release': 2, 'office:finish-close': 0,
  'office:settings': 1, 'office:recover': 1, 'office:discard-recovery': 1,
  'office:checkpoint': 1, 'office:export-pdf': 1, 'office:print': 1,
});
function invalid(message = 'Invalid IPC request'): never { throw new OfficeError('invalid', message); }

/** URL.origin is 'null' for this custom scheme in Node; inspect its components. */
export function isLocalAppUrl(value: string | URL): boolean {
  try {
    const url = typeof value === 'string' ? new URL(value) : value;
    return url.protocol === 'office-ex:' && url.hostname === 'app' && !url.port && !url.username && !url.password;
  } catch { return false; }
}
export interface SenderEvidence {
  windowAlive: boolean;
  expectedWebContents: boolean;
  mainFrame: boolean;
  /** Electron WebFrameMain.origin, not Node's URL.origin. */
  frameOrigin: string | null | undefined;
}
export function assertTrustedSender(evidence: SenderEvidence): void {
  if (evidence.windowAlive !== true || evidence.expectedWebContents !== true || evidence.mainFrame !== true || evidence.frameOrigin !== TRUSTED_APP_ORIGIN) {
    throw new OfficeError('unsafe', 'Untrusted IPC sender');
  }
}
export function assertIpcArity(channel: string, args: readonly unknown[]): void {
  if (!Object.hasOwn(ARITY, channel) || args.length !== ARITY[channel]) invalid('Invalid IPC argument count');
}
export function ipcId(value: unknown): asserts value is string {
  if (typeof value !== 'string' || !ID.test(value)) invalid('Invalid session identifier');
}
function fields(value: unknown, allowed: readonly string[]): Record<string, unknown> {
  if (!isRecord(value) || Object.getOwnPropertySymbols(value).length || Object.getOwnPropertyNames(value).some(key => !allowed.includes(key))) invalid();
  for (const descriptor of Object.values(Object.getOwnPropertyDescriptors(value))) {
    if (!('value' in descriptor) || !descriptor.enumerable) invalid('IPC accessors and hidden fields are forbidden');
  }
  return value;
}
function string(value: unknown, max: number, label: string): asserts value is string {
  if (typeof value !== 'string' || !value.length || value.length > max || /[\u0000]/.test(value)) invalid(`Invalid ${label}`);
}
function revision(value: unknown): asserts value is number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) invalid('Invalid document revision');
}
function bytes(value: unknown): asserts value is Uint8Array {
  if (!(value instanceof Uint8Array) || !value.byteLength || value.byteLength > IPC_MAX_DRAFT_BYTES) throw new OfficeError('limit', 'Invalid IPC payload size');
  if (typeof SharedArrayBuffer !== 'undefined' && value.buffer instanceof SharedArrayBuffer) throw new OfficeError('unsafe', 'Shared IPC byte buffers are forbidden');
}
function contentId(value: unknown): asserts value is string {
  if (typeof value !== 'string' || !CONTENT_ID.test(value)) invalid('Invalid persistent content identifier');
}
export function parseDroppedPath(value: unknown): string {
  string(value, 8192, 'dropped path');
  if (!isAbsolute(value)) invalid('Dropped file path must be absolute');
  return value;
}
export function parseTabStates(value: unknown): TabState[] {
  if (!Array.isArray(value) || value.length > IPC_MAX_TABS) throw new OfficeError('limit', 'Too many document tabs');
  const ids = new Set<string>();
  return value.map(candidate => {
    const item = fields(candidate, ['id', 'title', 'dirty', 'revision']);
    ipcId(item.id); string(item.title, 1024, 'tab title'); revision(item.revision);
    if (typeof item.dirty !== 'boolean' || ids.has(item.id)) invalid('Invalid or duplicate tab metadata');
    ids.add(item.id);
    return { id: item.id, title: item.title, dirty: item.dirty, revision: item.revision };
  });
}
export interface DraftSaveRequest { documentId: string; sourceToken?: string; bytes: Uint8Array; revision: number; contentId: string; saveAs: boolean; title: string }
export function parseDraftSave(value: unknown): DraftSaveRequest {
  const item = fields(value, ['documentId', 'sourceToken', 'bytes', 'revision', 'contentId', 'saveAs', 'title']);
  ipcId(item.documentId); if (item.sourceToken !== undefined) ipcId(item.sourceToken);
  bytes(item.bytes); revision(item.revision); contentId(item.contentId); string(item.title, 1024, 'document title');
  if (typeof item.saveAs !== 'boolean') invalid('Invalid Save As flag');
  return { documentId: item.documentId, ...(item.sourceToken === undefined ? {} : { sourceToken: item.sourceToken }), bytes: item.bytes, revision: item.revision, contentId: item.contentId, saveAs: item.saveAs, title: item.title };
}
/** DOCX IPC accepts only a bounded serialized model, never a caller-chosen ZIP.
 * Semantic verification occurs in the independent verifier renderer.
 */
export function parseDocxSave(value:unknown):DraftSaveRequest {
  const request=parseDraftSave(value);
  if(!request.contentId.startsWith(request.documentId+':content:')||!/^\d+$/.test(request.contentId.slice(request.documentId.length+9)))invalid('DOCX snapshot identity does not match the document');
  return request;
}
export type PrintRequest = {kind:'document';documentId:string;revision:number;contentId:string;bytes:Uint8Array;sourceToken?:string} | {kind:'pdf';documentId:string;revision:number;sourceToken:string};
/** PDF input is fetched from the main-owned source token. Document print input
 * is only a bounded draft, independently parsed in the isolated output worker. */
export function parsePrintRequest(value:unknown):PrintRequest {
  const item=fields(value,['kind','documentId','revision','contentId','bytes','sourceToken']);
  ipcId(item.documentId);revision(item.revision);
  if(item.sourceToken!==undefined)ipcId(item.sourceToken);
  if(item.kind==='pdf'){
    if(item.sourceToken===undefined||item.contentId!==undefined||item.bytes!==undefined)invalid('PDF print requires an original file capability');
    return {kind:'pdf',documentId:item.documentId,revision:item.revision,sourceToken:item.sourceToken};
  }
  if(item.kind!=='document')invalid('Invalid print kind');
  bytes(item.bytes);contentId(item.contentId);
  if(!item.contentId.startsWith(item.documentId+':content:')||!/^\d+$/.test(item.contentId.slice(item.documentId.length+9)))invalid('Print snapshot identity differs from document');
  return {kind:'document',documentId:item.documentId,revision:item.revision,contentId:item.contentId,bytes:item.bytes,...(item.sourceToken===undefined?{}:{sourceToken:item.sourceToken})};
}
export function parsePreservedCopy(value: unknown): { documentId: string; sourceToken: string } {
  const item = fields(value, ['documentId', 'sourceToken']); ipcId(item.documentId); ipcId(item.sourceToken);
  return { documentId: item.documentId, sourceToken: item.sourceToken };
}
export interface CheckpointRequest { documentId: string; contentId: string; revision: number; bytes: Uint8Array }
export function parseCheckpoint(value: unknown): CheckpointRequest {
  const item = fields(value, ['documentId', 'contentId', 'revision', 'bytes']);
  ipcId(item.documentId); contentId(item.contentId); revision(item.revision); bytes(item.bytes);
  return { documentId: item.documentId, contentId: item.contentId, revision: item.revision, bytes: item.bytes };
}
export function parseRelease(documentId: unknown, discard: unknown): { documentId: string; discard: boolean } {
  ipcId(documentId); if (typeof discard !== 'boolean') invalid('Invalid close request');
  return { documentId, discard };
}
export function parseSettingsPatch(value: unknown): { theme?: PublicSettings['theme']; language?: PublicSettings['language']; region?:PublicSettings['region']; documentLanguage?:PublicSettings['documentLanguage']; clearRecent?: boolean } {
  const item = fields(value, ['theme', 'language', 'region', 'documentLanguage', 'clearRecent']);
  if (item.theme !== undefined && item.theme !== 'system' && item.theme !== 'light' && item.theme !== 'dark') invalid('Invalid theme');
  if(item.language!==undefined&&!isLanguagePreference(item.language))invalid('Invalid language');
  if(item.region!==undefined&&!isRegionalPreference(item.region))invalid('Invalid region');
  if(item.documentLanguage!==undefined&&!isDocumentLanguagePreference(item.documentLanguage))invalid('Invalid document language preference');
  if (item.clearRecent !== undefined && typeof item.clearRecent !== 'boolean') invalid('Invalid recent-history option');
  return {
    ...(item.theme === undefined ? {} : { theme: item.theme }),
    ...(item.language === undefined ? {} : { language: item.language }),
    ...(item.region===undefined?{}:{region:item.region}),
    ...(item.documentLanguage===undefined?{}:{documentLanguage:item.documentLanguage}),
    ...(item.clearRecent === undefined ? {} : { clearRecent: item.clearRecent }),
  };
}
export function parseRecoveryId(value: unknown): string {
  if (typeof value !== 'string' || value.length > 256 || !RECOVERY_ID.test(value)) invalid('Invalid recovery identifier');
  return value;
}

/** Only explicitly supplied document path arguments become OS-open requests. */
export function commandLineFiles(argv: readonly string[], defaultApp: boolean, cwd: string): string[] {
  const result: string[] = [];
  const allowed = new Set(['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'oex', 'pdf']);
  for (const candidate of argv.slice(defaultApp ? 2 : 1)) {
    if (!candidate || candidate.length > 8192 || candidate.includes('\0') || candidate.startsWith('-')) continue;
    if (!isAbsolute(candidate) && /^[A-Za-z][A-Za-z0-9+.-]*:/.test(candidate)) continue;
    const extension = candidate.slice(candidate.lastIndexOf('.') + 1).toLowerCase();
    if (!allowed.has(extension)) continue;
    const path = isAbsolute(candidate) ? candidate : resolve(cwd, candidate);
    if (path.length <= 8192 && !result.includes(path)) result.push(path);
    if (result.length === IPC_MAX_TABS) break;
  }
  return result;
}
