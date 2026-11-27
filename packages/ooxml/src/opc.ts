// SPDX-License-Identifier: MIT
// Original OPC graph implementation. ECMA-376 Part 2, 5th edition (2021),
// §§6.2 (parts), 6.5 (relationships), 7 (ZIP mapping).
// https://ecma-international.org/publications-and-standards/standards/ecma-376/
import type { Element } from '@xmldom/xmldom';
import { OfficeError } from '@office-ex/contracts';
import { sha256 } from '@office-ex/core';
import { bytesEqual,checkCancelled,containerLimits,type ContainerOptions } from './limits.js';
import { childElements, parseXml } from './xml.js';
import { readZip, writeZip, type ZipWriteOptions } from './zip.js';

const CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const STRICT_OFFICE_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument';
const MAIN_CONTENT_TYPES: Readonly<Record<string, 'docx' | 'xlsx' | 'pptx'>> = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml': 'docx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml': 'xlsx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml': 'pptx',
};
const MAIN_NAMESPACES: Readonly<Record<string, readonly [string, string, string]>> = {
  docx: ['document', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main'],
  xlsx: ['workbook', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main'],
  pptx: ['presentation', 'http://schemas.openxmlformats.org/presentationml/2006/main', 'http://purl.oclc.org/ooxml/presentationml/main'],
};
export interface OpcRelationship {
  id: string;
  type: string;
  target: string;
  mode: 'Internal' | 'External';
  /** Case-resolved package path, no leading slash; null for external targets. */
  targetPart: string | null;
  fragment: string | null;
}
export interface OpcPackage {
  /** Entries and unchanged parts retain their exact uncompressed bytes. */
  parts: Map<string, Uint8Array>;
  contentTypes: Map<string, string>;
  /** Empty source key is the package root; other keys are case-resolved part paths. */
  relationships: Map<string, OpcRelationship[]>;
  mainPart: string | null;
  kind: 'docx' | 'xlsx' | 'pptx' | null;
  conformance: 'transitional' | 'strict' | 'unknown';
  hasMacros: boolean;
  hasSignatures: boolean;
  externalRelationships: number;
  originalBytes: Uint8Array;
}
interface OriginalPackage { partHashes: Map<string, string>; originalHash: string; hasMacros: boolean; hasSignatures: boolean }
const originalParts = new WeakMap<OpcPackage, OriginalPackage>();
function invalid(message: string): never { throw new OfficeError('invalid', message); }
function asciiFold(value: string): string { return value.replace(/[A-Z]/g, c => c.toLowerCase()); }
function attr(element: Element, name: string): string {
  if (!element.hasAttribute(name)) return invalid('Required OPC attribute missing');
  const value = element.getAttribute(name)!;
  if (!value) return invalid('Empty OPC attribute');
  return value;
}

/** Logical part path without its leading slash (directories are not parts). */
export function validatePartPath(path: string): string {
  if (!path || path.startsWith('/') || /[\\\u0000-\u0020\u007f?#]/.test(path)) throw new OfficeError('unsafe', 'Unsafe OPC part name');
  const segments = path.split('/');
  for (const segment of segments) {
    if (!segment || segment.endsWith('.') || /^\.+$/.test(segment)) invalid('Non-canonical OPC part segment');
    for (let pos = segment.indexOf('%'); pos >= 0; pos = segment.indexOf('%', pos + 3)) {
      const hex = segment.slice(pos + 1, pos + 3);
      if (!/^[0-9a-fA-F]{2}$/.test(hex)) invalid('Invalid OPC percent encoding');
      const code = Number.parseInt(hex, 16), decoded = String.fromCharCode(code);
      if (code === 0x2f || code === 0x5c || code < 0x20 || code === 0x7f || /[A-Za-z0-9._~-]/.test(decoded)) throw new OfficeError('unsafe', 'Ambiguous OPC percent encoding');
    }
    if (/[<>"{}|^`\[\]]/.test(segment)) invalid('Invalid OPC part character');
  }
  return path;
}
export function resolvePartTarget(sourcePart: string, target: string): { path: string; fragment: string | null } {
  if (!target || /[\\\u0000-\u0020\u007f?]/.test(target) || target.startsWith('//') || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(target)) throw new OfficeError('unsafe', 'Unsafe internal OPC target');
  const hash = target.indexOf('#'), fragment = hash < 0 ? null : target.slice(hash + 1);
  const targetPath = hash < 0 ? target : target.slice(0, hash);
  if (!targetPath) {
    if (!sourcePart) invalid('Package relationship has no target part');
    return { path: sourcePart, fragment };
  }
  const segments = targetPath.startsWith('/') ? [] : sourcePart.split('/').slice(0, -1);
  for (const segment of targetPath.replace(/^\//, '').split('/')) {
    if (segment === '.') continue;
    if (segment === '..') {
      if (!segments.length) throw new OfficeError('unsafe', 'OPC target escapes package');
      segments.pop();
    } else {
      if (!segment) invalid('Empty OPC target segment');
      segments.push(segment);
    }
  }
  return { path: validatePartPath(segments.join('/')), fragment };
}
function relationshipSource(path: string): string | null {
  const lower = asciiFold(path);
  if (lower === '_rels/.rels') return '';
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/i.exec(path);
  return match ? (match[1] ?? '') + match[2]! : null;
}
function validateContentType(type: string): void {
  // Office/OPC content types do not require MIME parameters in the M1 profile.
  if (!/^[!#$%&'*+.^_`|~0-9a-zA-Z-]+\/[!#$%&'*+.^_`|~0-9a-zA-Z-]+$/.test(type)) invalid('Invalid OPC content type');
}
function structure(parts: Map<string, Uint8Array>, options: ContainerOptions): Omit<OpcPackage, 'originalBytes'> {
  checkCancelled(options);
  const names = new Map<string, string>();
  for (const path of parts.keys()) {
    if (path.endsWith('/')) continue;
    if (path !== '[Content_Types].xml') validatePartPath(path);
    const canonical = asciiFold(path);
    if (names.has(canonical)) invalid('Equivalent OPC part names');
    names.set(canonical, path);
  }
  for (const canonical of names.keys()) {
    const segments = canonical.split('/'); segments.pop();
    while (segments.length) {
      if (names.has(segments.join('/'))) invalid('An OPC part is a parent of another part');
      segments.pop();
    }
  }
  const typeBytes = parts.get('[Content_Types].xml');
  if (!typeBytes) invalid('OPC content types part missing');
  const root = parseXml(typeBytes, options).documentElement;
  if (!root || root.namespaceURI !== CONTENT_TYPES || root.localName !== 'Types') invalid('Invalid OPC content types root');
  const defaults = new Map<string, string>(), overrides = new Map<string, string>();
  for (const child of childElements(root)) {
    checkCancelled(options);
    if (child.namespaceURI !== CONTENT_TYPES) invalid('Unknown OPC content type element');
    const type = attr(child, 'ContentType'); validateContentType(type);
    if (child.localName === 'Default') {
      const extension = asciiFold(attr(child, 'Extension'));
      if (!/^[A-Za-z0-9_-]+$/.test(extension) || defaults.has(extension)) invalid('Duplicate or invalid OPC default extension');
      defaults.set(extension, type);
    } else if (child.localName === 'Override') {
      const name = attr(child, 'PartName');
      if (!name.startsWith('/')) invalid('OPC override name must be absolute');
      const canonical = asciiFold(validatePartPath(name.slice(1)));
      if (overrides.has(canonical) || !names.has(canonical)) invalid('Duplicate or dangling OPC override');
      overrides.set(canonical, type);
    } else invalid('Unknown OPC content type declaration');
  }
  const contentTypes = new Map<string, string>();
  for (const [canonical, path] of names) {
    if (path === '[Content_Types].xml') continue;
    const filename = path.slice(path.lastIndexOf('/') + 1), dot = filename.lastIndexOf('.');
    const type = overrides.get(canonical) ?? (dot >= 0 ? defaults.get(asciiFold(filename.slice(dot + 1))) : undefined);
    if (!type) invalid('OPC part has no content type');
    contentTypes.set(path, type);
  }
  const relationships = new Map<string, OpcRelationship[]>();
  let externalRelationships = 0, hasMacros = false, hasSignatures = false;
  for (const [path, type] of contentTypes) {
    checkCancelled(options);
    if (/vbaProject|macroEnabled|macrosheet/i.test(type) || /(^|\/)vbaproject\.bin$/i.test(path)) hasMacros = true;
    if (/digital-signature/i.test(type) || /^_xmlsignatures\//i.test(path)) hasSignatures = true;
    const source = relationshipSource(path);
    if (source !== null) {
      if (type !== REL_CONTENT_TYPE) invalid('Relationship part has the wrong content type');
      const sourcePath = source ? names.get(asciiFold(source)) : '';
      if (sourcePath === undefined || (sourcePath && relationshipSource(sourcePath) !== null)) invalid('Invalid relationship source');
      if (relationships.has(sourcePath)) invalid('Duplicate relationship source');
      const relRoot = parseXml(parts.get(path)!, options).documentElement;
      if (!relRoot || relRoot.namespaceURI !== RELATIONSHIPS || relRoot.localName !== 'Relationships') invalid('Invalid relationship root');
      const ids = new Set<string>(), rels: OpcRelationship[] = [];
      for (const child of childElements(relRoot)) {
        if (child.namespaceURI !== RELATIONSHIPS || child.localName !== 'Relationship') invalid('Unknown relationship element');
        const id = attr(child, 'Id'), relType = attr(child, 'Type'), target = attr(child, 'Target'), modeAttribute = child.getAttribute('TargetMode');
        if (!/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(id) || ids.has(id)) invalid('Invalid or duplicate relationship ID');
        if (!/^[A-Za-z][A-Za-z0-9+.-]*:[^\s]+$/.test(relType)) invalid('Relationship type is not an absolute URI');
        if (modeAttribute && modeAttribute !== 'Internal' && modeAttribute !== 'External') invalid('Invalid relationship target mode');
        ids.add(id);
        const mode = modeAttribute === 'External' ? 'External' : 'Internal';
        let targetPart: string | null = null, fragment: string | null = null;
        if (mode === 'Internal') {
          const resolved = resolvePartTarget(sourcePath, target);
          targetPart = names.get(asciiFold(resolved.path)) ?? null; fragment = resolved.fragment;
          if (!targetPart || targetPart === '[Content_Types].xml' || relationshipSource(targetPart) !== null) invalid('Missing or forbidden internal relationship target');
        } else {
          if (/[\u0000-\u001f\u007f]/.test(target)) invalid('Invalid external relationship target');
          externalRelationships++;
        }
        rels.push({ id, type: relType, target, mode, targetPart, fragment });
      }
      relationships.set(sourcePath, rels);
    } else {
      if (type === REL_CONTENT_TYPE) invalid('Misnamed relationship part');
      // Validate all XML parts, including untouched unknown extensions. Preserve original bytes.
      if (/\+xml$|\/xml$/i.test(type)) parseXml(parts.get(path)!, options);
    }
  }
  const officeRelationships = (relationships.get('') ?? []).filter(rel => rel.type === OFFICE_REL || rel.type === STRICT_OFFICE_REL);
  if (officeRelationships.length > 1) invalid('Multiple Office main parts');
  const mainRel = officeRelationships[0];
  if (mainRel?.mode === 'External') invalid('Office main part cannot be external');
  const mainPart = mainRel?.targetPart ?? null;
  const kind = mainPart ? MAIN_CONTENT_TYPES[contentTypes.get(mainPart)!] ?? null : null;
  let conformance: OpcPackage['conformance'] = 'unknown';
  if (mainRel && mainPart && kind) {
    const mainRoot = parseXml(parts.get(mainPart)!, options).documentElement!, expected = MAIN_NAMESPACES[kind]!;
    if (mainRoot.localName !== expected[0] || (mainRoot.namespaceURI !== expected[1] && mainRoot.namespaceURI !== expected[2])) invalid('Office main content type and XML root disagree');
    conformance = mainRel.type === STRICT_OFFICE_REL ? 'strict' : 'transitional';
    if (mainRoot.namespaceURI !== expected[conformance === 'strict' ? 2 : 1]) invalid('Mixed Office conformance namespaces');
  }
  return { parts, contentTypes, relationships, mainPart, kind, conformance, hasMacros, hasSignatures, externalRelationships };
}

/** This opens the container graph only. It never claims semantic document edit support. */
function readOpcFresh(bytes:Uint8Array,options:ContainerOptions):OpcPackage{
  const parts=readZip(bytes,options),result:OpcPackage={...structure(parts,options),originalBytes:Uint8Array.from(bytes)};
  const cancellation = { get cancelled() { return options.signal?.aborted ?? false; }, throwIfCancelled: () => checkCancelled(options) };
  originalParts.set(result, { partHashes: new Map([...parts].map(([path, value]) => [path, sha256(value, cancellation)])), originalHash: sha256(result.originalBytes, cancellation), hasMacros: result.hasMacros, hasSignatures: result.hasSignatures });
  return result;
}
export interface OpcInspectionSummary{
 kind:OpcPackage['kind'];conformance:OpcPackage['conformance'];hasMacros:boolean;hasSignatures:boolean;externalRelationships:number;
 entries:{name:string;size:number;sha256:string}[];sha256:string;
}
function summarizeOpc(pkg:OpcPackage):OpcInspectionSummary{
 const hashes=originalParts.get(pkg)!;
 return {kind:pkg.kind,conformance:pkg.conformance,hasMacros:pkg.hasMacros,hasSignatures:pkg.hasSignatures,externalRelationships:pkg.externalRelationships,
  entries:[...pkg.parts].map(([name,bytes])=>({name,size:bytes.byteLength,sha256:hashes.partHashes.get(name)!})),sha256:hashes.originalHash};
}
/** Inspection needs metadata, not mutable package internals. Reuse the
 * hashes that structural verification already calculated. */
export function inspectOpc(bytes:Uint8Array,options:ContainerOptions={}):OpcInspectionSummary{
 return summarizeOpc(readOpcFresh(bytes,options));
}
const inspectionScopes=new WeakMap<ContainerOptions,OpcPackage>();
/** Single-use read capability owned by one inspection operation. The
 * verified graph is private until the semantic parser claims it. No
 * mutable part maps or caller-supplied graph can seed this capability.
 * The immutable options pin all limits; a changed source is rejected.
 * Once consumed, a later fallback must validate from scratch. */
export function prepareOpcInspection(bytes:Uint8Array,options:ContainerOptions={}):{summary:OpcInspectionSummary;options:ContainerOptions;release():void}{
 const pinned:ContainerOptions=Object.freeze({limits:Object.freeze(containerLimits(options)),signal:options.signal});
 const pkg=readOpcFresh(bytes,pinned);
 inspectionScopes.set(pinned,pkg);
 return {summary:summarizeOpc(pkg),options:pinned,release:()=>{inspectionScopes.delete(pinned);}};
}
export function openOpc(bytes:Uint8Array,options:ContainerOptions={}):OpcPackage{
 const pkg=inspectionScopes.get(options);
 if(!pkg)return readOpcFresh(bytes,options);
 inspectionScopes.delete(options);checkCancelled(options);
 if(!bytesEqual(bytes,pkg.originalBytes))throw new OfficeError('conflict','Source changed during OPC inspection');
 return pkg;
}

/** Replacements are whole parts supplied by a semantic adapter after its impact analysis. */
export function writeOpc(package_: OpcPackage, replacements: ReadonlyMap<string, Uint8Array> = new Map(), options: ZipWriteOptions = {}): Uint8Array {
  checkCancelled(options);
  const original = originalParts.get(package_);
  if (!original) throw new OfficeError('invalid', 'OPC package must originate from openOpc');
  const baseline = original.partHashes;
  const cancellation = { get cancelled() { return options.signal?.aborted ?? false; }, throwIfCancelled: () => checkCancelled(options) };
  // Do not permit mutations of the exposed map to bypass explicit replacement planning.
  if (package_.parts.size !== baseline.size || [...baseline].some(([path, value]) => !package_.parts.has(path) || sha256(package_.parts.get(path)!, cancellation) !== value)) {
    throw new OfficeError('conflict', 'OPC source parts were mutated; use explicit replacements');
  }
  if (sha256(package_.originalBytes, cancellation) !== original.originalHash) throw new OfficeError('conflict', 'OPC source bytes were mutated');
  if (replacements.size === 0) return package_.originalBytes.slice();
  if (original.hasSignatures || original.hasMacros) throw new OfficeError('unsafe', 'Signed or macro-bearing packages cannot be changed in M1');
  const parts = new Map(package_.parts);
  for (const [path, bytes] of replacements) {
    checkCancelled(options);
    if (!parts.has(path)) throw new OfficeError('unsupported', 'Part creation/removal requires a semantic adapter');
    parts.set(path, bytes.slice());
  }
  structure(parts, options);
  return writeZip(parts, options);
}
