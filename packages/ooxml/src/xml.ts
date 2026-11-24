// SPDX-License-Identifier: MIT
// DOM construction is delegated to xmldom (MIT); budget/policy checks are original.
// XML 1.0 (5th ed., 2008), ECMA-376 Part 2 (5th ed., 2021) §6.2.4.
import { DOMParser, XMLSerializer, type Document, type Element, type Node } from '@xmldom/xmldom';
import { OfficeError } from '@office-ex/contracts';
import { checkCancelled, containerLimits, type ContainerOptions } from './limits.js';

const encoder = new TextEncoder();
function invalid(message: string): never { throw new OfficeError('invalid', message); }
function decodeXml(source: Uint8Array | string, options: ContainerOptions): string {
  const limits = containerLimits(options);
  if (typeof source === 'string') {
    if (source.length > limits.maxXmlBytes || encoder.encode(source).length > limits.maxXmlBytes) throw new OfficeError('limit', 'XML byte budget exceeded');
    return source.charCodeAt(0) === 0xfeff ? source.slice(1) : source;
  }
  if (source.length > limits.maxXmlBytes) throw new OfficeError('limit', 'XML byte budget exceeded');
  let encoding = 'utf-8';
  if ((source[0] === 0xff && source[1] === 0xfe) || (source[0] === 0x3c && source[1] === 0)) encoding = 'utf-16le';
  if ((source[0] === 0xfe && source[1] === 0xff) || (source[0] === 0 && source[1] === 0x3c)) encoding = 'utf-16be';
  try {
    const text = new TextDecoder(encoding, { fatal: true }).decode(source);
    const declaration = /^<\?xml\s+[^?]*encoding\s*=\s*(['"])([^'"]+)\1[^?]*\?>/.exec(text);
    if (declaration && /^utf-(8|16)$/i.test(declaration[2]!) &&
        ((encoding === 'utf-8') !== (declaration[2]!.toLowerCase() === 'utf-8'))) invalid('XML encoding declaration disagrees with bytes');
    return text;
  } catch (error) { if (error instanceof OfficeError) throw error; return invalid('Invalid XML character encoding'); }
}
function validCodePoint(value: number): boolean {
  return value === 9 || value === 10 || value === 13 || (value >= 0x20 && value <= 0xd7ff) ||
    (value >= 0xe000 && value <= 0xfffd) || (value >= 0x10000 && value <= 0x10ffff);
}
function validateEntities(text: string): void {
  for (let pos = text.indexOf('&'); pos !== -1; pos = text.indexOf('&', pos + 1)) {
    const end = text.indexOf(';', pos + 1);
    if (end < 0 || end - pos > 32) invalid('Malformed XML entity reference');
    const entity = text.slice(pos + 1, end);
    if (/^(amp|lt|gt|quot|apos)$/.test(entity)) continue;
    let value: number;
    if (/^#x[0-9a-fA-F]+$/.test(entity)) value = Number.parseInt(entity.slice(2), 16);
    else if (/^#[0-9]+$/.test(entity)) value = Number.parseInt(entity.slice(1), 10);
    else throw new OfficeError('unsafe', 'Custom XML entities are forbidden');
    if (!validCodePoint(value)) invalid('Invalid XML character reference');
  }
}
function attributeValue(text: string): string {
  validateEntities(text);
  return text.replace(/\r\n?/g, '\n').replace(/[\t\n]/g, ' ').replace(/&([^;]+);/g, (_all, entity: string) => {
    const predefined: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
    if (predefined[entity] !== undefined) return predefined[entity]!;
    return String.fromCodePoint(entity.startsWith('#x') ? Number.parseInt(entity.slice(2), 16) : Number.parseInt(entity.slice(1), 10));
  });
}
function lexicalAttributes(text: string, count: () => void): [string, string][] {
  const result: [string, string][] = [];
  let pos = 0;
  while (pos < text.length) {
    const start = pos;
    while (/[ \t\r\n]/.test(text[pos] ?? '') && pos < text.length) pos++;
    if (pos === text.length) break;
    if (pos === start) invalid('XML attributes require whitespace');
    const nameStart = pos;
    while (pos < text.length && !/[ \t\r\n=]/.test(text[pos]!)) pos++;
    const name = text.slice(nameStart, pos);
    if (!name || /[<>"'\/]/.test(name)) invalid('Invalid XML attribute name');
    while (pos < text.length && /[ \t\r\n]/.test(text[pos]!)) pos++;
    if (text[pos++] !== '=') invalid('XML attribute requires an equals sign');
    while (pos < text.length && /[ \t\r\n]/.test(text[pos]!)) pos++;
    const quote = text[pos++];
    if (quote !== '"' && quote !== "'") invalid('XML attribute requires quotes');
    const end = text.indexOf(quote, pos);
    if (end < 0) invalid('Unterminated XML attribute');
    const value = text.slice(pos, end);
    if (value.includes('<')) invalid('XML attribute contains a markup delimiter');
    count(); result.push([name, attributeValue(value)]); pos = end + 1;
  }
  return result;
}
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
function attributeNamespaces(attributes: [string, string][], parent: ReadonlyMap<string, string> | undefined): Map<string, string> {
  const scope = new Map(parent ?? [['xml', XML_NAMESPACE]]), raw = new Set<string>(), expanded = new Set<string>();
  for (const [name, value] of attributes) {
    if (raw.has(name)) invalid('Duplicate XML attribute');
    raw.add(name);
    if (name === 'xmlns' || name.startsWith('xmlns:')) {
      const prefix = name === 'xmlns' ? '' : name.slice(6);
      if (prefix === 'xmlns' || prefix.includes(':') || (prefix && !value) ||
          (prefix === 'xml' && value !== XML_NAMESPACE) || (prefix !== 'xml' && value === XML_NAMESPACE) || value === XMLNS_NAMESPACE) invalid('Invalid XML namespace binding');
      scope.set(prefix, value);
    }
  }
  for (const [name] of attributes) {
    const segments = name.split(':');
    if (segments.length > 2 || segments.some(segment => !segment)) invalid('Invalid XML qualified attribute name');
    const prefix = segments.length === 2 ? segments[0]! : '', local = segments.at(-1)!;
    const namespace = prefix === 'xmlns' || name === 'xmlns' ? XMLNS_NAMESPACE : prefix ? scope.get(prefix) : '';
    if (namespace === undefined) invalid('Unbound XML attribute prefix');
    const key = `${namespace}\0${local}`;
    if (expanded.has(key)) invalid('Duplicate XML expanded attribute name');
    expanded.add(key);
  }
  return scope;
}
/** A bounded lexical pass precedes DOM allocation, including comments and attributes. */
function preflightXml(text: string, options: ContainerOptions): void {
  const limits = containerLimits(options), stack: string[] = [], namespaces: Map<string, string>[] = [];
  let nodes = 0, roots = 0;
  const count = (): void => { if (++nodes > limits.maxXmlNodes) throw new OfficeError('limit', 'XML node budget exceeded'); };
  const declaration = /^<\?xml\s+[^?]*encoding\s*=\s*(['"])([^'"]+)\1[^?]*\?>/.exec(text);
  if (declaration && !/^utf-(8|16)$/i.test(declaration[2]!)) throw new OfficeError('unsupported', 'OPC XML must use UTF-8 or UTF-16');
  let characterOffset = 0;
  for (const char of text) {
    if ((characterOffset++ & 0xfff) === 0) checkCancelled(options);
    if (!validCodePoint(char.codePointAt(0)!)) invalid('Invalid XML character');
  }
  let pos = 0;
  while (pos < text.length) {
    checkCancelled(options);
    if (text[pos] !== '<') {
      const next = text.indexOf('<', pos), end = next < 0 ? text.length : next;
      const content = text.slice(pos, end);
      if (content.includes(']]>')) invalid('CDATA terminator outside CDATA');
      if (!stack.length && content.trim()) invalid('Text outside XML root');
      validateEntities(content); count(); pos = end; continue;
    }
    if (text.startsWith('<!--', pos)) {
      const end = text.indexOf('-->', pos + 4);
      if (end < 0 || text.slice(pos + 4, end).includes('--')) invalid('Malformed XML comment');
      count(); pos = end + 3; continue;
    }
    if (text.startsWith('<![CDATA[', pos)) {
      const end = text.indexOf(']]>', pos + 9);
      if (end < 0 || !stack.length) invalid('Malformed XML CDATA');
      count(); pos = end + 3; continue;
    }
    if (text.startsWith('<!', pos)) throw new OfficeError('unsafe', 'XML declarations and DTDs are forbidden');
    if (text.startsWith('<?', pos)) {
      const end = text.indexOf('?>', pos + 2);
      if (end < 0) invalid('Truncated XML processing instruction');
      if (/^xml(?:\s|$)/i.test(text.slice(pos + 2, end)) && pos !== 0) invalid('Misplaced XML declaration');
      count(); pos = end + 2; continue;
    }
    let end = pos + 1, quote = '';
    for (; end < text.length; end++) {
      const char = text[end]!;
      if (quote) { if (char === quote) quote = ''; }
      else if (char === '"' || char === "'") quote = char;
      else if (char === '>') break;
      else if (char === '<') invalid('Nested XML tag delimiter');
    }
    if (end === text.length || quote) invalid('Unterminated XML tag');
    const body = text.slice(pos + 1, end);
    if (body.startsWith('/')) {
      const match = /^\/([^\s/>]+)\s*$/.exec(body);
      if (!match || stack.pop() !== match[1]) invalid('Mismatched XML end tag');
      namespaces.pop();
    } else {
      const match = /^([^\s/>]+)/.exec(body);
      if (!match) invalid('Malformed XML start tag');
      const attributes = body.slice(match[0].length), selfClosing = attributes.endsWith('/');
      if (/\/\s+$/.test(attributes)) invalid('Whitespace after XML self-closing slash');
      const parsed = lexicalAttributes(selfClosing ? attributes.slice(0, -1) : attributes, count);
      const scope = attributeNamespaces(parsed, namespaces.at(-1)), qualified = match[1]!.split(':');
      if (qualified.length > 2 || qualified.some(segment => !segment) || qualified[0] === 'xmlns' || (qualified.length === 2 && !scope.has(qualified[0]!))) invalid('Invalid XML qualified element name');
      count();
      if (!stack.length && ++roots > 1) invalid('Multiple XML roots');
      if (stack.length + 1 > limits.maxXmlDepth) throw new OfficeError('limit', 'XML depth budget exceeded');
      if (!selfClosing) { stack.push(match[1]!); namespaces.push(scope); }
    }
    pos = end + 1;
  }
  if (stack.length || roots !== 1) invalid('Incomplete XML document');
}

/** Never returns a recovered malformed DOM; all xmldom warnings/errors are fatal. */
export function parseXml(source: Uint8Array | string, options: ContainerOptions = {}): Document {
  checkCancelled(options);
  const text = decodeXml(source, options);
  preflightXml(text, options);
  let document: Document;
  try {
    document = new DOMParser({ onError: () => { throw new OfficeError('invalid', 'Malformed XML'); } }).parseFromString(text, 'application/xml');
  } catch (error) { if (error instanceof OfficeError) throw error; return invalid('Malformed XML'); }
  const stack: Node[] = [document], limits = containerLimits(options);
  let count = 0;
  while (stack.length) {
    checkCancelled(options);
    const node = stack.pop()!;
    if (++count > limits.maxXmlNodes) throw new OfficeError('limit', 'XML node budget exceeded');
    if (node.nodeType === 1) {
      const element = node as Element;
      if (element.namespaceURI === 'http://www.w3.org/2001/XInclude') throw new OfficeError('unsafe', 'XInclude is forbidden');
      if (element.prefix && !element.namespaceURI) invalid('Unbound XML namespace prefix');
      const attributeNames = new Set<string>();
      for (let i = 0; i < element.attributes.length; i++) {
        const attribute = element.attributes.item(i)!;
        if (attribute.prefix && !attribute.namespaceURI) invalid('Unbound XML attribute namespace prefix');
        const name = `${attribute.namespaceURI ?? ''}\0${attribute.localName}`;
        if (attributeNames.has(name)) invalid('Duplicate XML expanded attribute name');
        attributeNames.add(name);
      }
    }
    for (let child = node.lastChild; child; child = child.previousSibling) stack.push(child);
  }
  return document;
}

/** Revalidates a constructed/mutated DOM before handing bytes to a package writer. */
export function serializeXml(document: Document, options: ContainerOptions = {}): Uint8Array {
  checkCancelled(options);
  const text = new XMLSerializer().serializeToString(document).replace(
    /^(<\?xml\s+[^?]*encoding\s*=\s*['"])[^'"]+(['"][^?]*\?>)/, '$1UTF-8$2');
  parseXml(text, options);
  return encoder.encode(text);
}
export function childElements(element: Element): Element[] {
  const result: Element[] = [];
  for (let child = element.firstChild; child; child = child.nextSibling) if (child.nodeType === 1) result.push(child as Element);
  return result;
}
