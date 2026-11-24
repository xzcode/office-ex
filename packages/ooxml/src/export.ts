// SPDX-License-Identifier: MIT
// Original minimal OOXML writers. No external office engine.
export type Parts = Record<string, string>;
export type Run = { text: string; bold?: boolean };
export type Paragraph = { runs: Run[] };
export type Cell = { row: number; col: number; value: string | number | { formula: string } };
export type Sheet = { name: string; cells: Cell[] };

const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const main = 'http://schemas.openxmlformats.org/officeDocument/2006/';
const pkg = 'http://schemas.openxmlformats.org/package/2006/';
const rel = main + 'relationships/';

function escapeXml(s: string): string {
  for (const char of s) {
    const c = char.codePointAt(0)!;
    if (!(c === 9 || c === 10 || c === 13 || (c >= 32 && c <= 0xD7FF) ||
          (c >= 0xE000 && c <= 0xFFFD) || (c >= 0x10000 && c <= 0x10FFFF))) {
      throw new Error('Invalid XML character');
    }
  }
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;').replace(/\r/g, '&#13;');
}

function relationships(items: { id: string; type: string; target: string }[]): string {
  return xml + `<Relationships xmlns="${pkg}relationships">` + items.map(i =>
    `<Relationship Id="${i.id}" Type="${rel}${i.type}" Target="${i.target}"/>`).join('') + '</Relationships>';
}

function contentTypes(items: { path: string; type: string }[]): string {
  return xml + `<Types xmlns="${pkg}content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' + items.map(i =>
    `<Override PartName="/${i.path}" ContentType="${i.type}"/>`).join('') + '</Types>';
}

export function docxParts(paragraphs: Paragraph[]): Parts {
  const path = 'word/document.xml';
  const body = paragraphs.map(p => '<w:p>' + p.runs.map(r => {
    const text = r.text.split(/(\r\n|\r|\n|\t)/).map(t => t === '\t' ? '<w:tab/>' :
      /^\r\n$|^[\r\n]$/.test(t) ? '<w:br/>' : `<w:t xml:space="preserve">${escapeXml(t)}</w:t>`).join('');
    return '<w:r>' + (r.bold ? '<w:rPr><w:b/></w:rPr>' : '') + text + '</w:r>';
  }).join('') + '</w:p>').join('');
  return {
    '[Content_Types].xml': contentTypes([{path, type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'}]),
    '_rels/.rels': relationships([{id: 'rId1', type: 'officeDocument', target: path}]),
    [path]: xml + `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`
  };
}

function column(n: number): string {
  let result = '';
  while (n > 0) { n--; result = String.fromCharCode(65 + n % 26) + result; n = Math.floor(n / 26); }
  return result;
}

export function xlsxParts(sheets: Sheet[]): Parts {
  if (!sheets.length) throw new Error('Workbook needs at least one sheet');
  const names = new Set<string>();
  const parts: Parts = {};
  const types = [{path: 'xl/workbook.xml', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'}];
  const refs: {id: string; type: string; target: string}[] = [];
  const sheetTags = sheets.map((s, i) => {
    const key = s.name.toLowerCase();
    if (!s.name.trim() || s.name.length > 31 || /[\\/:?*\[\]]/.test(s.name) || /^'|'$/.test(s.name) || names.has(key)) throw new Error('Invalid or duplicate sheet name');
    names.add(key);
    const id = `rId${i+1}`, target = `worksheets/sheet${i+1}.xml`, path = `xl/${target}`;
    refs.push({id, type: 'worksheet', target});
    types.push({path, type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'});
    const rows = new Map<number, string[]>();
    const seen = new Set<string>();
    for (const c of [...s.cells].sort((a,b) => a.row-b.row || a.col-b.col)) {
      if (!Number.isInteger(c.row) || c.row < 1 || c.row > 1048576 || !Number.isInteger(c.col) || c.col < 1 || c.col > 16384) throw new Error('Cell outside spreadsheet bounds');
      const addr = column(c.col) + c.row;
      if (seen.has(addr)) throw new Error('Duplicate cell');
      seen.add(addr);
      let tag: string;
      if (typeof c.value === 'string') {
        if (c.value.length > 32767) throw new Error('Cell text too long');
        tag = `<c r="${addr}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(c.value)}</t></is></c>`;
      } else if (typeof c.value === 'number') {
        if (!Number.isFinite(c.value)) throw new Error('Non-finite cell value');
        tag = `<c r="${addr}"><v>${c.value}</v></c>`;
      } else {
        if (!c.value.formula.startsWith('=') || c.value.formula.length < 2) throw new Error('Formula must begin with =');
        tag = `<c r="${addr}"><f>${escapeXml(c.value.formula.slice(1))}</f></c>`;
      }
      if (!rows.has(c.row)) rows.set(c.row, []);
      rows.get(c.row)!.push(tag);
    }
    parts[path] = xml + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` +
      [...rows].map(([r,c]) => `<row r="${r}">${c.join('')}</row>`).join('') + '</sheetData></worksheet>';
    return `<sheet name="${escapeXml(s.name)}" sheetId="${i+1}" r:id="${id}"/>`;
  });
  parts['[Content_Types].xml'] = contentTypes(types);
  parts['_rels/.rels'] = relationships([{id: 'rId1', type: 'officeDocument', target: 'xl/workbook.xml'}]);
  parts['xl/_rels/workbook.xml.rels'] = relationships(refs);
  parts['xl/workbook.xml'] = xml + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${main}relationships"><sheets>${sheetTags.join('')}</sheets><calcPr fullCalcOnLoad="1"/></workbook>`;
  return parts;
}
