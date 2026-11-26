/** Sandbox-side adapter composition. Structural inspection stays independent of semantic profiles. */
import { OfficeError } from '@office-ex/contracts';
import { validateDocument, type OfficeDocument, type WriterDocument, type SheetsDocument, type SlidesDocument } from '@office-ex/core';
import { BoundedReader, sniffContainer, sortedDiagnostics, type Detection, type FormatId, type FormatDiagnostic, type CompatibilityReport } from '@office-ex/format-common';
import { openOpc, openWriterDocx, writeWriterDocx, saveWriterDocx, openSheetsXlsx, previewSheetsXlsx, type XlsxValuePreview, previewOfficeText, type OfficeTextPreview, writeSheetsXlsx, saveSheetsXlsx, openSlidesPptx, writeSlidesPptx, saveSlidesPptx } from '@office-ex/ooxml';
import type { CalculatedValue } from '@office-ex/sheets';
import { openLegacyXls, writeLegacyXls, saveLegacyXls } from './legacy-xls.js';
export { openLegacyXls, writeLegacyXls, saveLegacyXls } from './legacy-xls.js';
import { openLegacyPpt, writeLegacyPpt, saveLegacyPpt } from './legacy-ppt.js';
export { openLegacyPpt, writeLegacyPpt, saveLegacyPpt } from './legacy-ppt.js';
import { readCfb } from '@office-ex/cfb';
import { inspectLegacyDocText, openLegacyDoc, writeLegacyDoc, saveLegacyDoc, type LegacyDocText } from './legacy-doc.js';
export { inspectLegacyDocText, openLegacyDoc, writeLegacyDoc, saveLegacyDoc } from './legacy-doc.js';

export interface ContainerEntry { name: string; size: number; sha256: string }
export interface Inspection { semanticFailure?:{code:string;detail:string};xlsxValuePreview?:XlsxValuePreview;officeTextPreview?:OfficeTextPreview; detection: Detection; report: CompatibilityReport; entries: ContainerEntry[]; sha256: string; draft?: OfficeDocument; editableWriter?: WriterDocument; editableWriterFormat?: 'doc'|'docx'; legacyDocText?: LegacyDocText; editableSheets?: SheetsDocument; editableSheetsFormat?: 'xls'|'xlsx'; sheetsTemplateConversionBlocked?:boolean; sheetsCachedValues?: [string,CalculatedValue][]; editableSlides?: SlidesDocument; editableSlidesFormat?: 'ppt'|'pptx'; slidesTemplateConversionBlocked?: boolean }
/** Keep actionable diagnostics in the local inspection DTO, not server logs. */
function semanticFailure(inspection:Inspection,error:OfficeError):Inspection{
 return {...inspection,semanticFailure:{code:error.code,detail:error.message.slice(0,512)}};
}
function addOfficeTextPreview(result:Inspection,bytes:Uint8Array,signal?:AbortSignal):void{
 try{
  result.officeTextPreview=previewOfficeText(bytes,{signal});
  const diagnostics=sortedDiagnostics([...result.report.diagnostics.filter(d=>d.code!=='SEMANTIC_ADAPTER_PENDING'),diagnostic('OFFICE_TEXT_PREVIEW','FMT-04','format.officeTextPreview')]);
  result.report={...result.report,diagnostics};result.detection={...result.detection,diagnostics};
 }catch(error){if(!(error instanceof OfficeError)||!['unsupported','invalid','limit','unsafe'].includes(error.code))throw error;}
}
function addXlsxPreview(result:Inspection,bytes:Uint8Array,signal?:AbortSignal):void{
 try{
  result.xlsxValuePreview=previewSheetsXlsx(bytes,{signal});
  const diagnostics=sortedDiagnostics([...result.report.diagnostics.filter(d=>d.code!=='SEMANTIC_ADAPTER_PENDING'),diagnostic('XLSX_VALUE_PREVIEW','SHT-05','format.xlsxValuePreview')]);
  result.report={...result.report,diagnostics};result.detection={...result.detection,diagnostics};
 }catch(error){if(!(error instanceof OfficeError)||!['unsupported','invalid','limit','unsafe'].includes(error.code))throw error;}
}

const MAX_INPUT = 100 * 1024 * 1024;
const MAX_DRAFT = 32 * 1024 * 1024;
const decoder = new TextDecoder('utf-8', { fatal: true });
const diagnostic = (code: string, featureId: string, messageKey: string, severity: FormatDiagnostic['severity'] = 'warning'): FormatDiagnostic => ({code, featureId, messageKey, severity, safeAction: 'read-only'});
export async function sha256(bytes: Uint8Array): Promise<string> {
  return Array.from(new Uint8Array(await globalThis.crypto.subtle.digest('SHA-256', bytes.slice().buffer)), b => b.toString(16).padStart(2, '0')).join('');
}
export function parseDraft(bytes: Uint8Array): OfficeDocument {
  if (bytes.byteLength > MAX_DRAFT) throw new OfficeError('limit', 'Draft exceeds the 32 MiB budget');
  let envelope: unknown;
  try { envelope = JSON.parse(decoder.decode(bytes)); } catch { throw new OfficeError('invalid', 'Invalid Office EX draft'); }
  if (!envelope || typeof envelope !== 'object' || Array.isArray(envelope)) throw new OfficeError('invalid','Invalid draft envelope');
  const e = envelope as Record<string, unknown>;
  if (e.magic !== 'office-ex-draft' || e.version !== 1) throw new OfficeError('unsupported', 'Unsupported Office EX draft schema');
  if (Object.keys(e).some(k => !['magic','version','document'].includes(k))) throw new OfficeError('invalid','Unknown draft envelope field');
  const model = validateDocument(e.document);
  if(Object.values(model.resources).some(r=>r.dataBase64===undefined))throw new OfficeError('unsupported','Draft resource bytes are missing');
  return model;
}
export function serializeDraft(document: unknown): Uint8Array {
  const model = validateDocument(document);
  if(Object.values(model.resources).some(r=>r.dataBase64===undefined))throw new OfficeError('unsupported','Resource bytes must be packaged before this draft can be saved');
  const bytes = new TextEncoder().encode(JSON.stringify({magic:'office-ex-draft', version:1, document:model}));
  if (bytes.byteLength > MAX_DRAFT) throw new OfficeError('limit','Draft exceeds the 32 MiB budget');
  parseDraft(bytes); // Re-read our own serialized envelope before any filesystem commit.
  return bytes;
}
export async function inspectFile(bytes: Uint8Array, filename = '', signal?: AbortSignal): Promise<Inspection> {
  if (!(bytes instanceof Uint8Array) || bytes.byteLength > MAX_INPUT) throw new OfficeError('limit','Input exceeds the 100 MiB budget');
  if (signal?.aborted) throw new OfficeError('cancelled','Inspection cancelled');
  const container = sniffContainer(bytes);
  let format: FormatId = 'unknown'; let version = 'unknown'; let encrypted = false; let draft: OfficeDocument | undefined;
  const diagnostics: FormatDiagnostic[] = [];
  const parts: {name:string;bytes:Uint8Array}[] = [];
  if (container === 'zip') {
    const opc = openOpc(bytes, {signal});
    format = opc.kind ?? 'zip'; version = opc.conformance;
    for (const [name, data] of opc.parts) parts.push({name, bytes:data});
    if (opc.hasMacros) diagnostics.push(diagnostic('MACROS_DISABLED','SEC-02','format.macrosDisabled'));
    if (opc.hasSignatures) diagnostics.push(diagnostic('SIGNED_READ_ONLY','FMT-03','format.signedReadOnly'));
    if (opc.externalRelationships) diagnostics.push(diagnostic('EXTERNAL_LINKS_DISABLED','SEC-02','format.externalDisabled'));
    if (opc.conformance === 'strict') diagnostics.push(diagnostic('STRICT_READ_ONLY','FMT-05','format.strictReadOnly'));
  } else if (container === 'cfb') {
    const cfb = readCfb(bytes, {signal});
    const streams = new Map(cfb.entries.filter(e => e.type === 'stream').map(e => [e.path.join('/'), e.data]));
    for (const [name, data] of streams) parts.push({name,bytes:data});
    format = 'cfb'; version = `CFB v${cfb.majorVersion}`;
    // CFB stream lookup is case-insensitive; preserve original names in the report.
    const rootStreams = new Map([...streams].filter(([name]) => !name.includes('/')).map(([name, data]) => [name.toUpperCase(), data]));
    const hasStream = (name: string): boolean => rootStreams.has(name.toUpperCase());
    const getStream = (name: string): Uint8Array | undefined => rootStreams.get(name.toUpperCase());
    if (hasStream('EncryptedPackage') || hasStream('EncryptionInfo')) {
      if (!hasStream('EncryptedPackage') || !hasStream('EncryptionInfo')) throw new OfficeError('invalid','Incomplete encrypted OOXML container');
      if (getStream('EncryptedPackage')!.length < 8 || getStream('EncryptionInfo')!.length < 8) throw new OfficeError('invalid','Truncated encrypted OOXML header');
      encrypted = true; version = 'encrypted OOXML';
    } else {
      const candidates = [hasStream('WordDocument'), hasStream('Workbook') || hasStream('Book'), hasStream('PowerPoint Document') || hasStream('Current User')].filter(Boolean).length;
      if (candidates > 1) throw new OfficeError('invalid','Ambiguous Office root streams');
      if (hasStream('WordDocument')) {
        // MS-DOC 2.5.1 FibBase: wIdent,nFib,flags; full FIB mapping is M2-L.
        const r = new BoundedReader(getStream('WordDocument')!); r.range(0,32);
        if (r.u16(0) !== 0xa5ec) throw new OfficeError('invalid','Invalid Word FIB signature');
        const nFib = r.u16(2); const flags = r.u16(10);
        // Known FibBase versions only; nFibNew and full FIB semantics remain M2-L.
        // MS-DOC 2.5.2 / 2.5.6 (2022-11-15): FibBase and FibRgFcLcb.
        if (![0x00c1, 0x00d9, 0x0101, 0x010c, 0x0112].includes(nFib)) throw new OfficeError('unsupported','Unknown or pre-97 Word FIB version');
        if (!hasStream(flags & 0x0200 ? '1Table' : '0Table')) throw new OfficeError('invalid','Missing Word table stream');
        format = 'doc'; version = `FIB base 0x${nFib.toString(16)} (semantic validation pending)`;
        encrypted = Boolean(flags & 0x0100);
      } else if (hasStream('Workbook') || hasStream('Book')) {
        // MS-XLS 2.4.21 BOF. Only scan the bounded workbook-global records for FILEPASS.
        const r = new BoundedReader(getStream('Workbook') ?? getStream('Book')!); r.range(0,20);
        if (r.u16(0) !== 0x0809 || r.u16(2) !== 16) throw new OfficeError('unsupported','Unsupported Excel BOF');
        if (r.u16(4) !== 0x0600) throw new OfficeError('unsupported','Only BIFF8 is in the legacy profile');
        if (r.u16(6) !== 0x0005) throw new OfficeError('invalid','Missing workbook-global BOF');
        let offset = 0, count = 0, ended = false;
        while (offset < r.bytes.length) {
          if (signal?.aborted) throw new OfficeError('cancelled','Inspection cancelled');
          if (++count > 1_000_000) throw new OfficeError('limit','Binary record budget exceeded');
          const type = r.u16(offset), length = r.u16(offset + 2); r.range(offset + 4, length);
          if (type === 0x002f) encrypted = true;
          offset += 4 + length;
          if (type === 0x000a) {
            if (length !== 0) throw new OfficeError('invalid','Invalid workbook-global EOF length');
            ended = true; break;
          }
        }
        if (!ended) throw new OfficeError('invalid','Workbook-global EOF is missing');
        format = 'xls'; version = 'BIFF8 (semantic validation pending)';
      } else if (hasStream('PowerPoint Document') || hasStream('Current User')) {
        const main = getStream('PowerPoint Document'), user = getStream('Current User');
        if (!main || !user) throw new OfficeError('invalid','Missing PowerPoint streams');
        // MS-PPT 2.3.2 CurrentUserAtom. No username or external content is exposed.
        const r = new BoundedReader(user); r.range(0,32);
        if (r.u16(0) !== 0 || r.u16(2) !== 0x0ff6 || r.u32(8) !== 0x14 || r.u16(22) !== 0x03f4 || r.u8(24) !== 3 || r.u8(25) !== 0) throw new OfficeError('invalid','Invalid PowerPoint CurrentUserAtom');
        const userNameBytes = r.u16(20), recordBytes = r.u32(4);
        // Fixed20 + ANSI user name + release version4, plus optional UTF-16 name.
        // MS-PPT 2.3.2 (2021-02-16); no username is decoded or exposed.
        if (userNameBytes > 255 || recordBytes < 24 + userNameBytes) throw new OfficeError('invalid','Invalid PowerPoint CurrentUserAtom length');
        r.range(8, recordBytes);
        if (![8,9].includes(r.u32(28 + userNameBytes))) throw new OfficeError('invalid','Invalid PowerPoint release version');
        const token = r.u32(12);
        if (![0xe391c05f,0xf3d1c4df].includes(token)) throw new OfficeError('invalid','Invalid PowerPoint encryption marker');
        encrypted = token === 0xf3d1c4df;
        const edit = new BoundedReader(main), editOffset = r.u32(16); edit.range(editOffset,8);
        const editBytes = edit.u32(editOffset + 4);
        // MS-PPT 2.3.3 (2020-10-15); validate the atom envelope, not its full persist graph.
        if (edit.u16(editOffset) !== 0 || edit.u16(editOffset + 2) !== 0x0ff5 || ![28,32].includes(editBytes)) throw new OfficeError('invalid','Invalid PowerPoint current edit record');
        edit.range(editOffset + 8, editBytes);
        if (edit.u8(editOffset + 14) !== 0 || edit.u8(editOffset + 15) !== 3 || (encrypted && editBytes !== 32)) throw new OfficeError('invalid','Invalid PowerPoint edit version or encryption length');
        format = 'ppt'; version = 'PowerPoint 97–2003 (semantic validation pending)';
      }
      if (cfb.entries.some(e => e.path.some(name => /^(VBA|_VBA_PROJECT_CUR|Macros)$/i.test(name)))) diagnostics.push(diagnostic('MACROS_DISABLED','SEC-02','format.macrosDisabled'));
    }
  } else if (container === 'json') {
    draft = parseDraft(bytes); format = 'oex'; version = 'Office EX draft v1';
  } else if (container === 'pdf') {
    format = 'pdf'; version = 'PDF (reader planned for M4)';
  } else throw new OfficeError('unsupported','Unrecognized file signature');
  if (encrypted) diagnostics.push(diagnostic('ENCRYPTED_UNSUPPORTED','FMT-03','format.encryptedUnsupported','error'));
  if (!draft) diagnostics.push(diagnostic('SEMANTIC_ADAPTER_PENDING','FMT-04','format.semanticPending'));
  const extension = filename.includes('.') ? filename.split('.').pop()!.toLowerCase() : '';
  if (extension && extension !== format && ['doc','docx','xls','xlsx','ppt','pptx','oex','pdf'].includes(extension) && !(encrypted && container === 'cfb')) diagnostics.push(diagnostic('EXTENSION_MISMATCH','FMT-01','format.extensionMismatch'));
  const sorted = sortedDiagnostics(diagnostics);
  const entries: ContainerEntry[] = [];
  for (const part of parts) {
    if (signal?.aborted) throw new OfficeError('cancelled','Inspection cancelled');
    entries.push({name:part.name,size:part.bytes.byteLength,sha256:await sha256(part.bytes)});
  }
  entries.sort((a,b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0);
  const detection: Detection = {format,container,version,encrypted,diagnostics:sorted};
  const report: CompatibilityReport = {format,version,readOnly:!draft,diagnostics:sorted,features:[
    {featureId:container === 'cfb' ? 'LEG-01' : container === 'zip' ? 'FMT-05' : 'FMT-01',read:true,write:Boolean(draft),preserve:true,blocked:false,level:'C1'},
    {featureId:'FMT-04',read:Boolean(draft),write:Boolean(draft),preserve:true,blocked:!draft,level:draft ? 'C2':'C0'}
  ]};
  return {detection,report,entries,sha256:await sha256(bytes),...(draft ? {draft} : {})};
}

/** The desktop editor opts in only to the narrow, independently tested DOCX profile.
 * Structural inspection remains available for every file outside that profile.
 * A .docx extension is required for native editing, not for container detection.
 */
export async function inspectForEditor(bytes:Uint8Array,filename:string,documentId:string,signal?:AbortSignal):Promise<Inspection>{
  const inspection=await inspectFile(bytes,filename,signal);
  const format=inspection.detection.format;
  if((format==='ppt'||format==='pptx')&&filename.toLowerCase().endsWith('.'+format)&&bytes.byteLength<=MAX_DRAFT&&!inspection.detection.encrypted){
    try{
      const source=format==='pptx'?openSlidesPptx(bytes,documentId,filename,{signal}):openLegacyPpt(bytes,documentId,filename,{signal});
      const editableSlides=validateDocument(source.document) as SlidesDocument,featureId=format==='ppt'?'LEG-04':'SLD-03';
      const diagnostics=sortedDiagnostics([...inspection.report.diagnostics.filter(d=>d.code!=='SEMANTIC_ADAPTER_PENDING'),{code:format==='ppt'?'PPT_PROFILE_ACTIVE':'PPTX_PROFILE_ACTIVE',featureId,messageKey:format==='ppt'?'format.pptProfile':'format.pptxProfile',severity:'info',safeAction:'none'}]);
      return {...inspection,editableSlides,editableSlidesFormat:format,...('native' in source&&(source.native.layoutPaths.length>1||source.native.auxiliaryPaths.some(path=>path!==source.native.presPropsPath)||Boolean(source.document.theme?.majorFontDetails||source.document.theme?.minorFontDetails))?{slidesTemplateConversionBlocked:true}:{}),detection:{...inspection.detection,diagnostics},report:{...inspection.report,readOnly:false,diagnostics,features:[...inspection.report.features,{featureId,read:true,write:true,preserve:true,blocked:false,level:'C2'}]}};
    }catch(error){if(error instanceof OfficeError&&['unsupported','invalid','limit'].includes(error.code)){const result=semanticFailure(inspection,error);if(format==='pptx')addOfficeTextPreview(result,bytes,signal);return result;}throw error;}
  }
  if((format==='xls'||format==='xlsx')&&filename.toLowerCase().endsWith('.'+format)&&bytes.byteLength<=MAX_DRAFT&&!inspection.detection.encrypted){
    try{
      const xlsxSource=format==='xlsx'?openSheetsXlsx(bytes,documentId,filename,{signal}):undefined;
      const source=xlsxSource??openLegacyXls(bytes,documentId,filename);
      const editableSheets=validateDocument(source.document) as SheetsDocument,featureId=format==='xls'?'LEG-03':'SHT-05';
      const diagnostics=sortedDiagnostics([...inspection.report.diagnostics.filter(d=>d.code!=='SEMANTIC_ADAPTER_PENDING'),{code:format==='xls'?'XLS_PROFILE_ACTIVE':'XLSX_PROFILE_ACTIVE',featureId,messageKey:format==='xls'?'format.xlsProfile':'format.xlsxProfile',severity:'info',safeAction:'none'}]);
      return {...inspection,editableSheets,editableSheetsFormat:format,...(xlsxSource?.native.conversionBlocked?{sheetsTemplateConversionBlocked:true}:{}),sheetsCachedValues:source.cachedValues,detection:{...inspection.detection,diagnostics},report:{...inspection.report,readOnly:false,diagnostics,features:[...inspection.report.features,{featureId,read:true,write:true,preserve:true,blocked:false,level:'C2'}]}};
    }catch(error){
      if(!(error instanceof OfficeError)||!['unsupported','invalid','limit'].includes(error.code))throw error;
      const result=semanticFailure(inspection,error);
      if(format==='xlsx')addXlsxPreview(result,bytes,signal);
      return result;
    }
  }
  if(format==='xlsx'&&filename.toLowerCase().endsWith('.xlsx')&&!inspection.detection.encrypted&&bytes.byteLength>MAX_DRAFT){
    const result=semanticFailure(inspection,new OfficeError('limit','Workbook exceeds the editable model input budget'));
    addXlsxPreview(result,bytes,signal);
    return result;
  }
  if((format!=='docx'&&format!=='doc')||!filename.toLowerCase().endsWith('.'+format)||bytes.byteLength>MAX_DRAFT||inspection.detection.encrypted)return inspection;
  if(format==='doc'){
    try{inspection.legacyDocText=inspectLegacyDocText(bytes);}catch(error){if(!(error instanceof OfficeError)||!['unsupported','invalid','limit'].includes(error.code))throw error;}
  }
  try{
    const source=format==='docx'?openWriterDocx(bytes,documentId,filename,{signal}):openLegacyDoc(bytes,documentId,filename);
    const editableWriter=validateDocument(source.document) as WriterDocument;
    const featureId=format==='doc'?'LEG-02':'WRT-04';
    const diagnostics=sortedDiagnostics([
      ...inspection.report.diagnostics.filter(d=>d.code!=='SEMANTIC_ADAPTER_PENDING'),
      {code:format==='doc'?'DOC_PROFILE_ACTIVE':'DOCX_PROFILE_ACTIVE',featureId,messageKey:format==='doc'?'format.docProfile':'format.docxProfile',severity:'info',safeAction:'none'}
    ]);
    return {...inspection,editableWriter,editableWriterFormat:format,detection:{...inspection.detection,diagnostics},report:{
      ...inspection.report,readOnly:false,diagnostics,
      features:[...inspection.report.features,{featureId,read:true,write:true,preserve:true,blocked:false,level:'C2'}]
    }};
  }catch(error){
    if(error instanceof OfficeError&&['unsupported','invalid','limit'].includes(error.code)){const result=semanticFailure(inspection,error);if(format==='docx')addOfficeTextPreview(result,bytes,signal);return result;}
    throw error;
  }
}
/** Called only by the main-owned sandbox verifier when committing desktop DOCX.
 * The request contains an OEX snapshot, not caller-chosen output bytes. Existing
 * DOCX source bytes come exclusively from the host's authorized immutable cache.
 */
export function prepareWriterDocxSave(snapshotBytes:Uint8Array,documentId:string,sourceBytes?:Uint8Array):Uint8Array{
  const model=parseDraft(snapshotBytes);
  if(model.kind!=='writer'||model.documentId!==documentId)throw new OfficeError('invalid','DOCX save requires the matching Writer session');
  if(sourceBytes!==undefined&&(!(sourceBytes instanceof Uint8Array)||!sourceBytes.byteLength||sourceBytes.byteLength>MAX_DRAFT))
    throw new OfficeError('limit','DOCX source exceeds editable profile budget');
  const output=sourceBytes===undefined?writeWriterDocx(model):saveWriterDocx(openWriterDocx(sourceBytes,documentId,model.title),model);
  if(output.byteLength>MAX_DRAFT)throw new OfficeError('limit','DOCX output exceeds editable profile budget');
  return output;
}


/** Closed Word 97 semantic profile, verified independently from file commits. */
export function prepareWriterDocSave(snapshotBytes:Uint8Array,documentId:string,sourceBytes?:Uint8Array):Uint8Array{
  const model=parseDraft(snapshotBytes);
  if(model.kind!=='writer'||model.documentId!==documentId)throw new OfficeError('invalid','DOC save requires the matching Writer session');
  if(sourceBytes!==undefined&&(!(sourceBytes instanceof Uint8Array)||!sourceBytes.byteLength||sourceBytes.byteLength>MAX_DRAFT))throw new OfficeError('limit','DOC source exceeds editable profile budget');
  const output=sourceBytes===undefined?writeLegacyDoc(model):saveLegacyDoc(openLegacyDoc(sourceBytes,documentId,model.title),model);
  if(output.byteLength>MAX_DRAFT)throw new OfficeError('limit','DOC output exceeds editable profile budget');
  return output;
}

/** Main-owned sandbox accepts an immutable model snapshot and the authorized source.
 * Reopen generated files before allowing the atomic filesystem commit.
 */
export function prepareSheetsSave(format:'xls'|'xlsx',snapshotBytes:Uint8Array,documentId:string,sourceBytes?:Uint8Array):Uint8Array{
  const model=parseDraft(snapshotBytes);
  if(model.kind!=='sheets'||model.documentId!==documentId)throw new OfficeError('invalid','Spreadsheet save requires matching session identity');
  if(sourceBytes!==undefined&&(!(sourceBytes instanceof Uint8Array)||!sourceBytes.byteLength||sourceBytes.byteLength>MAX_DRAFT))throw new OfficeError('limit','Spreadsheet source exceeds editable profile budget');
  const output=format==='xlsx'?(sourceBytes===undefined?writeSheetsXlsx(model):saveSheetsXlsx(sourceBytes,model)):(sourceBytes===undefined?writeLegacyXls(model):saveLegacyXls(sourceBytes,model));
  if(output.byteLength>MAX_DRAFT)throw new OfficeError('limit','Spreadsheet output exceeds editable profile budget');
  if(format==='xlsx')openSheetsXlsx(output,documentId,model.title);else openLegacyXls(output,documentId,model.title);
  return output;
}

/** Main-owned sandbox: matching immutable model + authorized original bytes. */
export function prepareSlidesSave(format:'ppt'|'pptx',snapshotBytes:Uint8Array,documentId:string,sourceBytes?:Uint8Array):Uint8Array{
  const model=parseDraft(snapshotBytes);
  if(model.kind!=='slides'||model.documentId!==documentId)throw new OfficeError('invalid','Presentation save requires matching session identity');
  if(sourceBytes!==undefined&&(!(sourceBytes instanceof Uint8Array)||!sourceBytes.byteLength||sourceBytes.byteLength>MAX_DRAFT))throw new OfficeError('limit','Presentation source exceeds editable profile budget');
  const output=format==='pptx'?(sourceBytes===undefined?writeSlidesPptx(model):saveSlidesPptx(sourceBytes,model)):(sourceBytes===undefined?writeLegacyPpt(model):saveLegacyPpt(sourceBytes,model));
  if(output.byteLength>MAX_DRAFT)throw new OfficeError('limit','Presentation output exceeds editable profile budget');
  if(format==='pptx')openSlidesPptx(output,documentId,model.title);else openLegacyPpt(output,documentId,model.title);
  return output;
}
