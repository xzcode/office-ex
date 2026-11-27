import type { Result } from '@office-ex/contracts';
import type { LanguagePreference, RegionalPreference, DocumentLanguagePreference } from '@office-ex/i18n';
export interface PublicSettings { version:2;theme:'system'|'light'|'dark';language:LanguagePreference;region:RegionalPreference;documentLanguage:DocumentLanguagePreference;openInTabs:boolean;recent:{id:string;name:string;lastUsedAt?:number}[] }
export interface RecoveryItem { version:1;id:string;documentId:string;contentId:string;revision:number;createdAt:number;sha256:string;byteLength:number }
export interface OpenFileResult {token:string;documentId:string;name:string;bytes:Uint8Array;duplicate:boolean}
export interface TabState {id:string;title:string;dirty:boolean;revision:number}
export type PrintRequest = {kind:'document';documentId:string;revision:number;contentId:string;bytes:Uint8Array;sourceToken?:string} | {kind:'pdf';documentId:string;revision:number;sourceToken:string};
export interface NewDocumentRequest {kind:'writer'|'sheets'|'slides';format:'oex'|'docx'|'doc'|'xlsx'|'xls'|'pptx'|'ppt'}
export type WindowInitial={kind:'file';file:OpenFileResult}|{kind:'new';request:NewDocumentRequest}|{kind:'recovery';bytes:Uint8Array};
export type MenuAction = 'new-writer'|'new-sheets'|'new-slides'|'open'|'save'|'save-as'|'close-tab'|'undo'|'redo'|'settings'|'close-window'|'export-pdf'|'print';
export interface DesktopBridge {
  exportPdf(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{name:string;durability:string;pages:number;warnings:string[]}|null>>;
  print(request:PrintRequest):Promise<Result<{pages:number;warnings:string[]}>>;
  initialize(): Promise<Result<{settings:PublicSettings;recovery:{entries:RecoveryItem[];corrupt:number};platform:string;systemLocales:string[];appVersion:string;windowRole:'home'|'document';initial?:WindowInitial}>>;
  openFile():Promise<Result<OpenFileResult|null>>;
  openRecent(id:string):Promise<Result<OpenFileResult|null>>;
  openDroppedFile(file:File):Promise<Result<OpenFileResult|null>>;
  ready():Promise<Result<void>>;
  newWindow(request:NewDocumentRequest):Promise<Result<void>>;
  restoreWindow(id:string):Promise<Result<void>>;
  showHome():Promise<Result<void>>;
  onSettings(callback:(settings:PublicSettings)=>void):()=>void;
  saveDraft(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  saveDocx(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  saveDoc(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  saveXlsx(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  saveXls(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  savePptx(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  savePpt(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  saveCopy(request:{documentId:string;sourceToken:string}):Promise<Result<{name:string;durability:string;warnings?:string[]}|null>>;
  updateTabs(states:TabState[]):Promise<Result<void>>;
  confirmClose(documentId:string):Promise<Result<'save'|'discard'|'close'|'cancel'>>;
  releaseDocument(documentId:string,discardRecovery:boolean):Promise<Result<void>>;
  finishClose():Promise<Result<void>>;
  cancelWindowClose():Promise<Result<void>>;
  updateSettings(patch:{theme?:PublicSettings['theme'];language?:PublicSettings['language'];region?:PublicSettings['region'];documentLanguage?:PublicSettings['documentLanguage'];openInTabs?:boolean;clearRecent?:boolean}):Promise<Result<PublicSettings>>;
  recover(id:string):Promise<Result<{bytes:Uint8Array;metadata:RecoveryItem}>>;
  discardRecovery(id:string):Promise<Result<void>>;
  checkpoint(request:{documentId:string;contentId:string;revision:number;bytes:Uint8Array}):Promise<Result<void>>;
  onAction(callback:(action:MenuAction)=>void):()=>void;
  onFile(callback:(file:OpenFileResult)=>void):()=>void;
}
declare global { interface Window { office:DesktopBridge } }
