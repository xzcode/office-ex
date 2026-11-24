import type { Result } from '@office-ex/contracts';
import type { LanguagePreference, RegionalPreference, DocumentLanguagePreference } from '@office-ex/i18n';
export interface PublicSettings { version:2;theme:'system'|'light'|'dark';language:LanguagePreference;region:RegionalPreference;documentLanguage:DocumentLanguagePreference;recent:{id:string;name:string}[] }
export interface RecoveryItem { version:1;id:string;documentId:string;contentId:string;revision:number;createdAt:number;sha256:string;byteLength:number }
export interface OpenFileResult {token:string;documentId:string;name:string;bytes:Uint8Array;duplicate:boolean}
export interface TabState {id:string;title:string;dirty:boolean;revision:number}
export type MenuAction = 'new-writer'|'new-sheets'|'new-slides'|'open'|'save'|'save-as'|'close-tab'|'undo'|'redo'|'settings'|'close-window';
export interface DesktopBridge {
  initialize(): Promise<Result<{settings:PublicSettings;recovery:{entries:RecoveryItem[];corrupt:number};platform:string;systemLocales:string[]}>>;
  openFile():Promise<Result<OpenFileResult|null>>;
  openRecent(id:string):Promise<Result<OpenFileResult>>;
  openDroppedFile(file:File):Promise<Result<OpenFileResult>>;
  saveDraft(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  saveDocx(request:{documentId:string;sourceToken?:string;bytes:Uint8Array;revision:number;contentId:string;saveAs:boolean;title:string}):Promise<Result<{token:string;name:string;durability:string;warnings?:string[]}|null>>;
  saveCopy(request:{documentId:string;sourceToken:string}):Promise<Result<{name:string;durability:string;warnings?:string[]}|null>>;
  updateTabs(states:TabState[]):Promise<Result<void>>;
  confirmClose(documentId:string):Promise<Result<'save'|'discard'|'close'|'cancel'>>;
  releaseDocument(documentId:string,discardRecovery:boolean):Promise<Result<void>>;
  finishClose():Promise<Result<void>>;
  updateSettings(patch:{theme?:PublicSettings['theme'];language?:PublicSettings['language'];region?:PublicSettings['region'];documentLanguage?:PublicSettings['documentLanguage'];clearRecent?:boolean}):Promise<Result<PublicSettings>>;
  recover(id:string):Promise<Result<{bytes:Uint8Array;metadata:RecoveryItem}>>;
  discardRecovery(id:string):Promise<Result<void>>;
  checkpoint(request:{documentId:string;contentId:string;revision:number;bytes:Uint8Array}):Promise<Result<void>>;
  onAction(callback:(action:MenuAction)=>void):()=>void;
  onFile(callback:(file:OpenFileResult)=>void):()=>void;
}
declare global { interface Window { office:DesktopBridge } }
