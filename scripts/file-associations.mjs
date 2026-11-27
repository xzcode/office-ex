// SPDX-License-Identifier: MIT
/** Shared installer declarations. Register an available handler; never
 * assign the user's default application or modify Windows UserChoice. */
import {readFileSync} from 'node:fs';
export const fileTypes=Object.freeze(JSON.parse(readFileSync(new URL('../apps/desktop/file-types.json',import.meta.url),'utf8')).map(Object.freeze));
export const applicationName='Office EX Preview';
export const windowsCapabilities='Software\\OfficeEXPreview\\Capabilities';
const q=s=>'"'+s.replaceAll('"','""')+'"';
export const progId=extension=>'OfficeEXPreview.'+extension;
export const desktopEntry=()=>['[Desktop Entry]','Type=Application','Name='+applicationName,'Exec=office-ex-preview %F','Terminal=false','Icon=office-ex-preview','StartupWMClass=Office EX','Categories=Office;','MimeType='+fileTypes.map(f=>f.mime).join(';')+';',''].join('\n');
export const mimePackage=()=> '<?xml version="1.0" encoding="UTF-8"?>\n<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info"><mime-type type="application/x-office-ex"><comment>Office EX draft</comment><glob pattern="*.oex"/></mime-type></mime-info>\n';
export function windowsRegistry(){
 const lines=['[Registry]'];
 const add=(key,name,value,flags)=>lines.push('Root: HKA; Subkey: '+q(key)+'; ValueType: string; ValueName: '+q(name)+'; ValueData: '+q(value)+'; Flags: '+flags);
 const classes='Software\\Classes\\',exe='Applications\\Office EX.exe',owned='uninsdeletevalue uninsdeletekeyifempty';
 add('Software\\RegisteredApplications',applicationName,windowsCapabilities,'uninsdeletevalue');
 add(windowsCapabilities,'ApplicationName',applicationName,owned);
 add(windowsCapabilities,'ApplicationDescription','Offline office documents and PDF viewer',owned);
 add(windowsCapabilities,'ApplicationIcon','{app}\\Office EX.exe,0',owned);
 add(classes+exe,'FriendlyAppName',applicationName,owned);
 add(classes+exe+'\\shell\\open\\command','','"{app}\\Office EX.exe" "%1"',owned);
 for(const f of fileTypes){
  const ext='.'+f.extension,prog=progId(f.extension);
  add(classes+prog,'',f.name,owned);
  add(classes+prog+'\\DefaultIcon','','{app}\\Office EX.exe,0',owned);
  add(classes+prog+'\\shell\\open\\command','','"{app}\\Office EX.exe" "%1"',owned);
  add(classes+ext+'\\OpenWithProgids',prog,'','uninsdeletevalue uninsdeletekeyifempty');
  add(classes+exe+'\\SupportedTypes',ext,'',owned);
  add(windowsCapabilities+'\\FileAssociations',ext,prog,owned);
 }
 // Inno's private root keys may be removed as a whole; never delete an
 // extension's shared association key. Uninstall cleanup is scoped.
 lines.push('Root: HKA; Subkey: "Software\\OfficeEXPreview"; Flags: uninsdeletekeyifempty');
 return lines.join('\r\n')+'\r\n';
}
const escape=s=>s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const str=s=>'<string>'+escape(s)+'</string>',kv=(key,v)=>'<key>'+key+'</key>'+v,arr=a=>'<array>'+a.join('')+'</array>',dict=a=>'<dict>'+a.join('')+'</dict>';
export function macDocumentDeclarations(){
 const types=fileTypes.map(f=>dict([kv('CFBundleTypeName',str(f.name)),kv('CFBundleTypeRole',str(f.role)),kv('LSHandlerRank',str(f.extension==='oex'?'Owner':'Alternate')),kv('LSItemContentTypes',arr([str(f.uti)])),kv('CFBundleTypeExtensions',arr([str(f.extension)])),kv('CFBundleTypeIconFile',str('office-ex.icns'))]));
 const own=fileTypes.find(f=>f.extension==='oex');
 return {CFBundleDocumentTypes:arr(types),UTExportedTypeDeclarations:arr([dict([kv('UTTypeIdentifier',str(own.uti)),kv('UTTypeDescription',str(own.name)),kv('UTTypeConformsTo',arr([str('public.data'),str('public.content')])),kv('UTTypeTagSpecification',dict([kv('public.filename-extension',arr([str(own.extension)])),kv('public.mime-type',str(own.mime))]))])])};
}
export const plistDocument=value=>'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">'+value+'</plist>\n';
