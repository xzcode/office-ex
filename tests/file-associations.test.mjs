// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {fileTypes,desktopEntry,mimePackage,windowsRegistry,macDocumentDeclarations,progId} from '../scripts/file-associations.mjs';
test('M7.5 all eight types have consistent alternative native handler declarations without changing defaults',()=>{
 assert.deepEqual(fileTypes.map(f=>f.extension),['doc','docx','xls','xlsx','ppt','pptx','pdf','oex']);
 assert.equal(new Set(fileTypes.map(f=>f.mime)).size,8);assert.equal(new Set(fileTypes.map(f=>f.uti)).size,8);
 const linux=desktopEntry(),windows=windowsRegistry(),mac=macDocumentDeclarations();
 assert.match(linux,/^Exec=office-ex-preview %F$/m);
 assert.equal(linux.split('\n').find(v=>v.startsWith('MimeType=')).slice(9),fileTypes.map(f=>f.mime).join(';')+';');
 assert.match(mimePackage(),/pattern="\*\.oex"/);assert.match(mimePackage(),/application\/x-office-ex/);
 for(const f of fileTypes){
  assert.ok(windows.includes('ValueName: "'+progId(f.extension)+'"'));
  assert.ok(windows.includes('Subkey: "Software\\Classes\\.'+f.extension+'\\OpenWithProgids"'));
  assert.ok(windows.includes('ValueName: ".'+f.extension+'"; ValueData: "'+progId(f.extension)+'"'));
  assert.ok(mac.CFBundleDocumentTypes.includes('<string>'+f.uti+'</string>'));
  assert.ok(mac.CFBundleDocumentTypes.includes('<string>'+f.extension+'</string>'));
 }
 assert.ok(windows.includes('ValueData: """{app}\\Office EX.exe"" ""%1"""'));
 assert.doesNotMatch(windows,/UserChoice|Subkey: "Software\\Classes\\\.[^"\\]+";|Flags: uninsdeletekey(?:;|\r|$)/);
 assert.match(mac.CFBundleDocumentTypes,/<string>Viewer<\/string><key>LSHandlerRank<\/key><string>Alternate<\/string><key>LSItemContentTypes<\/key><array><string>com.adobe.pdf/);
 assert.match(mac.UTExportedTypeDeclarations,/<string>com.officex.preview.draft<\/string>/);
});
