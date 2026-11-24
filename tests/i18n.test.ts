// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { availableLocales, TARGET_LOCALES, resolveLocale, resolveRegion, localeDirection, message, validateCatalogs, formatNumber, formatDate, isLanguagePreference, diagnosticKey, errorKey, formatTemplate, formatPlural, formatSelect, pseudoLocalize, localeLabel } from '@office-ex/i18n';
test('I18N registry has 22 distinct target locales, eleven complete current catalogs, matching keys and placeholders',()=>{
  assert.equal(new Set(TARGET_LOCALES.map(x=>x[0])).size,22);
  assert.deepEqual(availableLocales(),['zh-CN','zh-TW','en-US','ja-JP','ko-KR','fr-FR','de-DE','es-ES','pt-BR','ru-RU','it-IT']);
  validateCatalogs();
  assert.equal(isLanguagePreference('ar'),false);
  assert.equal(isLanguagePreference(['en-US']),false);
  assert.equal(localeDirection('ar'),'rtl');assert.equal(localeDirection('he-IL'),'rtl');assert.equal(localeDirection('ja-JP'),'ltr');
});
test('I18N exact locale, script-aware Chinese, preferred-list fallback and explicit selection',()=>{
  for(const [requested,expected] of [['zh-Hant-HK','zh-TW'],['zh-Hans-TW','zh-CN'],['zh-SG','zh-CN'],['en-GB','en-US'],['ja','ja-JP'],['xx-YY','en-US'],['garbage locale','en-US']] as const)assert.equal(resolveLocale('system',[requested]),expected);
  assert.equal(resolveLocale('system',['ko-KR','ja-JP']),'ko-KR');
  assert.equal(resolveLocale('system',['ar','ja-JP']),'ja-JP');
  assert.equal(resolveLocale('system',['de-AT']),'de-DE');
  assert.equal(resolveLocale('system',['pt-PT']),'pt-BR');
  assert.equal(resolveLocale('system',['es-MX']),'es-ES');
  assert.equal(resolveLocale('zh-TW',['en-US']),'zh-TW');
});
test('I18N messages preserve dynamic parameters and reject missing ones; unknown keys cannot access prototypes',()=>{
  assert.equal(message('en-US','entriesLimited',{count:750}),'750 entries; first 500 shown.');
  assert.match(message('ja-JP','entriesLimited',{count:750}),/750/);
  assert.throws(()=>message('en-US','entriesLimited'));
  assert.equal(diagnosticKey('__proto__'),'diagUnknown');
  assert.equal(errorKey('constructor'),'errorInternal');
});
test('I18N independent regional formatting retains source numbers and dates',()=>{
  assert.equal(resolveRegion('de-DE',['ja-JP']),'de-DE');
  assert.equal(resolveRegion('system',['en-GB']),'en-GB');
  assert.equal(formatNumber('de-DE',1234.5),'1.234,5');
  assert.equal(formatNumber('en-US',1234.5),'1,234.5');
  assert.equal(formatDate('en-GB',Date.UTC(2026,10,24),{year:'numeric',month:'2-digit',day:'2-digit',timeZone:'UTC'}),'24/11/2026');
});

test('I18N all seven new catalogs expose translated save actions and valid localized language names',()=>{
  const expected={'ko-KR':'저장','fr-FR':'Enregistrer','de-DE':'Speichern','es-ES':'Guardar','pt-BR':'Salvar','ru-RU':'Сохранить','it-IT':'Salva'} as const;
  for(const locale of Object.keys(expected) as (keyof typeof expected)[]){assert.equal(message(locale,'save'),expected[locale]);assert.ok(localeLabel(locale,'zh-CN').length>0);assert.match(message(locale,'entriesLimited',{count:700}),/700/);}
});
test('I18N structured plural forms follow actual language rules and region formatting',()=>{
  const forms={one:'{count} файл',few:'{count} файла',many:'{count} файлов',other:'{count} файла'};
  assert.equal(formatPlural('ru-RU',1,forms),'1 файл');
  assert.equal(formatPlural('ru-RU',2,forms),'2 файла');
  assert.equal(formatPlural('ru-RU',5,forms),'5 файлов');
  assert.equal(formatPlural('de-DE',1000,{other:'{count} Dateien'}, {},'de-DE'),'1.000 Dateien');
  assert.throws(()=>formatPlural('en-US',NaN,{other:'{count} files'}));
});
test('I18N structured select and template refuse unsafe parameters without evaluating content',()=>{
  assert.equal(formatSelect('readonly',{readonly:'Read only: {title}',other:'Other'},{title:'<b>user text</b>'}),'Read only: <b>user text</b>');
  assert.equal(formatSelect('__proto__',{other:'Other'}),'Other');
  assert.throws(()=>formatTemplate('{name}',{get name():string{throw Error('accessor must not execute');}}));
  assert.throws(()=>formatTemplate('{count}',{count:Infinity}));
});
test('I18N pseudo messages expand visible labels while preserving parameters and optional RTL isolation',()=>{
  const source='Save document {title}';
  const expanded=pseudoLocalize(source);
  assert.ok(expanded.length>source.length*1.25);
  assert.ok(expanded.includes('{title}'));
  const rtl=pseudoLocalize(source,'rtl');assert.ok(rtl.startsWith('\u2067'));assert.ok(rtl.endsWith('\u2069'));
  assert.ok(formatTemplate(rtl,{title:'中文'}).includes('中文'));
});
