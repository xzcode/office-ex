// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { availableLocales, TARGET_LOCALES, resolveLocale, resolveRegion, localeDirection, message, validateCatalogs, formatNumber, formatDate, isLanguagePreference, diagnosticKey, errorKey } from '@office-ex/i18n';
test('I18N registry has 22 distinct target locales, four complete current catalogs, matching keys and placeholders',()=>{
  assert.equal(new Set(TARGET_LOCALES.map(x=>x[0])).size,22);
  assert.deepEqual(availableLocales(),['zh-CN','zh-TW','en-US','ja-JP']);
  validateCatalogs();
  assert.equal(isLanguagePreference('ar'),false);
  assert.equal(isLanguagePreference(['en-US']),false);
  assert.equal(localeDirection('ar'),'rtl');assert.equal(localeDirection('he-IL'),'rtl');assert.equal(localeDirection('ja-JP'),'ltr');
});
test('I18N exact locale, script-aware Chinese, preferred-list fallback and explicit selection',()=>{
  for(const [requested,expected] of [['zh-Hant-HK','zh-TW'],['zh-Hans-TW','zh-CN'],['zh-SG','zh-CN'],['en-GB','en-US'],['ja','ja-JP'],['xx-YY','en-US'],['garbage locale','en-US']] as const)assert.equal(resolveLocale('system',[requested]),expected);
  assert.equal(resolveLocale('system',['ko-KR','ja-JP']),'ja-JP');
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
