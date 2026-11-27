import test from 'node:test';
import assert from 'node:assert/strict';
import { CancellationToken, OfficeError } from '@office-ex/contracts';
import { columnName, columnNumber, formulaIsSupported, localizeFormula, parseFormula, serializeFormula, walkFormula } from '@office-ex/sheets';
const code = (expected: string) => (error: unknown): boolean => error instanceof OfficeError && error.code === expected;

test('SHT-03 recovered parser preserves operator precedence and associativity', () => {
  const ast = parseFormula('=1+2*3^4&"!"');
  assert.equal(ast.kind, 'binary');
  assert.equal(serializeFormula(ast), '=((1+(2*(3^4)))&"!")');
  assert.equal(serializeFormula(parseFormula('=2^3^2')), '=((2^3)^2)');
});
test('SHT-03 absolute, relative and mixed references survive AST serialization', () => {
  const ast = parseFormula("='Sheet O''Brien'!$A1+B$2");
  assert.deepEqual(ast, {kind:'binary',operator:'+',left:{kind:'reference',sheet:"Sheet O'Brien",row:1,column:1,rowAbsolute:false,columnAbsolute:true},right:{kind:'reference',row:2,column:2,rowAbsolute:true,columnAbsolute:false}});
  assert.deepEqual(parseFormula(serializeFormula(ast)), ast);
});
test('SHT-03 ranges are represented by bounds rather than expanded', () => {
  const ast = parseFormula('=SUM($A:$XFD,1:1048576)');
  assert.equal(ast.kind,'call');
  if (ast.kind !== 'call') throw new Error('Expected call');
  assert.equal(ast.args[0]?.kind,'range');
  assert.equal(ast.args[1]?.kind,'range');
  assert.deepEqual(parseFormula(serializeFormula(ast)), ast);
  assert.equal(serializeFormula(ast),'=SUM($A:$XFD,1:1048576)');
});
test('SHT-03 literal error, percent, missing argument and boolean nodes', () => {
  const ast = parseFormula('=IF(TRUE,10%,IF(FALSE,,#REF!))');
  const visited = new Set<string>();
  walkFormula(ast, (node) => visited.add(node.kind));
  for (const kind of ['call','boolean','unary','number','missing','error']) assert.ok(visited.has(kind));
  assert.deepEqual(parseFormula(serializeFormula(ast)),ast);
});
test('I18N-03 German formula alias and decimal separators are AST-equivalent', () => {
  const canonical = '=IF(A1="a,b",1.5,SUM(2,3))';
  const localized = localizeFormula(canonical,'de-DE');
  assert.ok(localized.includes('WENN('));
  assert.ok(localized.includes('1,5'));
  assert.ok(localized.includes('"a,b"'));
  assert.deepEqual(parseFormula(localized,{mode:'regional',locale:'de-DE'}),parseFormula(canonical));
});
test('I18N-03 French formula alias preserves quoted sheet names and strings', () => {
  const canonical = "='a,b.1'!A1+SUM(1.25,2)";
  const localized = localizeFormula(canonical,'fr-FR');
  assert.ok(localized.includes("'a,b.1'!"));
  assert.ok(localized.includes('SOMME(1,25;2)'));
  assert.deepEqual(parseFormula(localized,{mode:'regional',locale:'fr-FR'}),parseFormula(canonical));
});
test('SHT-03 unknown formula call remains parsed but is not declared supported', () => {
  const ast = parseFormula('=XLOOKUP(A1,B1,C1)');
  assert.equal(formulaIsSupported(ast),false);
  assert.equal(serializeFormula(ast),'=XLOOKUP(A1,B1,C1)');
  assert.equal(formulaIsSupported(parseFormula('=SUM(A1:B2)')),true);
});
test('SHT-03 incomplete input fails with a controlled OfficeError', () => {
  for (const source of ['=1+','=SUM(1,','=IF(','=A1:','=(', '="unterminated']) {
    assert.throws(()=>parseFormula(source), (error:unknown)=>error instanceof OfficeError);
  }
});
test('SHT-03 column and row limits and unsupported syntax are enforced', () => {
  assert.equal(columnName(16384),'XFD');
  assert.equal(columnNumber('XFD'),16384);
  assert.throws(()=>columnNumber('XFE'),code('invalid'));
  assert.throws(()=>parseFormula('=A1048577'),code('invalid'));
  assert.throws(()=>parseFormula('=[other.xlsx]Sheet1!A1'),code('unsupported'));
  assert.throws(()=>parseFormula('=SUM(Sheet1!A1:Sheet2!A2)'),code('unsupported'));
});
test('SHT-03 parser budgets and cancellation fail closed', () => {
  assert.throws(()=>parseFormula('=1+2+3',{maxTokens:2}),code('limit'));
  assert.throws(()=>parseFormula('=((((1))))',{maxDepth:3}),code('limit'));
  const cancellation = new CancellationToken();
  cancellation.cancel();
  assert.throws(()=>parseFormula('=1+2',{cancellation}),code('cancelled'));
});
