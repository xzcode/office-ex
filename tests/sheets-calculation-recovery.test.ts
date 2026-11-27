import test from 'node:test';
import assert from 'node:assert/strict';
import { CancellationToken, OfficeError } from '@office-ex/contracts';
import { createDocument, DocumentSession, type CellValue, type SheetsDocument } from '@office-ex/core';
import { WorkbookCalculator, calculationKey, setCells, parseFormula, type WorkbookSnapshot } from '@office-ex/sheets';

const formula = (source: string): CellValue => ({kind:'formula',source});
const number = (value: number): CellValue => ({kind:'number',value});
const string = (value: string): CellValue => ({kind:'string',value});
const boolean = (value: boolean): CellValue => ({kind:'boolean',value});
const err = (value: Extract<CellValue,{kind:'error'}>['value']): CellValue => ({kind:'error',value});
function book(values: Record<string,CellValue>, id = 'calc-book'): SheetsDocument {
  const doc = createDocument('sheets',id);
  for (const [key,value] of Object.entries(values)) doc.sheets[0].cells[key] = {value};
  return doc;
}
function snapshot(session: DocumentSession): WorkbookSnapshot {
  const doc = session.snapshot().document;
  if (doc.kind !== 'sheets') throw new Error('expected sheets');
  return doc;
}
const sid = 'sheet-1';
const code = (expected: string) => (e: unknown) => e instanceof OfficeError && e.code === expected;

test('SHT-03 arithmetic, comparison, coercion, error and text semantics', () => {
  const doc=book({'1:1':formula('=1+2*3'),'1:2':formula('=-2^2'),'1:3':formula('=2^3^2'),'1:4':formula('=10%+TRUE'),'1:5':formula('="2"+3'),'1:6':formula('="abc"+1'),'1:7':formula('=1/0'),'1:8':formula('=0^0'),'1:9':formula('="abc"="ABC"'),'1:10':formula('="+" & "(" & ")" & ","'),'1:11':formula('=Z99'),'1:12':formula('=1e308*1e308'),'1:13':formula('=2<"1"')});
  const result=new WorkbookCalculator().calculate(doc);
  const expected=[number(7),number(4),number(64),number(1.1),number(5),err('#VALUE!'),err('#DIV/0!'),err('#NUM!'),boolean(true),string('+(),'),number(0),err('#NUM!'),boolean(true)];
  expected.forEach((value,index)=>assert.deepEqual(result.get(sid,1,index+1).value,value));
});
test('SHT-03 all twelve P0 function IDs, range coercion and empty averages', () => {
  const doc=book({'1:1':number(4),'2:1':number(8),'3:1':string('3'),'4:1':boolean(true),'5:1':string(''),'6:1':err('#N/A'),
    '1:2':formula('=SUM(A1:A5,TRUE,"2")'),'2:2':formula('=AVERAGE(A1:A5)'),'3:2':formula('=MIN(A1:A5)'),'4:2':formula('=MAX(A1:A5)'),'5:2':formula('=COUNT(A1:A6,TRUE,"2")'),'6:2':formula('=COUNTA(A1:A6)'),'7:2':formula('=IF(FALSE,1/0,7)'),'8:2':formula('=AND(A1:A5,TRUE)'),'9:2':formula('=OR(FALSE,0,1)'),'10:2':formula('=NOT(FALSE)'),'11:2':formula('=ROUND(-1.005,2)+ABS(-2)'),'12:2':formula('=AVERAGE(Z1:Z9)'),'13:2':formula('=SUM(A1:A6)'),'14:2':formula('=ROUND(125,-1)'),'15:2':formula('=ABS(-9)')});
  const r=new WorkbookCalculator().calculate(doc);
  const expected=[number(15),number(6),number(4),number(8),number(4),number(6),number(7),boolean(true),boolean(true),boolean(true),number(.99),err('#DIV/0!'),err('#N/A'),number(130),number(9)];
  expected.forEach((value,i)=>assert.deepEqual(r.get(sid,i+1,2).value,value));
});
test('SHT-04 sparse full-column/row ranges do not expand empty cells', () => {
  const doc=book({'1:1':number(5),'1048576:1':number(7),'1048576:16384':number(11),'2:2':formula('=SUM(A:A)'),'3:2':formula('=SUM(1048576:1048576)')});
  const r=new WorkbookCalculator().calculate(doc,{maxSteps:200});
  assert.deepEqual(r.get(sid,2,2).value,number(12));
  assert.deepEqual(r.get(sid,3,2).value,number(18));
  assert.ok(r.steps<200);
});
test('SHT-04 quoted, case-insensitive cross-sheet references resolve stable sheet IDs', () => {
  const doc=book({'1:1':formula("=SUM('Sales 2026'!A1:A2)"),'1:2':formula("='sALES 2026'!$A$2"),'1:3':formula('=Missing!A1')});
  doc.sheets.push({id:'sales',name:'Sales 2026',cells:{'1:1':{value:number(8)},'2:1':{value:number(12)}},rows:{},columns:{},merges:[]});
  const r=new WorkbookCalculator().calculate(doc);
  assert.deepEqual(r.get(sid,1,1).value,number(20));
  assert.deepEqual(r.get(sid,1,2).value,number(12));
  assert.deepEqual(r.get(sid,1,3).value,err('#REF!'));
});
test('SHT-04 only affected formulas recalculate after cell transaction, undo and redo', () => {
  const session=new DocumentSession(book({'1:1':number(2),'1:2':formula('=A1+1'),'1:3':formula('=B1*2'),'1:4':formula('=40+2'),'2:1':number(10),'2:2':formula('=SUM(A2:A4)')}));
  const calc=new WorkbookCalculator();
  const first=calc.calculate(snapshot(session));
  assert.equal(first.computedCells,4);
  assert.equal(calc.calculate(snapshot(session)).computedCells,0);
  session.execute(setCells(session.snapshot(),sid,[{row:1,column:1,cell:{value:number(5)}}]));
  const second=calc.calculate(snapshot(session));
  assert.equal(second.computedCells,2);
  assert.deepEqual(second.get(sid,1,3).value,number(12));
  assert.equal(second.get(sid,1,4),first.get(sid,1,4));
  session.undo(); assert.deepEqual(calc.calculate(snapshot(session)).get(sid,1,3).value,number(6));
  session.redo(); assert.deepEqual(calc.calculate(snapshot(session)).get(sid,1,3).value,number(12));
  session.execute(setCells(session.snapshot(),sid,[{row:4,column:1,cell:{value:number(3)}}]));
  const third=calc.calculate(snapshot(session));
  assert.equal(third.computedCells,1);
  assert.deepEqual(third.get(sid,2,2).value,number(13));
  session.execute(setCells(session.snapshot(),sid,[{row:4,column:1,cell:null}]));
  assert.deepEqual(calc.calculate(snapshot(session)).get(sid,2,2).value,number(10));
});
test('SHT-04 replacement or removal of formula rebuilds reverse dependencies', () => {
  const session=new DocumentSession(book({'1:1':number(1),'1:2':number(2),'1:3':formula('=A1'),'1:4':formula('=C1+1')}));
  const calc=new WorkbookCalculator();
  calc.calculate(snapshot(session));
  session.execute(setCells(session.snapshot(),sid,[{row:1,column:3,cell:{value:formula('=B1')}}]));
  assert.deepEqual(calc.calculate(snapshot(session)).get(sid,1,4).value,number(3));
  session.execute(setCells(session.snapshot(),sid,[{row:1,column:1,cell:{value:number(9)}}]));
  assert.equal(calc.calculate(snapshot(session)).computedCells,0);
  session.execute(setCells(session.snapshot(),sid,[{row:1,column:2,cell:{value:number(8)}}]));
  assert.equal(calc.calculate(snapshot(session)).computedCells,2);
  session.execute(setCells(session.snapshot(),sid,[{row:1,column:3,cell:null}]));
  const result=calc.calculate(snapshot(session));
  assert.equal(result.cells.has(calculationKey(sid,1,3)),false);
  assert.deepEqual(result.get(sid,1,4).value,number(1));
});
test('SHT-04 cycle participation cannot be hidden by COUNT/COUNTA; unused IF branches stay lazy', () => {
  const doc=book({'1:1':formula('=COUNT(B1)'),'1:2':formula('=COUNTA(A1)'),'1:3':formula('=A1+1'),'1:4':formula('=IF(FALSE,D1,42)'),'1:5':formula('=IF(TRUE,8,UNKNOWN(1))')});
  const r=new WorkbookCalculator().calculate(doc);
  for (const col of [1,2]) assert.deepEqual(r.get(sid,1,col),{state:'error',value:err('#REF!'),reason:'cycle'});
  assert.deepEqual(r.get(sid,1,3).value,err('#REF!'));
  assert.deepEqual(r.get(sid,1,4).value,number(42));
  assert.deepEqual(r.get(sid,1,5).value,number(8));
});
test('SHT-04 unknown formulas retain only matching cached values; dependents remain stale', () => {
  const doc=book({'1:1':formula('=UNSUPPORTED(1)'),'1:2':formula('=A1+1'),'1:3':formula('=[external.xlsx]Sheet1!A1')});
  const a=calculationKey(sid,1,1),b=calculationKey(sid,1,2),c=calculationKey(sid,1,3);
  const calc=new WorkbookCalculator();
  const cache = new Map([[a,number(3)],[b,number(4)],[c,number(99)]] as [string,Extract<CellValue,{kind:'number'}>][]);
  const sources = new Map([[a,'=UNSUPPORTED(1)'],[b,'=A1+1'],[c,'=different']]);
  const first=calc.calculate(doc,{cachedValues:cache,cachedSources:sources});
  assert.deepEqual(first.get(sid,1,1),{state:'unsupported',reason:'unknown-function',value:number(3)});
  assert.deepEqual(first.get(sid,1,2),{state:'stale',reason:'dependency',value:number(4)});
  assert.deepEqual(first.get(sid,1,3),{state:'unsupported',reason:'unsupported-syntax'});
  const refreshed=new Map(cache);refreshed.set(a,{kind:'number',value:7});
  assert.deepEqual(calc.calculate(doc,{cachedValues:refreshed,cachedSources:sources}).get(sid,1,1).value,number(7));
  const changed=book({'1:1':formula('=UNSUPPORTED(2)'),'1:2':formula('=A1+1')});
  assert.equal(calc.calculate(changed,{cachedValues:cache,cachedSources:sources}).get(sid,1,1).value,undefined);
  assert.equal(calc.calculate(book({'1:1':formula('=UNSUPPORTED(1)')},'new-book')).get(sid,1,1).value,undefined);
});
test('SHT-04 cancellation, step and deadline abort without committing partial results', () => {
  const first=book({'1:1':number(1),'1:2':formula('=A1+1')});
  const changed=book({'1:1':number(5),'1:2':formula('=A1+1')});
  const calc=new WorkbookCalculator();calc.calculate(first);
  assert.throws(()=>calc.calculate(changed,{maxSteps:1}),code('limit'));
  const cancellation=new CancellationToken(); cancellation.cancel();
  assert.throws(()=>calc.calculate(changed,{cancellation}),code('cancelled'));
  assert.throws(()=>calc.calculate(changed,{deadline:0,now:()=>1}),code('limit'));
  assert.throws(()=>calc.calculate(changed,{deadline:0}),code('invalid'));
  assert.equal(calc.calculate(first).computedCells,0);
  assert.deepEqual(calc.calculate(changed).get(sid,1,2).value,number(6));
  calc.clear();assert.equal(calc.calculate(first).computedCells,1);
});
test('SHT-04 long acyclic chain evaluates topologically with bounded recursion', () => {
  const doc=book({'1:1':number(1)});
  for(let row=2;row<=1500;row++) doc.sheets[0].cells[row+':1']={value:formula('=A'+(row-1)+'+1')};
  const r=new WorkbookCalculator().calculate(doc,{maxDepth:16,maxSteps:200000});
  assert.deepEqual(r.get(sid,1500,1).value,number(1500));
  assert.equal(r.computedCells,1499);
});
test('SHT-03 string tokens cannot be mistaken for operators or references', () => {
  for(const text of ['+','-','(',')',',','%','!','A1']) assert.deepEqual(parseFormula('="'+text+'"'),{kind:'string',value:text});
  assert.throws(()=>parseFormula("='A1'"),code('unsupported'));
  assert.throws(()=>parseFormula('=A1:"B2"'),code('unsupported'));
  assert.equal(parseFormula('=IF(TRUE,")",",")').kind,'call');
});

test('M6.5 common scalar, text and predicate functions calculate with explicit error semantics',()=>{
 const cases:[string,CellValue][]=[
  ['=IFERROR(1/0,42)',number(42)],['=IFERROR(3,1/0)',number(3)],['=ISNUMBER(3)',boolean(true)],['=ISTEXT("3")',boolean(true)],['=ISLOGICAL(TRUE)',boolean(true)],
  ['=ISBLANK(Z99)',boolean(true)],['=ISBLANK("")',boolean(false)],['=ISERROR(#N/A)',boolean(true)],['=ISERR(#N/A)',boolean(false)],['=ISNA(#N/A)',boolean(true)],
  ['=PRODUCT(2,3,4)',number(24)],['=INT(-2.4)',number(-3)],['=MOD(-5,3)',number(1)],['=MOD(5,-3)',number(-1)],['=MOD(1,0)',err('#DIV/0!')],
  ['=ROUNDUP(-1.21,1)',number(-1.3)],['=ROUNDDOWN(-1.29,1)',number(-1.2)],['=TRUNC(-12.98,1)',number(-12.9)],['=TRUNC(-12.98)',number(-12)],
  ['=SIGN(-7)',number(-1)],['=SQRT(81)',number(9)],['=SQRT(-1)',err('#NUM!')],['=POWER(3,2)',number(9)],['=POWER(0,-1)',err('#DIV/0!')],
  ['=LEN("中文😀")',number(3)],['=LEFT("中文",1)',string('中')],['=RIGHT("abcd",2)',string('cd')],['=RIGHT("abcd",0)',string('')],['=MID("abcdef",2,3)',string('bcd')],['=MID("abc",0,2)',err('#VALUE!')],
  ['=UPPER("abc")',string('ABC')],['=LOWER("ABC")',string('abc')],['=TRIM("  A   B  ")',string('A B')],['=CONCATENATE("A",2,TRUE)',string('A2TRUE')]
 ];
 const d=book(Object.fromEntries(cases.map(([source],i)=>[(i+1)+':1',formula(source)])));
 const result=new WorkbookCalculator().calculate(d);
 cases.forEach(([,expected],i)=>assert.deepEqual(result.get(sid,i+1,1).value,expected,cases[i]![0]));
});
test('M6.5 COUNTBLANK uses sparse implicit cells, distinguishes empty strings and errors',()=>{
 const d=book({'1:1':number(2),'2:1':string(''),'3:1':err('#REF!'),'4:1':formula('=""'),'1:2':formula('=COUNTBLANK(A:A)'),'2:2':formula('=COUNTBLANK(A3)'),'3:2':formula('=COUNTBLANK(A4)')});
 const r=new WorkbookCalculator().calculate(d,{maxSteps:500});
 assert.deepEqual(r.get(sid,1,2).value,number(1048574));
 assert.deepEqual(r.get(sid,2,2).value,number(0));assert.deepEqual(r.get(sid,3,2).value,number(1));
 assert.ok(r.steps<500);
});
test('M6.5 IFERROR is lazy, does not promote unsupported caches or hide cycles, and recalculates dependencies',()=>{
 const s=new DocumentSession(book({'1:1':number(0),'1:2':formula('=IFERROR(10/A1,99)'),'2:2':formula('=IFERROR(1,UNSUPPORTED(1))'),'3:2':formula('=IFERROR(UNSUPPORTED(1),42)'),'4:2':formula('=IFERROR(B4,5)')}));
 const calc=new WorkbookCalculator(),r=calc.calculate(snapshot(s));
 assert.deepEqual(r.get(sid,1,2).value,number(99));assert.deepEqual(r.get(sid,2,2).value,number(1));
 assert.equal(r.get(sid,3,2).state,'unsupported');assert.deepEqual(r.get(sid,4,2),{state:'error',value:err('#REF!'),reason:'cycle'});
 s.execute(setCells(s.snapshot(),sid,[{row:1,column:1,cell:{value:number(2)}}]));
 assert.deepEqual(calc.calculate(snapshot(s)).get(sid,1,2).value,number(5));
 s.undo();assert.deepEqual(calc.calculate(snapshot(s)).get(sid,1,2).value,number(99));
});

test('M6.5 conditional aggregates pair relative range positions, match criteria and wildcards, and handle sparse blanks',()=>{
 const d=book({'1:1':string('Apple'),'2:1':string('apple'),'3:1':string('Pear'),'4:1':string('A*'),'1:2':string('East'),'2:2':string('West'),'3:2':string('East'),'4:2':string('East'),'1:3':number(10),'2:3':number(20),'3:3':number(30),'4:3':number(40)});
 const cases:[string,CellValue][]=[
  ['=COUNTIF(A1:A4,"apple")',number(2)],['=COUNTIF(A1:A4,"A*")',number(3)],['=COUNTIF(A1:A4,"A~*")',number(1)],['=COUNTIF(A:A,"")',number(1048572)],
  ['=COUNTIF(C1:C4,">=20")',number(3)],['=COUNTIFS(A1:A4,"A*",B1:B4,"East")',number(2)],['=SUMIF(A1:A4,"apple",C1:C4)',number(30)],
  ['=SUMIFS(C1:C4,B1:B4,"East",C1:C4,">15")',number(70)],['=AVERAGEIF(B1:B4,"East",C1:C4)',number(80/3)],['=AVERAGEIFS(C1:C4,A1:A4,"A*",B1:B4,"East")',number(25)],
  ['=AVERAGEIF(B1:B4,"missing",C1:C4)',err('#DIV/0!')],['=COUNTIFS(A1:A4,"apple",B1:B3,"East")',err('#VALUE!')],['=COUNTIFS(A:A,"",B:B,"")',number(1048572)]
 ];
 cases.forEach(([source],i)=>d.sheets[0]!.cells[(i+1)+':5']={value:formula(source)});
 const r=new WorkbookCalculator().calculate(d,{maxSteps:10000});
 cases.forEach(([source,expected],i)=>assert.deepEqual(r.get(sid,i+1,5).value,expected,source));
});
test('M6.5 lookup functions support exact wildcard, sorted numeric approximate match and scalar INDEX',()=>{
 const d=book({'1:1':string('Apple'),'2:1':string('Pear'),'3:1':string('apple'),'1:2':number(10),'2:2':number(20),'3:2':number(30),'1:3':number(1),'2:3':number(5),'3:3':number(10),'1:4':number(100),'2:4':number(500),'3:4':number(1000),'5:1':number(10),'5:2':number(5),'5:3':number(1)});
 const cases:[string,CellValue][]=[
  ['=VLOOKUP("apple",A1:B3,2,FALSE)',number(10)],['=VLOOKUP("Pe*",A1:B3,2,0)',number(20)],['=VLOOKUP("missing",A1:B3,2,FALSE)',err('#N/A')],
  ['=VLOOKUP(7,C1:D3,2,TRUE)',number(500)],['=VLOOKUP(0,C1:D3,2)',err('#N/A')],['=VLOOKUP("apple",A1:B3,3,FALSE)',err('#REF!')],
  ['=MATCH("a*",A1:A3,0)',number(1)],['=MATCH(7,C1:C3,1)',number(2)],['=MATCH(7,A5:C5,-1)',number(1)],
  ['=INDEX(A1:B3,2,2)',number(20)],['=INDEX(B1:B3,MATCH("Pear",A1:A3,0))',number(20)],['=INDEX(A5:C5,2)',number(5)],
  ['=HLOOKUP(5,A5:C5,1,FALSE)',number(5)]
 ];
 cases.forEach(([source],i)=>d.sheets[0]!.cells[(i+1)+':6']={value:formula(source)});
 const r=new WorkbookCalculator().calculate(d);
 cases.forEach(([source,expected],i)=>assert.deepEqual(r.get(sid,i+1,6).value,expected,source));
});
test('M6.5 unsupported implicit range expansion, array INDEX and unsorted lookup remain protected',()=>{
 const d=book({'1:1':number(5),'2:1':number(1),'1:2':number(3),'2:2':number(4),
  '1:3':formula('=SUMIF(A1:A2,">0",B1)'),'2:3':formula('=INDEX(A1:B2,0,1)'),'3:3':formula('=VLOOKUP(3,A1:B2,2,TRUE)')});
 const r=new WorkbookCalculator().calculate(d);
 for(let i=1;i<=3;i++)assert.equal(r.get(sid,i,3).state,'unsupported');
});
test('M6.5 cross-sheet conditional and lookup formulas recalculate on edit and undo',()=>{
 const d=book({'1:1':formula("=SUMIF('Sales 2026'!A1:A2,\"East\",'Sales 2026'!B1:B2)"),'1:2':formula("=VLOOKUP(\"East\",'Sales 2026'!A1:B2,2,FALSE)")});
 d.sheets.push({id:'sales',name:'Sales 2026',cells:{'1:1':{value:string('East')},'2:1':{value:string('West')},'1:2':{value:number(10)},'2:2':{value:number(20)}},rows:{},columns:{},merges:[]});
 const s=new DocumentSession(d),calc=new WorkbookCalculator();
 let r=calc.calculate(snapshot(s));assert.deepEqual(r.get(sid,1,1).value,number(10));assert.deepEqual(r.get(sid,1,2).value,number(10));
 s.execute(setCells(s.snapshot(),'sales',[{row:1,column:2,cell:{value:number(50)}}]));
 r=calc.calculate(snapshot(s));assert.deepEqual(r.get(sid,1,1).value,number(50));assert.deepEqual(r.get(sid,1,2).value,number(50));
 s.undo();assert.deepEqual(calc.calculate(snapshot(s)).get(sid,1,1).value,number(10));
});
