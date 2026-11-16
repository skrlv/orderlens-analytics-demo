import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { parseCSV, readWorkbook, workbookGrid, headersAt, mapRows, MAX_ROWS } from '../src/import.mjs';
import { validateRows, compute, money, FIELDS } from '../src/engine.mjs';
import { csvReport, jsonReport, reportRecords, RULES } from '../src/export.mjs';

// Read the exact locally shipped reader. No network, new dependency or Excel
// execution is involved; workbook objects below are in-memory parser fixtures.
const vendorContext = {};
vm.runInNewContext(fs.readFileSync(new URL('../public/vendor/xlsx.full.min.js', import.meta.url), 'utf8'), vendorContext);
const XLSX = vendorContext.XLSX;
const sample = name => fs.readFileSync(new URL(`../public/samples/${name}`, import.meta.url));
const fixture = JSON.parse(sample('fixtures.json'));
const expected = JSON.parse(sample('manual-expected.json'));
const canonicalMapping = Object.fromEntries(FIELDS.map((field, i) => [field, String(i)]));
const cells = values => values.map(v => ({t:'s',v}));
const canonicalGrid = () => [cells(FIELDS),...fixture.data.map(cells)];
function normalized(grid,{mapping=canonicalMapping,headerRow=1,sheet='CSV',book,decimal='.',mode='parts'}={}) {
  const mapped = mapRows(grid,mapping,{headerRow,sheet,book,XLSX,decimal});
  const validation = validateRows(mapped.rows,{headerRow,sheet,decimal,mode});
  return {mapped,validation};
}
const csvGrid = parseCSV(sample('orders-parts.csv').toString('utf8'), ',');
const csvDataset = normalized(csvGrid);
assert.equal(csvDataset.validation.ok,true);
assert.equal(csvDataset.mapped.errors.length,0);
const meta = {
  name:'orders-parts.csv',demo:true,mode:'parts',decimal:'.',headerRow:1,sheet:'CSV',
  mapping:canonicalMapping,stats:csvDataset.validation.stats,diagnostics:csvDataset.validation.diagnostics,
};
const report = compute(csvDataset.validation.events,{from:'2026-01-01',to:'2026-02-28',currency:'RUB',grain:'month'});

test('CSV supports BOM, UTF-8, quoted delimiters, escaped quotes and quoted newlines', () => {
  const rows=parseCSV('\ufeffname;note;empty\r\n"Заказ; 001";"Он сказал ""да""\r\nПроверено";\r\n',';');
  assert.deepEqual(rows.map(row=>row.map(c=>c.v)),[['name','note','empty'],['Заказ; 001','Он сказал "да"\r\nПроверено','']]);
  assert.deepEqual(parseCSV('a\tb\n1\t2','\t').map(r=>r.map(c=>c.v)),[['a','b'],['1','2']]);
  assert.deepEqual(parseCSV('a,b\r1,2\r',',').map(r=>r.map(c=>c.v)),[['a','b'],['1','2']]);
});
test('CSV rejects malformed quoting and unsupported delimiters', () => {
  for(const input of ['a,b\n"unfinished,2','a,b\n"closed"x,2','a,b\nun"escaped,2']) assert.throws(()=>parseCSV(input,','),/CSV/);
  assert.throws(()=>parseCSV('',';'),/пуст/);
  assert.throws(()=>parseCSV('a|b','|'),/разделитель/);
});
test('CSV enforces row/column limits before mapping', () => {
  assert.throws(()=>parseCSV(Array(65).fill('x').join(',')),/64/);
  assert.throws(()=>parseCSV(Array(MAX_ROWS+102).fill('x').join('\n')),/10 000/);
});
test('mapping and selectable header row keep physical row diagnostics', () => {
  const grid=[cells(['Служебный заголовок']),...canonicalGrid()];
  assert.deepEqual(headersAt(grid,2),FIELDS);
  const result=normalized(grid,{headerRow:2,sheet:'Лист 2'});
  assert.equal(result.validation.ok,true);
  assert.equal(result.validation.diagnostics[0].row,5);
  assert.equal(result.validation.diagnostics[0].sheet,'Лист 2');
  assert.deepEqual(result.validation.events,csvDataset.validation.events);
  for(const headerRow of [0,101,1.5])assert.throws(()=>headersAt(grid,headerRow));
  assert.throws(()=>headersAt(grid,90),/отсутствует/);
});
test('mapping supports arbitrary source column order and visible header labels', () => {
  const order=[7,6,5,4,3,2,1,0];
  const reversed=canonicalGrid().map(row=>order.map(index=>row[index]));
  const mapping=Object.fromEntries(FIELDS.map((field,index)=>[field,String(7-index)]));
  assert.deepEqual(headersAt(reversed,1),FIELDS.toReversed());
  const result=normalized(reversed,{mapping});
  assert.equal(result.validation.ok,true);
  assert.deepEqual(result.validation.events,csvDataset.validation.events);
  assert.throws(()=>mapRows(reversed,{event_id:'0',order_id:'0'},{}),/отдельная колонка/);
  const missing=normalized(canonicalGrid(),{mapping:{...canonicalMapping,customer_id:''}});
  assert.equal(missing.validation.ok,false);
  assert.ok(missing.validation.diagnostics.some(d=>d.field==='customer_id'));
});
test('a short CSV row leaves missing fields unknown rather than zero', () => {
  const grid=parseCSV(FIELDS.join(',')+'\nS,L,O,C,2026-01-01,sale');
  const result=normalized(grid);
  assert.equal(result.validation.ok,false);
  assert.deepEqual(result.validation.events,[]);
  assert.ok(result.validation.diagnostics.some(d=>d.field==='amount'&&d.severity==='error'));
  assert.ok(result.validation.diagnostics.some(d=>d.field==='currency'&&d.severity==='error'));
});
test('numeric Excel amount uses the explicitly selected comma separator', () => {
  const grid=canonicalGrid();
  grid[1][6]={t:'n',v:600.25};
  const result=mapRows(grid,canonicalMapping,{headerRow:1,sheet:'XLSX',decimal:',',XLSX});
  assert.equal(result.rows[0].amount,'600,25');
  const single=validateRows([result.rows[0]],{mode:'parts',decimal:','});
  assert.equal(single.ok,true);
  assert.equal(single.events[0].amount,60025n);
});
test('Excel amounts above the safe numeric boundary require decimal text', () => {
  const grid=canonicalGrid();
  for(const amount of [90071992547410,90071992547408.01,1000000000000.01]){
    grid[1][6]={t:'n',v:amount};
    const numeric=normalized(grid);
    assert.ok(numeric.mapped.errors.some(d=>d.field==='amount'&&/текст/i.test(d.message)));
    assert.equal(numeric.validation.ok,false);
  }
  grid[1][6]={t:'s',v:'90071992547410.01'};
  const textual=mapRows(grid,canonicalMapping,{XLSX});
  assert.equal(textual.errors.length,0);
  const single=validateRows([textual.rows[0]],{mode:'parts'});
  assert.equal(single.ok,true);assert.equal(single.events[0].amount,9007199254741001n);
});
test('semicolon/comma-decimal fixture matches dot-decimal CSV exactly', () => {
  const ru=normalized(parseCSV(sample('orders-parts-ru.csv').toString('utf8'),';'),{decimal:','});
  assert.equal(ru.validation.ok,true);
  assert.deepEqual(ru.validation.events,csvDataset.validation.events);
});
test('real shipped CSV and XLSX normalize to the same events and manual totals', () => {
  const book=readWorkbook(new Uint8Array(sample('orders-parts.xlsx')),XLSX);
  assert.equal(book.SheetNames.length,1);
  const sheet=book.SheetNames[0];
  const grid=workbookGrid(book,sheet,XLSX);
  assert.deepEqual(headersAt(grid,1),FIELDS);
  const xlsx=normalized(grid,{sheet,book});
  assert.equal(xlsx.mapped.errors.length,0);
  assert.equal(xlsx.validation.ok,true);
  assert.deepEqual(xlsx.validation.events,csvDataset.validation.events);
  for(const e of expected.cases){
    const result=compute(xlsx.validation.events,e);
    for(const key of ['gross','refunds','net','aov'])assert.equal(money(result[key]),e[key]);
    for(const key of ['orders','eligibleOrders','customers','repeatCustomers','repeatBasisPoints'])assert.equal(result[key],e[key]);
  }
});
test('XLSX reader rejects renamed text, legacy containers and truncated ZIPs', () => {
  for(const bytes of [new TextEncoder().encode('event_id,amount\nS1,100'),new Uint8Array([0xd0,0xcf,0x11,0xe0]),new Uint8Array([])])assert.throws(()=>readWorkbook(bytes,XLSX),/XLSX/);
  assert.throws(()=>readWorkbook(new Uint8Array([0x50,0x4b,3,4]),XLSX));
});
test('real shipped invalid CSV and XLSX both expose conflict/missing/invalid history', () => {
  const badCSV=normalized(parseCSV(sample('orders-errors.csv').toString('utf8'),','));
  const book=readWorkbook(new Uint8Array(sample('orders-errors.xlsx')),XLSX);
  const sheet=book.SheetNames[0];
  const badXLSX=normalized(workbookGrid(book,sheet,XLSX),{sheet,book});
  for(const r of [badCSV,badXLSX]){
    assert.equal(r.validation.ok,false);assert.deepEqual(r.validation.events,[]);
    for(const code of ['KEY_CONFLICT','AMOUNT','DATE','OVER_REFUND','REQUIRED','TYPE'])assert.ok(r.validation.diagnostics.some(d=>d.code===code),code);
  }
  assert.deepEqual(badCSV.validation.diagnostics.map(({sheet,...d})=>d),badXLSX.validation.diagnostics.map(({sheet,...d})=>d));
});
test('Excel 1900/1904 date systems use real format metadata; false leap day rejected', () => {
  const check=(serial,date1904,format='yyyy-mm-dd')=>{
    const grid=[cells(FIELDS),cells(fixture.data[3])];grid[1][4]={t:'n',v:serial,z:format};
    return normalized(grid,{book:{Workbook:{WBProps:{date1904}}}});
  };
  assert.equal(check(61,false).mapped.rows[0].date,'1900-03-01');
  assert.equal(check(0,true).mapped.rows[0].date,'1904-01-01');
  assert.equal(check(61,false).validation.ok,true);
  assert.equal(check(0,true).validation.ok,true);
  for(const r of [check(60,false),check(46032,false,'General'),check(61.5,false)]){
    assert.ok(r.mapped.errors.some(d=>d.field==='date'));
    assert.equal(r.validation.ok,false);
  }
});
test('required Excel formulas are rejected even with plausible cached values', () => {
  for(const field of FIELDS){
    const grid=canonicalGrid();const index=FIELDS.indexOf(field);
    grid[1][index]={...grid[1][index],f:'1+1'};
    const result=normalized(grid);
    assert.ok(result.mapped.errors.some(d=>d.field===field&&d.row===2&&/Формула/.test(d.message)));
  }
});
test('formula-only rows without cached values must not disappear as blank rows', () => {
  const grid=canonicalGrid();
  grid.push(FIELDS.map(()=>({t:'n',f:'1+1'})));
  const result=normalized(grid);
  assert.equal(result.mapped.rows.length,fixture.data.length+1);
  assert.equal(result.mapped.errors.filter(d=>d.row===17).length,FIELDS.length);
});
test('Excel error/boolean cells and numeric IDs require explicit correction', () => {
  for(const c of [{t:'e',v:7},{t:'b',v:true},{t:'n',v:1,z:'000'}]){
    const grid=canonicalGrid();grid[1][0]=c;
    const result=normalized(grid);
    assert.ok(result.mapped.errors.some(d=>d.field==='event_id'&&d.row===2));
    assert.equal(result.validation.ok,false);
  }
  const grid=canonicalGrid();grid[1][0]={t:'s',v:'001'};
  const result=mapRows(grid,canonicalMapping,{XLSX});
  assert.equal(result.rows[0].event_id,'001');
});
test('empty sheets and oversized workbook references are rejected before expansion', () => {
  for(const sheet of [{},{'!ref':'A1:BM2'},{'!ref':'A1:B10101'}]) assert.throws(()=>workbookGrid({Sheets:{X:sheet}},'X',XLSX));
  const sheet={'!ref':'B2:C2',B2:{t:'s',v:'x'},C2:{t:'s',v:'y'}};
  assert.deepEqual(workbookGrid({Sheets:{X:sheet}},'X',XLSX).map(row=>row.map(c=>c.v)),[['','',''],['','x','y']]);
});
test('direct engine numeric amounts are exact decimals or explicit errors', () => {
  for(const [amount,cents]of [[600.25,60025n],[0.1,10n]]){
    const r=validateRows([{...Object.fromEntries(FIELDS.map((f,i)=>[f,fixture.data[3][i]])),amount}],{mode:'event'});
    assert.equal(r.ok,true);assert.equal(r.events[0].amount,cents);
  }
  for(const amount of [NaN,Infinity,0.1+0.2,1e-7,true,null,{}]){
    const r=validateRows([{...Object.fromEntries(FIELDS.map((f,i)=>[f,fixture.data[3][i]])),amount}],{mode:'event'});
    assert.equal(r.ok,false);assert.deepEqual(r.events,[]);
  }
});
function parseExport(text){
  const grid=parseCSV(text,';').map(row=>row.map(c=>c.v));
  const columns=grid.shift();
  return grid.map(values=>Object.fromEntries(columns.map((field,i)=>[field,values[i]])));
}
test('CSV export matches every displayed summary and the exact dynamics', () => {
  const data=parseExport(csvReport(report,meta));
  const summary=Object.fromEntries(data.filter(r=>r.section==='summary').map(r=>[r.metric,r]));
  for(const [key,rule]of Object.entries(RULES)){
    const value=report[key];const expected=typeof value==='bigint'?money(value):value===null?'нет данных':key==='repeatBasisPoints'?(value/100).toFixed(2)+'%':String(value);
    assert.equal(summary[key].value,expected);
    assert.equal(summary[key].rule,rule);
    assert.equal(summary[key].currency,'RUB');
    assert.equal(summary[key].period,'2026-01-01..2026-02-28');
  }
  assert.equal(data.find(r=>r.section==='dynamics'&&r.metric==='net'&&r.period==='2026-02').value,'-500.00');
  const context=Object.fromEntries(data.filter(r=>r.section==='context').map(r=>[r.metric,r.value]));
  assert.equal(context.cutoff,'2026-02-28');assert.match(context.timezone,/Europe\/Moscow/);
  assert.equal(context.mode,'parts');assert.equal(context.decimal,'.');assert.equal(context.headerRow,'1');
  assert.equal(context.sheet,'CSV');assert.equal(context.events,'13');assert.equal(context.duplicates,'1');
  assert.match(context.dataset,/Вымышленные/);assert.match(context.scope,/Полный/);
  assert.deepEqual(JSON.parse(context.mapping),canonicalMapping);
  assert.equal(data.filter(r=>r.section==='diagnostic').length,1);
});
test('JSON export preserves exact money, report context, formula rules and full accepted history', () => {
  const parsed=JSON.parse(jsonReport(report,meta,csvDataset.validation.events));
  assert.equal(parsed.schema,'orderlens-report-v1');assert.deepEqual(parsed.rules,RULES);
  assert.deepEqual(parsed.context,meta);
  for(const key of ['gross','refunds','net','aov'])assert.equal(parsed.report[key],money(report[key]));
  for(const key of ['orders','eligibleOrders','customers','repeatCustomers','repeatBasisPoints','currency','from','to','cutoff','timezone'])assert.equal(parsed.report[key],report[key]);
  assert.equal(parsed.events.length,13);
  assert.equal(parsed.events.find(e=>e.event_id==='S1').amount,'1000.00');
  assert.equal(parsed.events.find(e=>e.event_id==='S6').currency,'USD');
  assert.equal(parsed.report.dynamics[1].net,'-500.00');
});
test('exports preserve missing averages/ratios and negative refund-only revenue', () => {
  const r=compute(csvDataset.validation.events,{from:'2026-02-06',to:'2026-02-06',currency:'RUB'});
  const csv=parseExport(csvReport(r,meta));
  assert.equal(csv.find(d=>d.section==='summary'&&d.metric==='aov').value,'нет данных');
  assert.equal(csv.find(d=>d.section==='summary'&&d.metric==='repeatBasisPoints').value,'нет данных');
  assert.equal(csv.find(d=>d.section==='summary'&&d.metric==='net').value,'-1000.00');
  const parsed=JSON.parse(jsonReport(r,meta,csvDataset.validation.events));
  assert.equal(parsed.report.aov,null);assert.equal(parsed.report.repeatBasisPoints,null);
});
test('CSV export neutralizes text formulas while preserving negative numeric metrics', () => {
  for(const name of ['=1+1','+SUM(A1:A2)','-SUM(A1:A2)','@SUM(A1:A2)','\t=1+1','\n=1+1','  =1+1']){
    const csv=parseExport(csvReport(report,{...meta,name}));
    assert.equal(csv.find(r=>r.section==='context'&&r.metric==='source').value,"'"+name);
    assert.equal(csv.find(r=>r.section==='dynamics'&&r.metric==='net'&&r.period==='2026-02').value,'-500.00');
  }
});
