import test from 'node:test';
import assert from 'node:assert/strict';
import { validateRows, compute, money, validDate, FIELDS } from '../src/engine.mjs';

const fixture = [
  ['S1','L1','O1','C1','2026-01-10','sale','600.00','RUB'],
  ['S1','L2','O1','C1','2026-01-10','sale','400.00','RUB'],
  ['S1','L1','O1','C1','2026-01-10','sale','600.00','RUB'],
  ['S2','L1','O2','C2','2026-01-15','sale','700.00','RUB'],
  ['S3','L1','O3','C1','2026-01-20','sale','500.00','RUB'],
  ['R1','L1','O2','C2','2026-01-22','refund','200.00','RUB'],
  ['X1','L1','O4','C3','2026-01-25','cancel','0.00','RUB'],
  ['S4','L1','O5','C2','2026-02-05','sale','300.00','RUB'],
  ['R2','L1','O1','C1','2026-02-06','refund','1000.00','RUB'],
  ['S5','L1','O6','C2','2026-02-10','sale','200.00','RUB'],
  ['R3','L1','O5','C2','2026-02-11','refund','100.00','RUB'],
  ['S6','L1','O7','C4','2026-02-12','sale','50.00','USD'],
  ['S7','L1','O8','C5','2026-02-20','sale','400.00','RUB'],
  ['R4','L1','O8','C5','2026-02-21','refund','400.00','RUB'],
  ['S8','L1','O9','C6','2026-02-22','sale','100.00','RUB'],
].map(row => Object.fromEntries(FIELDS.map((field, index) => [field, row[index]])));
const opts = { mode: 'parts', decimal: '.', sheet: 'Контроль', headerRow: 1 };
const clone = () => structuredClone(fixture);
const accepted = validateRows(fixture, opts);
assert.equal(accepted.ok, true);

test('manual fixture deduplicates parts before event aggregation', () => {
  assert.deepEqual(accepted.stats, {inputRows:15,uniqueRows:14,duplicates:1,events:13,sales:8,refunds:4,cancellations:1});
  assert.equal(accepted.events.find(event => event.event_id === 'S1').amount, 100000n);
  assert.equal(accepted.diagnostics.length, 1);
  assert.deepEqual(accepted.diagnostics[0], { severity:'warning',code:'DUPLICATE',sheet:'Контроль',row:4,field:'event_id + line_id',message:'Точный дубль строки 2: учтён один раз.' });
});

const cases = [
  ['January RUB','2026-01-01','2026-01-31','RUB',[220000n,20000n,200000n,3,73333n,3,2,1,5000]],
  ['February RUB','2026-02-01','2026-02-28','RUB',[100000n,150000n,-50000n,4,25000n,3,2,1,5000]],
  ['both months RUB','2026-01-01','2026-02-28','RUB',[320000n,170000n,150000n,7,45714n,5,3,1,3333]],
  ['February USD','2026-02-01','2026-02-28','USD',[5000n,0n,5000n,1,5000n,1,1,0,0]],
];
for (const [name,from,to,currency,expected] of cases) test(`independent manual totals: ${name}`, () => {
  const r = compute(accepted.events, {from,to,currency});
  assert.deepEqual([r.gross,r.refunds,r.net,r.orders,r.aov,r.eligibleOrders,r.customers,r.repeatCustomers,r.repeatBasisPoints], expected);
  assert.equal(r.cutoff, to);
  assert.equal(r.timezone, 'Europe/Moscow');
  assert.equal(r.dynamics.reduce((sum, bucket) => sum + bucket.net, 0n), r.net);
});

test('refund-only period has negative revenue and no average or repeat denominator', () => {
  const r = compute(accepted.events, {from:'2026-02-06',to:'2026-02-06',currency:'RUB',grain:'day'});
  assert.deepEqual([r.gross,r.refunds,r.net,r.orders,r.aov,r.customers,r.repeatBasisPoints], [0n,100000n,-100000n,0,null,0,null]);
  assert.deepEqual(r.dynamics, [{period:'2026-02-06',gross:0n,refunds:100000n,net:-100000n,orders:0}]);
});
test('empty interval in accepted history is a real zero, never an unknown amount', () => {
  const r = compute(accepted.events, {from:'2026-03-01',to:'2026-03-31',currency:'RUB'});
  assert.deepEqual([r.gross,r.refunds,r.net,r.orders,r.aov,r.repeatBasisPoints,r.dynamics], [0n,0n,0n,0,null,null,[]]);
});
test('cutoff excludes later refunds and preserves historical repeat rates', () => {
  const r = compute(accepted.events, {from:'2026-01-01',to:'2026-01-31',currency:'RUB'});
  assert.equal(r.eligibleOrders, 3);
  assert.equal(r.repeatCustomers, 1);
});
test('full refund preserves paid order count and paid average but excludes retained purchase', () => {
  const r = compute(accepted.events, {from:'2026-02-20',to:'2026-02-21',currency:'RUB'});
  assert.deepEqual([r.orders,r.aov,r.net,r.eligibleOrders,r.customers,r.repeatBasisPoints], [1,40000n,0n,0,0,null]);
});

function reject(mutator, code, options=opts) {
  const rows = clone(); mutator(rows);
  const r = validateRows(rows, options);
  assert.equal(r.ok, false);
  assert.deepEqual(r.events, []);
  assert.ok(r.diagnostics.some(d => d.code === code), JSON.stringify(r.diagnostics));
  assert.ok(r.diagnostics.filter(d => d.severity === 'error').every(d => d.sheet === (options.sheet ?? 'CSV') && Number.isInteger(d.row) && d.message.length > 0));
}
const invalidCases = [
  ['conflicting duplicate', rows => rows[2].amount='601', 'KEY_CONFLICT'],
  ['missing amount', rows => rows[0].amount='', 'AMOUNT'],
  ['missing customer', rows => rows[0].customer_id='', 'REQUIRED'],
  ['missing line ID', rows => rows[0].line_id='', 'REQUIRED'],
  ['invalid date', rows => rows[0].date='2026-02-30', 'DATE'],
  ['ambiguous date', rows => rows[0].date='01/02/2026', 'DATE'],
  ['unsupported timestamp', rows => rows[0].date='2026-01-10T00:00:00+03:00', 'DATE'],
  ['refund over paid total', rows => rows[8].amount='1001', 'OVER_REFUND'],
  ['orphan refund', rows => rows.splice(0,3), 'ORPHAN_REFUND'],
  ['refund customer mismatch', rows => rows[8].customer_id='OTHER', 'ORDER_CUSTOMER'],
  ['refund currency mismatch', rows => rows[8].currency='USD', 'ORDER_CURRENCY'],
  ['refund before payment', rows => rows[8].date='2026-01-09', 'REFUND_BEFORE_SALE'],
  ['negative refund not silently flipped', rows => rows[8].amount='-1000', 'AMOUNT'],
  ['fractional cent not rounded', rows => rows[8].amount='1000.001', 'AMOUNT'],
  ['unsupported grouping separator', rows => rows[8].amount='1 000.00', 'AMOUNT'],
  ['unsupported type', rows => rows[8].type='pending', 'TYPE'],
  ['unknown currency not guessed', rows => rows[8].currency='XXX', 'CURRENCY'],
  ['missing currency not guessed', rows => rows[8].currency='', 'CURRENCY'],
  ['nonzero unpaid cancellation', rows => rows[6].amount='1', 'CANCEL_AMOUNT'],
  ['zero sale', rows => rows[3].amount='0', 'AMOUNT_POSITIVE'],
  ['zero refund', rows => rows[8].amount='0', 'AMOUNT_POSITIVE'],
  ['event parts metadata disagree', rows => rows[1].date='2026-01-11', 'EVENT_CONFLICT'],
  ['two different payment events for one order', rows => rows.push({...rows[3],event_id:'S-EXTRA'}), 'MULTIPLE_SALES'],
  ['paid and cancelled order', rows => rows.push({...rows[3],event_id:'X-EXTRA',type:'cancel',amount:'0'}), 'SALE_CANCEL'],
];
for (const [name,mutator,code] of invalidCases) test(`rejects atomically: ${name}`, () => reject(mutator,code));

test('event mode treats a complete event as one order and ignores unused line IDs', () => {
  const rows = [fixture[3],{...fixture[3],line_id:'OTHER'}];
  const r = validateRows(rows,{mode:'event'});
  assert.equal(r.ok,true);
  assert.equal(r.events.length,1);
  assert.equal(r.stats.duplicates,1);
  assert.equal(r.events[0].amount,70000n);
});
test('parts fixture cannot silently be interpreted as complete events', () => reject(()=>{},'KEY_CONFLICT',{...opts,mode:'event'}));
test('comma decimal is explicit and strict', () => {
  const r = validateRows([{...fixture[3],amount:'700,01'}],{mode:'event',decimal:','});
  assert.equal(r.ok,true);
  assert.equal(r.events[0].amount,70001n);
  assert.equal(validateRows([{...fixture[3],amount:'700.01'}],{mode:'event',decimal:','}).ok,false);
});
test('string IDs keep leading zeroes; Map keys safely handle prototype-like identifiers', () => {
  const rows = [{...fixture[3],event_id:'__proto__',order_id:'001',customer_id:'constructor'}, {...fixture[3],event_id:'constructor',order_id:'1',customer_id:'constructor'}];
  const r=validateRows(rows,{mode:'event'});
  assert.equal(r.ok,true);
  assert.deepEqual(new Set(r.events.map(e=>e.order_id)),new Set(['001','1']));
  assert.equal(compute(r.events,{from:'2026-01-01',to:'2026-01-31',currency:'RUB'}).repeatCustomers,1);
});
test('composite keys cannot collide on delimiters inside identifiers', () => {
  const rows = [{...fixture[3],event_id:'a|b',line_id:'c'}, {...fixture[3],event_id:'a',line_id:'b|c',order_id:'OTHER'}];
  const r=validateRows(rows,opts);
  assert.equal(r.ok,true);
  assert.equal(r.events.length,2);
});
test('exact decimal arithmetic beyond Number safe integer and half-up average rounding', () => {
  const r=validateRows([{...fixture[3],amount:'9007199254740993.01'},{...fixture[3],event_id:'BIG2',order_id:'BIG2',amount:'0.02'}],{mode:'event'});
  assert.equal(r.ok,true);
  const c=compute(r.events,{from:'2026-01-01',to:'2026-01-31',currency:'RUB'});
  assert.equal(c.gross,900719925474099303n);
  assert.equal(c.aov,450359962737049652n);
  assert.equal(money(c.gross),'9007199254740993.03');
  const cents=validateRows([{...fixture[3],amount:'0.1'},{...fixture[3],event_id:'CENTS2',order_id:'CENTS2',amount:'0.2'}],{mode:'event'});
  assert.equal(compute(cents.events,{from:'2026-01-01',to:'2026-01-31',currency:'RUB'}).gross,30n);
});
test('unsafe numeric amounts are rejected while large decimal strings remain exact', () => {
  for(const amount of [90071992547410,90071992547408.01,1000000000000.01]){
    const numeric=validateRows([{...fixture[3],amount}],{mode:'event'});
    assert.equal(numeric.ok,false);assert.deepEqual(numeric.events,[]);
    assert.ok(numeric.diagnostics.some(d=>d.code==='AMOUNT_UNSAFE_NUMBER'&&d.field==='amount'&&d.message.includes('1 000 000 000 000')));
  }
  const boundary=validateRows([{...fixture[3],amount:1000000000000}],{mode:'event'});
  assert.equal(boundary.ok,true);assert.equal(boundary.events[0].amount,100000000000000n);
  const textual=validateRows([{...fixture[3],amount:'90071992547410.01'}],{mode:'event'});
  assert.equal(textual.ok,true);assert.equal(textual.events[0].amount,9007199254741001n);
  for(const [amount,cents]of [['90071992547408.01',9007199254740801n],['1000000000000.01',100000000000001n]]){
    const exact=validateRows([{...fixture[3],amount}],{mode:'event'});
    assert.equal(exact.ok,true);assert.equal(exact.events[0].amount,cents);
  }
});
test('repeat import is replacement-ready and deterministic without mutation', () => {
  const again=validateRows(fixture,opts);
  assert.deepEqual(again,accepted);
  assert.equal(fixture[0].amount,'600.00');
  assert.equal(again.events.length,13);
});
test('money handles zero and negative minor units exactly', () => {
  assert.equal(money(0n),'0.00');
  assert.equal(money(-1n),'-0.01');
  assert.equal(money(-50000n),'-500.00');
  assert.throws(()=>money(1),TypeError);
});
test('strict calendar validation and explicit configuration failures', () => {
  assert.equal(validDate('2024-02-29'),true);
  for (const value of ['2026-02-29','1900-02-29','0000-01-01','2026-13-01','2026-00-01','2026-01-00','2026-1-01',null]) assert.equal(validDate(value),false);
  assert.equal(validateRows([],opts).ok,false);
  assert.equal(validateRows(fixture,{}).ok,false);
  assert.equal(validateRows(fixture,{...opts,decimal:';'}).ok,false);
  assert.throws(()=>compute(accepted.events,{from:'2026-03-01',to:'2026-02-01',currency:'RUB'}),RangeError);
  assert.throws(()=>compute(accepted.events,{from:'2026-01-01',to:'2026-02-01',currency:'XXX'}),RangeError);
  assert.throws(()=>compute(accepted.events,{from:'2026-01-01',to:'2026-02-01',currency:'RUB',grain:'year'}),RangeError);
});
