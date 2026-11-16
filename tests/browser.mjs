import {chromium} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import vm from 'node:vm';
import {parseCSV} from '../src/import.mjs';
const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const base=process.env.BASE_URL||'http://127.0.0.1:4188/';
const out=path.resolve(process.env.EVIDENCE_DIR||path.join(project,'test-results'));
await fs.mkdir(out,{recursive:true});
const browser=await chromium.launch({headless:true});
const checks=[],errors=[],requests=[],httpErrors=[];
const check=(name)=>{checks.push(name);console.log('PASS',name);};
const norm=s=>s.replace(/[\s\u202f\u00a0]/g,'');
async function ready(page){await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>document.getElementById('net').textContent!=='—');assert.match(await page.locator('#notice').innerText(),/Принят полный набор/);}
async function period(page,from,to,currency='RUB'){await page.locator('#from').fill(from);await page.locator('#to').fill(to);await page.locator('#currency').selectOption(currency);await page.locator('#to').dispatchEvent('change');}
async function metrics(page,{net,aov,repeat,orders}){assert.equal(norm(await page.locator('#net').innerText()),net);assert.equal(norm(await page.locator('#aov').innerText()),aov);assert.equal(norm(await page.locator('#repeat').innerText()),repeat);assert.match(await page.locator('#orderCount').innerText(),new RegExp(': '+orders+'$'));}
async function download(page,selector,name){const wait=page.waitForEvent('download');await page.locator(selector).click();const d=await wait;assert.equal(await d.failure(),null);const p=path.join(out,name||d.suggestedFilename());await d.saveAs(p);assert.ok((await fs.stat(p)).size>0);return p;}
async function upload(page,file,{mode='parts',delimiter=',',decimal='.'}={}){if(!await page.locator('#importPanel').isVisible())await page.locator('#uploadButton').click();await page.locator('#delimiter').selectOption(delimiter);await page.locator('#decimal').selectOption(decimal);await page.locator('#file').setInputFiles(file);await page.locator('#confirmImport').waitFor({state:'visible'});await page.waitForFunction(()=>!document.getElementById('confirmImport').disabled);await page.locator('#mode').selectOption(mode);await page.locator('#confirmImport').click();}
try{
  const context=await browser.newContext({viewport:{width:1440,height:1080},locale:'ru-RU',acceptDownloads:true});
  const page=await context.newPage();page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>requests.push({method:r.method(),url:r.url()}));page.on('response',r=>{if(r.status()>=400)httpErrors.push({status:r.status(),url:r.url()});});
  await ready(page);
  await metrics(page,{net:'1500,00₽',aov:'457,14₽',repeat:'33,33%',orders:7});check('initial demo matches combined manual totals');
  await period(page,'2026-01-01','2026-01-31');await metrics(page,{net:'2000,00₽',aov:'733,33₽',repeat:'50,00%',orders:3});check('January cutoff ignores future full refund');
  await period(page,'2026-02-01','2026-02-28');await metrics(page,{net:'-500,00₽',aov:'250,00₽',repeat:'50,00%',orders:4});
  const csvPath=await download(page,'#exportCSV','feb-rub-report.csv');
  const csv=parseCSV(await fs.readFile(csvPath,'utf8'),';').map(r=>r.map(c=>c.v));
  const summary=Object.fromEntries(csv.slice(1).filter(r=>r[0]==='summary').map(r=>[r[1],r[4]]));
  assert.deepEqual(summary,{gross:'1000.00',refunds:'1500.00',net:'-500.00',orders:'4',aov:'250.00',eligibleOrders:'3',customers:'2',repeatCustomers:'1',repeatBasisPoints:'50.00%'});
  assert.ok(csv.some(r=>r[1]==='cutoff'&&r[4]==='2026-02-28'));assert.ok(csv.some(r=>r[1]==='timezone'&&r[4].includes('Europe/Moscow')));
  const jsonPath=await download(page,'#exportJSON','feb-rub-report.json');const json=JSON.parse(await fs.readFile(jsonPath,'utf8'));assert.equal(json.report.net,'-500.00');assert.equal(json.events.length,13);assert.equal(json.context.stats.duplicates,1);assert.ok(json.rules.aov.includes('G / O'));check('actual CSV and JSON downloads agree with visible February totals and context');
  await period(page,'2026-02-01','2026-02-28','USD');await metrics(page,{net:'50,00$',aov:'50,00$',repeat:'0,00%',orders:1});check('USD isolated from RUB');
  await period(page,'2026-02-06','2026-02-06','RUB');await metrics(page,{net:'-1000,00₽',aov:'Нетзаказов',repeat:'Нетданных',orders:0});check('refund-only period keeps negative revenue and missing average');
  await period(page,'2026-03-01','2026-03-31');await metrics(page,{net:'0,00₽',aov:'Нетзаказов',repeat:'Нетданных',orders:0});check('empty period has zero events and no false average');
  await period(page,'2026-02-28','2026-01-01');assert.equal(await page.locator('#exportCSV').isDisabled(),true);assert.match(await page.locator('#filterError').innerText(),/предыдущий отчёт/);check('invalid period disables export and labels previous result');
  await period(page,'2026-01-01','2026-02-28');
  const downloadedCSV=await download(page,'a[href="./samples/orders-parts.csv"]','downloaded-sample.csv');
  await upload(page,downloadedCSV);await period(page,'2026-01-01','2026-02-28');await metrics(page,{net:'1500,00₽',aov:'457,14₽',repeat:'33,33%',orders:7});
  const csvJsonPath=await download(page,'#exportJSON','csv-import-report.json');const csvJson=JSON.parse(await fs.readFile(csvJsonPath,'utf8'));check('downloaded sample CSV imports through explicit mapping and parts mode');
  const downloadedXLSX=await download(page,'a[href="./samples/orders-parts.xlsx"]','downloaded-sample.xlsx');await upload(page,downloadedXLSX);await period(page,'2026-01-01','2026-02-28');await metrics(page,{net:'1500,00₽',aov:'457,14₽',repeat:'33,33%',orders:7});
  const xlsxJsonPath=await download(page,'#exportJSON','xlsx-import-report.json');const xlsxJson=JSON.parse(await fs.readFile(xlsxJsonPath,'utf8'));assert.deepEqual(xlsxJson.report,csvJson.report);assert.deepEqual(xlsxJson.events,csvJson.events);check('downloaded XLSX normalizes identically to CSV and report');
  const sandbox={};vm.runInNewContext(await fs.readFile(path.join(project,'public/vendor/xlsx.full.min.js'),'utf8'),sandbox);const XLSX=sandbox.XLSX;const source=XLSX.read(await fs.readFile(downloadedXLSX),{type:'buffer'});const rows=XLSX.utils.sheet_to_json(source.Sheets[source.SheetNames[0]],{header:1,raw:true});const multi=XLSX.utils.book_new();XLSX.utils.book_append_sheet(multi,XLSX.utils.aoa_to_sheet([['Инструкция'],['Выберите лист История и строку 3']]),'Описание');XLSX.utils.book_append_sheet(multi,XLSX.utils.aoa_to_sheet([['Вымышленная выгрузка'],[],...rows]),'История');const multiFile=path.join(out,'multi-header.xlsx');await fs.writeFile(multiFile,XLSX.write(multi,{type:'buffer',bookType:'xlsx'}));await upload(page,multiFile);assert.match(await page.locator('#notice').innerText(),/не принят/);await page.locator('#sheet').selectOption('История');await page.locator('#headerRow').fill('3');await page.locator('#headerRow').dispatchEvent('change');await page.locator('#mode').selectOption('parts');await page.locator('#confirmImport').click();await period(page,'2026-01-01','2026-02-28');await metrics(page,{net:'1500,00₽',aov:'457,14₽',repeat:'33,33%',orders:7});check('real workbook sheet selection and third-row headers');
  await upload(page,path.join(project,'public/samples/orders-parts-ru.csv'),{delimiter:';',decimal:','});await period(page,'2026-01-01','2026-02-28');await metrics(page,{net:'1500,00₽',aov:'457,14₽',repeat:'33,33%',orders:7});check('UTF-8 BOM / semicolon / comma decimal imported explicitly');
  await upload(page,path.join(project,'public/samples/orders-errors.csv'),{delimiter:','});assert.match(await page.locator('#notice').innerText(),/отклонён/);assert.match(await page.locator('#diagnostics').innerText(),/amount|date|customer_id/);await metrics(page,{net:'1500,00₽',aov:'457,14₽',repeat:'33,33%',orders:7});assert.match(await page.locator('#sourceName').innerText(),/orders-parts-ru/);check('bad import shows row-level diagnostics and preserves accepted data');
  await page.locator('#file').setInputFiles({name:'corrupt.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from('invalid workbook')});await page.waitForTimeout(200);assert.match(await page.locator('#notice').innerText(),/не принят/);assert.equal(await page.locator('#confirmImport').isDisabled(),true);await metrics(page,{net:'1500,00₽',aov:'457,14₽',repeat:'33,33%',orders:7});check('corrupt workbook preserves data and disables acceptance');
  await page.locator('#file').setInputFiles({name:'custom.csv',mimeType:'text/csv',buffer:Buffer.from('eid,oid,cid,day,kind,paid,ccy\n001,0001,01,2026-05-01,sale,12.30,EUR\n')});await page.locator('#delimiter').selectOption(',');await page.locator('#decimal').selectOption('.');await page.locator('#mode').selectOption('event');
  for(const [f,i]of Object.entries({event_id:'0',order_id:'1',customer_id:'2',date:'3',type:'4',amount:'5',currency:'6'}))await page.locator('#map-'+f).selectOption(i);
  await page.locator('#confirmImport').click();assert.match(await page.locator('#notice').innerText(),/Принят полный набор/);await metrics(page,{net:'12,30€',aov:'12,30€',repeat:'0,00%',orders:1});assert.match(await page.locator('#events').innerText(),/0001/);check('custom column mapping, event mode and leading-zero identifiers');
  await page.locator('#loadDemo').click();await page.waitForFunction(()=>document.getElementById('sourceName').textContent==='orders-parts.csv');await page.locator('#closeImport').click();
  await page.screenshot({path:path.join(out,'desktop.png'),fullPage:true});await page.screenshot({path:path.join(out,'preview.png')});check('desktop screenshot captured');await context.close();
  const mobile=await browser.newContext({viewport:{width:390,height:844},isMobile:true,hasTouch:true,locale:'ru-RU',acceptDownloads:true});const m=await mobile.newPage();m.on('pageerror',e=>errors.push(e.message));m.on('request',r=>requests.push({method:r.method(),url:r.url()}));await ready(m);assert.ok(await m.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));await upload(m,downloadedXLSX);await period(m,'2026-02-01','2026-02-28');await metrics(m,{net:'-500,00₽',aov:'250,00₽',repeat:'50,00%',orders:4});await m.locator('#closeImport').click();const mobileCSV=await download(m,'#exportCSV','mobile-feb-report.csv');
  const mrows=parseCSV(await fs.readFile(mobileCSV,'utf8'),';').map(r=>r.map(c=>c.v));assert.ok(mrows.some(r=>r[0]==='summary'&&r[1]==='net'&&r[4]==='-500.00'));assert.ok(await m.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));await m.screenshot({path:path.join(out,'mobile.png'),fullPage:true});check('390px touch path: XLSX upload, mapping, February metrics, real export, no page overflow');await mobile.close();
  assert.deepEqual(errors,[]);assert.deepEqual(httpErrors,[]);assert.ok(requests.every(r=>r.method==='GET'&&new URL(r.url).origin===new URL(base).origin));check('no JavaScript/HTTP errors; only same-origin GET requests, no table upload');
  await fs.writeFile(path.join(out,'browser-result.json'),JSON.stringify({base,checkedAt:new Date().toISOString(),checks,errors,httpErrors,requests,passed:checks.length},null,2));
}finally{await browser.close();}
