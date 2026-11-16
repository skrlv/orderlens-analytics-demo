import {chromium} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {parseCSV} from '../src/import.mjs';
const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const base=process.env.BASE_URL||'http://127.0.0.1:4188/';
const out=path.resolve(process.env.EVIDENCE_DIR||path.join(project,'test-results/xlsx-safety'));
await fs.mkdir(out,{recursive:true});
const browser=await chromium.launch({headless:true});
const checks=[],errors=[],requests=[],workers=[];
const pass=name=>{checks.push(name);console.log('PASS',name);};
const clean=s=>s.replace(/[\s\u202f\u00a0]/g,'');
async function workersClosed(page){for(let i=0;i<30&&page.workers().length;i++)await page.waitForTimeout(100);assert.equal(page.workers().length,0);}
async function exported(page,label){const waiting=page.waitForEvent('download');await page.locator('#exportCSV').click();const d=await waiting;assert.equal(await d.failure(),null);const p=path.join(out,label+'.csv');await d.saveAs(p);const rows=parseCSV(await fs.readFile(p,'utf8'),';').map(r=>r.map(c=>c.v));assert.ok(rows.some(r=>r[0]==='summary'&&r[1]==='net'&&r[4]==='-500.00'));return rows;}
async function feb(page){await page.locator('#from').fill('2026-02-01');await page.locator('#to').fill('2026-02-28');await page.locator('#to').dispatchEvent('change');assert.equal(clean(await page.locator('#net').innerText()),'-500,00₽');}
async function load(page){await page.goto(base,{waitUntil:'networkidle'});await page.waitForFunction(()=>document.getElementById('net').textContent!=='—');await feb(page);await page.locator('#uploadButton').click();}
async function file(page,name,buffer){await page.locator('#file').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer});}
async function rejected(page){await page.waitForFunction(()=>document.getElementById('notice').textContent.includes('Новый файл не принят'));assert.equal(await page.locator('#confirmImport').isDisabled(),true);assert.equal(await page.locator('#exportCSV').isEnabled(),true);assert.equal(clean(await page.locator('#net').innerText()),'-500,00₽');assert.match(await page.locator('#diagnostics').innerText(),/файл/);}
async function healthy(page,bytes,name='healthy.xlsx'){await file(page,name,bytes);await page.waitForFunction(()=>!document.getElementById('confirmImport').disabled);await page.locator('#mode').selectOption('parts');await page.locator('#confirmImport').click();await feb(page);assert.equal(await page.locator('#sourceName').innerText(),name);}
try{
 const context=await browser.newContext({viewport:{width:1280,height:900},locale:'ru-RU',acceptDownloads:true});const page=await context.newPage();
 page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>requests.push({method:r.method(),url:r.url()}));page.on('worker',w=>{const entry={url:w.url(),closed:false};workers.push(entry);w.on('close',()=>entry.closed=true);});
 await load(page);const xlsx=await fs.readFile(path.join(project,'public/samples/orders-parts.xlsx'));
 const hundred=await fs.readFile(path.join(project,'tests/fixtures/truncated-100.xlsx'));assert.equal(hundred.length,100);
 await file(page,'review-truncated-100.xlsx',hundred);await rejected(page);assert.match(await page.locator('#diagnostics').innerText(),/XLSX|ZIP|структур|поврежд|обрезан/i);await exported(page,'truncated-100-preserved');await page.screenshot({path:path.join(out,'corrupt-diagnostic.png'),fullPage:true});pass('exact 100-byte review file rejects promptly and previous CSV export remains usable');
 for(const [name,bytes]of [['missing-end.xlsx',xlsx.subarray(0,xlsx.length-22)],['mid-file.xlsx',xlsx.subarray(0,Math.floor(xlsx.length/2))],['invalid-zip.xlsx',Buffer.from([0x50,0x4b,3,4,...new Array(80).fill(0)])]]){await file(page,name,bytes);await rejected(page);}
 await exported(page,'other-corruption-preserved');pass('other truncated and invalid ZIP structures reject without replacing accepted data');
 await healthy(page,xlsx);await exported(page,'healthy-worker-import');assert.equal(await page.locator('#cancelRead').isHidden(),true);pass('valid XLSX imports in same-origin worker under unchanged script restrictions');
 const workerURL=new URL('./src/xlsx-worker.js',base).href;
 await context.route(workerURL,route=>route.fulfill({status:200,contentType:'text/javascript',body:'self.onmessage = () => { while (true) {} };'}));
 await file(page,'slow-timeout.xlsx',xlsx);await page.locator('#cancelRead').waitFor({state:'visible'});
 const ticks=await page.evaluate(async()=>{let ticks=0;await new Promise(resolve=>{const id=setInterval(()=>{if(++ticks===5){clearInterval(id);resolve();}},30);});return ticks;});assert.equal(ticks,5);await exported(page,'during-slow-read');assert.equal(await page.locator('#sourceName').innerText(),'healthy.xlsx');pass('while worker is busy, main-thread heartbeat, report controls and real CSV download stay responsive');
 await rejected(page);assert.match(await page.locator('#diagnostics').innerText(),/5 секунд/);await page.waitForFunction(()=>document.getElementById('cancelRead').hidden);await exported(page,'after-timeout');await workersClosed(page);pass('five-second watchdog terminates non-returning reader and keeps previous report/export');
 await file(page,'cancel-me.xlsx',xlsx);await page.locator('#cancelRead').waitFor({state:'visible'});await page.locator('#cancelRead').click();assert.match(await page.locator('#notice').innerText(),/Чтение отменено/);await workersClosed(page);await exported(page,'after-manual-cancel');pass('manual cancellation terminates busy reader without losing accepted data');
 await file(page,'superseded.xlsx',xlsx);await page.locator('#cancelRead').waitFor({state:'visible'});
 await page.locator('#file').setInputFiles({name:'latest.csv',mimeType:'text/csv',buffer:await fs.readFile(path.join(project,'public/samples/orders-parts.csv'))});await page.waitForFunction(()=>!document.getElementById('confirmImport').disabled);await page.locator('#mode').selectOption('parts');await page.locator('#confirmImport').click();await feb(page);await page.waitForTimeout(700);assert.equal(await page.locator('#sourceName').innerText(),'latest.csv');await workersClosed(page);await exported(page,'latest-import-preserved');pass('new CSV cancels previous read; stale worker cannot overwrite the newer import');
 await context.unroute(workerURL);await healthy(page,xlsx,'recovered.xlsx');await exported(page,'recovered-valid-import');pass('valid XLSX still works after timeout, cancellation and replacement');
 await page.locator('#closeImport').click();await page.screenshot({path:path.join(out,'desktop.png'),fullPage:true});await page.screenshot({path:path.join(out,'preview.png')});await context.close();
 const mobile=await browser.newContext({viewport:{width:390,height:844},isMobile:true,hasTouch:true,acceptDownloads:true});const m=await mobile.newPage();m.on('pageerror',e=>errors.push(e.message));await load(m);await file(m,'mobile-truncated.xlsx',hundred);await rejected(m);await exported(m,'mobile-corrupt-preserved');await healthy(m,xlsx,'mobile-valid.xlsx');await m.locator('#closeImport').click();await exported(m,'mobile-valid-export');assert.ok(await m.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));await m.screenshot({path:path.join(out,'mobile.png'),fullPage:true});await mobile.close();pass('390px mobile corruption/recovery path preserves export and has no page overflow');
 assert.deepEqual(errors,[]);assert.ok(workers.every(w=>w.closed));assert.ok(requests.every(r=>r.method==='GET'&&new URL(r.url).origin===new URL(base).origin));pass('all workers closed; no JavaScript errors, external requests or table upload');
 await fs.writeFile(path.join(out,'xlsx-browser-result.json'),JSON.stringify({base,checkedAt:new Date().toISOString(),checks,passed:checks.length,errors,workers,requests},null,2));
}finally{await browser.close();}
