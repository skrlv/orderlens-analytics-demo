import {FIELDS,validateRows,compute,money} from './engine.mjs';
import {parseCSV,workbookGrid,headersAt,mapRows} from './import.mjs';
import {readWorkbookAsync} from './xlsx-reader.mjs';
import {csvReport,jsonReport} from './export.mjs';
const $=id=>document.getElementById(id);
const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const cash=(v,currency)=>money(v).replace('.',',').replace(/\B(?=(\d{3})+(?!\d))/g,'\u202f')+' '+({RUB:'₽',USD:'$',EUR:'€'}[currency]||currency);
let accepted=null,pending=null,currentReport=null,page=0,fileToken=0,readController=null;
function reading(active){$('readStatus').hidden=!active;$('cancelRead').hidden=!active;$('importPanel').setAttribute('aria-busy',String(active));}
function clearPending(){pending=null;$('confirmImport').disabled=true;$('mapping').innerHTML='';$('preview').innerHTML='';}
function stopReading(){readController?.abort();readController=null;reading(false);}
$('cancelRead').onclick=()=>{++fileToken;stopReading();clearPending();diagnostics([]);notice('Чтение отменено. Прежние данные и экспорт доступны.');};
function notice(message,type=''){ $('notice').textContent=message;$('notice').className='notice '+type; }
function showImport(){ $('importPanel').hidden=false;$('importPanel').scrollIntoView({behavior:'smooth',block:'start'}); }
for(const id of ['openImport','uploadButton','showImport'])$(id).addEventListener('click',showImport);
$('closeImport').onclick=()=>{$('importPanel').hidden=true;};
function mapping(){return Object.fromEntries(FIELDS.map(f=>[f,$('map-'+f)?.value??'']));}
function table(headers,rows){return `<table><thead><tr>${headers.map(h=>`<th scope="col">${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;}
function diagnostics(items){
  $('diagnostics').innerHTML=items.length?`<h3 class="diagtitle">Диагностика нового файла · ${items.length}</h3><div class="tablewrap">${table(['Уровень','Лист','Запись / строка','Поле','Причина'],items.slice(0,200).map(d=>[d.severity==='error'?'Ошибка':'Предупреждение',d.sheet,d.row,d.field,d.message]))}</div>${items.length>200?'<p>Показаны первые 200 сообщений. Исправьте файл и проверьте снова.</p>':''}`:'';
}
function prepare(){
  $('confirmImport').disabled=true;
  if(!pending)return;
  try{
    const headerRow=Number($('headerRow').value);
    if(pending.book)pending.grid=workbookGrid(pending.book,$('sheet').value,globalThis.XLSX);
    else pending.grid=parseCSV(pending.text,$('delimiter').value==='tab'?'\t':$('delimiter').value);
    const headers=headersAt(pending.grid,headerRow);
    if(!headers.length)throw Error('Нет заголовков.');
    pending.headers=headers;
    $('mapping').innerHTML='<h3 class="preview-title">Сопоставьте колонки</h3><div class="mappinggrid">'+FIELDS.map(f=>`<label>${f}${f==='line_id'?' · части':' *'}<select id="map-${f}"><option value="">Не выбрано</option>${headers.map((h,i)=>`<option value="${i}" ${h===f?'selected':''}>${esc(h)} · ${i+1}</option>`).join('')}</select></label>`).join('')+'</div>';
    $('preview').innerHTML=`<h3 class="preview-title">Предпросмотр · первые 5 записей · ${esc(pending.name)}</h3><div class="tablewrap">${table(headers,pending.grid.slice(headerRow,headerRow+5).map(r=>r.map(c=>c.f!==undefined?'[формула]':String(c.v??''))))}</div>`;
    $('confirmImport').disabled=false;diagnostics([]);
    notice('Файл готов к проверке. Проверьте режим строк, формат суммы и сопоставление колонок.');
  }catch(e){$('mapping').innerHTML='';$('preview').innerHTML='';diagnostics([{severity:'error',code:'PARSE',sheet:$('sheet').value,row:'—',field:'файл',message:e.message}]);notice('Новый файл не принят. Ранее принятые данные сохранены.','error');}
}
$('file').addEventListener('change',async()=>{
  const file=$('file').files[0];if(!file)return;const token=++fileToken;
  stopReading();clearPending();diagnostics([]);
  const controller=new AbortController();readController=controller;
  try{
    if(file.size>5*1024*1024)throw Error('Файл больше 5 МБ. Разделите выгрузку.');
    if(!/\.(csv|xlsx)$/i.test(file.name))throw Error('Поддерживаются только CSV и XLSX.');
    const isXlsx=/\.xlsx$/i.test(file.name);
    if(isXlsx){reading(true);notice('Читаем Excel отдельно от интерфейса. Прежний отчёт и экспорт доступны; чтение можно отменить.');}
    const buffer=await file.arrayBuffer();if(token!==fileToken)return;
    const content=isXlsx?{book:await readWorkbookAsync(buffer,{signal:controller.signal})}:{text:new TextDecoder('utf-8',{fatal:true}).decode(buffer)};
    if(token!==fileToken)return;
    pending={name:file.name,demo:false,...content};
    $('mode').value='';$('headerRow').value='1';
    if(isXlsx){$('sheet').disabled=false;$('sheet').innerHTML=pending.book.SheetNames.map(n=>`<option>${esc(n)}</option>`).join('');}
    else{$('sheet').disabled=true;$('sheet').innerHTML='<option>CSV</option>';}
    prepare();
  }catch(e){if(token!==fileToken)return;clearPending();diagnostics([{severity:'error',code:'FILE',sheet:'—',row:'—',field:'файл',message:e.message}]);notice('Новый файл не принят. Ранее принятые данные и экспорт доступны.','error');}
  finally{if(token===fileToken){readController=null;reading(false);}}
});
for(const id of ['delimiter','sheet','headerRow'])$(id).addEventListener('change',prepare);
function importPending(){
  if(!pending)return false;
  try{
    const mode=$('mode').value;if(!mode)throw Error('Выберите, что означает строка: событие целиком или часть события.');
    const map=mapping();for(const f of FIELDS.filter(f=>f!=='line_id'||mode==='parts'))if(map[f]==='')throw Error(`Выберите колонку ${f}.`);
    if(mode==='event')delete map.line_id;
    const opts={mode,decimal:$('decimal').value,sheet:$('sheet').value,headerRow:Number($('headerRow').value)};
    const {rows,errors}=mapRows(pending.grid,map,{...opts,book:pending.book,XLSX:globalThis.XLSX});
    const result=validateRows(rows,opts);const all=[...errors,...result.diagnostics];diagnostics(all);
    if(errors.length||!result.ok){notice('Новый файл отклонён. Исправьте ошибки; текущие данные и отчёт сохранены.','error');return false;}
    accepted={events:result.events,meta:{name:pending.name,demo:pending.demo,...opts,mapping:map,stats:result.stats,diagnostics:all}};page=0;
    const dates=result.events.map(e=>e.date).sort();$('from').value=dates[0]||'2026-01-01';$('to').value=dates.at(-1)||'2026-01-01';
    const currencies=[...new Set(result.events.map(e=>e.currency))].sort();$('currency').innerHTML=currencies.map(c=>`<option>${esc(c)}</option>`).join('');
    $('currency').value=currencies.includes('RUB')?'RUB':currencies[0];
    render();notice(`Принят полный набор: ${result.stats.inputRows} строк → ${result.stats.events} событий. ${result.stats.duplicates?`Точных дублей: ${result.stats.duplicates}, учтены один раз.`:'Ошибок нет.'}${pending.demo?' Все данные вымышлены.':''}`,'success');
    return true;
  }catch(e){diagnostics([{severity:'error',code:'SETTINGS',sheet:$('sheet').value,row:'—',field:'настройки',message:e.message}]);notice('Новый файл не принят. Текущий отчёт сохранён.','error');return false;}
}
$('confirmImport').onclick=importPending;
async function demo(initial=false){
  const token=++fileToken;
  stopReading();clearPending();
  try{$('file').value='';const response=await fetch('./samples/orders-parts.csv');if(!response.ok)throw Error('Не удалось прочитать контрольный файл.');const text=await response.text();if(token!==fileToken)return;pending={name:'orders-parts.csv',text,demo:true};$('delimiter').value=',';$('decimal').value='.';$('headerRow').value='1';$('sheet').innerHTML='<option>CSV</option>';$('sheet').disabled=true;prepare();$('mode').value='parts';if(importPending()){$('from').value='2026-01-01';$('to').value='2026-02-28';render();if(initial)$('importPanel').hidden=true;}}catch(e){if(token===fileToken)notice(e.message,'error');}
}
$('loadDemo').onclick=()=>demo();
function chart(report){
  const points=report.dynamics;
  if(!points.length){$('chart').innerHTML='<div class="empty">В этом периоде нет событий.</div>';return;}
  if(points.length>60){$('chart').innerHTML='<div class="empty">Для графика выберите до 60 активных периодов. Все значения доступны в таблице и экспорте.</div>';return;}
  const max=points.reduce((m,d)=>[d.gross,d.refunds,d.net<0n?-d.net:d.net].reduce((a,b)=>a>b?a:b,m),1n);
  const height=v=>Number((v<0n?-v:v)*6500n/max)/100;
  const w=600,step=500/points.length,start=65;
  let svg='<svg class="chartsvg" viewBox="0 0 600 205" role="img" aria-label="Оплаты, возвраты и разница по периодам. Точные суммы в таблице ниже.">';
  for(let j=0;j<3;j++){const y=40+j*65;svg+=`<line x1="50" x2="585" y1="${y}" y2="${y}" stroke="#eaf0f6"/>`;}
  svg+='<line x1="50" x2="585" y1="105" y2="105" stroke="#cedaeb"/><text x="26" y="109">0</text>';
  points.forEach((d,i)=>{const x=start+i*step+step/2;const bw=Math.min(25,step*.24);const gh=height(d.gross),rh=height(d.refunds),nh=height(d.net);svg+=`<rect x="${x-bw-3}" y="${105-gh}" width="${bw}" height="${gh}" rx="3" fill="#376ce3"/><rect x="${x+3}" y="${105-rh}" width="${bw}" height="${rh}" rx="3" fill="#c7d6ef"/><circle cx="${x}" cy="${d.net<0n?105+nh:105-nh}" r="4" fill="${d.net<0n?'#b3473f':'#1b8a79'}"/>`;if(points.length<=12)svg+=`<text x="${x}" y="${d.net<0n?117+nh:95-nh}" text-anchor="middle" class="${d.net<0n?'negative':'netlabel'}">${esc(cash(d.net,report.currency))}</text>`;if(points.length<=16)svg+=`<text x="${x}" y="195" text-anchor="middle">${esc(d.period)}</text>`;});
  $('chart').innerHTML=svg+'</svg>';
}
function render(){
  if(!accepted)return;
  try{
    currentReport=compute(accepted.events,{from:$('from').value,to:$('to').value,currency:$('currency').value,grain:$('grain').value});
    const r=currentReport,m=accepted.meta,s=m.stats;
    $('filterError').hidden=true;$('exportCSV').disabled=false;$('exportJSON').disabled=false;
    $('sourceName').textContent=m.name;$('sourceInfo').textContent=`${m.demo?'Вымышленные данные':'Файл пользователя'} · ${s.inputRows} строк · ${s.events} событий · полный набор`;
    $('net').textContent=cash(r.net,r.currency);$('aov').textContent=r.aov===null?'Нет заказов':cash(r.aov,r.currency);
    $('grossInline').textContent='Оплаты '+cash(r.gross,r.currency);$('refundsInline').textContent='Возвраты '+cash(r.refunds,r.currency);
    $('orderCount').textContent=`Оплаченных заказов: ${r.orders}`;$('repeat').textContent=r.repeatBasisPoints===null?'Нет данных':(r.repeatBasisPoints/100).toFixed(2).replace('.',',')+'%';
    $('repeatDetails').textContent=`${r.repeatCustomers} из ${r.customers} клиентов · ${r.eligibleOrders} заказов с остатком`;$('cutoff').textContent='Дата отсечения: '+r.to;
    $('qualityTitle').textContent='Набор проверен';$('qualityText').textContent=s.duplicates?`Удалено точных дублей: ${s.duplicates}. Конфликтов и ошибок нет.`:'Обязательные поля и история заказов проверены.';
    $('qualityStats').innerHTML=[['Строк в файле',s.inputRows],['Уникальных частей / строк',s.uniqueRows],['Событий после свёртки',s.events],['Неустранённых ошибок',0]].map(([k,v])=>`<div class="qualityrow"><span>${esc(k)}</span><b>${v}</b></div>`).join('');
    chart(r);$('dynamics').innerHTML=r.dynamics.length?table(['Период','Оплаты','Возвраты','Разница','Заказы'],r.dynamics.map(d=>[d.period,cash(d.gross,r.currency),cash(d.refunds,r.currency),cash(d.net,r.currency),d.orders])):'<p class="empty">Нет событий. Оплаты и возвраты равны нулю; средний чек и доля отсутствуют.</p>';
    renderEvents();
  }catch(e){currentReport=null;$('filterError').textContent=e.message+' Ниже показан предыдущий отчёт. Экспорт недоступен до исправления периода.';$('filterError').hidden=false;$('exportCSV').disabled=true;$('exportJSON').disabled=true;}
}
function renderEvents(){
  const all=accepted.events,perPage=12,last=Math.max(0,Math.ceil(all.length/perPage)-1);page=Math.min(last,page);
  const events=all.slice(page*perPage,(page+1)*perPage);
  $('eventsCaption').textContent=`${all.length} событий всей истории · свёрнуты части, удалены дубли · фильтр дат применяется к показателям выше`;
  $('events').innerHTML=`<table><thead><tr>${['Событие','Заказ','Клиент','Дата','Тип','Сумма','Валюта'].map(h=>`<th scope="col">${h}</th>`).join('')}</tr></thead><tbody>${events.map(e=>`<tr><td>${esc(e.event_id)}</td><td>${esc(e.order_id)}</td><td>${esc(e.customer_id)}</td><td>${e.date}</td><td><span class="eventtype ${e.type}">${{sale:'Оплата',refund:'Возврат',cancel:'Отмена'}[e.type]}</span></td><td class="num">${esc(cash(e.amount,e.currency))}</td><td>${e.currency}</td></tr>`).join('')}</tbody></table>`;
  $('pageInfo').textContent=`${page+1} / ${last+1}`;$('prevPage').disabled=page===0;$('nextPage').disabled=page===last;
}
for(const id of ['from','to','currency','grain'])$(id).addEventListener('change',render);
$('prevPage').onclick=()=>{page--;renderEvents();};$('nextPage').onclick=()=>{page++;renderEvents();};
function download(content,extension,type){if(!currentReport)return;const url=URL.createObjectURL(new Blob([content],{type}));const a=document.createElement('a');a.href=url;a.download=`orderlens-${currentReport.currency}-${currentReport.from}-${currentReport.to}.${extension}`;a.click();setTimeout(()=>URL.revokeObjectURL(url),10000);}
$('exportCSV').onclick=()=>{if(currentReport)download(csvReport(currentReport,accepted.meta),'csv','text/csv;charset=utf-8');};
$('exportJSON').onclick=()=>{if(currentReport)download(jsonReport(currentReport,accepted.meta,accepted.events),'json','application/json');};
await demo(true);
