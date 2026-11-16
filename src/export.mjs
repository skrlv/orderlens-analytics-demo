import {money} from './engine.mjs';
export const RULES={gross:'G = сумма sale в периоде; фактически оплачено после скидок',refunds:'R = сумма refund в периоде, включая возвраты прежних заказов',net:'N = G − R; денежные события, не прибыль',orders:'O = уникальные order_id с sale в периоде; возврат не стирает оплату',aov:'G / O до возвратов; округление half-up до 0.01; O=0 → нет оплаченных заказов',eligibleOrders:'E = оплаченные в периоде заказы с остатком > 0 на дату отсечения',customers:'C = уникальные клиенты заказов E, только выбранная валюта',repeatCustomers:'C₂ = клиенты с ≥ 2 разными заказами из E в периоде',repeatBasisPoints:'C₂ / C; округление half-up до 0.01%; C=0 → нет данных для доли'};
export function reportRecords(report,meta){
  const base={period:`${report.from}..${report.to}`,currency:report.currency};
  const records=[];
  for(const [metric,value]of Object.entries({source:meta.name,dataset:meta.demo?'Вымышленные данные':'Файл пользователя',scope:'Полный принятый набор; ошибки блокируют замену',timezone:'Europe/Moscow; даты без времени',cutoff:report.to,mode:meta.mode,decimal:meta.decimal,inputRows:meta.stats.inputRows,duplicates:meta.stats.duplicates,events:meta.stats.events,mapping:JSON.stringify(meta.mapping),headerRow:meta.headerRow,sheet:meta.sheet,grain:report.grain,exportedAt:new Date().toISOString()}))records.push({section:'context',metric,...base,value:String(value),rule:''});
  for(const metric of Object.keys(RULES)){let value=report[metric];value=typeof value==='bigint'?money(value):value===null?'нет данных':metric==='repeatBasisPoints'?(value/100).toFixed(2)+'%':String(value);records.push({section:'summary',metric,...base,value,rule:RULES[metric]});}
  for(const d of report.dynamics)for(const metric of ['gross','refunds','net','orders'])records.push({section:'dynamics',metric,period:d.period,currency:report.currency,value:typeof d[metric]==='bigint'?money(d[metric]):String(d[metric]),rule:RULES[metric]});
  for(const d of meta.diagnostics)records.push({section:'diagnostic',metric:d.code,...base,value:d.message,rule:`${d.sheet} / запись ${d.row} / ${d.field}`});
  return records;
}
export function csvReport(report,meta){
  const columns=['section','metric','period','currency','value','rule'];
  const cell=(v,numeric=false)=>{v=String(v);if(!numeric&&(/^\s*[=+\-@]/.test(v)||/^[\t\r\n]/.test(v)))v="'"+v;return '"'+v.replaceAll('"','""')+'"';};
  return '\ufeff'+[columns.map(x=>cell(x)).join(';'),...reportRecords(report,meta).map(r=>columns.map(k=>cell(r[k]??'',k==='value'&&/^-?\d+(\.\d+)?%?$/.test(r[k]))).join(';'))].join('\r\n')+'\r\n';
}
export function jsonReport(report,meta,events){return JSON.stringify({schema:'orderlens-report-v1',report,context:meta,rules:RULES,events},(_k,v)=>typeof v==='bigint'?money(v):v,2);}
