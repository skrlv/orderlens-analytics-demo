import {validateXlsxContainer} from './xlsx-container.mjs';
export const MAX_ROWS=10000;
export function parseCSV(text,delimiter=',') {
  if(![',',';','\t'].includes(delimiter)) throw Error('Выберите разделитель CSV.');
  text=text.replace(/^\uFEFF/,'');
  const rows=[]; let row=[],field='',quoted=false,closed=false,start=true;
  const endField=()=>{row.push({t:'s',v:field});field='';closed=false;start=true;};
  const endRow=()=>{endField();rows.push(row);row=[];if(rows.length>MAX_ROWS+100)throw Error('Лимит: 10 000 строк данных и до 100 строк перед заголовком.');};
  for(let i=0;i<text.length;i++) {const ch=text[i];
    if(quoted){if(ch==='"'){if(text[i+1]==='"'){field+='"';i++;}else{quoted=false;closed=true;}}else field+=ch;continue;}
    if(ch===delimiter){endField();continue;}
    if(ch==='\n'||ch==='\r'){if(ch==='\r'&&text[i+1]==='\n')i++;endRow();continue;}
    if(closed)throw Error(`CSV: лишний символ после кавычки (запись ${rows.length+1}).`);
    if(ch==='"'){if(!start)throw Error(`CSV: кавычка внутри неэкранированного поля (запись ${rows.length+1}).`);quoted=true;start=false;continue;}
    field+=ch;start=false;
  }
  if(quoted)throw Error('CSV: не закрыты кавычки.');
  if(field||row.length||closed)endRow();
  if(!rows.length)throw Error('Файл пуст.');
  if(rows.some(r=>r.length>64))throw Error('Лимит: 64 колонки.');
  return rows;
}
/** Synchronous low-level read. The app invokes this only inside xlsx-worker.js. */
export function readWorkbook(buffer,XLSX){validateXlsxContainer(buffer);return XLSX.read(buffer,{type:'array',cellFormula:true,cellNF:true,cellText:true,cellDates:false,bookVBA:false});}
export function workbookGrid(book,name,XLSX){
  const sheet=book.Sheets[name]; if(!sheet?.['!ref'])throw Error('На выбранном листе нет ячеек.');
  const range=XLSX.utils.decode_range(sheet['!ref']);
  if(range.e.r>=MAX_ROWS+100||range.e.c>=64)throw Error('Лимит листа: 10 000 строк данных, до 100 строк перед заголовком, 64 колонки.');
  const rows=[];
  for(let r=0;r<=range.e.r;r++){const row=[];for(let c=0;c<=range.e.c;c++)row.push(sheet[XLSX.utils.encode_cell({r,c})]||{t:'s',v:''});rows.push(row);}return rows;
}
export function headersAt(grid,headerRow){
  if(!Number.isInteger(headerRow)||headerRow<1||headerRow>100)throw Error('Строка заголовков: целое число от 1 до 100.');
  const row=grid[headerRow-1];if(!row)throw Error('Строка заголовков отсутствует.');
  return row.map((c,i)=>String(c.v??'').trim()||`Колонка ${i+1}`);
}
export function mapRows(grid,mapping,{headerRow=1,sheet='CSV',book,XLSX,decimal='.'}){
  const used=Object.values(mapping).filter(x=>x!=='');
  if(new Set(used).size!==used.length)throw Error('Каждому полю нужна отдельная колонка.');
  const rows=[]; const errors=[];
  for(let r=headerRow;r<grid.length;r++){
    if(grid[r].every(c=>c.f===undefined&&c.t!=='e'&&(c.v===undefined||String(c.v).trim()==='')))continue;
    const row={_row:r+1};
    for(const [field,index] of Object.entries(mapping)){
      if(index===''){row[field]='';continue;}
      const c=grid[r][Number(index)]||{v:''}; let value=c.v??'';
      const problem=message=>errors.push({severity:'error',code:'CELL',sheet,row:r+1,field,message});
      if(c.f!==undefined){problem('Формула не принимается. Выгрузите рассчитанные значения.');value='';}
      else if(c.t==='e'||c.t==='b'){problem('Ошибка Excel или логическое значение в обязательном поле.');value='';}
      else if(field.endsWith('_id')&&typeof value==='number'){problem('ID должен быть текстовой ячейкой Excel, чтобы сохранить ведущие нули.');value='';}
      else if(field==='date'&&typeof value==='number'){
        const is1904=!!book?.Workbook?.WBProps?.date1904;
        if(!XLSX?.SSF.is_date(c.z||'')||!Number.isInteger(value)||(!is1904&&value===60)){problem('Нужна дата Excel без времени или текст YYYY-MM-DD; обычные числа не считаются датой.');value='';}
        else{const d=XLSX.SSF.parse_date_code(value,{date1904:is1904});value=d?`${String(d.y).padStart(4,'0')}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`:'';}
      }else if(field==='amount'&&typeof value==='number'){
        if(!Number.isFinite(value)||Math.abs(value)>1000000000000){problem('Слишком большая числовая сумма Excel. Сохраните сумму текстом для точного чтения копеек/центов.');value='';}
        else value=String(value).replace('.',decimal);
      }
      row[field]=value;
    }
    rows.push(row);
  }
  if(rows.length>MAX_ROWS)throw Error('Лимит: 10 000 непустых строк данных.');
  return {rows,errors};
}
