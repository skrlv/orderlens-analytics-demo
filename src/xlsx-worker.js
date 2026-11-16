/* Static same-origin worker: no blob/eval and no table data sent to a server. */
importScripts('../vendor/xlsx.full.min.js');
self.onmessage = async ({data}) => {
  try {
    const {readWorkbook} = await import('./import.mjs');
    const parsed = readWorkbook(data.buffer, self.XLSX);
    if (!parsed.SheetNames?.length || parsed.SheetNames.length > 32) throw Error('Книга должна содержать от 1 до 32 листов.');
    const book = {SheetNames:parsed.SheetNames, Workbook:{WBProps:{date1904:!!parsed.Workbook?.WBProps?.date1904}}, Sheets:Object.create(null)};
    let cells = 0, textSize = 0;
    for (const name of book.SheetNames) {
      const source = parsed.Sheets[name], sheet = {};
      if (source?.['!ref']) {
        const range = self.XLSX.utils.decode_range(source['!ref']);
        if (range.e.r >= 10100 || range.e.c >= 64) throw Error('Лимит листа: 10 000 строк данных, до 100 строк перед заголовком, 64 колонки.');
        sheet['!ref'] = source['!ref'];
      }
      for (const [address, cell] of Object.entries(source || {})) {
        if (!/^[A-Z]+[1-9]\d*$/.test(address)) continue;
        if (++cells > 650000) throw Error('Книга содержит больше 650 000 ячеек. Оставьте только нужный лист.');
        const clean = {};
        for (const key of ['t','v','f','z']) if (cell[key] !== undefined) {
          clean[key] = cell[key];
          if (typeof cell[key] === 'string') textSize += cell[key].length;
        }
        if (textSize > 32*1024*1024) throw Error('Слишком много текста в книге. Разделите выгрузку.');
        sheet[address] = clean;
      }
      book.Sheets[name] = sheet;
    }
    self.postMessage({ok:true,book});
  } catch (error) { self.postMessage({ok:false,message:error.message || 'Повреждённая книга Excel.'}); }
  finally { self.close(); }
};
