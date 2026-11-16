export const XLSX_TIMEOUT_MS = 5000;

/** A fresh same-origin worker per read; every terminal path releases it. */
export function readWorkbookAsync(buffer, {signal, timeoutMs = XLSX_TIMEOUT_MS, workerFactory = () => new Worker(new URL('./xlsx-worker.js', import.meta.url))} = {}) {
  return new Promise((resolve, reject) => {
    let worker, timer, settled = false;
    const finish = (error, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener('abort', abort);
      worker?.terminate();
      if (error) reject(error); else resolve(value);
    };
    const abort = () => finish(new DOMException('Чтение Excel отменено. Прежний отчёт сохранён.', 'AbortError'));
    if (signal?.aborted) return abort();
    try {
      worker = workerFactory();
      signal?.addEventListener('abort', abort, {once:true});
      worker.onmessage = ({data}) => {
        if (data?.ok === true && data.book?.SheetNames?.length) finish(null, data.book);
        else finish(new Error(data?.message || 'Не удалось прочитать структуру книги Excel.'));
      };
      worker.onerror = event => { event.preventDefault?.(); finish(new Error('Не удалось безопасно прочитать Excel. Проверьте файл и повторите загрузку.')); };
      worker.onmessageerror = () => finish(new Error('Ошибка передачи прочитанной книги Excel. Прежний отчёт сохранён.'));
      timer = setTimeout(() => finish(new Error('Чтение Excel превысило 5 секунд и остановлено. Файл может быть повреждён или слишком сложен; сохраните его заново либо выгрузите CSV. Прежний отчёт доступен.')), timeoutMs);
      worker.postMessage({buffer}, [buffer]);
    } catch (error) { finish(error); }
  });
}
