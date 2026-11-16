import test from 'node:test';
import assert from 'node:assert/strict';
import {readWorkbookAsync, XLSX_TIMEOUT_MS} from '../src/xlsx-reader.mjs';

const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const workbook = () => ({
  SheetNames: ['Заказы'],
  Workbook: {WBProps: {date1904: true}},
  Sheets: {'Заказы': {
    '!ref': 'A1:B2',
    A1: {t: 's', v: 'event_id'}, B1: {t: 's', v: 'date'},
    A2: {t: 's', v: '001'}, B2: {t: 'n', v: 44500, z: 'yyyy-mm-dd'},
  }},
});

/**
 * A worker double that uses real structuredClone transfer semantics. It can
 * deliberately deliver late callbacks even after terminate(), to exercise the
 * reader's own settle guard rather than relying on the browser suppressing them.
 */
class FakeWorker {
  onmessage = null;
  onerror = null;
  onmessageerror = null;
  calls = [];
  terminationCount = 0;
  postMessageError = null;
  afterPost = null;
  postMessage(payload, transfer) {
    this.calls.push({payload, transfer});
    if (this.postMessageError) throw this.postMessageError;
    this.received = structuredClone(payload, {transfer});
    this.afterPost?.();
  }
  terminate() { this.terminationCount += 1; }
  message(data) { this.onmessage?.({data: structuredClone(data)}); }
  runtimeError(event = {}) { this.onerror?.(event); }
  messageError() { this.onmessageerror?.({}); }
}

/** Observe per-read abort-listener registration without modifying global APIs. */
function trackedAbortController() {
  const controller = new AbortController();
  const listeners = new Set();
  const signal = {
    get aborted() { return controller.signal.aborted; },
    addEventListener(type, listener, options) {
      assert.equal(type, 'abort');
      listeners.add(listener);
      controller.signal.addEventListener(type, listener, options);
    },
    removeEventListener(type, listener) {
      assert.equal(type, 'abort');
      listeners.delete(listener);
      controller.signal.removeEventListener(type, listener);
    },
  };
  return {signal, abort: () => controller.abort(), listeners};
}

function start({timeoutMs = 1000, signal, worker = new FakeWorker(), buffer = new Uint8Array([80, 75, 3, 4, 1, 2]).buffer} = {}) {
  const promise = readWorkbookAsync(buffer, {timeoutMs, signal, workerFactory: () => worker});
  return {promise, worker, buffer};
}

test('public default deadline is five seconds', () => {
  assert.equal(XLSX_TIMEOUT_MS, 5000);
});

test('transfers the original ArrayBuffer instead of copying and returns a cloneable workbook', async () => {
  const abort = trackedAbortController();
  const input = new Uint8Array([80, 75, 3, 4, 7, 9]).buffer;
  const {worker, promise} = start({buffer: input, signal: abort.signal});
  assert.equal(worker.calls.length, 1);
  assert.strictEqual(worker.calls[0].payload.buffer, input);
  assert.equal(worker.calls[0].transfer.length, 1);
  assert.strictEqual(worker.calls[0].transfer[0], input);
  assert.equal(input.byteLength, 0, 'the original buffer is detached by transfer');
  assert.deepEqual([...new Uint8Array(worker.received.buffer)], [80, 75, 3, 4, 7, 9]);
  assert.equal(abort.listeners.size, 1);
  const book = workbook();
  worker.message({ok: true, book});
  const result = await promise;
  assert.deepEqual(result, book);
  assert.notStrictEqual(result, book, 'simulate the worker structured-clone boundary');
  assert.equal(result.Sheets['Заказы'].A2.v, '001');
  assert.equal(result.Workbook.WBProps.date1904, true);
  assert.equal(worker.terminationCount, 1);
  assert.equal(abort.listeners.size, 0);
});

test('worker-reported parse failure rejects with its diagnostic and releases worker/listener', async () => {
  const abort = trackedAbortController();
  const {worker, promise} = start({signal: abort.signal});
  const rejected = assert.rejects(promise, /Повреждённый ZIP/);
  worker.message({ok: false, message: 'Повреждённый ZIP'});
  await rejected;
  assert.equal(worker.terminationCount, 1);
  assert.equal(abort.listeners.size, 0);
});

test('a malformed success cannot be accepted as a workbook', async () => {
  for (const data of [undefined, null, {}, {ok: true}, {ok: true, book: {SheetNames: []}}, {ok: false}]) {
    const {worker, promise} = start();
    const rejected = assert.rejects(promise, /структуру книги Excel/);
    worker.message(data);
    await rejected;
    assert.equal(worker.terminationCount, 1);
  }
});

test('worker runtime error is handled and prevents an uncaught browser error', async () => {
  const abort = trackedAbortController();
  const {worker, promise} = start({signal: abort.signal});
  const rejected = assert.rejects(promise, /безопасно прочитать Excel/);
  let prevented = 0;
  worker.runtimeError({preventDefault() { prevented += 1; }});
  await rejected;
  assert.equal(prevented, 1);
  assert.equal(worker.terminationCount, 1);
  assert.equal(abort.listeners.size, 0);
});

test('structured-clone delivery failure rejects and cleans up', async () => {
  const abort = trackedAbortController();
  const {worker, promise} = start({signal: abort.signal});
  const rejected = assert.rejects(promise, /Ошибка передачи прочитанной книги/);
  worker.messageError();
  await rejected;
  assert.equal(worker.terminationCount, 1);
  assert.equal(abort.listeners.size, 0);
});

test('a nonresponding parser is terminated at a short test deadline', async () => {
  const abort = trackedAbortController();
  const {worker, promise} = start({timeoutMs: 20, signal: abort.signal});
  await assert.rejects(promise, /превысило .*остановлено/);
  assert.equal(worker.terminationCount, 1);
  assert.equal(abort.listeners.size, 0);
  // A buffered callback after the deadline cannot revive the parse.
  worker.message({ok: true, book: workbook()});
  abort.abort();
  assert.equal(worker.terminationCount, 1);
});

test('abort before reading creates no worker and leaves the caller buffer intact', async () => {
  const abort = trackedAbortController();
  abort.abort();
  const input = new Uint8Array([1, 2, 3]).buffer;
  let factories = 0;
  await assert.rejects(readWorkbookAsync(input, {
    signal: abort.signal,
    timeoutMs: 20,
    workerFactory() { factories += 1; return new FakeWorker(); },
  }), {name: 'AbortError'});
  assert.equal(factories, 0);
  assert.equal(input.byteLength, 3);
  assert.equal(abort.listeners.size, 0);
});

test('abort during parsing promptly rejects, terminates once, and ignores all late events', async () => {
  const abort = trackedAbortController();
  const {worker, promise} = start({timeoutMs: 20, signal: abort.signal});
  const rejected = assert.rejects(promise, {name: 'AbortError'});
  abort.abort();
  await rejected;
  worker.message({ok: true, book: workbook()});
  worker.message({ok: false, message: 'late failure'});
  worker.runtimeError();
  worker.messageError();
  await delay(35);
  assert.equal(worker.terminationCount, 1);
  assert.equal(abort.listeners.size, 0);
});

test('a completed read clears its deadline and ignores later abort/error/messages', async () => {
  const abort = trackedAbortController();
  const {worker, promise} = start({timeoutMs: 20, signal: abort.signal});
  const original = workbook();
  worker.message({ok: true, book: original});
  const result = await promise;
  abort.abort();
  worker.message({ok: false, message: 'too late'});
  worker.runtimeError();
  worker.messageError();
  worker.message({ok: true, book: {...workbook(), SheetNames: ['Wrong']}});
  await delay(35);
  assert.deepEqual(result, original);
  assert.equal(worker.terminationCount, 1);
  assert.equal(abort.listeners.size, 0);
});

test('a failed transfer rejects its original error and cleans up a created worker', async () => {
  const abort = trackedAbortController();
  const worker = new FakeWorker();
  const error = new DOMException('Cannot transfer this input', 'DataCloneError');
  worker.postMessageError = error;
  const {promise} = start({worker, signal: abort.signal, timeoutMs: 20});
  await assert.rejects(promise, e => e === error);
  assert.equal(worker.terminationCount, 1);
  assert.equal(abort.listeners.size, 0);
  await delay(35);
  assert.equal(worker.terminationCount, 1);
});

test('worker construction failure rejects without registering an abort listener', async () => {
  const abort = trackedAbortController();
  const error = new Error('Worker blocked by CSP');
  await assert.rejects(readWorkbookAsync(new ArrayBuffer(4), {
    signal: abort.signal,
    timeoutMs: 20,
    workerFactory() { throw error; },
  }), e => e === error);
  assert.equal(abort.listeners.size, 0);
});

test('caller cancellation isolates the latest read from a late response of its predecessor', async () => {
  const firstAbort = trackedAbortController();
  const oldRead = start({signal: firstAbort.signal, timeoutMs: 20});
  let accepted = null;
  const oldOutcome = oldRead.promise.then(book => { accepted = book; }, error => error.name);
  // This is the replacement sequence used by app.mjs: abort old, start new.
  firstAbort.abort();
  const currentRead = start({timeoutMs: 1000});
  const currentBook = {...workbook(), SheetNames: ['Current']};
  const currentOutcome = currentRead.promise.then(book => { accepted = book; return book; });
  oldRead.worker.message({ok: true, book: {...workbook(), SheetNames: ['Old']}});
  assert.equal(await oldOutcome, 'AbortError');
  assert.equal(accepted, null);
  currentRead.worker.message({ok: true, book: currentBook});
  await currentOutcome;
  oldRead.worker.message({ok: true, book: {...workbook(), SheetNames: ['LateOld']}});
  assert.deepEqual(accepted, currentBook);
  assert.equal(oldRead.worker.terminationCount, 1);
  assert.equal(currentRead.worker.terminationCount, 1);
});
