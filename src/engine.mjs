/** OrderLens data contract v1. Money is represented exclusively in minor units. */
export const FIELDS = Object.freeze(['event_id', 'line_id', 'order_id', 'customer_id', 'date', 'type', 'amount', 'currency']);
export const CURRENCIES = Object.freeze(['RUB', 'USD', 'EUR']);
export const TIMEZONE = 'Europe/Moscow';

export function validDate(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const [year, month, day] = value.split('-').map(Number);
  if (year < 1 || month < 1 || month > 12 || day < 1) return false;
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  return day <= [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1];
}

function text(value) {
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' && Number.isFinite(value)) return String(value);
  return '';
}

function amountInCents(value, decimal) {
  // A Number above this conservative boundary may already have lost cents
  // before reaching the validator. Large decimal strings remain exact.
  if (typeof value === 'number' && Math.abs(value) > 1000000000000) return null;
  const input = text(value);
  const expression = decimal === ',' ? /^\d+(?:,\d{1,2})?$/ : /^\d+(?:\.\d{1,2})?$/;
  if (!expression.test(input)) return null;
  const [whole, fraction = ''] = input.split(decimal);
  return BigInt(whole) * 100n + BigInt(fraction.padEnd(2, '0'));
}

const signature = event => JSON.stringify([
  event.event_id, event.order_id, event.customer_id, event.date,
  event.type, event.amount.toString(), event.currency,
]);
const commonSignature = event => JSON.stringify([
  event.event_id, event.order_id, event.customer_id, event.date, event.type, event.currency,
]);

/**
 * Atomic validation: on any error no events are returned. Exact duplicates are
 * warnings. A caller replaces its accepted dataset only when result.ok is true.
 */
export function validateRows(rows, options = {}) {
  const { mode, decimal = '.', sheet = 'CSV', headerRow = 1 } = options;
  const diagnostics = [];
  const stats = { inputRows: Array.isArray(rows) ? rows.length : 0, uniqueRows: 0, duplicates: 0, events: 0, sales: 0, refunds: 0, cancellations: 0 };
  const diagnostic = (severity, code, row, field, message) => diagnostics.push({ severity, code, sheet, row, field, message });
  const result = events => ({ ok: !diagnostics.some(item => item.severity === 'error'), events, diagnostics, stats });

  if (mode !== 'event' && mode !== 'parts') diagnostic('error', 'MODE', headerRow, 'mode', 'Выберите режим: событие целиком или часть события.');
  if (decimal !== '.' && decimal !== ',') diagnostic('error', 'DECIMAL', headerRow, 'amount', 'Выберите десятичный разделитель: точка или запятая.');
  if (!Array.isArray(rows) || rows.length === 0) diagnostic('error', 'EMPTY', headerRow, '', 'В таблице нет строк данных.');
  if (diagnostics.length) return result([]);

  const unique = new Map();
  rows.forEach((source, index) => {
    const row = Number.isInteger(source?._row) && source._row > 0 ? source._row : headerRow + index + 1;
    const event = { _row: row };
    const before = diagnostics.length;
    for (const field of ['event_id', 'order_id', 'customer_id']) {
      event[field] = text(source?.[field]);
      if (!event[field]) diagnostic('error', 'REQUIRED', row, field, 'Обязательный идентификатор не заполнен или имеет неподдерживаемый тип.');
    }
    if (mode === 'parts') {
      event.line_id = text(source?.line_id);
      if (!event.line_id) diagnostic('error', 'REQUIRED', row, 'line_id', 'Для частей события нужен line_id.');
    }
    event.date = text(source?.date);
    if (!validDate(event.date)) diagnostic('error', 'DATE', row, 'date', 'Нужна существующая дата YYYY-MM-DD. Время и неоднозначные даты не поддерживаются.');
    event.type = text(source?.type).toLowerCase();
    if (!['sale', 'refund', 'cancel'].includes(event.type)) diagnostic('error', 'TYPE', row, 'type', 'Допустимы только sale, refund и cancel.');
    event.currency = text(source?.currency).toUpperCase();
    if (!CURRENCIES.includes(event.currency)) diagnostic('error', 'CURRENCY', row, 'currency', 'Укажите RUB, USD или EUR; все поддерживаемые валюты имеют 2 десятичных знака.');
    event.amount = amountInCents(source?.amount, decimal);
    if (typeof source?.amount === 'number' && Math.abs(source.amount) > 1000000000000) diagnostic('error', 'AMOUNT_UNSAFE_NUMBER', row, 'amount', 'Числовая сумма больше 1 000 000 000 000. Сохраните её десятичным текстом для точного чтения копеек/центов.');
    else if (event.amount === null) diagnostic('error', 'AMOUNT', row, 'amount', `Нужна явно заданная неотрицательная сумма, без разделителей тысяч и не более 2 знаков после «${decimal}».`);
    else if ((event.type === 'sale' || event.type === 'refund') && event.amount <= 0n) diagnostic('error', 'AMOUNT_POSITIVE', row, 'amount', 'Для оплаты и возврата нужна положительная фактическая сумма.');
    else if (event.type === 'cancel' && event.amount !== 0n) diagnostic('error', 'CANCEL_AMOUNT', row, 'amount', 'Неоплаченная отмена должна иметь явно заданную сумму 0.');
    if (diagnostics.length !== before) return;
    const key = JSON.stringify(mode === 'parts' ? [event.event_id, event.line_id] : [event.event_id]);
    const previous = unique.get(key);
    if (previous) {
      if (signature(previous) === signature(event)) {
        stats.duplicates += 1;
        diagnostic('warning', 'DUPLICATE', row, mode === 'parts' ? 'event_id + line_id' : 'event_id', `Точный дубль строки ${previous._row}: учтён один раз.`);
      } else diagnostic('error', 'KEY_CONFLICT', row, mode === 'parts' ? 'event_id + line_id' : 'event_id', `Ключ уже встречался в строке ${previous._row}, но содержимое отличается.`);
    } else unique.set(key, event);
  });
  stats.uniqueRows = unique.size;

  const byEvent = new Map();
  for (const line of unique.values()) {
    const existing = byEvent.get(line.event_id);
    if (!existing) byEvent.set(line.event_id, { ...line });
    else if (commonSignature(existing) !== commonSignature(line)) diagnostic('error', 'EVENT_CONFLICT', line._row, 'event_id', `Части события ${line.event_id} должны иметь одинаковые заказ, клиента, дату, тип и валюту.`);
    else existing.amount += line.amount;
  }
  const candidates = [...byEvent.values()];
  const byOrder = new Map();
  for (const event of candidates) {
    const group = byOrder.get(event.order_id) ?? [];
    group.push(event);
    byOrder.set(event.order_id, group);
  }
  for (const [orderId, events] of byOrder) {
    const first = events[0];
    for (const event of events.slice(1)) {
      if (event.customer_id !== first.customer_id) diagnostic('error', 'ORDER_CUSTOMER', event._row, 'customer_id', `У заказа ${orderId} встречаются разные клиенты.`);
      if (event.currency !== first.currency) diagnostic('error', 'ORDER_CURRENCY', event._row, 'currency', `У заказа ${orderId} встречаются разные валюты.`);
    }
    const sales = events.filter(event => event.type === 'sale');
    const refunds = events.filter(event => event.type === 'refund');
    const cancellations = events.filter(event => event.type === 'cancel');
    if (sales.length > 1) diagnostic('error', 'MULTIPLE_SALES', sales[1]._row, 'order_id', `У заказа ${orderId} несколько отдельных оплат. Рассрочки и повторные оплаты не входят в договор данных.`);
    if (cancellations.length > 1) diagnostic('error', 'MULTIPLE_CANCELLATIONS', cancellations[1]._row, 'order_id', `У заказа ${orderId} несколько разных событий отмены. Нужна одна итоговая неоплаченная отмена.`);
    if (sales.length && cancellations.length) diagnostic('error', 'SALE_CANCEL', cancellations[0]._row, 'type', `Заказ ${orderId} и оплачен, и отменён. Для списания оплаченной суммы нужен refund.`);
    if (refunds.length && sales.length === 0) diagnostic('error', 'ORPHAN_REFUND', refunds[0]._row, 'order_id', `Для возврата заказа ${orderId} нет исходной оплаты. Загрузите полную историю до фильтрации периода.`);
    if (sales.length === 1) {
      const sale = sales[0];
      let refunded = 0n;
      for (const refund of refunds.toSorted((a, b) => a.date.localeCompare(b.date) || a.event_id.localeCompare(b.event_id))) {
        if (refund.date < sale.date) diagnostic('error', 'REFUND_BEFORE_SALE', refund._row, 'date', `Возврат заказа ${orderId} датирован раньше оплаты.`);
        refunded += refund.amount;
        if (refunded > sale.amount) diagnostic('error', 'OVER_REFUND', refund._row, 'amount', `Накопленные возвраты заказа ${orderId} превышают оплаченную сумму.`);
      }
    }
  }

  stats.events = candidates.length;
  stats.sales = candidates.filter(event => event.type === 'sale').length;
  stats.refunds = candidates.filter(event => event.type === 'refund').length;
  stats.cancellations = candidates.filter(event => event.type === 'cancel').length;
  if (diagnostics.some(item => item.severity === 'error')) return result([]);
  const events = candidates.map(({ _row, line_id, ...event }) => event).toSorted((a, b) => a.date.localeCompare(b.date) || a.event_id.localeCompare(b.event_id));
  return result(events);
}

const sum = events => events.reduce((total, event) => total + event.amount, 0n);
const roundRatio = (numerator, denominator) => (numerator * 2n + denominator) / (denominator * 2n);

/** Inclusive calendar dates in Europe/Moscow; input is accepted history. */
export function compute(events, { from, to, currency, grain = 'month' } = {}) {
  if (!validDate(from) || !validDate(to) || from > to) throw new RangeError('Укажите существующие даты начала и конца в правильном порядке.');
  if (!CURRENCIES.includes(currency)) throw new RangeError('Выберите одну поддерживаемую валюту.');
  if (!['month', 'day'].includes(grain)) throw new RangeError('Допустима динамика по дням или месяцам.');
  const history = events.filter(event => event.currency === currency);
  const periodEvents = history.filter(event => event.date >= from && event.date <= to);
  const sales = periodEvents.filter(event => event.type === 'sale');
  const refunds = periodEvents.filter(event => event.type === 'refund');
  const gross = sum(sales);
  const returned = sum(refunds);
  const orders = new Set(sales.map(event => event.order_id)).size;
  const refundsToCutoff = new Map();
  for (const event of history) {
    if (event.type === 'refund' && event.date <= to) refundsToCutoff.set(event.order_id, (refundsToCutoff.get(event.order_id) ?? 0n) + event.amount);
  }
  const customers = new Map();
  let eligibleOrders = 0;
  for (const sale of sales) {
    if (sale.amount - (refundsToCutoff.get(sale.order_id) ?? 0n) > 0n) {
      eligibleOrders += 1;
      const customerOrders = customers.get(sale.customer_id) ?? new Set();
      customerOrders.add(sale.order_id);
      customers.set(sale.customer_id, customerOrders);
    }
  }
  const repeatCustomers = [...customers.values()].filter(customerOrders => customerOrders.size >= 2).length;
  const buckets = new Map();
  for (const event of periodEvents) {
    const period = grain === 'month' ? event.date.slice(0, 7) : event.date;
    const bucket = buckets.get(period) ?? { period, gross: 0n, refunds: 0n, net: 0n, orders: 0 };
    if (event.type === 'sale') { bucket.gross += event.amount; bucket.orders += 1; }
    if (event.type === 'refund') bucket.refunds += event.amount;
    bucket.net = bucket.gross - bucket.refunds;
    buckets.set(period, bucket);
  }
  return {
    from, to, cutoff: to, timezone: TIMEZONE, currency, grain,
    gross, refunds: returned, net: gross - returned, orders,
    aov: orders ? roundRatio(gross, BigInt(orders)) : null,
    eligibleOrders, customers: customers.size, repeatCustomers,
    repeatBasisPoints: customers.size ? Number(roundRatio(BigInt(repeatCustomers) * 10000n, BigInt(customers.size))) : null,
    dynamics: [...buckets.values()].toSorted((a, b) => a.period.localeCompare(b.period)),
  };
}

/** Canonical decimal output, including negative event revenue. */
export function money(cents) {
  if (typeof cents !== 'bigint') throw new TypeError('money expects integer minor units (BigInt).');
  const absolute = cents < 0n ? -cents : cents;
  return `${cents < 0n ? '-' : ''}${absolute / 100n}.${String(absolute % 100n).padStart(2, '0')}`;
}
