/**
 * Bounded ZIP/OOXML structural preflight. It never decompresses file contents,
 * evaluates XML or verifies content CRCs. A separately terminable parser is
 * still required for corrupt DEFLATE/XML inside a structurally valid archive.
 * ZIP field layout reference: PKWARE APPNOTE 6.3.10, sections 4.3.7/12/16.
 * https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
 */
export const XLSX_CONTAINER_LIMITS = Object.freeze({
  fileBytes: 5 * 1024 * 1024,
  entries: 1024,
  totalUncompressedBytes: 32 * 1024 * 1024,
  entryUncompressedBytes: 16 * 1024 * 1024,
});

const LOCAL = 0x04034b50;
const CENTRAL = 0x02014b50;
const EOCD = 0x06054b50;
const DESCRIPTOR = 0x08074b50;
const REQUIRED = ['[Content_Types].xml', 'xl/workbook.xml'];
const fail = message => { throw new Error(`XLSX: ${message}`); };

function bytesOf(buffer) {
  if (buffer instanceof ArrayBuffer) return new Uint8Array(buffer);
  if (ArrayBuffer.isView(buffer)) return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  return fail('нужен двоичный файл XLSX. Выберите файл заново.');
}

/** Returns plain metadata on success; throws a Russian Error on rejection. */
export function validateXlsxContainer(buffer) {
  const bytes = bytesOf(buffer);
  const size = bytes.byteLength;
  if (size > XLSX_CONTAINER_LIMITS.fileBytes) fail('файл больше 5 МБ. Разделите выгрузку.');
  if (size < 52) fail('файл обрезан: отсутствует полная структура ZIP. Выгрузите XLSX заново.');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = offset => view.getUint16(offset, true);
  const u32 = offset => view.getUint32(offset, true);
  if (u32(0) !== LOCAL) fail('файл не является поддерживаемой книгой XLSX (ZIP / OOXML).');
  const bounded = (offset, length, end = size) => Number.isInteger(offset) && Number.isInteger(length) && offset >= 0 && length >= 0 && offset <= end && length <= end - offset;

  // EOCD is at most 22 + 65535 bytes from EOF. Its comment must end at EOF;
  // trailing junk and truncated comments are not accepted as a second archive.
  let eocd = -1;
  const searchStart = Math.max(0, size - 22 - 65535);
  for (let pos = size - 22; pos >= searchStart; pos -= 1) {
    if (u32(pos) === EOCD && pos + 22 + u16(pos + 20) === size) { eocd = pos; break; }
  }
  if (eocd < 0) fail('файл обрезан или повреждён: не найден конец ZIP и его полный комментарий. Выгрузите XLSX заново.');
  const disk = u16(eocd + 4);
  const centralDisk = u16(eocd + 6);
  const diskEntries = u16(eocd + 8);
  const entryCount = u16(eocd + 10);
  const centralDirectoryBytes = u32(eocd + 12);
  const centralDirectoryOffset = u32(eocd + 16);
  if (disk === 0xffff || centralDisk === 0xffff || diskEntries === 0xffff || entryCount === 0xffff || centralDirectoryBytes === 0xffffffff || centralDirectoryOffset === 0xffffffff) fail('ZIP64 не поддерживается. Сохраните небольшую обычную книгу XLSX или CSV.');
  if (disk !== 0 || centralDisk !== 0 || diskEntries !== entryCount) fail('многотомные ZIP не поддерживаются. Нужен один полный файл XLSX.');
  if (entryCount === 0 || entryCount > XLSX_CONTAINER_LIMITS.entries) fail('число файлов внутри книги должно быть от 1 до 1024.');
  if (!bounded(centralDirectoryOffset, centralDirectoryBytes, eocd) || centralDirectoryOffset + centralDirectoryBytes !== eocd || centralDirectoryBytes < entryCount * 46) fail('повреждены границы центрального каталога ZIP. Выгрузите XLSX заново.');

  const extraFields = (start, length) => {
    const end = start + length;
    for (let pos = start; pos < end;) {
      if (!bounded(pos, 4, end)) fail('повреждён заголовок дополнительного поля ZIP.');
      const id = u16(pos);
      const fieldBytes = u16(pos + 2);
      if (!bounded(pos + 4, fieldBytes, end)) fail('дополнительное поле ZIP выходит за границы записи.');
      if (id === 0x0001) fail('ZIP64 не поддерживается. Сохраните обычный XLSX или CSV.');
      if (id === 0x9901) fail('зашифрованные XLSX не поддерживаются. Снимите пароль в исходной программе.');
      pos += 4 + fieldBytes;
    }
  };
  const settings = (version, flags, method) => {
    if (version >= 45) fail('ZIP64 и расширенные варианты ZIP не поддерживаются. Сохраните обычный XLSX или CSV.');
    if (flags & (0x0001 | 0x0040 | 0x2000)) fail('зашифрованные XLSX не поддерживаются. Снимите пароль в исходной программе.');
    if (flags & ~0x080e) fail('неподдерживаемые флаги ZIP. Сохраните обычный XLSX или CSV.');
    if (method !== 0 && method !== 8) fail('поддерживаются только ZIP без сжатия или Deflate. Сохраните обычный XLSX.');
  };
  const utf8 = new TextDecoder('utf-8', { fatal: true });
  const filename = (start, length, flags) => {
    if (!length) fail('пустое имя файла внутри ZIP.');
    const nameBytes = bytes.subarray(start, start + length);
    let name;
    if (flags & 0x0800) {
      try { name = utf8.decode(nameBytes); } catch { fail('имя файла ZIP содержит неверный UTF-8.'); }
    } else {
      // Non-UTF8 names are compared byte-for-byte. Required OOXML paths are
      // ASCII; no locale-dependent filename conversion is performed here.
      name = Array.from(nameBytes, byte => String.fromCharCode(byte)).join('');
    }
    if (/[\u0000-\u001f\u007f\\]/.test(name) || name.startsWith('/') || /^[A-Za-z]:/.test(name) || name.split('/').some(part => part === '.' || part === '..')) fail('недопустимое имя или путь внутри ZIP.');
    return name;
  };

  const entries = [];
  const names = new Set();
  const localOffsets = new Set();
  let totalUncompressedBytes = 0;
  let pos = centralDirectoryOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (!bounded(pos, 46, eocd) || u32(pos) !== CENTRAL) fail('число записей или заголовок центрального каталога ZIP повреждены.');
    const version = u16(pos + 6);
    const flags = u16(pos + 8);
    const method = u16(pos + 10);
    const crc = u32(pos + 16);
    const compressedSize = u32(pos + 20);
    const uncompressedSize = u32(pos + 24);
    const nameBytes = u16(pos + 28);
    const extraBytes = u16(pos + 30);
    const commentBytes = u16(pos + 32);
    const startDisk = u16(pos + 34);
    const localOffset = u32(pos + 42);
    const recordBytes = 46 + nameBytes + extraBytes + commentBytes;
    if (!bounded(pos, recordBytes, eocd)) fail('запись центрального каталога ZIP обрезана.');
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff || localOffset === 0xffffffff || startDisk === 0xffff) fail('ZIP64 не поддерживается. Сохраните обычный XLSX или CSV.');
    if (startDisk !== 0) fail('многотомные ZIP не поддерживаются.');
    settings(version, flags, method);
    if (uncompressedSize > XLSX_CONTAINER_LIMITS.entryUncompressedBytes) fail('отдельный файл внутри книги больше 16 МБ после распаковки. Уменьшите книгу или используйте CSV.');
    totalUncompressedBytes += uncompressedSize;
    if (totalUncompressedBytes > XLSX_CONTAINER_LIMITS.totalUncompressedBytes) fail('книга больше 32 МБ после распаковки. Разделите её или используйте CSV.');
    if (method === 0 && compressedSize !== uncompressedSize) fail('размер несжатой записи ZIP противоречит каталогу.');
    if (compressedSize === 0 && uncompressedSize > 0) fail('данные файла ZIP отсутствуют, хотя заявлен непустой файл.');
    const nameStart = pos + 46;
    const name = filename(nameStart, nameBytes, flags);
    if (names.has(name) || localOffsets.has(localOffset)) fail('дублирующиеся имена или смещения файлов в ZIP не поддерживаются.');
    names.add(name); localOffsets.add(localOffset);
    extraFields(nameStart + nameBytes, extraBytes);
    entries.push({ name, method, flags, compressedSize, uncompressedSize, localOffset, _version: version, _crc: crc, _nameStart: nameStart, _nameBytes: nameBytes });
    pos += recordBytes;
  }
  if (pos !== eocd) fail('размер или точное число записей центрального каталога ZIP не совпадают.');
  for (const name of REQUIRED) if (!entries.some(entry => entry.name === name && entry.uncompressedSize > 0)) fail(`не найдена обязательная часть книги ${name}. Нужен XLSX, а не произвольный ZIP.`);

  // Compute fixed local bounds first; reject overlaps before scanning variable
  // name/extra fields. Each local interval is subsequently visited only once.
  for (const entry of entries) {
    const start = entry.localOffset;
    if (!bounded(start, 30, centralDirectoryOffset) || u32(start) !== LOCAL) fail('локальный заголовок ZIP отсутствует, обрезан или имеет неверное смещение.');
    const version = u16(start + 4);
    const flags = u16(start + 6);
    const method = u16(start + 8);
    const nameBytes = u16(start + 26);
    const extraBytes = u16(start + 28);
    if (version !== entry._version || flags !== entry.flags || method !== entry.method) fail('локальный заголовок ZIP не совпадает с центральным каталогом.');
    const dataOffset = start + 30 + nameBytes + extraBytes;
    if (!bounded(start, 30 + nameBytes + extraBytes, centralDirectoryOffset) || !bounded(dataOffset, entry.compressedSize, centralDirectoryOffset)) fail('сжатые данные или имя локального файла ZIP обрезаны.');
    entry.dataOffset = dataOffset;
    entry.dataEnd = dataOffset + entry.compressedSize;
    entry._localNameBytes = nameBytes;
    entry._localExtraBytes = extraBytes;
  }
  const ordered = entries.toSorted((a, b) => a.localOffset - b.localOffset);
  if (ordered[0].localOffset !== 0) fail('начало локальных записей ZIP повреждено.');
  for (let index = 0; index < ordered.length; index += 1) {
    const entry = ordered[index];
    const nextOffset = ordered[index + 1]?.localOffset ?? centralDirectoryOffset;
    if (entry.dataEnd > nextOffset) fail('локальные записи или сжатые данные ZIP перекрываются.');
    const nameStart = entry.localOffset + 30;
    if (entry._localNameBytes !== entry._nameBytes) fail('имя локального файла ZIP не совпадает с каталогом.');
    for (let byte = 0; byte < entry._nameBytes; byte += 1) if (bytes[nameStart + byte] !== bytes[entry._nameStart + byte]) fail('имя локального файла ZIP не совпадает с каталогом.');
    extraFields(nameStart + entry._localNameBytes, entry._localExtraBytes);
    const localCRC = u32(entry.localOffset + 14);
    const localCompressed = u32(entry.localOffset + 18);
    const localUncompressed = u32(entry.localOffset + 22);
    if (localCompressed === 0xffffffff || localUncompressed === 0xffffffff) fail('ZIP64 не поддерживается.');
    if (entry.flags & 0x0008) {
      if ((localCRC !== 0 && localCRC !== entry._crc) || (localCompressed !== 0 && localCompressed !== entry.compressedSize) || (localUncompressed !== 0 && localUncompressed !== entry.uncompressedSize)) fail('размеры локальной записи ZIP не совпадают с каталогом.');
      const descriptorBytes = nextOffset - entry.dataEnd;
      let descriptorStart = entry.dataEnd;
      if (descriptorBytes === 16 && u32(descriptorStart) === DESCRIPTOR) descriptorStart += 4;
      else if (descriptorBytes !== 12) fail('дескриптор сжатых данных ZIP отсутствует или повреждён.');
      if (u32(descriptorStart) !== entry._crc || u32(descriptorStart + 4) !== entry.compressedSize || u32(descriptorStart + 8) !== entry.uncompressedSize) fail('дескриптор данных ZIP не совпадает с каталогом.');
    } else {
      if (localCRC !== entry._crc || localCompressed !== entry.compressedSize || localUncompressed !== entry.uncompressedSize) fail('контрольные поля или размеры локальной записи ZIP не совпадают с каталогом.');
      if (entry.dataEnd !== nextOffset) fail('между локальными записями ZIP найдены неподдерживаемые или потерянные данные.');
    }
  }
  return {
    fileBytes: size, entryCount, totalUncompressedBytes, centralDirectoryOffset, centralDirectoryBytes,
    commentBytes: u16(eocd + 20),
    entries: entries.map(({ _version, _crc, _nameStart, _nameBytes, _localNameBytes, _localExtraBytes, ...entry }) => entry),
  };
}
