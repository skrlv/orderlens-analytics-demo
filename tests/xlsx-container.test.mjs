import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { validateXlsxContainer, XLSX_CONTAINER_LIMITS } from '../src/xlsx-container.mjs';

const sample = fs.readFileSync(new URL('../public/samples/orders-parts.xlsx', import.meta.url));
const truncated = fs.readFileSync(new URL('./fixtures/truncated-100.xlsx', import.meta.url));
const u16 = (b,o,v) => b.writeUInt16LE(v,o);
const u32 = (b,o,v) => b.writeUInt32LE(v>>>0,o);
const required = () => [{name:'[Content_Types].xml',data:'<Types/>'},{name:'xl/workbook.xml',data:'<workbook/>'}];

// Minimal synthetic ZIP metadata. Test payloads are never decompressed or
// parsed as workbooks. CRC fields intentionally test structural equality only.
function zip(input=required(),options={}) {
  const chunks=[];const entries=[];let offset=0;
  for(const spec of input){
    const name=Buffer.from(spec.name,'utf8');const data=Buffer.from(spec.data??'x');
    const extra=Buffer.from(spec.extra??[]);const flags=spec.flags??0;const method=spec.method??0;
    const uncompressed=spec.uncompressed??data.length;const crc=0x12345678;
    const local=Buffer.alloc(30);u32(local,0,0x04034b50);u16(local,4,20);u16(local,6,flags);u16(local,8,method);
    if(!(flags&8)){u32(local,14,crc);u32(local,18,data.length);u32(local,22,uncompressed);}
    u16(local,26,name.length);u16(local,28,extra.length);
    const entry={spec,name,data,extra,flags,method,uncompressed,crc,localOffset:offset};
    chunks.push(local,name,extra,data);offset+=local.length+name.length+extra.length+data.length;
    if(flags&8){const signed=spec.signedDescriptor!==false;const descriptor=Buffer.alloc(signed?16:12);let d=0;if(signed){u32(descriptor,0,0x08074b50);d=4;}u32(descriptor,d,crc);u32(descriptor,d+4,data.length);u32(descriptor,d+8,uncompressed);chunks.push(descriptor);offset+=descriptor.length;}
    entries.push(entry);
  }
  const cdOffset=offset;const centralOffsets=[];
  const ordered=options.reverseCentral?entries.toReversed():entries;
  for(const entry of ordered){
    const central=Buffer.alloc(46);const extra=Buffer.from(entry.spec.centralExtra??entry.extra);
    u32(central,0,0x02014b50);u16(central,4,20);u16(central,6,20);u16(central,8,entry.flags);u16(central,10,entry.method);
    u32(central,16,entry.crc);u32(central,20,entry.data.length);u32(central,24,entry.uncompressed);
    u16(central,28,entry.name.length);u16(central,30,extra.length);u32(central,42,entry.localOffset);
    centralOffsets.push(offset);chunks.push(central,entry.name,extra);offset+=46+entry.name.length+extra.length;
  }
  const eocdOffset=offset;const eocd=Buffer.alloc(22);const comment=Buffer.from(options.comment??'');
  u32(eocd,0,0x06054b50);u16(eocd,8,entries.length);u16(eocd,10,entries.length);
  u32(eocd,12,eocdOffset-cdOffset);u32(eocd,16,cdOffset);u16(eocd,20,comment.length);
  chunks.push(eocd,comment);
  return {bytes:Buffer.concat(chunks),entries,centralOffsets,cdOffset,eocdOffset};
}
function reject(bytes,pattern=/XLSX:/){assert.throws(()=>validateXlsxContainer(bytes),pattern);}
function changed(original,change){const bytes=Buffer.from(original.bytes);change(bytes,original);return bytes;}

test('real 4437-byte sample passes with exact bounded container metadata',()=>{
  const metadata=validateXlsxContainer(sample);
  assert.equal(metadata.fileBytes,4437);assert.equal(metadata.entryCount,8);
  assert.equal(metadata.totalUncompressedBytes,16001);assert.equal(metadata.commentBytes,0);
  assert.equal(metadata.centralDirectoryOffset,3900);assert.equal(metadata.centralDirectoryBytes,515);
  assert.ok(metadata.entries.some(e=>e.name==='[Content_Types].xml'));
  assert.ok(metadata.entries.some(e=>e.name==='xl/workbook.xml'));
  for(const e of metadata.entries)assert.ok(e.dataOffset<e.dataEnd&&e.dataEnd<=metadata.centralDirectoryOffset);
});
test('accepts ArrayBuffer, byte views with offsets and Node Buffer without mutation',()=>{
  const before=Buffer.from(sample);const padded=Buffer.concat([Buffer.from('prefix'),sample,Buffer.from('suffix')]);
  const view=new Uint8Array(padded.buffer,padded.byteOffset+6,sample.length);
  const dataView=new DataView(padded.buffer,padded.byteOffset+6,sample.length);
  const arrayBuffer=sample.buffer.slice(sample.byteOffset,sample.byteOffset+sample.byteLength);
  for(const input of [sample,view,dataView,arrayBuffer])assert.equal(validateXlsxContainer(input).entryCount,8);
  assert.deepEqual(sample,before);reject('text');reject(null);
});
test('exact reviewer 100-byte truncation fails before any workbook parser',()=>{
  assert.equal(truncated.length,100);assert.equal(truncated.readUInt32LE(0),0x04034b50);
  assert.deepEqual(truncated,sample.subarray(0,100));
  reject(truncated,/обрезан|повреждён/);
});
test('diverse truncations including missing EOCD or partial comment fail promptly',()=>{
  for(const length of [0,1,4,21,29,30,52,99,100,101,500,3899,3900,3910,4415,4436])reject(sample.subarray(0,length));
  const commented=zip(required(),{comment:'Коммент ZIP'});
  reject(commented.bytes.subarray(0,-1));
});
test('ZIP comments, maximum-length comments and reversed central order are supported',()=>{
  for(const comment of ['comment PK\x05\x06 not an EOCD',Buffer.alloc(65535,65)]){
    const built=zip(required(),{comment,reverseCentral:true});const m=validateXlsxContainer(built.bytes);
    assert.equal(m.entryCount,2);assert.equal(m.commentBytes,Buffer.byteLength(comment));
    assert.equal(m.entries[0].name,'xl/workbook.xml');
  }
});
test('EOCD signature, comment length, trailing bytes and directory bounds are checked',()=>{
  const z=zip();
  reject(changed(z,(b,z)=>u32(b,z.eocdOffset,0)));
  reject(changed(z,(b,z)=>u16(b,z.eocdOffset+20,1)));
  reject(Buffer.concat([z.bytes,Buffer.from('trailing junk')]));
  reject(changed(z,(b,z)=>u32(b,z.eocdOffset+16,z.cdOffset+1)));
  reject(changed(z,(b,z)=>u32(b,z.eocdOffset+12,0xffffffff)),/ZIP64/);
  reject(changed(z,(b,z)=>u32(b,z.eocdOffset+16,0xfffffffe)));
});
test('exact central record count and entry cap are enforced',()=>{
  const z=zip();
  for(const count of [0,1,3,1025])reject(changed(z,(b,z)=>{u16(b,z.eocdOffset+8,count);u16(b,z.eocdOffset+10,count);}));
  reject(changed(z,(b,z)=>u16(b,z.eocdOffset+10,0xffff)),/ZIP64/);
  reject(changed(z,(b,z)=>u32(b,z.centralOffsets[0],0)));
  reject(changed(z,(b,z)=>u16(b,z.centralOffsets[0]+28,0xffff)));
});
test('actual archives accept exactly 1024 entries and reject the 1025th',()=>{
  const entries=[...required(),...Array.from({length:1022},(_,i)=>({name:`xl/extra/part-${i}.xml`}))];
  assert.equal(validateXlsxContainer(zip(entries).bytes).entryCount,1024);
  reject(zip([...entries,{name:'xl/extra/too-many.xml'}]).bytes,/1024/);
});
test('multi-disk, encryption, ZIP64 fields/extras and unsupported methods are rejected',()=>{
  const z=zip();
  for(const [field,value]of [[4,1],[6,1],[8,1]])reject(changed(z,(b,z)=>u16(b,z.eocdOffset+field,value)),/многотом/);
  reject(changed(z,(b,z)=>u16(b,z.centralOffsets[0]+34,1)),/многотом/);
  for(const flag of [1,0x40,0x2000])reject(changed(z,(b,z)=>u16(b,z.centralOffsets[0]+8,flag)),/зашифрован/);
  for(const field of [20,24,42])reject(changed(z,(b,z)=>u32(b,z.centralOffsets[0]+field,0xffffffff)),/ZIP64/);
  reject(changed(z,(b,z)=>u16(b,z.centralOffsets[0]+6,45)),/ZIP64/);
  reject(zip([{...required()[0],extra:[1,0,0,0]},required()[1]]).bytes,/ZIP64/);
  reject(zip([{...required()[0],extra:[1,0x99,0,0]},required()[1]]).bytes,/зашифрован/);
  reject(changed(z,(b,z)=>u16(b,z.centralOffsets[0]+10,12)),/Deflate/);
  reject(changed(z,(b,z)=>u16(b,z.centralOffsets[0]+8,0x10)),/флаги/);
});
test('local signature, header lengths, offsets, names and methods match the directory',()=>{
  const z=zip();const first=z.centralOffsets[0];
  reject(changed(z,(b)=>u32(b,0,0)));
  reject(changed(z,(b)=>u16(b,26,0xffff)));
  reject(changed(z,(b)=>u16(b,28,0xffff)));
  reject(changed(z,(b)=>u16(b,8,8)),/заголовок/);
  reject(changed(z,(b)=>u16(b,6,8)),/заголовок/);
  reject(changed(z,(b)=>{b[30]=0x58;}),/имя/);
  reject(changed(z,(b)=>u32(b,first+42,z.cdOffset)),/локальный/);
  reject(changed(z,(b)=>u32(b,z.centralOffsets[1]+42,z.entries[0].localOffset)),/дублирующ/);
});
test('compressed data bounds, local/central sizes and overlap are checked',()=>{
  const z=zip();
  reject(changed(z,(b,z)=>{u16(b,8,8);u16(b,z.centralOffsets[0]+10,8);u32(b,z.centralOffsets[0]+20,0xfffffff0);u32(b,18,0xfffffff0);}));
  reject(changed(z,b=>u32(b,18,0)),/размеры/);
  reject(changed(z,b=>u32(b,22,0)),/размеры/);
  reject(changed(z,b=>u32(b,14,0)),/контрольные/);
  reject(changed(z,b=>u32(b,18,0xffffffff)),/ZIP64/);
  // Move the second local header into the first compressed payload.
  reject(changed(z,(b,z)=>u32(b,z.centralOffsets[1]+42,z.entries[1].localOffset-1)));
});
test('missing required OOXML parts, duplicates and unsafe paths are rejected',()=>{
  reject(zip([{name:'arbitrary.txt'}]).bytes,/обязательная часть/);
  reject(zip([required()[0]]).bytes,/xl\/workbook.xml/);
  reject(zip([...required(),required()[0]]).bytes,/дублирующ/);
  for(const name of ['../bad.xml','/bad.xml','xl\\bad.xml','C:/bad.xml','xl/./bad.xml','bad\u0000.xml'])reject(zip([...required(),{name}]).bytes,/имя|путь/);
});
test('uncompressed limits are checked from metadata without allocating the declared payload',()=>{
  const MiB=1024*1024;
  reject(zip([{...required()[0],method:8,uncompressed:16*MiB+1},required()[1]]).bytes,/16 МБ/);
  reject(zip([...required(),{name:'a',method:8,uncompressed:16*MiB},{name:'b',method:8,uncompressed:16*MiB}]).bytes,/32 МБ/);
  reject(Buffer.alloc(XLSX_CONTAINER_LIMITS.fileBytes+1),/5 МБ/);
  const boundary=zip([{...required()[0],method:8,uncompressed:16*MiB},{...required()[1],method:8,uncompressed:16*MiB}]);
  assert.equal(validateXlsxContainer(boundary.bytes).totalUncompressedBytes,32*MiB);
});
test('data descriptors with and without signatures are bounded and consistent',()=>{
  for(const signedDescriptor of [true,false]){
    const z=zip([{...required()[0],flags:8,signedDescriptor},required()[1]]);
    assert.equal(validateXlsxContainer(z.bytes).entryCount,2);
    const descriptor=z.entries[1].localOffset-(signedDescriptor?16:12);
    reject(changed(z,b=>u32(b,descriptor+(signedDescriptor?8:4),123)),/дескриптор/);
  }
});
test('malformed extra fields and UTF-8 paths reject safely',()=>{
  reject(zip([{...required()[0],extra:[1]},required()[1]]).bytes,/дополнительн/);
  reject(zip([{...required()[0],extra:[2,0,10,0]},required()[1]]).bytes,/дополнительн/);
  const z=zip([{...required()[0],flags:0x800},required()[1]]);
  const corrupted=changed(z,(b,z)=>{b[z.centralOffsets[0]+46]=0xff;});
  reject(corrupted,/UTF-8/);
});
