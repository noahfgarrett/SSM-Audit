/* ---- Minimal ZIP writer with real deflate ----
   SheetJS bundles its own (weak) compressor: a large audit workbook comes out
   roughly 3x larger than the same XML zipped by zlib. Browsers expose zlib-grade
   deflate through CompressionStream, so the workbook's entries are re-zipped
   here. Output is a plain ZIP (method 8, UTF-8 names, no data descriptors),
   which is all an .xlsx is. Callers fall back to SheetJS's own bytes when
   CompressionStream is unavailable. */

const CRC_TABLE=(()=>{const table=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;table[n]=c>>>0;}return table;})();
export function crc32(bytes){let crc=0xFFFFFFFF;for(let i=0;i<bytes.length;i++)crc=CRC_TABLE[(crc^bytes[i])&0xFF]^(crc>>>8);return (crc^0xFFFFFFFF)>>>0;}

export function zipDeflateAvailable(){return typeof CompressionStream==='function'&&typeof Response==='function';}

async function deflateRaw(bytes){
  const stream=new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/* MS-DOS time/date fields: a fixed, valid stamp keeps output deterministic. */
const DOS_TIME=0,DOS_DATE=(1<<5)|1|((2026-1980)<<9);
const FLAG_UTF8=0x0800;

/* entries: [{name, data:Uint8Array}] → Uint8Array of a ZIP file. */
export async function zipEntries(entries,onProgress){
  const encoder=new TextEncoder(),locals=[],centrals=[];let offset=0;
  const totalBytes=entries.reduce((sum,entry)=>sum+entry.data.length,0)||1;let doneBytes=0;
  for(const entry of entries){
    const name=encoder.encode(entry.name),data=entry.data,crc=crc32(data),packed=await deflateRaw(data);
    const local=new Uint8Array(30+name.length),view=new DataView(local.buffer);
    view.setUint32(0,0x04034b50,true);view.setUint16(4,20,true);view.setUint16(6,FLAG_UTF8,true);view.setUint16(8,8,true);
    view.setUint16(10,DOS_TIME,true);view.setUint16(12,DOS_DATE,true);view.setUint32(14,crc,true);view.setUint32(18,packed.length,true);view.setUint32(22,data.length,true);
    view.setUint16(26,name.length,true);view.setUint16(28,0,true);local.set(name,30);
    const central=new Uint8Array(46+name.length),cview=new DataView(central.buffer);
    cview.setUint32(0,0x02014b50,true);cview.setUint16(4,20,true);cview.setUint16(6,20,true);cview.setUint16(8,FLAG_UTF8,true);cview.setUint16(10,8,true);
    cview.setUint16(12,DOS_TIME,true);cview.setUint16(14,DOS_DATE,true);cview.setUint32(16,crc,true);cview.setUint32(20,packed.length,true);cview.setUint32(24,data.length,true);
    cview.setUint16(28,name.length,true);cview.setUint16(30,0,true);cview.setUint16(32,0,true);cview.setUint16(34,0,true);cview.setUint16(36,0,true);cview.setUint32(38,0,true);cview.setUint32(42,offset,true);central.set(name,46);
    locals.push(local,packed);centrals.push(central);offset+=local.length+packed.length;
    doneBytes+=entry.data.length;if(onProgress)await onProgress(doneBytes/totalBytes,entry.name);
  }
  const centralSize=centrals.reduce((sum,part)=>sum+part.length,0);
  const end=new Uint8Array(22),eview=new DataView(end.buffer);
  eview.setUint32(0,0x06054b50,true);eview.setUint16(4,0,true);eview.setUint16(6,0,true);eview.setUint16(8,entries.length,true);eview.setUint16(10,entries.length,true);eview.setUint32(12,centralSize,true);eview.setUint32(16,offset,true);eview.setUint16(20,0,true);
  const total=offset+centralSize+end.length,out=new Uint8Array(total);let at=0;
  for(const part of [...locals,...centrals,end]){out.set(part,at);at+=part.length;}
  return out;
}
