import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import vm from 'node:vm'
import { Worker as Thread } from 'node:worker_threads'
import { auditSnapshotFromWorkbook } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { applyCompletedEquipment } from '../src/ui/audit.js'

const root=new URL('../',import.meta.url);
const html=readFileSync(new URL('SSM-Audit.html',root),'utf8');
const vendor=html.match(/<script id="sheetjs-runtime">([\s\S]*?)<\/script>/)[1];
const workerSource=JSON.parse(html.match(/const AUDIT_IMPORT_WORKER_SOURCE=("(?:[^"\\]|\\.)*");/)[1]);
const runtime=vm.createContext({});vm.runInContext(vendor,runtime);
globalThis.XLSX=runtime.XLSX;
const client=readFileSync(new URL('src/io/import-client.js',root),'utf8').replaceAll('export function','function');

function fixture(count=32,withStatus=true,bookType='xlsx'){
  const rows=[EXTO_REV21_COLUMNS.map(column=>column.header)];
  for(let i=0;i<count;i++){
    const values={equipmentId:`TEST-PUMP-${i}`,equipmentDescription:'Pump',closestParent:i?'TEST-PUMP-0':'',upn:'111',discipline:'MECHANICAL WET',building:'TEST',dependencies:i?'TEST-PUMP-0':''};
    rows.push(EXTO_REV21_COLUMNS.map(column=>values[column.field]||''));
  }
  const workbook=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet(rows),'Full Export');
  if(withStatus)XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([
    ['Equipment Name','RR OA/BT','EQ OA/BT'],
    ...Array.from({length:count},(_,i)=>[`TEST-PUMP-${i}`,i%2?'Not Started':'Completed','']),
    ['UNMATCHED','Completed',''],['TEST-PUMP-1','Completed','Not Started'],
  ]),'Equipment Status Report');
  return new Uint8Array(XLSX.write(workbook,{type:'array',bookType}));
}

// Run the exact packaged worker in an isolated JS thread with browser globals.
// This tests offline processing, not browser UI or URL-policy behavior.
async function background(bytes,audit=true){
  const thread=new Thread(`const {parentPort,workerData}=require('node:worker_threads');
    const vm=require('node:vm');
    const context=vm.createContext({TextEncoder,Uint8Array,self:{postMessage:(data,transfer)=>parentPort.postMessage(data,transfer)}});
    vm.runInContext(workerData,context);
    parentPort.on('message',data=>context.self.onmessage({data}));`,{eval:true,workerData:vendor+'\n'+workerSource});
  const messages=[];
  try{
    return await new Promise((resolve,reject)=>{
      thread.on('error',reject);
      thread.on('message',data=>{messages.push(data);if(data.type==='result'||data.type==='error')resolve({data,messages});});
      thread.postMessage({file:new Blob([bytes]),fileName:'synthetic.xlsx',audit});
    });
  }finally{await thread.terminate();}
}

for(const withStatus of [false,true])test(`background import matches synchronous audit with status=${withStatus}`,async()=>{
  const bytes=fixture(32,withStatus);
  const expected=await auditSnapshotFromWorkbook(XLSX.read(bytes,{type:'array',dense:true}),'synthetic.xlsx');
  const {data,messages}=await background(bytes);
  assert.equal(data.type,'result');
  assert.deepEqual(data.snapshot,expected);
  assert.deepEqual(data.rawResult,runSsmAudit(expected));
  assert.deepEqual(data.bytes,new Uint8Array(bytes),'original workbook bytes survive intact');
  if(withStatus){
    assert.equal(data.status.matched,16);assert.equal(data.status.completed.size,17);
    assert.ok(data.status.completed instanceof Set);
    const filtered=applyCompletedEquipment(data.rawResult,data.status.completed);
    assert.ok(filtered.findings.every(f=>!f.equipmentId||!data.status.completed.has(f.equipmentId)));
    assert.equal(data.snapshot.rows.length,32,'completed parents remain in the model');
  }else assert.equal(data.status,null);
  const progress=messages.filter(message=>message.type==='progress');
  assert.equal(progress[0].fraction,.02);
  assert.ok(progress.every((value,index)=>!index||value.fraction>=progress[index-1].fraction));
});

test('comparison reference skips audit and does not return workbook bytes',async()=>{
  const {data}=await background(fixture(),false);
  assert.equal(data.type,'result');assert.equal(data.rawResult,null);assert.equal(data.status,null);assert.equal(data.bytes,null);
  assert.equal(data.snapshot.rows.length,32);
});

test('bad workbook produces an actionable error instead of a pending import',async()=>{
  const {data}=await background(new Uint8Array([1,2,3]));
  assert.equal(data.type,'error');assert.ok(data.message);
});

test('legacy XLS workbooks use the same offline worker and audit rules',async()=>{
  const bytes=fixture(12,true,'biff8');
  const expected=await auditSnapshotFromWorkbook(XLSX.read(bytes,{type:'array',dense:true}),'synthetic.xlsx');
  const {data}=await background(bytes);
  assert.equal(data.type,'result');assert.deepEqual(data.rawResult,runSsmAudit(expected));
  assert.equal(data.status.matched,6);
});

test('multiple registry tabs retain physical source rows and duplicate evidence',async()=>{
  const workbook=XLSX.read(fixture(12,false),{type:'array',dense:true});
  const rows=XLSX.utils.sheet_to_json(workbook.Sheets['Full Export'],{header:1});
  XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([[],...rows,rows[1]]),'Registry Copy');
  const bytes=new Uint8Array(XLSX.write(workbook,{type:'array',bookType:'xlsx'}));
  const expected=await auditSnapshotFromWorkbook(XLSX.read(bytes,{type:'array',dense:true}),'synthetic.xlsx');
  const {data}=await background(bytes);
  assert.equal(data.type,'result');assert.deepEqual(data.snapshot,expected);assert.deepEqual(data.rawResult,runSsmAudit(expected));
  assert.equal(data.snapshot.rows.length,13);assert.equal(data.snapshot.source.ignoredDuplicateRows,12);
});

test('large background import leaves the controlling event loop responsive',async(t)=>{
  const count=Number(process.env.SSM_IMPORT_STRESS_ROWS)||2000,bytes=fixture(count);
  let ticks=0,maxGap=0,last=performance.now();
  const started=last,timer=setInterval(()=>{const now=performance.now();maxGap=Math.max(maxGap,now-last);last=now;ticks++;},10);
  let imported;
  try{imported=await background(bytes);}finally{clearInterval(timer);}
  assert.equal(imported.data.type,'result');assert.equal(imported.data.snapshot.rows.length,count);
  assert.equal(imported.data.status.matched,Math.ceil(count/2));
  assert.ok(ticks>=3,'timer continues while workbook processing runs');
  t.diagnostic(`${count} synthetic rows: ${Math.round(performance.now()-started)} ms, ${ticks} timer ticks, largest timer gap ${Math.round(maxGap)} ms`);
});

function clientHarness(mode){
  const events=[],callbacks=[];
  class FakeWorker{
    constructor(){if(mode==='construct')throw new Error('blocked');events.push('created');callbacks.push(this);}
    postMessage(data){events.push(data);if(mode==='post')throw new Error('cannot clone');}
    terminate(){events.push('terminated');}
  }
  const context=vm.createContext({Worker:mode==='unavailable'?undefined:FakeWorker,Blob,AUDIT_IMPORT_WORKER_SOURCE:workerSource,
    document:{getElementById:()=>({textContent:vendor})},URL:{createObjectURL:()=>{events.push('url');return 'blob:local';},revokeObjectURL:()=>events.push('revoked')}});
  vm.runInContext(client,context);
  return {run:context.importAuditWorkbook,review:context.prepareAuditReview,events,callbacks};
}

test('client forwards progress, returns results and frees worker resources once',async()=>{
  const h=clientHarness(),progress=[],pending=h.run({name:'synthetic.xlsx'},{report:(...args)=>progress.push(args)}),worker=h.callbacks[0];
  worker.onmessage({data:{type:'progress',fraction:.2,label:'Reading'}});
  worker.onmessage({data:{type:'result',snapshot:{rows:[]}}});
  worker.onmessage({data:{type:'result'}});
  assert.deepEqual(progress,[[.2,'Reading']]);assert.ok((await pending).snapshot);
  assert.equal(h.events.filter(event=>event==='terminated').length,1);assert.equal(h.events.filter(event=>event==='revoked').length,1);
});

test('review client reuses the worker and releases it with the registry session',async()=>{
 const h=clientHarness(),session={baselineSnapshot:{rows:[]},sourceBytes:new Uint8Array([80,75]),changes:[],references:{}};
 const first=h.review(session,[],{},false),worker=h.callbacks[0];worker.onmessage({data:{type:'result',prepared:{snapshot:{rows:[]}}}});await first;
 const second=h.review(session,[],{},false);assert.equal(h.callbacks.length,1);const posted=h.events.filter(e=>e?.kind==='review');assert.ok(posted[0].baseline);assert.equal(posted[1].baseline,undefined);
 const competing=h.review(session,[],{},false);await assert.rejects(competing,/current review/);
 session.disposeReviewWorker();await assert.rejects(second,/registry changed/);
 assert.equal(h.events.filter(e=>e==='terminated').length,1);assert.equal(h.events.filter(e=>e==='revoked').length,1);assert.equal(session.reviewWorker,null);
});

for(const mode of ['construct','post','unavailable','error','messageerror','reported'])test(`client safely rejects ${mode} failures`,async()=>{
  const h=clientHarness(mode),pending=h.run({name:'synthetic.xlsx'});
  if(mode==='error')h.callbacks[0].onerror({});
  if(mode==='messageerror')h.callbacks[0].onmessageerror();
  if(mode==='reported')h.callbacks[0].onmessage({data:{type:'error',message:'Invalid registry'}});
  await assert.rejects(pending);
  if(mode!=='unavailable')assert.ok(h.events.includes('revoked'));
  if(!['construct','unavailable'].includes(mode))assert.ok(h.events.includes('terminated'));
});
