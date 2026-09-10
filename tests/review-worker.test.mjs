import test from 'node:test'
import assert from 'node:assert/strict'
import {readFileSync} from 'node:fs'
import vm from 'node:vm'
import {Worker} from 'node:worker_threads'
import {auditSnapshotFromAoa} from '../src/audit/model.js'
import {auditMakeCorrection,auditApplyCorrections} from '../src/audit/actions.js'
import {auditSessionResult} from '../src/audit/review.js'
import {EXTO_REV21_COLUMNS} from '../src/exto/rev21-contract.js'
const html=readFileSync(new URL('../SSM-Audit.html',import.meta.url),'utf8');
const vendor=html.match(/<script id="sheetjs-runtime">([\s\S]*?)<\/script>/)[1];
const source=JSON.parse(html.match(/const AUDIT_IMPORT_WORKER_SOURCE=("(?:[^"\\]|\\.)*");/)[1]);
const runtime=vm.createContext({});vm.runInContext(vendor,runtime);const XLSX=runtime.XLSX;
test('packaged review worker validates large drafts off-thread and reuses its source workbook',async()=>{
 const count=Number(process.env.SSM_REVIEW_BENCH_ROWS)||5000,aoa=[EXTO_REV21_COLUMNS.map(c=>c.header)];
 for(let i=0;i<count;i++){const r={equipmentId:`DEMO-PANEL-${i}`,equipmentDescription:'Electrical panel',building:'DEMO',upn:'602',systemName:'602  Medium Voltage',discipline:'ELECTRICAL',closestParent:'602  Medium Voltage',milestoneParent:'DEMO-L1-M1-10',milestone:'DEMO-L2-M1-20 UPN 602',dependencyProject:'Old project'};aoa.push(EXTO_REV21_COLUMNS.map(c=>r[c.field]||''));}
 const baseline=auditSnapshotFromAoa(aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
 const file=new Blob([XLSX.write(book,{type:'array',bookType:'xlsx'})]);
 const worker=new Worker(`const {parentPort,workerData}=require('node:worker_threads'),vm=require('node:vm');const context=vm.createContext({TextEncoder,TextDecoder,Uint8Array,setTimeout,clearTimeout,self:{postMessage:data=>parentPort.postMessage(data)}});vm.runInContext(workerData,context);let reads=0;const read=context.XLSX.read;context.XLSX.read=(...args)=>{reads++;return read(...args);};parentPort.on('message',async data=>{await context.self.onmessage({data});parentPort.postMessage({type:'reads',reads});});`,{eval:true,workerData:vendor+'\n'+source});
 let ticks=0;const timer=setInterval(()=>ticks++,10),started=performance.now();
 const send=data=>new Promise((resolve,reject)=>{let prepared;const onError=error=>{worker.off('message',onMessage);reject(error);},onMessage=message=>{if(message.type==='error'){worker.off('message',onMessage);worker.off('error',onError);reject(new Error(message.message));}if(message.type==='result')prepared=message.prepared;if(message.type==='reads'){worker.off('message',onMessage);worker.off('error',onError);resolve({prepared,reads:message.reads});}};worker.on('message',onMessage);worker.once('error',onError);worker.postMessage({kind:'review',references:{},migration:{enabled:false,profile:null},migrationChanged:false,...data});});
 try{
  const initialResult=auditSessionResult(baseline);
  const actionsExport=await send({kind:'actions-export',baseline,file,actionsWorkbook:{result:{...initialResult,findings:initialResult.findings.slice(0,100)},sessionName:'Demo',options:{}}});
  assert.equal(actionsExport.reads,0,'Actions export needs no source-workbook parsing');
  const actionsBook=XLSX.read(actionsExport.prepared.bytes,{type:'array',cellStyles:true});
  assert.equal(actionsBook.SheetNames[0],'Actionable');assert.ok(actionsBook.SheetNames.length>1);
  const change=auditMakeCorrection(baseline.rows[0],'Dependency Project','');
  const first=await send({previousChanges:[],changes:[change]});
  assert.equal(first.reads,1);assert.equal(first.prepared.snapshot.rows[0].dependencyProject,'');assert.equal(first.prepared.exportCheck.cellCount,1);
  const expected=auditSessionResult(auditApplyCorrections(baseline,[change]));assert.deepEqual(first.prepared.result,expected);
  const changes=[change,auditMakeCorrection(baseline.rows[1],'Dependency Project','')];
  const second=await send({previousChanges:[change],changes});assert.equal(second.reads,1,'original workbook was not reparsed');assert.equal(second.prepared.exportCheck.cellCount,2);assert.equal(second.prepared.snapshot.rows[1].dependencyProject,'');
  const exported=await send({kind:'export',changes,completedEquipmentIds:[baseline.rows[0].equipmentId]});
  assert.equal(exported.reads,1,'export reuses the validated source workbook');
  const output=XLSX.read(exported.prepared.bytes,{type:'array',cellStyles:true}),sheet=output.Sheets['Upload Template'];
  assert.deepEqual([...output.SheetNames],['Upload Template']);assert.equal(sheet.K2.v,baseline.rows[1].equipmentId);assert.equal(sheet.AO2.v,'');assert.equal(sheet.AO2.s.fgColor.rgb,'FFF2CC');
  assert.equal(XLSX.utils.decode_range(sheet['!ref']).e.r,1);
  assert.deepEqual(Array.from(XLSX.utils.sheet_to_json(sheet,{header:1})[0]),EXTO_REV21_COLUMNS.map(c=>c.header));
  assert.ok(ticks>0,'the calling thread remains responsive while the worker computes');
  console.log(JSON.stringify({reviewBenchmark:{rows:count,twoReviewsMs:Math.round(performance.now()-started),responsiveTicks:ticks,workbookReads:second.reads}}));
 }finally{clearInterval(timer);await worker.terminate();}
});
