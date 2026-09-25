import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import {readFileSync} from 'node:fs'
import {clean,esc,natCmp} from '../src/core/text.js'
import {auditNormId} from '../src/audit/model.js'

const source=readFileSync(new URL('../src/audit/export.js',import.meta.url),'utf8');
const start=source.indexOf('export async function exportRemainingRegistryXlsx()');
const end=source.indexOf('export async function exportUpdatedRegistryXlsx()',start);
assert.ok(start>=0&&end>start);
const code=source.slice(start,end).replace('export async function','async function');
function harness(overrides={}){
 const session={sourceBytes:new Uint8Array([80,75]),baselineSnapshot:{rows:[]},changes:[],changesRev:0,status:{completed:new Set(['DEMO-DONE'])},completedSearch:'hide everything',...overrides};
 const downloads=[],messages=[],requests=[];
 const context=vm.createContext({S:{session},Blob,toast:message=>messages.push(message),downloadBlob:(name,blob)=>downloads.push({name,blob}),runWithProgress:async(title,description,work)=>work(async()=>{},()=>{}),prepareAuditReview:async(...args)=>{
  requests.push(args);return {bytes:new Uint8Array([80,75]),filename:'Registry_Remaining.xlsx',summary:{exportedRows:2,excludedCompletedRows:1}};
 }});
 vm.runInContext(code,context);
 return {session,context,downloads,messages,requests,run:()=>context.exportRemainingRegistryXlsx()};
}

test('remaining registry download works without edits and does not use display filters',async()=>{
 const h=harness();assert.equal(await h.run(),true);
 assert.equal(h.requests.length,1);assert.equal(h.requests[0][0],h.session);
 assert.deepEqual({...h.requests[0][5]},{export:true,remainingRegistry:true});
 assert.equal(h.downloads.length,1);assert.equal(h.downloads[0].blob.type,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
 assert.match(h.messages[0],/2 remaining equipment rows exported; 1 completed rows removed/);
 assert.equal(h.session.remainingExportBusy,false);
});

test('remaining registry export requires status and refuses concurrent operations',async()=>{
 for(const overrides of [{status:null},{reviewBusy:true},{remainingExportBusy:true},{sourceBytes:null}]){
  const h=harness(overrides);assert.equal(await h.run(),false);assert.equal(h.requests.length,0);assert.equal(h.downloads.length,0);assert.equal(h.messages.length,1);
 }
});

test('remaining registry never downloads stale results and clears its busy state on failure',async()=>{
 for(const mode of ['revision','session','status','error']){
  const h=harness();
  h.context.prepareAuditReview=async()=>{
   if(mode==='error')throw new Error('No incomplete equipment remains to export.');
   if(mode==='revision')h.session.changesRev++;
   if(mode==='session')h.context.S.session={};
   if(mode==='status')h.session.status={completed:new Set()};
   return {bytes:new Uint8Array([80,75]),filename:'Never.xlsx'};
  };
  assert.equal(await h.run(),false);assert.equal(h.downloads.length,0);assert.equal(h.session.remainingExportBusy,false);
  assert.match(h.messages[0],mode==='error'?/No incomplete equipment/:/changed while exporting/);
 }
});

test('review client forwards remaining-registry mode and the full completion set to the worker',async()=>{
 const client=readFileSync(new URL('../src/io/import-client.js',import.meta.url),'utf8').replaceAll('export function','function');
 let sent;
 const session={status:{completed:new Set(['DEMO-DONE','DEMO-ALSO-DONE'])},changes:[]};
 session.reviewWorker={pending:null,worker:{postMessage:data=>{sent=data;session.reviewWorker.pending.resolve({ok:true});session.reviewWorker.pending=null;}}};
 const context=vm.createContext({Blob});vm.runInContext(client,context);
 await context.prepareAuditReview(session,[],null,false,()=>{},{export:true,remainingRegistry:true});
 assert.equal(sent.kind,'export');assert.equal(sent.remainingRegistry,true);
 assert.deepEqual(Array.from(sent.completedEquipmentIds),['DEMO-DONE','DEMO-ALSO-DONE']);
});

test('Completed Equipment offers a guarded export button and the guide documents retained tags',()=>{
 const ui=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
 const guide=readFileSync(new URL('../src/ui/guide-content.js',import.meta.url),'utf8');
 const completed=ui.slice(ui.indexOf('export function renderCompletedEquipment('),ui.indexOf('function rerenderModifications('));
 assert.match(completed,/id="exportRemainingRegistry"/);
 assert.match(completed,/button\.disabled=true;[\s\S]*await exportRemainingRegistryXlsx\(\);[\s\S]*button\.disabled=false/);
 assert.match(guide,/In-progress, incomplete and unmatched tags remain/);
 assert.match(guide,/Display filters do not limit this export/);
});

test('completed screen lists report members without steps and ignores obsolete BT filters',()=>{
 const ui=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
 const code=ui.slice(ui.indexOf('const COMPLETED_CHUNK='),ui.indexOf('function rerenderModifications(')).replace('export function','function');
 const nodes=new Map(),node=selector=>{
  if(!nodes.has(selector))nodes.set(selector,{innerHTML:'',setAttribute(){},focus(){}});
  return nodes.get(selector);
 };
 const session={name:'Synthetic registry',rawResult:{},status:{equipment:[{name:'DEMO-2'},{name:'DEMO-1'}],matched:1},snapshot:{rows:[{equipmentId:'DEMO-1',upn:'111',_source:{row:2}}]},completedStep:'RR OA/BT',completedSort:'step'};
 const context=vm.createContext({S:{session},clean,esc,natCmp,auditNormId,$:node,ic:()=>'',document:{body:{classList:{remove(){}}}},teardownAuditFilters(){},exportRemainingRegistryXlsx(){}});
 vm.runInContext(code,context);
 context.renderCompletedEquipment(()=>assert.fail('Should not redirect'));
 const html=node('#view').innerHTML;
 assert.match(html,/data-completed-open="DEMO-1"/);assert.match(html,/data-completed-open="DEMO-2"/);
 assert.ok(html.indexOf('data-completed-open="DEMO-1"')<html.indexOf('data-completed-open="DEMO-2"'));
 assert.match(html,/This report must contain only completed equipment/);
 assert.doesNotMatch(html,/Step completed|By step|data-completed-step|OA\/BT/);
 assert.equal(typeof node('#exportRemainingRegistry').onclick,'function');
});
