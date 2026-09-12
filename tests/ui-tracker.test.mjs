import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { esc } from '../src/core/text.js'

test('Tracker export selector keeps its choice and passes it to the download', async () => {
  const source = readFileSync(new URL('../src/ui/audit.js', import.meta.url), 'utf8')
  const start = source.indexOf('function renderExportOptions(){')
  const end = source.indexOf('\nexport function openExportOptions()', start)
  const session = { exportKind: 'tracker', result: { findings: [] } }, nodes = new Map(), downloads = []
  const inputs = [{ value: 'milestone' }, { value: 'discipline' }]
  const node = selector => {
    if (!nodes.has(selector)) nodes.set(selector, { innerHTML: '' })
    return nodes.get(selector)
  }
  const context = vm.createContext({
    S: { session }, esc, ic: () => '', $: node,
    $$: selector => selector === 'input[name="tracker-signoff"]' ? inputs : [],
    exportPlan: () => ({ levels: {}, rules: {} }), exportPlanSummary: () => ({}),
    SSM_AUDIT_SEVERITIES: [], closeExportOptions() {},
    exportTrackerXlsx: async mode => downloads.push(mode),
  })
  vm.runInContext(source.slice(start, end), context)
  vm.runInContext('renderExportOptions()', context)
  assert.match(node('#exportModalBody').innerHTML, /value="milestone" checked/)
  assert.match(node('#exportModalBody').innerHTML, /All checkmarks start empty/)
  inputs[1].onchange()
  assert.equal(session.trackerSignOffBy, 'discipline')
  assert.match(node('#exportModalBody').innerHTML, /value="discipline" checked/)
  await node('#exportGo').onclick()
  inputs[0].onchange()
  await node('#exportGo').onclick()
  assert.deepEqual(downloads, ['discipline', 'milestone'])
})

test('Actions export modal passes precisely its filtered findings to the Actions download',async()=>{
  const source=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
  const start=source.indexOf('function renderExportOptions(){'),end=source.indexOf('\nexport function openExportOptions()',start);
  const findings=[{id:'selected'}],downloads=[],nodes=new Map();
  const node=selector=>{if(!nodes.has(selector))nodes.set(selector,{innerHTML:''});return nodes.get(selector);};
  const context=vm.createContext({S:{session:{exportKind:'actions',result:{findings:[]}}},esc,ic:()=>'', $:node,$$:()=>[],
    exportPlan:()=>({levels:{},rules:{}}),exportPlanSummary:()=>({}),SSM_AUDIT_SEVERITIES:[],closeExportOptions(){},
    actionsExportFindings:()=>findings,exportActionsXlsx:async rows=>downloads.push(rows),
  });
  vm.runInContext(source.slice(start,end),context);vm.runInContext('renderExportOptions()',context);
  assert.match(node('#exportModalBody').innerHTML,/1 active findings/);
  assert.match(node('#exportModalBody').innerHTML,/Check Actionable/);
  assert.match(node('#exportModalBody').innerHTML,/check Actioned/);
  await node('#exportGo').onclick();assert.equal(downloads[0],findings);
});

test('Updated Registry modal explains the batches and refreshes the Actions summary after download',async()=>{
  const source=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
  const start=source.indexOf('function renderExportOptions(){'),end=source.indexOf('\nexport function openExportOptions()',start);
  const nodes=new Map(),S={screen:'modify',session:{exportKind:'updated',result:{findings:[]},changes:[{}],sourceBytes:new Uint8Array([1])}};
  let downloads=0,renders=0;
  const node=selector=>{if(!nodes.has(selector))nodes.set(selector,{innerHTML:''});return nodes.get(selector);};
  const context=vm.createContext({S,esc,ic:()=>'', $:node,$$:()=>[],exportPlan:()=>({levels:{},rules:{}}),exportPlanSummary:()=>({}),SSM_AUDIT_SEVERITIES:[],closeExportOptions(){},currentNavigate(){},
    exportUpdatedRegistryXlsx:async()=>{downloads++;return true;},rerenderModifications:()=>renders++});
  vm.runInContext(source.slice(start,end),context);vm.runInContext('renderExportOptions()',context);
  assert.match(node('#exportModalBody').innerHTML,/1,950 changed equipment rows per file/);
  assert.match(node('#exportModalBody').innerHTML,/Extract the ZIP/);
  await node('#exportGo').onclick();assert.equal(downloads,1);assert.equal(renders,1);
  S.screen='dashboard';await node('#exportGo').onclick();assert.equal(downloads,2);assert.equal(renders,1,'finishing export does not navigate away from another screen');
});
