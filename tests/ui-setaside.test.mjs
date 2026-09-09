import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import {readFileSync} from 'node:fs'
import {esc} from '../src/core/text.js'

const source=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
function harness(){
  const rule={id:'parent.cross-upn'},findings=[
    {id:'a',why:'Child UPN 650, parent UPN 101',rule},
    {id:'b',why:'Child UPN 650, parent UPN 101',rule},
    {id:'c',why:'Child UPN 603, parent UPN 602',rule},
    {id:'d',why:'Child UPN 603, parent UPN 602',rule},
  ];
  const session={excluded:new Set(),changes:[{tag:'DEMO-1',field:'Building',value:'DEMO-B'}],snapshot:{rows:[{equipmentId:'DEMO-1',building:'DEMO-B'}]},rawResult:{findings}};
  const nodes=new Map(),messages=[],calls={refresh:0,counts:0};
  const node=selector=>{if(!nodes.has(selector)){const classes=new Set();nodes.set(selector,{innerHTML:'',checked:true,disabled:false,classList:{toggle:(name,on)=>on?classes.add(name):classes.delete(name),contains:name=>classes.has(name)}});}return nodes.get(selector);};
  const context=vm.createContext({S:{session},esc,isExcludedId:id=>session.excluded.has(id),modifyPatternMap:new Map(),modifyPatternKey:(_rule,why)=>why,
    modifyAmount:value=>String(value),modifyHighlight:esc,modifyPctBtn:()=>'',modifyItemMasterSwap:()=>'',ic:name=>`<i>${name}</i>`,SEVERITY_LABELS:{error:'Invalid'},
    $:node,$$:()=>[],saveExcluded:()=>{session.reviewDirty=true;},refreshSessionResult:()=>{calls.refresh++;},updateModifyCounts:()=>{calls.counts++;},toast:message=>messages.push(message)});
  vm.runInContext(source.slice(source.indexOf('function modifyActiveFindings('),source.indexOf('function modifyFillPatternRows(')),context);
  vm.runInContext(source.slice(source.indexOf('function syncModifyPatternBox('),source.indexOf('function modifyRuleSeverity(')),context);
  const api=vm.runInContext('({modifyPatterns,modifyActiveFindings,modifyPatternHtml,setModifyPatternAside,syncModifyPatternBox})',context);
  const groups=api.modifyPatterns({rule,matches:findings}).groups;
  for(const group of groups)context.modifyPatternMap.set(group.key,group.findings);
  return {api,session,groups,node,calls,messages,findings};
}

test('setting aside one pattern leaves its sibling active and preserves registry data and staged corrections',()=>{
  const h=harness(),before=JSON.stringify([h.session.snapshot,h.session.changes]),group=h.groups[0];
  h.api.setModifyPatternAside(group.key,true);
  assert.deepEqual([...h.session.excluded],['a','b']);
  assert.deepEqual(Array.from(h.api.modifyActiveFindings(h.findings),f=>f.id),['c','d']);
  assert.equal(h.node(`[data-mod-pattern="${group.key}"]`).classList.contains('is-excluded'),true);
  assert.match(h.node(`[data-mod-aside-group="${group.key}"]`).innerHTML,/Restore/);
  assert.equal(h.node(`[data-mod-action-group="${group.key}"]`).disabled,true);
  assert.equal(JSON.stringify([h.session.snapshot,h.session.changes]),before);
  assert.equal(h.calls.refresh,1);assert.equal(h.calls.counts,1);assert.equal(h.session.reviewDirty,true);
  h.api.setModifyPatternAside(group.key,false);
  assert.equal(h.session.excluded.size,0);assert.equal(h.node(`[data-mod-action-group="${group.key}"]`).disabled,false);
  assert.match(h.node(`[data-mod-aside-group="${group.key}"]`).innerHTML,/Set aside/);
  assert.equal(JSON.stringify([h.session.snapshot,h.session.changes]),before);
});

test('partially set-aside groups retain tri-state selection and bulk actions use only active findings',()=>{
  const h=harness(),group=h.groups[0];h.session.excluded.add('a');h.api.syncModifyPatternBox(group.key);
  assert.equal(h.node(`input[data-mod-group="${group.key}"]`).indeterminate,true);
  assert.deepEqual(Array.from(h.api.modifyActiveFindings(group.findings),f=>f.id),['b']);
  h.api.setModifyPatternAside(group.key,true);assert.deepEqual([...h.session.excluded],['a','b']);
  assert.match(h.api.modifyPatternHtml(group),/is-excluded/);assert.match(h.api.modifyPatternHtml(group),/Restore/);
});

test('set-aside controls cannot change review state while an action is being prepared',()=>{
  const h=harness();h.session.reviewBusy=true;h.api.setModifyPatternAside(h.groups[0].key,true);
  assert.equal(h.session.excluded.size,0);assert.equal(h.calls.refresh,0);
});

test('pattern button and filtered batch action handlers preserve set-aside scope',()=>{
  assert.match(source,/data-mod-aside-group/);
  assert.match(source,/modifyActiveFindings\(collectMatches\(\)\)/);
  assert.match(source,/modifyActiveFindings\(modifyPatternMap\.get\(actionGroup\.dataset\.modActionGroup\)\|\|\[\]\)/);
});
