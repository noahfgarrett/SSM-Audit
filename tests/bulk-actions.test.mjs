import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import {readFileSync} from 'node:fs'
import {auditReadyFindingIds,auditCorrectionKey} from '../src/audit/actions.js'
import {auditReferenceRecommendation,auditReadReferenceAoa} from '../src/audit/references.js'

const ui=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
const change={source:{sheet:'Registry',row:2},prop:'upn',before:'650',value:'101'};
function session(){return {changes:[change],draftResolved:new Set(['fixed']),reviewHistory:[{disposition:'corrected-draft',findingIds:['fixed','still-open'],changes:[change]}]};}
test('ready means an applied correction still exists and its finding remains resolved',()=>{
 const s=session();assert.deepEqual([...auditReadyFindingIds(s)],['fixed']);
 s.changes=[];assert.equal(auditReadyFindingIds(s).size,0);
 s.changes=[change];s.draftResolved.clear();assert.equal(auditReadyFindingIds(s).size,0);
 s.draftResolved.add('fixed');s.reviewHistory[0].disposition='exception';assert.equal(auditReadyFindingIds(s).size,0);
 s.reviewHistory[0].disposition='corrected-draft';s.changes=[{...change,value:'102'}];assert.equal(auditReadyFindingIds(s).size,0);
});
test('ready state derives from saved review data without persisting a separate readiness claim',()=>{
 const s=session(),saved=JSON.parse(JSON.stringify({changes:s.changes,history:s.reviewHistory}));
 assert.deepEqual([...auditReadyFindingIds({changes:saved.changes,reviewHistory:saved.history,draftResolved:new Set(['fixed'])})],['fixed']);
});
test('rule selection scopes to shown pending findings and drops hidden or fully applied rules',()=>{
 const s={...session(),selectedActionRules:new Set(['a','hidden']),excluded:new Set(['dismissed'])};
 const groups=[{rules:[{rule:{id:'a'},matches:[{id:'pending'},{id:'fixed'},{id:'dismissed'}]},{rule:{id:'b'},matches:[{id:'next'}]},{rule:{id:'done'},matches:[{id:'fixed'}]}]}];
 const nodes=new Map(),node=k=>{if(!nodes.has(k))nodes.set(k,{});return nodes.get(k);};
 const context=vm.createContext({S:{session:s},auditReadyFindingIds,ic:()=>'',esc:x=>x,modifyGroups:()=>groups,$:node,$$:()=>[],isExcludedId:id=>s.excluded.has(id)});
 vm.runInContext(ui.slice(ui.indexOf('function modifyReadySet('),ui.indexOf('function excludedInBase(')),context);
 vm.runInContext(ui.slice(ui.indexOf('function modifyActiveFindings('),ui.indexOf('function actionsExportFindings(')),context);
 const api=vm.runInContext('({modifySelectableRules,modifyActiveFindings,syncModifyRuleSelection})',context);
 api.syncModifyRuleSelection();assert.deepEqual([...s.selectedActionRules],['a']);
 assert.equal(node('#modifySelectAll').indeterminate,true);assert.equal(node('#modifyActionSelected').disabled,false);
 assert.deepEqual(Array.from(api.modifySelectableRules(),e=>e.rule.id),['a','b']);
 assert.deepEqual(Array.from(api.modifyActiveFindings(groups[0].rules[0].matches),f=>f.id),['pending']);
 s.selectedActionRules.add('b');api.syncModifyRuleSelection();assert.equal(node('#modifySelectAll').checked,true);
 groups[0].rules=groups[0].rules.slice(1);api.syncModifyRuleSelection();assert.deepEqual([...s.selectedActionRules],['b']);
});
test('milestone lookups reuse their indexes across 500 suggestions and refresh for a new reference',()=>{
 const reference=auditReadReferenceAoa([['L2 ID','Title','L1 ID'],...Array.from({length:1000},(_,i)=>[`DEMO-L2-M1-${i+1}`,`Phase ${i+1}`,'DEMO-L1-M1-10'])],'milestones','Register');
 const entries=[...reference.entries];let scans=0;entries[Symbol.iterator]=function*(){scans++;yield* Array.prototype.values.call(this);};
 const refs={milestones:{...reference,entries}},prepared={};
 for(let i=0;i<500;i++)assert.match(auditReferenceRecommendation({milestone:`DEMO-L2-M1-${i+1}`,milestoneParent:''},'Milestone Parent',refs,prepared).value,/DEMO-L1-M1-10/);
 assert.equal(scans,2,'one milestone index and one parent index, not two scans per row');
 const replacement=auditReadReferenceAoa([['L2 ID','Title','L1 ID'],['DEMO-L2-M1-1','Phase 1','DEMO-L1-M1-20']],'milestones','Register');
 assert.equal(auditReferenceRecommendation({milestone:'DEMO-L2-M1-1'},'Milestone Parent',{milestones:replacement},prepared).value,'DEMO-L1-M1-20');
});
test('cached catalog suggestions preserve ambiguity and refresh when the selected catalog changes',()=>{
 const catalog=names=>auditReadReferenceAoa([['Item Master Name'],...names.map(n=>[n])],'itemMasters','Catalog'),prepared={};
 assert.equal(auditReferenceRecommendation({itemMaster:'DEMO_EL_PANEL'},'Item Master',{itemMasters:catalog(['VF1_EL_PANEL'])},prepared).value,'VF1_EL_PANEL');
 assert.equal(auditReferenceRecommendation({itemMaster:'DEMO_EL_PANEL'},'Item Master',{itemMasters:catalog(['VF1_EL_PANEL','VF2_EL_PANEL'])},prepared),null);
});
