import test from 'node:test'
import assert from 'node:assert/strict'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { auditMakeCorrection, auditApplyCorrections, auditMergeCorrections, auditCorrectionImpact, auditRecommendationContext, auditProposeCorrection, auditRegistryRevision, auditReviewDocument, auditReadReviewDocument } from '../src/audit/actions.js'
import { S, resetSession } from '../src/state.js'
import { sessionAudit, activeRules } from '../src/ui/audit.js'
import { auditReadReferenceAoa } from '../src/audit/references.js'

const base={building:'DEMO',upn:'602',discipline:'ELECTRICAL',systemName:'602  Medium Voltage',closestParent:'602  Medium Voltage',equipmentDescription:'Electrical panel'}
function snapshot(records){return auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),...records.map(row=>EXTO_REV21_COLUMNS.map(c=>({...base,...row})[c.field]||''))],{sheet:'Registry',file:'synthetic.xlsx'})}
const finding=(row,rule='metadata.system-upn-mismatch',field='System Name')=>({id:'test:1',sheet:row._source.sheet,row:row._source.row,equipmentId:row.equipmentId,field,rule:{id:rule}})

test('corrections target the physical duplicate occurrence and never mutate source rows',()=>{
  const original=snapshot([{equipmentId:'PNL-1'},{equipmentId:'PNL-1'}]);
  const change=auditMakeCorrection(original.rows[1],'UPN','603',finding(original.rows[1]));
  const draft=auditApplyCorrections(original,[change]);
  assert.equal(draft.rows[0].upn,'602');assert.equal(draft.rows[1].upn,'603');assert.equal(original.rows[1].upn,'602');
  assert.throws(()=>auditApplyCorrections(original,[{...change,source:{...change.source,row:9}}]),/match/);
});
test('stale values and conflicting cell proposals fail atomically',()=>{
  const original=snapshot([{equipmentId:'PNL-1'}]),change=auditMakeCorrection(original.rows[0],'UPN','603');
  assert.throws(()=>auditApplyCorrections(original,[{...change,before:'650'}]),/original cell/);
  assert.throws(()=>auditApplyCorrections(original,[change,change]),/more than once/);
  assert.throws(()=>auditMergeCorrections(original,[],[change,{...change,value:'650'}]),/different values/);
  assert.equal(original.rows[0].upn,'602');
});
test('editing an already changed cell keeps the original baseline and restoring it removes the patch',()=>{
  const original=snapshot([{equipmentId:'PNL-1'}]),first=auditMakeCorrection(original.rows[0],'UPN','603');
  const draft=auditApplyCorrections(original,[first]),next=auditMakeCorrection(draft.rows[0],'UPN','650');
  const merged=auditMergeCorrections(original,[first],[next]);assert.equal(merged[0].before,'602');assert.equal(merged[0].value,'650');
  assert.equal(auditMergeCorrections(original,merged,[{...next,value:'602'}]).length,0);
});
test('a changed error message is not counted as a resolved finding',()=>{
  const issue=(id,severity='error')=>({id,rule:{id:'parent.invalid'},sheet:'Registry',row:2,field:'Closest Parent',severity});
  const impact=auditCorrectionImpact({findings:[issue('before')]},{findings:[issue('after')]});
  assert.equal(impact.resolved.length,0);assert.equal(impact.changed.length,1);assert.equal(impact.unsafe.length,1);
});
test('a supported System Name correction clears only issues the engine actually resolves',()=>{
  const original=snapshot([{equipmentId:'PNL-1',systemName:'650  Facility Management System'}]);
  const before=runSsmAudit(original),issue=before.findings.find(f=>f.rule.id==='metadata.system-upn-mismatch');assert.ok(issue);
  const proposal=auditProposeCorrection(issue,auditRecommendationContext(original));assert.ok(proposal);assert.match(proposal.reason,/Confirm the UPN/);
  const draft=auditApplyCorrections(original,proposal.changes),after=runSsmAudit(draft),impact=auditCorrectionImpact(before,after);
  assert.ok(impact.resolved.some(f=>f.id===issue.id));assert.ok(after.findings.length>0);
});
test('a UPN majority is not used to guess missing milestones',()=>{
  const original=snapshot([{equipmentId:'PNL-1',milestone:'L2-DEMO-1',milestoneParent:'L1-DEMO-1'},{equipmentId:'PNL-2'}]);
  const proposal=auditProposeCorrection(finding(original.rows[1],'milestone.required-pair','L2 Milestone'),auditRecommendationContext(original));assert.equal(proposal,null);
});
test('review files bind to every mapped row, not only tag identities',async()=>{
  const original=snapshot([{equipmentId:'PNL-1'}]),changed=snapshot([{equipmentId:'PNL-1',equipmentDescription:'Changed description'}]);
  assert.notEqual(await auditRegistryRevision(original),await auditRegistryRevision(changed));
  const document=await auditReviewDocument({baselineSnapshot:original,actioned:new Set(['test:1'])});
  assert.deepEqual([...(await auditReadReviewDocument(document,original)).actioned],['test:1']);
  await assert.rejects(auditReadReviewDocument(document,changed),/different registry revision/);
});
test('review files reject altered reference context and malformed history',async()=>{
  const original=snapshot([{equipmentId:'PNL-1'}]),document=await auditReviewDocument({baselineSnapshot:original});
  await assert.rejects(auditReadReviewDocument(document,original,{milestones:{entries:[]}}),/same reference/);
  await assert.rejects(auditReadReviewDocument({...document,history:[{reason:'x',owner:''}]},original),/history is invalid/);
});
test('reference upload order does not prevent restoring the same review',async()=>{
  const original=snapshot([{equipmentId:'PNL-1'}]),milestones={kind:'milestones',entries:[]},itemMasters={kind:'itemMasters',entries:[]};
  const document=await auditReviewDocument({baselineSnapshot:original,references:{milestones,itemMasters}});
  await assert.doesNotReject(auditReadReviewDocument(document,original,{itemMasters,milestones}));
});
test('filter views travel only in explicit review files and malformed values are rejected',async()=>{
  resetSession();const original=snapshot([{equipmentId:'PNL-1'}]);
  S.session.filterViews=[{name:'My scope',filters:{dimFilters:{building:['DEMO']}}}];
  const document=await auditReviewDocument({...S.session,baselineSnapshot:original});
  const restored=await auditReadReviewDocument(document,original);
  assert.deepEqual(restored.filterViews[0].filters.dimFilters.building,['DEMO']);
  await assert.rejects(auditReadReviewDocument({...document,filterViews:[{name:'Bad',filters:{dimFilters:{building:'not a list'}}}]},original),/invalid filter views/);
  resetSession();assert.deepEqual(S.session.filterViews,[]);
});
test('optional references are excluded from checks-ran counts until supplied',()=>{
  resetSession();const original=snapshot([{equipmentId:'PNL-1',milestone:'L2-DEMO-1',milestoneParent:'L1-DEMO-9'}]);
  const before=sessionAudit(original),baseCount=activeRules().length;
  S.session.references={milestones:auditReadReferenceAoa([['L2 ID','L1 ID'],['L2-DEMO-1','L1-DEMO-1']],'milestones','Register')};
  const after=sessionAudit(original);assert.equal(after.summary.checks,before.summary.checks+4);assert.equal(activeRules().length,baseCount+4);
  assert.ok(after.summary.source.reference>0);assert.equal(Object.values(after.summary.category).reduce((a,b)=>a+b,0),after.findings.length);
  resetSession();
});
