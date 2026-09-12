import test from 'node:test'
import assert from 'node:assert/strict'
import { EXTO_REV21_COLUMNS, extoRev21SystemsForUpn } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { auditMakeCorrection, auditApplyCorrections, auditMergeCorrections, auditCorrectionImpact, auditRecommendationContext, auditProposeCorrection, auditRegistryRevision, auditReviewDocument, auditReadReviewDocument, auditActionEntry } from '../src/audit/actions.js'
import { S, resetSession } from '../src/state.js'
import { sessionAudit, activeRules } from '../src/ui/audit.js'
import { auditReadReferenceAoa } from '../src/audit/references.js'

const base={building:'DEMO',upn:'602',discipline:'ELECTRICAL',systemName:'602  Medium Voltage',closestParent:'602  Medium Voltage',equipmentDescription:'Electrical panel'}
function snapshot(records){return auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),...records.map(row=>EXTO_REV21_COLUMNS.map(c=>({...base,...row})[c.field]||''))],{sheet:'Registry',file:'synthetic.xlsx'})}
const finding=(row,rule='metadata.system-upn-mismatch',field='System Name')=>({id:'test:1',sheet:row._source.sheet,row:row._source.row,equipmentId:row.equipmentId,field,rule:{id:rule}})

for(const [ruleId,prop] of [['parent.cross-building','building'],['parent.cross-discipline','discipline']])test(`${ruleId} asks for reconciliation and never copies the other row as a suggested fix`,()=>{
  const original=snapshot([{equipmentId:'PARENT',building:'DEMO-A',discipline:'MECHANICAL DRY'}, {equipmentId:'CHILD',closestParent:'PARENT',building:'DEMO-B',discipline:'MECHANICAL WET'}]);
  const issue=runSsmAudit(original).findings.find(f=>f.rule.id===ruleId);assert.ok(issue);
  assert.match(issue.recommendation,/Reconcile/);assert.doesNotMatch(issue.recommendation,/pick a parent|as a dependency/);
  const context=auditRecommendationContext(original);
  for(const [target,index] of [['child',1],['parent',0]]){
    const entry=auditActionEntry(issue,context,target);assert.equal(entry.confidence,'Engineer review');
    assert.equal(entry.changes.length,1);assert.equal(entry.changes[0].prop,prop);
    assert.equal(entry.changes[0].value,original.rows[index][prop]);
    assert.equal(entry.changes[0].before,entry.changes[0].value);
  }
});

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

function servedDrive({upn='101',child={},parent={}}={}){
  return snapshot([
    {equipmentId:`F77-MAH${upn}-01-00`,equipmentDescription:'Makeup Air Handler MAH',upn,systemName:extoRev21SystemsForUpn(upn)[0],discipline:'MECHANICAL DRY',closestParent:'',...parent},
    {equipmentId:`F77-VFD${upn}-01-00`,equipmentDescription:'Variable Frequency Drive',upn:'650',systemName:extoRev21SystemsForUpn('650')[0],discipline:'FACILITIES MONITORING SYSTEM',closestParent:`F77-MAH${upn}-01-00`,dependencies:'SUPPLY-PANEL',...child},
  ]);
}
for(const upn of ['101','104'])test(`tag and served equipment agree on UPN ${upn}: prefill coordinated corrections`,()=>{
  const original=servedDrive({upn}),before=runSsmAudit(original),issue=before.findings.find(f=>f.rule.id==='parent.cross-upn');assert.ok(issue);
  assert.match(issue.recommendation,new RegExp(`Change F77-VFD${upn}-01-00's UPN from 650 to ${upn}`));
  assert.match(issue.recommendation,/System Name/);assert.doesNotMatch(issue.recommendation,/pick a parent|as a dependency/);
  const proposal=auditProposeCorrection(issue,auditRecommendationContext(original));assert.ok(proposal);
  assert.deepEqual(proposal.changes.map(c=>c.field),['UPN','System Name','Discipline']);
  const draft=auditApplyCorrections(original,proposal.changes),row=draft.rows[1];
  assert.equal(row.upn,upn);assert.equal(row.systemName,extoRev21SystemsForUpn(upn)[0]);assert.equal(row.discipline,'MECHANICAL DRY');
  assert.equal(row.closestParent,original.rows[1].closestParent);assert.equal(row.dependencies,'SUPPLY-PANEL');
  assert.ok(!runSsmAudit(draft).findings.some(f=>f.rule.id==='parent.cross-upn'));
  assert.equal(auditCorrectionImpact(before,runSsmAudit(draft)).unsafe.length,0);
});
for(const child of [
  {equipmentId:'F77-VFD650-01-00'},
  {equipmentId:'F77-VFD-01-00'},
  {equipmentId:'F77-VFD101-RIO650-01-00'},
  {equipmentDescription:'Remote I/O Panel',equipmentId:'F77-RIO650-01-00'},
  {discipline:'ELECTRICAL'},
  {building:'OTHER'},
])test(`uncertain or conflicting tag evidence has no UPN proposal: ${JSON.stringify(child)}`,()=>{
  const original=servedDrive({child}),row=original.rows[1];
  assert.equal(auditProposeCorrection(finding(row,'parent.cross-upn','Closest Parent'),auditRecommendationContext(original)),null);
});
test('UPN guidance identifies incorrect parent metadata without renaming equipment',()=>{
  const original=servedDrive({parent:{upn:'650'},child:{upn:'101',systemName:extoRev21SystemsForUpn('101')[0]}});
  const issue=runSsmAudit(original).findings.find(f=>f.rule.id==='parent.cross-upn');assert.ok(issue);
  assert.match(issue.recommendation,/Change F77-MAH101-01-00's UPN from 650 to 101/);
  assert.match(issue.recommendation,/Keep the equipment tags and existing links unchanged/);
});
test('UPN guidance supports grandchildren through an intermediate parent',()=>{
  const original=servedDrive();
  const rows=original.rows.map(row=>({...row}));
  rows.push({...rows[1],equipmentId:'F77-TET101-01-00',upn:'101',closestParent:rows[1].equipmentId,equipmentDescription:'Temperature transmitter'});
  const result=runSsmAudit(snapshot(rows));
  const issue=result.findings.find(f=>f.rule.id==='parent.cross-upn'&&f.equipmentId==='F77-TET101-01-00');
  assert.match(issue.recommendation,/Change F77-VFD101-01-00's UPN from 650 to 101/);
});
for(const equipmentId of ['F77-VFD-101-01','F77-VFD101-RIO650-01','F77-VFD650-01','BLDG101-VFD-01'])test(`uncertain UPN guidance requests metadata reconciliation: ${equipmentId}`,()=>{
  const original=servedDrive({child:{equipmentId}}),issue=runSsmAudit(original).findings.find(f=>f.rule.id==='parent.cross-upn');
  assert.match(issue.recommendation,/Reconcile the UPN metadata/);assert.doesNotMatch(issue.recommendation,/Change .*UPN from/);
});
test('a numeric building prefix is not UPN evidence',()=>{
  const original=servedDrive({child:{equipmentId:'BLDG650-VFD101-01-00'}}),issue=runSsmAudit(original).findings.find(f=>f.rule.id==='parent.cross-upn');
  assert.match(issue.recommendation,/UPN from 650 to 101/);
  assert.ok(auditProposeCorrection(issue,auditRecommendationContext(original)));
});
test('conflicting parent tag evidence does not prefill a metadata correction',()=>{
  const original=servedDrive({parent:{equipmentId:'F77-MAH104-01-00'},child:{closestParent:'F77-MAH104-01-00'}});
  const issue=runSsmAudit(original).findings.find(f=>f.rule.id==='parent.cross-upn');
  assert.match(issue.recommendation,/Reconcile/);assert.equal(auditProposeCorrection(issue,auditRecommendationContext(original)),null);
});
test('an instrument under a matching system header suggests metadata, never a replacement parent',()=>{
  const original=servedDrive({parent:{equipmentId:'DEMO-HEADER',equipmentDescription:'System header',itemMaster:'VF_Blank'},child:{equipmentId:'F77-TET101-02-01',equipmentDescription:'Temperature Transmitter',closestParent:'DEMO-HEADER'}});
  const proposal=auditProposeCorrection(finding(original.rows[1],'parent.cross-upn','Closest Parent'),auditRecommendationContext(original));
  assert.ok(proposal);
  assert.deepEqual(proposal.changes.map(c=>[c.field,c.value]),[['UPN','101'],['System Name',extoRev21SystemsForUpn('101')[0]]]);
  const draft=auditApplyCorrections(original,proposal.changes);
  assert.equal(draft.rows[1].closestParent,'DEMO-HEADER');
  assert.equal(draft.rows[1].dependencies,'SUPPLY-PANEL');
});
test('tag and correct System Name repair the UPN instead of overwriting the correct system',()=>{
  const original=servedDrive({child:{systemName:extoRev21SystemsForUpn('101')[0]}}),row=original.rows[1];
  const issue=runSsmAudit(original).findings.find(f=>f.rule.id==='metadata.system-upn-mismatch'&&f.row===row._source.row);
  const proposal=auditProposeCorrection(issue,auditRecommendationContext(original));
  assert.deepEqual(proposal.changes.map(c=>[c.field,c.value]),[['UPN','101']]);
});
test('unknown non-electrical nomenclature does not select a System Name just from the current UPN',()=>{
  const original=servedDrive({child:{equipmentId:'F77-UNKNOWN',systemName:extoRev21SystemsForUpn('101')[0]}}),row=original.rows[1];
  assert.equal(auditProposeCorrection(finding(row),auditRecommendationContext(original)),null);
});

test('rows missing milestones get the closest-parent pair as a suggested fix with a measured confidence',()=>{
  const l2='L2-M1-1220 - UPN 602 MV Enabling Green Tag',l1='L1-M1-110 - BC1 Available'
  const original=snapshot([{equipmentId:'MV-ROOT',milestone:l2,milestoneParent:l1},{equipmentId:'MV-CHILD',closestParent:'MV-ROOT'}]);
  const issue=runSsmAudit(original).findings.find(f=>f.rule.id==='milestone.incomplete-pair'&&f.equipmentId==='MV-CHILD');assert.ok(issue,'the child is flagged for missing milestones');
  const proposal=auditProposeCorrection(issue,auditRecommendationContext(original));
  assert.ok(proposal,'a recommendation is proposed');
  assert.deepEqual(proposal.changes.map(change=>[change.field,change.value]).sort(),[['L1 Milestone Parent',l1],['L2 Milestone',l2]]);
  assert.match(proposal.confidence,/^95% — parent chain$/);
  assert.match(proposal.reason,/MV-ROOT/);
});
