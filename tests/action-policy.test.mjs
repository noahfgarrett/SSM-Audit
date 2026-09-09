import test from 'node:test'
import assert from 'node:assert/strict'
import { SSM_AUDIT_RULES, runSsmAudit } from '../src/audit/engine.js'
import { SSM_AUDIT_REFERENCE_RULES } from '../src/audit/references.js'
import { auditActionEntry,auditActionPolicy,auditRecommendationContext,auditProposeCorrection,auditApplyCorrections } from '../src/audit/actions.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
const snapshot=rows=>auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),...rows.map(r=>EXTO_REV21_COLUMNS.map(c=>r[c.field]||''))],{sheet:'Registry'});
const issue=(row,id,field='Closest Parent')=>({id:'test:'+id,rule:{id},sheet:row._source.sheet,row:row._source.row,field});
test('same-project references clear even when repeated or dependency tags are missing',()=>{
 for(const dependencyProject of ['DEMO','demo; DEMO','DEMO ; DEMO ; DEMO']){
  const s=snapshot([{equipmentId:'CHILD',project:'Demo',dependencyProject,dependencies:'NOT-IN-REGISTRY'}]),result=runSsmAudit(s);
  const issue=result.findings.find(f=>f.rule.id==='dependency.project-not-needed');assert.ok(issue);
  assert.equal(result.findings.some(f=>f.rule.id==='dependency.project-multiple'),dependencyProject.includes(';'));
  assert.ok(result.findings.some(f=>f.rule.id==='dependency.unresolved'));
  const entry=auditActionEntry(issue,auditRecommendationContext(s));assert.equal(entry.changes[0].value,'');
  const after=runSsmAudit(auditApplyCorrections(s,entry.changes));assert.ok(after.findings.some(f=>f.rule.id==='dependency.unresolved'));
 }
});
test('external project duplicates collapse but distinct projects never receive a guessed clear',()=>{
 for(const [dependencyProject,expected] of [['OTHER; OTHER','OTHER'],['OTHER; THIRD','OTHER; THIRD']]){
  const s=snapshot([{equipmentId:'CHILD',project:'DEMO',dependencyProject,dependencies:'EXTERNAL-TAG'}]),result=runSsmAudit(s);
  assert.ok(!result.findings.some(f=>f.rule.id==='dependency.project-not-needed'));
  const issue=result.findings.find(f=>f.rule.id==='dependency.project-multiple');assert.ok(issue);
  const entry=auditActionEntry(issue,auditRecommendationContext(s));assert.equal(entry.changes[0].value,expected);
 }
});
test('optional migration populates the unique VF name from the selected catalog',()=>{
 const s=snapshot([{equipmentId:'DEMO-1',itemMaster:'CAMPUS_DEMO_EL_PANEL'}]);
 const f=issue(s.rows[0],'item-master.migration-advisory','Item Master Unique Identifier');
 const catalog=names=>({itemMasters:{kind:'itemMasters',entries:names.map(name=>({name}))}});
 const entry=auditActionEntry(f,auditRecommendationContext(s,catalog(['VF_DEMO_EL_PANEL'])));
 assert.equal(entry.changes[0].before,'CAMPUS_DEMO_EL_PANEL');assert.equal(entry.changes[0].value,'VF_DEMO_EL_PANEL');
 const fallback=auditActionEntry(f,auditRecommendationContext(s,catalog(['VF_EL_PANEL'])));
 assert.equal(fallback.changes[0].value,'VF_EL_PANEL');assert.match(fallback.reason,/Leading site text/);
 for(const names of [[],['VF_UNRELATED'],['VF1_DEMO_EL_PANEL','VF2_DEMO_EL_PANEL']]){
  const uncertain=auditActionEntry(f,auditRecommendationContext(s,catalog(names)));
  assert.equal(uncertain.changes[0].value,uncertain.changes[0].before);assert.match(uncertain.reason,/No reliable suggestion/);
 }
});
test('every shipped rule has an explicit action policy, including engineer-only decisions',()=>{
 for(const rule of [...Object.values(SSM_AUDIT_RULES),...Object.values(SSM_AUDIT_REFERENCE_RULES)]){
  const policy=auditActionPolicy({rule});assert.ok(policy.known||policy.targets,rule.id);
 }
});
for(const [id,field,prop] of [['parent.cross-building','Building','building'],['parent.cross-discipline','Discipline','discipline']])test(`${id} edits the selected physical equipment, never a relationship`,()=>{
 const s=snapshot([{equipmentId:'CHILD',closestParent:'PARENT',building:'BLDG-1',discipline:'ELECTRICAL',dependencies:'POWER'},{equipmentId:'PARENT',building:'BLDG-2',discipline:'MECHANICAL DRY'}]),context=auditRecommendationContext(s),f=issue(s.rows[0],id);
 assert.equal(auditProposeCorrection(f,context),null);
 const child=auditActionEntry(f,context,'child'),parent=auditActionEntry(f,context,'parent');
 assert.deepEqual(child.changes.map(c=>c.field),[field]);assert.deepEqual(parent.changes.map(c=>c.field),[field]);
 assert.equal(child.changes[0].value,s.rows[1][prop]);assert.equal(parent.changes[0].value,s.rows[0][prop]);
 assert.equal(parent.changes[0].tag,'PARENT');
 const updated=auditApplyCorrections(s,parent.changes);assert.equal(updated.rows[0].closestParent,'PARENT');assert.equal(updated.rows[0].dependencies,'POWER');assert.equal(s.rows[1][prop],prop==='building'?'BLDG-2':'MECHANICAL DRY');
});
test('missing and ambiguous parents cannot be edited by guessing a row',()=>{
 for(const additional of [[],[{equipmentId:'PARENT'},{equipmentId:'PARENT'}]]){
  const s=snapshot([{equipmentId:'CHILD',closestParent:'PARENT',building:'BLDG-1'},...additional]);
  assert.equal(auditActionEntry(issue(s.rows[0],'parent.cross-building'),auditRecommendationContext(s),'parent'),null);
 }
});
test('unknown corrections keep relevant fields editable without manufactured assignments',()=>{
 const s=snapshot([{equipmentId:'CHILD',upn:'101',dependencies:'POWER',closestParent:'PARENT'}]),context=auditRecommendationContext(s);
 for(const [id,fields] of [['milestone.incomplete-pair',['L1 Milestone Parent','L2 Milestone']],['dependency.unresolved',['Dependencies','Dependency Project']],['dependency.precedence-cycle',['Closest Parent','Dependencies']]]){
  const entry=auditActionEntry(issue(s.rows[0],id),context);assert.deepEqual(entry.changes.map(c=>c.field),fields);assert.ok(entry.changes.every(c=>c.value===c.before));
 }
});
test('clearing stale Dependency Project does not clear Dependencies',()=>{
 const s=snapshot([{equipmentId:'CHILD',dependencies:'PARENT',dependencyProject:'Old project'},{equipmentId:'PARENT'}]),context=auditRecommendationContext(s),f=issue(s.rows[0],'dependency.project-not-needed','Dependency Project');
 const entry=auditActionEntry(f,context);assert.equal(entry.changes.length,1);assert.equal(entry.changes[0].value,'');
 const updated=auditApplyCorrections(s,entry.changes);assert.equal(updated.rows[0].dependencyProject,'');assert.equal(updated.rows[0].dependencies,'PARENT');
});
