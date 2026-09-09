import test from 'node:test'
import assert from 'node:assert/strict'
import { SSM_AUDIT_RULES } from '../src/audit/engine.js'
import { SSM_AUDIT_REFERENCE_RULES } from '../src/audit/references.js'
import { auditActionEntry,auditActionPolicy,auditRecommendationContext,auditProposeCorrection,auditApplyCorrections } from '../src/audit/actions.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
const snapshot=rows=>auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),...rows.map(r=>EXTO_REV21_COLUMNS.map(c=>r[c.field]||''))],{sheet:'Registry'});
const issue=(row,id,field='Closest Parent')=>({id:'test:'+id,rule:{id},sheet:row._source.sheet,row:row._source.row,field});
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
