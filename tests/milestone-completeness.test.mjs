import test from 'node:test'
import assert from 'node:assert/strict'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'

const run=rows=>runSsmAudit(auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),...rows.map(row=>EXTO_REV21_COLUMNS.map(c=>row[c.field]||''))],{sheet:'Registry'})).findings.filter(f=>f.rule.id==='milestone.incomplete-pair');
test('missing either or both milestones is Rule Broken, including wholly unassigned registries',()=>{
  for(const values of [{},{milestone:'DEMO-L2-20'},{milestoneParent:'DEMO-L1-10'}]){
    const findings=run([{equipmentId:'DEMO-1',...values}]);
    assert.equal(findings.length,1);assert.equal(findings[0].severity,'error');
    assert.equal(findings[0].rule.confidence,'required');
  }
  const findings=run([{equipmentId:'DEMO-1'},{equipmentId:'DEMO-2'}]);
  assert.equal(findings.length,2);
});
test('placeholder milestones are missing but populated pairs are not',()=>{
  for(const value of [' ','N/A','NA','TBD','TBC','None','NULL','Not Applicable','-','---']){
    assert.equal(run([{equipmentId:'DEMO-1',milestone:value,milestoneParent:'DEMO-L1-10'}])[0].severity,'error');
    assert.equal(run([{equipmentId:'DEMO-1',milestone:'DEMO-L2-20',milestoneParent:value}])[0].severity,'error');
  }
  assert.equal(run([{equipmentId:'DEMO-1',milestone:'DEMO-L2-20',milestoneParent:'DEMO-L1-10'}]).length,0);
});
test('milestone readiness remains distinct from Exto gating',()=>{
  for(const field of ['milestone','milestoneParent'])assert.equal(EXTO_REV21_COLUMNS.find(c=>c.field===field).gating,false);
});
