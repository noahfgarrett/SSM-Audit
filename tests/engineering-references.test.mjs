import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { auditReadEngineeringAoa, auditReadEngineeringWorkbook, auditEngineeringFindings } from '../src/audit/engineering-references.js'
import { auditSessionResult } from '../src/audit/review.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { auditActionPolicy, auditProposeCorrection, auditRecommendationContext, auditReviewDocument, auditReadReviewDocument } from '../src/audit/actions.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js',import.meta.url),'utf8'));
const parse=(rows,kind='mel')=>auditReadEngineeringAoa(rows,kind,'Selected sheet');
const snap=(rows)=>auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),...rows.map(row=>EXTO_REV21_COLUMNS.map(c=>row[c.field]||''))],{sheet:'Registry'});
const row={equipmentId:'TEST-PUMP111-01',upn:'111',building:'TEST',discipline:'MECHANICAL WET',systemName:'111 Chilled Water',closestParent:'111 Chilled Water',equipmentDescription:'Pump'};
const elec={upn:'603',building:'TEST',discipline:'ELECTRICAL',systemName:'603 Low Voltage',closestParent:'603 Low Voltage',equipmentDescription:'Panel'};
const fms={upn:'650',building:'TEST',discipline:'FACILITIES MONITORING SYSTEM',systemName:'650 Facilities Monitoring System',closestParent:'650 Facilities Monitoring System',equipmentDescription:'RIO'};

test('MEL row 2 headers, formatted fields, explicit parent and system description are read',()=>{
  const ref=parse([['Synthetic'],['Equipment Tag','UPN','Bldg','Discipline','System Description','System Parent Equipment Tag(s)'],[row.equipmentId,'111','TEST','MECHANICAL WET','Chilled Water','TEST-ROLLUP']]);
  assert.equal(ref.headerRow,2);assert.equal(ref.entries[0].sourceRow,3);assert.equal(ref.entries[0].systemName,'111 Chilled Water');assert.equal(ref.entries[0].closestParent,'TEST-ROLLUP');
});
test('missing or mismatched headers are rejected without inventing mappings',()=>{
  assert.equal(parse([['Equipment Tag'],['TEST-1']]),null);
  assert.equal(parse([['Panel'],['TEST-1']],'pmd'),null);
  assert.equal(parse([['Equipment ID','UPN'],['TEST-1','111']],'cable'),null);
});
test('sheet discovery supports dense, sparse and offset headers',()=>{
  const rows=[[],[null,'Equipment Tag','UPN'],[null,row.equipmentId,7]];
  const sheet=XLSX.utils.aoa_to_sheet(rows);sheet.C3.z='000';
  const book={SheetNames:['Notes','MEL'],Sheets:{Notes:XLSX.utils.aoa_to_sheet([['Notes']]),MEL:sheet}};
  assert.equal(auditReadEngineeringWorkbook(book,'mel')[0].entries[0].upn,'007');
  const bytes=XLSX.write(book,{type:'array',bookType:'xlsx'});
  const dense=XLSX.read(bytes,{type:'array',dense:true});
  assert.equal(auditReadEngineeringWorkbook(dense,'mel')[0].entries[0].equipmentId,row.equipmentId);
  dense.Sheets.MEL[2][2].w='007';
  assert.equal(auditReadEngineeringWorkbook(dense,'mel')[0].entries[0].upn,'007');
});
test('EasyPower downstream levels sort numerically, cross gaps and collapse adjacent copies',()=>{
  const ref=parse([['Starting Source','Downstream 10','Downstream 2','Downstream 3','Final Source','ID Name'],['TEST-GIS','TEST-LVS','TEST-XFM','','TEST-LVS','TEST-LOAD']],'easyPower');
  assert.deepEqual(ref.entries.map(e=>[e.equipmentId,e.parent]),[['TEST-GIS',''],['TEST-XFM','TEST-GIS'],['TEST-LVS','TEST-XFM'],['TEST-LOAD','TEST-LVS']]);
});
test('cable schedule order is irrelevant and all feeds survive',()=>{
  const ref=parse([['Load Name (To)','Panel (From)'],['TEST-PANEL','TEST-GIS'],['TEST-LOAD','TEST-PANEL'],['TEST-LOAD','TEST-SECOND']],'cable');
  const findings=auditEngineeringFindings(snap([{...row,equipmentId:'TEST-LOAD',dependencies:'TEST-PANEL'},{...elec,equipmentId:'TEST-PANEL',closestParent:'TEST-GIS'},{...elec,equipmentId:'TEST-GIS'},{...elec,equipmentId:'TEST-SECOND'}]),{cable:ref});
  assert.equal(findings.length,1);assert.equal(findings[0].expected,'TEST-SECOND');
});
test('PMD repeated rows deduplicate findings and semicolon feeds are retained',()=>{
  const ref=parse([['Panel','Instrument Tag'],['TEST-RIO; TEST-PLC',row.equipmentId],['TEST-RIO',row.equipmentId]],'pmd');
  const findings=auditEngineeringFindings(snap([row,{...fms,equipmentId:'TEST-RIO'},{...fms,equipmentId:'TEST-PLC'}]),{pmd:ref});
  assert.equal(findings.length,1);assert.equal(findings[0].expected,'TEST-RIO; TEST-PLC');
});
test('a cross-group feed needs a local dependency; Closest Parent or a remote dependency does not count',()=>{
  const ref=parse([['Panel','Instrument Tag'],['TEST-RIO',row.equipmentId]],'pmd');
  for(const changes of [{dependencies:'TEST-RIO'},{dependencies:'test-rio',project:'TEST',site:'TEST-SITE',dependencyProject:'TEST;TEST'}])assert.equal(auditEngineeringFindings(snap([{...row,...changes},{...fms,equipmentId:'TEST-RIO'}]),{pmd:ref}).length,0);
  for(const changes of [{closestParent:'TEST-RIO'},{dependencies:'TEST-RIO',project:'TEST',site:'TEST-SITE',dependencyProject:'REMOTE'}]){
    const findings=auditEngineeringFindings(snap([{...row,...changes},{...fms,equipmentId:'TEST-RIO'}]),{pmd:ref});
    assert.equal(findings.length,1);assert.equal(findings[0].rule.id,'reference.engineering-expected-dependency');
  }
});
test('a feed inside one group is nesting, not a missing dependency',()=>{
  const ref=parse([['Panel','Instrument Tag'],['TEST-SKID',row.equipmentId]],'pmd');
  assert.equal(auditEngineeringFindings(snap([row,{...row,equipmentId:'TEST-SKID'}]),{pmd:ref}).length,0);
});
test('the session result carries the dependency check when power or control documents are loaded',()=>{
  const ref=parse([['Load Name (To)','Panel (From)'],[row.equipmentId,'TEST-PANEL']],'cable');
  const result=auditSessionResult(snap([row,{...elec,equipmentId:'TEST-PANEL'}]),{cable:ref});
  assert.deepEqual(result.dependencyCheck.counts,{compared:1,match:0,missing:1,notInRegistry:0});
  assert.equal(result.findings.filter(finding=>finding.rule.id==='reference.engineering-expected-dependency').length,1);
  assert.ok(!result.findings.some(finding=>finding.rule.id==='reference.engineering-relationship'));
});
test('case and whitespace normalize but suffixes remain separate tags',()=>{
  const ref=parse([['Equipment Tag','UPN'],[' test-pump111-01 ','111'],['TEST-PUMP111-01-A','111']]);
  const findings=auditEngineeringFindings(snap([row]),{mel:ref});
  assert.equal(findings.length,1);assert.equal(findings[0].equipmentId,'TEST-PUMP111-01-A');assert.equal(findings[0].row,0);
});
test('only populated source fields are compared; blank registry metadata is visible',()=>{
  const ref=parse([['Equipment Tag','UPN','Building','Discipline'],[row.equipmentId,'111','OTHER','']]);
  const findings=auditEngineeringFindings(snap([{...row,building:''}]),{mel:ref});
  assert.equal(findings.length,1);assert.equal(findings[0].field,'Building');assert.equal(findings[0].expected,'OTHER');assert.match(findings[0].why,/row 2/);
});
test('conflicting MEL values do not produce a correction proposal',()=>{
  const ref=parse([['Equipment Tag','UPN'],[row.equipmentId,'111'],[row.equipmentId,'112']]),snapshot=snap([row]);
  const finding=auditEngineeringFindings(snapshot,{mel:ref})[0];
  assert.equal(finding.rule.id,'reference.engineering-conflicting-mel');
  assert.deepEqual(auditActionPolicy(finding).fields,['UPN']);
  assert.equal(auditProposeCorrection(finding,auditRecommendationContext(snapshot,{mel:ref})),null);
});
test('I&C discrepancy explains why it must not be copied as Exto discipline',()=>{
  const finding=auditEngineeringFindings(snap([row]),{mel:parse([['Equipment Tag','Discipline'],[row.equipmentId,'I&C']])})[0];
  assert.match(finding.recommendation,/Do not copy I&C/);assert.equal(finding.severity,'warning');
});
test('absent MEL tags are scope notes and Blank headers are excluded',()=>{
  const findings=auditEngineeringFindings(snap([row,{...row,equipmentId:'HEADER',itemMaster:'VF_Blank'}]),{mel:parse([['Equipment Tag','UPN'],['OTHER','111']])});
  assert.equal(findings.length,2);assert.ok(findings.every(f=>f.equipmentId!=='HEADER'));assert.ok(findings.every(f=>f.severity==='missing'&&f.category==='missing-tags'));
});
test('session audit works without milestones, preserves no-reference baseline and resolves corrected metadata',()=>{
  const snapshot=snap([row]),refs={mel:parse([['Equipment Tag','Building'],[row.equipmentId,'OTHER']])};
  assert.deepEqual(auditSessionResult(snapshot),runSsmAudit(snapshot));
  const before=auditSessionResult(snapshot,refs);assert.ok(before.findings.some(f=>f.rule.id==='reference.engineering-metadata'));
  const after=auditSessionResult(snap([{...row,building:'OTHER'}]),refs);assert.ok(!after.findings.some(f=>f.rule.id==='reference.engineering-metadata'));
  assert.equal(before.summary.findings,before.findings.length);assert.equal(Object.values(before.summary.severity).reduce((a,b)=>a+b,0),before.findings.length);
});
test('review progress fingerprints engineering inputs without embedding source data',async()=>{
  const snapshot=snap([row]),refs={mel:parse([['Equipment Tag','Building'],[row.equipmentId,'PRIVATE-SYNTHETIC-VALUE']])};
  const document=await auditReviewDocument({snapshot,references:refs});
  assert.ok(!JSON.stringify(document).includes('PRIVATE-SYNTHETIC-VALUE'));
  await auditReadReviewDocument(document,snapshot,refs);
  await assert.rejects(auditReadReviewDocument(document,snapshot,{}),/same reference/);
});
