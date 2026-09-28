import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { auditEmailDirectory, auditPlanEmailUpdates, buildAuditUpdateRowsBytes } from '../src/audit/export.js'
import { auditMakeCorrection, auditMergeCorrections, auditReviewDocument, auditReadReviewDocument } from '../src/audit/actions.js'
import { auditPrepareInWorker } from '../src/audit/review.js'
import { S, resetSession } from '../src/state.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js',import.meta.url),'utf8'));
function fixture(records=[
    {equipmentId:'DEMO-VFD101-01',discipline:'Facilities Monitoring System',upn:'650'},
    {equipmentId:'DEMO-MAH101-01',discipline:'Mechanical',upn:'101'},
    {equipmentId:'DEMO-UNASSIGNED',discipline:'Unknown',upn:'101'},
  ]){
  records=records.map(row=>({...row,intelPmEmail:'old.pm@example.com',superintendentEmail:'old.sup@example.com',cxEngineerEmail:'old.cx@example.com',dependencies:'DEMO-FEED'}));
  const aoa=[EXTO_REV21_COLUMNS.map(c=>c.header),...records.map(row=>EXTO_REV21_COLUMNS.map(c=>row[c.field]||''))];
  const baseline=auditSnapshotFromAoa(aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([
    ['Discipline','Intel PM Email Address','Superintendent Email Address','Cx Engineer Email Address'],
    ['Facilities Monitoring System','fms.pm@example.com','fms.sup@example.com','fms.cx@example.com'],
    ['Mechanical','mech.pm@example.com','','mech.cx@example.com'],
    ['LIFE SAFETY SYSTEM','life.pm@example.com','life.sup@example.com','life.cx@example.com'],
    ['SECURITY','security.pm@example.com','security.sup@example.com','security.cx@example.com'],
  ]),'Emails');
  const bytes=XLSX.write(book,{type:'array',bookType:'xlsx'});
  return {baseline,book,bytes,directory:auditEmailDirectory(book)};
}

for(const [upn,discipline,prefix] of [['630','LIFE SAFETY SYSTEM','life'],['650','FACILITIES MONITORING SYSTEM','fms'],['SEC','SECURITY','security']]){
  test(`imported UPN ${upn} overrides conflicting discipline in manual updates, automatic updates and exports`,async()=>{
    const f=fixture([{equipmentId:'DEMO-TAG',discipline:'Mechanical',upn}]),row=f.baseline.rows[0];
    const manual=auditPlanEmailUpdates(f.baseline,[],f.directory,{scope:'all'});
    assert.equal(manual.emailUpdates.length,3);
    assert.ok(manual.emailUpdates.every(c=>c.value.startsWith(prefix+'.')));
    assert.ok(manual.emailUpdates.every(c=>c.reason.includes(discipline)));
    const edits=[auditMakeCorrection(row,'UPN','101'),auditMakeCorrection(row,'Discipline','Electrical')];
    const automatic=auditPlanEmailUpdates(f.baseline,edits,f.directory,{scope:'changed',previousChanges:[]});
    assert.equal(automatic.emailUpdates.length,3);
    assert.ok(automatic.emailUpdates.every(c=>c.value.startsWith(prefix+'.')),'metadata edits do not change imported routing');
    for(const [changes,updateEmails] of [[automatic.changes,false],[edits,true]]){
      const bytes=await buildAuditUpdateRowsBytes(f.bytes,f.baseline,changes,{uploadTemplate:true,updateEmails});
      const sheet=XLSX.read(bytes,{type:'array',cellStyles:true}).Sheets['Upload Template'];
      for(const [field,role] of [['intelPmEmail','pm'],['superintendentEmail','sup'],['cxEngineerEmail','cx']]){
        const cell=sheet[XLSX.utils.encode_cell({r:2,c:EXTO_REV21_COLUMNS.find(c=>c.field===field).index})];
        assert.equal(cell.v,`${prefix}.${role}@example.com`);assert.equal(cell.s.fgColor.rgb,'FFF2CC');
      }
    }
    assert.equal(f.baseline.rows[0].discipline,'Mechanical');assert.equal(f.baseline.rows[0].upn,upn);
  });
}

test('UPN email overrides normalize case and spaces without matching other identifiers',()=>{
  const values=[[630,'life'],[' 650 ','fms'],[' sec ','security'],['SEC-1','mech'],['1630','mech'],['6500','mech'],['','mech']];
  const f=fixture(values.map(([upn],i)=>({equipmentId:`DEMO-${i}`,upn,discipline:'Mechanical'})));
  const plan=auditPlanEmailUpdates(f.baseline,[],f.directory);
  values.forEach(([,prefix],i)=>assert.equal(plan.changes.find(c=>c.tag===`DEMO-${i}`&&c.prop==='intelPmEmail').value,`${prefix}.pm@example.com`));
});

test('missing override directory entries never fall back to a conflicting discipline',()=>{
  const f=fixture(['630','650','SEC'].map((upn,i)=>({equipmentId:`DEMO-${i}`,upn,discipline:'Mechanical'})));
  for(const key of ['LIFE SAFETY SYSTEM','FACILITIES MONITORING SYSTEM','SECURITY'])f.directory.byDiscipline.delete(key);
  const plan=auditPlanEmailUpdates(f.baseline,[],f.directory);
  assert.equal(plan.changes.length,0);
  assert.deepEqual(plan.emailSummary.missingDisciplines,['FACILITIES MONITORING SYSTEM','LIFE SAFETY SYSTEM','SECURITY']);
});

test('UPN override copies blank email cells and includes completed equipment',()=>{
  const f=fixture([{equipmentId:'DEMO-DONE',upn:'630',discipline:'Mechanical'},{equipmentId:'DEMO-OPEN',upn:'SEC',discipline:''}]);
  f.directory.byDiscipline.get('SECURITY').superintendentEmail='';
  const plan=auditPlanEmailUpdates(f.baseline,[],f.directory,{completedEquipmentIds:['demo-done']});
  assert.equal(plan.emailUpdates.length,6);
  assert.equal(plan.emailUpdates.filter(c=>c.tag==='DEMO-DONE').length,3);
  assert.equal(plan.emailUpdates.find(c=>c.tag==='DEMO-OPEN'&&c.prop==='superintendentEmail').value,'');
});

test('changing a non-override UPN to an override preserves its original discipline emails',()=>{
  const f=fixture([{equipmentId:'DEMO-1',upn:'101',discipline:'Mechanical'}]),row=f.baseline.rows[0];
  for(const upn of ['630','650','SEC']){
    const plan=auditPlanEmailUpdates(f.baseline,[auditMakeCorrection(row,'UPN',upn)],f.directory,{scope:'changed'});
    assert.equal(plan.emailUpdates.length,3);assert.ok(plan.emailUpdates.every(c=>c.value===''||c.value.startsWith('mech.')));
  }
});

test('review worker applies UPN email precedence to each newly affected row',async()=>{
  const f=fixture(['630','650','SEC'].map((upn,i)=>({equipmentId:`DEMO-${i}`,upn,discipline:'Mechanical'})));
  const changes=f.baseline.rows.map(row=>auditMakeCorrection(row,'UPN','101'));
  const prepared=await auditPrepareInWorker({},{baseline:f.baseline,file:new Blob([f.bytes]),changes,previousChanges:[],references:{},emailUpdateScope:'changed'});
  assert.equal(prepared.emailSummary.changedCells,9);
  assert.deepEqual(prepared.snapshot.rows.map(row=>row.intelPmEmail),['life.pm@example.com','fms.pm@example.com','security.pm@example.com']);
  assert.ok(prepared.snapshot.rows.every(row=>row.upn==='101'));
});

test('email settings start off; explicit all-row update ignores filters and includes completed equipment',()=>{
  resetSession();assert.equal(S.session.updateEmails,false);
  const f=fixture(),before=structuredClone(f.baseline.rows);
  const plan=auditPlanEmailUpdates(f.baseline,[],f.directory,{scope:'all',completedEquipmentIds:['DEMO-MAH101-01']});
  assert.equal(plan.changes.length,6);assert.equal(plan.emailSummary.matchedRows,2);
  assert.equal(plan.changes.filter(c=>c.tag==='DEMO-MAH101-01').length,3);
  assert.deepEqual(plan.emailSummary.missingDisciplines,['Unknown']);
  assert.deepEqual(f.baseline.rows,before);
});

test('FMS-to-Mechanical corrections retain imported FMS email matching',()=>{
  const f=fixture(),row=f.baseline.rows[0];
  const edits=[auditMakeCorrection(row,'Discipline','Mechanical'),auditMakeCorrection(row,'UPN','101')];
  const plan=auditPlanEmailUpdates(f.baseline,edits,f.directory,{scope:'changed',previousChanges:[]});
  assert.equal(plan.emailUpdates.length,3);
  assert.ok(plan.emailUpdates.every(c=>c.value.startsWith('fms.')));
  assert.equal(plan.changes.length,5);
  assert.equal(f.baseline.rows[0].discipline,'Facilities Monitoring System');
});

test('automatic email updates touch only newly affected equipment, including cleared metadata',()=>{
  const f=fixture(),previous=[auditMakeCorrection(f.baseline.rows[0],'UPN','101')];
  const changes=auditMergeCorrections(f.baseline,previous,[auditMakeCorrection(f.baseline.rows[1],'Dependencies','')]);
  const plan=auditPlanEmailUpdates(f.baseline,changes,f.directory,{scope:'changed',previousChanges:previous});
  assert.equal(plan.emailUpdates.length,3);
  assert.ok(plan.emailUpdates.every(c=>c.tag==='DEMO-MAH101-01'));
  assert.equal(plan.emailUpdates.find(c=>c.prop==='superintendentEmail').value,'','blank directory entry clears the old address');
});

test('directory values replace prior draft email edits in all-row and opted-in changed-row updates',()=>{
  const f=fixture(),row=f.baseline.rows[0];
  const prior=[auditMakeCorrection(row,'Intel PM Email Address','manual.pm@example.com')];
  const all=auditPlanEmailUpdates(f.baseline,prior,f.directory,{scope:'all'});
  assert.equal(all.changes.find(c=>c.tag===row.equipmentId&&c.prop==='intelPmEmail').value,'fms.pm@example.com');
  for(const [changes,previousChanges] of [[prior,[]],[prior,prior]]){
    const edits=changes===previousChanges?[...changes,auditMakeCorrection(row,'Building','UPDATED')]:changes;
    const plan=auditPlanEmailUpdates(f.baseline,edits,f.directory,{scope:'changed',previousChanges});
    assert.equal(plan.emailUpdates.length,3);
    assert.ok(plan.emailUpdates.every(c=>c.tag===row.equipmentId));
    assert.equal(plan.emailUpdates.find(c=>c.prop==='intelPmEmail').before,'manual.pm@example.com');
    assert.equal(plan.changes.find(c=>c.prop==='intelPmEmail').value,'fms.pm@example.com');
  }
});

test('blank directory rows clear all three emails and a second synchronization makes no changes',async()=>{
  const f=fixture([{equipmentId:'DEMO-1',upn:'SEC',discipline:'Mechanical'}]);
  f.directory.byDiscipline.set('SECURITY',{intelPmEmail:'',superintendentEmail:'',cxEngineerEmail:''});
  const plan=auditPlanEmailUpdates(f.baseline,[],f.directory);
  assert.equal(plan.emailUpdates.length,3);assert.ok(plan.changes.every(c=>c.value===''));
  const again=auditPlanEmailUpdates(f.baseline,plan.changes,f.directory);
  assert.equal(again.emailUpdates.length,0);assert.deepEqual(again.changes,plan.changes);
  const bytes=await buildAuditUpdateRowsBytes(f.bytes,f.baseline,plan.changes,{uploadTemplate:true});
  const sheet=XLSX.read(bytes,{type:'array',cellStyles:true}).Sheets['Upload Template'];
  for(const field of ['intelPmEmail','superintendentEmail','cxEngineerEmail']){
    const cell=sheet[XLSX.utils.encode_cell({r:2,c:EXTO_REV21_COLUMNS.find(c=>c.field===field).index})];
    assert.equal(cell.v,'');assert.equal(cell.s.fgColor.rgb,'FFF2CC');
  }
});

test('all-row email sync changes completed rows in the draft but completed rows remain excluded from exports',async()=>{
  const f=fixture(),cache={},completedEquipmentIds=['DEMO-MAH101-01'];
  const prepared=await auditPrepareInWorker(cache,{baseline:f.baseline,file:new Blob([f.bytes]),changes:[],previousChanges:[],references:{},emailUpdateScope:'all',completedEquipmentIds});
  assert.equal(prepared.snapshot.rows[1].intelPmEmail,'mech.pm@example.com');
  assert.equal(prepared.snapshot.rows[1].superintendentEmail,'');
  const output=await buildAuditUpdateRowsBytes(f.bytes,f.baseline,prepared.changes,{uploadTemplate:true,completedEquipmentIds});
  const sheet=XLSX.read(output,{type:'array'}).Sheets['Upload Template'];
  const rows=XLSX.utils.sheet_to_json(sheet,{range:1});
  assert.deepEqual(rows.map(row=>row['Equipment ID']),['DEMO-VFD101-01']);
});

test('missing columns and unknown disciplines are reported without guessed addresses',()=>{
  const f=fixture(),row={...f.baseline.rows[0],_source:{...f.baseline.rows[0]._source,columns:{...f.baseline.rows[0]._source.columns}}};
  delete row._source.columns.cxEngineerEmail;
  const plan=auditPlanEmailUpdates({...f.baseline,rows:[row,f.baseline.rows[2]]},[],f.directory,{scope:'all'});
  assert.equal(plan.changes.length,2);
  assert.deepEqual(plan.emailSummary.unavailableFields,['Cx Engineer Email Address']);
  assert.deepEqual(plan.emailSummary.missingDisciplines,['Unknown']);
  assert.throws(()=>auditPlanEmailUpdates(f.baseline,[],null),/Emails tab/);
});

test('conflicting directory entries fail closed and identical duplicates are harmless',()=>{
  const f=fixture(),rows=XLSX.utils.sheet_to_json(f.book.Sheets.Emails,{header:1});
  rows.push([...rows[1]]);f.book.Sheets.Emails=XLSX.utils.aoa_to_sheet(rows);
  assert.equal(auditEmailDirectory(f.book).error,'');
  rows.at(-1)[1]='conflict@example.com';f.book.Sheets.Emails=XLSX.utils.aoa_to_sheet(rows);
  assert.throws(()=>auditPlanEmailUpdates(f.baseline,[],auditEmailDirectory(f.book)),/conflicting entries/);
});

test('review worker stages all email changes explicitly and restores exact drafts without resyncing',async()=>{
  const f=fixture(),cache={},input={baseline:f.baseline,file:new Blob([f.bytes]),changes:[],previousChanges:[],references:{}};
  const untouched=await auditPrepareInWorker(cache,input);
  assert.equal(untouched.snapshot.rows[0].intelPmEmail,'old.pm@example.com');
  const prepared=await auditPrepareInWorker(cache,{changes:[],previousChanges:[],references:{},emailUpdateScope:'all'});
  assert.equal(prepared.changes.length,6);assert.equal(prepared.emailSummary.changedCells,6);
  assert.equal(prepared.snapshot.rows[0].intelPmEmail,'fms.pm@example.com');
  assert.equal(prepared.impact.unsafe.length,0);
  const restored=await auditPrepareInWorker(cache,{changes:[],previousChanges:prepared.changes,references:{}});
  assert.equal(restored.snapshot.rows[0].intelPmEmail,'old.pm@example.com');
});

test('email corrections and opt-in setting survive saved review; older reviews remain off',async()=>{
  const f=fixture(),plan=auditPlanEmailUpdates(f.baseline,[],f.directory);
  const saved=await auditReviewDocument({baselineSnapshot:f.baseline,changes:plan.changes,updateEmails:true});
  const restored=await auditReadReviewDocument(saved,f.baseline);
  assert.equal(restored.updateEmails,true);assert.equal(restored.snapshot.rows[0].cxEngineerEmail,'fms.cx@example.com');
  delete saved.updateEmails;
  assert.equal((await auditReadReviewDocument(saved,f.baseline)).updateEmails,false);
  saved.updateEmails='yes';await assert.rejects(auditReadReviewDocument(saved,f.baseline),/email setting/);
});

test('explicit email-only edits export all metadata and yellow cells with automatic export updates off',async()=>{
  const f=fixture(),plan=auditPlanEmailUpdates(f.baseline,[],f.directory);
  const bytes=await buildAuditUpdateRowsBytes(f.bytes,f.baseline,plan.changes,{uploadTemplate:true});
  const sheet=XLSX.read(bytes,{type:'array',cellStyles:true}).Sheets['Upload Template'];
  const col=field=>EXTO_REV21_COLUMNS.find(c=>c.field===field).index;
  const cell=(r,field)=>sheet[XLSX.utils.encode_cell({r,c:col(field)})];
  assert.equal(XLSX.utils.decode_range(sheet['!ref']).e.r,3);
  assert.equal(cell(2,'intelPmEmail').v,'fms.pm@example.com');
  assert.equal(cell(2,'intelPmEmail').s.fgColor.rgb,'FFF2CC');
  assert.equal(cell(2,'discipline').v,'Facilities Monitoring System');
  assert.equal(cell(2,'dependencies').v,'DEMO-FEED');
  assert.equal(cell(3,'superintendentEmail').v,'');
  assert.equal(cell(3,'superintendentEmail').s.fgColor.rgb,'FFF2CC');
});
