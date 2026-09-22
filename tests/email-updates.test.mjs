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
function fixture(){
  const records=[
    {equipmentId:'DEMO-VFD101-01',discipline:'Facilities Monitoring System',upn:'650'},
    {equipmentId:'DEMO-MAH101-01',discipline:'Mechanical',upn:'101'},
    {equipmentId:'DEMO-UNASSIGNED',discipline:'Unknown',upn:'101'},
  ].map(row=>({...row,intelPmEmail:'old.pm@example.com',superintendentEmail:'old.sup@example.com',cxEngineerEmail:'old.cx@example.com',dependencies:'DEMO-FEED'}));
  const aoa=[EXTO_REV21_COLUMNS.map(c=>c.header),...records.map(row=>EXTO_REV21_COLUMNS.map(c=>row[c.field]||''))];
  const baseline=auditSnapshotFromAoa(aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([
    ['Discipline','Intel PM Email Address','Superintendent Email Address','Cx Engineer Email Address'],
    ['Facilities Monitoring System','fms.pm@example.com','fms.sup@example.com','fms.cx@example.com'],
    ['Mechanical','mech.pm@example.com','','mech.cx@example.com'],
  ]),'Emails');
  const bytes=XLSX.write(book,{type:'array',bookType:'xlsx'});
  return {baseline,book,bytes,directory:auditEmailDirectory(book)};
}

test('email settings start off; explicit all-row update ignores filters and skips completed equipment',()=>{
  resetSession();assert.equal(S.session.updateEmails,false);
  const f=fixture(),before=structuredClone(f.baseline.rows);
  const plan=auditPlanEmailUpdates(f.baseline,[],f.directory,{scope:'all',completedEquipmentIds:['DEMO-MAH101-01']});
  assert.equal(plan.changes.length,3);assert.ok(plan.changes.every(c=>c.tag==='DEMO-VFD101-01'));
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
  assert.equal(plan.emailUpdates.length,2);
  assert.ok(plan.emailUpdates.every(c=>c.tag==='DEMO-MAH101-01'));
  assert.ok(!plan.emailUpdates.some(c=>c.prop==='superintendentEmail'),'blank directory entry is not a deletion');
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
  assert.equal(prepared.changes.length,5);assert.equal(prepared.emailSummary.changedCells,5);
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
  assert.equal(cell(3,'superintendentEmail').v,'old.sup@example.com');
});
