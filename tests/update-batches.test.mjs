import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import {readFileSync} from 'node:fs'
import {EXTO_REV21_COLUMNS} from '../src/exto/rev21-contract.js'
import {auditSnapshotFromAoa} from '../src/audit/model.js'
import {auditMakeCorrection} from '../src/audit/actions.js'
import {AUDIT_UPDATE_BATCH_SIZE,buildAuditUpdateBatches,auditUpdateExportDate,auditUpdateExportSummary,auditEmailDirectory} from '../src/audit/export.js'
import {crc32,zipEntries} from '../src/core/zip.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js',import.meta.url),'utf8'));
const headers=EXTO_REV21_COLUMNS.map(c=>c.header);
function fixture(count,{duplicateLast=false}={}){
 const aoa=[headers];
 for(let i=0;i<count;i++){
  const values=Object.fromEntries(EXTO_REV21_COLUMNS.map(c=>[c.field,`Metadata ${i} ${c.field}`]));
  values.equipmentId=`DEMO-EQ-${duplicateLast&&i===count-1?0:i}`;values.upn='602';values.dependencyProject='DEMO';
  aoa.push(EXTO_REV21_COLUMNS.map(c=>values[c.field]));
 }
 const baseline=auditSnapshotFromAoa(aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['Not for export']]),'Other');
 const source=new Uint8Array(XLSX.write(book,{bookType:'xlsx',type:'buffer'}));
 return {source,baseline,book,aoa,changes:baseline.rows.flatMap(row=>[auditMakeCorrection(row,'Dependency Project',''),auditMakeCorrection(row,'UPN','603')])};
}
function entries(bytes){
 const archive=XLSX.CFB.read(bytes,{type:'array'}),result=new Map();
 archive.FileIndex.forEach((file,i)=>{
  const name=archive.FullPaths[i].slice(archive.FullPaths[0].length);
  if(file.type===2&&name&&!name.startsWith('\u0001'))result.set(name,new Uint8Array(file.content));
 });
 return result;
}

test('VF upload rows receive Yes using final item masters with yellow changes and accurate counts',async()=>{
 const f=fixture(7),item=26,association=38;
 const values=[['VF1_PANEL',''],['VF_Blank','No'],['VF2_PANEL','Yes'],['DEMO_PANEL','No'],['DEMO_PANEL',''],['VF1_PANEL','No'],['VFD_PANEL','No']];
 values.forEach(([im,fat],i)=>{f.aoa[i+1][item]=im;f.aoa[i+1][association]=fat;});
 const baseline=auditSnapshotFromAoa(f.aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(f.aoa),'Registry');
 const source=new Uint8Array(XLSX.write(book,{bookType:'xlsx',type:'buffer'})),original=source.slice();
 const changes=baseline.rows.map(row=>auditMakeCorrection(row,'UPN','603'));
 changes.push(auditMakeCorrection(baseline.rows[4],'Item Master Unique Identifier','VF3_PANEL'));
 changes.push(auditMakeCorrection(baseline.rows[5],'Item Master Unique Identifier','DEMO_PANEL'));
 const result=await buildAuditUpdateBatches(source,baseline,changes);
 const output=XLSX.read([...entries(result.bytes).values()][0],{type:'array',cellStyles:true}).Sheets['Upload Template'];
 assert.deepEqual(values.map((_,i)=>output[`AM${i+3}`]?.v),['Yes','Yes','Yes','No','Yes','No','No']);
 for(const row of [3,4,7])assert.equal(output[`AM${row}`].s.fgColor.rgb,'FFF2CC');
 assert.notEqual(output.AM5.s?.fgColor?.rgb,'FFF2CC','existing Yes is unchanged');
 assert.equal(result.summary.exportedCells,12);assert.deepEqual(source,original);
});

test('VF association is supplied when the source has no association column',async()=>{
 const f=fixture(1);f.aoa[1][26]='VF1_PANEL';
 const aoa=f.aoa.map(row=>row.filter((_,index)=>index!==38));
 const baseline=auditSnapshotFromAoa(aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
 const source=new Uint8Array(XLSX.write(book,{bookType:'xlsx',type:'buffer'}));
 const result=await buildAuditUpdateBatches(source,baseline,[auditMakeCorrection(baseline.rows[0],'UPN','603')]);
 const sheet=XLSX.read([...entries(result.bytes).values()][0],{type:'array',cellStyles:true}).Sheets['Upload Template'];
 assert.equal(sheet.AM3.v,'Yes');assert.equal(sheet.AM3.s.fgColor.rgb,'FFF2CC');assert.equal(result.summary.exportedCells,2);
});

for(const [count,sizes] of [[1,[1]],[1950,[1950]],[1951,[1950,1]],[3900,[1950,1950]],[4300,[1950,1950,400]]]){
 test(`upload export batches ${count} rows without loss, metadata changes or overlap`,async()=>{
  const f=fixture(count),original=f.source.slice(),stages=[];
  const result=await buildAuditUpdateBatches(f.source,f.baseline,f.changes,{exportDate:'2026-09-11',sourceWorkbook:f.book,onStage:f=>stages.push(f)});
  assert.equal(result.filename,'Registry_Automated_Update_2026-09-11.zip');
  assert.deepEqual(result.summary.batches.map(b=>b.rows),sizes);
  assert.equal(result.summary.exportedRows,count);assert.equal(result.summary.exportedCells,count*2);
  const files=entries(result.bytes),seen=new Set();assert.equal(files.size,sizes.length);
  for(const [i,batch] of result.summary.batches.entries()){
   assert.equal(batch.name,`Registry_Automated_Update_2026-09-11_Batch${String(i+1).padStart(2,'0')}_of_${String(sizes.length).padStart(2,'0')}.xlsx`);
   const book=XLSX.read(files.get(batch.name),{type:'array',cellStyles:true}),sheet=book.Sheets['Upload Template'];
   assert.deepEqual(book.SheetNames,['Upload Template']);
   const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''});
   assert.deepEqual(rows[0],EXTO_REV21_COLUMNS.map(c=>c.gating?'Gating':'Non Gating'));assert.deepEqual(rows[1],headers);
   assert.equal(rows.length-2,sizes[i]);assert.ok(rows.length-2<=AUDIT_UPDATE_BATCH_SIZE);
   assert.equal(sheet['!autofilter'].ref,`A2:AR${sizes[i]+2}`);
   for(const c of EXTO_REV21_COLUMNS)assert.equal(sheet[`${XLSX.utils.encode_col(c.index)}1`].s.fgColor.rgb,c.gating?'FFC7CE':'FFEB9C');
   const reimport=auditSnapshotFromAoa(rows,{sheet:'Upload Template'});
   assert.equal(reimport.headerRow,2);assert.equal(reimport.rows.length,sizes[i]);assert.equal(reimport.rows[0]._source.row,3);
   for(let r=2;r<rows.length;r++){
    const tag=sheet[`K${r+1}`].v;assert.ok(!seen.has(tag));seen.add(tag);
    const input=f.aoa[Number(tag.split('-').at(-1))+1];
    for(const c of EXTO_REV21_COLUMNS){
     const cell=sheet[XLSX.utils.encode_cell({r,c:c.index})],changed=['upn','dependencyProject'].includes(c.field);
     assert.equal(cell?.v??'',c.field==='upn'?'603':c.field==='dependencyProject'?'':input[c.index],c.field);
     if(changed)assert.equal(cell.s.fgColor.rgb,'FFF2CC');
     else assert.notEqual(cell?.s?.fgColor?.rgb,'FFF2CC');
    }
   }
  }
  assert.equal(seen.size,count);assert.deepEqual(f.source,original);
  assert.ok(stages.every((value,i)=>!i||value>=stages[i-1]),'progress never moves backward');
 });
}

test('completed rows are removed before batching and summary distinguishes cells from rows',async()=>{
 const f=fixture(1952),result=await buildAuditUpdateBatches(f.source,f.baseline,f.changes,{sourceWorkbook:f.book,completedEquipmentIds:['demo-eq-0','DEMO-EQ-1']});
 assert.deepEqual(result.summary.batches.map(b=>b.rows),[1950]);
 assert.equal(result.summary.excludedCompletedRows,2);assert.equal(result.summary.exportedCells,3900);
 const output=XLSX.read([...entries(result.bytes).values()][0],{type:'array'});
 const tags=XLSX.utils.sheet_to_json(output.Sheets['Upload Template'],{range:1}).map(r=>r['Equipment ID']);
 assert.equal(tags.length,1950);
 assert.ok(!tags.includes('DEMO-EQ-0'));assert.ok(!tags.includes('DEMO-EQ-1'));
 assert.match(auditUpdateExportSummary(result.summary),/1,950 equipment rows exported in 1 batch \(1,950\).*2 completed rows excluded.*3,900 changed cells/);
});

test('duplicate physical rows remain together without exceeding a batch limit',async()=>{
 const f=fixture(1951,{duplicateLast:true}),result=await buildAuditUpdateBatches(f.source,f.baseline,f.changes,{sourceWorkbook:f.book});
 assert.deepEqual(result.summary.batches.map(b=>b.rows),[1950,1]);
 const counts=[...entries(result.bytes).values()].map(bytes=>{
  const book=XLSX.read(bytes,{type:'array'});return XLSX.utils.sheet_to_json(book.Sheets['Upload Template'],{range:1}).filter(r=>r['Equipment ID']==='DEMO-EQ-0').length;
 });
 assert.deepEqual(counts,[2,0]);
});

test('empty and conflicting exports fail rather than creating partial batches',async()=>{
 const f=fixture(1951);
 await assert.rejects(buildAuditUpdateBatches(f.source,f.baseline,f.changes,{sourceWorkbook:f.book,completedEquipmentIds:f.baseline.rows.map(r=>r.equipmentId)}),/No changed, incomplete equipment/);
 f.changes.at(-1).before='Conflict';
 await assert.rejects(buildAuditUpdateBatches(f.source,f.baseline,f.changes,{sourceWorkbook:f.book}),/original|match|changed|conflict/i);
 const small=fixture(1);
 await assert.rejects(buildAuditUpdateBatches(small.source,small.baseline,small.changes,{exportDate:'../../unsafe'}),/export date is invalid/);
 assert.equal(auditUpdateExportDate(new Date(2026,8,11,23,59)),'2026-09-11');
});

test('stored outer ZIP works without browser compression and has correct CRC and method',async t=>{
 const original=globalThis.CompressionStream;globalThis.CompressionStream=undefined;t.after(()=>{globalThis.CompressionStream=original;});
 const data=new Uint8Array([1,2,3,4]),bytes=await zipEntries([{name:'batch.xlsx',data}],null,{store:true});
 const header=new DataView(bytes.buffer);assert.equal(header.getUint16(8,true),0);assert.equal(header.getUint32(14,true),crc32(data));
 assert.deepEqual(entries(bytes).get('batch.xlsx'),data);
});

test('a two-row upload export can be reloaded and corrected again without shifting milestone cells',async()=>{
 const f=fixture(1),first=await buildAuditUpdateBatches(f.source,f.baseline,f.changes,{sourceWorkbook:f.book});
 const source=[...entries(first.bytes).values()][0],book=XLSX.read(source,{type:'array',cellStyles:true});
 const baseline=auditSnapshotFromAoa(XLSX.utils.sheet_to_json(book.Sheets['Upload Template'],{header:1,defval:''}),{sheet:'Upload Template'});
 const changes=[auditMakeCorrection(baseline.rows[0],'L1 Milestone Parent','DEMO-L1-M1-10'),auditMakeCorrection(baseline.rows[0],'L2 Milestone','DEMO-L2-M1-20')];
 const next=await buildAuditUpdateBatches(source,baseline,changes,{sourceWorkbook:book});
 const output=XLSX.read([...entries(next.bytes).values()][0],{type:'array',cellStyles:true}),sheet=output.Sheets['Upload Template'];
 assert.deepEqual(output.SheetNames,['Upload Template']);assert.equal(sheet.K3.v,'DEMO-EQ-0');
 assert.equal(sheet.Y3.v,'DEMO-L2-M1-20');assert.equal(sheet.Z3.v,'DEMO-L1-M1-10');
 assert.equal(sheet.Y3.s.fgColor.rgb,'FFF2CC');assert.equal(sheet.Z3.s.fgColor.rgb,'FFF2CC');
 assert.equal(sheet.Y1.v,'Non Gating');assert.equal(sheet.Z1.v,'Non Gating');
 assert.equal(sheet.Y2.v,'Milestone');assert.equal(sheet.Z2.v,'Milestone Parent');
 assert.equal(next.summary.exportedRows,1);assert.equal(next.summary.exportedCells,2);
});

test('an Emails tab fills the three email columns by corrected discipline, in yellow, and reports misses',async()=>{
 const f=fixture(5),discipline=7,[pm,sup,cx]=[29,30,31];
 const disciplines=['Mechanical','ELECTRICAL','I&C','Mechanical','Structural'];
 disciplines.forEach((value,i)=>{f.aoa[i+1][discipline]=value;});
 f.aoa[4][pm]='pm.mech@example.com';f.aoa[4][sup]='old.sup@example.com';   // row 4: PM already right, superintendent stale
 const baseline=auditSnapshotFromAoa(f.aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(f.aoa),'Registry');
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([
  ['Discipline','Intel PM Email Address','Superintendent Email Address','Cx Engineer Email Address'],
  ['MECHANICAL','pm.mech@example.com','sup.mech@example.com','cx.mech@example.com'],
  ['Electrical','pm.elec@example.com','sup.elec@example.com',''],
  ['Facilities Monitoring System','pm.fms@example.com','sup.fms@example.com','cx.fms@example.com'],
 ]),'Emails');
 const source=new Uint8Array(XLSX.write(book,{bookType:'xlsx',type:'buffer'}));
 const directory=auditEmailDirectory(XLSX.read(source,{type:'array'}));
 assert.equal(directory.sheet,'Emails');assert.equal(directory.rows,3);assert.equal(directory.byDiscipline.get('MECHANICAL').cxEngineerEmail,'cx.mech@example.com');
 const changes=baseline.rows.map(row=>auditMakeCorrection(row,'UPN','603'));
 changes.push(auditMakeCorrection(baseline.rows[4],'Discipline','ELECTRICAL'));   // structural row corrected to electrical
 const result=await buildAuditUpdateBatches(source,baseline,changes);
 const sheet=XLSX.read([...entries(result.bytes).values()][0],{type:'array',cellStyles:true}).Sheets['Upload Template'];
 const cell=(r,c)=>sheet[XLSX.utils.encode_cell({r:r+2,c})];
 assert.equal(cell(0,pm).v,'pm.mech@example.com');assert.equal(cell(0,sup).v,'sup.mech@example.com');assert.equal(cell(0,cx).v,'cx.mech@example.com');
 for(const c of [pm,sup,cx])assert.equal(cell(0,c).s.fgColor.rgb,'FFF2CC');
 assert.equal(cell(1,pm).v,'pm.elec@example.com');assert.equal(cell(1,cx).v,f.aoa[2][cx],'a blank Emails cell leaves the original value');assert.notEqual(cell(1,cx).s?.fgColor?.rgb,'FFF2CC');
 assert.equal(cell(2,pm).v,'pm.fms@example.com','I&C maps to the FMS discipline');
 assert.equal(cell(3,pm).v,'pm.mech@example.com');assert.notEqual(cell(3,pm).s?.fgColor?.rgb,'FFF2CC','an address already in place is not a change');
 assert.equal(cell(3,sup).v,'sup.mech@example.com');assert.equal(cell(3,sup).s.fgColor.rgb,'FFF2CC');
 assert.equal(cell(4,pm).v,'pm.elec@example.com','the corrected discipline decides the emails');assert.equal(cell(4,discipline).v,'ELECTRICAL');
 assert.equal(result.summary.emailCells,3+2+3+2+2);assert.equal(result.summary.exportedCells,6+12);
 assert.deepEqual(result.summary.emailMisses,[]);assert.equal(result.summary.emailsTab,'Emails');
 assert.match(auditUpdateExportSummary(result.summary),/12 email cells filled from the Emails tab/);
 const stranger=await buildAuditUpdateBatches(source,baseline,[auditMakeCorrection(baseline.rows[4],'UPN','603')]);
 assert.deepEqual(stranger.summary.emailMisses,['Structural']);assert.match(auditUpdateExportSummary(stranger.summary),/No email entry for: Structural/);
 const plain=await buildAuditUpdateBatches(f.source,f.baseline,f.changes.slice(0,2));
 assert.equal(plain.summary.emailsTab,'');assert.doesNotMatch(auditUpdateExportSummary(plain.summary),/email/);
});

test('a malformed Emails tab stops the export instead of writing partial addresses',async()=>{
 const f=fixture(1),book=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(f.aoa),'Registry');
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['Discipline','Intel PM Email Address'],['Mechanical','pm@example.com']]),'Emails');
 const source=new Uint8Array(XLSX.write(book,{bookType:'xlsx',type:'buffer'}));
 await assert.rejects(buildAuditUpdateBatches(source,f.baseline,f.changes.slice(0,1)),/Emails tab is missing SUPERINTENDENT EMAIL ADDRESS, CX ENGINEER EMAIL ADDRESS/);
});
