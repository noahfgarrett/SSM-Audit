import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import {readFileSync} from 'node:fs'
import {EXTO_REV21_COLUMNS} from '../src/exto/rev21-contract.js'
import {auditSnapshotFromAoa} from '../src/audit/model.js'
import {auditMakeCorrection} from '../src/audit/actions.js'
import {AUDIT_UPDATE_BATCH_SIZE,buildAuditUpdateBatches,auditUpdateExportDate,auditUpdateExportSummary} from '../src/audit/export.js'
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
   const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''});assert.deepEqual(rows[0],headers);
   assert.equal(rows.length-1,sizes[i]);assert.ok(rows.length-1<=AUDIT_UPDATE_BATCH_SIZE);
   for(let r=1;r<rows.length;r++){
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
 const tags=XLSX.utils.sheet_to_json(output.Sheets['Upload Template']).map(r=>r['Equipment ID']);
 assert.ok(!tags.includes('DEMO-EQ-0'));assert.ok(!tags.includes('DEMO-EQ-1'));
 assert.match(auditUpdateExportSummary(result.summary),/1,950 equipment rows exported in 1 batch \(1,950\).*2 completed rows excluded.*3,900 changed cells/);
});

test('duplicate physical rows remain together without exceeding a batch limit',async()=>{
 const f=fixture(1951,{duplicateLast:true}),result=await buildAuditUpdateBatches(f.source,f.baseline,f.changes,{sourceWorkbook:f.book});
 assert.deepEqual(result.summary.batches.map(b=>b.rows),[1950,1]);
 const counts=[...entries(result.bytes).values()].map(bytes=>{
  const book=XLSX.read(bytes,{type:'array'});return XLSX.utils.sheet_to_json(book.Sheets['Upload Template']).filter(r=>r['Equipment ID']==='DEMO-EQ-0').length;
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
