import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import {readFileSync} from 'node:fs'
import {buildAuditActionsWorkbook,AUDIT_EXPORT_TICK,AUDIT_EXPORT_UNTICKED} from '../src/audit/export.js'
import {workbookBytesCompact} from '../src/core/download.js'
import {runSsmAudit} from '../src/audit/engine.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js',import.meta.url),'utf8'));
const rule={id:'parent.cross-upn',title:'Parent is in a different UPN',statement:'Parent and child share a UPN.'};
function fixture(){
  const rows=[1,2,3,4].map((n)=>({equipmentId:`DEMO-${n}`,equipmentDescription:'Instrument',closestParent:'DEMO-PARENT',dependencies:'DEMO-PANEL',upn:n<3?'650':'603',systemName:'Demo system',discipline:'Demo discipline',building:'Demo building',milestoneParent:'DEMO-L1-1',milestone:'DEMO-L2-1',_source:{sheet:'Registry',row:n+1}}));
  const findings=rows.map((row,at)=>({id:String(at),rule,severity:'error',equipmentId:row.equipmentId,sheet:'Registry',row:at+2,why:at<2?'Child UPN 650, parent UPN 101':'Child UPN 603, parent UPN 602',actual:row.upn,expected:at<2?'101':'602',recommendation:'Verify the assigned system.',field:'UPN'}));
  return {rows,findings};
}
test('Actions workbook keeps each pattern together, alternates whole groups and provides a tick per finding',async()=>{
  const result=fixture(),before=JSON.stringify(result),book=buildAuditActionsWorkbook(result,'Demo');
  assert.deepEqual(book.SheetNames,['Actionable','Index',rule.title,'_Action queue']);const sheet=book.Sheets[rule.title];
  assert.equal(sheet.A6.v,AUDIT_EXPORT_UNTICKED);assert.equal(sheet.A9.v,AUDIT_EXPORT_UNTICKED);
  assert.equal(sheet.C6.v,sheet.C7.v);assert.equal(sheet.C8.v,sheet.C9.v);assert.notEqual(sheet.C7.v,sheet.C8.v);
  assert.equal(sheet.D6.s.fill.fgColor.rgb,'FFFFFF');assert.equal(sheet.D7.s.fill.fgColor.rgb,'FFFFFF');assert.equal(sheet.D8.s.fill.fgColor.rgb,'F2F2F2');
  assert.equal(sheet.W8.s.fill.fgColor.rgb,'F2F2F2');assert.equal(sheet['!freeze'],'A6');assert.equal(sheet['!autofilter'].ref,'A5:W9');
  assert.equal(sheet.F3.f,`COUNTIF(B6:B9,"${AUDIT_EXPORT_TICK}")`);assert.equal(book.Sheets.Index.D6.f,`'${rule.title}'!F3`);
  assert.ok(sheet['!xmlExtras'].dataValidations[0].includes('sqref="A6:B9"'));
  assert.equal(sheet.K6.v,'DEMO-PARENT');assert.equal(sheet.L6.v,'DEMO-PANEL');assert.equal(sheet.Q6.v,'DEMO-L1-1');
  const bytes=await workbookBytesCompact(book),reread=XLSX.read(bytes,{type:'array',cellStyles:true});
  assert.equal(reread.Sheets[rule.title].D8.s.fgColor.rgb,'F2F2F2');assert.equal(reread.Sheets[rule.title].A6.v,AUDIT_EXPORT_UNTICKED);
  const zip=XLSX.CFB.read(bytes,{type:'array'}),xml=new TextDecoder().decode(XLSX.CFB.find(zip,'/xl/worksheets/sheet3.xml').content);
  assert.match(xml,/dataValidation type="list"/);assert.match(xml,/sqref="A6:B9"/);
  assert.equal(JSON.stringify(result),before);
});
test('dismissed and completed findings are omitted and disabled rules have no tab',()=>{
  const result=fixture();result.findings.push({...result.findings[0],id:'off',rule:{id:'disabled',title:'Off'}});
  const book=buildAuditActionsWorkbook(result,'Demo',{excludedIds:['0'],completedEquipmentIds:['demo-3'],disabledRules:['disabled'],actionedIds:['1']});
  const sheet=book.Sheets[rule.title];assert.deepEqual([sheet.D6.v,sheet.D7.v].sort(),['DEMO-2','DEMO-4']);assert.equal(sheet.A6.v,AUDIT_EXPORT_UNTICKED);assert.equal(sheet.D3.v,2);
  assert.equal(book.SheetNames.length,4);
});
test('Actionable is first, with hidden queue calculations and links back to rule rows',async()=>{
  const book=buildAuditActionsWorkbook(fixture(),'Demo'),front=book.Sheets.Actionable,queue=book.Sheets['_Action queue'];
  assert.equal(front.B3.f,"INT('_Action queue'!A6/5)");
  assert.equal(queue.A3.f,`A2+IF(AND('${rule.title}'!A6="${AUDIT_EXPORT_TICK}",'${rule.title}'!B6<>"${AUDIT_EXPORT_TICK}"),5,0)+1`);
  assert.equal(front.K6.f,"IF(ROWS($A$6:A6)<=$B$3,MATCH(ROWS($A$6:A6)*5,'_Action queue'!$A$2:$A$6,1)+1,0)");
  assert.match(front.I6.f,/HYPERLINK/);assert.equal(queue.L3.v,`#'${rule.title}'!A6`);
  assert.equal(front['!cols'][10].hidden,true);assert.equal(book.Workbook.Sheets.at(-1).Hidden,1);
  const bytes=await workbookBytesCompact(book),reread=XLSX.read(bytes,{type:'array'});
  assert.equal(reread.Workbook.Sheets.at(-1).Hidden,1);assert.equal(reread.Sheets.Actionable.A6.f,front.A6.f);
  assert.equal(reread.Sheets[rule.title].B6.v,AUDIT_EXPORT_UNTICKED);
});
test('rule tab names are legal, unique and have correct index links',()=>{
  const result=fixture(),titles=['Index','A very long rule title / with several conflicts','A very long rule title : with several conflicts',"Parent's category"];
  result.findings=result.findings.map((finding,i)=>({...finding,rule:{id:String(i),title:titles[i]}}));
  const book=buildAuditActionsWorkbook(result,'Demo');assert.equal(new Set(book.SheetNames.map(n=>n.toLowerCase())).size,7);
  for(const name of book.SheetNames){assert.ok(name.length<=31);assert.doesNotMatch(name,/[\[\]:*?/\\]/);}
  for(let r=6;r<10;r++){const formula=book.Sheets.Index[`D${r}`].f;assert.ok(book.SheetNames.some(name=>formula===`'${name.replaceAll("'","''")}'!F3`));}
});
test('metadata joins use physical rows for duplicate tags and never guess an ambiguous row',()=>{
  const result=fixture();result.rows[1].equipmentId=result.rows[0].equipmentId;result.rows[1].building='Other building';result.findings[1].equipmentId=result.rows[0].equipmentId;
  result.findings.push({...result.findings[0],id:'unknown',sheet:'Unknown',row:99});
  const sheet=buildAuditActionsWorkbook(result,'Demo').Sheets[rule.title],grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''}).slice(5);
  assert.equal(grid.find(row=>row[21]==='Registry'&&row[22]===3)[15],'Other building');
  assert.equal(grid.find(row=>row[21]==='Unknown')[15],'');
});
test('empty scope exports a readable index, not phantom rules or broken progress formulas',()=>{
  const book=buildAuditActionsWorkbook({rows:[],findings:[]},'Demo');assert.deepEqual(book.SheetNames,['Actionable','Index']);assert.match(book.Sheets.Index.A6.v,/No active findings/);
});
test('Actions workbook What to do carries the tag-supported metadata correction',()=>{
  const parent={equipmentId:'F77-MAH101-01-0',upn:'101',discipline:'MECHANICAL DRY',systemName:'101 Makeup Air',_source:{sheet:'Registry',row:2}};
  const child={equipmentId:'F77-VFD101-01-0',upn:'650',discipline:'FACILITIES MONITORING SYSTEM',closestParent:parent.equipmentId,_source:{sheet:'Registry',row:3}};
  const result={...runSsmAudit({rows:[parent,child],missingHeaders:[]})};
  result.findings=result.findings.filter(f=>f.rule.id==='parent.cross-upn');assert.equal(result.findings.length,1);
  const book=buildAuditActionsWorkbook(result,'Synthetic');
  assert.match(book.Sheets[rule.title].I6.v,/Change F77-VFD101-01-0's UPN from 650 to 101/);
  assert.match(book.Sheets[rule.title].I6.v,/align its System Name/);
});
test('distinct Item Master replacements remain separate patterns and formula-like tags remain literal text',()=>{
  const result=fixture();result.findings=result.findings.map((f,i)=>({...f,rule:{id:'item-master.migration-advisory',title:'Item Master review'},why:'Review the replacement',actual:i<2?'DEMO_PANEL':'DEMO_PUMP',expected:i<2?'VF_PANEL':'VF_PUMP'}));
  result.findings[0].equipmentId='=1+1';const sheet=buildAuditActionsWorkbook(result,'Demo').Sheets['Item Master review'];
  const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''}).slice(5);assert.equal(new Set(grid.map(row=>row[2])).size,2);
  const at=grid.findIndex(row=>row[3]==='=1+1')+6;assert.equal(sheet[`D${at}`].t,'s');assert.equal(sheet[`D${at}`].f,undefined);
});

test('browser exports use direct bytes with native compression and the fallback, retaining every finding',async()=>{
  const context=vm.createContext({Uint8Array,ArrayBuffer,TextEncoder,TextDecoder});
  vm.runInContext(readFileSync(new URL('../src/vendor/sheetjs.js',import.meta.url),'utf8'),context);
  assert.equal(vm.runInContext('typeof Buffer',context),'undefined');
  const savedXlsx=globalThis.XLSX,savedCompression=globalThis.CompressionStream;
  const writer=context.XLSX.write,formats=[];
  context.XLSX.write=(book,options)=>{
    formats.push(options.type);
    assert.equal(options.type,'buffer','array output expands the ZIP into a per-byte character array');
    return writer(book,options);
  };
  try{
    globalThis.XLSX=context.XLSX;
    for(const native of [true,false]){
      globalThis.CompressionStream=native?savedCompression:undefined;
      const seed=fixture(),count=1800;
      const findings=Array.from({length:count},(_,i)=>({...seed.findings[i%4],id:String(i),equipmentId:`DEMO-${i}`,row:i+2}));
      const book=buildAuditActionsWorkbook({rows:[],findings},'Synthetic large export');
      const bytes=await workbookBytesCompact(book),read=context.XLSX.read(bytes,{type:'array',cellStyles:true});
      const sheet=read.Sheets[rule.title],last=count+5;
      assert.equal(sheet.D3.v,count);assert.equal(context.XLSX.utils.decode_range(sheet['!ref']).e.r,last-1);
      assert.equal(sheet[`B${last}`].v,AUDIT_EXPORT_UNTICKED);
      assert.equal(sheet[`B${last}`].s.fgColor.rgb,'F2F2F2');
      assert.ok(read.Sheets.Actionable[`A${last}`].f.includes('INDEX'));
      assert.equal(read.Workbook.Sheets.at(-1).Hidden,1);
      const zip=context.XLSX.CFB.read(bytes,{type:'array'});
      const xml=new TextDecoder().decode(context.XLSX.CFB.find(zip,'/xl/worksheets/sheet3.xml').content);
      assert.match(xml,/dataValidation type="list"/);
    }
    assert.deepEqual(formats,['buffer','buffer']);
  }finally{globalThis.XLSX=savedXlsx;globalThis.CompressionStream=savedCompression;}
});
