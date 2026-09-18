import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { clean, esc } from '../src/core/text.js'
import { auditNormId, auditSnapshotFromAoa } from '../src/audit/model.js'
import { auditReadEngineeringAoa, AUDIT_ENGINEERING_KINDS } from '../src/audit/engineering-references.js'
import { auditPrepareInWorker } from '../src/audit/review.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { auditApplyCorrections, auditMakeCorrection, auditCorrectionImpact } from '../src/audit/actions.js'
import { auditReadMigrationSettings } from '../src/audit/milestone-migration.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js',import.meta.url),'utf8'));
const ui=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
const state=readFileSync(new URL('../src/state.js',import.meta.url),'utf8').replaceAll('export ','');
const helpers=ui.slice(ui.indexOf('function engineeringUploadsHtml()'),ui.indexOf('\nfunction clearComparisonTarget()')).replace('export async function','async function');
const install=ui.slice(ui.indexOf('function reviewInstallDraft('),ui.indexOf('\nasync function reviewPrepare('));
const row={equipmentId:'TEST-PUMP111-01',building:'TEST',upn:'111',closestParent:'111 Chilled Water',systemName:'111 Chilled Water',discipline:'MECHANICAL WET'};
const mel=auditReadEngineeringAoa([['Equipment Tag','Building'],[row.equipmentId,'OTHER']],'mel','Synthetic MEL');
function fixture(){
  const aoa=[EXTO_REV21_COLUMNS.map(c=>c.header),EXTO_REV21_COLUMNS.map(c=>row[c.field]||'')],snapshot=auditSnapshotFromAoa(aoa,{sheet:'Registry'});
  const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
  return {snapshot,status:null,rawResult:runSsmAudit(snapshot),bytes:XLSX.write(book,{type:'array',bookType:'xlsx'})};
}
function harness(){
  const cache={},file=fixture(),calls={render:0,progress:0},errors=[];
  const ctx=vm.createContext({AUDIT_ENGINEERING_KINDS,auditNormId,esc,clean,XLSX,auditReadMigrationSettings,auditCorrectionImpact,modifyRecommendationContext:null,
    ic:()=>'<svg></svg>',renderUpload:()=>calls.render++,refreshSessionResult:()=>{},scheduleReviewAutosave:()=>{},
    runWithProgress:async(_a,_b,fn)=>fn(async()=>{},()=>calls.progress++),
    importAuditWorkbook:async()=>file,
    prepareAuditReview:async(session,changes,migration,migrationChanged,report,options={})=>auditPrepareInWorker(cache,{changes,previousChanges:session.changes||[],migration,migrationChanged,references:options.references||session.references,referencesChanged:options.references!==undefined,...(!cache.baseline?{baseline:session.baselineSnapshot,baselineResult:session.importResult,file:new Blob([session.sourceBytes])}:{})},report),
    clearComparisonTarget:()=>{},loadActioned:()=>{},loadExcluded:()=>{},loadChanges:()=>{},reviewAutosaveRestore:async()=>{},applyRulePreferences:r=>r,
    toast:message=>errors.push(message),console:{error:(_a,error)=>errors.push(error.message)},$ :()=>null,$$:()=>[],
  });
  vm.runInContext(state+'\n'+helpers+'\n'+install,ctx);
  const getSession=()=>vm.runInContext('S.session',ctx);
  return {ctx,getSession,file,calls,errors};
}
test('dedicated input UI keeps unselected sheets inactive and escapes workbook text',async()=>{
  const h=harness();
  await h.ctx.installEngineeringInput('mel',{name:'<private>.xlsx',references:[mel,{...mel,sheetName:'Other'}],selected:[]},()=>{});
  assert.equal(h.getSession().references.mel,undefined);
  const html=h.ctx.engineeringUploadsHtml();
  for(const kind of Object.keys(AUDIT_ENGINEERING_KINDS)){
    assert.ok(html.includes(`data-engineering-browse="${kind}"`));
    assert.ok(html.includes(`data-engineering-drop="${kind}"`));
  }
  assert.ok(html.includes('&lt;private&gt;.xlsx'));assert.ok(!html.includes('<private>'));assert.match(html,/Select the project sheets/);
});
test('references selected before the first registry participate in its background audit',async()=>{
  const h=harness();await h.ctx.installEngineeringInput('mel',{name:'synthetic.xlsx',references:[mel],selected:[0]},()=>{});
  await h.ctx.addAuditTarget({name:'registry.xlsx'},()=>{});
  assert.deepEqual(h.errors,[]);assert.ok(h.getSession().rawResult.findings.some(f=>f.rule.id==='reference.engineering-metadata'));
  assert.equal(h.getSession().engineeringInputs.mel.selected[0],0);assert.ok(h.calls.progress>0);
});
test('adding and removing engineering evidence reruns checks without dropping staged edits',async()=>{
  const h=harness();await h.ctx.addAuditTarget({name:'registry.xlsx'},()=>{});
  const session=h.getSession(),change=auditMakeCorrection(session.snapshot.rows[0],'Building','OTHER');
  session.changes=[change];session.snapshot=auditApplyCorrections(session.baselineSnapshot,session.changes);
  await h.ctx.installEngineeringInput('mel',{name:'synthetic.xlsx',references:[mel],selected:[0]},()=>{});
  assert.equal(session.changes.length,1);assert.equal(session.snapshot.rows[0].building,'OTHER');
  assert.ok(session.baselineResult.findings.some(f=>f.rule.id==='reference.engineering-metadata'));
  assert.ok(!session.rawResult.findings.some(f=>f.rule.id==='reference.engineering-metadata'));
  await h.ctx.installEngineeringInput('mel',null,()=>{});
  assert.equal(session.references.mel,undefined);assert.equal(session.changes.length,1);assert.equal(session.engineeringInputs.mel,undefined);
});
test('a replacement registry cannot silently inherit the previous project sources',async()=>{
  const h=harness();await h.ctx.installEngineeringInput('mel',{name:'synthetic.xlsx',references:[mel],selected:[0]},()=>{});
  await h.ctx.addAuditTarget({name:'first.xlsx'},()=>{});await h.ctx.addAuditTarget({name:'second.xlsx'},()=>{});
  assert.equal(Object.keys(h.getSession().references).length,0);assert.equal(Object.keys(h.getSession().engineeringInputs).length,0);
});

function uploadHarness(){
  const h=harness(),elements=new Map(),imports=[];
  for(const kind of Object.keys(AUDIT_ENGINEERING_KINDS)){
    const classes=new Set();
    elements.set(`[data-engineering-drop="${kind}"]`,{classList:{add:value=>classes.add(value),remove:value=>classes.delete(value)},classes});
    elements.set(`[data-engineering-file="${kind}"]`,{files:[],value:'',click(){this.clicked=true;}});
    elements.set(`[data-engineering-browse="${kind}"]`,{addEventListener(_event,handler){this.onclick=handler;}});
  }
  h.ctx.$=selector=>elements.get(selector);
  h.ctx.importAuditWorkbook=async(file,options)=>{imports.push({file,kind:options.referenceKind});return {references:[{...mel,kind:options.referenceKind}]};};
  h.ctx.wireEngineeringUploads(()=>{});
  return {...h,elements,imports,drop:kind=>elements.get(`[data-engineering-drop="${kind}"]`),input:kind=>elements.get(`[data-engineering-file="${kind}"]`)};
}
function dragEvent(files=[],types=['Files']){
  return {dataTransfer:{files,types},preventDefault(){this.prevented=true;},stopPropagation(){this.stopped=true;}};
}
for(const kind of Object.keys(AUDIT_ENGINEERING_KINDS))test(`${kind} drop imports into only its own document slot`,async()=>{
  const h=uploadHarness(),file={name:`synthetic-${kind}.xlsx`},event=dragEvent([file]);
  await h.drop(kind).ondrop(event);
  assert.equal(event.prevented,true);assert.equal(event.stopped,true);
  assert.equal(h.imports.length,1);assert.equal(h.imports[0].kind,kind);
  assert.deepEqual(Object.keys(h.getSession().references),[kind]);
  assert.equal(h.getSession().engineeringInputs[kind].name,file.name);
  assert.deepEqual(h.errors,[]);
});
test('document drop highlighting survives nested targets and clears on leave or drop',async()=>{
  const h=uploadHarness(),drop=h.drop('mel');
  drop.ondragenter(dragEvent([],['text/plain']));assert.equal(drop.classes.size,0);
  drop.ondragenter(dragEvent());drop.ondragenter(dragEvent());drop.ondragleave(dragEvent());
  assert.ok(drop.classes.has('dragging'));
  const over=dragEvent();drop.ondragover(over);assert.equal(over.dataTransfer.dropEffect,'copy');
  drop.ondragleave(dragEvent());assert.equal(drop.classes.size,0);
  drop.ondragenter(dragEvent());await drop.ondrop(dragEvent());assert.equal(drop.classes.size,0);
});
test('document drops reject unsupported files and batches without changing loaded evidence',async()=>{
  const h=uploadHarness();await h.drop('mel').ondrop(dragEvent([{name:'first.xlsx'}]));
  const prior=h.getSession().references.mel;
  await h.drop('mel').ondrop(dragEvent([{name:'wrong.pdf'}]));
  await h.drop('mel').ondrop(dragEvent([{name:'a.xlsx'},{name:'b.xlsx'}]));
  assert.equal(h.imports.length,1);assert.equal(h.getSession().references.mel,prior);
  assert.match(h.errors[0],/Excel workbook/);assert.match(h.errors[1],/one workbook/);
});
test('document picker retains its browse behavior and uses the same loader',async()=>{
  const h=uploadHarness(),input=h.input('pmd');
  h.elements.get('[data-engineering-browse="pmd"]').onclick();assert.equal(input.clicked,true);
  input.files=[{name:'picked.XLS'}];input.value='picked.XLS';await input.onchange();
  assert.equal(input.value,'');assert.equal(h.imports[0].kind,'pmd');
});
test('simultaneous document drops do not race and failed loads can be retried',async()=>{
  const h=uploadHarness();let reject;
  h.ctx.importAuditWorkbook=()=>new Promise((_resolve,fail)=>{reject=fail;});
  const first=h.drop('mel').ondrop(dragEvent([{name:'first.xlsx'}]));
  await Promise.resolve();await Promise.resolve();
  await h.drop('cable').ondrop(dragEvent([{name:'second.xlsx'}]));
  assert.match(h.errors[0],/Wait for/);
  reject(new Error('Synthetic parse failure'));await first;
  assert.match(h.errors[1],/Synthetic parse failure/);
  h.ctx.importAuditWorkbook=async()=>({references:[mel]});
  await h.drop('mel').ondrop(dragEvent([{name:'retry.xlsx'}]));
  assert.equal(h.getSession().engineeringInputs.mel.name,'retry.xlsx');
});
