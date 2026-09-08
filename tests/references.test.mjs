import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { MODULES } from '../build/manifest.mjs'
import { auditReadReferenceAoa, auditReadReferenceWorkbook, auditReferenceSheets, auditMilestoneReferenceMatch, auditReferenceRecommendation, auditReferenceFindings, SSM_AUDIT_REFERENCE_RULES } from '../src/audit/references.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js',import.meta.url),'utf8'),{filename:'sheetjs.js'});

const milestoneHeaders=['Milestone #','Title','L1 MS #','Parent Title'];
const milestoneRows=[milestoneHeaders,['DEMO-L2-M3-901','Phase Alpha readiness','DEMO-L1-M8','Phase Alpha parent']];
function reference(aoa=milestoneRows,kind='milestones'){return auditReadReferenceAoa(aoa,kind,kind==='milestones'?'Milestone Register':'Current Catalog');}
function catalog(names){return reference([['Item Master Name'],...names.map(name=>[name])],'itemMasters');}
function workbook(sheets){return {SheetNames:Object.keys(sheets),Sheets:sheets};}
function frozen(value){if(value&&typeof value==='object'){Object.values(value).forEach(frozen);Object.freeze(value);}return value;}

test('milestone header on row 2 preserves authoritative identifiers, titles, parent, and source row',()=>{
  const parsed=reference([['Synthetic register'],...milestoneRows]);
  assert.equal(parsed.kind,'milestones');
  assert.equal(parsed.sheetName,'Milestone Register');
  assert.equal(parsed.warning,undefined);
  assert.deepEqual(parsed.entries,[{id:'DEMO-L2-M3-901',title:'Phase Alpha readiness',parentId:'DEMO-L1-M8',parentTitle:'Phase Alpha parent',phase:'',milestone:'DEMO-L2-M3-901 Phase Alpha readiness',milestoneParent:'DEMO-L1-M8 Phase Alpha parent',sourceRow:3}]);
});

test('milestone header synonyms and optional parent titles are supported',()=>{
  for(const idHeader of ['Milestone #','L2 ID','Milestone ID','L2 Milestone']){
    for(const parentHeader of ['L1 MS #','L1 ID','Milestone Parent']){
      const parsed=reference([[idHeader,'Milestone Description',parentHeader],['L2-M3-901','Phase Alpha','L1-M8']]);
      assert.equal(parsed.entries[0].parentId,'L1-M8');
      assert.equal(parsed.entries[0].parentTitle,'');
      assert.equal(parsed.warning,undefined);
    }
  }
  assert.equal(reference([[' L2\nID ','Description',' L1 MS\n# '],['L2-9','Ready','L1-3']]).entries.length,1);
});

test('catalog recognizes all main-name headers and prefers Item Master Name over a separate ID',()=>{
  for(const header of ['Item Master Unique Identifier','Item Master','Item Master Name','Item Master ID']){
    assert.equal(reference([['Synthetic catalog'],[header],['VF_EL_PANEL']],'itemMasters').entries[0].name,'VF_EL_PANEL');
  }
  const parsed=reference([['Item Master ID','Item Master Name','Discipline'],['synthetic-uid','VF_EL_PANEL','Electrical']],'itemMasters');
  assert.deepEqual(parsed.entries,[{id:'synthetic-uid',name:'VF_EL_PANEL',discipline:'Electrical',sourceRow:2}]);
});

test('discovery finds relevant sheets only and reading always requires an explicit choice',()=>{
  const book=workbook({Instructions:[['How to use a register']],Register:milestoneRows,OtherRegister:milestoneRows,'Current Catalog':[['Item Master Name'],['VF_EL_PANEL']]});
  assert.deepEqual(auditReferenceSheets(book,'milestones'),['Register','OtherRegister']);
  assert.deepEqual(auditReferenceSheets(book,'itemMasters'),['Current Catalog']);
  for(const selection of [undefined,'','Missing','register']){
    const parsed=auditReadReferenceWorkbook(book,'milestones',selection);
    assert.deepEqual(parsed.entries,[]);
    assert.ok(parsed.warning);
  }
  assert.equal(auditReadReferenceWorkbook(book,'milestones','Register').entries.length,1);
  assert.deepEqual(auditReadReferenceWorkbook(book,'itemMasters','Register').entries,[]);
});

test('history and obsolete sheets are excluded without reading their contents, even if selected',()=>{
  const forbidden=['History','Change Log','ChangeHistory','Archived Catalog','Archive','Old Catalog','Obsolete','Deprecated','Superseded','Retired','Comments'];
  const book=workbook(Object.fromEntries([...forbidden.map(name=>[name,null]),['Current Catalog',[['Item Master Name'],['VF_EL_PANEL']]]]));
  for(const name of forbidden)Object.defineProperty(book.Sheets,name,{get(){throw new Error('Historical sheet must never be read');}});
  assert.deepEqual(auditReferenceSheets(book,'itemMasters'),['Current Catalog']);
  for(const name of forbidden){
    assert.deepEqual(auditReadReferenceWorkbook(book,'itemMasters',name).entries,[]);
    assert.deepEqual(auditReadReferenceAoa(milestoneRows,'milestones',name).entries,[]);
  }
});

test('only the selected catalog contributes names and the first or another sheet cannot supply a candidate',()=>{
  const book=workbook({'Old Catalog':[['Item Master Name'],['VF_EL_OLD']],'Main Catalog':[['Item Master Name'],['VF_EL_PANEL']],'Alternate Catalog':[['Item Master Name'],['VF_EL_OTHER']]});
  const parsed=auditReadReferenceWorkbook(book,'itemMasters','Main Catalog');
  for(const itemMaster of ['DEMO_EL_OLD','DEMO_EL_OTHER'])assert.equal(auditReferenceRecommendation({itemMaster},'itemMaster',{itemMasters:parsed}),null);
  assert.equal(auditReferenceRecommendation({itemMaster:'DEMO_EL_PANEL'},'itemMaster',{itemMasters:parsed}).value,'VF_EL_PANEL');
});

test('header search honors the first 200 physical rows, not just populated rows',()=>{
  for(const format of ['aoa','sparse','dense','legacyDense']){
    const convert=aoa=>format==='aoa'?aoa:format==='legacyDense'?Object.assign(aoa.map(row=>row.map(v=>({v,t:'s'}))),{'!ref':`A1:D${aoa.length}`}):format==='dense'?{'!ref':`A1:D${aoa.length}`,'!data':aoa.map(row=>row.map(v=>({v,t:'s'})))}:XLSX.utils.aoa_to_sheet(aoa);
    const inside=Array.from({length:199},()=>[]).concat(milestoneRows),outside=Array.from({length:200},()=>[]).concat(milestoneRows);
    const book=workbook({Inside:convert(inside),Outside:convert(outside)});
    assert.deepEqual(auditReferenceSheets(book,'milestones'),['Inside'],format);
    assert.equal(auditReadReferenceWorkbook(book,'milestones','Inside').entries[0].sourceRow,201,format);
    assert.deepEqual(auditReadReferenceWorkbook(book,'milestones','Outside').entries,[],format);
  }
});

test('dense and sparse workbooks support formatted IDs and offset columns without mutation',()=>{
  const aoa=[[],[null,'L2 ID','Description','L1 ID'],[null,7,'Synthetic gate','L1-9']];
  const sparse=XLSX.utils.aoa_to_sheet(aoa);sparse.B3.z='000';
  const dense={'!ref':'B2:D3','!data':[[],[null,{v:'L2 ID'},{v:'Description'},{v:'L1 ID'}],[null,{v:7,t:'n',z:'000'},{v:'Synthetic gate'},{v:'L1-9'}]]};
  for(const sheet of [sparse,dense]){
    const book=frozen(workbook({Register:sheet})),before=JSON.stringify(book);
    const parsed=auditReadReferenceWorkbook(book,'milestones','Register');
    assert.equal(parsed.entries[0].id,'007');
    assert.equal(parsed.entries[0].sourceRow,3);
    assert.equal(JSON.stringify(book),before);
  }
});

test('oversized formatting ranges do not expand sparse references into empty-cell scans',()=>{
  const sheet=XLSX.utils.aoa_to_sheet([['Synthetic register'],...milestoneRows]);
  sheet['!ref']='A1:XFD1048576';
  const book=workbook({Register:sheet}),started=performance.now();
  assert.deepEqual(auditReferenceSheets(book,'milestones'),['Register']);
  assert.equal(auditReadReferenceWorkbook(book,'milestones','Register').entries[0].sourceRow,3);
  assert.ok(performance.now()-started<1000);
});

test('current-table parsing skips repeated headers, comments, and obsolete status rows',()=>{
  const rows=[['Item Master Name','Status','Comments','Row Type'],['VF_EL_PANEL','Current','Historical comment about a valid record',''],['Comment: retired names below'],['VF_EL_OLD','Obsolete','',''],['VF_EL_REMOVED','Retired','',''],['VF_EL_NOTE','Current','','Comment'],['Item Master Name','Status','Comments','Row Type'],['VF_EL_ACTIVE','Active','',''],['History'],['VF_EL_ARCHIVED','Current','','']];
  assert.deepEqual(reference(rows,'itemMasters').entries.map(entry=>entry.name),['VF_EL_PANEL','VF_EL_ACTIVE']);
  const milestones=reference([milestoneHeaders,['Comments: not a milestone','Do not use'],['L1-9','Parent-only row'],...milestoneRows.slice(1),['History'],['L2-OLD','Old gate','L1-OLD']]);
  assert.equal(milestones.entries.length,1);
});

test('missing inputs and required values fail closed without guessing',()=>{
  for(const book of [null,{},workbook({Empty:[]})]){
    assert.deepEqual(auditReferenceSheets(book,'milestones'),[]);
    assert.ok(auditReadReferenceWorkbook(book,'milestones','Empty').warning);
  }
  assert.deepEqual(auditReferenceSheets(workbook({Register:milestoneRows}),'other'),[]);
  assert.ok(auditReadReferenceAoa(milestoneRows,'other','Register').warning);
  assert.ok(auditReadReferenceAoa(null,'milestones','Register').warning);
  const incomplete=reference([['L2 ID'],['L2-7'],[''],[null]]);
  assert.equal(incomplete.entries.length,1);
  assert.match(incomplete.warning,/parent identifier/);
  assert.equal(auditReferenceRecommendation({milestone:'L2-7'},'milestoneParent',{milestones:incomplete}),null);
  assert.deepEqual(reference([milestoneHeaders,['','Title without ID','L1-9']]).entries,[]);
  assert.deepEqual(reference([['Item Master Name'],[''],[null]],'itemMasters').entries,[]);
});

test('milestone matching preserves the full site code while tolerating title variance and whitespace',()=>{
  const parsed=reference();
  for(const milestone of ['DEMO-L2-M3-901','DEMO-L2-M3-901 Different display wording','DEMO-L2-M3-901 - Different display wording','DEMO-L2-M3-901: Different display wording','  demo - l2 - m3 - 901   Phase\nAlpha readiness ','Phase Alpha readiness','DEMO-L2 M3 901']){
    assert.equal(auditMilestoneReferenceMatch({milestone},parsed).status,'matched',milestone);
  }
  for(const milestone of ['L2-M3-901','OTHER-CAMPUS9-L2-M3-901 Phase Alpha readiness','DEMO-L2-M3-90','DEMO-L2-M3-999 Phase Alpha readiness','Please use DEMO-L2-M3-901','Alpha','901','DEMO-L1-M3-901']){
    assert.equal(auditMilestoneReferenceMatch({milestone},parsed).status,'missing',milestone);
  }
});

test('combined identifier/name cells support exact code and full-name matching',()=>{
  const parsed=reference([['Milestone','Milestone Parent'],['DEMO-L2-M3-901 Phase Alpha','DEMO-L1-M8 Phase Alpha']]);
  for(const milestone of ['DEMO-L2-M3-901','DEMO-L2-M3-901 Different wording','Phase Alpha'])assert.equal(auditMilestoneReferenceMatch({milestone},parsed).status,'matched');
  for(const milestone of ['L2-M3-901','OTHER-L2-M3-901 Phase Alpha'])assert.equal(auditMilestoneReferenceMatch({milestone},parsed).status,'missing');
  assert.equal(auditReferenceRecommendation({milestone:'DEMO-L2-M3-901'},'Milestone Parent',{milestones:parsed}).value,'DEMO-L1-M8 Phase Alpha');
});

test('canonical L2 names may be recommended only from an already exact unique identity',()=>{
  const milestones=reference();
  for(const milestone of ['DEMO-L2-M3-901','DEMO-L2-M3-901 Different wording','Phase Alpha readiness']){
    const result=auditReferenceRecommendation({milestone},'L2 Milestone',{milestones});
    assert.equal(result.value,'DEMO-L2-M3-901 Phase Alpha readiness');
    assert.equal(result.confidence,'verified');
  }
  assert.equal(auditReferenceRecommendation({milestone:'OTHER-L2-M3-901 Phase Alpha readiness'},'L2 Milestone',{milestones}),null);
  assert.equal(auditReferenceRecommendation({milestone:'L2-M3-901 Wrong phase',upn:'901'},'L2 Milestone',{milestones}),null);
});

test('comment appendices and error or uncached formula cells never supply catalog names',()=>{
  const sheet={'!data':[[{v:'Item Master Name'}],[{f:'UNSUPPORTED()'}],[{v:42,t:'e',w:'#VALUE!'}],[{v:'VF_EL_PANEL'}],[{v:'Comments'}],[{v:'VF_EL_OLD'}]],'!ref':'A1:A6'};
  assert.deepEqual(auditReadReferenceWorkbook(workbook({Main:sheet}),'itemMasters','Main').entries.map(entry=>entry.name),['VF_EL_PANEL']);
});

test('duplicate milestone codes with different phases remain ambiguous until the full name is exact',()=>{
  const parsed=reference([milestoneHeaders,...milestoneRows.slice(1),['DEMO-L2-M3-901','Phase Beta readiness','DEMO-L1-M9','Phase Beta parent']]);
  for(const milestone of ['DEMO-L2-M3-901','DEMO-L2-M3-901 Unrecognized phase']){
    assert.deepEqual(auditMilestoneReferenceMatch({milestone,milestoneParent:'DEMO-L1-M8',upn:'901'},parsed),{status:'ambiguous'});
    assert.equal(auditReferenceRecommendation({milestone},'milestoneParent',{milestones:parsed}),null);
  }
  const match=auditMilestoneReferenceMatch({milestone:'DEMO-L2-M3-901 - Phase Beta readiness'},parsed);
  assert.equal(match.status,'matched');
  assert.equal(match.entry.parentId,'DEMO-L1-M9');
  assert.equal(auditMilestoneReferenceMatch({milestone:'Phase Beta readiness'},parsed).status,'matched');
  assert.equal(auditMilestoneReferenceMatch({milestone:'OTHER-L2-M3-901 Phase Beta readiness'},parsed).status,'missing');
  const sameName=reference([milestoneHeaders,...milestoneRows.slice(1),['DEMO-L2-M3-901','Phase Alpha readiness','DEMO-L1-M99','Other parent']]);
  assert.equal(auditMilestoneReferenceMatch({milestone:'Phase Alpha readiness'},sameName).status,'ambiguous');
  assert.equal(auditMilestoneReferenceMatch({milestone:'DEMO-L2-M3-901 Phase Alpha readiness'},sameName).status,'ambiguous');
  const phaseColumn=reference([['L2 ID','Title','L1 ID','Phase'],['L2-9','Ready','L1-8','Alpha'],['L2-9','Ready','L1-8','Beta']]);
  assert.equal(auditMilestoneReferenceMatch({milestone:'L2-9'},phaseColumn).status,'ambiguous');
  assert.equal(auditMilestoneReferenceMatch({milestone:'L2-9 Ready'},phaseColumn).status,'ambiguous');
});

test('site namespaces and identifier suffix punctuation are never discarded or equated',()=>{
  const parsed=reference([milestoneHeaders,
    ['DEMO-WEST.9-L2-M3-901_1','Shared name','DEMO-WEST.9-L1-M8','Shared parent'],
    ['OTHER-L2-M3-901_1','Shared name','OTHER-L1-M8','Shared parent'],
  ]);
  for(const id of ['DEMO-WEST.9-L2-M3-901_1','OTHER-L2-M3-901_1']){
    const match=auditMilestoneReferenceMatch({milestone:`${id} Different wording`},parsed);
    assert.equal(match.status,'matched');
    assert.equal(match.entry.id,id);
  }
  for(const milestone of ['L2-M3-901_1 Shared name','DEMO-L2-M3-901_1 Shared name','DEMO_WEST.9-L2-M3-901_1 Shared name','DEMO-WEST.9-L2-M3-901-1 Shared name'])assert.equal(auditMilestoneReferenceMatch({milestone},parsed).status,'missing',milestone);
  assert.equal(auditMilestoneReferenceMatch({milestone:'Shared name'},parsed).status,'ambiguous');
});

test('an explicit unmatched or foreign code cannot fall back to a coincident exact title',()=>{
  const parsed=reference([milestoneHeaders,
    ...milestoneRows.slice(1),
    ['DEMO-L2-M3-902','OTHER-L2-M3-901 Phase Alpha readiness','DEMO-L1-M8','Phase Alpha parent'],
    ['DEMO-L2-M3-903','Please use OTHER-L2-M3-901','DEMO-L1-M8','Phase Alpha parent'],
  ]);
  for(const milestone of ['OTHER-L2-M3-901 Phase Alpha readiness','Please use OTHER-L2-M3-901','DEMO-L2-M3-999 Phase Alpha readiness']){
    assert.equal(auditMilestoneReferenceMatch({milestone},parsed).status,'missing',milestone);
    assert.equal(auditReferenceRecommendation({milestone},'L1 Milestone Parent',{milestones:parsed}),null);
  }
});

test('identical repeated records deduplicate but distinct IDs sharing a name are ambiguous',()=>{
  const parsed=reference([...milestoneRows,...milestoneRows.slice(1)]);
  assert.equal(parsed.entries.length,1);
  assert.equal(auditMilestoneReferenceMatch({milestone:'DEMO-L2-M3-901'},parsed).status,'matched');
  const sharedTitle=reference([milestoneHeaders,...milestoneRows.slice(1),['DEMO-L2-M4-902','Phase Alpha readiness','DEMO-L1-M8','Phase Alpha parent']]);
  assert.equal(auditMilestoneReferenceMatch({milestone:'Phase Alpha readiness'},sharedTitle).status,'ambiguous');
});

test('missing references or a missing L2 stay unverified and UPN alone never supplies a milestone',()=>{
  for(const parsed of [null,reference([]),catalog(['VF_EL_PANEL'])])assert.deepEqual(auditMilestoneReferenceMatch({milestone:'L2-M3-901'},parsed),{status:'unverified'});
  for(const row of [null,{}, {upn:'901',systemName:'Phase Alpha readiness',equipmentDescription:'Phase Alpha readiness',milestoneParent:'DEMO-L1-M8'}]){
    assert.deepEqual(auditMilestoneReferenceMatch(row,reference()),{status:'unverified'});
    assert.equal(auditReferenceRecommendation(row,'milestone',{milestones:reference()}),null);
    assert.equal(auditReferenceRecommendation(row,'milestoneParent',{milestones:reference()}),null);
  }
});

test('parent recommendations use the explicit register link rather than deriving an L1 from L2 or UPN',()=>{
  for(const field of ['milestoneParent','Milestone Parent','L1 Milestone Parent']){
    const result=auditReferenceRecommendation({milestone:'DEMO-L2-M3-901',milestoneParent:'L1-M3',upn:'999'},field,{milestones:reference()});
    assert.equal(result.value,'DEMO-L1-M8 Phase Alpha parent');
    assert.equal(result.confidence,'verified');
    assert.match(result.reason,/explicitly maps/);
  }
  for(const milestoneParent of ['DEMO-L1-M8','DEMO-L1-M8 Different display wording','Phase Alpha parent'])assert.equal(auditReferenceRecommendation({milestone:'DEMO-L2-M3-901',milestoneParent},'milestoneParent',{milestones:reference()}),null);
  for(const milestoneParent of ['L1-M8','ANOTHER-L1-M8 Phase Alpha parent'])assert.equal(auditReferenceRecommendation({milestone:'DEMO-L2-M3-901',milestoneParent},'milestoneParent',{milestones:reference()}).value,'DEMO-L1-M8 Phase Alpha parent');
  const missingParent=reference([milestoneHeaders,['L2-7','Ready','','Parent title alone']]);
  assert.equal(auditReferenceRecommendation({milestone:'L2-7'},'milestoneParent',{milestones:missingParent}),null);
  const wrongLevel=reference([milestoneHeaders,['L2-7','Ready','L2-8','Not L1']]);
  assert.equal(auditReferenceRecommendation({milestone:'L2-7'},'milestoneParent',{milestones:wrongLevel}),null);
});

test('an arbitrary site-prefix catalog equivalent is review-only even with descriptive equipment metadata',()=>{
  const itemMasters=catalog(['VF_I&C_RIO W/O SUD']);
  for(const prefix of ['PROJECT99','UNLISTED-CAMPUS.4','XYZ123']){
    const result=auditReferenceRecommendation({itemMaster:`${prefix}_I&C_RIO W/O SUD`,equipmentDescription:'Remote I/O Panel',equipmentClassification:'RIO',discipline:'I&C'},'Item Master Unique Identifier',{itemMasters});
    assert.equal(result.value,'VF_I&C_RIO W/O SUD');
    assert.equal(result.confidence,'review');
    assert.match(result.reason,/not a literal catalog member/);
    assert.match(result.reason,/legacy prefix invalid/);
    assert.match(result.reason,/compatibility is unverified/);
  }
});

test('catalog equivalence tolerates whitespace but preserves all functional detail and punctuation',()=>{
  const itemMasters=catalog(['VF_PROC_EQ_ELEC_I&C','VF_I&C_RIO W/O SUD','VF_EL_AC/DC']);
  assert.equal(auditReferenceRecommendation({itemMaster:'  campus9 _ PROC _ EQ _ ELEC _ I&C '},'itemMaster',{itemMasters}).value,'VF_PROC_EQ_ELEC_I&C');
  for(const itemMaster of ['CAMPUS9_PROC_EQ','CAMPUS9_PROC_EQ_ELEC','CAMPUS9_PROC_EQ_ELEC_I&C_EXTRA','CAMPUS9_I&C_RIO W/SUD','CAMPUS9_EL_ACDC','CAMPUS9_EL_AC-DC','Use VF_PROC_EQ_ELEC_I&C for this row'])assert.equal(auditReferenceRecommendation({itemMaster},'itemMaster',{itemMasters}),null,itemMaster);
  const shorter=catalog(['VF_EQ_ELEC','VF_PANEL']);
  for(const itemMaster of ['VF_PROC_EQ_ELEC','CAMPUS9_EL_PANEL'])assert.equal(auditReferenceRecommendation({itemMaster},'itemMaster',{itemMasters:shorter}),null);
});

test('multi-token site prefixes require explicit site evidence, not an arbitrary suffix match',()=>{
  const itemMasters=catalog(['VF_EL_PANEL']);
  assert.equal(auditReferenceRecommendation({itemMaster:'CAMPUS_WEST_EL_PANEL',site:'CAMPUS_WEST'},'itemMaster',{itemMasters}).value,'VF_EL_PANEL');
  assert.equal(auditReferenceRecommendation({itemMaster:'CAMPUS_WEST_EL_PANEL'},'itemMaster',{itemMasters}),null);
});

test('ambiguous catalog equivalents, absent fields, and nonmembers without equivalents produce no suggestion',()=>{
  const itemMasters=catalog(['VF1_EL_PANEL','VF2_EL_PANEL']);
  assert.equal(auditReferenceRecommendation({itemMaster:'CAMPUS9_EL_PANEL'},'itemMaster',{itemMasters}),null);
  const duplicateIds=reference([['Item Master ID','Item Master Name'],['ID-1','VF_EL_PANEL'],['ID-2','VF_EL_PANEL']],'itemMasters');
  assert.equal(auditReferenceRecommendation({itemMaster:'CAMPUS9_EL_PANEL'},'itemMaster',{itemMasters:duplicateIds}),null);
  for(const row of [null,{}, {itemMaster:''},{itemMaster:'PROJECT9_UNKNOWN'},{equipmentDescription:'Electrical Panel',upn:'901'}])assert.equal(auditReferenceRecommendation(row,'itemMaster',{itemMasters}),null);
  assert.equal(auditReferenceRecommendation({itemMaster:'CAMPUS9_EL_PANEL'},'itemMaster',null),null);
  assert.equal(auditReferenceRecommendation({milestone:'L2-M3-901'},'upn',{milestones:reference()}),null);
});

test('literal catalog membership needs no rename and never implies equipment compatibility',()=>{
  const itemMasters=catalog(['VF_EL_PANEL','LEGACY9_EL_PANEL']);
  for(const itemMaster of ['VF_EL_PANEL','LEGACY9_EL_PANEL',' vf_el_panel '])assert.equal(auditReferenceRecommendation({itemMaster,equipmentDescription:'Unrelated equipment'},'itemMaster',{itemMasters}),null);
  const separateId=reference([['Item Master ID','Item Master Name'],['UID-9','VF_EL_PANEL']],'itemMasters');
  assert.equal(auditReferenceRecommendation({itemMaster:'UID-9'},'itemMaster',{itemMasters:separateId}),null);
});

test('reference parsing, matching, and recommendations are deterministic and do not mutate inputs',()=>{
  const aoa=frozen(milestoneRows.map(row=>[...row])),parsed=frozen(reference(aoa)),row=frozen({milestone:'DEMO-L2-M3-901 Changed wording',milestoneParent:'Wrong L1'}),references=frozen({milestones:parsed,itemMasters:null});
  const before=JSON.stringify({aoa,row,references});
  assert.deepEqual(reference(aoa),reference(aoa));
  assert.deepEqual(auditMilestoneReferenceMatch(row,parsed),auditMilestoneReferenceMatch(row,parsed));
  assert.deepEqual(auditReferenceRecommendation(row,'milestoneParent',references),auditReferenceRecommendation(row,'milestoneParent',references));
  assert.equal(JSON.stringify({aoa,row,references}),before);
});

test('placeholder identities and parents are unavailable, never verified recommendations',()=>{
  const parsed=reference([milestoneHeaders,['N/A','Not a milestone','L1-3'],['L2-7','Ready','TBD'],['L2-8','Ready','-']]);
  assert.equal(parsed.entries.length,2);
  for(const milestone of ['N/A','TBD','-'])assert.equal(auditMilestoneReferenceMatch({milestone},parsed).status,'unverified');
  for(const milestone of ['L2-7','L2-8'])assert.equal(auditReferenceRecommendation({milestone},'milestoneParent',{milestones:parsed}),null);
});

test('optional findings are engine-shaped and distinguish absence, ambiguity, and positive contradictions',()=>{
  const parsed=reference([milestoneHeaders,...milestoneRows.slice(1),['L2-8','Other phase','L1-2','Other parent'],['L2-8','Later phase','L1-3','Later parent']]);
  const rows=[
    {equipmentId:'UNKNOWN',milestone:'L2-999'},
    {equipmentId:'AMBIGUOUS',milestone:'L2-8'},
    {equipmentId:'CONTRADICTION',milestone:'DEMO-L2-M3-901 Different wording',milestoneParent:'L1-M3'},
    {equipmentId:'MISSING-PARENT',milestone:'DEMO-L2-M3-901'},
    {equipmentId:'LOCAL-ALIAS',milestone:'DEMO-L2-M3-901',milestoneParent:'A local alias not in the register'},
    {equipmentId:'VALID',milestone:'DEMO-L2-M3-901 Different wording',milestoneParent:'DEMO-L1-M8 Different parent wording'},
    {equipmentId:'NO-L2',upn:'901'},
  ].map((row,index)=>({...row,_source:{sheet:'Synthetic registry',row:index+3}}));
  const snapshot=frozen({rows}),references=frozen({milestones:parsed,itemMasters:null}),before=JSON.stringify({snapshot,references});
  const findings=auditReferenceFindings(snapshot,references);
  assert.deepEqual(findings.map(finding=>[finding.equipmentId,finding.severity]),[['UNKNOWN','info'],['AMBIGUOUS','info'],['CONTRADICTION','error'],['MISSING-PARENT','info'],['LOCAL-ALIAS','info']]);
  assert.equal(findings[0].rule,SSM_AUDIT_REFERENCE_RULES.milestoneUnknown);
  assert.equal(findings[1].rule,SSM_AUDIT_REFERENCE_RULES.milestoneAmbiguous);
  assert.equal(findings[2].expected,'DEMO-L1-M8 Phase Alpha parent');
  assert.equal(findings[2].field,'L1 Milestone Parent');
  for(const finding of findings){
    assert.equal(finding.schemaVersion,1);
    assert.equal(finding.id,`${finding.rule.id}:${finding.fingerprint}`);
    assert.match(finding.fingerprint,/^[a-f0-9]{8}$/);
    assert.equal(finding.sheet,'Synthetic registry');
    assert.ok(finding.row>=3);
    assert.equal(finding.category,'milestones');
    assert.equal(finding.relationship,null);
    assert.equal(finding.relatedEquipmentId,'');
    assert.ok(finding.searchKey.includes(finding.equipmentId));
    assert.ok(finding.why&&finding.expected&&finding.recommendation);
    assert.ok(Object.isFrozen(finding));
  }
  assert.equal(new Set(findings.map(finding=>finding.id)).size,findings.length);
  assert.deepEqual(auditReferenceFindings(snapshot,references),findings);
  assert.equal(JSON.stringify({snapshot,references}),before);
});

test('reference findings remain optional and produce no catalog or legacy-prefix flags',()=>{
  const snapshot={rows:[{equipmentId:'DEMO-1',milestone:'L2-M3-901',itemMaster:'LEGACY_EL_PANEL'}]};
  for(const references of [null,{}, {milestones:null,itemMasters:catalog(['VF_EL_PANEL'])},{milestones:reference([])}])assert.deepEqual(auditReferenceFindings(snapshot,references),[]);
  for(const empty of [null,{}, {rows:null},{rows:[null,{}]}])assert.deepEqual(auditReferenceFindings(empty,{milestones:reference()}),[]);
});

test('parent contradictions compare full codes, not differing display titles',()=>{
  const milestones=reference([milestoneHeaders,...milestoneRows.slice(1),['L2-9','Other','GATE-OTHER','Known other parent']]);
  const rows=[
    {equipmentId:'SAME-CODE',milestone:'DEMO-L2-M3-901',milestoneParent:'DEMO-L1-M8 Phase Beta parent'},
    {equipmentId:'FOREIGN-CODE',milestone:'DEMO-L2-M3-901',milestoneParent:'OTHER-L1-M8 Phase Alpha parent'},
    {equipmentId:'NO-NAMESPACE',milestone:'DEMO-L2-M3-901',milestoneParent:'L1-M8 Phase Alpha parent'},
    {equipmentId:'KNOWN-OTHER',milestone:'DEMO-L2-M3-901',milestoneParent:'Known other parent'},
  ];
  const findings=auditReferenceFindings({rows},{milestones});
  assert.deepEqual(findings.map(finding=>[finding.equipmentId,finding.severity]),[['FOREIGN-CODE','error'],['NO-NAMESPACE','error'],['KNOWN-OTHER','error']]);
  assert.ok(Object.values(SSM_AUDIT_REFERENCE_RULES).every(rule=>rule.enabled&&rule.title===rule.standardRef&&!/\.xlsx|\.xls|\.csv/i.test(rule.title)));
});

test('parent title-only validation requires a unique parent identity and preserves unknown aliases as review',()=>{
  const milestones=reference([milestoneHeaders,
    ...milestoneRows.slice(1),
    ['OTHER-L2-M3-901','Other milestone','OTHER-L1-M8','Phase Alpha parent'],
    ['DEMO-L2-M3-902','Another milestone','DEMO-L1-M8','Alternate parent wording'],
  ]);
  const rows=[
    {equipmentId:'AMBIGUOUS-TITLE',milestone:'DEMO-L2-M3-901',milestoneParent:'Phase Alpha parent'},
    {equipmentId:'SAME-PARENT-ALIAS',milestone:'DEMO-L2-M3-901',milestoneParent:'Alternate parent wording'},
    {equipmentId:'UNKNOWN-TITLE',milestone:'DEMO-L2-M3-901',milestoneParent:'Unlisted parent alias'},
    {equipmentId:'FULL-PARENT',milestone:'DEMO-L2-M3-901',milestoneParent:'DEMO-L1-M8 Changed wording'},
  ];
  assert.deepEqual(auditReferenceFindings({rows},{milestones}).map(finding=>[finding.equipmentId,finding.severity]),[['AMBIGUOUS-TITLE','info'],['UNKNOWN-TITLE','info']]);
  assert.equal(auditReferenceRecommendation(rows[1],'L1 Milestone Parent',{milestones}),null);
  assert.equal(auditReferenceRecommendation(rows[0],'L1 Milestone Parent',{milestones}).value,'DEMO-L1-M8 Phase Alpha parent');
});

test('reference findings build a local index for a large registry without retaining global state',()=>{
  const count=3000,milestones=reference([milestoneHeaders,...Array.from({length:count},(_,index)=>[`L2-${index}`,`Gate ${index}`,`L1-${index}`,''])]);
  const snapshot={rows:Array.from({length:count},(_,index)=>({equipmentId:`DEMO-${index}`,milestone:`L2-${index}`,milestoneParent:`L1-${index}`}))};
  const started=performance.now();
  assert.deepEqual(auditReferenceFindings(snapshot,{milestones}),[]);
  assert.ok(performance.now()-started<3000);
});

test('new reference helpers compile with the existing single-scope bundle without writing build artifacts',()=>{
  const paths=[...new Set([...MODULES,'src/audit/references.js'])];
  const source=paths.map(path=>readFileSync(new URL(`../${path}`,import.meta.url),'utf8').replace(/^\s*import\s+[^\n]+\n/gm,'').replace(/^\s*export\s+\{[^}]*\};?\s*$/gm,'').replace(/\bexport\s+(?=(?:async\s+)?function\b|class\b|const\b|let\b|var\b)/g,'')).join('\n');
  assert.doesNotThrow(()=>new vm.Script(source));
});
test('current catalog IM Name and milestone driver headers are supported',()=>{
  const catalog=auditReadReferenceAoa([['Discipline','IM Name'],['Electrical','VF_EL_PANEL']],'itemMasters','Current Catalog');
  assert.equal(catalog.entries[0].name,'VF_EL_PANEL');
  const schedule=auditReadReferenceAoa([['Milestone #','L1 MS #','L1 MS / Driver'],['L2-DEMO-1','L1-DEMO-1','Facility readiness']],'milestones','Schedule');
  assert.equal(schedule.entries[0].parentTitle,'Facility readiness');
  assert.equal(auditReadReferenceAoa([['IM Name'],['VF_OLD']],'itemMasters','Legacy').entries.length,0);
});
test('schedule identity uses MS instead of the separate milestone sequence number',()=>{
  const register=auditReadReferenceAoa([['Milestone #','MS','MS Description','L1 MS #','L1 MS / Driver'],['17','L2-DEMO-101','Enable the system','L1-DEMO-1','Facility readiness']],'milestones','Register');
  assert.equal(register.entries[0].id,'L2-DEMO-101');assert.equal(register.entries[0].title,'Enable the system');
  assert.equal(auditMilestoneReferenceMatch({milestone:'L2-DEMO-101'},register).status,'matched');
});
