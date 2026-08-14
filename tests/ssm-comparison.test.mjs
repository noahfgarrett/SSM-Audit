import test from 'node:test'
import assert from 'node:assert/strict'

import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { compareSsmRegistries } from '../src/audit/compare.js'

const headers=EXTO_REV21_COLUMNS.map(column=>column.header)
const index=Object.fromEntries(EXTO_REV21_COLUMNS.map(column=>[column.field,column.index]))
function row(values){const cells=new Array(headers.length).fill('');for(const [field,value] of Object.entries(values))cells[index[field]]=value;return cells;}
function snapshot(values,name){return auditSnapshotFromAoa([headers,...values],{file:`${name}.xlsx`,sheet:'Registry'});}
function equipment({building,tag,parent,description,upn='101',discipline='MECHANICAL DRY',classification='',dependencies=''}){
  return row({building,equipmentId:tag,closestParent:parent,closestParentStatus:'NEW',equipmentDescription:description,equipmentClassification:classification,dependencies,upn,discipline,systemName:`${upn} Example System`});
}

test('project comparison aligns equivalent hierarchies while excluding Building',()=>{
  const target=snapshot([
    equipment({building:'BLD-A',tag:'BLD-A-SYSTEM-GROUP',parent:'101 Example System',description:'HEADER'}),
    equipment({building:'BLD-A',tag:'BLD-A-PANEL-01',parent:'BLD-A-SYSTEM-GROUP',description:'Electrical Panel',classification:'PNL',discipline:'ELECTRICAL'}),
    equipment({building:'BLD-A',tag:'BLD-A-RIO-01',parent:'BLD-A-SYSTEM-GROUP',description:'Remote I/O Panel',classification:'RIO',discipline:'FACILITIES MONITORING SYSTEM',dependencies:'BLD-A-PANEL-01'}),
  ],'target');
  const reference=snapshot([
    equipment({building:'BLD-Z (AREA 2)',tag:'BLD-Z-SYSTEM-GROUP',parent:'101 Example System',description:'HEADER'}),
    equipment({building:'BLD-Z (AREA 2)',tag:'BLD-Z-PANEL-01',parent:'BLD-Z-SYSTEM-GROUP',description:'Electrical Panel',classification:'PNL',discipline:'ELECTRICAL'}),
    equipment({building:'BLD-Z (AREA 2)',tag:'BLD-Z-RIO-01',parent:'BLD-Z-SYSTEM-GROUP',description:'Remote I/O Panel',classification:'RIO',discipline:'FACILITIES MONITORING SYSTEM',dependencies:'BLD-Z-PANEL-01'}),
  ],'reference');
  const result=compareSsmRegistries(target,reference),system=result.systems[0];
  assert.equal(system.upn,'101');
  assert.equal(system.status,'aligned');
  assert.equal(system.observations.length,0);
  assert.deepEqual(system.pairSummary,{aligned:3,changed:0,targetOnly:0,referenceOnly:0});
});

test('project comparison explains I&C parent and header placement differences',()=>{
  const target=snapshot([
    equipment({building:'A1',tag:'A1-CONTROLS',parent:'650 Example System',description:'HEADER',upn:'650',discipline:'FACILITIES MONITORING SYSTEM'}),
    equipment({building:'A1',tag:'A1-RIO-01',parent:'A1-CONTROLS',description:'Remote I/O Panel',classification:'RIO',upn:'650',discipline:'FACILITIES MONITORING SYSTEM'}),
  ],'target');
  const reference=snapshot([
    equipment({building:'B9',tag:'B9-PANEL-01',parent:'650 Example System',description:'Electrical Panel',classification:'PNL',upn:'650',discipline:'ELECTRICAL'}),
    equipment({building:'B9',tag:'B9-RIO-01',parent:'B9-PANEL-01',description:'Remote I/O Panel',classification:'RIO',upn:'650',discipline:'FACILITIES MONITORING SYSTEM'}),
  ],'reference');
  const system=compareSsmRegistries(target,reference).systems[0],rio=system.pairs.find(pair=>pair.target&&pair.target.role==='Remote I/O Panel');
  assert.equal(system.status,'different');
  assert.ok(system.observations.some(item=>item.type==='controls'&&/I&C nesting/.test(item.title)));
  assert.ok(system.observations.some(item=>item.type==='headers'&&item.subject==='REMOTE I/O PANEL'));
  assert.equal(rio.status,'changed');
  assert.ok(rio.differences.includes('Parent equipment type differs'));
  assert.ok(rio.differences.includes('Organizational header usage differs'));
});

test('project comparison recognizes instruments nested within non-controls disciplines',()=>{
  const target=snapshot([
    equipment({building:'A1',tag:'A1-MAH-01',parent:'101 Example System',description:'Makeup Air Handler'}),
    equipment({building:'A1',tag:'A1-TET-01',parent:'A1-MAH-01',description:'Temperature Transmitter',classification:'TET'}),
  ],'target');
  const reference=snapshot([
    equipment({building:'B1',tag:'B1-MAH-01',parent:'101 Example System',description:'Makeup Air Handler'}),
    equipment({building:'B1',tag:'B1-PNL-01',parent:'101 Example System',description:'Electrical Panel',discipline:'ELECTRICAL'}),
    equipment({building:'B1',tag:'B1-TET-01',parent:'B1-PNL-01',description:'Temperature Transmitter',classification:'TET'}),
  ],'reference');
  const system=compareSsmRegistries(target,reference).systems[0];
  assert.equal(system.targetControls,1);
  assert.equal(system.referenceControls,1);
  assert.ok(system.observations.some(item=>item.type==='controls'&&item.subject.includes('TEMPERATURE TRANSMITTER')));
});

test('project comparison aligns repeated electrical branches by neutral identity and topology',()=>{
  const target=snapshot([
    equipment({building:'A1',tag:'A1-XFM-A-01',parent:'603 Example System',description:'Electrical Transformer',classification:'TFRM',upn:'603',discipline:'ELECTRICAL'}),
    equipment({building:'A1',tag:'A1-XFM-B-02',parent:'603 Example System',description:'Electrical Transformer',classification:'TFRM',upn:'603',discipline:'ELECTRICAL'}),
    equipment({building:'A1',tag:'A1-LVS-A-01',parent:'A1-XFM-A-01',description:'Low Voltage Switchgear',classification:'SWG',upn:'603',discipline:'ELECTRICAL'}),
    equipment({building:'A1',tag:'A1-LVS-B-02',parent:'A1-XFM-B-02',description:'Low Voltage Switchgear',classification:'SWG',upn:'603',discipline:'ELECTRICAL'}),
  ],'target');
  const reference=snapshot([
    equipment({building:'Z9',tag:'Z9-TRANSFORMER-A-02',parent:'603 Example System',description:'Transformer',classification:'TFRM',upn:'603',discipline:'ELECTRICAL'}),
    equipment({building:'Z9',tag:'Z9-TRANSFORMER-Z-01',parent:'603 Example System',description:'Transformer',classification:'TFRM',upn:'603',discipline:'ELECTRICAL'}),
    equipment({building:'Z9',tag:'Z9-SWITCHGEAR-A-02',parent:'Z9-TRANSFORMER-A-02',description:'Unit Substation Switchgear',classification:'SWG',upn:'603',discipline:'ELECTRICAL'}),
    equipment({building:'Z9',tag:'Z9-SWITCHGEAR-Z-01',parent:'Z9-TRANSFORMER-Z-01',description:'Unit Substation Switchgear',classification:'SWG',upn:'603',discipline:'ELECTRICAL'}),
  ],'reference');
  const system=compareSsmRegistries(target,reference).systems[0],first=system.pairs.find(pair=>pair.target&&pair.target.tag==='A1-XFM-A-01'),second=system.pairs.find(pair=>pair.target&&pair.target.tag==='A1-XFM-B-02');
  assert.equal(first.reference.tag,'Z9-TRANSFORMER-Z-01');
  assert.equal(second.reference.tag,'Z9-TRANSFORMER-A-02');
  assert.equal(first.matchReason,'Same equipment type, parent type, and neutral tag identity');
  assert.equal(system.pairSummary.targetOnly,0);
  assert.equal(system.pairSummary.referenceOnly,0);
});

test('project comparison normalizes common FMS network nomenclature without tag equality',()=>{
  const target=snapshot([
    equipment({building:'A1',tag:'A1-NET-01',parent:'650 Example System',description:'Stratix Enclosure',classification:'STRATIX ENCLOSURE',upn:'650',discipline:'FACILITIES MONITORING SYSTEM'}),
    equipment({building:'A1',tag:'A1-SW-01',parent:'A1-NET-01',description:'Cisco Switch',classification:'CISCO SWITCH',upn:'650',discipline:'FACILITIES MONITORING SYSTEM'}),
    equipment({building:'A1',tag:'A1-OIT-01',parent:'A1-NET-01',description:'Operator Interface Terminal',classification:'OIT',upn:'650',discipline:'FACILITIES MONITORING SYSTEM'}),
  ],'target');
  const reference=snapshot([
    equipment({building:'B2',tag:'B2-BAY-01',parent:'650 Example System',description:'FMS Bay Stratix Panel',classification:'BSTX',upn:'650',discipline:'I&C'}),
    equipment({building:'B2',tag:'B2-STX-01',parent:'B2-BAY-01',description:'Startix Switch',classification:'STX',upn:'650',discipline:'I&C'}),
    equipment({building:'B2',tag:'B2-HMI-01',parent:'B2-BAY-01',description:'HMI Touchscreen',classification:'HMI',upn:'650',discipline:'I&C'}),
  ],'reference');
  const system=compareSsmRegistries(target,reference).systems[0];
  assert.equal(system.pairSummary.targetOnly,0);
  assert.equal(system.pairSummary.referenceOnly,0);
  assert.ok(system.pairs.some(pair=>pair.target?.semanticRole==='Industrial network switch'&&pair.reference?.semanticRole==='Industrial network switch'));
  assert.ok(system.pairs.some(pair=>pair.target?.semanticRole==='Operator interface'&&pair.reference?.semanticRole==='Operator interface'));
});

test('project comparison builds one synchronized tree from paired hierarchy branches',()=>{
  const target=snapshot([
    equipment({building:'A1',tag:'A1-MAH-01',parent:'101 Example System',description:'Makeup Air Handler'}),
    equipment({building:'A1',tag:'A1-TET-01',parent:'A1-MAH-01',description:'Temperature Transmitter',classification:'TET'}),
  ],'target');
  const reference=snapshot([
    equipment({building:'B1',tag:'B1-MAH-01',parent:'101 Example System',description:'Makeup Air Handler'}),
    equipment({building:'B1',tag:'B1-PNL-01',parent:'101 Example System',description:'Electrical Panel',discipline:'ELECTRICAL'}),
    equipment({building:'B1',tag:'B1-TET-01',parent:'B1-PNL-01',description:'Temperature Transmitter',classification:'TET'}),
  ],'reference');
  const system=compareSsmRegistries(target,reference).systems[0],mah=system.pairs.find(pair=>pair.target&&pair.target.tag==='A1-MAH-01'),tet=system.pairs.find(pair=>pair.target&&pair.target.tag==='A1-TET-01');
  assert.ok(system.treeRoots.includes(mah.id));
  assert.ok(mah.childrenIds.includes(tet.id));
  assert.equal(tet.parentPairId,mah.id);
  assert.equal(tet.placementMismatch,true);
  assert.notEqual(tet.targetParentPairId,tet.referenceParentPairId);
});

test('project comparison promotes New parent references into visible session headers',()=>{
  const target=snapshot([
    equipment({building:'A1',tag:'A1-TET-01',parent:'A1-INSTRUMENTS',description:'Temperature Transmitter',classification:'TET'}),
  ],'target');
  const reference=snapshot([
    equipment({building:'B1',tag:'B1-INSTRUMENTS',parent:'101 Example System',description:'HEADER'}),
    equipment({building:'B1',tag:'B1-TET-01',parent:'B1-INSTRUMENTS',description:'Temperature Transmitter',classification:'TT'}),
  ],'reference');
  const system=compareSsmRegistries(target,reference).systems[0],header=system.pairs.find(pair=>pair.target&&pair.target.isSyntheticHeader),instrument=system.pairs.find(pair=>pair.target&&pair.target.tag==='A1-TET-01');
  assert.equal(system.targetRows,1);
  assert.equal(system.targetHeaders,1);
  assert.ok(header);
  assert.ok(header.childrenIds.includes(instrument.id));
  assert.ok(system.treeRoots.includes(header.id));
});

test('project comparison keeps target-only and completed-project-only systems explicit',()=>{
  const target=snapshot([equipment({building:'A1',tag:'A1-MAH-01',parent:'101 Example System',description:'Makeup Air Handler'})],'target');
  const reference=snapshot([equipment({building:'B1',tag:'B1-PUMP-01',parent:'111 Example System',description:'Centrifugal Pump',upn:'111',discipline:'MECHANICAL WET'})],'reference');
  const result=compareSsmRegistries(target,reference);
  assert.deepEqual(result.systems.map(system=>[system.upn,system.status]),[['101','target-only'],['111','reference-only']]);
  assert.equal(result.summary.targetOnlySystems,1);
  assert.equal(result.summary.referenceOnlySystems,1);
});

test('project comparison is deterministic and handles large flat registries efficiently',()=>{
  const count=6000,targetRows=[],referenceRows=[];
  for(let position=0;position<count;position++){
    const upn=String(101+position%12),suffix=String(position).padStart(5,'0'),description=position%8===0?'Remote I/O Panel':'Centrifugal Pump',discipline=position%8===0?'FACILITIES MONITORING SYSTEM':'MECHANICAL WET';
    targetRows.push(equipment({building:'A',tag:`A-EQ-${suffix}`,parent:`${upn} Example System`,description,upn,discipline,classification:position%8===0?'RIO':'PMP'}));
    referenceRows.push(equipment({building:'Z',tag:`Z-EQ-${suffix}`,parent:`${upn} Example System`,description,upn,discipline,classification:position%8===0?'RIO':'PMP'}));
  }
  const target=snapshot(targetRows,'target'),reference=snapshot(referenceRows,'reference'),started=performance.now(),first=compareSsmRegistries(target,reference),elapsed=performance.now()-started,second=compareSsmRegistries(target,reference);
  assert.equal(first.summary.alignedRows,count);
  assert.equal(first.summary.changedRows,0);
  assert.deepEqual(first.summary,second.summary);
  assert.ok(elapsed<5000,`comparison took ${elapsed.toFixed(0)}ms`);
});
