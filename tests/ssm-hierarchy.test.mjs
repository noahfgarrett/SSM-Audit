import test from 'node:test'
import assert from 'node:assert/strict'

import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { buildSsmHierarchy } from '../src/audit/hierarchy.js'

const headers=EXTO_REV21_COLUMNS.map(column=>column.header)
const index=Object.fromEntries(EXTO_REV21_COLUMNS.map(column=>[column.field,column.index]))
function row(values){const cells=new Array(headers.length).fill('');for(const [field,value] of Object.entries(values))cells[index[field]]=value;return cells;}
function snapshot(values){return auditSnapshotFromAoa([headers,...values],{file:'registry.xlsx',sheet:'Registry'});}
function equipment({tag,parent,description,building='B1',discipline='MECHANICAL DRY',upn='101',systemName=`${upn} Example System`,status='Existing',dependencies=''}){
  return row({building,discipline,upn,systemName,equipmentId:tag,closestParent:parent,closestParentStatus:status,equipmentDescription:description,dependencies});
}
function equipmentNode(hierarchy,tag){return [...hierarchy.nodeByKey.values()].find(node=>node.type==='equipment'&&node.tag===tag);}

test('single-registry hierarchy follows Building, Discipline, System, and closest-parent nesting',()=>{
  const hierarchy=buildSsmHierarchy(snapshot([
    equipment({tag:'B1-MAH-01',parent:'101 Example System',description:'Makeup Air Handler'}),
    equipment({tag:'B1-TET-01',parent:'B1-MAH-01',description:'Temperature Transmitter',dependencies:'B1-RIO-01'}),
  ]));
  const building=hierarchy.nodeByKey.get(hierarchy.rootKeys[0]),discipline=hierarchy.nodeByKey.get(building.children[0]),system=hierarchy.nodeByKey.get(discipline.children[0]),mah=equipmentNode(hierarchy,'B1-MAH-01'),tet=equipmentNode(hierarchy,'B1-TET-01');
  assert.deepEqual([building.type,discipline.type,system.type],['building','discipline','system']);
  assert.equal(mah.parentKey,system.key);
  assert.equal(tet.parentKey,mah.key);
  assert.ok(mah.children.includes(tet.key));
  assert.equal(hierarchy.summary.dependencies,1);
  assert.equal(building.equipmentCount,2);
});

test('single-registry hierarchy materializes New parent references as visible headers',()=>{
  const hierarchy=buildSsmHierarchy(snapshot([
    equipment({tag:'B1-TET-01',parent:'B1-INSTRUMENTS',description:'Temperature Transmitter',status:'New'}),
  ]));
  const header=equipmentNode(hierarchy,'B1-INSTRUMENTS'),instrument=equipmentNode(hierarchy,'B1-TET-01');
  assert.ok(header);
  assert.equal(header.isSyntheticHeader,true);
  assert.equal(instrument.parentKey,header.key);
  assert.ok(header.children.includes(instrument.key));
  assert.equal(hierarchy.summary.equipment,1);
  assert.equal(hierarchy.summary.generatedHeaders,1);
});

test('single-registry hierarchy keeps unresolved parents visible at the assigned system root',()=>{
  const hierarchy=buildSsmHierarchy(snapshot([
    equipment({tag:'B1-PUMP-01',parent:'B1-MISSING-HEADER',description:'Centrifugal Pump',status:'Unknown'}),
  ]));
  const pump=equipmentNode(hierarchy,'B1-PUMP-01'),system=hierarchy.nodeByKey.get(pump.parentKey);
  assert.equal(pump.unresolvedParent,true);
  assert.equal(system.type,'system');
  assert.equal(hierarchy.summary.unresolvedParents,1);
});

test('single-registry hierarchy breaks parent cycles without dropping equipment',()=>{
  const hierarchy=buildSsmHierarchy(snapshot([
    equipment({tag:'B1-EQ-A',parent:'B1-EQ-B',description:'Equipment A'}),
    equipment({tag:'B1-EQ-B',parent:'B1-EQ-A',description:'Equipment B'}),
  ]));
  const first=equipmentNode(hierarchy,'B1-EQ-A'),second=equipmentNode(hierarchy,'B1-EQ-B');
  assert.equal(first.cycleBreak,true);
  assert.equal(second.cycleBreak,true);
  assert.equal(hierarchy.nodeByKey.get(first.parentKey).type,'system');
  assert.equal(hierarchy.nodeByKey.get(second.parentKey).type,'system');
  assert.equal(hierarchy.summary.cycleBreaks,2);
  assert.equal(hierarchy.summary.equipment,2);
});

test('single-registry hierarchy aggregates equipment findings through its branches',()=>{
  const target=snapshot([
    equipment({tag:'B1-MAH-01',parent:'101 Example System',description:'Makeup Air Handler'}),
    equipment({tag:'B1-TET-01',parent:'B1-MAH-01',description:'Temperature Transmitter'}),
  ]),finding={equipmentId:'B1-TET-01',severity:'warning',why:'Synthetic finding'};
  const hierarchy=buildSsmHierarchy(target,[finding]),instrument=equipmentNode(hierarchy,'B1-TET-01'),building=hierarchy.nodeByKey.get(hierarchy.rootKeys[0]);
  assert.equal(instrument.findingCount,1);
  assert.deepEqual(instrument.findings,[finding]);
  assert.equal(building.findingCount,1);
});

test('single-registry hierarchy handles large registries without recursive parent resolution',()=>{
  const values=[],count=8000;for(let position=0;position<count;position++)values.push(equipment({tag:`B1-EQ-${String(position).padStart(5,'0')}`,parent:'101 Example System',description:'Equipment'}));
  const started=performance.now(),hierarchy=buildSsmHierarchy(snapshot(values)),elapsed=performance.now()-started;
  assert.equal(hierarchy.summary.equipment,count);
  assert.ok(elapsed<3000,`hierarchy build took ${elapsed.toFixed(0)}ms`);
});

test('single-registry hierarchy counts a deeply nested branch without overflowing the stack',()=>{
  const values=[],count=20000;
  for(let position=0;position<count;position++)values.push(equipment({tag:`B1-EQ-${String(position).padStart(5,'0')}`,parent:position?`B1-EQ-${String(position-1).padStart(5,'0')}`:'101 Example System',description:'Equipment'}));
  const hierarchy=buildSsmHierarchy(snapshot(values)),building=hierarchy.nodeByKey.get(hierarchy.rootKeys[0]);
  assert.equal(hierarchy.summary.equipment,count);
  assert.equal(building.equipmentCount,count);
});

test('single-registry hierarchy accepts an Existing parent outside the local registry',()=>{
  const hierarchy=buildSsmHierarchy(snapshot([
    equipment({tag:'B1-PUMP-01',parent:'OTHER-PROJECT-PARENT',description:'Centrifugal Pump',status:'Existing'}),
  ])),pump=equipmentNode(hierarchy,'B1-PUMP-01');
  assert.equal(pump.unresolvedParent,false);
  assert.equal(hierarchy.summary.unresolvedParents,0);
  assert.equal(hierarchy.nodeByKey.get(pump.parentKey).type,'system');
});

test('duplicate Equipment IDs receive only findings from their own source row',()=>{
  const target=snapshot([
    equipment({tag:'B1-DUPLICATE',parent:'101 Example System',description:'Pump'}),
    equipment({tag:'B1-DUPLICATE',parent:'101 Example System',description:'Pump'}),
  ]),findings=target.rows.map((item,index)=>({equipmentId:item.equipmentId,severity:'warning',why:`Finding ${index+1}`,sheet:item._source.sheet,row:item._source.row}));
  const hierarchy=buildSsmHierarchy(target,findings),duplicates=[...hierarchy.nodeByKey.values()].filter(node=>node.type==='equipment'&&node.tag==='B1-DUPLICATE'),building=hierarchy.nodeByKey.get(hierarchy.rootKeys[0]);
  assert.equal(duplicates.length,2);
  assert.deepEqual(duplicates.map(node=>node.findingCount),[1,1]);
  assert.equal(building.findingCount,2);
});
