import { clean, natCmp } from '../core/text.js'
import { auditNormId, auditSplitReferences } from './model.js'
import { auditRegistryModel } from './compare.js'

export const SSM_HIERARCHY_SCHEMA_VERSION=1;
const UNASSIGNED_BUILDING='Unassigned Building',UNASSIGNED_DISCIPLINE='Unassigned Discipline',UNASSIGNED_SYSTEM='Unassigned System';

function normalizedKey(value,fallback){return auditNormId(value)||fallback;}
function groupIdentity(node){
  const row=node.row||{},building=clean(row.building)||UNASSIGNED_BUILDING,discipline=clean(row.discipline)||UNASSIGNED_DISCIPLINE,upn=clean(node.upn)||'UNASSIGNED',systemName=clean(row.systemName)||(upn==='UNASSIGNED'?UNASSIGNED_SYSTEM:`UPN ${upn}`);
  const buildingKey=`building|${normalizedKey(building,'UNASSIGNED')}`,disciplineKey=`discipline|${buildingKey}|${normalizedKey(discipline,'UNASSIGNED')}`,systemKey=`system|${disciplineKey}|${normalizedKey(upn,'UNASSIGNED')}|${normalizedKey(systemName,'UNASSIGNED')}`;
  return {building,discipline,upn,systemName,buildingKey,disciplineKey,systemKey};
}
function sourceOrder(node){const source=node.row&&node.row._source||{};return `${clean(source.sheet)}\u0001${String(source.row||0).padStart(10,'0')}\u0001${node.tag}`;}
function parentCandidate(item,candidates){
  const available=(candidates||[]).filter(candidate=>candidate.key!==item.key);if(!available.length)return null;
  return [...available].sort((left,right)=>{
    const score=candidate=>(candidate.systemKey===item.systemKey?8:0)+(candidate.upn===item.upn?4:0)+(candidate.buildingKey===item.buildingKey?2:0)+(candidate.disciplineKey===item.disciplineKey?1:0)+(candidate.isSyntheticHeader?0:.25);
    return score(right)-score(left)||natCmp(sourceOrder(left),sourceOrder(right));
  })[0];
}
function hierarchyNodeFromRegistry(node,findingsByTag){
  const group=groupIdentity(node),row=node.row||{},dependencies=auditSplitReferences(row.dependencies).filter(reference=>!/^N\/?A$/i.test(reference)),findings=findingsByTag.get(node.id)||[];
  return {...group,key:`equipment|${node.key}`,type:'equipment',tag:clean(node.tag),label:clean(node.tag),description:clean(node.role),classification:clean(node.classification),normalizedTag:node.id,closestParent:clean(row.closestParent),closestParentStatus:clean(row.closestParentStatus),dependencies,isHeader:!!node.isHeader,isSyntheticHeader:!!node.isSyntheticHeader,source:row._source||{},children:[],parentKey:'',unresolvedParent:false,cycleBreak:false,findings:[...findings],findingCount:findings.length,equipmentCount:node.isSyntheticHeader?0:1,dependencyCount:dependencies.length,searchKey:auditNormId([node.tag,node.role,node.classification,row.building,row.discipline,row.upn,row.systemName,row.closestParent,dependencies.join(' ')].join(' '))};
}
function groupNode(key,type,label,metadata={}){return {key,type,label,description:type==='building'?'Building':type==='discipline'?'Discipline':'System Name',children:[],parentKey:'',equipmentCount:0,findingCount:0,dependencyCount:0,searchKey:auditNormId([label,metadata.upn].join(' ')),...metadata};}
function groupSort(left,right){
  const rank={building:0,discipline:1,system:2,equipment:3};return (rank[left.type]-rank[right.type])||natCmp(left.upn||'',right.upn||'')||natCmp(left.label,right.label);
}
function equipmentSort(left,right){return Number(right.isHeader)-Number(left.isHeader)||natCmp(left.description,right.description)||natCmp(left.tag,right.tag)||natCmp(sourceOrder(left),sourceOrder(right));}
function parentCycleKeys(equipment,desiredParents){
  const done=new Set(),cycles=new Set();
  for(const start of equipment){
    if(done.has(start.key))continue;const path=[],position=new Map();let current=start;
    while(current&&!done.has(current.key)){
      if(position.has(current.key)){for(let index=position.get(current.key);index<path.length;index++)cycles.add(path[index].key);break;}
      position.set(current.key,path.length);path.push(current);current=desiredParents.get(current.key)||null;
    }
    for(const node of path)done.add(node.key);
  }
  return cycles;
}
function annotateCounts(nodeByKey,key,visiting=new Set()){
  const node=nodeByKey.get(key);if(!node||visiting.has(key))return {equipment:0,findings:0,dependencies:0};visiting.add(key);
  let equipment=node.type==='equipment'&&!node.isSyntheticHeader?1:0,findings=node.type==='equipment'?node.findings.length:0,dependencies=node.type==='equipment'?node.dependencies.length:0;
  for(const childKey of node.children){const child=annotateCounts(nodeByKey,childKey,visiting);equipment+=child.equipment;findings+=child.findings;dependencies+=child.dependencies;}
  visiting.delete(key);node.equipmentCount=equipment;node.findingCount=findings;node.dependencyCount=dependencies;return {equipment,findings,dependencies};
}

export function buildSsmHierarchy(snapshot,findings=[]){
  const registry=auditRegistryModel(snapshot),findingsByTag=new Map();
  for(const finding of findings){const key=auditNormId(finding.equipmentId);if(!key)continue;const list=findingsByTag.get(key)||[];list.push(finding);findingsByTag.set(key,list);}
  const equipment=registry.nodes.map(node=>hierarchyNodeFromRegistry(node,findingsByTag)),nodeByKey=new Map(equipment.map(node=>[node.key,node])),equipmentByTag=new Map();
  for(const node of equipment){const list=equipmentByTag.get(node.normalizedTag)||[];list.push(node);equipmentByTag.set(node.normalizedTag,list);}
  const desiredParents=new Map();
  for(const node of equipment){
    const parentId=auditNormId(node.closestParent),systemId=auditNormId(node.systemName);if(!parentId||parentId===systemId)continue;
    const parent=parentCandidate(node,equipmentByTag.get(parentId));if(parent)desiredParents.set(node.key,parent);else node.unresolvedParent=true;
  }
  const cycleKeys=parentCycleKeys(equipment,desiredParents);
  const buildings=new Map(),disciplines=new Map(),systems=new Map();
  function ensureGroups(node){
    let building=buildings.get(node.buildingKey);if(!building){building=groupNode(node.buildingKey,'building',node.building,{buildingKey:node.buildingKey});buildings.set(node.buildingKey,building);nodeByKey.set(building.key,building);}
    let discipline=disciplines.get(node.disciplineKey);if(!discipline){discipline=groupNode(node.disciplineKey,'discipline',node.discipline,{buildingKey:node.buildingKey,disciplineKey:node.disciplineKey});discipline.parentKey=building.key;building.children.push(discipline.key);disciplines.set(node.disciplineKey,discipline);nodeByKey.set(discipline.key,discipline);}
    let system=systems.get(node.systemKey);if(!system){system=groupNode(node.systemKey,'system',node.systemName,{buildingKey:node.buildingKey,disciplineKey:node.disciplineKey,systemKey:node.systemKey,upn:node.upn});system.parentKey=discipline.key;discipline.children.push(system.key);systems.set(node.systemKey,system);nodeByKey.set(system.key,system);}
    return system;
  }
  for(const node of equipment){
    const parent=desiredParents.get(node.key),cycle=cycleKeys.has(node.key);if(cycle)node.cycleBreak=true;
    if(parent&&!cycle){node.parentKey=parent.key;parent.children.push(node.key);}
    else{const system=ensureGroups(node);node.parentKey=system.key;system.children.push(node.key);}
  }
  for(const node of equipment)node.children.sort((left,right)=>equipmentSort(nodeByKey.get(left),nodeByKey.get(right)));
  for(const node of systems.values())node.children.sort((left,right)=>equipmentSort(nodeByKey.get(left),nodeByKey.get(right)));
  for(const node of disciplines.values())node.children.sort((left,right)=>groupSort(nodeByKey.get(left),nodeByKey.get(right)));
  for(const node of buildings.values())node.children.sort((left,right)=>groupSort(nodeByKey.get(left),nodeByKey.get(right)));
  const rootKeys=[...buildings.values()].sort(groupSort).map(node=>node.key);for(const key of rootKeys)annotateCounts(nodeByKey,key);
  const summary={rows:registry.rows.length,equipment:equipment.filter(node=>!node.isSyntheticHeader).length,generatedHeaders:equipment.filter(node=>node.isSyntheticHeader).length,buildings:buildings.size,disciplines:disciplines.size,systems:systems.size,dependencies:equipment.reduce((total,node)=>total+node.dependencies.length,0),findings:findings.length,unresolvedParents:equipment.filter(node=>node.unresolvedParent).length,cycleBreaks:equipment.filter(node=>node.cycleBreak).length};
  return Object.freeze({schemaVersion:SSM_HIERARCHY_SCHEMA_VERSION,standard:'Registry SSM hierarchy',rootKeys:Object.freeze(rootKeys),nodeByKey,summary:Object.freeze(summary),groups:Object.freeze({buildings:Object.freeze([...buildings.values()].sort(groupSort)),disciplines:Object.freeze([...disciplines.values()].sort(groupSort)),systems:Object.freeze([...systems.values()].sort(groupSort))})});
}
