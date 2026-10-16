import { clean } from '../core/text.js'
import { auditNormId, auditSplitReferences } from './model.js'

/* Expected dependencies from power and control documents.
   Mirrors SSManagement's partition fold: every EasyPower / Cable Schedule / PMD
   connection starts as a parent link, and a link whose two ends sit in different
   registry groups (Building | Discipline | System Name) is a dependency instead.
   The electrical feed chain may nest across discipline and system inside one
   building. A group with a blank part is unknown, and unknown never demotes —
   a crossing must be proven, not assumed. Pure: snapshot + references in. */

export const AUDIT_DEPENDENCY_SOURCE_KINDS=Object.freeze({easyPower:'EasyPower',cable:'Cable Schedule',pmd:'PMD'});
const AUDIT_DEPENDENCY_MAX_DEPTH=25;
const AUDIT_DEPENDENCY_GROUP_PARTS=Object.freeze([['building','building'],['discipline','discipline'],['systemName','system']]);

function auditDependencyGroup(row){
  const group={};
  for(const [field] of AUDIT_DEPENDENCY_GROUP_PARTS){const value=auditNormId(row?.[field]);if(!value)return null;group[field]=value;}
  return group;
}
/* Returns the reason U is a dependency of T, or '' when the link nests. */
function auditDependencyCrossing(child,feeder){
  const childGroup=auditDependencyGroup(child),feederGroup=auditDependencyGroup(feeder);
  if(!childGroup||!feederGroup)return '';
  const differs=AUDIT_DEPENDENCY_GROUP_PARTS.filter(([field])=>childGroup[field]!==feederGroup[field]);
  if(!differs.length)return '';
  if(childGroup.discipline==='ELECTRICAL'&&feederGroup.discipline==='ELECTRICAL'&&childGroup.building===feederGroup.building)return '';
  return differs.map(([field,label])=>`different ${label} (${clean(child[field])} vs ${clean(feeder[field])})`).join(', ');
}
/* Dependency i belongs to Dependency Project i (or the single listed project). */
function auditLocalDependencies(row){
  const dependencies=auditSplitReferences(row.dependencies),projects=auditSplitReferences(row.dependencyProject),own=auditNormId(row.project);
  return new Set(dependencies.filter((_,index)=>{
    const project=projects.length===1?projects[0]:projects[index];
    return !clean(project)||auditNormId(project)===own;
  }).map(auditNormId));
}
function auditDependencyEdges(references){
  const feedersOf=new Map(),sourcesOf=new Map(),kinds=[];
  const remember=(map,key,value)=>{let list=map.get(key);if(!list){list=[];map.set(key,list);}list.push(value);};
  for(const [kind,label] of Object.entries(AUDIT_DEPENDENCY_SOURCE_KINDS)){
    const reference=references?.[kind];if(!reference?.entries?.length)continue;
    kinds.push(kind);
    for(const entry of reference.entries){
      const child=clean(entry.equipmentId),parent=clean(entry.parent);if(!child)continue;
      const source={kind,label,sheet:clean(entry.sheetName||reference.sheetName),row:entry.sourceRow||0};
      remember(sourcesOf,auditNormId(child),{tag:child,source});
      if(!parent||auditNormId(parent)===auditNormId(child))continue;
      remember(sourcesOf,auditNormId(parent),{tag:parent,source});
      let feeders=feedersOf.get(auditNormId(child));if(!feeders){feeders=new Map();feedersOf.set(auditNormId(child),feeders);}
      const key=auditNormId(parent);let feeder=feeders.get(key);if(!feeder){feeder={tag:parent,sources:[]};feeders.set(key,feeder);}
      if(!feeder.sources.some(item=>item.kind===source.kind&&item.sheet===source.sheet&&item.row===source.row))feeder.sources.push(source);
    }
  }
  return {feedersOf,sourcesOf,kinds};
}
/* Breadth-first walk to the nearest registry tags upstream of `start`. */
function auditResolveFeeders(start,feedersOf,registry){
  const startKey=auditNormId(start),found=new Map(),seen=new Set([startKey]);let loop=false;
  const queue=[...(feedersOf.get(startKey)?.values()||[])].map(feeder=>({feeder,skipped:[],sources:feeder.sources,depth:1}));
  while(queue.length){
    const {feeder,skipped,sources,depth}=queue.shift(),key=auditNormId(feeder.tag);
    if(key===startKey){loop=true;continue;}
    if(registry.has(key)){
      const existing=found.get(key);
      if(!existing)found.set(key,{tag:registry.get(key).equipmentId,skipped,sources:[...sources]});
      else for(const source of sources)if(!existing.sources.some(item=>item.kind===source.kind&&item.sheet===source.sheet&&item.row===source.row))existing.sources.push(source);
      continue;
    }
    if(seen.has(key))continue;
    seen.add(key);
    if(depth>=AUDIT_DEPENDENCY_MAX_DEPTH){loop=true;continue;}
    for(const next of feedersOf.get(key)?.values()||[])queue.push({feeder:next,skipped:[...skipped,feeder.tag],sources,depth:depth+1});
  }
  return {found:[...found.values()],loop};
}

export function auditExpectedDependencies(snapshot,references={}){
  const {feedersOf,sourcesOf,kinds}=auditDependencyEdges(references),registry=new Map();
  for(const row of snapshot?.rows||[]){const key=auditNormId(row.equipmentId);if(key&&!registry.has(key))registry.set(key,row);}
  const rows=[],affects=new Map();
  for(const [key,row] of registry){
    if(!feedersOf.has(key))continue;
    const {found,loop}=auditResolveFeeders(row.equipmentId,feedersOf,registry),expected=[];
    for(const feeder of found){
      const reason=auditDependencyCrossing(row,registry.get(auditNormId(feeder.tag)));
      for(const tag of feeder.skipped){let list=affects.get(auditNormId(tag));if(!list){list=new Set();affects.set(auditNormId(tag),list);}list.add(row.equipmentId);}
      if(reason)expected.push({tag:feeder.tag,reason,skipped:feeder.skipped,sources:feeder.sources});
    }
    if(!expected.length)continue;
    const local=auditLocalDependencies(row),parent=auditNormId(row.closestParent);
    const missing=expected.filter(entry=>!local.has(auditNormId(entry.tag))).map(entry=>entry.tag);
    const notes=[];
    for(const tag of missing)if(auditNormId(tag)===parent)notes.push(`${tag} is listed as Closest Parent, not as a dependency`);
    const skipped=[...new Set(expected.flatMap(entry=>entry.skipped))];
    if(skipped.length)notes.push(`${skipped.join(', ')} ${skipped.length===1?'is':'are'} not in the registry; compared against the next tag upstream`);
    if(loop)notes.push('The source chain loops back to this tag');
    rows.push({equipmentId:row.equipmentId,status:missing.length?'missing':'match',expected,missing,notes,
      registryDependencies:auditSplitReferences(row.dependencies),closestParent:clean(row.closestParent),
      building:clean(row.building),discipline:clean(row.discipline),upn:clean(row.upn),systemName:clean(row.systemName),_source:row._source||null});
  }
  const notInRegistry=[];
  for(const [key,mentions] of sourcesOf){
    if(registry.has(key))continue;
    const sources=[];
    for(const {source} of mentions)if(!sources.some(item=>item.kind===source.kind&&item.sheet===source.sheet&&item.row===source.row))sources.push(source);
    notInRegistry.push({tag:mentions[0].tag,sources,affects:[...(affects.get(key)||[])]});
  }
  const byTag=(left,right)=>clean(left.equipmentId??left.tag).localeCompare(clean(right.equipmentId??right.tag),undefined,{numeric:true});
  rows.sort(byTag);notInRegistry.sort(byTag);
  const match=rows.filter(entry=>entry.status==='match').length;
  return {rows,notInRegistry,counts:{compared:rows.length,match,missing:rows.length-match,notInRegistry:notInRegistry.length},kinds};
}
