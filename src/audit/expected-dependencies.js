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
/* Upstream links repeat on every EasyPower row below them, so a tag can be
   mentioned tens of thousands of times. Locations are deduplicated by key and
   only the first few are kept for display; the count stays exact. */
const AUDIT_DEPENDENCY_KEPT_SOURCES=20;
function auditDependencySourceBucket(){return {list:[],keys:new Set(),labels:new Set()};}
function auditDependencySourceKey(source){return `${source.kind}\u0001${source.sheet}\u0001${source.row}`;}
function auditDependencyAddSource(bucket,source,key=auditDependencySourceKey(source)){
  if(bucket.keys.has(key))return;
  bucket.keys.add(key);bucket.labels.add(source.label);
  if(bucket.list.length<AUDIT_DEPENDENCY_KEPT_SOURCES)bucket.list.push(source);
}
function auditDependencyMergeSources(target,bucket){
  for(const source of bucket.list)auditDependencyAddSource(target,source);
  for(const key of bucket.keys)target.keys.add(key);
  for(const label of bucket.labels)target.labels.add(label);
}
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
  const mention=(tag,source)=>{const key=auditNormId(tag);let entry=sourcesOf.get(key);if(!entry){entry={tag,bucket:auditDependencySourceBucket()};sourcesOf.set(key,entry);}auditDependencyAddSource(entry.bucket,source);};
  for(const [kind,label] of Object.entries(AUDIT_DEPENDENCY_SOURCE_KINDS)){
    const reference=references?.[kind];if(!reference?.entries?.length)continue;
    kinds.push(kind);
    for(const entry of reference.entries){
      const child=clean(entry.equipmentId),parent=clean(entry.parent);if(!child)continue;
      const source={kind,label,sheet:clean(entry.sheetName||reference.sheetName),row:entry.sourceRow||0};
      mention(child,source);
      if(!parent||auditNormId(parent)===auditNormId(child))continue;
      mention(parent,source);
      let feeders=feedersOf.get(auditNormId(child));if(!feeders){feeders=new Map();feedersOf.set(auditNormId(child),feeders);}
      const key=auditNormId(parent);let feeder=feeders.get(key);if(!feeder){feeder={tag:parent,bucket:auditDependencySourceBucket()};feeders.set(key,feeder);}
      auditDependencyAddSource(feeder.bucket,source);
    }
  }
  return {feedersOf,sourcesOf,kinds};
}
/* Breadth-first walk to the nearest registry tags upstream of `start`. */
function auditResolveFeeders(start,feedersOf,registry){
  const startKey=auditNormId(start),found=new Map(),seen=new Set([startKey]);let loop=false;
  const queue=[...(feedersOf.get(startKey)?.values()||[])].map(feeder=>({feeder,skipped:[],bucket:feeder.bucket,depth:1}));
  for(let head=0;head<queue.length;head++){
    const {feeder,skipped,bucket,depth}=queue[head],key=auditNormId(feeder.tag);
    if(key===startKey){loop=true;continue;}
    if(registry.has(key)){
      let existing=found.get(key);
      if(!existing){existing={tag:registry.get(key).equipmentId,skipped,bucket:auditDependencySourceBucket()};found.set(key,existing);}
      auditDependencyMergeSources(existing.bucket,bucket);
      continue;
    }
    if(seen.has(key))continue;
    seen.add(key);
    if(depth>=AUDIT_DEPENDENCY_MAX_DEPTH){loop=true;continue;}
    for(const next of feedersOf.get(key)?.values()||[])queue.push({feeder:next,skipped:[...skipped,feeder.tag],bucket,depth:depth+1});
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
      if(reason)expected.push({tag:feeder.tag,reason,skipped:feeder.skipped,sources:feeder.bucket.list,sourceCount:feeder.bucket.keys.size});
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
  for(const [key,{tag,bucket}] of sourcesOf){
    if(registry.has(key))continue;
    notInRegistry.push({tag,sources:bucket.list,sourceCount:bucket.keys.size,documents:[...bucket.labels],affects:[...(affects.get(key)||[])]});
  }
  const byTag=(left,right)=>clean(left.equipmentId??left.tag).localeCompare(clean(right.equipmentId??right.tag),undefined,{numeric:true});
  rows.sort(byTag);notInRegistry.sort(byTag);
  const match=rows.filter(entry=>entry.status==='match').length;
  return {rows,notInRegistry,counts:{compared:rows.length,match,missing:rows.length-match,notInRegistry:notInRegistry.length},kinds};
}

/* ---- Power hierarchy for one tag (on-demand double check) ----
   Every upstream path from a Starting Source down to the tag, what the tag
   feeds, and the rule's verdict on each link. Pure; computed on click. */
const AUDIT_POWER_MAX_PATHS=12,AUDIT_POWER_MAX_CHILDREN=50;
function auditPowerContext(row){return row?{building:clean(row.building),discipline:clean(row.discipline),systemName:clean(row.systemName),upn:clean(row.upn)}:null;}
function auditPowerVerdict(child,feeder,childTag='',feederTag=''){
  if(!child||!feeder){
    const absent=[!feeder&&feederTag,!child&&childTag].filter(Boolean);
    return {kind:'unverified',reason:absent.length?`${absent.join(' and ')} ${absent.length===1?'is':'are'} not in the registry`:'Not compared: one end is not in the registry'};
  }
  if(!auditDependencyGroup(child)||!auditDependencyGroup(feeder))return {kind:'unverified',reason:'Not compared: Building, Discipline or System Name is blank'};
  const crossing=auditDependencyCrossing(child,feeder);
  if(crossing)return {kind:'dependency',reason:crossing};
  return {kind:'nests',reason:auditDependencyGroup(child).discipline==='ELECTRICAL'&&auditDependencyGroup(feeder).discipline==='ELECTRICAL'&&auditNormId(child.systemName)!==auditNormId(feeder.systemName)?'Electrical feed inside one building':'Same Building, Discipline and System Name'};
}
export function auditPowerHierarchy(snapshot,references,tag){
  const {feedersOf,sourcesOf}=auditDependencyEdges(references),registry=new Map();
  for(const row of snapshot?.rows||[]){const key=auditNormId(row.equipmentId);if(key&&!registry.has(key))registry.set(key,row);}
  const startKey=auditNormId(tag),spell=key=>registry.get(key)?.equipmentId||sourcesOf.get(key)?.tag||key;
  const link=(childKey,parentKey)=>{
    const bucket=feedersOf.get(childKey)?.get(parentKey)?.bucket;
    return {...auditPowerVerdict(registry.get(childKey),registry.get(parentKey),spell(childKey),spell(parentKey)),sources:bucket?bucket.list:[],sourceCount:bucket?bucket.keys.size:0};
  };
  const paths=[];let loop=false,truncated=false;
  const climb=trail=>{
    if(truncated)return;
    const key=trail[trail.length-1],next=[];
    for(const parentKey of feedersOf.get(key)?.keys()||[]){if(trail.includes(parentKey))loop=true;else next.push(parentKey);}
    if(!next.length||trail.length>=AUDIT_DEPENDENCY_MAX_DEPTH){
      if(trail.length<2)return;
      if(paths.length>=AUDIT_POWER_MAX_PATHS){truncated=true;return;}
      paths.push([...trail].reverse());return;
    }
    for(const parentKey of next)climb([...trail,parentKey]);
  };
  climb([startKey]);
  const children=[];let childCount=0;
  for(const [childKey,feeders] of feedersOf){
    if(!feeders.has(startKey)||childKey===startKey)continue;
    childCount++;
    if(children.length<AUDIT_POWER_MAX_CHILDREN)children.push({tag:spell(childKey),registry:auditPowerContext(registry.get(childKey)),link:link(childKey,startKey)});
  }
  const nearest=registry.has(startKey)?auditResolveFeeders(registry.get(startKey).equipmentId,feedersOf,registry).found.map(found=>{
    const verdict=auditPowerVerdict(registry.get(startKey),registry.get(auditNormId(found.tag)));
    return {tag:found.tag,skipped:found.skipped,kind:verdict.kind,reason:verdict.reason};
  }):[];
  return {
    tag:spell(startKey),registry:auditPowerContext(registry.get(startKey)),
    paths:paths.map(path=>path.map((key,index)=>({tag:spell(key),registry:auditPowerContext(registry.get(key)),link:index?link(key,path[index-1]):null}))),
    truncated,loop,children,childCount,nearest,
  };
}
