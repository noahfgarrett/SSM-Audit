import { clean } from '../core/text.js'
import { EXTO_REV21_COLUMNS, extoRev21SystemsForUpn, extoRev21EffectiveDiscipline, extoRev21Canonical, extoRev21UpnCandidates } from '../exto/rev21-contract.js'
import { auditNormId, auditSourceKey, auditSplitReferences } from './model.js'
import { auditCommissioningRole, auditIsBlankItemMaster, auditMilestoneBranchCandidates, auditMilestoneCohortCandidates, auditItemMasterCanonicalCandidates } from './engine.js'
import { VF_ITEM_MASTER_NAMES } from '../exto/vf-item-masters.js'
import { auditReferenceRecommendation } from './references.js'
import { auditReadMigrationSettings } from './milestone-migration.js'

export const AUDIT_ACTION_FIELDS=Object.freeze({
  'UPN':'upn','Discipline':'discipline','System Name':'systemName','Closest Parent':'closestParent',
  'Dependencies':'dependencies','Dependency Project':'dependencyProject','L2 Milestone':'milestone',
  'L1 Milestone Parent':'milestoneParent','Milestone Parent':'milestoneParent',
  'Item Master Unique Identifier':'itemMaster','Equipment Classification':'equipmentClassification',
});
const AUDIT_ACTION_COLUMNS=new Map(EXTO_REV21_COLUMNS.map(column=>[column.field,column]));
const AUDIT_ACTION_PROPS=new Set(Object.values(AUDIT_ACTION_FIELDS));

export function auditCorrectionKey(change){return JSON.stringify([change.source?.sheet||'',change.source?.row||0,change.prop||AUDIT_ACTION_FIELDS[change.field]]);}
export function auditActionIndex(snapshot){
  const bySource=new Map(),byTag=new Map(),bySystem=new Map();
  for(const row of snapshot.rows){
    bySource.set(auditSourceKey(row),row);
    const tag=auditNormId(row.equipmentId),tags=byTag.get(tag)||[];tags.push(row);byTag.set(tag,tags);
    const group=JSON.stringify([auditNormId(row.building),auditNormId(row.upn)]),peers=bySystem.get(group)||[];peers.push(row);bySystem.set(group,peers);
  }
  return {bySource,byTag,bySystem};
}
export function auditFindingRow(finding,index){return index.bySource.get(JSON.stringify([finding.sheet||'',finding.row||0]))||null;}
export function auditMakeCorrection(row,field,value,finding,reason=''){
  const prop=AUDIT_ACTION_FIELDS[field],column=AUDIT_ACTION_COLUMNS.get(prop);
  if(!prop||!row?._source||row._source.columns?.[prop]==null)throw new Error('This field is not available on the source row.');
  return {tag:row.equipmentId,field,prop,header:column.header,before:clean(row[prop]),value:clean(value),source:row._source,ruleId:finding?.rule?.id||'',findingId:finding?.id||'',reason};
}
export function auditMergeCorrections(baseline,existing,incoming){
  const index=auditActionIndex(baseline),merged=new Map(existing.map(change=>[auditCorrectionKey(change),change]));
  const batch=new Map();
  for(const change of incoming){
    const row=index.bySource.get(auditSourceKey({_source:change.source})),prop=change.prop||AUDIT_ACTION_FIELDS[change.field];
    if(!row||auditNormId(row.equipmentId)!==auditNormId(change.tag)||!AUDIT_ACTION_PROPS.has(prop))throw new Error('A correction no longer matches its source row.');
    const key=auditCorrectionKey(change),value=clean(change.value);
    if(batch.has(key)&&batch.get(key)!==value)throw new Error('This batch proposes different values for the same cell. Review those rows separately.');
    batch.set(key,value);
    const next={...change,prop,header:AUDIT_ACTION_COLUMNS.get(prop).header,source:row._source,before:clean(row[prop]),value};
    if(value===next.before)merged.delete(key);else merged.set(key,next);
  }
  return [...merged.values()];
}
export function auditApplyCorrections(baseline,changes){
  const index=auditActionIndex(baseline),patches=new Map(),seen=new Set();
  for(const change of changes){
    const prop=change.prop||AUDIT_ACTION_FIELDS[change.field],key=auditSourceKey({_source:change.source}),row=index.bySource.get(key);
    if(!row||!AUDIT_ACTION_PROPS.has(prop)||AUDIT_ACTION_FIELDS[change.field]!==prop||row._source.columns?.[prop]==null||auditNormId(row.equipmentId)!==auditNormId(change.tag))throw new Error('A correction does not match this registry.');
    if(typeof change.value!=='string'||change.value.length>32767||clean(row[prop])!==change.before)throw new Error('The original cell value has changed. Review the correction again.');
    const cell=JSON.stringify([key,prop]);if(seen.has(cell))throw new Error('The same cell is corrected more than once.');seen.add(cell);
    const values=patches.get(key)||{};values[prop]=change.value;patches.set(key,values);
  }
  const rows=baseline.rows.map(row=>patches.has(auditSourceKey(row))?Object.freeze({...row,...patches.get(auditSourceKey(row))}):row);
  return Object.freeze({...baseline,rows:Object.freeze(rows)});
}
export function auditCorrectionImpact(before,after){
  const oldIds=new Set(before.findings.map(finding=>finding.id)),newIds=new Set(after.findings.map(finding=>finding.id));
  const key=finding=>JSON.stringify([finding.rule.id,finding.sheet,finding.row,finding.field]);
  const remaining=new Map(),changed=[],resolved=[];
  for(const finding of after.findings)if(!oldIds.has(finding.id)){const group=remaining.get(key(finding))||[];group.push(finding);remaining.set(key(finding),group);}
  for(const finding of before.findings)if(!newIds.has(finding.id)){
    const group=remaining.get(key(finding));if(group?.length)changed.push({before:finding,after:group.shift()});else resolved.push(finding);
  }
  const introduced=[...remaining.values()].flat();
  return {resolved,introduced,changed,unsafe:[...introduced,...changed.map(pair=>pair.after)].filter(finding=>finding.severity==='blocker'||finding.severity==='error')};
}
export async function auditRegistryRevision(snapshot){
  return auditReviewDigest(snapshot.rows.map(row=>[auditSourceKey(row),EXTO_REV21_COLUMNS.map(column=>clean(row[column.field])),row._source?.columns||{},row._sources||[]]));
}
async function auditReviewDigest(value){
  if(!globalThis.crypto?.subtle)throw new Error('Review-file verification is unavailable in this browser.');
  const bytes=await globalThis.crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(value)));
  return [...new Uint8Array(bytes)].map(value=>value.toString(16).padStart(2,'0')).join('');
}
export async function auditReviewDocument(session){
  return {format:'ssm-audit-review',version:1,baselineRevision:await auditRegistryRevision(session.baselineSnapshot||session.snapshot),
    referencesRevision:await auditReviewReferenceRevision(session.references),changes:session.changes||[],actioned:[...(session.actioned||[])],reviewed:[...(session.reviewedIds||session.actioned||[])],excluded:[...(session.excluded||[])],history:session.reviewHistory||[],filterViews:session.filterViews||[],
    milestoneMigration:session.milestoneMigration||{enabled:false,profile:null},createdAt:new Date().toISOString()};
}
function auditReviewReferenceRevision(references={}){return auditReviewDigest(['milestones','itemMasters'].map(kind=>[kind,references[kind]||null]));}
function auditReviewFilterViews(views=[]){
  if(!Array.isArray(views)||views.length>100)throw new Error('The review file contains invalid filter views.');
  const list=values=>{
    if(values===undefined)return [];
    if(!Array.isArray(values)||values.length>100000||values.some(value=>typeof value!=='string'||value.length>4000))throw new Error('The review file contains invalid filter views.');
    return [...values];
  };
  return views.map(view=>{
    if(!view||typeof view.name!=='string'||view.name.length>40||!view.filters||typeof view.filters!=='object')throw new Error('The review file contains invalid filter views.');
    const filters=Object.fromEntries(['hiddenSeverities','hiddenSources','hiddenCategories','hiddenRules'].map(key=>[key,list(view.filters[key])]));
    filters.dimFilters=Object.fromEntries(['discipline','milestone','upn','building'].map(key=>[key,list(view.filters.dimFilters?.[key])]));
    return {name:view.name,filters};
  });
}
export async function auditReadReviewDocument(document,baseline,references={}){
  if(document?.format!=='ssm-audit-review'||document.version!==1||document.baselineRevision!==await auditRegistryRevision(baseline))throw new Error('This review belongs to a different registry revision. No changes were loaded.');
  if(document.referencesRevision!==await auditReviewReferenceRevision(references))throw new Error('Select the same reference workbooks before loading this review.');
  if(!Array.isArray(document.changes)||document.changes.length>baseline.rows.length*AUDIT_ACTION_PROPS.size)throw new Error('The review file contains too many corrections.');
  const snapshot=auditApplyCorrections(baseline,document.changes);
  const ids=value=>{if(!Array.isArray(value)||value.length>1000000||value.some(id=>typeof id!=='string'||id.length>100))throw new Error('The review file contains invalid decisions.');return new Set(value);};
  const history=Array.isArray(document.history)?document.history:[];
  if(history.length>10000||history.some(entry=>!entry||typeof entry.id!=='string'||entry.id.length>100||!Number.isFinite(Date.parse(entry.at))||!['corrected-draft','reviewed','exception'].includes(entry.disposition)||typeof entry.reason!=='string'||entry.reason.length>4000||typeof entry.owner!=='string'||entry.owner.length>200||!Array.isArray(entry.findingIds)||entry.findingIds.length>1000000||entry.findingIds.some(id=>typeof id!=='string'||id.length>100)||!Array.isArray(entry.changes)||entry.changes.some(change=>!change||!AUDIT_ACTION_FIELDS[change.field]||typeof change.before!=='string'||typeof change.value!=='string'||change.before.length>32767||change.value.length>32767)))throw new Error('The review history is invalid.');
  return {snapshot,changes:document.changes,actioned:ids(document.actioned),reviewed:ids(document.reviewed||document.actioned),excluded:ids(document.excluded),history,filterViews:auditReviewFilterViews(document.filterViews),milestoneMigration:auditReadMigrationSettings(document.milestoneMigration)};
}

export function auditRecommendationContext(snapshot,references={}){
  const index=auditActionIndex(snapshot),cohorts=new Map(),branches=new Map();
  for(const candidate of auditMilestoneCohortCandidates(snapshot.rows))cohorts.set(auditSourceKey(candidate.row),candidate);
  for(const candidate of auditMilestoneBranchCandidates(snapshot.rows))branches.set(auditSourceKey(candidate.row),candidate);
  return {snapshot,index,references,cohorts,branches};
}
function auditActionSuffix(row){return auditNormId(row.equipmentId).split('-').slice(-2).join('-');}
function auditActionTagUpn(row){
  if(auditIsBlankItemMaster(row)||auditNormId(row.discipline)==='ELECTRICAL')return '';
  const body=auditNormId(row.equipmentId);
  const candidates=new Set();
  for(const match of body.matchAll(/(?:^|[-_ ])([A-Z]{2,})(\d{3})/g)){
    for(const upn of extoRev21UpnCandidates(match[1]+match[2]))candidates.add(upn);
  }
  return candidates.size===1?[...candidates][0]:'';
}
export function auditParentRecommendations(row,context){
  const group=context.index.bySystem.get(JSON.stringify([auditNormId(row.building),auditNormId(row.upn)]))||[];
  const role=auditCommissioningRole(row),suffix=auditActionSuffix(row),current=auditNormId(row.closestParent);
  const list=[];
  for(const peer of group){
    if(peer===row||auditNormId(peer.equipmentId)===current||auditIsBlankItemMaster(peer)||context.index.byTag.get(auditNormId(peer.equipmentId))?.length!==1)continue;
    const parentRole=auditCommissioningRole(peer);
    const roleMatch=role==='drive'?parentRole==='driven-equipment':['instrument','control-valve','room-sensor','lcp'].includes(role)?['driven-equipment','panel','skid'].includes(parentRole):role==='rio'?['plc','rio'].includes(parentRole):false;
    const suffixMatch=suffix.includes('-')&&auditActionSuffix(peer)===suffix;
    if(!roleMatch||!suffixMatch)continue;
    list.push({tag:peer.equipmentId,description:peer.equipmentDescription,reason:'Same building and UPN; equipment type and tag ending agree.'});
  }
  return list.slice(0,12);
}
export function auditProposeCorrection(finding,context){
  const row=auditFindingRow(finding,context.index);if(!row)return null;
  const changes=[],add=(field,value)=>{if(clean(row[AUDIT_ACTION_FIELDS[field]])!==clean(value))changes.push(auditMakeCorrection(row,field,value,finding));};
  let reason='',confidence='Review';
  try{
    const reference=auditReferenceRecommendation(row,finding.field,context.references);
    if(reference){add(finding.field,reference.value);reason=reference.reason;confidence=reference.confidence||'Reference';}
    else if(finding.rule.id==='parent.cross-upn'){
      const parents=context.index.byTag.get(auditNormId(row.closestParent));
      if(parents?.length!==1||context.index.byTag.get(auditNormId(row.equipmentId))?.length!==1)return null;
      const parent=parents[0],role=auditCommissioningRole(row);
      const tagUpn=auditActionTagUpn(row);
      if(!tagUpn||tagUpn!==auditNormId(parent.upn))return null;
      if(!clean(row.building)||auditNormId(row.building)!==auditNormId(parent.building))return null;
      const systems=extoRev21SystemsForUpn(tagUpn),parentSystem=systems.find(value=>auditNormId(value)===auditNormId(parent.systemName));
      const currentSystem=systems.find(value=>auditNormId(value)===auditNormId(row.systemName));
      const system=currentSystem||parentSystem;if(!system)return null;
      // A UPN conflict does not by itself mean the equipment has the wrong
      // parent. Tag and parent agreement support correcting system metadata.
      add('UPN',tagUpn);add('System Name',system);
      const directInstrument=['instrument','control-valve','room-sensor'].includes(role)&&/\d+-\d+[A-Z]?$/.test(auditActionSuffix(row))&&auditActionSuffix(row)===auditActionSuffix(parent);
      const discipline=extoRev21Canonical('discipline',parent.discipline);
      if((role==='drive'||directInstrument)&&auditCommissioningRole(parent)==='driven-equipment'&&!auditIsBlankItemMaster(parent)&&discipline)add('Discipline',discipline);
      reason=`The equipment tag and ${parent.equipmentId} both identify UPN ${tagUpn}. Keep this parent and its dependencies; correct the system metadata. Confirm the equipment belongs to this system.`;
      confidence='Supported; confirm system';
    }else if(['metadata.system-upn-mismatch','metadata.upn-inconsistent'].includes(finding.rule.id)&&finding.field==='System Name'){
      const tagUpn=auditActionTagUpn(row);
      if(auditNormId(row.discipline)!=='ELECTRICAL'){
        if(!tagUpn)return null;
        if(tagUpn!==auditNormId(row.upn)){
          const systems=extoRev21SystemsForUpn(tagUpn),system=systems.find(value=>auditNormId(value)===auditNormId(row.systemName));
          const parents=context.index.byTag.get(auditNormId(row.closestParent));
          if(!system||parents&&(parents.length!==1||auditNormId(parents[0].upn)!==tagUpn))return null;
          add('UPN',tagUpn);
          reason=`The equipment tag and its assigned System Name both identify UPN ${tagUpn}. Correct the UPN and keep that System Name. Confirm the system assignment.`;
          return changes.length?{finding,row,changes,reason,confidence:'Supported; confirm system'}:null;
        }
      }
      const systems=extoRev21SystemsForUpn(row.upn);if(systems.length!==1)return null;
      add('System Name',systems[0]);
      if(auditNormId(row.closestParent)===auditNormId(row.systemName)&&!context.index.byTag.has(auditNormId(row.closestParent)))add('Closest Parent',systems[0]);
      reason='One approved System Name exists for the current UPN. Confirm the UPN before applying.';
    }else if(finding.rule.id==='metadata.ic-discipline'){
      add('Discipline',extoRev21EffectiveDiscipline(row.discipline));reason='Use the available controls discipline; confirm the assigned system.';
    }else if(['dependency.duplicate','dependency.self'].includes(finding.rule.id)){
      const seen=new Set();const refs=auditSplitReferences(row.dependencies).filter(tag=>{const key=auditNormId(tag);if(key===auditNormId(row.equipmentId)||seen.has(key))return false;seen.add(key);return true;});
      add('Dependencies',refs.join('; '));reason='Remove repeated and self-referencing entries while keeping the other dependencies.';confidence='Direct check';
    }else if(finding.rule.id==='dependency.project-not-needed'){
      add('Dependency Project','');reason='All listed dependencies resolve inside this registry. Confirm no external project is intended.';
    }else if(['milestone.local-cohort-outlier','milestone.branch-outlier'].includes(finding.rule.id)){
      const candidate=(finding.rule.id==='milestone.branch-outlier'?context.branches:context.cohorts).get(auditSourceKey(row));if(!candidate)return null;
      add('L2 Milestone',candidate.expectedMilestone);add('L1 Milestone Parent',candidate.expectedParent);
      reason=`${candidate.agreementCount} comparable records agree on this pair. Confirm this equipment is in the same phase.`;
    }else if(finding.rule.id==='item-master.standardized-assignment'){
      const candidates=auditItemMasterCanonicalCandidates(row.itemMaster,VF_ITEM_MASTER_NAMES);if(candidates.length!==1)return null;
      add('Item Master Unique Identifier',candidates[0]);reason='One standardized name has the same functional name. Confirm checklist compatibility; the site prefix alone is not an error.';
    }else if(finding.field==='Closest Parent'){
      const parents=auditParentRecommendations(row,context);if(parents.length!==1)return null;
      return auditCustomCorrection(finding,parents[0].tag,context,parents[0].reason+' Confirm the served equipment.');
    }
  }catch{return null;}
  return changes.length?{finding,row,changes,reason,confidence}:null;
}
export function auditCustomCorrection(finding,value,context,reason='Engineer-selected value.'){
  const row=auditFindingRow(finding,context.index);if(!row||!AUDIT_ACTION_FIELDS[finding.field])throw new Error('This finding needs a row-specific correction.');
  const changes=[auditMakeCorrection(row,finding.field,value,finding,reason)];
  if(finding.field==='Closest Parent'&&clean(value)!==clean(row.closestParent)&&['parent.cross-upn','parent.cross-building','logic.drive-parent-unexpected'].includes(finding.rule.id)){
    const old=context.index.byTag.get(auditNormId(row.closestParent));
    if(old?.length===1&&!auditIsBlankItemMaster(old[0])&&auditNormId(row.closestParent)!==auditNormId(row.equipmentId)){
      const refs=auditSplitReferences(row.dependencies);if(!refs.some(tag=>auditNormId(tag)===auditNormId(row.closestParent)))refs.push(row.closestParent);
      changes.push(auditMakeCorrection(row,'Dependencies',refs.join('; '),finding,'Preserve the previous equipment relationship as a dependency.'));
    }
  }
  return {finding,row,changes,reason,confidence:'Engineer review'};
}
