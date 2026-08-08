import { clean, natCmp } from '../core/text.js'
import { EXTO_REV21_COLUMNS, extoRev21Canonical, extoRev21Norm, extoRev21UpnCandidates, extoRev21ValidatePartial } from '../exto/rev21-contract.js'
import { auditFingerprint, auditNormId, auditSplitReferences } from './model.js'

export const SSM_AUDIT_STANDARD='SSM SOP + Exto Rev21';
export const SSM_AUDIT_SEVERITIES=['blocker','error','warning','info'];
export const SSM_AUDIT_CATEGORIES=['upload','structure','dependencies','metadata','milestones','item-masters','headers'];
const SSM_AUDIT_SEVERITY_RANK={blocker:0,error:1,warning:2,info:3};

function auditRule(id,category,standardRef){return {id,version:1,category,standardRef};}
const SSM_AUDIT_RULES={
  missingHeader:auditRule('upload.missing-header','upload','Exto Rev21 upload contract'),
  missingGating:auditRule('upload.missing-gating-field','upload','Exto Rev21 gating fields'),
  invalidDropdown:auditRule('metadata.invalid-dropdown-value','metadata','Exto Rev21 approved dropdowns'),
  duplicateId:auditRule('identity.duplicate-equipment-id','metadata','SSM unique equipment identity'),
  blankParent:auditRule('parent.blank','structure','SOP parent-child hierarchy'),
  selfParent:auditRule('parent.self','structure','SOP parent-child hierarchy'),
  unresolvedParent:auditRule('parent.unresolved','structure','Exto Closest Parent Status guidance'),
  generatedHeader:auditRule('parent.generated-header-review','headers','Exto New parent/header guidance'),
  crossUpn:auditRule('parent.cross-upn','structure','SOP child UPN must match parent UPN'),
  crossDiscipline:auditRule('parent.cross-discipline','structure','SOP Discipline > UPN hierarchy'),
  parentAsDependency:auditRule('dependency.repeats-parent','dependencies','SOP child does not repeat its parent as a dependency'),
  parentCycle:auditRule('parent.cycle','structure','SSM acyclic parent hierarchy'),
  literalNa:auditRule('dependency.literal-na','dependencies','Exto blank dependency semantics'),
  selfDependency:auditRule('dependency.self','dependencies','SSM predecessor logic'),
  duplicateDependency:auditRule('dependency.duplicate','dependencies','SSM predecessor logic'),
  unresolvedDependency:auditRule('dependency.unresolved','dependencies','SSM dependency references'),
  precedenceCycle:auditRule('dependency.precedence-cycle','dependencies','SOP sequential startup logic'),
  systemUpn:auditRule('metadata.system-upn-mismatch','metadata','Exto UPN and System Name contract'),
  icDiscipline:auditRule('metadata.ic-discipline','metadata','Exto approved Discipline vocabulary'),
  upnInconsistent:auditRule('metadata.upn-inconsistent','metadata','SOP Discipline > UPN/System hierarchy'),
  milestonePair:auditRule('milestone.incomplete-pair','milestones','SOP L1/L2 milestone hierarchy'),
  milestoneInconsistent:auditRule('milestone.parent-inconsistent','milestones','SOP L1/L2 milestone hierarchy'),
  legacyItemMaster:auditRule('item-master.legacy','item-masters','VF Item Master standard'),
  headerDependency:auditRule('header.has-dependency','headers','SOP organizational header semantics'),
  unusedHeader:auditRule('header.unused','headers','SOP organizational header semantics'),
};

function auditFinding(rule,severity,row,details={}){
  const equipmentId=clean(row&&row.equipmentId),source=row&&row._source||{},actual=details.actual==null?'':details.actual;
  const fingerprint=auditFingerprint([rule.id,auditNormId(equipmentId),source.sheet||'',source.row||'',JSON.stringify(actual)].join('|'));
  return Object.freeze({schemaVersion:1,id:`${rule.id}:${fingerprint}`,fingerprint,rule,severity,category:rule.category,
    equipmentId,row:source.row||0,sheet:source.sheet||'',field:details.field||'',why:details.why||'',actual,
    expected:details.expected==null?'':details.expected,recommendation:details.recommendation||'',relatedEquipmentId:details.relatedEquipmentId||'',
    searchKey:auditNormId([rule.id,rule.category,equipmentId,source.sheet,details.field,details.why,actual,details.expected,details.relatedEquipmentId].join(' '))});
}
function auditPolarity(discipline){return /elec|life safety|security|fire/i.test(discipline||'')?'top-down':'bottom-up';}
export function auditCyclePaths(nodes,edges){
  const state=new Map(),path=[],pathIndex=new Map(),cycles=[],seen=new Set();
  for(const start of nodes){
    if(state.get(start))continue;
    state.set(start,1);pathIndex.set(start,path.length);path.push(start);
    const frames=[{node:start,next:[...(edges.get(start)||[])],index:0}];
    while(frames.length){
      const frame=frames[frames.length-1];
      if(frame.index<frame.next.length){
        const next=frame.next[frame.index++],nextState=state.get(next);
        if(nextState===1){const from=pathIndex.get(next),cycle=[...path.slice(from),next],key=[...new Set(cycle)].sort(natCmp).join('|');if(!seen.has(key)){seen.add(key);cycles.push(cycle);}}
        else if(!nextState){state.set(next,1);pathIndex.set(next,path.length);path.push(next);frames.push({node:next,next:[...(edges.get(next)||[])],index:0});}
      }else{
        frames.pop();state.set(frame.node,2);pathIndex.delete(frame.node);path.pop();
      }
    }
  }
  return cycles;
}

export function runSsmAudit(snapshot){
  const rows=snapshot&&snapshot.rows||[],findings=[];let checks=0;
  const add=(rule,severity,row,details)=>findings.push(auditFinding(rule,severity,row,details));
  const missingHeaderEntries=snapshot&&snapshot.missingHeaderEntries||((snapshot&&snapshot.missingHeaders)||[]).map(header=>({header,sheet:'',headerRow:0}));
  for(const entry of missingHeaderEntries){checks++;add(SSM_AUDIT_RULES.missingHeader,'blocker',{_source:{sheet:entry.sheet,row:entry.headerRow}},{field:entry.header,why:`The registry tab does not contain the ${entry.header} column.`,expected:'Rev21 semantic header present',recommendation:'Add or correctly name the missing column before relying on this audit.'});}
  const rowsById=new Map(),duplicates=new Map();
  for(const row of rows){const key=auditNormId(row.equipmentId);if(!key)continue;if(rowsById.has(key)){const list=duplicates.get(key)||[rowsById.get(key)];list.push(row);duplicates.set(key,list);}else rowsById.set(key,row);}
  for(const list of duplicates.values())for(const row of list){checks++;add(SSM_AUDIT_RULES.duplicateId,'blocker',row,{field:'Equipment ID',why:'Equipment IDs must be unique in one SSM register.',actual:row.equipmentId,expected:'One row per Equipment ID',recommendation:'Merge duplicate records or assign the correct unique equipment tag.'});}
  const children=new Map(),parentEdges=new Map(),precedenceEdges=new Map(),generatedHeaders=new Map(),nodes=new Set(rowsById.keys());
  const addEdge=(map,from,to)=>{if(!from||!to)return;const set=map.get(from)||new Set();set.add(to);map.set(from,set);};
  const upnGroups=new Map(),milestoneGroups=new Map();
  for(const row of rows){
    const id=auditNormId(row.equipmentId),parentId=auditNormId(row.closestParent),status=extoRev21Canonical('closestParentStatus',row.closestParentStatus)||clean(row.closestParentStatus).toUpperCase();
    for(const column of EXTO_REV21_COLUMNS.filter(column=>column.gating)){checks++;if(!clean(row[column.field]))add(SSM_AUDIT_RULES.missingGating,'blocker',row,{field:column.header,why:`${column.header} is a gating Exto field and is blank.`,expected:'Populated gating value',recommendation:`Populate ${column.header} before upload.`});}
    for(const issue of extoRev21ValidatePartial(row)){checks++;add(SSM_AUDIT_RULES.invalidDropdown,issue.field==='systemName'||issue.field==='upn'||issue.field==='discipline'?'blocker':'error',row,{field:EXTO_REV21_COLUMNS.find(column=>column.field===issue.field)?.header||issue.field,why:issue.reason,actual:issue.value,expected:`A value from ${issue.dropdown}`,recommendation:'Choose the approved value that matches the project source data.'});}
    if(/^I\s*(?:&|AND)\s*C$/i.test(clean(row.discipline))){checks++;const candidates=extoRev21UpnCandidates(row.equipmentId);add(SSM_AUDIT_RULES.icDiscipline,'blocker',row,{field:'Discipline',why:'I&C is not an approved Exto Discipline. The equipment tag must identify its FMS UPN.',actual:`${row.discipline}; tag UPN candidates: ${candidates.join(', ')||'none'}`,expected:'FACILITIES MONITORING SYSTEM with one tag-derived UPN',recommendation:candidates.length===1?`Use UPN ${candidates[0]} and the approved ${candidates[0]} System Name.`:'Resolve the tag-derived UPN before assigning the Exto Discipline and System Name.'});}
    if(row.systemName&&row.upn){checks++;const prefix=extoRev21Norm(row.systemName).split(' ')[0],systemUpn=extoRev21Canonical('upn',prefix);if(systemUpn&&extoRev21Norm(systemUpn)!==extoRev21Norm(row.upn))add(SSM_AUDIT_RULES.systemUpn,'blocker',row,{field:'System Name',why:'The approved System Name identifies a different UPN than the row UPN.',actual:`UPN ${row.upn}; System ${row.systemName}`,expected:`The approved System Name for UPN ${row.upn}`,recommendation:'Correct the UPN or select the approved System Name for this equipment.'});}
    if(row.milestone&&!row.milestoneParent||!row.milestone&&row.milestoneParent){checks++;add(SSM_AUDIT_RULES.milestonePair,'warning',row,{field:'Milestone',why:'L1 Milestone Parent and L2 Milestone should be assigned as a pair.',actual:`L2: ${row.milestone||'blank'}; L1: ${row.milestoneParent||'blank'}`,expected:'Both populated or both blank under the project milestone policy',recommendation:'Confirm the project milestone program and complete the missing side of the pair.'});}
    if(row.milestone){const parents=milestoneGroups.get(auditNormId(row.milestone))||new Map();parents.set(auditNormId(row.milestoneParent),(parents.get(auditNormId(row.milestoneParent))||0)+1);milestoneGroups.set(auditNormId(row.milestone),parents);}
    if(/^CA_/i.test(clean(row.itemMaster))){checks++;add(SSM_AUDIT_RULES.legacyItemMaster,'warning',row,{field:'Item Master Unique Identifier',why:'This is a legacy CA Item Master rather than the current VF vocabulary.',actual:row.itemMaster,expected:'Approved VF Item Master',recommendation:'Review against the current Item Master Revision Template before upload.'});}
    const upn=auditNormId(row.upn);if(upn){const group=upnGroups.get(upn)||{systems:new Set(),disciplines:new Set(),wbs:new Set(),rows:[]};if(row.systemName)group.systems.add(auditNormId(row.systemName));if(row.discipline)group.disciplines.add(auditNormId(row.discipline));if(row.wbs)group.wbs.add(auditNormId(row.wbs));group.rows.push(row);upnGroups.set(upn,group);}
    checks++;if(!parentId)add(SSM_AUDIT_RULES.blankParent,'blocker',row,{field:'Closest Parent',why:'Every Exto row requires a structural parent or System Name root.',expected:'Equipment parent or the row System Name',recommendation:'Assign the correct same-UPN parent; use System Name for a root.'});
    else if(parentId===id)add(SSM_AUDIT_RULES.selfParent,'blocker',row,{field:'Closest Parent',why:'Equipment cannot be its own parent.',actual:row.closestParent,expected:'A different parent tag or System Name',recommendation:'Select the actual parent and re-check the branch.'});
    else{
      const parent=rowsById.get(parentId);
      if(parent){
        addEdge(parentEdges,id,parentId);const list=children.get(parentId)||[];list.push(row);children.set(parentId,list);
        const childUpn=auditNormId(row.upn),parentUpn=auditNormId(parent.upn);checks++;
        if(childUpn&&parentUpn&&childUpn!==parentUpn)add(SSM_AUDIT_RULES.crossUpn,'error',row,{field:'Closest Parent',why:'A structural child must share its parent’s UPN.',actual:`${row.equipmentId} (${row.upn}) → ${parent.equipmentId} (${parent.upn})`,expected:`A parent in UPN ${row.upn}`,recommendation:'Move the crossing relationship to Dependencies and select a same-UPN structural parent.',relatedEquipmentId:parent.equipmentId});
        const childDiscipline=auditNormId(row.discipline),parentDiscipline=auditNormId(parent.discipline);checks++;
        if(childDiscipline&&parentDiscipline&&childDiscipline!==parentDiscipline)add(SSM_AUDIT_RULES.crossDiscipline,'error',row,{field:'Closest Parent',why:'Structural parent-child equipment should stay inside one Discipline partition.',actual:`${row.discipline} → ${parent.discipline}`,expected:`Parent in ${row.discipline}`,recommendation:'Represent the crossing as a dependency unless an approved SOP waiver applies.',relatedEquipmentId:parent.equipmentId});
        if(auditPolarity(row.discipline)==='top-down')addEdge(precedenceEdges,parentId,id);else addEdge(precedenceEdges,id,parentId);
      }else if(status==='NEW'){
        if(parentId!==auditNormId(row.systemName)){
          const entry=generatedHeaders.get(parentId)||{row,parent:row.closestParent,count:0};entry.count++;generatedHeaders.set(parentId,entry);
        }
      }else if(status!=='EXISTING')add(SSM_AUDIT_RULES.unresolvedParent,'blocker',row,{field:'Closest Parent',why:'The parent is not in this registry and is not identified as New or Existing.',actual:row.closestParent,expected:'Resolvable equipment, New header, or Existing external parent',recommendation:'Correct the parent reference and Closest Parent Status.'});
    }
    const dependencies=auditSplitReferences(row.dependencies),seenDeps=new Set();
    for(const dependency of dependencies){const depId=auditNormId(dependency);checks++;
      if(/^N\/?A$/i.test(dependency))add(SSM_AUDIT_RULES.literalNa,'blocker',row,{field:'Dependencies',why:'Exto expects a blank cell when there is no dependency; N/A is treated as a tag.',actual:dependency,expected:'Blank dependency cell',recommendation:'Remove the literal N/A value.'});
      else if(depId===id)add(SSM_AUDIT_RULES.selfDependency,'error',row,{field:'Dependencies',why:'Equipment cannot depend on itself.',actual:dependency,expected:'A different predecessor tag',recommendation:'Remove the self-reference.'});
      if(seenDeps.has(depId))add(SSM_AUDIT_RULES.duplicateDependency,'warning',row,{field:'Dependencies',why:'The same dependency is listed more than once.',actual:dependency,expected:'One reference per dependency',recommendation:'Remove the duplicate reference.'});seenDeps.add(depId);
      if(depId&&depId===parentId)add(SSM_AUDIT_RULES.parentAsDependency,'error',row,{field:'Dependencies',why:'A child does not also list its structural parent as a dependency.',actual:dependency,expected:'Parent only in Closest Parent',recommendation:'Remove the duplicate dependency unless the hierarchy itself is being corrected.'});
      if(depId&&!rowsById.has(depId)&&!clean(row.dependencyProject)&&!/^N\/?A$/i.test(dependency))add(SSM_AUDIT_RULES.unresolvedDependency,'error',row,{field:'Dependencies',why:'The dependency tag is absent and no external Dependency Project is supplied.',actual:dependency,expected:'Resolvable tag or external Dependency Project',recommendation:'Correct the tag or identify the project containing it.'});
      if(rowsById.has(depId)&&depId!==id&&depId!==parentId)addEdge(precedenceEdges,depId,id);
    }
  }
  for(const entry of generatedHeaders.values()){checks++;add(SSM_AUDIT_RULES.generatedHeader,'warning',entry.row,{field:'Closest Parent',why:`A New parent is absent from the registry and is referenced by ${entry.count.toLocaleString()} equipment row${entry.count===1?'':'s'}.`,actual:entry.parent,expected:'An intentional approved header or the row System Name',recommendation:'Confirm this is one intentional organizational header; otherwise add or correct the missing equipment record.'});}
  for(const [upn,group] of upnGroups){checks++;if(group.systems.size>1||group.disciplines.size>1)add(SSM_AUDIT_RULES.upnInconsistent,'error',group.rows[0],{field:'UPN',why:'One UPN maps to multiple System Names or Disciplines in this registry.',actual:`Systems: ${[...group.systems].join('; ')}; Disciplines: ${[...group.disciplines].join('; ')}`,expected:'One approved System Name and Discipline per UPN',recommendation:`Review all rows assigned to UPN ${upn}.`});}
  for(const [milestone,parents] of milestoneGroups){checks++;if(parents.size>1)add(SSM_AUDIT_RULES.milestoneInconsistent,'error',rows.find(row=>auditNormId(row.milestone)===milestone),{field:'Milestone Parent',why:'The same L2 milestone maps to multiple L1 parents.',actual:[...parents.keys()].join('; '),expected:'One L1 parent per L2 milestone',recommendation:'Correct the inconsistent milestone assignment.'});}
  for(const cycle of auditCyclePaths(nodes,parentEdges)){checks++;const row=rowsById.get(cycle[0]);add(SSM_AUDIT_RULES.parentCycle,'blocker',row,{field:'Closest Parent',why:'The Closest Parent chain contains a cycle.',actual:cycle.map(id=>rowsById.get(id)?.equipmentId||id).join(' → '),expected:'Acyclic parent hierarchy',recommendation:'Break the cycle by correcting at least one Closest Parent.'});}
  for(const cycle of auditCyclePaths(nodes,precedenceEdges)){checks++;const row=rowsById.get(cycle[0]);add(SSM_AUDIT_RULES.precedenceCycle,'blocker',row,{field:'Dependencies',why:'Parent and dependency relationships create a startup sequence cycle.',actual:cycle.map(id=>rowsById.get(id)?.equipmentId||id).join(' → '),expected:'Acyclic commissioning sequence',recommendation:'Review the involved parent/dependency links and break the circular gating logic.'});}
  for(const row of rows){const headerLike=/^VF[_ -]?BLANK/i.test(clean(row.itemMaster))&&/^HEADER$/i.test(clean(row.equipmentDescription));if(!headerLike)continue;const id=auditNormId(row.equipmentId),childCount=(children.get(id)||[]).length;
    if(auditSplitReferences(row.dependencies).length){checks++;add(SSM_AUDIT_RULES.headerDependency,'warning',row,{field:'Dependencies',why:'An organizational header normally groups children and should not carry unexplained dependencies.',actual:row.dependencies,expected:'Blank unless explicitly approved',recommendation:'Confirm the header is intentional and document why it gates other equipment.'});}
    if(!childCount){checks++;add(SSM_AUDIT_RULES.unusedHeader,'warning',row,{field:'Equipment ID',why:'This header-like row has no children.',actual:row.equipmentId,expected:'At least one grouped child',recommendation:'Remove the unused header or attach the intended children.'});}
  }
  findings.sort((a,b)=>(SSM_AUDIT_SEVERITY_RANK[a.severity]-SSM_AUDIT_SEVERITY_RANK[b.severity])||natCmp(a.category,b.category)||natCmp(a.equipmentId,b.equipmentId)||a.row-b.row||natCmp(a.rule.id,b.rule.id));
  const severity=Object.fromEntries(SSM_AUDIT_SEVERITIES.map(level=>[level,findings.filter(finding=>finding.severity===level).length]));
  const category=Object.fromEntries(SSM_AUDIT_CATEGORIES.map(name=>[name,findings.filter(finding=>finding.category===name).length]));
  return Object.freeze({schemaVersion:1,standard:SSM_AUDIT_STANDARD,rows:Object.freeze([...rows]),findings:Object.freeze(findings),
    summary:Object.freeze({rows:rows.length,checks,findings:findings.length,severity:Object.freeze(severity),category:Object.freeze(category),status:severity.blocker?'blocked':severity.error||severity.warning?'review':'ready'})});
}
