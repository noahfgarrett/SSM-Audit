import { clean, natCmp } from '../core/text.js'
import { extoRev21Canonical, extoRev21Norm, extoRev21UpnCandidates } from '../exto/rev21-contract.js'
import { auditFingerprint, auditNormId, auditSplitReferences } from './model.js'

export const SSM_AUDIT_STANDARD='Registry Integrity + SSM Rules';
export const SSM_AUDIT_SEVERITIES=['blocker','error','warning','info'];
export const SSM_AUDIT_CATEGORIES=['structure','dependencies','metadata','milestones','item-masters','headers'];
export const SSM_AUDIT_SOURCES=Object.freeze([
  Object.freeze({id:'registry',label:'Registry Integrity',description:'Internal identity, references, and metadata consistency in an existing SSM.'}),
  Object.freeze({id:'sop',label:'SSM SOP',description:'Parent-child structure, dependencies, sequencing, milestones, and headers.'}),
  Object.freeze({id:'logic',label:'Commissioning Logic',description:'Confidence-rated control, electrical, and process-enabling relationships.'}),
]);
const SSM_AUDIT_SEVERITY_RANK={blocker:0,error:1,warning:2,info:3};

function auditRule(id,source,category,title,statement,options={}){return Object.freeze({id,version:1,source,category,title,statement,standardRef:title,confidence:options.confidence||'required',enabled:options.enabled!==false,disabledReason:options.disabledReason||''});}
export const SSM_AUDIT_RULES=Object.freeze({
  duplicateId:auditRule('identity.duplicate-equipment-id','sop','metadata','Unique equipment identity','Each Equipment ID must occur once in an SSM.'),
  blankParent:auditRule('parent.blank','sop','structure','Complete hierarchy','Each equipment row must have a parent or System Name root.'),
  selfParent:auditRule('parent.self','sop','structure','No self-parenting','Equipment cannot be its own parent.'),
  unresolvedParent:auditRule('parent.unresolved','registry','structure','Resolvable parent','A parent must resolve to equipment, an intentional new header, or an existing external parent.'),
  generatedHeader:auditRule('parent.generated-header-review','registry','headers','Intentional new header','A proposed organizational header must be intentional and consistently reused.'),
  crossUpn:auditRule('parent.cross-upn','sop','structure','Parent and child share a UPN','Structural parent-child equipment must remain within one UPN.'),
  crossDiscipline:auditRule('parent.cross-discipline','sop','structure','Parent and child share a discipline','Cross-discipline relationships belong in Dependencies unless an approved exception applies.'),
  parentCycle:auditRule('parent.cycle','sop','structure','Acyclic hierarchy','A parent chain cannot loop back to itself.'),
  selfDependency:auditRule('dependency.self','sop','dependencies','No self-dependency','Equipment cannot depend on itself.'),
  duplicateDependency:auditRule('dependency.duplicate','sop','dependencies','Unique dependency references','List each dependency once per equipment row.'),
  unresolvedDependency:auditRule('dependency.unresolved','sop','dependencies','Resolvable dependency','A dependency must resolve locally or identify its external project.'),
  precedenceCycle:auditRule('dependency.precedence-cycle','sop','dependencies','Acyclic commissioning sequence','Parent and dependency links cannot create a circular startup sequence.'),
  systemUpn:auditRule('metadata.system-upn-mismatch','registry','metadata','System Name matches UPN','System Name and UPN must identify the same system.'),
  icDiscipline:auditRule('metadata.ic-discipline','registry','metadata','Approved controls mapping','I&C equipment must map to the approved controls discipline and tag-derived UPN.'),
  upnInconsistent:auditRule('metadata.upn-inconsistent','sop','metadata','Consistent UPN grouping','One UPN should map to one System Name and Discipline.'),
  milestonePair:auditRule('milestone.incomplete-pair','sop','milestones','Complete milestone pair','When a project uses milestones, every equipment row must include both L1 Milestone Parent and L2 Milestone.'),
  milestoneIntent:auditRule('milestone.intent-mismatch','sop','milestones','Aligned milestone intent','L1 and L2 milestone intent must not conflict on an explicit phase, capacity, or percentage target.'),
  milestoneInconsistent:auditRule('milestone.parent-inconsistent','sop','milestones','Consistent milestone meaning','One L2 milestone should map to one consistent L1 milestone intent.'),
  itemMasterStandard:auditRule('item-master.standardized-assignment','registry','item-masters','Standardized Item Master','Item Master assignments should match the approved VF standard by equipment meaning, not by project prefix.',{enabled:false,confidence:'pending',disabledReason:'Waiting for the approved VF Item Master standard.'}),
  headerDependency:auditRule('header.has-dependency','sop','headers','Header groups only','An organizational header should not carry unexplained dependencies.'),
  unusedHeader:auditRule('header.unused','sop','headers','Header has children','An organizational header must group at least one child.'),
  controlLink:auditRule('logic.control-link-missing','logic','dependencies','Control source present','A control device with a repeatable relationship should identify its RIO, PLC, VFD, or control-panel source.',{confidence:'description-rated'}),
  drivenElectricalPath:auditRule('logic.driven-electrical-path-missing','logic','dependencies','Electrical enabling path present','Electrically driven equipment should trace to an electrical predecessor.',{confidence:'strong'}),
  controlElectricalPath:auditRule('logic.control-electrical-path-missing','logic','dependencies','Control power path present','Control equipment should trace to an electrical supply predecessor.',{confidence:'strong'}),
  rioControlPath:auditRule('logic.rio-control-path-missing','logic','dependencies','RIO controller path present','A RIO should link to an I/O cluster, PLC, upstream RIO, or equivalent controller.',{confidence:'strong'}),
  driveParent:auditRule('logic.drive-parent-unexpected','logic','structure','Drive grouped with driven equipment','A VFD or motor starter should be structurally grouped with the equipment it operates.',{confidence:'strong'}),
  heatTraceChain:auditRule('logic.heat-trace-chain-missing','logic','structure','Heat-trace chain present','Heat-trace panels and connection boxes should follow their transformer, panel, or upstream connection-box chain.',{confidence:'strong'}),
  fduDependency:auditRule('logic.fdu-supported-equipment-missing','logic','dependencies','Fiber unit identifies served equipment','A fiber distribution unit should identify the equipment or system it enables.',{confidence:'strong'}),
  vesdaFireAlarm:auditRule('logic.vesda-fire-alarm-missing','logic','dependencies','VESDA fire-alarm dependency present','A VESDA system should identify its associated fire alarm panel as a dependency.',{confidence:'required'}),
});

function auditFinding(rule,severity,row,details={}){
  const equipmentId=clean(row&&row.equipmentId),source=row&&row._source||{},actual=details.actual==null?'':details.actual;
  const fingerprint=auditFingerprint([rule.id,auditNormId(equipmentId),source.sheet||'',source.row||'',JSON.stringify(actual)].join('|'));
  return Object.freeze({schemaVersion:1,id:`${rule.id}:${fingerprint}`,fingerprint,rule,severity,category:rule.category,
    equipmentId,row:source.row||0,sheet:source.sheet||'',field:details.field||'',why:details.why||'',actual,
    expected:details.expected==null?'':details.expected,recommendation:details.recommendation||'',relatedEquipmentId:details.relatedEquipmentId||'',
    searchKey:auditNormId([rule.id,rule.source,rule.title,rule.category,equipmentId,source.sheet,details.field,details.why,actual,details.expected,details.relatedEquipmentId].join(' '))});
}
function auditPolarity(discipline){return /elec|life safety|security|fire/i.test(discipline||'')?'top-down':'bottom-up';}
function auditDescription(row){return auditNormId(row&&row.equipmentDescription);}
function auditReferences(value){return auditSplitReferences(value).filter(reference=>!/^N\/?A$/i.test(reference));}
function auditIsOrganizationalHeader(row){return /^HEADER$/i.test(clean(row&&row.equipmentDescription));}
function auditIsElectrical(row){return auditNormId(row&&row.discipline)==='ELECTRICAL';}
function auditIsRio(row){return !auditIsOrganizationalHeader(row)&&/REMOTE\s*I\/?O|\bRIO\b/.test(auditDescription(row));}
function auditIsDriveOrStarter(row){return !auditIsOrganizationalHeader(row)&&/^VARIABLE FREQUENCY DRIVE\b|\bVFD SKID\b|\bMOTOR STARTER\b/.test(auditDescription(row));}
function auditIsController(row){return !auditIsOrganizationalHeader(row)&&(/I\/?O CLUSTER|IO CLUSTER|REMOTE\s*I\/?O|\bRIO\b|\bPLC\b|PROGRAMMABLE LOGIC CONTROLLER|CONTROL PANEL|FIRE ALARM PANEL|GAS DETECTION PANEL|SECURITY CONTROL PANEL/.test(auditDescription(row))||auditIsDriveOrStarter(row));}
function auditRequiresRioController(row){return auditIsRio(row)&&!/\bPLC\b|PROGRAMMABLE LOGIC CONTROLLER/.test(auditDescription(row));}
function auditIsControlEquipment(row){return auditIsController(row)||/FMS (?:CABINET|CONTROLLER)/.test(auditDescription(row));}
function auditIsDrivenEquipment(row){const description=auditDescription(row);return !auditIsOrganizationalHeader(row)&&!auditIsElectrical(row)&&!/\bPLC\b|PROGRAMMABLE LOGIC CONTROLLER|CONTROL PANEL|\bVFD SKID\b|^VARIABLE FREQUENCY DRIVE\b|\bMOTOR STARTER\b/.test(description)&&/AIR HANDLER|\bMAH\b|\bAHU\b|PUMP|\bFAN\b|\bCHILLER\b|\bCOMPRESSOR\b|\bBLOWER\b|\bSCRUBBER\b|\bBOILER\b|COOLING TOWER|\bMIXER\b/.test(description);}
function auditIsTransformer(row){return /TRANSFORMER/.test(auditDescription(row));}
function auditIsHeatTracePanel(row){return /HEAT TRACE PANEL|HEAT TRACING PANEL/.test(auditDescription(row));}
function auditIsHeatTraceConnection(row){return /HEAT TRACE POWER CONNECTION/.test(auditDescription(row));}
function auditIsFdu(row){return /FIBER OPTIC DISTRIBUTION UNIT|\bFDU\b/.test(auditDescription(row));}
function auditIsVesda(row){return /\bVESDA\b|VERY EARLY SMOKE/.test(auditDescription(row));}
function auditIsFireAlarmPanel(row){return /FIRE ALARM PANEL/.test(auditDescription(row));}
function auditControlExpectation(row){
  if(auditIsOrganizationalHeader(row)||auditIsElectrical(row))return '';
  const description=auditDescription(row);
  if(/VOLTAGE TRANSMITTER|PRESSURE SWITCH HIGH|LEVEL SWITCH LOW|HUMIDITY CONTROL VALVE|ANALYZER TRANSMITTER[- ]MOISTURE|LEVEL TRANSMITTER - GWR/.test(description))return 'strong';
  if(/POWER SUPPLY ALARM|CURRENT SWITCH|HAND SWITCH.*(?:POSITION|HOA)|POSITION INDICATING TRANSMITTER|TEMPERATURE TRANSMITTER.*LINE|CONTROL VALVE.*GLOBE|HUMIDITY ELEMENT TRANSMITTER|FLOW SWITCH|FLOW TRANSMITTER|TEMPERATURE TRANSMITTER|STATUS INDICATOR/.test(description))return 'advisory';
  return '';
}
function auditMilestoneTokens(value){
  const text=auditNormId(value).replace(/^L[12]-[^ ]+\s*-?\s*/,'');if(!text)return [];
  const tokens=[];let match;
  if((match=text.match(/\b(25|30|50|100)\s*(?:PERCENT|%)/)))tokens.push(`percent-${match[1]}`);
  if(/END[- ]STATE|\bCAPACITY\b/.test(text))tokens.push('capacity');
  if(/\bENABLING\b/.test(text))tokens.push('enabling');
  if(/ENERGIZ/.test(text))tokens.push('energization');
  if(/BLOW\s*DOWN|\bBD START\b/.test(text))tokens.push('blowdown');
  if(/OP(?:ERATIONAL)? READY/.test(text))tokens.push('operational-ready');
  if(/ROOMS? .*READY|ROOM READY/.test(text))tokens.push('room-ready');
  if(/EQUIPMENT SET/.test(text))tokens.push('equipment-set');
  if(/DATA COLLECTION/.test(text))tokens.push('data-collection');
  return [...new Set(tokens)];
}
function auditMilestoneIntentKey(value){const tokens=auditMilestoneTokens(value);return tokens.length?tokens.join('|'):'unclassified';}
function auditMilestoneConflict(l1,l2){
  const left=auditMilestoneTokens(l1),right=auditMilestoneTokens(l2),leftPercent=left.find(token=>token.startsWith('percent-')),rightPercent=right.find(token=>token.startsWith('percent-'));
  if(leftPercent&&rightPercent&&leftPercent!==rightPercent)return `${leftPercent.replace('percent-','')}% L1 versus ${rightPercent.replace('percent-','')}% L2`;
  if(left.includes('capacity')&&right.includes('enabling')||left.includes('enabling')&&right.includes('capacity'))return 'capacity versus enabling';
  return '';
}
function auditReachableFrom(seeds,edges){
  const reached=new Set(seeds),queue=[...seeds];
  for(let index=0;index<queue.length;index++)for(const next of edges.get(queue[index])||[])if(!reached.has(next)){reached.add(next);queue.push(next);}
  return reached;
}
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
  const rowsById=new Map(),duplicates=new Map();
  for(const row of rows){const key=auditNormId(row.equipmentId);if(!key)continue;if(rowsById.has(key)){const list=duplicates.get(key)||[rowsById.get(key)];list.push(row);duplicates.set(key,list);}else rowsById.set(key,row);}
  for(const list of duplicates.values())for(const row of list){checks++;add(SSM_AUDIT_RULES.duplicateId,'blocker',row,{field:'Equipment ID',why:'Equipment IDs must be unique in one SSM register.',actual:row.equipmentId,expected:'One row per Equipment ID',recommendation:'Merge duplicate records or assign the correct unique equipment tag.'});}
  const children=new Map(),parentEdges=new Map(),precedenceEdges=new Map(),generatedHeaders=new Map(),nodes=new Set(rowsById.keys());
  const addEdge=(map,from,to)=>{if(!from||!to)return;const set=map.get(from)||new Set();set.add(to);map.set(from,set);};
  const upnGroups=new Map(),milestoneGroups=new Map(),projectUsesMilestones=rows.some(row=>clean(row.milestone)||clean(row.milestoneParent));
  for(const row of rows){
    const id=auditNormId(row.equipmentId),parentId=auditNormId(row.closestParent),status=extoRev21Canonical('closestParentStatus',row.closestParentStatus)||clean(row.closestParentStatus).toUpperCase();
    if(/^I\s*(?:&|AND)\s*C$/i.test(clean(row.discipline))){checks++;const candidates=extoRev21UpnCandidates(row.equipmentId);add(SSM_AUDIT_RULES.icDiscipline,'blocker',row,{field:'Discipline',why:'I&C does not match the allowed Cx discipline mapping.',actual:`${row.discipline}; tag UPN candidates: ${candidates.join(', ')||'none'}`,expected:'FACILITIES MONITORING SYSTEM with one tag-derived UPN',recommendation:candidates.length===1?`Use UPN ${candidates[0]} and its allowed System Name.`:'Resolve the tag-derived UPN before assigning Discipline and System Name.'});}
    if(row.systemName&&row.upn){checks++;const match=extoRev21Norm(row.systemName).match(/^([0-9]+)(?:\s|$)/),systemUpn=match&&match[1];if(systemUpn&&extoRev21Norm(systemUpn)!==extoRev21Norm(row.upn))add(SSM_AUDIT_RULES.systemUpn,'blocker',row,{field:'System Name',why:'System Name identifies a different UPN than the row assignment.',actual:`UPN ${row.upn}; System ${row.systemName}`,expected:`System Name beginning with UPN ${row.upn}`,recommendation:'Correct the UPN or select the System Name assigned to that UPN.'});}
    if(projectUsesMilestones){checks++;if(!clean(row.milestone)||!clean(row.milestoneParent))add(SSM_AUDIT_RULES.milestonePair,'error',row,{field:'Milestones',why:'This project uses milestones, but this equipment does not have a complete L1/L2 pair.',actual:`L1: ${row.milestoneParent||'blank'}; L2: ${row.milestone||'blank'}`,expected:'Both L1 Milestone Parent and L2 Milestone populated',recommendation:'Assign the equipment to its governing L1 milestone and corresponding L2 milestone.'});}
    if(row.milestone&&row.milestoneParent){
      const conflict=auditMilestoneConflict(row.milestoneParent,row.milestone);checks++;
      if(conflict)add(SSM_AUDIT_RULES.milestoneIntent,'warning',row,{field:'Milestones',why:`The milestone pair has conflicting intent: ${conflict}.`,actual:`L1: ${row.milestoneParent}; L2: ${row.milestone}`,expected:'L1 and L2 milestones with compatible phase and capacity intent',recommendation:'Confirm the equipment belongs to the same startup phase represented by both milestones.'});
      const key=auditNormId(row.milestone),entry=milestoneGroups.get(key)||{row,parents:new Map()};
      const intent=auditMilestoneIntentKey(row.milestoneParent),values=entry.parents.get(intent)||new Set();values.add(clean(row.milestoneParent));entry.parents.set(intent,values);milestoneGroups.set(key,entry);
    }
    const upn=auditNormId(row.upn);if(upn){const group=upnGroups.get(upn)||{systems:new Set(),disciplines:new Set(),wbs:new Set(),rows:[]};if(row.systemName)group.systems.add(auditNormId(row.systemName));if(row.discipline)group.disciplines.add(auditNormId(row.discipline));if(row.wbs)group.wbs.add(auditNormId(row.wbs));group.rows.push(row);upnGroups.set(upn,group);}
    checks++;if(!parentId)add(SSM_AUDIT_RULES.blankParent,'blocker',row,{field:'Closest Parent',why:'This equipment has no structural parent or System Name root.',expected:'Equipment parent or the row System Name',recommendation:'Assign the correct same-UPN parent; use System Name for a root.'});
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
      }else if(status!=='EXISTING')add(SSM_AUDIT_RULES.unresolvedParent,'blocker',row,{field:'Closest Parent',why:'Closest Parent cannot be resolved for Cx upload.',actual:row.closestParent,expected:'Resolvable equipment, New header, or Existing external parent',recommendation:'Correct the parent reference and Closest Parent Status.'});
    }
    const dependencies=auditReferences(row.dependencies),seenDeps=new Set();
    for(const dependency of dependencies){const depId=auditNormId(dependency);checks++;
      if(depId===id)add(SSM_AUDIT_RULES.selfDependency,'error',row,{field:'Dependencies',why:'Equipment cannot depend on itself.',actual:dependency,expected:'A different predecessor tag',recommendation:'Remove the self-reference.'});
      if(seenDeps.has(depId))add(SSM_AUDIT_RULES.duplicateDependency,'warning',row,{field:'Dependencies',why:'The same dependency is listed more than once.',actual:dependency,expected:'One reference per dependency',recommendation:'Remove the duplicate reference.'});seenDeps.add(depId);
      if(depId&&!rowsById.has(depId)&&!clean(row.dependencyProject))add(SSM_AUDIT_RULES.unresolvedDependency,'error',row,{field:'Dependencies',why:'The dependency tag is absent and no external Dependency Project is supplied.',actual:dependency,expected:'Resolvable tag or external Dependency Project',recommendation:'Correct the tag or identify the project containing it.'});
      if(rowsById.has(depId)&&depId!==id&&depId!==parentId)addEdge(precedenceEdges,depId,id);
    }
  }
  for(const entry of generatedHeaders.values()){checks++;add(SSM_AUDIT_RULES.generatedHeader,'warning',entry.row,{field:'Closest Parent',why:`A proposed new header is referenced by ${entry.count.toLocaleString()} equipment row${entry.count===1?'':'s'}.`,actual:entry.parent,expected:'An intentional upload header or the row System Name',recommendation:'Confirm this is one intentional organizational header; otherwise add or correct the missing equipment record.'});}
  for(const [upn,group] of upnGroups){checks++;if(group.systems.size>1||group.disciplines.size>1)add(SSM_AUDIT_RULES.upnInconsistent,'error',group.rows[0],{field:'UPN',why:'One UPN maps to multiple System Names or Disciplines in this registry.',actual:`Systems: ${[...group.systems].join('; ')}; Disciplines: ${[...group.disciplines].join('; ')}`,expected:'One approved System Name and Discipline per UPN',recommendation:`Review all rows assigned to UPN ${upn}.`});}
  for(const entry of milestoneGroups.values()){
    const meaningful=[...entry.parents].filter(([intent])=>intent!=='unclassified');checks++;
    if(meaningful.length>1)add(SSM_AUDIT_RULES.milestoneInconsistent,'error',entry.row,{field:'Milestone Parent',why:'The same L2 milestone maps to conflicting L1 milestone meanings.',actual:meaningful.map(([intent])=>intent).join('; '),expected:'One consistent L1 phase and readiness intent per L2 milestone',recommendation:'Review the affected milestone assignments by meaning rather than milestone number.'});
  }
  for(const cycle of auditCyclePaths(nodes,parentEdges)){checks++;const row=rowsById.get(cycle[0]);add(SSM_AUDIT_RULES.parentCycle,'blocker',row,{field:'Closest Parent',why:'The Closest Parent chain contains a cycle.',actual:cycle.map(id=>rowsById.get(id)?.equipmentId||id).join(' → '),expected:'Acyclic parent hierarchy',recommendation:'Break the cycle by correcting at least one Closest Parent.'});}
  for(const cycle of auditCyclePaths(nodes,precedenceEdges)){checks++;const row=rowsById.get(cycle[0]);add(SSM_AUDIT_RULES.precedenceCycle,'blocker',row,{field:'Dependencies',why:'Parent and dependency relationships create a startup sequence cycle.',actual:cycle.map(id=>rowsById.get(id)?.equipmentId||id).join(' → '),expected:'Acyclic commissioning sequence',recommendation:'Review the involved parent/dependency links and break the circular gating logic.'});}
  for(const row of rows){if(!auditIsOrganizationalHeader(row))continue;const id=auditNormId(row.equipmentId),childCount=(children.get(id)||[]).length;
    if(auditReferences(row.dependencies).length){checks++;add(SSM_AUDIT_RULES.headerDependency,'warning',row,{field:'Dependencies',why:'An organizational header normally groups children and should not carry unexplained dependencies.',actual:row.dependencies,expected:'Blank unless explicitly approved',recommendation:'Confirm the header is intentional and document why it gates other equipment.'});}
    if(!childCount){checks++;add(SSM_AUDIT_RULES.unusedHeader,'warning',row,{field:'Equipment ID',why:'This header-like row has no children.',actual:row.equipmentId,expected:'At least one grouped child',recommendation:'Remove the unused header or attach the intended children.'});}
  }
  const electricalSeeds=new Set(rows.filter(auditIsElectrical).map(row=>auditNormId(row.equipmentId)).filter(Boolean));
  const electricalPath=auditReachableFrom(electricalSeeds,precedenceEdges);
  for(const row of rows){
    const id=auditNormId(row.equipmentId);if(!id)continue;
    const parent=rowsById.get(auditNormId(row.closestParent)),dependencies=auditReferences(row.dependencies).map(value=>rowsById.get(auditNormId(value))).filter(Boolean),related=[parent,...dependencies].filter(Boolean);
    const controlExpectation=auditControlExpectation(row);
    if(controlExpectation){checks++;if(!related.some(auditIsController))add(SSM_AUDIT_RULES.controlLink,controlExpectation==='strong'?'warning':'info',row,{field:'Dependencies',why:'This control device has no traceable RIO, PLC, VFD, or control-panel source.',actual:row.dependencies||'Blank',expected:'A control source linked as a parent or dependency',recommendation:'Confirm the control source and add the missing relationship.'});}
    if(auditIsDrivenEquipment(row)){checks++;if(!electricalPath.has(id))add(SSM_AUDIT_RULES.drivenElectricalPath,'warning',row,{field:'Dependencies',why:'No electrical enabling path was found for this driven equipment.',actual:row.dependencies||'Blank',expected:'A commissioning path to its electrical supply equipment',recommendation:'Trace the supplying panel, starter, or VFD and add the missing relationship.'});}
    if(!auditIsElectrical(row)&&auditIsControlEquipment(row)){checks++;if(!electricalPath.has(id))add(SSM_AUDIT_RULES.controlElectricalPath,auditIsRio(row)?'warning':'info',row,{field:'Dependencies',why:'This control equipment has no traceable electrical supply.',actual:row.dependencies||'Blank',expected:'A commissioning path to its supplying panel, circuit, or transformer',recommendation:'Confirm the supplying electrical equipment and add the missing relationship.'});}
    if(auditRequiresRioController(row)){checks++;if(!related.some(auditIsController))add(SSM_AUDIT_RULES.rioControlPath,'warning',row,{field:'Dependencies',why:'This RIO has no traceable controller or upstream I/O relationship.',actual:row.dependencies||'Blank',expected:'I/O cluster, PLC, upstream RIO, or equivalent controller',recommendation:'Add the controlling or upstream I/O relationship.'});}
    if(auditIsDriveOrStarter(row)){checks++;if(!parent||!auditIsDrivenEquipment(parent))add(SSM_AUDIT_RULES.driveParent,'warning',row,{field:'Closest Parent',why:'This drive or starter is not grouped with driven equipment.',actual:parent?`${row.closestParent} (${parent.equipmentDescription||'description unavailable'})`:(row.closestParent||'Blank'),expected:'Pump, fan, air handler, chiller, compressor, scrubber, boiler, cooling tower, or mixer',recommendation:'Move the drive beneath the equipment it operates, then preserve its electrical feed as a dependency.'});}
    if(auditIsHeatTracePanel(row)){checks++;if(!parent||!auditIsTransformer(parent))add(SSM_AUDIT_RULES.heatTraceChain,'warning',row,{field:'Closest Parent',why:'This heat-trace panel does not follow a transformer in the structural chain.',actual:parent?`${row.closestParent} (${parent.equipmentDescription||'description unavailable'})`:(row.closestParent||'Blank'),expected:'Supplying transformer',recommendation:'Confirm and assign the transformer that supplies this panel.'});}
    if(auditIsHeatTraceConnection(row)){checks++;if(!parent||(!auditIsHeatTracePanel(parent)&&!auditIsHeatTraceConnection(parent)))add(SSM_AUDIT_RULES.heatTraceChain,'warning',row,{field:'Closest Parent',why:'This heat-trace connection box is outside its panel or connection-box chain.',actual:parent?`${row.closestParent} (${parent.equipmentDescription||'description unavailable'})`:(row.closestParent||'Blank'),expected:'Heat-trace panel or upstream heat-trace connection box',recommendation:'Attach the connection box to the correct heat-trace branch.'});}
    if(auditIsFdu(row)){checks++;if(!dependencies.length)add(SSM_AUDIT_RULES.fduDependency,'warning',row,{field:'Dependencies',why:'This fiber distribution unit does not identify the equipment or system it enables.',actual:'Blank',expected:'At least one served-equipment dependency',recommendation:'Add the supported equipment or system as a dependency.'});}
    if(auditIsVesda(row)){checks++;if(!dependencies.some(auditIsFireAlarmPanel))add(SSM_AUDIT_RULES.vesdaFireAlarm,'error',row,{field:'Dependencies',why:'This VESDA system does not identify its associated fire alarm panel.',actual:row.dependencies||'Blank',expected:'Associated fire alarm panel dependency',recommendation:'Add the fire alarm panel that handles the VESDA alarms and actions.'});}
  }
  findings.sort((a,b)=>(SSM_AUDIT_SEVERITY_RANK[a.severity]-SSM_AUDIT_SEVERITY_RANK[b.severity])||natCmp(a.category,b.category)||natCmp(a.equipmentId,b.equipmentId)||a.row-b.row||natCmp(a.rule.id,b.rule.id));
  const severity=Object.fromEntries(SSM_AUDIT_SEVERITIES.map(level=>[level,findings.filter(finding=>finding.severity===level).length]));
  const category=Object.fromEntries(SSM_AUDIT_CATEGORIES.map(name=>[name,findings.filter(finding=>finding.category===name).length]));
  const source=Object.fromEntries(SSM_AUDIT_SOURCES.map(item=>[item.id,findings.filter(finding=>finding.rule.source===item.id).length]));
  return Object.freeze({schemaVersion:1,standard:SSM_AUDIT_STANDARD,rows:Object.freeze([...rows]),findings:Object.freeze(findings),
    summary:Object.freeze({rows:rows.length,checks,findings:findings.length,severity:Object.freeze(severity),category:Object.freeze(category),source:Object.freeze(source),status:severity.blocker?'blocked':severity.error||severity.warning?'review':'ready'})});
}
