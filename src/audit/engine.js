import { clean, natCmp } from '../core/text.js'
import { extoRev21Canonical, extoRev21EffectiveDiscipline, extoRev21IsSystemName, extoRev21IsUpn, extoRev21Norm, extoRev21SystemsForUpn, extoRev21UpnCandidates } from '../exto/rev21-contract.js'
import { VF_ITEM_MASTER_NAMES } from '../exto/vf-item-masters.js'
import { auditFingerprint, auditNormId, auditSplitReferences } from './model.js'

/* ---- SSM Audit rule engine ----
   Every rule carries a plain-language `statement` (what must be true) and every
   finding carries a plain-language `why` (what was seen). Both are written for
   an engineer picking the report up cold: one or two short sentences, the
   guideline it comes from, no implementation talk. Rule ids are stable
   fingerprint inputs — never rename an id casually. */

export const SSM_AUDIT_STANDARD='Registry Integrity + SSM Rules';
export const SSM_AUDIT_SEVERITIES=['blocker','error','warning','info'];
export const SSM_AUDIT_CATEGORIES=['structure','dependencies','metadata','milestones','item-masters','headers'];
export const SSM_AUDIT_SOURCES=Object.freeze([
  Object.freeze({id:'registry',label:'Registry Integrity',description:'The registry has to agree with itself and with the approved Rev21 lists — the same lists Exto validates against.'}),
  Object.freeze({id:'sop',label:'SSM SOP',description:'How the SSM SOP says equipment nests, depends, sequences, and rolls up to milestones.'}),
  Object.freeze({id:'logic',label:'Commissioning Logic',description:'Relationships that commissioning practice expects — a drive under the equipment it runs, a VESDA tied to its fire alarm panel.'}),
]);
const SSM_AUDIT_SEVERITY_RANK={blocker:0,error:1,warning:2,info:3};

function auditRule(id,source,category,title,statement,options={}){return Object.freeze({id,version:1,source,category,title,statement,standardRef:title,confidence:options.confidence||'required',enabled:options.enabled!==false,disabledReason:options.disabledReason||''});}
export const SSM_AUDIT_RULES=Object.freeze({
  /* identity + structure */
  duplicateId:auditRule('identity.duplicate-equipment-id','sop','metadata','One row per Equipment ID','Each Equipment ID appears once. Two rows with the same tag cannot both be right.'),
  blankParent:auditRule('parent.blank','sop','structure','Every row has a parent','Each equipment row names a Closest Parent — another piece of equipment, or its own System Name when it is the top of that system.'),
  selfParent:auditRule('parent.self','sop','structure','No self-parenting','A row cannot name itself as its Closest Parent.'),
  unresolvedParent:auditRule('parent.unresolved','registry','structure','Parent can be found','A Closest Parent must be a tag in this registry, an intentional header, or an Existing parent from another project.'),
  generatedHeader:auditRule('parent.generated-header-review','registry','headers','New header is intentional','A Closest Parent that is not any equipment row is a proposed header. Confirm it is meant to be one and is reused consistently.'),
  crossUpn:auditRule('parent.cross-upn','sop','structure','Parent shares the UPN','A structural child stays inside its parent’s UPN. Anything that crosses UPNs is a dependency of the downstream equipment, not its parent.'),
  crossDiscipline:auditRule('parent.cross-discipline','sop','structure','Parent shares the discipline','A structural child stays inside its parent’s discipline. Controls devices nesting under the equipment they serve are the approved exception.'),
  parentCycle:auditRule('parent.cycle','sop','structure','No loops in the hierarchy','Following Closest Parent upward must never lead back to the same equipment.'),
  /* dependencies */
  selfDependency:auditRule('dependency.self','sop','dependencies','No self-dependency','A row cannot list itself as a dependency.'),
  duplicateDependency:auditRule('dependency.duplicate','sop','dependencies','Each dependency listed once','The same dependency should not appear twice on one row.'),
  unresolvedDependency:auditRule('dependency.unresolved','sop','dependencies','Dependency can be found','A dependency must be a tag in this registry, or name the other project it lives in.'),
  precedenceCycle:auditRule('dependency.precedence-cycle','sop','dependencies','Startup order has no loops','Parents and dependencies together define what starts before what. That order cannot go in a circle.'),
  sameUpnBottomUp:auditRule('dependency.same-upn-bottom-up','sop','dependencies','Same-system dependency is deliberate','In bottom-up disciplines (mechanical, process, waste, and similar) equipment inside one UPN is sequenced by the hierarchy itself. A same-UPN dependency there is normal for instruments and devices; on other equipment it is worth a second look.',{confidence:'strong'}),
  /* metadata */
  upnNotApproved:auditRule('metadata.upn-not-approved','registry','metadata','UPN is on the approved list','The UPN must be one of the values in the Rev21 upload template. Three-digit numbers are the norm; RR, SEC, and MISC are the approved letter codes.'),
  systemUpn:auditRule('metadata.system-upn-mismatch','registry','metadata','System Name belongs to the UPN','System Name must be one of the approved Rev21 System Names for the row’s UPN. Numeric UPNs normally lead the name (602 Medium Voltage); letter codes such as RR map to their own approved names.'),
  icDiscipline:auditRule('metadata.ic-discipline','registry','metadata','I&C uses the approved discipline','Instrumentation and controls rows use the approved controls discipline, and their UPN comes from the tag.'),
  upnInconsistent:auditRule('metadata.upn-inconsistent','sop','metadata','One UPN, one system','Every row on a UPN should carry the same System Name and Discipline.'),
  /* milestones — enabled as review-grade checks; engineers decide */
  milestonePair:auditRule('milestone.incomplete-pair','sop','milestones','L1 and L2 assigned together','When a project uses milestones, each row carries both an L1 Milestone Parent and an L2 Milestone.',{confidence:'strong'}),
  milestoneUpn:auditRule('milestone.l2-upn-mismatch','sop','milestones','L2 milestone names this UPN','The L2 milestone name should contain the row’s own UPN. When it names a different UPN the assignment needs review.',{confidence:'strong'}),
  milestoneIntent:auditRule('milestone.intent-mismatch','sop','milestones','L1 and L2 agree','L1 and L2 milestones on one row should describe the same phase — not, for example, 30% capacity on one and 100% on the other.',{confidence:'strong'}),
  milestoneInconsistent:auditRule('milestone.parent-inconsistent','sop','milestones','One L2, one meaning','The same L2 milestone should roll up to one consistent L1 meaning across the registry.',{confidence:'strong'}),
  milestoneCohort:auditRule('milestone.local-cohort-outlier','sop','milestones','Matches comparable equipment','Equipment of the same kind, in the same building and system, normally shares one milestone pair.',{confidence:'strong'}),
  milestoneLevel:auditRule('milestone.level-field-mismatch','sop','milestones','L1 and L2 in the right fields','An identifier written for L1 belongs in L1 Milestone Parent; one written for L2 belongs in L2 Milestone.',{confidence:'strong'}),
  milestoneBranch:auditRule('milestone.branch-outlier','sop','milestones','Matches its branch','A child normally shares the milestone pair of its parent and siblings unless there is a documented reason.',{confidence:'strong'}),
  /* item masters */
  itemMasterStandard:auditRule('item-master.standardized-assignment','registry','item-masters','Item Master is the VF standard','Item Masters use the approved VF names. Legacy project-prefixed names (CA_…) should be replaced with their VF equivalent.',{confidence:'strong'}),
  /* headers */
  headerItemMaster:auditRule('header.item-master-not-blank','sop','headers','Header uses a Blank Item Master','A row that other equipment nests under as an organizational header should carry a Blank Item Master (VF_Blank) so no checklists are applied to it.'),
  headerDependency:auditRule('header.has-dependency','sop','headers','Header carries no dependencies','An organizational header only groups equipment. It should not carry dependencies of its own.'),
  unusedHeader:auditRule('header.unused','sop','headers','Header has children','A row set up as a header (Blank Item Master) should have at least one piece of equipment nested under it.'),
  /* commissioning logic */
  controlLink:auditRule('logic.control-link-missing','logic','dependencies','Control device names its source','A control device should point to what controls it — an RIO, PLC, VFD, or control panel — as its parent or a dependency.',{confidence:'description-rated'}),
  drivenElectricalPath:auditRule('logic.driven-electrical-path-missing','logic','dependencies','Driven equipment traces to power','Pumps, fans, air handlers, chillers, and similar equipment should trace back to the electrical gear that powers them.',{confidence:'strong'}),
  controlElectricalPath:auditRule('logic.control-electrical-path-missing','logic','dependencies','Control equipment traces to power','Control panels and RIOs should trace back to the electrical supply that powers them.',{confidence:'strong'}),
  rioControlPath:auditRule('logic.rio-control-path-missing','logic','dependencies','RIO names its controller','An RIO should link to the PLC, I/O cluster, or upstream RIO that runs it.',{confidence:'strong'}),
  driveParent:auditRule('logic.drive-parent-unexpected','logic','structure','Drive sits under its equipment','A VFD or motor starter nests under the equipment it drives (the pump, fan, or air handler), with its power feed kept as a dependency.',{confidence:'strong'}),
  vfdDependencies:auditRule('sop.vfd-dependencies','sop','dependencies','VFD depends on its panel and PLC','Per the SOP a VFD lists its electrical panel and its PLC as dependencies.',{confidence:'strong'}),
  heatTraceChain:auditRule('logic.heat-trace-chain-missing','logic','structure','Heat trace follows its chain','Heat-trace panels sit under their transformer; heat-trace connection boxes sit under their panel or the upstream connection box.',{confidence:'strong'}),
  fduDependency:auditRule('logic.fdu-supported-equipment-missing','logic','dependencies','Fiber unit names what it serves','A fiber distribution unit should list the equipment or system it enables as a dependency.',{confidence:'strong'}),
  vesdaFireAlarm:auditRule('logic.vesda-fire-alarm-missing','logic','dependencies','VESDA depends on its fire alarm panel','Per the SOP a VESDA system lists its fire alarm panel as a dependency.',{confidence:'required'}),
  /* SOP nesting conventions */
  instrumentUpn:auditRule('sop.instrument-parent-upn','sop','structure','Instrument nests in the UPN in its tag','An instrument’s tag carries its UPN. It should nest under equipment in that UPN — not directly under the System Name, and never under equipment in a different UPN.',{confidence:'strong'}),
  fmsIoUnderVfd:auditRule('sop.fms-io-under-vfd','sop','structure','FMS I/O nests under the VFD','Per the SOP, FMS hardwired I/O for a drive nests under that VFD, with the PLC as a dependency.',{confidence:'strong'}),
  lcpPlacement:auditRule('sop.lcp-placement','sop','structure','LCP sits with its equipment','A local control panel nests under the equipment or skid it serves (MAH > SKID > LCP), or is a system root. It never nests under equipment in another UPN.',{confidence:'strong'}),
  untiedInstrumentRollup:auditRule('sop.untied-instrument-rollup','sop','structure','Loose instrument follows the UPN','An instrument not tied to equipment may sit under a piping or duct roll-up header only when that header is in the instrument’s own UPN.',{confidence:'strong'}),
  controlValveParent:auditRule('sop.control-valve-parent','sop','structure','Control valve sits under equipment','Per the SOP a control valve nests under the equipment it serves, not directly under the System Name.',{confidence:'strong'}),
  roomSensorParent:auditRule('sop.room-sensor-parent','sop','structure','Room sensor sits under its equipment','Per the SOP a room sensor nests under the equipment it controls, not under the room or area system.',{confidence:'strong'}),
});

function auditFinding(rule,severity,row,details={}){
  const equipmentId=clean(row&&row.equipmentId),source=row&&row._source||{},actual=details.actual==null?'':details.actual;
  const fingerprint=auditFingerprint([rule.id,auditNormId(equipmentId),source.sheet||'',source.row||'',JSON.stringify(actual)].join('|'));
  return Object.freeze({schemaVersion:1,id:`${rule.id}:${fingerprint}`,fingerprint,rule,severity,category:rule.category,
    equipmentId,row:source.row||0,sheet:source.sheet||'',field:details.field||'',why:details.why||'',actual,
    expected:details.expected==null?'':details.expected,recommendation:details.recommendation||'',relatedEquipmentId:details.relatedEquipmentId||'',
    /* Optional structured relationship for the UI to draw: {kind:'parent'|'dependency'|'loop', nodes:[{tag,role,upn,discipline}]}
       role ∈ 'this' (the flagged equipment) | 'parent' | 'dependency' | 'step' (a loop member). Text fields stay the human record. */
    relationship:details.relationship?Object.freeze({kind:details.relationship.kind,nodes:Object.freeze((details.relationship.nodes||[]).map(node=>Object.freeze({...node})))}):null,
    searchKey:auditNormId([rule.id,rule.source,rule.title,rule.category,equipmentId,source.sheet,details.field,details.why,actual,details.expected,details.relatedEquipmentId].join(' '))});
}
function auditRelNode(row,role,fallbackTag){return {tag:clean(row&&row.equipmentId)||clean(fallbackTag),role,upn:clean(row&&row.upn),discipline:clean(row&&row.discipline)};}
export function auditPolarity(discipline){return /elec|life safety|security|fire/i.test(discipline||'')?'top-down':'bottom-up';}
function auditDescription(row){return auditNormId(row&&row.equipmentDescription);}
function auditReferences(value){return auditSplitReferences(value).filter(reference=>!/^N\/?A$/i.test(reference));}
export function auditIsBlankItemMaster(row){return /(^|[^A-Z])BLANK([^A-Z]|$)/.test(auditNormId(row&&row.itemMaster));}
function auditIsElectrical(row){return auditNormId(row&&row.discipline)==='ELECTRICAL';}
function auditIsRio(row){return /REMOTE\s*I\/?O|\bRIO\b/.test(auditDescription(row));}
function auditIsDriveOrStarter(row){return /^VARIABLE FREQUENCY DRIVE\b|\bVFD SKID\b|\bMOTOR STARTER\b|^VFD\b/.test(auditDescription(row));}
function auditIsVfd(row){return /^VARIABLE FREQUENCY DRIVE\b|^VFD\b|\bVFD SKID\b/.test(auditDescription(row));}
function auditIsPlc(row){return /\bPLC\b|PROGRAMMABLE LOGIC CONTROLLER/.test(auditDescription(row));}
function auditIsPanel(row){return /\bPANEL\b|SWITCHBOARD|SWITCHGEAR|\bMCC\b|PANELBOARD|DISTRIBUTION BOARD/.test(auditDescription(row));}
function auditIsController(row){return /I\/?O CLUSTER|IO CLUSTER|REMOTE\s*I\/?O|\bRIO\b|\bPLC\b|PROGRAMMABLE LOGIC CONTROLLER|CONTROL PANEL|FIRE ALARM PANEL|GAS DETECTION PANEL|SECURITY CONTROL PANEL/.test(auditDescription(row))||auditIsDriveOrStarter(row);}
function auditRequiresRioController(row){return auditIsRio(row)&&!auditIsPlc(row);}
function auditIsControlEquipment(row){return auditIsController(row)||/FMS (?:CABINET|CONTROLLER)/.test(auditDescription(row));}
function auditIsDrivenEquipment(row){const description=auditDescription(row);return !auditIsElectrical(row)&&!/\bPLC\b|PROGRAMMABLE LOGIC CONTROLLER|CONTROL PANEL|\bVFD SKID\b|^VARIABLE FREQUENCY DRIVE\b|\bMOTOR STARTER\b/.test(description)&&/AIR HANDLER|\bMAH\b|\bAHU\b|PUMP|\bFAN\b|\bCHILLER\b|\bCOMPRESSOR\b|\bBLOWER\b|\bSCRUBBER\b|\bBOILER\b|COOLING TOWER|\bMIXER\b/.test(description);}
function auditIsTransformer(row){return /TRANSFORMER/.test(auditDescription(row));}
function auditIsHeatTracePanel(row){return /HEAT TRACE PANEL|HEAT TRACING PANEL/.test(auditDescription(row));}
function auditIsHeatTraceConnection(row){return /HEAT TRACE POWER CONNECTION/.test(auditDescription(row));}
function auditIsFdu(row){return /FIBER OPTIC DISTRIBUTION UNIT|\bFDU\b/.test(auditDescription(row));}
function auditIsVesda(row){return /\bVESDA\b|VERY EARLY SMOKE/.test(auditDescription(row));}
function auditIsFireAlarmPanel(row){return /FIRE ALARM PANEL/.test(auditDescription(row));}
/* Instrument-like rows: transmitters, switches, sensors, valves, detectors. Panels and
   controllers are excluded even when their descriptions mention a switch or valve. */
export function auditIsInstrument(row){const d=auditDescription(row);return !auditIsController(row)&&!auditIsPanel(row)&&/TRANSMITTER|\bSWITCH\b|SENSOR|INDICAT|ANALYZER|ELEMENT|VALVE|DETECTOR|THERMOSTAT|GAUGE|\bMETER\b|PROBE/.test(d);}
function auditIsControlValve(row){return /CONTROL VALVE|\bTCV\b|\bPCV\b|\bFCV\b|\bLCV\b|\bHCV\b/.test(auditDescription(row));}
function auditIsRoomSensor(row){return /ROOM (?:TEMPERATURE|HUMIDITY|PRESSURE|SENSOR|TEMP)|SPACE (?:TEMPERATURE|HUMIDITY|SENSOR)|ROOM (?:TEMP |)(?:TRANSMITTER|ELEMENT)/.test(auditDescription(row));}
function auditIsLcp(row){return /LOCAL CONTROL PANEL|\bLCP\b/.test(auditDescription(row));}
function auditIsFmsIo(row){return /FMS (?:HARDWIRED |)I\/?O|HARDWIRED I\/?O/.test(auditDescription(row));}
function auditIsPipingRollup(row){return /PIPING|DUCTWORK|\bDUCT\b|PIPE ROLL|DISTRIBUTION PIPING/.test(auditDescription(row));}
function auditControlExpectation(row){
  if(auditIsElectrical(row))return '';
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
function auditMilestonePairParts(row){
  const parent=clean(row&&row.milestoneParent),milestone=clean(row&&row.milestone);if(!parent||!milestone)return null;
  return {key:`${auditNormId(parent)} => ${auditNormId(milestone)}`,parent,milestone};
}
function auditLooksLikeMilestoneLevel(value,level){return new RegExp(`(?:^|[-_\\s])L${level}[-_\\s]+(?:M\\d+[-_\\s]+)?\\d`,'i').test(clean(value));}
export function auditMilestoneLevelIssues(row){
  const issues=[];
  if(auditLooksLikeMilestoneLevel(row&&row.milestoneParent,2))issues.push({row,field:'L1 Milestone Parent',actual:clean(row.milestoneParent),expected:'L1 milestone identifier',foundLevel:'L2'});
  if(auditLooksLikeMilestoneLevel(row&&row.milestone,1))issues.push({row,field:'L2 Milestone',actual:clean(row.milestone),expected:'L2 milestone identifier',foundLevel:'L1'});
  return issues;
}
export function auditMilestoneCohortCandidates(rows,{minimumPeers=10,minimumAgreement=.95}={}){
  const groups=new Map(),result=[];
  for(const row of rows||[]){const parts=auditMilestonePairParts(row),cohort=[row&&row.building,row&&row.upn,row&&row.equipmentDescription].map(auditNormId);if(!parts||cohort.some(value=>!value))continue;
    const key=cohort.join('|'),entry=groups.get(key)||{rows:[],counts:new Map(),representatives:new Map()};entry.rows.push(row);entry.counts.set(parts.key,(entry.counts.get(parts.key)||0)+1);if(!entry.representatives.has(parts.key))entry.representatives.set(parts.key,parts);groups.set(key,entry);}
  for(const entry of groups.values())for(const row of entry.rows){const current=auditMilestonePairParts(row),peerCount=entry.rows.length-1;if(peerCount<minimumPeers)continue;
    let expectedKey='',agreementCount=0;for(const [pairKey,count] of entry.counts){const peerMatches=count-(pairKey===current.key?1:0);if(peerMatches>agreementCount||peerMatches===agreementCount&&natCmp(pairKey,expectedKey)<0){expectedKey=pairKey;agreementCount=peerMatches;}}
    if(!expectedKey||expectedKey===current.key||agreementCount/peerCount<minimumAgreement)continue;const expected=entry.representatives.get(expectedKey);result.push({row,peerCount,agreementCount,expectedParent:expected.parent,expectedMilestone:expected.milestone});}
  return result;
}
export function auditMilestoneBranchCandidates(rows,{minimumPeers=5,minimumAgreement=.95}={}){
  const rowsById=new Map(),children=new Map(),result=[];
  for(const row of rows||[]){const id=auditNormId(row&&row.equipmentId);if(id&&!rowsById.has(id))rowsById.set(id,row);const parentId=auditNormId(row&&row.closestParent);if(parentId){const list=children.get(parentId)||[];list.push(row);children.set(parentId,list);}}
  for(const row of rows||[]){const current=auditMilestonePairParts(row),parent=rowsById.get(auditNormId(row&&row.closestParent)),parentPair=auditMilestonePairParts(parent);if(!current||!parentPair||current.key===parentPair.key)continue;
    const partition=[row&&row.building,row&&row.upn,row&&row.discipline].map(auditNormId),parentPartition=[parent&&parent.building,parent&&parent.upn,parent&&parent.discipline].map(auditNormId);if(partition.some(value=>!value)||partition.some((value,index)=>value!==parentPartition[index]))continue;
    const peers=(children.get(auditNormId(row.closestParent))||[]).filter(peer=>peer!==row&&auditMilestonePairParts(peer)&&[peer.building,peer.upn,peer.discipline].map(auditNormId).every((value,index)=>value===partition[index]));if(peers.length<minimumPeers)continue;
    const agreementCount=peers.filter(peer=>auditMilestonePairParts(peer).key===parentPair.key).length;if(agreementCount/peers.length<minimumAgreement)continue;result.push({row,parent,peerCount:peers.length,agreementCount,expectedParent:parentPair.parent,expectedMilestone:parentPair.milestone});
  }
  return result;
}
function auditItemMasterTokens(value){return clean(value).toUpperCase().replace(/[^A-Z0-9]+/g,' ').split(/\s+/).filter(Boolean);}
export function auditItemMasterCanonicalCandidates(value,canonicalValues){
  const input=auditItemMasterTokens(value),matches=[];if(!input.length)return matches;
  for(const canonicalValue of canonicalValues||[]){const canonical=auditItemMasterTokens(canonicalValue);if(canonical.length<2||!/^VF\d*$/.test(canonical[0]))continue;const body=canonical.slice(1);if(input.length<body.length)continue;const suffix=input.slice(-body.length);if(body.every((token,index)=>token===suffix[index]))matches.push(clean(canonicalValue));}
  return [...new Set(matches)].sort(natCmp);
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
/* The set of tags used as a Closest Parent that carry a Blank Item Master —
   the site's actual organizational headers. Never keyed on description text. */
export function auditHeaderIds(rows){
  const rowsById=new Map(),parentRefs=new Set();
  for(const row of rows||[]){const id=auditNormId(row&&row.equipmentId);if(id&&!rowsById.has(id))rowsById.set(id,row);const parentId=auditNormId(row&&row.closestParent);if(parentId)parentRefs.add(parentId);}
  const headers=new Set();
  for(const [id,row] of rowsById)if(parentRefs.has(id)&&auditIsBlankItemMaster(row))headers.add(id);
  return headers;
}

export function runSsmAudit(snapshot,options={}){
  const rows=snapshot&&snapshot.rows||[],findings=[];let checks=0;
  /* The approved VF Item Master names ship with the app; a caller may pass a
     narrower or newer list, and an explicit empty list disables vocabulary checks. */
  const itemMasterVocabulary=Array.isArray(options.itemMasterVocabulary)?options.itemMasterVocabulary:[...VF_ITEM_MASTER_NAMES];
  const add=(rule,severity,row,details)=>{if(rule.enabled)findings.push(auditFinding(rule,severity,row,details));};
  const rowsById=new Map(),duplicates=new Map();
  for(const row of rows){const key=auditNormId(row.equipmentId);if(!key)continue;if(rowsById.has(key)){const list=duplicates.get(key)||[rowsById.get(key)];list.push(row);duplicates.set(key,list);}else rowsById.set(key,row);}
  for(const list of duplicates.values())for(const row of list){checks++;add(SSM_AUDIT_RULES.duplicateId,'blocker',row,{field:'Equipment ID',why:'This tag appears on more than one row.',actual:row.equipmentId,expected:'One row per Equipment ID',recommendation:'Keep one row for this tag and merge or correct the others.'});}
  const headerIds=auditHeaderIds(rows),systemNames=new Set(rows.map(row=>auditNormId(row.systemName)).filter(Boolean));
  const children=new Map(),parentEdges=new Map(),precedenceEdges=new Map(),generatedHeaders=new Map(),nodes=new Set(rowsById.keys());
  const addEdge=(map,from,to)=>{if(!from||!to)return;const set=map.get(from)||new Set();set.add(to);map.set(from,set);};
  const upnGroups=new Map(),milestoneGroups=new Map();
  const projectUsesMilestones=SSM_AUDIT_RULES.milestonePair.enabled&&rows.some(row=>clean(row.milestone)||clean(row.milestoneParent));
  const rowUpn=row=>auditNormId(row&&row.upn);
  const populatedItemMasters=rows.filter(row=>clean(row.itemMaster)&&!auditIsBlankItemMaster(row));
  const legacyItemMasterRegistry=populatedItemMasters.length>=50&&populatedItemMasters.filter(row=>!/^VF/.test(auditNormId(row.itemMaster))).length/populatedItemMasters.length>0.5;
  const isHeaderRow=row=>headerIds.has(auditNormId(row&&row.equipmentId));
  for(const row of rows){
    const id=auditNormId(row.equipmentId),parentId=auditNormId(row.closestParent),status=extoRev21Canonical('closestParentStatus',row.closestParentStatus)||clean(row.closestParentStatus).toUpperCase();
    const upn=rowUpn(row);
    /* --- approved lists --- */
    if(upn){checks++;if(!extoRev21IsUpn(upn))add(SSM_AUDIT_RULES.upnNotApproved,'blocker',row,{field:'UPN',why:'This UPN is not on the Rev21 approved list.',actual:row.upn,expected:'A UPN from the Rev21 upload template',recommendation:'Correct the UPN to the approved value for this system.'});}
    if(extoRev21EffectiveDiscipline(row.discipline)==='FACILITIES MONITORING SYSTEM'&&extoRev21Norm(row.discipline)!=='FACILITIES MONITORING SYSTEM'){checks++;const candidates=extoRev21UpnCandidates(row.equipmentId);add(SSM_AUDIT_RULES.icDiscipline,'blocker',row,{field:'Discipline',why:'This I&C row is not using the approved controls discipline.',actual:`${row.discipline}; UPN in tag: ${candidates.join(', ')||'none found'}`,expected:'FACILITIES MONITORING SYSTEM with the UPN taken from the tag',recommendation:candidates.length===1?`Set the discipline to the approved value and use UPN ${candidates[0]}.`:'Set the approved discipline and confirm the UPN from the tag before assigning the System Name.'});}
    if(row.systemName&&upn&&extoRev21IsUpn(upn)){checks++;
      const approved=extoRev21SystemsForUpn(upn),normalizedApproved=new Set(approved.map(extoRev21Norm)),sys=extoRev21Norm(row.systemName);
      if(!normalizedApproved.has(sys)){
        const inVocabulary=extoRev21IsSystemName(row.systemName);
        add(SSM_AUDIT_RULES.systemUpn,'blocker',row,{field:'System Name',
          why:inVocabulary?'This System Name belongs to a different UPN than the row is assigned to.':'This System Name is not an approved Rev21 name for this UPN.',
          actual:`UPN ${row.upn}; System ${row.systemName}`,
          expected:approved.length?`One of: ${approved.join(' | ')}`:`An approved Rev21 System Name for UPN ${row.upn}`,
          recommendation:inVocabulary?'Either the UPN or the System Name is wrong — correct whichever does not match this equipment.':'Pick the approved System Name for this UPN from the Rev21 list.'});
      }
    }
    /* --- milestones --- */
    /* Info, not warning: real registries roll milestones out gradually, and
       thousands of not-yet-assigned rows must not bury the actionable list. */
    if(projectUsesMilestones){checks++;if(!clean(row.milestone)||!clean(row.milestoneParent))add(SSM_AUDIT_RULES.milestonePair,'info',row,{field:'Milestones',why:'This project uses milestones, but this row is missing its L1 or L2.',actual:`L1: ${row.milestoneParent||'blank'}; L2: ${row.milestone||'blank'}`,expected:'Both L1 Milestone Parent and L2 Milestone filled in',recommendation:'Assign the governing L1 milestone and its L2 milestone.'});}
    if(clean(row.milestone)&&upn&&SSM_AUDIT_RULES.milestoneUpn.enabled){checks++;
      const inName=extoRev21UpnCandidates(row.milestone);
      if(inName.length&&!inName.includes(upn))add(SSM_AUDIT_RULES.milestoneUpn,'warning',row,{field:'L2 Milestone',why:`The L2 milestone names UPN ${inName.join('/')}, but this row is on UPN ${row.upn}.`,actual:row.milestone,expected:`An L2 milestone for UPN ${row.upn}`,recommendation:'Confirm the row belongs to this L2 milestone; if not, assign the L2 milestone for its own UPN.'});
      else if(!inName.length&&/^[0-9]+$/.test(upn))add(SSM_AUDIT_RULES.milestoneUpn,'info',row,{field:'L2 Milestone',why:'The L2 milestone name does not mention a UPN, so it cannot be checked against this row.',actual:row.milestone,expected:`An L2 milestone naming UPN ${row.upn}`,recommendation:'Confirm the milestone applies to this system.'});
    }
    if(row.milestone&&row.milestoneParent){
      if(SSM_AUDIT_RULES.milestoneIntent.enabled){const conflict=auditMilestoneConflict(row.milestoneParent,row.milestone);checks++;
        if(conflict)add(SSM_AUDIT_RULES.milestoneIntent,'warning',row,{field:'Milestones',why:`L1 and L2 describe different phases: ${conflict}.`,actual:`L1: ${row.milestoneParent}; L2: ${row.milestone}`,expected:'L1 and L2 milestones for the same phase',recommendation:'Confirm which phase this equipment belongs to and align both milestones.'});}
      if(SSM_AUDIT_RULES.milestoneInconsistent.enabled){const key=auditNormId(row.milestone),entry=milestoneGroups.get(key)||{row,parents:new Map()};
        const intent=auditMilestoneIntentKey(row.milestoneParent),values=entry.parents.get(intent)||new Set();values.add(clean(row.milestoneParent));entry.parents.set(intent,values);milestoneGroups.set(key,entry);}
    }
    if(upn){const group=upnGroups.get(upn)||{systems:new Set(),disciplines:new Set(),rows:[]};if(row.systemName)group.systems.add(auditNormId(row.systemName));if(row.discipline)group.disciplines.add(auditNormId(row.discipline));group.rows.push(row);upnGroups.set(upn,group);}
    /* --- parents --- */
    checks++;if(!parentId)add(SSM_AUDIT_RULES.blankParent,'blocker',row,{field:'Closest Parent',why:'No Closest Parent is filled in.',expected:'The equipment this nests under, or the row’s own System Name if it is the top of the system',recommendation:'Enter the parent equipment in the same UPN, or the System Name for a root.'});
    else if(parentId===id)add(SSM_AUDIT_RULES.selfParent,'blocker',row,{field:'Closest Parent',why:'The row lists itself as its own parent.',actual:row.closestParent,expected:'A different parent tag or the System Name',recommendation:'Enter the real parent.'});
    else{
      const parent=rowsById.get(parentId);
      if(parent){
        addEdge(parentEdges,id,parentId);const list=children.get(parentId)||[];list.push(row);children.set(parentId,list);
        const childUpn=upn,parentUpn=rowUpn(parent);checks++;
        if(childUpn&&parentUpn&&childUpn!==parentUpn)add(SSM_AUDIT_RULES.crossUpn,'error',row,{field:'Closest Parent',why:`This row is on UPN ${row.upn} but its parent is on UPN ${parent.upn}. A parent must be in the same UPN.`,actual:`${row.equipmentId} (${row.upn}) → ${parent.equipmentId} (${parent.upn})`,expected:`A parent in UPN ${row.upn}`,recommendation:`Keep ${parent.equipmentId} as a dependency and pick a parent inside UPN ${row.upn} (or the System Name if this is the top).`,relatedEquipmentId:parent.equipmentId,relationship:{kind:'parent',nodes:[auditRelNode(row,'this'),auditRelNode(parent,'parent')]}});
        const childDiscipline=auditNormId(row.discipline),parentDiscipline=auditNormId(parent.discipline);checks++;
        const approvedControlsChild=extoRev21EffectiveDiscipline(row.discipline)==='FACILITIES MONITORING SYSTEM'||auditIsControlEquipment(row)||auditIsInstrument(row);
        if(childDiscipline&&parentDiscipline&&childDiscipline!==parentDiscipline&&!approvedControlsChild)add(SSM_AUDIT_RULES.crossDiscipline,'warning',row,{field:'Closest Parent',why:`This row is ${row.discipline} but its parent is ${parent.discipline}.`,actual:`${row.discipline} → ${parent.discipline}`,expected:`A parent in ${row.discipline}, unless this is a controls device under the equipment it serves`,recommendation:'Confirm the exception, or record the relationship as a dependency instead.',relatedEquipmentId:parent.equipmentId,relationship:{kind:'parent',nodes:[auditRelNode(row,'this'),auditRelNode(parent,'parent')]}});
        if(auditPolarity(row.discipline)==='top-down')addEdge(precedenceEdges,parentId,id);else addEdge(precedenceEdges,id,parentId);
      }else if(status==='NEW'){
        if(parentId!==auditNormId(row.systemName)){
          const entry=generatedHeaders.get(parentId)||{row,parent:row.closestParent,count:0};entry.count++;generatedHeaders.set(parentId,entry);
        }
      }else if(status!=='EXISTING')add(SSM_AUDIT_RULES.unresolvedParent,'blocker',row,{field:'Closest Parent',why:'The Closest Parent is not a tag in this registry and is not marked as Existing.',actual:row.closestParent,expected:'A tag in this registry, an intentional header, or an Existing parent from another project',recommendation:'Fix the parent tag, or set Closest Parent Status to Existing if it lives in another project.'});
    }
    /* --- dependencies --- */
    const dependencies=auditReferences(row.dependencies),seenDeps=new Set();
    for(const dependency of dependencies){const depId=auditNormId(dependency);checks++;
      if(depId===id)add(SSM_AUDIT_RULES.selfDependency,'error',row,{field:'Dependencies',why:'The row lists itself as a dependency.',actual:dependency,expected:'A different predecessor tag',recommendation:'Remove the self-reference.'});
      if(seenDeps.has(depId))add(SSM_AUDIT_RULES.duplicateDependency,'warning',row,{field:'Dependencies',why:'The same dependency is listed twice.',actual:dependency,expected:'Each dependency once',recommendation:'Remove the duplicate.'});seenDeps.add(depId);
      const target=rowsById.get(depId);
      if(depId&&!target&&!clean(row.dependencyProject))add(SSM_AUDIT_RULES.unresolvedDependency,'error',row,{field:'Dependencies',why:'This dependency is not a tag in the registry, and no other project is named for it.',actual:dependency,expected:'A tag in this registry, or a Dependency Project naming where it lives',recommendation:'Correct the tag, or fill in the Dependency Project.'});
      if(target&&depId!==id&&depId!==parentId)addEdge(precedenceEdges,depId,id);
      if(target&&depId!==id&&SSM_AUDIT_RULES.sameUpnBottomUp.enabled&&upn&&rowUpn(target)===upn&&auditPolarity(row.discipline)==='bottom-up'&&auditNormId(target.discipline)===auditNormId(row.discipline)&&!auditIsInstrument(row)&&!isHeaderRow(row)){
        checks++;add(SSM_AUDIT_RULES.sameUpnBottomUp,'info',row,{field:'Dependencies',why:`This dependency is inside the same UPN (${row.upn}) and discipline. In a bottom-up discipline the hierarchy already sets the order, so a same-system dependency is only needed for a real sequencing reason.`,actual:dependency,expected:'Same-UPN dependencies mainly on instruments and devices; on other equipment, only for a documented sequencing reason',recommendation:'Confirm the dependency is intended. If the equipment simply nests, the hierarchy already covers the order.',relatedEquipmentId:target.equipmentId,relationship:{kind:'dependency',nodes:[auditRelNode(row,'this'),auditRelNode(target,'dependency')]}});
      }
    }
    /* --- item masters ---
       Any site-prefixed name (CA_NB_…, SP_NB_…, EL_…) is a legacy assignment the
       VF standard replaces; the VF equivalent is proposed by matching the name's
       tail. When a registry is wholesale non-VF this is a migration, not 20,000
       separate mistakes, so the per-row note drops to info — the count is the
       signal, and the structural findings stay on top. */
    if(SSM_AUDIT_RULES.itemMasterStandard.enabled&&clean(row.itemMaster)){checks++;
      const im=auditNormId(row.itemMaster),known=itemMasterVocabulary.some(value=>auditNormId(value)===im);
      const legacy=/^[A-Z]{1,4}_/.test(im)&&!/^VF/.test(im);
      /* A legacy site prefix is always reportable. A VF-looking name that is
         merely absent from the list is only judged when a vocabulary is in
         force — an explicit empty list means "do not judge unknown names". */
      if(!known&&!auditIsBlankItemMaster(row)&&(legacy||itemMasterVocabulary.length)){
        const candidates=auditItemMasterCanonicalCandidates(row.itemMaster,itemMasterVocabulary);
        add(SSM_AUDIT_RULES.itemMasterStandard,legacyItemMasterRegistry?'info':'warning',row,{field:'Item Master Unique Identifier',
          why:legacy?`This is a site-prefixed Item Master (${im.split('_')[0]}_). The VF standard name should be used.`:'This Item Master is not one of the approved VF names in the Standardized Item Master Template.',
          actual:row.itemMaster,expected:candidates.length===1?candidates[0]:candidates.length?`One of: ${candidates.join(' | ')}`:'The matching VF_ Item Master',
          recommendation:candidates.length===1?`Replace with ${candidates[0]}.`:'Replace with the VF Item Master that matches this equipment.'});
      }
    }
  }
  for(const entry of generatedHeaders.values()){checks++;add(SSM_AUDIT_RULES.generatedHeader,'warning',entry.row,{field:'Closest Parent',why:`${entry.count.toLocaleString()} row${entry.count===1?' nests':'s nest'} under a parent that is not any equipment in this registry.`,actual:entry.parent,expected:'An intentional header row (with a Blank Item Master) or the row’s System Name',recommendation:'If this is meant to be a header, add it as a row with a Blank Item Master. If not, correct the parent.'});}
  for(const [upn,group] of upnGroups){checks++;if(group.systems.size>1||group.disciplines.size>1)add(SSM_AUDIT_RULES.upnInconsistent,'error',group.rows[0],{field:'UPN',why:`Rows on UPN ${upn} use more than one System Name or Discipline.`,actual:`Systems: ${[...group.systems].join('; ')}; Disciplines: ${[...group.disciplines].join('; ')}`,expected:'One System Name and one Discipline for the whole UPN',recommendation:`Review every row on UPN ${upn} and align them.`});}
  for(const entry of milestoneGroups.values()){
    const meaningful=[...entry.parents].filter(([intent])=>intent!=='unclassified');checks++;
    if(meaningful.length>1)add(SSM_AUDIT_RULES.milestoneInconsistent,'warning',entry.row,{field:'Milestone Parent',why:'This L2 milestone rolls up to L1 milestones that mean different things.',actual:meaningful.map(([intent])=>intent).join('; '),expected:'One consistent L1 meaning per L2 milestone',recommendation:'Review the rows on this L2 milestone and settle on one L1.'});
  }
  const milestoneCohortRows=new Set();
  if(SSM_AUDIT_RULES.milestoneLevel.enabled)for(const row of rows)for(const issue of auditMilestoneLevelIssues(row)){checks++;add(SSM_AUDIT_RULES.milestoneLevel,'warning',row,{field:issue.field,why:`${issue.field} holds an identifier written for ${issue.foundLevel}.`,actual:issue.actual,expected:issue.expected,recommendation:'Move the milestone to the right field and confirm the pair.'});}
  if(SSM_AUDIT_RULES.milestoneCohort.enabled)for(const candidate of auditMilestoneCohortCandidates(rows)){checks++;milestoneCohortRows.add(candidate.row);add(SSM_AUDIT_RULES.milestoneCohort,'warning',candidate.row,{field:'Milestones',why:`${candidate.agreementCount.toLocaleString()} comparable rows in the same building and system use a different milestone pair.`,actual:`L1: ${candidate.row.milestoneParent}; L2: ${candidate.row.milestone}`,expected:`L1: ${candidate.expectedParent}; L2: ${candidate.expectedMilestone}`,recommendation:'Confirm this row is really an exception; otherwise match its neighbours.'});}
  if(SSM_AUDIT_RULES.milestoneBranch.enabled)for(const candidate of auditMilestoneBranchCandidates(rows)){if(milestoneCohortRows.has(candidate.row))continue;checks++;add(SSM_AUDIT_RULES.milestoneBranch,'warning',candidate.row,{field:'Milestones',why:`Its parent and ${candidate.agreementCount.toLocaleString()} siblings share one milestone pair; this row differs.`,actual:`L1: ${candidate.row.milestoneParent}; L2: ${candidate.row.milestone}`,expected:`L1: ${candidate.expectedParent}; L2: ${candidate.expectedMilestone}`,recommendation:'Confirm the difference is intended; otherwise match the branch.',relatedEquipmentId:candidate.parent.equipmentId});}
  for(const cycle of auditCyclePaths(nodes,parentEdges)){checks++;const row=rowsById.get(cycle[0]);add(SSM_AUDIT_RULES.parentCycle,'blocker',row,{field:'Closest Parent',why:'Following Closest Parent upward comes back to this equipment.',actual:cycle.map(id=>rowsById.get(id)?.equipmentId||id).join(' → '),expected:'A hierarchy with no loops',recommendation:'Change one Closest Parent in this loop.',relationship:{kind:'loop',nodes:cycle.map((id,index)=>auditRelNode(rowsById.get(id),index===0?'this':'step',id))}});}
  for(const cycle of auditCyclePaths(nodes,precedenceEdges)){checks++;const row=rowsById.get(cycle[0]);add(SSM_AUDIT_RULES.precedenceCycle,'blocker',row,{field:'Dependencies',why:'These parent and dependency links make the startup order go in a circle.',actual:cycle.map(id=>rowsById.get(id)?.equipmentId||id).join(' → '),expected:'A startup order with no loops',recommendation:'Break the circle by removing or moving one of these links.',relationship:{kind:'loop',nodes:cycle.map((id,index)=>auditRelNode(rowsById.get(id),index===0?'this':'step',id))}});}
  /* --- headers: rows used as parents that carry a Blank Item Master --- */
  const parentRefCounts=new Map();for(const row of rows){const p=auditNormId(row.closestParent);if(p)parentRefCounts.set(p,(parentRefCounts.get(p)||0)+1);}
  for(const row of rows){
    const id=auditNormId(row.equipmentId);if(!id)continue;
    const referenced=parentRefCounts.get(id)||0,blank=auditIsBlankItemMaster(row);
    if(referenced&&!blank&&isLikelyHeaderName(row,children.get(id)||[])){checks++;add(SSM_AUDIT_RULES.headerItemMaster,'warning',row,{field:'Item Master Unique Identifier',why:`${referenced.toLocaleString()} row${referenced===1?' nests':'s nest'} under this row as a group, but it carries a real Item Master, so checklists would be applied to a header.`,actual:row.itemMaster||'Blank',expected:'VF_Blank (or the site’s Blank Item Master)',recommendation:'If this row is an organizational header, set its Item Master to VF_Blank. If it is real equipment, leave it.'});}
    if(blank&&referenced){
      if(auditReferences(row.dependencies).length){checks++;add(SSM_AUDIT_RULES.headerDependency,'warning',row,{field:'Dependencies',why:'This header groups equipment but also carries dependencies.',actual:row.dependencies,expected:'No dependencies on a header',recommendation:'Move the dependencies to the equipment that actually needs them.'});}
    }else if(blank&&!referenced){checks++;add(SSM_AUDIT_RULES.unusedHeader,'info',row,{field:'Equipment ID',why:'This row has a Blank Item Master (a header) but nothing nests under it.',actual:row.equipmentId,expected:'At least one piece of equipment nested under it',recommendation:'Attach the intended equipment, or remove the header if it is not needed.'});}
  }
  /* --- commissioning logic + SOP nesting --- */
  const electricalSeeds=new Set(rows.filter(auditIsElectrical).map(row=>auditNormId(row.equipmentId)).filter(Boolean));
  const electricalPath=auditReachableFrom(electricalSeeds,precedenceEdges);
  for(const row of rows){
    const id=auditNormId(row.equipmentId);if(!id||isHeaderRow(row))continue;
    const parent=rowsById.get(auditNormId(row.closestParent)),dependencyReferences=auditReferences(row.dependencies),dependencies=dependencyReferences.map(value=>rowsById.get(auditNormId(value))).filter(Boolean),hasExternalDependency=!!clean(row.dependencyProject)&&dependencyReferences.some(value=>!rowsById.has(auditNormId(value))),hasExternalParent=!parent&&extoRev21Canonical('closestParentStatus',row.closestParentStatus)==='EXISTING',related=[parent,...dependencies].filter(Boolean);
    const upn=rowUpn(row),parentIsSystem=!parent&&auditNormId(row.closestParent)===auditNormId(row.systemName);
    const controlExpectation=auditControlExpectation(row);
    if(controlExpectation){checks++;if(!hasExternalDependency&&!hasExternalParent&&!related.some(auditIsController))add(SSM_AUDIT_RULES.controlLink,controlExpectation==='strong'?'warning':'info',row,{field:'Dependencies',why:'This control device does not point to what controls it (an RIO, PLC, VFD, or control panel).',actual:row.dependencies||'Blank',expected:'The control source as parent or dependency',recommendation:'Add the RIO, PLC, VFD, or control panel that runs this device.'});}
    if(auditIsDrivenEquipment(row)){checks++;if(!hasExternalDependency&&!hasExternalParent&&!electricalPath.has(id))add(SSM_AUDIT_RULES.drivenElectricalPath,'warning',row,{field:'Dependencies',why:'There is no path from this equipment back to the electrical gear that powers it.',actual:row.dependencies||'Blank',expected:'A parent or dependency chain reaching its panel, starter, or VFD',recommendation:'Add the supplying panel, starter, or VFD as a dependency.'});}
    if(!auditIsElectrical(row)&&auditIsControlEquipment(row)){checks++;if(!hasExternalDependency&&!hasExternalParent&&!electricalPath.has(id))add(SSM_AUDIT_RULES.controlElectricalPath,auditIsRio(row)?'warning':'info',row,{field:'Dependencies',why:'There is no path from this control equipment back to its power supply.',actual:row.dependencies||'Blank',expected:'A chain reaching its supplying panel, circuit, or transformer',recommendation:'Add the electrical equipment that powers it as a dependency.'});}
    if(auditRequiresRioController(row)){checks++;if(!hasExternalDependency&&!hasExternalParent&&!related.some(auditIsController))add(SSM_AUDIT_RULES.rioControlPath,'warning',row,{field:'Dependencies',why:'This RIO does not name the controller or upstream I/O that runs it.',actual:row.dependencies||'Blank',expected:'PLC, I/O cluster, or upstream RIO',recommendation:'Add the PLC or upstream I/O as a dependency.'});}
    if(auditIsDriveOrStarter(row)){checks++;if(!parent||!auditIsDrivenEquipment(parent))add(SSM_AUDIT_RULES.driveParent,'warning',row,{field:'Closest Parent',why:'This drive or starter is not nested under the equipment it runs.',actual:parent?`${row.closestParent} (${parent.equipmentDescription||'no description'})`:(row.closestParent||'Blank'),expected:'The pump, fan, air handler, chiller, or similar equipment it drives',recommendation:'Move it under the equipment it drives and keep its power feed as a dependency.'});
      if(auditIsVfd(row)&&SSM_AUDIT_RULES.vfdDependencies.enabled&&!hasExternalDependency){checks++;const hasPanel=dependencies.some(d=>auditIsPanel(d)&&!auditIsController(d))||dependencies.some(auditIsElectrical),hasPlc=dependencies.some(auditIsPlc)||dependencies.some(auditIsController);
        if(!hasPanel||!hasPlc)add(SSM_AUDIT_RULES.vfdDependencies,'warning',row,{field:'Dependencies',why:`Per the SOP a VFD lists both its electrical panel and its PLC as dependencies. ${!hasPanel&&!hasPlc?'Neither is listed.':!hasPanel?'No electrical panel is listed.':'No PLC is listed.'}`,actual:row.dependencies||'Blank',expected:'The supplying panel and the controlling PLC',recommendation:'Add the missing panel or PLC as a dependency.'});}
    }
    if(auditIsHeatTracePanel(row)){checks++;if(!parent||!auditIsTransformer(parent))add(SSM_AUDIT_RULES.heatTraceChain,'warning',row,{field:'Closest Parent',why:'This heat-trace panel is not nested under a transformer.',actual:parent?`${row.closestParent} (${parent.equipmentDescription||'no description'})`:(row.closestParent||'Blank'),expected:'The transformer that supplies it',recommendation:'Nest the panel under its supplying transformer.'});}
    if(auditIsHeatTraceConnection(row)){checks++;if(!parent||(!auditIsHeatTracePanel(parent)&&!auditIsHeatTraceConnection(parent)))add(SSM_AUDIT_RULES.heatTraceChain,'warning',row,{field:'Closest Parent',why:'This heat-trace connection box is not under its panel or the upstream connection box.',actual:parent?`${row.closestParent} (${parent.equipmentDescription||'no description'})`:(row.closestParent||'Blank'),expected:'A heat-trace panel or upstream connection box',recommendation:'Nest it in the correct heat-trace branch.'});}
    if(auditIsFdu(row)){checks++;if(!dependencyReferences.length)add(SSM_AUDIT_RULES.fduDependency,'warning',row,{field:'Dependencies',why:'This fiber distribution unit does not say what it serves.',actual:'Blank',expected:'The equipment or system it enables',recommendation:'Add the served equipment or system as a dependency.'});}
    if(auditIsVesda(row)){checks++;if(!hasExternalDependency&&!dependencies.some(auditIsFireAlarmPanel))add(SSM_AUDIT_RULES.vesdaFireAlarm,'error',row,{field:'Dependencies',why:'Per the SOP a VESDA system depends on its fire alarm panel; none is listed.',actual:row.dependencies||'Blank',expected:'The associated fire alarm panel',recommendation:'Add the fire alarm panel as a dependency.'});}
    /* SOP nesting conventions — the tag's UPN is the guiderail */
    const tagUpns=extoRev21UpnCandidates(row.equipmentId),tagUpn=tagUpns.length===1?tagUpns[0]:'';
    if(auditIsInstrument(row)&&SSM_AUDIT_RULES.instrumentUpn.enabled&&tagUpn){checks++;
      if(parent&&rowUpn(parent)&&rowUpn(parent)!==tagUpn&&!isHeaderRow(parent))add(SSM_AUDIT_RULES.instrumentUpn,'error',row,{field:'Closest Parent',why:`This instrument’s tag carries UPN ${tagUpn}, but it is nested under equipment on UPN ${parent.upn}.`,actual:`${row.closestParent} (UPN ${parent.upn})`,expected:`Equipment in UPN ${tagUpn}`,recommendation:`Nest it under the equipment it serves in UPN ${tagUpn}.`,relatedEquipmentId:parent.equipmentId,relationship:{kind:'parent',nodes:[auditRelNode(row,'this'),auditRelNode(parent,'parent')]}});
      else if(parentIsSystem&&!auditIsControlValve(row))add(SSM_AUDIT_RULES.instrumentUpn,'warning',row,{field:'Closest Parent',why:'This instrument sits directly under the System Name instead of the equipment it belongs to.',actual:row.closestParent,expected:`Equipment (or an approved roll-up header) in UPN ${tagUpn}`,recommendation:'Nest it under the equipment or skid it is mounted on.'});
    }
    if(auditIsInstrument(row)&&SSM_AUDIT_RULES.untiedInstrumentRollup.enabled&&tagUpn&&parent&&isHeaderRow(parent)&&auditIsPipingRollup(parent)&&rowUpn(parent)&&rowUpn(parent)!==tagUpn){checks++;
      add(SSM_AUDIT_RULES.untiedInstrumentRollup,'error',row,{field:'Closest Parent',why:`This instrument sits under a piping/duct roll-up on UPN ${parent.upn}, but its tag carries UPN ${tagUpn}.`,actual:`${row.closestParent} (UPN ${parent.upn})`,expected:`A roll-up or equipment in UPN ${tagUpn}`,recommendation:`Move it to the roll-up for UPN ${tagUpn}.`,relatedEquipmentId:parent.equipmentId,relationship:{kind:'parent',nodes:[auditRelNode(row,'this'),auditRelNode(parent,'parent')]}});}
    if(auditIsControlValve(row)&&SSM_AUDIT_RULES.controlValveParent.enabled&&parentIsSystem){checks++;add(SSM_AUDIT_RULES.controlValveParent,'warning',row,{field:'Closest Parent',why:'This control valve sits directly under the System Name.',actual:row.closestParent,expected:'The equipment the valve serves',recommendation:'Nest the valve under its equipment.'});}
    if(auditIsRoomSensor(row)&&SSM_AUDIT_RULES.roomSensorParent.enabled&&(parentIsSystem||(parent&&auditNormId(parent.discipline)==='ROOM/AREA/BAY-READY'))){checks++;add(SSM_AUDIT_RULES.roomSensorParent,'warning',row,{field:'Closest Parent',why:'This room sensor is placed under the room or area rather than the equipment it controls.',actual:row.closestParent,expected:'The equipment this sensor controls (for example the air handler serving the room)',recommendation:'Nest the sensor under the equipment it controls.'});}
    if(auditIsLcp(row)&&SSM_AUDIT_RULES.lcpPlacement.enabled&&parent&&rowUpn(parent)&&upn&&rowUpn(parent)!==upn){checks++;add(SSM_AUDIT_RULES.lcpPlacement,'error',row,{field:'Closest Parent',why:`This local control panel is nested under equipment on UPN ${parent.upn}, but it is on UPN ${row.upn}.`,actual:`${row.closestParent} (UPN ${parent.upn})`,expected:`Its own equipment or skid in UPN ${row.upn}, or the System Name`,recommendation:'Nest the LCP under the equipment or skid it serves in its own UPN.',relatedEquipmentId:parent.equipmentId,relationship:{kind:'parent',nodes:[auditRelNode(row,'this'),auditRelNode(parent,'parent')]}});}
    if(auditIsFmsIo(row)&&SSM_AUDIT_RULES.fmsIoUnderVfd.enabled){checks++;const underVfd=parent&&auditIsVfd(parent),hasPlc=dependencies.some(auditIsPlc)||dependencies.some(auditIsController);
      if(!underVfd||!hasPlc)add(SSM_AUDIT_RULES.fmsIoUnderVfd,'warning',row,{field:underVfd?'Dependencies':'Closest Parent',why:!underVfd?'This FMS hardwired I/O is not nested under a VFD.':'This FMS hardwired I/O does not list its PLC as a dependency.',actual:!underVfd?(row.closestParent||'Blank'):(row.dependencies||'Blank'),expected:!underVfd?'The VFD it is wired to':'The controlling PLC as a dependency',recommendation:!underVfd?'Nest it under its VFD, with the PLC as a dependency.':'Add the PLC as a dependency.'});}
  }
  findings.sort((a,b)=>(SSM_AUDIT_SEVERITY_RANK[a.severity]-SSM_AUDIT_SEVERITY_RANK[b.severity])||natCmp(a.category,b.category)||natCmp(a.equipmentId,b.equipmentId)||a.row-b.row||natCmp(a.rule.id,b.rule.id));
  const severity=Object.fromEntries(SSM_AUDIT_SEVERITIES.map(level=>[level,findings.filter(finding=>finding.severity===level).length]));
  const category=Object.fromEntries(SSM_AUDIT_CATEGORIES.map(name=>[name,findings.filter(finding=>finding.category===name).length]));
  const source=Object.fromEntries(SSM_AUDIT_SOURCES.map(item=>[item.id,findings.filter(finding=>finding.rule.source===item.id).length]));
  return Object.freeze({schemaVersion:1,standard:SSM_AUDIT_STANDARD,rows:Object.freeze([...rows]),findings:Object.freeze(findings),headerIds:Object.freeze([...headerIds]),
    summary:Object.freeze({rows:rows.length,checks,findings:findings.length,severity:Object.freeze(severity),category:Object.freeze(category),source:Object.freeze(source),status:severity.blocker?'blocked':severity.error||severity.warning?'review':'ready'})});
}
/* A parented row without a Blank Item Master is only called out as a probable
   header when it reads like one: no description of its own equipment kind and
   several children. Real equipment with children (an AHU with its instruments)
   is never flagged this way. */
function isLikelyHeaderName(row,children){
  const description=auditDescription(row);
  if(children.length<3)return false;
  return !description||/HEADER|CLUSTER|RING|GROUP|ROLL[- ]?UP|PIPING|DUCTWORK|SKID/.test(description)&&!auditIsInstrument(row)&&!auditIsController(row);
}
