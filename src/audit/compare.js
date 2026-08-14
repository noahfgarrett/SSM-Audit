import { clean, natCmp } from '../core/text.js'
import { auditNormId, auditSplitReferences } from './model.js'

export const SSM_COMPARISON_SCHEMA_VERSION=1;
const EMPTY_ROLE='Unclassified equipment';
const CLASSIFICATION_FAMILIES=new Map(Object.entries({
  'AFD-VFD':'Variable frequency drive',VFD:'Variable frequency drive',
  TET:'Temperature transmitter',TT:'Temperature transmitter',TIT:'Temperature transmitter',
  FIT:'Flow transmitter',FT:'Flow transmitter',PIT:'Pressure transmitter',PT:'Pressure transmitter',
  PDIT:'Differential pressure transmitter',PDT:'Differential pressure transmitter',
  AIT:'Analyzer transmitter',AT:'Analyzer transmitter',
  TCV:'Control valve',CV:'Control valve',FV:'Control valve',
  TFRM:'Transformer',DP:'Distribution panel',DTP:'Distribution panel',BRP:'Branch panel',
  GIS:'Gas-insulated switchgear','SWG-GIS':'Gas-insulated switchgear',
  MVSUB:'Medium-voltage substation',LVSUB:'Low-voltage substation',CB:'Circuit breaker',
  BAT:'Battery system',MBAT:'Battery monitoring system',CHRG:'Battery charger',
  UPS:'Uninterruptible power supply','UPS-10':'Uninterruptible power supply',
  RIO:'Remote I/O panel',PLC:'Programmable logic controller',LCP:'Local control panel',
  OIT:'Operator interface',HMI:'Operator interface',GTW:'Network gateway',RACK:'Equipment rack',
  STX:'Industrial network switch','CISCO SWITCH':'Industrial network switch',
  BSTX:'Industrial network switch panel','STRATIX ENCLOSURE':'Industrial network switch enclosure',
  PMP:'Pump',LAT:'Lateral',MAIN:'Main distribution',HS:'Hand switch',
  ZSO:'Position switch',ZSC:'Position switch',QS:'Current switch',QL:'Status indicator',
  XV:'On/off valve',TK:'Tank',FDU:'Fiber distribution unit','FDU ENCLOSURE':'Fiber distribution unit enclosure',
}));
const DESCRIPTION_FAMILIES=[
  [/GAS[- ]INSULATED SWITCHGEAR|\bGIS\b/,'Gas-insulated switchgear'],
  [/MEDIUM[- ]VOLTAGE SUBSTATION|\bMVSS\b/,'Medium-voltage substation'],
  [/LOW[- ]VOLTAGE SUBSTATION|\bLVSS\b/,'Low-voltage substation'],
  [/\bREMOTE\s*I\/?O\b|\bRIO\b/,'Remote I/O panel'],
  [/PROGRAMMABLE LOGIC CONTROLLER|\bPLC\b/,'Programmable logic controller'],
  [/OPERATOR INTERFACE|\bOIT\b|\bHMI\b/,'Operator interface'],
  [/STRATIX|STARTIX|INDUSTRIAL NETWORK SWITCH|CISCO SWITCH/,'Industrial network switch'],
  [/PROFIBUS GATEWAY|NETWORK GATEWAY|SEGMENT COUPLER/,'Network gateway'],
  [/VARIABLE FREQUENCY DRIVE|ADJUSTABLE[- ]VARIABLE FREQUENCY DRIVE|\bVFD\b/,'Variable frequency drive'],
  [/FAULT CURRENT LIMITER/,'Fault current limiter'],
  [/\bTRANSFORMER\b/,'Transformer'],
  [/\bSWITCHGEAR\b|\bSWITCHBOARD\b/,'Switchgear'],
  [/\bCIRCUIT BREAKER\b/,'Circuit breaker'],
  [/\bDISCONNECT(?:OR)?(?: SWITCH)?\b/,'Disconnect switch'],
  [/\bDISTRIBUTION PANEL\b/,'Distribution panel'],
  [/\bBRANCH (?:CIRCUIT )?PANEL\b/,'Branch panel'],
  [/\bLIGHTING FIXTURE\b/,'Lighting fixture'],
  [/\bBATTERY CHARGER\b/,'Battery charger'],
  [/\bBATTERY (?:CABINET|RACK|SYSTEM)\b/,'Battery system'],
  [/\bUNINTERRUPTIBLE POWER SUPPLY\b|(?:^|\s)UPS(?:\s|$)|LIGHTING INVERTER/,'Uninterruptible power supply'],
  [/DIESEL (?:ENGINE )?GENERATOR|EMERGENCY GENERATOR/,'Emergency generator'],
  [/TEMPERATURE (?:ELEMENT )?(?:INDICATING )?TRANSMITTER/,'Temperature transmitter'],
  [/FLOW (?:INDICATING )?TRANSMITTER/,'Flow transmitter'],
  [/PRESSURE (?:DIFF(?:ERENTIAL)? )?(?:INDICATING )?TRANSMITTER/,'Pressure transmitter'],
  [/ANALY(?:ZER|TICAL) (?:INDICATING )?TRANSMITTER/,'Analyzer transmitter'],
  [/CONTROL VALVE|FLOW VALVE/,'Control valve'],
  [/SOLENOID VALVE|ON\/OFF VALVE/,'On/off valve'],
  [/\bPUMP\b/,'Pump'],
  [/\bLATERAL\b/,'Lateral'],
  [/MAIN DISTRIBUTION/,'Main distribution'],
  [/DISTRIBUTION PIPING|DISTRTIBUTION PIPING/,'Distribution piping'],
  [/\bHAND SWITCH\b/,'Hand switch'],
  [/POSITION SWITCH/,'Position switch'],
  [/CURRENT SWITCH/,'Current switch'],
  [/STATUS INDICATOR/,'Status indicator'],
  [/FIBER OPTIC DISTRIBUTION UNIT|\bFDU\b/,'Fiber distribution unit'],
];

function mapPush(map,key,value){const list=map.get(key)||[];list.push(value);map.set(key,list);}
function normalizedMapCount(map,key,amount=1){map.set(key,(map.get(key)||0)+amount);}
function displayRole(row){return clean(row&&row.equipmentDescription)||clean(row&&row.equipmentClassification)||EMPTY_ROLE;}
function roleKey(row){return auditNormId(displayRole(row));}
function isHeader(row){return /^HEADER$/i.test(clean(row&&row.equipmentDescription));}
function systemKey(row){
  const match=clean(row&&row.systemName).match(/^([0-9]{3,4})(?:\s|$)/);if(match)return match[1];
  return auditNormId(row&&row.upn)||'UNASSIGNED';
}
function stripLeadingSystemNumber(value){return clean(value).replace(/^\s*[0-9]{3,4}\s*[-:]?\s*/,'').replace(/\s+/g,' ').trim();}
function buildingNeutralTag(row){
  const id=auditNormId(row&&row.equipmentId),building=auditNormId(row&&row.building);if(!id||!building)return id;
  const aliases=[building,...building.match(/[A-Z0-9]+/g)||[]].filter(value=>value.length>1).sort((a,b)=>b.length-a.length);
  for(const alias of aliases){const escaped=alias.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),neutral=id.replace(new RegExp(`^${escaped}(?:\\s*[-_/:.]\\s*|\\s+)`),'');if(neutral!==id)return neutral;}
  return id;
}
function disciplineKind(row){
  const discipline=auditNormId(row&&row.discipline),upn=systemKey(row);
  if(upn==='650'||/FACILITIES MONITORING|INSTRUMENTATION|(?:^|\s)I\s*(?:&|AND)\s*C(?:\s|$)/.test(discipline))return 'Controls / I&C';
  return clean(row&&row.discipline)||'Unassigned discipline';
}
function isControls(row){
  if(disciplineKind(row)==='Controls / I&C')return true;
  const description=auditNormId(row&&row.equipmentDescription);
  return /REMOTE\s*I\/?O|\bRIO\b|\bPLC\b|PROGRAMMABLE LOGIC CONTROLLER|I\/?O CLUSTER|CONTROL PANEL|FMS (?:CABINET|CONTROLLER)|\bTRANSMITTER\b|\bSENSOR\b|\bINSTRUMENT\b|\bANALYZER\b|CONTROL VALVE|POSITION (?:SWITCH|INDICATOR)|LIMIT SWITCH|FLOW SWITCH|PRESSURE SWITCH|TEMPERATURE ELEMENT|HUMIDITY ELEMENT|STATUS INDICATOR|CURRENT SWITCH/.test(description);
}
function classificationFamily(row){const key=auditNormId(row&&row.equipmentClassification);return CLASSIFICATION_FAMILIES.get(key)||key;}
function semanticRole(row){
  if(isHeader(row))return 'Organizational header';
  const description=auditNormId(row&&row.equipmentDescription);for(const [pattern,label] of DESCRIPTION_FAMILIES)if(pattern.test(description))return label;
  const classification=classificationFamily(row);if(classification&&CLASSIFICATION_FAMILIES.has(auditNormId(row&&row.equipmentClassification)))return classification;
  return classification||displayRole(row);
}
function tagIdentityKey(row){
  const tag=buildingNeutralTag(row),upn=systemKey(row),tokens=tag.match(/\d+[A-Z]*/g)||[],useful=tokens.filter(token=>token!==upn);
  return useful.slice(-3).join('|');
}
function sortedMode(values){
  const counts=new Map(),display=new Map();for(const value of values){const key=auditNormId(value);if(!key)continue;normalizedMapCount(counts,key);if(!display.has(key))display.set(key,clean(value));}
  const winner=[...counts].sort((a,b)=>b[1]-a[1]||natCmp(a[0],b[0]))[0];return winner&&display.get(winner[0])||'';
}

export function auditRegistryModel(snapshot){
  const rows=snapshot&&snapshot.rows||[],rowsById=new Map(),childrenById=new Map(),generatedHeaders=new Map();
  for(const row of rows){const id=auditNormId(row.equipmentId);if(id&&!rowsById.has(id))rowsById.set(id,row);}
  for(const row of rows){
    const parent=auditNormId(row.closestParent),upn=systemKey(row);if(parent)mapPush(childrenById,parent,row);
    if(!parent||rowsById.has(parent)||auditNormId(row.closestParentStatus)!=='NEW'||parent===auditNormId(row.systemName))continue;
    const key=`${upn}|${parent}`;if(generatedHeaders.has(key))continue;
    generatedHeaders.set(key,{building:clean(row.building),discipline:clean(row.discipline),upn:clean(row.upn),systemName:clean(row.systemName),equipmentId:clean(row.closestParent),equipmentDescription:'HEADER',equipmentClassification:'HEADER',closestParent:clean(row.systemName),closestParentStatus:'New',dependencies:'',_source:row._source,_synthetic:true});
  }
  const nodeByRow=new Map(),systems=new Map();
  for(const row of [...rows,...generatedHeaders.values()]){
    const id=auditNormId(row.equipmentId),parentId=auditNormId(row.closestParent),upn=systemKey(row),parentRow=rowsById.get(parentId)||generatedHeaders.get(`${upn}|${parentId}`);
    const source=row._source||{},key=`${clean(source.sheet)}:${source.row||0}:${id}`;
    const semantic=semanticRole(row),parentSemantic=parentRow?semanticRole(parentRow):'System / external root';
    const node={row,key,id,upn,tag:clean(row.equipmentId),coreTag:buildingNeutralTag(row),identityKey:tagIdentityKey(row),role:displayRole(row),roleKey:roleKey(row),semanticRole:semantic,semanticRoleKey:auditNormId(semantic),classification:clean(row.equipmentClassification),classificationKey:auditNormId(row.equipmentClassification),classificationFamilyKey:auditNormId(classificationFamily(row)),discipline:clean(row.discipline),disciplineKind:disciplineKind(row),isControls:isControls(row),isHeader:isHeader(row),isSyntheticHeader:!!row._synthetic,parentId,parentRow,parentRole:parentRow?displayRole(parentRow):'System / external root',parentRoleKey:parentRow?roleKey(parentRow):'SYSTEM / EXTERNAL ROOT',parentSemanticRole:parentSemantic,parentSemanticRoleKey:auditNormId(parentSemantic),childCount:(childrenById.get(id)||[]).length,childSemanticRoleKey:'',ancestorSemanticPathKey:'',depth:0,underHeader:false,headerName:'',dependencyRoles:[],dependencyRoleKey:''};
    nodeByRow.set(row,node);mapPush(systems,upn,node);
  }
  for(const node of nodeByRow.values()){
    const seen=new Set([node.id]),roles=[],semanticRoles=[];let current=node.parentRow,depth=0;
    while(current&&depth<100){const parentNode=nodeByRow.get(current);if(!parentNode||seen.has(parentNode.id))break;seen.add(parentNode.id);depth++;roles.push(parentNode.role);semanticRoles.push(parentNode.semanticRoleKey);if(parentNode.isHeader&&!node.underHeader){node.underHeader=true;node.headerName=parentNode.tag;}current=parentNode.parentRow;}
    node.depth=depth;node.path=[...roles.reverse(),node.role];
    node.ancestorSemanticPathKey=semanticRoles.reverse().join('>');
    node.childSemanticRoleKey=(childrenById.get(node.id)||[]).map(child=>nodeByRow.get(child)).filter(Boolean).map(child=>child.semanticRoleKey).sort(natCmp).join(';');
    const dependencyRows=auditSplitReferences(node.row.dependencies).filter(value=>!/^N\/?A$/i.test(value)).map(reference=>rowsById.get(auditNormId(reference))||null);
    node.dependencyRoles=dependencyRows.map(linked=>linked?displayRole(linked):'External dependency').sort(natCmp);
    node.dependencyRoleKey=dependencyRows.map(linked=>auditNormId(linked?semanticRole(linked):'External dependency')).sort(natCmp).join(';');
  }
  return {rows,nodes:[...nodeByRow.values()],systems,rowsById};
}

function nodeSort(left,right){return natCmp(left.semanticRole,right.semanticRole)||natCmp(left.parentSemanticRole,right.parentSemanticRole)||natCmp(left.ancestorSemanticPathKey,right.ancestorSemanticPathKey)||natCmp(left.childSemanticRoleKey,right.childSemanticRoleKey)||natCmp(left.dependencyRoleKey,right.dependencyRoleKey)||natCmp(left.identityKey,right.identityKey)||natCmp(left.role,right.role)||natCmp(left.coreTag,right.coreTag)||natCmp(left.tag,right.tag);}
function pairNodes(targetNodes,referenceNodes){
  const target=[...targetNodes].sort(nodeSort),reference=[...referenceNodes].sort(nodeSort),targetOpen=new Set(target),referenceOpen=new Set(reference),pairs=[];
  const pairPass=(keyFor,reason)=>{
    const left=new Map(),right=new Map();
    for(const node of target)if(targetOpen.has(node)){const key=keyFor(node);if(key)mapPush(left,key,node);}
    for(const node of reference)if(referenceOpen.has(node)){const key=keyFor(node);if(key)mapPush(right,key,node);}
    for(const key of [...left.keys()].filter(value=>right.has(value)).sort(natCmp)){
      const targets=left.get(key).sort(nodeSort),references=right.get(key).sort(nodeSort),count=Math.min(targets.length,references.length);
      for(let index=0;index<count;index++){const targetNode=targets[index],referenceNode=references[index];targetOpen.delete(targetNode);referenceOpen.delete(referenceNode);pairs.push(comparePair(targetNode,referenceNode,reason));}
    }
  };
  pairPass(node=>node.coreTag&&`TAG|${node.coreTag}`,'Same equipment tag after Building removal');
  pairPass(node=>node.identityKey?`TYPE-PARENT-ID|${node.semanticRoleKey}|${node.parentSemanticRoleKey}|${node.identityKey}`:'','Same equipment type, parent type, and neutral tag identity');
  pairPass(node=>node.childSemanticRoleKey||node.dependencyRoleKey?`TYPE-CONTEXT|${node.semanticRoleKey}|${node.parentSemanticRoleKey}|${node.childSemanticRoleKey}|${node.dependencyRoleKey}|${Number(node.underHeader)}`:'','Same equipment type and hierarchy context');
  pairPass(node=>`FULL|${node.semanticRoleKey}|${node.classificationFamilyKey}|${node.parentSemanticRoleKey}|${Number(node.isHeader)}`,'Same equipment type and parent type');
  pairPass(node=>`TYPE-PARENT|${node.semanticRoleKey}|${node.parentSemanticRoleKey}`,'Same equipment type and parent type');
  pairPass(node=>node.ancestorSemanticPathKey?`TYPE-PATH|${node.semanticRoleKey}|${node.ancestorSemanticPathKey}`:'','Same equipment type and upstream equipment path');
  pairPass(node=>node.identityKey?`TYPE-ID|${node.semanticRoleKey}|${node.identityKey}`:'','Same equipment type and neutral tag identity');
  pairPass(node=>node.semanticRoleKey&&node.semanticRoleKey!==auditNormId(EMPTY_ROLE)?`TYPE|${node.semanticRoleKey}`:'','Same equipment type');
  pairPass(node=>node.classificationFamilyKey?`CLASS-PARENT|${node.classificationFamilyKey}|${node.parentSemanticRoleKey}`:'','Same classification family and parent type');
  pairPass(node=>node.classificationFamilyKey?`CLASS|${node.classificationFamilyKey}`:'','Same classification family');
  for(const node of [...targetOpen].sort(nodeSort))pairs.push(comparePair(node,null,'Only in target'));
  for(const node of [...referenceOpen].sort(nodeSort))pairs.push(comparePair(null,node,'Only in completed project'));
  const rank={changed:0,'target-only':1,'reference-only':2,aligned:3};
  return pairs.sort((a,b)=>rank[a.status]-rank[b.status]||nodeSort(a.target||a.reference,b.target||b.reference));
}
function comparePair(target,reference,matchReason){
  if(!reference)return {target,reference:null,status:'target-only',matchReason,differences:['Only in target']};
  if(!target)return {target:null,reference,status:'reference-only',matchReason,differences:['Only in completed project']};
  const differences=[];
  if(target.semanticRoleKey!==reference.semanticRoleKey)differences.push('Equipment type differs');
  if(target.parentSemanticRoleKey!==reference.parentSemanticRoleKey)differences.push('Parent equipment type differs');
  if(target.underHeader!==reference.underHeader)differences.push('Organizational header usage differs');
  if(target.disciplineKind!==reference.disciplineKind)differences.push('Discipline placement differs');
  if(target.dependencyRoleKey!==reference.dependencyRoleKey)differences.push('Dependency equipment types differ');
  return {target,reference,status:differences.length?'changed':'aligned',matchReason,differences};
}

function pairTree(pairs){
  const targetById=new Map(),referenceById=new Map();
  for(const pair of pairs){
    if(pair.target&&!targetById.has(pair.target.id))targetById.set(pair.target.id,pair);
    if(pair.reference&&!referenceById.has(pair.reference.id))referenceById.set(pair.reference.id,pair);
  }
  for(let index=0;index<pairs.length;index++){
    const pair=pairs[index],targetParent=pair.target&&targetById.get(pair.target.parentId),referenceParent=pair.reference&&referenceById.get(pair.reference.parentId);
    pair.id=`pair-${index}`;
    pair.targetParentPair=targetParent||null;
    pair.referenceParentPair=referenceParent||null;
  }
  const pairById=new Map(pairs.map(pair=>[pair.id,pair]));
  for(const pair of pairs){
    const targetParentId=pair.targetParentPair&&pair.targetParentPair.id||'',referenceParentId=pair.referenceParentPair&&pair.referenceParentPair.id||'';
    pair.targetParentPairId=targetParentId;
    pair.referenceParentPairId=referenceParentId;
    pair.parentPairId=pair.target?targetParentId:referenceParentId;
    if(pair.parentPairId===pair.id)pair.parentPairId='';
    pair.placementMismatch=!!(pair.target&&pair.reference&&targetParentId!==referenceParentId);
    pair.childrenIds=[];
    delete pair.targetParentPair;
    delete pair.referenceParentPair;
  }
  for(const pair of pairs){const parent=pairById.get(pair.parentPairId);if(parent)parent.childrenIds.push(pair.id);}
  const sortPairs=(leftId,rightId)=>nodeSort(pairById.get(leftId).target||pairById.get(leftId).reference,pairById.get(rightId).target||pairById.get(rightId).reference);
  for(const pair of pairs)pair.childrenIds.sort(sortPairs);
  const roots=pairs.filter(pair=>!pairById.has(pair.parentPairId)).sort((left,right)=>nodeSort(left.target||left.reference,right.target||right.reference)).map(pair=>pair.id);
  return {roots};
}

function nodeCounts(nodes,keyFor){const counts=new Map();for(const node of nodes)normalizedMapCount(counts,keyFor(node));return counts;}
function countObservations(type,title,left,right,subjectLabel){
  const observations=[],keys=new Set([...left.keys(),...right.keys()]);
  for(const key of [...keys].sort(natCmp)){const target=left.get(key)||0,reference=right.get(key)||0;if(target===reference)continue;observations.push({type,title,subject:subjectLabel(key),target,reference});}
  return observations;
}
function headerPlacementObservations(targetNodes,referenceNodes){
  const summarize=nodes=>{const groups=new Map();for(const node of nodes){if(node.isHeader)continue;const value=groups.get(node.semanticRoleKey)||{inside:0,outside:0};value[node.underHeader?'inside':'outside']++;groups.set(node.semanticRoleKey,value);}return groups;},target=summarize(targetNodes),reference=summarize(referenceNodes),observations=[];
  const state=value=>!value.inside?'Not under a header':!value.outside?'Under a header':'Mixed header placement';
  for(const role of [...target.keys()].filter(key=>reference.has(key)).sort(natCmp)){const targetState=state(target.get(role)),referenceState=state(reference.get(role));if(targetState===referenceState)continue;observations.push({type:'headers',title:'Organizational header use differs',subject:role,target:targetState,reference:referenceState});}
  return observations;
}
function compareSystem(upn,targetNodes=[],referenceNodes=[]){
  const pairs=pairNodes(targetNodes,referenceNodes),observations=[];
  const tree=pairTree(pairs);
  const targetActual=targetNodes.filter(node=>!node.isSyntheticHeader),referenceActual=referenceNodes.filter(node=>!node.isSyntheticHeader);
  const targetNames=targetActual.map(node=>node.row.systemName),referenceNames=referenceActual.map(node=>node.row.systemName),targetName=sortedMode(targetNames),referenceName=sortedMode(referenceNames);
  if(targetActual.length&&!referenceActual.length)observations.push({type:'coverage',title:'System occurs only in target',subject:`UPN ${upn}`,target:targetActual.length,reference:0});
  else if(referenceActual.length&&!targetActual.length)observations.push({type:'coverage',title:'System occurs only in completed project',subject:`UPN ${upn}`,target:0,reference:referenceActual.length});
  if(targetName&&referenceName&&auditNormId(stripLeadingSystemNumber(targetName))!==auditNormId(stripLeadingSystemNumber(referenceName)))observations.push({type:'system',title:'System descriptions differ',subject:`${targetName} / ${referenceName}`,target:targetName,reference:referenceName});
  observations.push(...countObservations('role','Equipment mix differs',nodeCounts(targetNodes,node=>node.semanticRoleKey),nodeCounts(referenceNodes,node=>node.semanticRoleKey),key=>key));
  observations.push(...countObservations('hierarchy','Hierarchy pattern differs',nodeCounts(targetNodes,node=>`${node.parentSemanticRoleKey} > ${node.semanticRoleKey}`),nodeCounts(referenceNodes,node=>`${node.parentSemanticRoleKey} > ${node.semanticRoleKey}`),key=>key));
  observations.push(...headerPlacementObservations(targetNodes,referenceNodes));
  observations.push(...countObservations('controls','I&C nesting differs',nodeCounts(targetNodes.filter(node=>node.isControls),node=>`${node.semanticRoleKey} > ${node.parentSemanticRoleKey}`),nodeCounts(referenceNodes.filter(node=>node.isControls),node=>`${node.semanticRoleKey} > ${node.parentSemanticRoleKey}`),key=>key));
  const pairSummary={aligned:0,changed:0,targetOnly:0,referenceOnly:0};
  for(const pair of pairs){if(pair.status==='target-only')pairSummary.targetOnly++;else if(pair.status==='reference-only')pairSummary.referenceOnly++;else pairSummary[pair.status]++;}
  const targetHeaders=targetNodes.filter(node=>node.isHeader).length,referenceHeaders=referenceNodes.filter(node=>node.isHeader).length;
  return {upn,label:targetName||referenceName||(upn==='UNASSIGNED'?'Unassigned UPN':`UPN ${upn}`),targetName,referenceName,targetRows:targetActual.length,referenceRows:referenceActual.length,targetHeaders,referenceHeaders,targetControls:targetActual.filter(node=>node.isControls).length,referenceControls:referenceActual.filter(node=>node.isControls).length,pairs,treeRoots:tree.roots,observations,pairSummary,status:!targetActual.length?'reference-only':!referenceActual.length?'target-only':observations.length||pairSummary.changed||pairSummary.targetOnly||pairSummary.referenceOnly?'different':'aligned',differenceCount:observations.length+pairSummary.changed+pairSummary.targetOnly+pairSummary.referenceOnly};
}

export function compareSsmRegistries(targetSnapshot,referenceSnapshot){
  const target=auditRegistryModel(targetSnapshot),reference=auditRegistryModel(referenceSnapshot),upns=new Set([...target.systems.keys(),...reference.systems.keys()]);
  const systems=[...upns].sort(natCmp).map(upn=>compareSystem(upn,target.systems.get(upn)||[],reference.systems.get(upn)||[]));
  const summary={systems:systems.length,alignedSystems:0,differentSystems:0,targetOnlySystems:0,referenceOnlySystems:0,targetRows:target.rows.length,referenceRows:reference.rows.length,alignedRows:0,changedRows:0,targetOnlyRows:0,referenceOnlyRows:0,observations:0};
  for(const system of systems){if(system.status==='aligned')summary.alignedSystems++;else if(system.status==='target-only')summary.targetOnlySystems++;else if(system.status==='reference-only')summary.referenceOnlySystems++;else summary.differentSystems++;summary.alignedRows+=system.pairSummary.aligned;summary.changedRows+=system.pairSummary.changed;summary.targetOnlyRows+=system.pairSummary.targetOnly;summary.referenceOnlyRows+=system.pairSummary.referenceOnly;summary.observations+=system.observations.length;}
  return Object.freeze({schemaVersion:SSM_COMPARISON_SCHEMA_VERSION,standard:'Building-neutral project comparison',summary:Object.freeze(summary),systems:Object.freeze(systems)});
}

export function comparisonSystemTypes(system){
  const counts={coverage:0,system:0,role:0,hierarchy:0,headers:0,controls:0};for(const item of system&&system.observations||[])counts[item.type]=(counts[item.type]||0)+1;return counts;
}
