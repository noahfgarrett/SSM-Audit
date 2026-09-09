import { clean } from '../core/text.js'
import { auditFingerprint } from './model.js'

// Session-only reference records. All helpers are prefixed for the single-scope bundle.
const AUDIT_REF_HEADER_LIMIT=200;
const AUDIT_REF_HEADERS={
  milestones:{
    id:['L2 ID','L2 Milestone ID','Milestone ID','Milestone Code','MS','MS Code','Milestone #','Milestone No','Milestone Number','L2 Milestone #','L2 MS #','L2 Milestone','Milestone'],
    title:['L2 Milestone Title','L2 Milestone Name','L2 Milestone Description','Milestone Title','Milestone Name','Milestone Description','MS Description','L2 Title','L2 Name','L2 Description','Title','Description'],
    parentId:['L1 ID','L1 MS #','L1 Milestone ID','L1 Milestone #','Milestone Parent','L1 Milestone Parent','L1 Milestone','Parent Milestone ID','Parent ID'],
    parentTitle:['L1 Milestone Title','L1 Milestone Name','L1 Milestone Description','L1 Milestone Parent Title','L1 MS Title','L1 MS Name','L1 MS Description','L1 MS / Driver','L1 Title','L1 Name','L1 Description','Parent Milestone Title','Parent Milestone Name','Milestone Parent Title','Milestone Parent Name','Milestone Parent Description','Parent Title','Parent Description'],
    phase:['Milestone Phase','Phase'],
  },
  itemMasters:{
    name:['Item Master Name','Item Master','Item Master Unique Identifier','Item Master ID','IM Name'],
    id:['Item Master Unique Identifier','Item Master ID','Item Master Name','Item Master','IM Name'],
    discipline:['Discipline','Item Master Discipline'],
  },
};
const AUDIT_REF_ROW_HEADERS={status:['Record Status','Lifecycle Status','Status','State'],rowType:['Record Type','Row Type','Entry Type']};

function auditRefText(value){return clean(value).replace(/\s+/g,' ');}
function auditRefValue(value){const text=auditRefText(value);return /^(?:N\/?A|NOT APPLICABLE|TBD|TBC|NONE|NULL|-+)$/i.test(text)?'':text;}
function auditRefNorm(value){return auditRefText(value).toUpperCase();}
function auditRefHeaderKey(value){return auditRefNorm(value).replace(/#/g,' NUMBER ').replace(/[^A-Z0-9]/g,'');}
function auditRefKind(kind){return kind==='milestones'||kind==='itemMasters';}
function auditRefExcludedSheet(name){
  const words=clean(name).replace(/([a-z])([A-Z])/g,'$1 $2').replace(/[^A-Za-z0-9]+/g,' ');
  return /\b(?:history|historical|changes?|changelog|revisions?|legacy|archives?|archived|obsolete|old|deprecated|superseded|retired|comments?|notes?)\b/i.test(words);
}
function auditRefHistorical(value){return /^(?:history|historical|change(?:s|\s*log)?|archive[ds]?|obsolete|old|deprecated|superseded|retired|inactive|deleted|removed)(?:$|\b[\s:(-])/i.test(auditRefText(value));}
function auditRefComment(value){return /^(?:#|\/\/|(?:comments?|notes?)(?:$|\b[\s:(-]))/i.test(auditRefText(value));}
function auditRefFirstCell(row){return (Array.isArray(row)?row:[]).find(value=>auditRefText(value))||'';}
function auditRefColumns(row,kind){
  if(!Array.isArray(row))return null;
  const keys=row.map(auditRefHeaderKey),fields={};
  for(const [field,aliases] of Object.entries({...AUDIT_REF_HEADERS[kind],...AUDIT_REF_ROW_HEADERS})){
    for(const alias of aliases){const index=keys.indexOf(auditRefHeaderKey(alias));if(index!==-1){fields[field]=index;break;}}
  }
  return fields[kind==='milestones'?'id':'name']==null?null:fields;
}
function auditRefDetectHeader(aoa,kind){
  let best=null;
  for(let index=0;index<Math.min(AUDIT_REF_HEADER_LIMIT,aoa.length);index++){
    const row=aoa[index],first=auditRefFirstCell(row);
    if(auditRefHistorical(first))break;
    if(auditRefComment(first))continue;
    const fields=auditRefColumns(row,kind);if(!fields)continue;
    const score=new Set(Object.values(fields)).size;
    if(!best||score>best.score)best={index,fields,score};
  }
  return best;
}
function auditRefCellText(cell){
  if(cell==null)return '';
  if(typeof cell!=='object')return auditRefText(cell);
  if(cell.t==='e')return '';
  if(cell.w!=null)return auditRefText(cell.w);
  if(cell.v==null)return '';
  // Format a copy: SheetJS may cache display text on the cell it receives.
  return auditRefText(typeof XLSX!=='undefined'&&cell.t?XLSX.utils.format_cell({...cell}):cell.v);
}
function auditRefSheetAoa(sheet,maxRows=Infinity){
  const dense=Array.isArray(sheet)?sheet:sheet&&Array.isArray(sheet['!data'])?sheet['!data']:null;
  if(dense)return dense.slice(0,maxRows).map(row=>Array.isArray(row)?row.map(auditRefCellText):[]);
  if(!sheet||!sheet['!ref'])return [];
  if(typeof XLSX==='undefined')throw new Error('The local workbook reader is unavailable.');
  const range=XLSX.utils.decode_range(sheet['!ref']),aoa=[];
  // Visit populated cells, not a potentially huge formatted range. Keep physical
  // offsets so blank rows still count toward the header-search limit.
  for(const address of Object.keys(sheet)){
    if(!/^[A-Z]+[1-9][0-9]*$/.test(address))continue;
    const {r,c}=XLSX.utils.decode_cell(address);
    if(r<range.s.r||r>range.e.r||r>=maxRows||c<range.s.c||c>range.e.c)continue;
    const value=auditRefCellText(sheet[address]);if(!value)continue;
    if(!aoa[r])aoa[r]=[];
    aoa[r][c]=value;
  }
  return aoa;
}
function auditRefResult(kind,sheetName,entries=[],warning=''){
  return {kind,sheetName,entries,...(warning?{warning}:{})};
}
function auditRefJoin(id,title){
  if(!id)return title;
  if(!title||auditRefNorm(id)===auditRefNorm(title)||auditRefNorm(id).endsWith(` ${auditRefNorm(title)}`))return id;
  return `${id} ${title}`;
}
function auditRefMilestoneParts(value){
  const text=auditRefNorm(value).replace(/\b(L[12])\s+(?=(?:M\d+|\d+)\b)/g,'$1-')
    .replace(/\b(L[12]-M\d+)\s+(?:-\s*)?(?=\d+\b)/g,'$1-');
  // Read a compact code before accepting spaced code separators, so a display
  // separator does not become part of the ID. Namespace and suffixes stay intact.
  const match=text.match(/^((?:[A-Z0-9][A-Z0-9.]*[-_])*(L[12])(?:[-_][A-Z0-9.]+)+)(?:\s+|[:|]\s*|$)(.*)$/)
    ||text.match(/^((?:[A-Z0-9][A-Z0-9.]*\s*[-_]\s*)*(L[12])(?:\s*[-_]\s*[A-Z0-9.]+)+)(?:\s+|[:|]\s*|$)(.*)$/);
  return match?{code:match[1].replace(/\s+/g,''),title:match[3].replace(/^[-:|\u2013\u2014]\s*/,''),level:match[2]}:null;
}
function auditRefHasMilestoneCode(value){return /(?:^|[^A-Z0-9])L[12](?:\s*[-_]\s*[A-Z0-9]|\s+(?:M\d+|\d+)\b)/i.test(auditRefText(value));}
function auditRefMilestoneKey(value){
  const parts=auditRefMilestoneParts(value);
  return parts?auditRefJoin(parts.code,parts.title):auditRefNorm(value);
}
function auditRefMilestoneAliases(entry,parent=false){
  const id=parent?entry.parentId:entry.id,title=parent?entry.parentTitle:entry.title,parts=auditRefMilestoneParts(id);
  return [id,title,auditRefJoin(id,title),parts&&parts.code,parts&&parts.title].filter(Boolean).map(auditRefMilestoneKey);
}
function auditRefEntryKey(entry,kind){
  const fields=kind==='milestones'?['id','title','parentId','parentTitle','phase']:['id','name','discipline'];
  return JSON.stringify(fields.map(field=>auditRefNorm(entry[field])));
}
function auditRefUsable(reference,kind){return reference&&reference.kind===kind&&Array.isArray(reference.entries)&&reference.entries.length>0;}

/** Parse a caller-selected AOA. Milestones expose id/title/parentId/parentTitle,
 * milestone/milestoneParent display values, phase, and the one-based sourceRow.
 * Item Masters expose id/name/discipline and sourceRow. No inputs are mutated. */
export function auditReadReferenceAoa(aoa,kind,sheetName){
  if(!auditRefKind(kind))return auditRefResult(kind,sheetName,[],'Unsupported reference kind.');
  if(typeof sheetName!=='string'||!sheetName.trim())return auditRefResult(kind,sheetName,[],'Select a reference sheet explicitly.');
  if(auditRefExcludedSheet(sheetName))return auditRefResult(kind,sheetName,[],'History, change, archive, obsolete, and comment sheets cannot be references.');
  if(!Array.isArray(aoa))return auditRefResult(kind,sheetName,[],'The selected sheet could not be read.');
  const detected=auditRefDetectHeader(aoa,kind);
  if(!detected)return auditRefResult(kind,sheetName,[],'No relevant reference header was found in the first 200 rows.');
  const entries=[],seen=new Set(),warnings=[],fields=detected.fields;
  if(kind==='milestones'&&fields.title==null)warnings.push('No separate milestone title/description column was found.');
  if(kind==='milestones'&&fields.parentId==null)warnings.push('No L1 parent identifier column was found; parent recommendations are unavailable.');
  for(let index=detected.index+1;index<aoa.length;index++){
    const cells=aoa[index];if(!Array.isArray(cells))continue;
    const get=field=>fields[field]==null?'':auditRefText(cells[fields[field]]),first=auditRefFirstCell(cells);
    // A historical appendix is not part of the selected current table.
    if(auditRefHistorical(first)&&(cells.filter(value=>auditRefText(value)).length===1||/^(?:history|historical|change\s*log|archive)(?:$|\b[\s:(-])/i.test(auditRefText(first))))break;
    if(cells.filter(value=>auditRefText(value)).length===1&&/^(?:comments|notes)\s*:?$/i.test(auditRefText(first)))break;
    if(auditRefHistorical(first)||auditRefComment(first)||auditRefHistorical(get('status'))||auditRefHistorical(get('rowType'))||auditRefComment(get('rowType'))||auditRefColumns(cells,kind))continue;
    let entry;
    if(kind==='milestones'){
      const id=auditRefValue(get('id')),title=get('title'),parentId=auditRefValue(get('parentId')),parentTitle=get('parentTitle');
      if(!id||auditRefHistorical(id)||auditRefComment(id)||auditRefMilestoneParts(id)?.level==='L1')continue;
      entry={id,title,parentId,parentTitle,phase:get('phase'),milestone:auditRefJoin(id,title),milestoneParent:auditRefJoin(parentId,parentTitle),sourceRow:index+1};
    }else{
      const name=auditRefValue(get('name'));if(!name||auditRefHistorical(name)||auditRefComment(name))continue;
      entry={id:get('id'),name,discipline:get('discipline'),sourceRow:index+1};
    }
    const key=auditRefEntryKey(entry,kind);if(seen.has(key))continue;
    seen.add(key);entries.push(entry);
  }
  if(!entries.length)warnings.push('No current reference entries were found in the selected sheet.');
  return auditRefResult(kind,sheetName,entries,warnings.join(' '));
}

/** Sheet discovery is advisory only; it never chooses a sheet for the caller. */
export function auditReferenceSheets(workbook,kind){
  if(!auditRefKind(kind)||!Array.isArray(workbook?.SheetNames))return [];
  return [...new Set(workbook.SheetNames)].filter(name=>{
    if(typeof name!=='string'||auditRefExcludedSheet(name)||!Object.hasOwn(workbook.Sheets||{},name))return false;
    try{return !!auditRefDetectHeader(auditRefSheetAoa(workbook.Sheets[name],AUDIT_REF_HEADER_LIMIT),kind);}catch{return false;}
  });
}

export function auditReadReferenceWorkbook(workbook,kind,sheetName){
  if(!auditRefKind(kind)||typeof sheetName!=='string'||!sheetName.trim()||auditRefExcludedSheet(sheetName))return auditReadReferenceAoa([],kind,sheetName);
  if(!Array.isArray(workbook?.SheetNames)||!workbook.SheetNames.includes(sheetName)||!Object.hasOwn(workbook.Sheets||{},sheetName))return auditRefResult(kind,sheetName,[],'The explicitly selected reference sheet is unavailable.');
  try{return auditReadReferenceAoa(auditRefSheetAoa(workbook.Sheets[sheetName]),kind,sheetName);}
  catch{return auditRefResult(kind,sheetName,[],'The selected reference sheet could not be read locally.');}
}

function auditRefMilestoneIndex(reference){
  if(!auditRefUsable(reference,'milestones'))return null;
  const codes=new Map(),names=new Map();
  const add=(index,key,entry)=>{
    const candidates=index.get(key)||new Map();candidates.set(auditRefEntryKey(entry,'milestones'),entry);index.set(key,candidates);
  };
  for(const entry of reference.entries){
    if(!entry||!auditRefValue(entry.id))continue;
    const parts=auditRefMilestoneParts(entry.id);if(parts?.level==='L1')continue;
    if(parts)add(codes,parts.code,entry);
    for(const alias of auditRefMilestoneAliases(entry))add(names,alias,entry);
  }
  return names.size?{codes,names}:null;
}
function auditRefMatchMilestone(row,index){
  const value=auditRefMilestoneKey(auditRefValue(row?.milestone));
  if(!value||!index)return {status:'unverified'};
  const parts=auditRefMilestoneParts(value);
  if(!parts&&auditRefHasMilestoneCode(value))return {status:'missing'};
  let candidates=parts?index.codes.get(parts.code):index.names.get(value);
  if(!candidates)return {status:'missing'};
  // A title can narrow duplicate full codes, never replace an explicit code.
  if(parts?.title&&candidates.size>1){
    const names=index.names.get(value),exact=new Map([...candidates].filter(([key])=>names?.has(key)));
    if(exact.size)candidates=exact;
  }
  if(candidates.size>1)return {status:'ambiguous'};
  return {status:'matched',entry:candidates.values().next().value};
}
/** This verifies the named L2, not whether that milestone suits the equipment.
 * UPN, equipment descriptions, and the existing (possibly wrong) L1 are not evidence. */
export function auditMilestoneReferenceMatch(row,reference){
  return auditRefMatchMilestone(row,auditRefMilestoneIndex(reference));
}

function auditRefParentIdentity(entry){return auditRefMilestoneParts(entry.parentId)?.code||auditRefMilestoneKey(entry.parentId);}
function auditRefParentIndex(reference){
  const index=new Map();
  for(const entry of reference.entries){
    if(!entry||!auditRefValue(entry.parentId)||auditRefMilestoneParts(entry.parentId)?.level==='L2')continue;
    const identity=auditRefParentIdentity(entry);
    for(const alias of auditRefMilestoneAliases(entry,true)){
      const identities=index.get(alias)||new Set();identities.add(identity);index.set(alias,identities);
    }
  }
  return index;
}
function auditRefParentMatches(current,entry,parentIndex){
  const actual=auditRefMilestoneParts(current),expected=auditRefMilestoneParts(entry.parentId);
  if(actual)return !!expected&&actual.code===expected.code;
  if(auditRefHasMilestoneCode(current))return false;
  const identities=parentIndex.get(auditRefMilestoneKey(current));
  return !!identities&&identities.size===1&&identities.has(auditRefParentIdentity(entry));
}

function auditRefItemKey(value){return auditRefNorm(value).replace(/\s*([_/&,])\s*/g,'$1');}
function auditRefItemBody(value,site=''){
  const name=auditRefItemKey(value),siteKey=auditRefItemKey(site);
  // An explicit site field can delimit a multi-token prefix. Never trim arbitrary
  // suffixes: ELEC, I&C, W/O, etc. must remain part of the functional body.
  if(!/^VF\d*_/.test(name)&&siteKey&&name.startsWith(`${siteKey}_`))return name.slice(siteKey.length+1);
  const match=name.match(/^[A-Z0-9][A-Z0-9.-]*_(.+)$/);
  return match?match[1]:'';
}

/** Confidence is 'verified' for register identity/parent relationships and 'review'
 * for catalog equivalents. Literal catalog membership never proves compatibility. */
export function auditReferenceRecommendation(row,field,references){
  const key=auditRefHeaderKey(field);
  if(['ITEMMASTER','ITEMMASTERNAME','ITEMMASTERID','ITEMMASTERUNIQUEIDENTIFIER'].includes(key)){
    const reference=references?.itemMasters,value=auditRefNorm(auditRefValue(row?.itemMaster));
    if(!value||!auditRefUsable(reference,'itemMasters'))return null;
    const entries=reference.entries.filter(entry=>entry&&auditRefText(entry.name));
    if(entries.some(entry=>[entry.name,entry.id].some(name=>auditRefNorm(name)===value)))return null;
    const body=auditRefItemBody(row.itemMaster,row.site);if(!body)return null;
    const candidates=new Map();
    for(const entry of entries){
      if(/^VF\d*_/.test(auditRefItemKey(entry.name))&&auditRefItemBody(entry.name)===body)candidates.set(auditRefEntryKey(entry,'itemMasters'),entry);
    }
    if(candidates.size!==1)return null;
    return {value:candidates.values().next().value.name,reason:'The selected current catalog contains one equivalent name with the same complete functional body after site-prefix normalization. The existing value is not a literal catalog member; that alone does not make a legacy prefix invalid. Equipment compatibility is unverified and requires review.',confidence:'review'};
  }
  const parent=['MILESTONEPARENT','L1MILESTONEPARENT','L1ID','L1MILESTONE'].includes(key),milestone=['MILESTONE','L2MILESTONE','MILESTONEID','L2ID'].includes(key);
  if(!parent&&!milestone)return null;
  const match=auditMilestoneReferenceMatch(row,references?.milestones);if(match.status!=='matched')return null;
  const entry=match.entry;
  if(parent&&(!auditRefValue(entry.parentId)||auditRefMilestoneParts(entry.parentId)?.level==='L2'))return null;
  const current=parent?row.milestoneParent:row.milestone,value=parent?auditRefJoin(entry.parentId,entry.parentTitle):auditRefJoin(entry.id,entry.title);
  if(!value||(parent?auditRefParentMatches(current,entry,auditRefParentIndex(references.milestones)):auditRefMilestoneKey(current)===auditRefMilestoneKey(value)))return null;
  return {value,reason:parent?'The selected authoritative register explicitly maps this uniquely matched L2 milestone to this L1 parent; no relationship was inferred from UPN.':'This is the uniquely matched milestone in the selected authoritative register; equipment applicability is not inferred.',confidence:'verified'};
}

function auditRefRule(id,title,statement,confidence='strong'){
  return Object.freeze({id,version:1,source:'reference',category:'milestones',title,statement,standardRef:title,confidence,enabled:true,disabledReason:''});
}
export const SSM_AUDIT_REFERENCE_RULES=Object.freeze({
  milestoneUnknown:auditRefRule('reference.milestone-unknown','Milestone not found in the selected register','A milestone not found in the selected register needs confirmation; absence alone does not prove an incorrect assignment.'),
  milestoneAmbiguous:auditRefRule('reference.milestone-ambiguous','Milestone matches more than one register entry','The milestone must identify one register entry before its parent can be verified.'),
  milestoneParentMissing:auditRefRule('reference.milestone-parent-missing','Milestone Parent is missing','A uniquely verified L2 milestone requires the L1 parent explicitly listed in the selected register.','required'),
  milestoneParentMismatch:auditRefRule('reference.milestone-parent-mismatch','Milestone Parent differs from the register','The L1 parent must agree with the explicit parent of the uniquely verified L2 milestone in the selected register.','required'),
});
function auditRefFinding(rule,severity,row,details){
  const equipmentId=clean(row?.equipmentId),source=row?._source||{},actual=details.actual||'',expected=details.expected||'';
  const fingerprint=auditFingerprint([rule.id,auditRefNorm(equipmentId),source.sheet||'',source.row||'',JSON.stringify(actual),JSON.stringify(expected)].join('|'));
  return Object.freeze({schemaVersion:1,id:`${rule.id}:${fingerprint}`,fingerprint,rule,severity,category:rule.category,
    equipmentId,row:source.row||0,sheet:source.sheet||'',field:details.field,why:details.why,actual,expected,
    recommendation:details.recommendation||'',relatedEquipmentId:'',relationship:null,
    searchKey:auditRefNorm([rule.id,rule.source,rule.title,rule.category,equipmentId,source.sheet,details.field,details.why,actual,expected].join(' '))});
}
function auditRefParentContradiction(current,entry,parentIndex){
  const actual=auditRefMilestoneParts(current),expected=auditRefMilestoneParts(entry.parentId);
  if(actual&&expected)return actual.code!==expected.code;
  if(!actual&&auditRefHasMilestoneCode(current))return false;
  // A different literal parent known to this register is also positive evidence.
  const identities=parentIndex.get(actual?actual.code:auditRefMilestoneKey(current));
  return !!identities&&!identities.has(auditRefParentIdentity(entry));
}

/** Optional, engine-shaped findings only. No catalog membership or migration rule
 * is duplicated here, and neither counts nor snapshot state are modified. */
export function auditReferenceFindings(snapshot,references){
  const reference=references?.milestones,index=auditRefMilestoneIndex(reference);
  if(!index||!Array.isArray(snapshot?.rows))return [];
  const findings=[],parentIndex=auditRefParentIndex(reference);
  for(const row of snapshot.rows){
    const match=auditRefMatchMilestone(row,index);
    if(match.status==='unverified')continue;
    if(match.status==='missing'||match.status==='ambiguous'){
      const ambiguous=match.status==='ambiguous';
      findings.push(auditRefFinding(ambiguous?SSM_AUDIT_REFERENCE_RULES.milestoneAmbiguous:SSM_AUDIT_REFERENCE_RULES.milestoneUnknown,'info',row,{
        field:'L2 Milestone',actual:clean(row.milestone),expected:'One exact milestone identifier or name in the selected register',
        why:ambiguous?'More than one register entry matches this milestone. Its L1 parent cannot be verified.':'No exact identifier or name matches this milestone in the selected register. This does not establish that the assignment is wrong.',
        recommendation:'Confirm the governing milestone and selected register; include the full milestone name when a code is shared by different phases.'}));
      continue;
    }
    const entry=match.entry,current=auditRefValue(row.milestoneParent);
    if(!auditRefValue(entry.parentId)||auditRefMilestoneParts(entry.parentId)?.level==='L2'||auditRefParentMatches(current,entry,parentIndex))continue;
    const missing=!current,contradiction=!missing&&auditRefParentContradiction(current,entry,parentIndex),expected=auditRefJoin(entry.parentId,entry.parentTitle);
    findings.push(auditRefFinding(missing?SSM_AUDIT_REFERENCE_RULES.milestoneParentMissing:SSM_AUDIT_REFERENCE_RULES.milestoneParentMismatch,missing||contradiction?'error':'info',row,{
      field:'L1 Milestone Parent',actual:clean(row.milestoneParent),expected,
      why:missing?'This L2 milestone is uniquely verified, but its L1 parent is blank.':contradiction?'This L2 milestone is uniquely verified, and its assigned L1 contradicts the parent explicitly recorded in the selected register.':'This L2 milestone is uniquely verified, but the L1 text cannot be verified as its recorded parent. Confirm any local alias before changing it.',
      recommendation:'Confirm the L2 assignment, then use its explicit L1 parent from the selected register.'}));
  }
  return findings;
}
