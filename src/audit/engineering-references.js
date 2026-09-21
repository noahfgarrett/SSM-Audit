import { clean } from '../core/text.js'
import { auditNormId, auditFingerprint, auditSplitReferences } from './model.js'
import { auditRefSheetAoa } from './references.js'
import { auditIsBlankItemMaster, SSM_AUDIT_SEVERITIES } from './engine.js'

export const AUDIT_FINDING_LEVELS=Object.freeze([...SSM_AUDIT_SEVERITIES.slice(0,2),'missing',...SSM_AUDIT_SEVERITIES.slice(2)]);

export const AUDIT_ENGINEERING_KINDS=Object.freeze({mel:'MEL',easyPower:'EasyPower',cable:'Cable Schedule',pmd:'PMD'});
const AUDIT_ENGINEERING_FIELDS=Object.freeze({upn:'UPN',building:'Building',discipline:'Discipline',systemName:'System Name',equipmentDescription:'Equipment Description',equipmentClassification:'Equipment Classification'});
const AUDIT_ENGINEERING_CONTEXT_FIELDS=Object.freeze(['building','discipline','upn','systemName','equipmentDescription','equipmentClassification','closestParent']);
const AUDIT_ENGINEERING_TAG_PLACEHOLDERS=new Set(['NA','N/A']);
const AUDIT_ENGINEERING_ALIASES={
  equipmentId:['equipmenttag','equipmentid','tag','tagname'],upn:['upn','upncode'],building:['bldg','building','buildingname'],
  discipline:['discipline'],systemName:['systemname'],systemDescription:['systemdescription'],equipmentDescription:['equipmentdescription','description'],
  equipmentClassification:['equipmentclassification'],closestParent:['systemparentequipmenttags','systemparentequipmenttag'],
  panel:['panelfrom','frompanel','panel'],load:['loadnameto','loadname','toequipment'],instrument:['instrumenttag','instrumenttagno','instrumenttagname'],
  start:['startingsource','startsource'],id:['idname'],final:['finalsource'],
};
function auditEngineeringTag(value){const tag=clean(value);return AUDIT_ENGINEERING_TAG_PLACEHOLDERS.has(auditNormId(tag))?'':tag;}
function auditEngineeringTagList(value){return auditSplitReferences(value).map(auditEngineeringTag).filter(Boolean);}
function auditEngineeringParent(value){return auditEngineeringTagList(value).join('; ');}
function auditEngineeringContextValue(field,entry){
  const value=field==='closestParent'?(entry.closestParent??entry.parent):entry[field];
  return field==='closestParent'?auditEngineeringParent(value):clean(value);
}
function auditEngineeringSourceContext(entries){
  const context={};
  for(const field of AUDIT_ENGINEERING_CONTEXT_FIELDS){
    const choices=new Map();
    for(const entry of entries){const value=auditEngineeringContextValue(field,entry);if(value&&!choices.has(auditNormId(value)))choices.set(auditNormId(value),value);}
    context[field]=choices.size===1?[...choices.values()][0]:'';
  }
  return context;
}
function auditEngineeringHeader(aoa,kind){
  for(let row=0;row<Math.min(200,aoa.length);row++){
    const fields={},downstream=[];
    (aoa[row]||[]).forEach((value,column)=>{
      const key=auditNormId(value).replace(/[^A-Z0-9]/g,'').toLowerCase();
      for(const [field,aliases] of Object.entries(AUDIT_ENGINEERING_ALIASES))if(aliases.includes(key)&&fields[field]==null)fields[field]=column;
      const match=key.match(/^(?:downstream|ds)0*(\d+)$/);if(match)downstream.push({column,level:Number(match[1])});
    });
    const valid=kind==='mel'?fields.equipmentId!=null&&Object.keys(AUDIT_ENGINEERING_FIELDS).some(field=>fields[field]!=null):kind==='cable'?fields.panel!=null&&fields.load!=null:kind==='pmd'?fields.panel!=null&&fields.instrument!=null:kind==='easyPower'?fields.start!=null&&fields.id!=null:false;
    if(valid)return {row,fields,downstream:downstream.sort((a,b)=>a.level-b.level)};
  }
  return null;
}
export function auditReadEngineeringAoa(aoa,kind,sheetName=''){
  const header=auditEngineeringHeader(aoa,kind);if(!header)return null;
  const entries=[];
  for(let index=header.row+1;index<aoa.length;index++){
    const values=aoa[index]||[],read=field=>clean(values[header.fields[field]]),sourceRow=index+1;
    if(kind==='mel'){
      const equipmentId=auditEngineeringTag(read('equipmentId'));if(!equipmentId)continue;
      const entry={equipmentId,sourceRow,sheetName};
      for(const field of Object.keys(AUDIT_ENGINEERING_FIELDS))entry[field]=read(field);
      const description=read('systemDescription');
      if(!entry.systemName&&description&&entry.upn)entry.systemName=auditNormId(description).startsWith(`${auditNormId(entry.upn)} `)?description:`${entry.upn} ${description}`;
      entry.closestParent=auditEngineeringParent(read('closestParent'));entries.push(entry);
    }else if(kind==='easyPower'){
      // Preserve gaps in the export without dropping everything after a blank.
      const chain=[read('start'),...header.downstream.map(item=>clean(values[item.column])),read('final'),read('id')].map(auditEngineeringTag).filter(Boolean);
      chain.forEach((equipmentId,i)=>{if(!i||auditNormId(equipmentId)!==auditNormId(chain[i-1]))entries.push({equipmentId,parent:i?chain[i-1]:'',sourceRow,sheetName});});
    }else{
      const parents=auditEngineeringTagList(read('panel')),children=auditEngineeringTagList(read(kind==='pmd'?'instrument':'load'));
      for(const equipmentId of children){
        if(parents.length)for(const parent of parents)entries.push({equipmentId,parent,sourceRow,sheetName});
        else entries.push({equipmentId,parent:'',sourceRow,sheetName});
      }
      for(const equipmentId of parents)entries.push({equipmentId,parent:'',sourceRow,sheetName});
    }
  }
  return {kind,sheetName,headerRow:header.row+1,entries};
}
export function auditReadEngineeringWorkbook(workbook,kind){
  if(!Object.hasOwn(AUDIT_ENGINEERING_KINDS,kind))return [];
  const references=[];
  for(const name of workbook.SheetNames||[]){
    const sheet=workbook.Sheets[name];if(!auditEngineeringHeader(auditRefSheetAoa(sheet,200),kind))continue;
    const parsed=auditReadEngineeringAoa(auditRefSheetAoa(sheet),kind,name);if(parsed?.entries.length)references.push(parsed);
  }
  return references;
}
function auditEngineeringRule(id,title,category,statement){return Object.freeze({id:`reference.engineering-${id}`,version:1,source:'reference',category,title,statement,standardRef:title,confidence:'strong',enabled:true,disabledReason:''});}
export const SSM_AUDIT_ENGINEERING_RULES=Object.freeze({
  missing:auditEngineeringRule('missing-tag','Source tag not found in registry','missing-tags','Tags in the selected project sheets should be reconciled with registry scope. Absence alone does not prove that equipment must be added.'),
  absent:auditEngineeringRule('not-in-mel','Registry tag not found in selected MEL','missing-tags','Physical equipment not found in the selected MEL needs a scope or naming check. Organizational headers are excluded.'),
  metadata:auditEngineeringRule('metadata','Registry metadata differs from MEL','metadata','For an exact equipment tag, reconcile nonblank MEL metadata with registry values. A difference does not establish which source is correct.'),
  ambiguous:auditEngineeringRule('conflicting-mel','MEL contains conflicting values','metadata','Conflicting values for the same exact tag must be resolved before a source-based correction is selected.'),
  relation:auditEngineeringRule('relationship','Source connection not represented in registry','dependencies','Confirm an explicit power or control connection as a closest parent or dependency. Electrical flow is not the same as SSM nesting.'),
  parent:auditEngineeringRule('mel-parent','Closest Parent differs from MEL','structure','Review an explicit MEL system parent against the registry parent. Confirm UPN and commissioning intent before changing nesting.'),
});
export function auditEngineeringRuleActive(rule,references={}){
  const kinds=Object.keys(AUDIT_ENGINEERING_KINDS).filter(kind=>references[kind]?.entries?.length);
  if(rule.id===SSM_AUDIT_ENGINEERING_RULES.missing.id)return kinds.length>0;
  if(rule.id===SSM_AUDIT_ENGINEERING_RULES.relation.id)return kinds.some(kind=>kind!=='mel');
  return kinds.includes('mel');
}
function auditEngineeringFinding(rule,row,kind,details){
  const equipmentId=clean(row.equipmentId),source=row._source||{},fingerprint=auditFingerprint(JSON.stringify([rule.id,kind,auditNormId(equipmentId),source.sheet||'',source.row||0,details.field,details.actual,details.expected]));
  return {schemaVersion:1,id:`${rule.id}:${fingerprint}`,fingerprint,rule,category:rule.category,equipmentId,sheet:source.sheet||'',row:source.row||0,relatedEquipmentId:'',relationship:null,...details,severity:rule.category==='missing-tags'?'missing':'warning',searchKey:auditNormId([rule.title,kind,equipmentId,details.field,details.actual,details.expected,details.why].join(' '))};
}
export function auditCoverageRows(result){
  const rows=Array.isArray(result?.rows)?result.rows:[],findings=Array.isArray(result?.findings)?result.findings:[],actualIds=new Set();
  for(const row of rows){const id=auditNormId(row?.equipmentId);if(id)actualIds.add(id);}
  const groups=new Map();
  for(const finding of findings){
    if((finding?.rule?.id||finding?.ruleId)!==SSM_AUDIT_ENGINEERING_RULES.missing.id)continue;
    const equipmentId=auditEngineeringTag(finding.equipmentId),tag=auditNormId(equipmentId);if(!tag||actualIds.has(tag))continue;
    const sheet=finding.sheet||'',row=finding.row||0,key=JSON.stringify([tag,clean(sheet),row]);
    let group=groups.get(key);
    if(!group){group={equipmentId,source:{sheet,row},values:Object.fromEntries(AUDIT_ENGINEERING_CONTEXT_FIELDS.map(field=>[field,new Map()]))};groups.set(key,group);}
    const sourceContext=finding.sourceContext||{};
    for(const field of AUDIT_ENGINEERING_CONTEXT_FIELDS){
      const value=field==='closestParent'?auditEngineeringParent(sourceContext[field]):clean(sourceContext[field]);
      const normalized=auditNormId(value),values=group.values[field];if(normalized&&!values.has(normalized))values.set(normalized,value);
    }
  }
  return [...groups.values()].map(group=>{
    const context=Object.fromEntries(AUDIT_ENGINEERING_CONTEXT_FIELDS.map(field=>{
      const values=group.values[field];return [field,values.size===1?[...values.values()][0]:''];
    }));
    return {equipmentId:group.equipmentId,...context,_source:group.source,_coverageOnly:true};
  });
}
export function auditEngineeringFindings(snapshot,references={}){
  const findings=[],rows=snapshot?.rows||[],registry=new Map();
  for(const row of rows){const key=auditNormId(row.equipmentId);if(!key)continue;const matches=registry.get(key)||[];matches.push(row);registry.set(key,matches);}
  for(const kind of Object.keys(AUDIT_ENGINEERING_KINDS)){
    const reference=references[kind];if(!reference?.entries?.length)continue;
    const label=AUDIT_ENGINEERING_KINDS[kind],byTag=new Map();
    for(const entry of reference.entries){const key=auditNormId(auditEngineeringTag(entry.equipmentId));if(!key)continue;const group=byTag.get(key)||[];group.push(entry);byTag.set(key,group);}
    for(const [tag,entries] of byTag){
      const targets=registry.get(tag),equipmentId=auditEngineeringTag(entries[0].equipmentId);
      const locations=[...new Set(entries.map(entry=>`${entry.sheetName||reference.sheetName}, row ${entry.sourceRow}`))];
      const evidence=`${label}: ${locations.slice(0,4).join('; ')}${locations.length>4?`; and ${locations.length-4} more rows`:''}.`;
      if(!targets){findings.push(auditEngineeringFinding(SSM_AUDIT_ENGINEERING_RULES.missing,{equipmentId},kind,{field:'Equipment ID',actual:'Not found in registry',expected:equipmentId,sourceContext:auditEngineeringSourceContext(entries),why:`An exact tag from ${label} is not in this registry. ${evidence}`,recommendation:'Confirm selected sheet, project scope and tag spelling before adding equipment.'}));continue;}
      if(kind==='mel'){
        for(const field of [...Object.keys(AUDIT_ENGINEERING_FIELDS),'closestParent']){
          const choices=new Map();for(const entry of entries){const value=auditEngineeringContextValue(field,entry);if(value)choices.set(auditNormId(value),value);}
          if(!choices.size)continue;
          const fieldName=AUDIT_ENGINEERING_FIELDS[field]||'Closest Parent',expected=[...choices.values()].join(' | '),ambiguous=choices.size>1;
          for(const row of targets){
            if(!ambiguous&&choices.has(auditNormId(row[field])))continue;
            const rule=ambiguous?SSM_AUDIT_ENGINEERING_RULES.ambiguous:field==='closestParent'?SSM_AUDIT_ENGINEERING_RULES.parent:SSM_AUDIT_ENGINEERING_RULES.metadata;
            const inc=field==='discipline'&&/^(?:I&C|INC|INSTRUMENTATION\s*(?:AND|&)\s*CONTROL)$/i.test(expected);
            findings.push(auditEngineeringFinding(rule,row,kind,{field:fieldName,actual:clean(row[field]),expected,why:`${ambiguous?`The selected MEL has conflicting ${fieldName} values for this tag.`:`${fieldName} differs: registry "${clean(row[field])||'(blank)'}"; MEL "${expected}".`} ${evidence}`,recommendation:ambiguous?'Resolve the source conflict before entering a correction.':inc?'Do not copy I&C into the Exto discipline. Confirm the tag UPN, then reconcile the system and its approved discipline.':'Reconcile the difference using the governing project information; neither source is changed automatically.'}));
          }
        }
      }else{
        const parents=new Map();for(const entry of entries){const parent=auditEngineeringTag(entry.parent);if(parent&&auditNormId(parent)!==tag)parents.set(auditNormId(parent),parent);}
        for(const row of targets){
          const linked=new Set([auditNormId(row.closestParent),...(!clean(row.dependencyProject)||auditSplitReferences(row.dependencyProject).every(project=>auditNormId(project)===auditNormId(row.project))?auditSplitReferences(row.dependencies).map(auditNormId):[])]);
          const missing=[...parents].filter(([key])=>!linked.has(key)).map(([,value])=>value);
          if(missing.length)findings.push(auditEngineeringFinding(SSM_AUDIT_ENGINEERING_RULES.relation,row,kind,{field:'Dependencies',actual:[row.closestParent,row.dependencies].filter(Boolean).join('; '),expected:missing.join('; '),why:`A ${label} connection is not represented as a local parent or dependency. ${evidence}`,recommendation:'Confirm this feed or control connection and source scope. Use the appropriate parent or dependency; do not replace existing links solely to clear this check.'}));
        }
      }
    }
    if(kind==='mel')for(const row of rows)if(row.equipmentId&&!byTag.has(auditNormId(row.equipmentId))&&!auditIsBlankItemMaster(row))findings.push(auditEngineeringFinding(SSM_AUDIT_ENGINEERING_RULES.absent,row,kind,{severity:'info',field:'Equipment ID',actual:row.equipmentId,expected:'Exact tag in selected MEL',why:'This registry tag is not in the selected MEL sheets. The MEL may cover only part of the project.',recommendation:'Confirm scope, naming aliases and the selected sheets. Do not delete equipment based on absence alone.'}));
  }
  return findings;
}
