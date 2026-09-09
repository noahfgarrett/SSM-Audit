import { runSsmAudit } from './engine.js'
import { auditReferenceFindings, SSM_AUDIT_REFERENCE_RULES } from './references.js'
import { auditMigrationReferences, auditMigrationImpact } from './milestone-migration.js'
import { auditApplyCorrections, auditCorrectionImpact } from './actions.js'
import { validateAuditCorrections } from './export.js'

export function auditSessionResult(snapshot,references={},migration){
  references=auditMigrationReferences(references,migration);
  const catalog=references.itemMasters;
  const raw=runSsmAudit(snapshot,catalog?{itemMasterVocabulary:catalog.entries.map(entry=>entry.name||entry.value).filter(Boolean)}:{});
  const extra=auditReferenceFindings(snapshot,references);if(!references.milestones)return raw;
  const findings=[...raw.findings,...extra],severity={blocker:0,error:0,warning:0,info:0},category={...raw.summary.category},source={...raw.summary.source,reference:extra.length};
  for(const finding of extra)category[finding.category]=(category[finding.category]||0)+1;
  for(const finding of findings)severity[finding.severity]++;
  return {...raw,findings,summary:{...raw.summary,checks:raw.summary.checks+Object.values(SSM_AUDIT_REFERENCE_RULES).length,findings:findings.length,severity,category,source,status:severity.blocker?'blocked':severity.error||severity.warning||raw.summary.unverified||extra.length?'review':'ready'}};
}

// This cache belongs to one worker and one original registry, never storage.
export async function auditPrepareInWorker(cache,data,report=()=>{}){
  if(data.baseline){cache.baseline=data.baseline;cache.file=data.file;cache.workbook=null;cache.contextKey=null;cache.previous=null;}
  if(!cache.baseline)throw new Error('Open the original registry before reviewing changes.');
  const baseline=cache.baseline,{changes,references,migration,previousChanges}=data;
  report(.15,'Validating source cells');
  const snapshot=auditApplyCorrections(baseline,changes);
  let exportCheck={cellCount:0,sheetCount:0};
  if(changes.length){
    if(!cache.workbook){const bytes=new Uint8Array(await cache.file.arrayBuffer());if(bytes[0]!==0x50||bytes[1]!==0x4b)throw new Error('Draft corrections require an original XLSX registry.');cache.workbook=XLSX.read(bytes,{type:'array',cellStyles:true});}
    exportCheck=validateAuditCorrections(cache.workbook,baseline,changes);
  }
  const contextKey=JSON.stringify([references,migration]),previousKey=JSON.stringify(previousChanges);
  report(.4,'Running audit checks');
  if(cache.contextKey!==contextKey){cache.baselineResult=auditSessionResult(baseline,references,migration);cache.previous=null;cache.contextKey=contextKey;}
  const before=cache.previous?.key===previousKey?cache.previous:{snapshot:auditApplyCorrections(baseline,previousChanges),result:null};
  before.result||=previousChanges.length?auditSessionResult(before.snapshot,references,migration):cache.baselineResult;
  const result=changes.length?auditSessionResult(snapshot,references,migration):cache.baselineResult;
  report(.85,'Comparing findings');
  const impact=auditMigrationImpact(auditCorrectionImpact(before.result,result),before.snapshot,snapshot,migration);
  const draftResolvedIds=auditCorrectionImpact(cache.baselineResult,result).resolved.map(f=>f.id);
  cache.previous={key:JSON.stringify(changes),snapshot,result};
  return {snapshot,result,exportCheck,impact,draftResolvedIds,...(data.migrationChanged?{milestoneMigration:migration,baselineResult:cache.baselineResult}:{})};
}
