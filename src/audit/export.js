import { clean } from '../core/text.js'
import { downloadBlob, styleHeaderRow, workbookBlob } from '../core/download.js'
import { S } from '../state.js'
import { toast } from '../ui/feedback.js'
import { SSM_AUDIT_RULES } from './engine.js'

function addSheet(workbook,sheet,name){XLSX.utils.book_append_sheet(workbook,sheet,name);}
function printable(value){return typeof value==='string'?value:JSON.stringify(value);}

export function exportSsmAuditXlsx(){
  const result=S.session&&S.session.result;if(!result){toast('Run an SSM Audit first');return;}
  const workbook=XLSX.utils.book_new(),summary=result.summary;
  const summaryRows=[
    ['SSM Audit','Value'],
    ['Standard',result.standard],
    ['Source file',S.session.name],
    ['Status',summary.status.toUpperCase()],
    ['Rows audited',summary.rows],
    ['Checks completed',summary.checks],
    ['Total findings',summary.findings],
    ['Blockers',summary.severity.blocker],
    ['Errors',summary.severity.error],
    ['Warnings',summary.severity.warning],
    ['Advisories',summary.severity.info],
  ];
  const summarySheet=XLSX.utils.aoa_to_sheet(summaryRows);
  summarySheet['!cols']=[{wch:24},{wch:72}];styleHeaderRow(summarySheet);addSheet(workbook,summarySheet,'Audit Summary');

  const findingHeaders=['Severity','Confidence','Category','Rule Source','Rule','Rule ID','Equipment ID','Sheet','Row','Field','Why flagged','Found','Expected','Recommended correction','Fingerprint'];
  const findingRows=result.findings.map(finding=>[
    finding.severity.toUpperCase(),finding.rule.confidence,finding.category,finding.rule.source,finding.rule.title,finding.rule.id,finding.equipmentId,finding.sheet,finding.row,finding.field,finding.why,
    printable(finding.actual),printable(finding.expected),finding.recommendation,finding.fingerprint,
  ]);
  const findingsSheet=XLSX.utils.aoa_to_sheet([findingHeaders,...findingRows]);
  findingsSheet['!cols']=[{wch:11},{wch:18},{wch:18},{wch:20},{wch:32},{wch:34},{wch:28},{wch:22},{wch:9},{wch:25},{wch:58},{wch:45},{wch:45},{wch:58},{wch:12}];
  styleHeaderRow(findingsSheet);findingsSheet['!autofilter']={ref:`A1:O${Math.max(1,findingRows.length+1)}`};addSheet(workbook,findingsSheet,'Findings');

  const byRule=new Map(Object.values(SSM_AUDIT_RULES).map(rule=>[rule.id,{rule,blocker:0,error:0,warning:0,info:0,total:0}]));
  for(const finding of result.findings){
    const entry=byRule.get(finding.rule.id)||{rule:finding.rule,blocker:0,error:0,warning:0,info:0,total:0};
    entry[finding.severity]++;entry.total++;byRule.set(finding.rule.id,entry);
  }
  const ruleRows=[['Rule Source','Rule','Status','Confidence','Rule Statement','Rule ID','Category','Blockers','Errors','Warnings','Advisories','Total'],
    ...[...byRule.values()].sort((a,b)=>Number(b.rule.enabled)-Number(a.rule.enabled)||b.total-a.total||a.rule.id.localeCompare(b.rule.id)).map(entry=>[entry.rule.source,entry.rule.title,entry.rule.enabled?'Enabled':'Disabled',entry.rule.confidence,entry.rule.statement+(entry.rule.disabledReason?` ${entry.rule.disabledReason}`:''),entry.rule.id,entry.rule.category,entry.blocker,entry.error,entry.warning,entry.info,entry.total])];
  const rulesSheet=XLSX.utils.aoa_to_sheet(ruleRows);rulesSheet['!cols']=[{wch:20},{wch:32},{wch:12},{wch:18},{wch:70},{wch:40},{wch:20},{wch:11},{wch:11},{wch:11},{wch:11},{wch:11}];
  styleHeaderRow(rulesSheet);rulesSheet['!autofilter']={ref:`A1:L${Math.max(1,ruleRows.length)}`};addSheet(workbook,rulesSheet,'Rule Summary');

  const base=clean(S.session.name).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  downloadBlob(`${base}-Audit.xlsx`,workbookBlob(workbook));toast('SSM Audit report exported');
}

export function exportSsmComparisonXlsx(){
  const comparison=S.comparison,result=comparison&&comparison.result;if(!result){toast('Compare two registries first');return;}
  const workbook=XLSX.utils.book_new(),summary=result.summary;
  const summaryRows=[
    ['Project Comparison','Value'],
    ['Method',result.standard],
    ['Target registry',comparison.targetName],
    ['Completed project reference',comparison.referenceName],
    ['Systems compared',summary.systems],
    ['Systems with differences',summary.differentSystems+summary.targetOnlySystems+summary.referenceOnlySystems],
    ['Aligned systems',summary.alignedSystems],
    ['Target rows',summary.targetRows],
    ['Completed project rows',summary.referenceRows],
    ['Aligned equipment pairs',summary.alignedRows],
    ['Changed equipment pairs',summary.changedRows],
    ['Target-only equipment',summary.targetOnlyRows],
    ['Completed-project-only equipment',summary.referenceOnlyRows],
    ['Building comparison','Excluded'],
  ];
  const summarySheet=XLSX.utils.aoa_to_sheet(summaryRows);summarySheet['!cols']=[{wch:34},{wch:80}];styleHeaderRow(summarySheet);addSheet(workbook,summarySheet,'Comparison Summary');

  const systemRows=[['UPN','System','Status','Differences','Target Rows','Completed Project Rows','Target Headers','Completed Project Headers','Target I&C','Completed Project I&C'],...result.systems.map(system=>[system.upn,system.label,system.status,system.differenceCount,system.targetRows,system.referenceRows,system.targetHeaders,system.referenceHeaders,system.targetControls,system.referenceControls])];
  const systemSheet=XLSX.utils.aoa_to_sheet(systemRows);systemSheet['!cols']=[{wch:12},{wch:58},{wch:22},{wch:13},{wch:13},{wch:22},{wch:15},{wch:24},{wch:12},{wch:22}];styleHeaderRow(systemSheet);systemSheet['!autofilter']={ref:`A1:J${Math.max(1,systemRows.length)}`};addSheet(workbook,systemSheet,'Systems');

  const differenceRows=[['UPN','Difference Type','Finding','Pattern or Equipment Type','Target','Completed Project']];
  for(const system of result.systems)for(const item of system.observations)differenceRows.push([system.upn,item.type,item.title,item.subject,item.target,item.reference]);
  const differenceSheet=XLSX.utils.aoa_to_sheet(differenceRows);differenceSheet['!cols']=[{wch:12},{wch:20},{wch:38},{wch:72},{wch:18},{wch:24}];styleHeaderRow(differenceSheet);differenceSheet['!autofilter']={ref:`A1:F${Math.max(1,differenceRows.length)}`};addSheet(workbook,differenceSheet,'Observed Differences');

  const pairRows=[['UPN','Status','Match Basis','Target Equipment ID','Target Description','Target Parent Type','Target Header','Target Dependencies','Completed Project Equipment ID','Completed Project Description','Completed Project Parent Type','Completed Project Header','Completed Project Dependencies','Differences']];
  for(const system of result.systems)for(const pair of system.pairs){const target=pair.target,reference=pair.reference;pairRows.push([system.upn,pair.status,pair.matchReason,target&&target.tag||'',target&&target.role||'',target&&target.parentRole||'',target&&target.headerName||'',target&&target.dependencyRoles.join('; ')||'',reference&&reference.tag||'',reference&&reference.role||'',reference&&reference.parentRole||'',reference&&reference.headerName||'',reference&&reference.dependencyRoles.join('; ')||'',pair.differences.join('; ')]);}
  const pairSheet=XLSX.utils.aoa_to_sheet(pairRows);pairSheet['!cols']=[{wch:12},{wch:20},{wch:38},{wch:30},{wch:34},{wch:34},{wch:30},{wch:40},{wch:36},{wch:34},{wch:34},{wch:30},{wch:40},{wch:54}];styleHeaderRow(pairSheet);pairSheet['!autofilter']={ref:`A1:N${Math.max(1,pairRows.length)}`};addSheet(workbook,pairSheet,'Equipment Mapping');

  const base=clean(comparison.targetName).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  downloadBlob(`${base}-Project-Comparison.xlsx`,workbookBlob(workbook));toast('Project comparison exported');
}
