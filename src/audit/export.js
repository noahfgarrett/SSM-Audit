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
