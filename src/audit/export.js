import { clean } from '../core/text.js'
import { downloadBlob, styleHeaderRow, workbookBlob } from '../core/download.js'
import { S } from '../state.js'
import { toast } from '../ui/feedback.js'

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

  const findingHeaders=['Severity','Category','Rule ID','Equipment ID','Sheet','Row','Field','Why flagged','Found','Expected','Recommended correction','Standard reference','Fingerprint'];
  const findingRows=result.findings.map(finding=>[
    finding.severity.toUpperCase(),finding.category,finding.rule.id,finding.equipmentId,finding.sheet,finding.row,finding.field,finding.why,
    printable(finding.actual),printable(finding.expected),finding.recommendation,finding.rule.standardRef,finding.fingerprint,
  ]);
  const findingsSheet=XLSX.utils.aoa_to_sheet([findingHeaders,...findingRows]);
  findingsSheet['!cols']=[{wch:11},{wch:18},{wch:32},{wch:28},{wch:22},{wch:9},{wch:25},{wch:58},{wch:45},{wch:45},{wch:58},{wch:42},{wch:12}];
  styleHeaderRow(findingsSheet);findingsSheet['!autofilter']={ref:`A1:M${Math.max(1,findingRows.length+1)}`};addSheet(workbook,findingsSheet,'Findings');

  const byRule=new Map();
  for(const finding of result.findings){
    const entry=byRule.get(finding.rule.id)||{rule:finding.rule,blocker:0,error:0,warning:0,info:0,total:0};
    entry[finding.severity]++;entry.total++;byRule.set(finding.rule.id,entry);
  }
  const ruleRows=[['Rule ID','Category','Standard reference','Blockers','Errors','Warnings','Advisories','Total'],
    ...[...byRule.values()].sort((a,b)=>b.total-a.total||a.rule.id.localeCompare(b.rule.id)).map(entry=>[entry.rule.id,entry.rule.category,entry.rule.standardRef,entry.blocker,entry.error,entry.warning,entry.info,entry.total])];
  const rulesSheet=XLSX.utils.aoa_to_sheet(ruleRows);rulesSheet['!cols']=[{wch:34},{wch:20},{wch:44},{wch:11},{wch:11},{wch:11},{wch:11},{wch:11}];
  styleHeaderRow(rulesSheet);rulesSheet['!autofilter']={ref:`A1:H${Math.max(1,ruleRows.length)}`};addSheet(workbook,rulesSheet,'Rule Summary');

  const base=clean(S.session.name).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  downloadBlob(`${base}-Audit.xlsx`,workbookBlob(workbook));toast('SSM Audit report exported');
}
