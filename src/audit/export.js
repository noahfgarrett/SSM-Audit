import { clean, natCmp } from '../core/text.js'
import { downloadBlob, sheetAutoFilter, sheetCellStyle, sheetFormulaCell, sheetFreezeRows, sheetLinkCell, sheetSetCell, sheetStyleCell, sheetXmlExtras, styleHeaderRow, workbookBlob, workbookBlobCompact } from '../core/download.js'
import { S } from '../state.js'
import { runWithProgress, toast } from '../ui/feedback.js'
import { SSM_AUDIT_RULES } from './engine.js'
import { auditColumnName, auditNormId } from './model.js'

function addSheet(workbook,sheet,name){XLSX.utils.book_append_sheet(workbook,sheet,name);}
function printable(value){return typeof value==='string'?value:JSON.stringify(value);}
const EXPORT_SOURCE_LABELS={registry:'Registry Integrity',sop:'SSM SOP',logic:'Commissioning Logic'},EXPORT_CONFIDENCE_LABELS={required:'Required',strong:'Strong pattern','description-rated':'Description based'};

/* ---- SSM Audit workbook ----
   One tab per L2 milestone so a commissioning engineer can work a milestone end
   to end: the full equipment tree of that milestone, one line per finding, and an
   `Actioned` column with a tick-box dropdown (☐ / ☑). Index and Dashboard read
   those ticks back with live COUNTIF formulas, so progress updates itself. */
const AUDIT_EXPORT_DASHBOARD_SHEET='Dashboard',AUDIT_EXPORT_INDEX_SHEET='Index',AUDIT_EXPORT_FINDINGS_SHEET='All Findings',AUDIT_EXPORT_RULES_SHEET='Rules',AUDIT_EXPORT_CALC_SHEET='Calc';
const AUDIT_EXPORT_NO_MILESTONE='No milestone',AUDIT_EXPORT_BAR_SEGMENTS=10,AUDIT_EXPORT_WIDE_BAR_SEGMENTS=25,AUDIT_EXPORT_BACK_LINK='← Index';
export const AUDIT_EXPORT_TICK='☑',AUDIT_EXPORT_UNTICKED='☐';
const AUDIT_EXPORT_ACTIONED_NOTE=`Click the Actioned cell on an equipment’s first line and choose ${AUDIT_EXPORT_TICK} when it is closed out. The row turns green, and Index and Dashboard progress update from those ticks. Cells shaded red are the ones the finding is about.`;
const AUDIT_EXPORT_PALETTE=Object.freeze({ink:'173F5F',accent:'F26722',headerText:'FFFFFF',body:'21323F',black:'000000',muted:'8A96A3',repeat:'5B6773',band:'F4F7FA',line:'D9E1E9',link:'1B5FAA',flag:'FBE3E1',done:'E3F5E8',barTrack:'FDEFE6'});
/* One distinct hue per nest level (not shades of one colour) so siblings under
   the same parent read as a band and a child is obviously a different band. All
   light enough for black text. Level 7+ cycles. */
const AUDIT_EXPORT_NEST_FILLS=Object.freeze(['DCE8F4','DFF2E3','FFF0CC','EADDF6','D6F0F3','FBE0EA','ECEFD3','F4E3D2']);
const AUDIT_EXPORT_SEVERITY_LABELS=Object.freeze({blocker:'INVALID',error:'RULE BROKEN',warning:'CHECK THIS',info:'NOTE'});
const AUDIT_EXPORT_SEVERITY_COLORS=Object.freeze({blocker:{fill:'8C1D18',color:'FFFFFF'},error:{fill:'D9531E',color:'FFFFFF'},warning:{fill:'F2B441',color:'40320A'},info:{fill:'6E8598',color:'FFFFFF'}});
/* Milestone tab columns. Closest Parent and Dependencies sit side by side so a
   hierarchy question can be checked without scrolling. */
const AUDIT_EXPORT_MILESTONE_HEADERS=Object.freeze(['Actioned','Nest','Equipment ID','Description','Closest Parent','Dependencies','Discipline','UPN','System Name','Building','Item Master','Severity','Finding','Why','What to do','Actioned By','Note']);
const AUDIT_EXPORT_MILESTONE_WIDTHS=Object.freeze([10,6,36,38,28,34,22,8,30,10,26,12,32,54,54,16,30]);
const AUDIT_EXPORT_MILESTONE_ALIGN=Object.freeze(['center','center','left','left','left','left','left','right','left','left','left','center','left','left','left','left','left']);
/* Which column a finding's `field` points at, so that cell can be shaded. */
const AUDIT_EXPORT_FIELD_COLUMNS=Object.freeze({'Equipment ID':'C','Closest Parent':'E','Dependencies':'F','Dependency Project':'F','Discipline':'G','UPN':'H','System Name':'I','Building':'J','Item Master Unique Identifier':'K'});
const AUDIT_EXPORT_ROW_COLUMNS=Object.freeze(['D','E','F','G','H','I','J','K']);
const AUDIT_EXPORT_FINDING_HEADERS=Object.freeze(['Severity','Milestone','Equipment ID','Description','Rule','Why','What to do','Field','Found','Expected','Sheet','Row']);
const AUDIT_EXPORT_FINDING_WIDTHS=Object.freeze([12,34,30,40,34,54,54,26,44,44,22,8]);
const AUDIT_EXPORT_RULE_HEADERS=Object.freeze(['Rule','What must be true','Source','Confidence','Findings count']);
const AUDIT_EXPORT_RULE_WIDTHS=Object.freeze([36,84,24,20,16]);

const AUDIT_EXPORT_STYLES=Object.freeze({
  title:sheetCellStyle({bold:true,size:18,color:AUDIT_EXPORT_PALETTE.ink}),
  subtitle:sheetCellStyle({size:11,color:AUDIT_EXPORT_PALETTE.muted}),
  section:sheetCellStyle({bold:true,size:12,color:AUDIT_EXPORT_PALETTE.ink,vertical:'center'}),
  note:sheetCellStyle({italic:true,size:10,color:AUDIT_EXPORT_PALETTE.muted}),
  back:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.link,underline:true}),
  kpiLabel:sheetCellStyle({bold:true,size:9,color:AUDIT_EXPORT_PALETTE.muted,align:'center',bottom:AUDIT_EXPORT_PALETTE.accent,bottomWeight:'medium'}),
  kpiValue:sheetCellStyle({bold:true,size:16,color:AUDIT_EXPORT_PALETTE.ink,align:'center'}),
  kpiPercent:sheetCellStyle({bold:true,size:16,color:AUDIT_EXPORT_PALETTE.accent,align:'center',numFmt:'0%'}),
  overallLabel:sheetCellStyle({bold:true,size:10,color:AUDIT_EXPORT_PALETTE.muted}),
  overallBar:sheetCellStyle({size:20,color:AUDIT_EXPORT_PALETTE.accent,fill:AUDIT_EXPORT_PALETTE.barTrack,align:'left',vertical:'center'}),
  overallPercent:sheetCellStyle({bold:true,size:26,color:AUDIT_EXPORT_PALETTE.accent,fill:AUDIT_EXPORT_PALETTE.barTrack,align:'center',vertical:'center',numFmt:'0%'}),
  overallCaption:sheetCellStyle({bold:true,size:11,color:AUDIT_EXPORT_PALETTE.ink,fill:AUDIT_EXPORT_PALETTE.barTrack,align:'left',vertical:'center'}),
  header:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.headerText,fill:AUDIT_EXPORT_PALETTE.ink,align:'left',vertical:'center',wrap:true}),
  headerRight:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.headerText,fill:AUDIT_EXPORT_PALETTE.ink,align:'right',vertical:'center',wrap:true}),
  headerCenter:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.headerText,fill:AUDIT_EXPORT_PALETTE.ink,align:'center',vertical:'center',wrap:true}),
  text:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,vertical:'top'}),
  wrap:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,vertical:'top',wrap:true}),
  wrapMuted:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.muted,vertical:'top',wrap:true}),
  muted:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.muted,vertical:'top'}),
  repeat:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.repeat,vertical:'top'}),
  flag:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.black,fill:AUDIT_EXPORT_PALETTE.flag,vertical:'top'}),
  linkText:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.link,underline:true,vertical:'top'}),
  linkBand:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.link,underline:true,fill:AUDIT_EXPORT_PALETTE.band,vertical:'top'}),
  label:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.body,vertical:'center'}),
  labelBand:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.body,fill:AUDIT_EXPORT_PALETTE.band,vertical:'center'}),
  number:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,align:'right'}),
  numberBand:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,fill:AUDIT_EXPORT_PALETTE.band,align:'right'}),
  numberMid:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,align:'right',vertical:'center'}),
  numberMidBand:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,fill:AUDIT_EXPORT_PALETTE.band,align:'right',vertical:'center'}),
  percent:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.ink,align:'right',numFmt:'0%'}),
  percentBand:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.ink,fill:AUDIT_EXPORT_PALETTE.band,align:'right',numFmt:'0%'}),
  percentBig:sheetCellStyle({bold:true,size:13,color:AUDIT_EXPORT_PALETTE.ink,align:'right',vertical:'center',numFmt:'0%'}),
  percentBigBand:sheetCellStyle({bold:true,size:13,color:AUDIT_EXPORT_PALETTE.ink,fill:AUDIT_EXPORT_PALETTE.band,align:'right',vertical:'center',numFmt:'0%'}),
  bar:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.accent,align:'left'}),
  barBand:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.accent,fill:AUDIT_EXPORT_PALETTE.band,align:'left'}),
  barWide:sheetCellStyle({size:14,color:AUDIT_EXPORT_PALETTE.accent,align:'left',vertical:'center'}),
  barWideBand:sheetCellStyle({size:14,color:AUDIT_EXPORT_PALETTE.accent,fill:AUDIT_EXPORT_PALETTE.band,align:'left',vertical:'center'}),
  nest:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.black,align:'center'}),
  actioned:sheetCellStyle({align:'center',size:14,color:AUDIT_EXPORT_PALETTE.ink,border:AUDIT_EXPORT_PALETTE.line}),
  actionedRepeat:sheetCellStyle({align:'center',fill:AUDIT_EXPORT_PALETTE.band,border:AUDIT_EXPORT_PALETTE.line}),
  entry:sheetCellStyle({border:AUDIT_EXPORT_PALETTE.line,vertical:'top'}),
});
/* Equipment IDs are always black: bold on the equipment's own line, regular on
   the repeat lines that carry its further findings. */
const AUDIT_EXPORT_NEST_STYLES=Object.freeze(AUDIT_EXPORT_NEST_FILLS.map(fill=>Object.freeze({
  own:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.black,fill,vertical:'top'}),
  repeat:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.black,fill,vertical:'top'}),
  level:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.black,fill,align:'center'}),
})));
const AUDIT_EXPORT_SEVERITY_STYLES=Object.freeze(Object.fromEntries(Object.entries(AUDIT_EXPORT_SEVERITY_COLORS).map(([severity,colors])=>[severity,sheetCellStyle({bold:true,size:10,color:colors.color,fill:colors.fill,align:'center',vertical:'center'})])));
/* Differential style 0: the green an actioned row turns. Referenced by the
   conditional format each milestone tab carries. */
const AUDIT_EXPORT_DXFS=Object.freeze([`<dxf><fill><patternFill><bgColor rgb="FF${AUDIT_EXPORT_PALETTE.done}"/></patternFill></fill></dxf>`]);

function auditExportEmptyBar(segments=AUDIT_EXPORT_BAR_SEGMENTS){return '░'.repeat(segments);}
function auditExportDate(value){
  try{return value.toLocaleString(undefined,{year:'numeric',month:'short',day:'2-digit',hour:'2-digit',minute:'2-digit'});}
  catch(_){return value.toISOString();}
}
function auditExportSheetRef(name){return `'${String(name).replace(/'/g,"''")}'`;}
function auditExportRowKey(row){const source=row&&row._source||{};return `${auditNormId(row&&row.equipmentId)}${clean(source.sheet)}${source.row||0}`;}
function auditExportFindingKey(finding){return `${auditNormId(finding&&finding.equipmentId)}${clean(finding&&finding.sheet)}${finding&&finding.row||0}`;}

/* Excel forbids [ ] : * ? / \ in a tab name, caps it at 31 characters, and will
   not open a workbook with two tabs sharing a name. Milestone names routinely
   break all three, so the final name is stored on the group and every hyperlink
   is written from that stored name. */
export function auditExportSheetName(label,used){
  const base=clean(label).replace(/[[\]:*?/\\]/g,' ').replace(/\s+/g,' ').replace(/^'+|'+$/g,'').trim().slice(0,31).trim()||'Milestone';
  let name=base,attempt=2;
  while(used&&used.has(name.toLowerCase())){const tag=` (${attempt++})`;name=`${base.slice(0,31-tag.length).trim()}${tag}`;}
  if(used)used.add(name.toLowerCase());
  return name;
}

/* Depth from the system root, walked over the whole registry rather than one
   milestone, so a child keeps its true depth even when its parent sits on a
   different milestone. Iterative: registry hierarchies run thousands deep. */
export function auditExportNestLevels(rows){
  const rowsById=new Map();
  for(const row of rows||[]){const id=auditNormId(row&&row.equipmentId);if(id&&!rowsById.has(id))rowsById.set(id,row);}
  const levelById=new Map();
  for(const row of rows||[]){
    const chain=[],guard=new Set();let current=row;
    while(current){
      const id=auditNormId(current.equipmentId);
      if(!id||guard.has(id)||levelById.has(id))break;
      guard.add(id);chain.push(id);
      const parentId=auditNormId(current.closestParent);
      current=parentId&&parentId!==id?rowsById.get(parentId)||null:null;
    }
    let base=0;
    if(current){const id=auditNormId(current.equipmentId);if(id&&levelById.has(id))base=levelById.get(id)+1;}
    for(let index=chain.length-1;index>=0;index--)levelById.set(chain[index],base+(chain.length-1-index));
  }
  return row=>{
    const id=auditNormId(row&&row.equipmentId);
    if(id&&levelById.has(id))return levelById.get(id);
    const parentId=auditNormId(row&&row.closestParent);
    return parentId&&levelById.has(parentId)?levelById.get(parentId)+1:0;
  };
}

/* Depth-first over the rows of one milestone, siblings in natural tag order, so
   the tab reads as a tree and a parent always precedes its children. Rows whose
   parent is on another milestone start their own branch; rows inside a parent
   cycle are appended rather than dropped. */
export function auditExportOrderRows(groupRows){
  const present=new Set(),childrenByParent=new Map(),roots=[];
  for(const row of groupRows||[]){const id=auditNormId(row&&row.equipmentId);if(id)present.add(id);}
  for(const row of groupRows||[]){
    const id=auditNormId(row&&row.equipmentId),parentId=auditNormId(row&&row.closestParent);
    if(parentId&&parentId!==id&&present.has(parentId)){const list=childrenByParent.get(parentId)||[];list.push(row);childrenByParent.set(parentId,list);}
    else roots.push(row);
  }
  const sorted=list=>[...list].sort((left,right)=>natCmp(clean(left.equipmentId),clean(right.equipmentId)));
  const ordered=[],visited=new Set(),stack=sorted(roots).reverse();
  while(stack.length){
    const row=stack.pop();if(visited.has(row))continue;
    visited.add(row);ordered.push(row);
    const children=sorted(childrenByParent.get(auditNormId(row.equipmentId))||[]);
    for(let index=children.length-1;index>=0;index--)if(!visited.has(children[index]))stack.push(children[index]);
  }
  for(const row of groupRows||[])if(!visited.has(row)){visited.add(row);ordered.push(row);}
  return ordered;
}

export function auditExportGroups(result){
  const rows=result&&result.rows||[],findings=result&&result.findings||[],levelFor=auditExportNestLevels(rows);
  const findingsByRow=new Map();
  for(const finding of findings){const key=auditExportFindingKey(finding),list=findingsByRow.get(key)||[];list.push(finding);findingsByRow.set(key,list);}
  const groups=new Map();
  for(const row of rows){
    const label=clean(row.milestone)||AUDIT_EXPORT_NO_MILESTONE;
    const group=groups.get(label)||{label,sheetName:'',rows:[],lines:[],equipmentCount:0,findingCount:0};
    group.rows.push(row);groups.set(label,group);
  }
  const ordered=[...groups.values()].sort((left,right)=>
    (left.label===AUDIT_EXPORT_NO_MILESTONE?1:0)-(right.label===AUDIT_EXPORT_NO_MILESTONE?1:0)||natCmp(left.label,right.label));
  for(const group of ordered){
    group.equipmentCount=group.rows.length;
    for(const row of auditExportOrderRows(group.rows)){
      const rowFindings=findingsByRow.get(auditExportRowKey(row))||[];
      group.findingCount+=rowFindings.length;
      group.lines.push({row,level:levelFor(row),findings:rowFindings});
    }
  }
  return ordered;
}

/* Live progress: the milestone tab owns the truth (column A, typed Y), Index and
   Dashboard only read it back. Every formula carries a cached value so the file
   is readable before Excel recalculates. */
function auditExportActionedCell(group,style){return sheetFormulaCell(`COUNTIF(${auditExportSheetRef(group.sheetName)}!A:A,"${AUDIT_EXPORT_TICK}")`,0,style||AUDIT_EXPORT_STYLES.number);}
/* Actioned equipment in one discipline: one COUNTIFS per milestone tab lives on
   the hidden Calc sheet (one formula per cell, always short); the Dashboard sums
   that column. Chaining every tab into one Dashboard formula overran Excel's
   8,192-character formula limit on registries with many milestones. */
function auditExportDisciplineActionedCell(disciplineIndex,groupCount,style){
  const column=auditColumnName(disciplineIndex+1),first=2,last=first+Math.max(0,groupCount-1);
  return sheetFormulaCell(groupCount?`SUM(${auditExportSheetRef(AUDIT_EXPORT_CALC_SHEET)}!${column}${first}:${column}${last})`:'0',0,style);
}
/* Calc: rows = milestone tabs, columns = disciplines; each cell counts that tab's
   ticked rows in that discipline. Hidden, but it is a normal sheet. */
function auditExportCalcSheet(groups,disciplines){
  const aoa=[['Milestone tab',...disciplines.map(discipline=>discipline.label)]];
  for(const group of groups)aoa.push([group.sheetName,...disciplines.map(()=>0)]);
  const sheet=XLSX.utils.aoa_to_sheet(aoa.length>1?aoa:[['Milestone tab','(no milestones)']]);
  groups.forEach((group,rowOffset)=>{
    const rowIndex=rowOffset+2,tab=auditExportSheetRef(group.sheetName);
    disciplines.forEach((discipline,columnOffset)=>{
      const value=String(discipline.label).replace(/"/g,'""');
      sheetSetCell(sheet,`${auditColumnName(columnOffset+1)}${rowIndex}`,sheetFormulaCell(`COUNTIFS(${tab}!G:G,"${value}",${tab}!A:A,"${AUDIT_EXPORT_TICK}")`,0,AUDIT_EXPORT_STYLES.number));
    });
  });
  sheet['!cols']=[{wch:34},...disciplines.map(()=>({wch:14}))];
  return sheet;
}
function auditExportPercentCell(equipmentCell,actionedCell,style){return sheetFormulaCell(`IF(${equipmentCell}=0,0,${actionedCell}/${equipmentCell})`,0,style);}
function auditExportBarCell(percentCell,style,segments=AUDIT_EXPORT_BAR_SEGMENTS){
  const filled=`MIN(${segments},ROUND(${percentCell}*${segments},0))`;
  return sheetFormulaCell(`REPT("█",${filled})&REPT("░",${segments}-${filled})`,auditExportEmptyBar(segments),style);
}
function auditExportTickValidation(range){return `<dataValidation type="list" allowBlank="1" showDropDown="0" showErrorMessage="1" errorTitle="Actioned" error="Pick ${AUDIT_EXPORT_TICK} or ${AUDIT_EXPORT_UNTICKED} from the list." sqref="${range}"><formula1>"${AUDIT_EXPORT_TICK},${AUDIT_EXPORT_UNTICKED}"</formula1></dataValidation>`;}
function auditExportDoneFormat(range,firstRow){return `<conditionalFormatting sqref="${range}"><cfRule type="expression" dxfId="0" priority="1"><formula>$A${firstRow}="${AUDIT_EXPORT_TICK}"</formula></cfRule></conditionalFormatting>`;}
/* ---- export plan ----
   plan.levels[severity] and plan.rules[ruleId] each hold 'include' | 'pretick'
   | 'skip'. A rule entry overrides its level; everything else defaults to
   include. 'skip' findings are left out of the workbook entirely; 'pretick'
   findings are exported with the Actioned box already ticked. */
export function auditExportPlanMode(plan,finding){
  const byRule=plan&&plan.rules&&plan.rules[finding.rule.id];
  if(byRule==='include'||byRule==='pretick'||byRule==='skip')return byRule;
  const byLevel=plan&&plan.levels&&plan.levels[finding.severity];
  return byLevel==='pretick'||byLevel==='skip'?byLevel:'include';
}
/* The result the workbook is built from: skipped findings removed, summary
   recounted, and the set of findings that arrive pre-ticked. */
export function auditExportApplyPlan(result,plan){
  const findings=[],preticked=new Set();
  for(const finding of result&&result.findings||[]){
    const mode=auditExportPlanMode(plan,finding);
    if(mode==='skip')continue;
    findings.push(finding);if(mode==='pretick')preticked.add(finding);
  }
  const severity={blocker:0,error:0,warning:0,info:0};
  for(const finding of findings)severity[finding.severity]=(severity[finding.severity]||0)+1;
  const summary=Object.assign({},result&&result.summary,{findings:findings.length,severity});
  return {result:Object.assign({},result,{findings,summary}),preticked,skipped:(result&&result.findings||[]).length-findings.length};
}

function auditExportDisciplines(result){
  const totals=new Map();
  for(const row of result&&result.rows||[]){const label=clean(row.discipline)||'No discipline';const entry=totals.get(label)||{label,equipmentCount:0,findingCount:0};entry.equipmentCount++;totals.set(label,entry);}
  const keyFor=row=>clean(row.discipline)||'No discipline';
  const rowsByKey=new Map();for(const row of result&&result.rows||[])rowsByKey.set(auditExportRowKey(row),keyFor(row));
  for(const finding of result&&result.findings||[]){const label=rowsByKey.get(auditExportFindingKey(finding));if(label&&totals.has(label))totals.get(label).findingCount++;}
  return [...totals.values()].sort((left,right)=>right.equipmentCount-left.equipmentCount||natCmp(left.label,right.label));
}
function auditExportHeaderRow(sheet,rowIndex,headers,alignments){
  for(let column=0;column<headers.length;column++){
    const style=alignments&&alignments[column]==='right'?AUDIT_EXPORT_STYLES.headerRight:alignments&&alignments[column]==='center'?AUDIT_EXPORT_STYLES.headerCenter:AUDIT_EXPORT_STYLES.header;
    sheetStyleCell(sheet,`${auditColumnName(column)}${rowIndex}`,style);
  }
  sheet['!rows']=sheet['!rows']||[];sheet['!rows'][rowIndex-1]={hpt:26};
}
function auditExportBackLink(sheet){
  sheetSetCell(sheet,'A1',{t:'s',v:AUDIT_EXPORT_BACK_LINK});
  sheetStyleCell(sheet,'A1',AUDIT_EXPORT_STYLES.back);
  sheetLinkCell(sheet,'A1',`#${auditExportSheetRef(AUDIT_EXPORT_INDEX_SHEET)}!A1`,'Back to the milestone index');
}

/* Dashboard layout (columns A–F):
     1–3  title block
     5    OVERALL PROGRESS caption        6  wide bar (A–E merged) + big % (F)
     8–9  KPI strip
     11   "By discipline" header, rows follow        (bars 25 segments, tall rows)
     then "By milestone" header, rows follow          (same bars; links to tabs)
   Every progress number is a formula over the ticks on the milestone tabs. */
function auditExportDashboardSheet(result,groups,disciplines,sessionName,generated){
  const summary=result&&result.summary||{},severity=summary.severity||{};
  const columns=['Discipline','Equipment','Findings','Actioned','%','Progress'];
  const aoa=[
    [`SSM Audit — ${clean(sessionName)||'Registry'}`,'','','','',''],
    [`Generated ${auditExportDate(generated)}`,'','','','',''],
    [`Standard: ${clean(result&&result.standard)}`,'','','','',''],
    ['','','','','',''],
    ['OVERALL PROGRESS — equipment actioned across every milestone','','','','',''],
    [auditExportEmptyBar(AUDIT_EXPORT_WIDE_BAR_SEGMENTS),'','','','',0],
    ['','','','','',''],
    ['Rows audited','Findings','Invalid','Rule broken','Check this','Notes'],
    [summary.rows||0,summary.findings||0,severity.blocker||0,severity.error||0,severity.warning||0,severity.info||0],
    ['','','','','',''],
    columns,
  ];
  const disciplineFirst=aoa.length+1;
  for(const discipline of disciplines)aoa.push([discipline.label,discipline.equipmentCount,discipline.findingCount,0,0,auditExportEmptyBar(AUDIT_EXPORT_WIDE_BAR_SEGMENTS)]);
  if(!disciplines.length)aoa.push(['No equipment rows in this registry','','','','','']);
  const disciplineLast=disciplineFirst+Math.max(0,disciplines.length-1);
  aoa.push(['','','','','','']);
  aoa.push(['Milestone','Equipment','Findings','Actioned','%','Progress']);
  const milestoneHeader=aoa.length,milestoneFirst=aoa.length+1;
  for(const group of groups)aoa.push([group.label,group.equipmentCount,group.findingCount,0,0,auditExportEmptyBar(AUDIT_EXPORT_WIDE_BAR_SEGMENTS)]);
  if(!groups.length)aoa.push(['No equipment rows in this registry','','','','','']);
  const milestoneLast=milestoneFirst+Math.max(0,groups.length-1);
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=[{wch:46},{wch:12},{wch:11},{wch:11},{wch:9},{wch:62}];
  sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:5}},{s:{r:1,c:0},e:{r:1,c:5}},{s:{r:2,c:0},e:{r:2,c:5}},{s:{r:4,c:0},e:{r:4,c:5}},{s:{r:5,c:0},e:{r:5,c:4}}];
  sheet['!rows']=[{hpt:30},{hpt:18},{hpt:18},{hpt:10},{hpt:18},{hpt:46},{hpt:12},{hpt:18},{hpt:28},{hpt:14}];
  sheetStyleCell(sheet,'A1',AUDIT_EXPORT_STYLES.title);
  sheetStyleCell(sheet,'A2',AUDIT_EXPORT_STYLES.subtitle);
  sheetStyleCell(sheet,'A3',AUDIT_EXPORT_STYLES.subtitle);
  sheetStyleCell(sheet,'A5',AUDIT_EXPORT_STYLES.overallLabel);
  /* Overall: ticks over equipment, summed from the milestone rows below. */
  const overallPercent=groups.length?`IF(SUM(B${milestoneFirst}:B${milestoneLast})=0,0,SUM(D${milestoneFirst}:D${milestoneLast})/SUM(B${milestoneFirst}:B${milestoneLast}))`:'';
  if(overallPercent)sheetSetCell(sheet,'F6',sheetFormulaCell(overallPercent,0,AUDIT_EXPORT_STYLES.overallPercent));else sheetStyleCell(sheet,'F6',AUDIT_EXPORT_STYLES.overallPercent);
  if(overallPercent)sheetSetCell(sheet,'A6',auditExportBarCell('F6',AUDIT_EXPORT_STYLES.overallBar,AUDIT_EXPORT_WIDE_BAR_SEGMENTS));else sheetStyleCell(sheet,'A6',AUDIT_EXPORT_STYLES.overallBar);
  for(const column of ['B','C','D','E'])sheetStyleCell(sheet,`${column}6`,AUDIT_EXPORT_STYLES.overallBar);
  for(let column=0;column<6;column++){
    sheetStyleCell(sheet,`${auditColumnName(column)}8`,AUDIT_EXPORT_STYLES.kpiLabel);
    sheetStyleCell(sheet,`${auditColumnName(column)}9`,AUDIT_EXPORT_STYLES.kpiValue);
  }
  const progressRow=(rowIndex,band,actionedCell,link)=>{
    sheet['!rows'][rowIndex-1]={hpt:24};
    if(link){sheetStyleCell(sheet,`A${rowIndex}`,band?AUDIT_EXPORT_STYLES.linkBand:AUDIT_EXPORT_STYLES.linkText);sheetLinkCell(sheet,`A${rowIndex}`,link.target,link.tooltip);}
    else sheetStyleCell(sheet,`A${rowIndex}`,band?AUDIT_EXPORT_STYLES.labelBand:AUDIT_EXPORT_STYLES.label);
    sheetStyleCell(sheet,`B${rowIndex}`,band?AUDIT_EXPORT_STYLES.numberMidBand:AUDIT_EXPORT_STYLES.numberMid);
    sheetStyleCell(sheet,`C${rowIndex}`,band?AUDIT_EXPORT_STYLES.numberMidBand:AUDIT_EXPORT_STYLES.numberMid);
    sheetSetCell(sheet,`D${rowIndex}`,actionedCell);
    sheetSetCell(sheet,`E${rowIndex}`,auditExportPercentCell(`B${rowIndex}`,`D${rowIndex}`,band?AUDIT_EXPORT_STYLES.percentBigBand:AUDIT_EXPORT_STYLES.percentBig));
    sheetSetCell(sheet,`F${rowIndex}`,auditExportBarCell(`E${rowIndex}`,band?AUDIT_EXPORT_STYLES.barWideBand:AUDIT_EXPORT_STYLES.barWide,AUDIT_EXPORT_WIDE_BAR_SEGMENTS));
  };
  auditExportHeaderRow(sheet,disciplineFirst-1,columns,['left','right','right','right','right','left']);
  disciplines.forEach((discipline,offset)=>{
    const rowIndex=disciplineFirst+offset,band=offset%2===1;
    progressRow(rowIndex,band,auditExportDisciplineActionedCell(offset,groups.length,band?AUDIT_EXPORT_STYLES.numberMidBand:AUDIT_EXPORT_STYLES.numberMid),null);
  });
  auditExportHeaderRow(sheet,milestoneHeader,['Milestone','Equipment','Findings','Actioned','%','Progress'],['left','right','right','right','right','left']);
  groups.forEach((group,offset)=>{
    const rowIndex=milestoneFirst+offset,band=offset%2===1;
    progressRow(rowIndex,band,auditExportActionedCell(group,band?AUDIT_EXPORT_STYLES.numberMidBand:AUDIT_EXPORT_STYLES.numberMid),{target:`#${auditExportSheetRef(group.sheetName)}!A1`,tooltip:`Open ${group.label}`});
  });
  return sheet;
}

function auditExportIndexSheet(groups){
  const firstRow=5;
  const aoa=[
    ['SSM Audit — Index','','','','',''],
    [AUDIT_EXPORT_DASHBOARD_SHEET,AUDIT_EXPORT_FINDINGS_SHEET,AUDIT_EXPORT_RULES_SHEET,'','',''],
    ['','','','','',''],
    ['Milestone','Equipment','Findings','Actioned','%','Progress'],
  ];
  for(const group of groups)aoa.push([group.label,group.equipmentCount,group.findingCount,0,0,auditExportEmptyBar()]);
  if(!groups.length)aoa.push(['No equipment rows in this registry','','','','','']);
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=[{wch:44},{wch:12},{wch:11},{wch:11},{wch:9},{wch:16}];
  sheet['!rows']=[{hpt:28},{hpt:18},{hpt:10}];
  sheetStyleCell(sheet,'A1',AUDIT_EXPORT_STYLES.title);
  [[AUDIT_EXPORT_DASHBOARD_SHEET,'A2'],[AUDIT_EXPORT_FINDINGS_SHEET,'B2'],[AUDIT_EXPORT_RULES_SHEET,'C2']].forEach(([name,address])=>{
    sheetStyleCell(sheet,address,AUDIT_EXPORT_STYLES.back);
    sheetLinkCell(sheet,address,`#${auditExportSheetRef(name)}!A1`,`Open ${name}`);
  });
  auditExportHeaderRow(sheet,4,['Milestone','Equipment','Findings','Actioned','%','Progress'],['left','right','right','right','right','left']);
  groups.forEach((group,offset)=>{
    const rowIndex=firstRow+offset,band=offset%2===1;
    sheetStyleCell(sheet,`A${rowIndex}`,band?AUDIT_EXPORT_STYLES.linkBand:AUDIT_EXPORT_STYLES.linkText);
    sheetLinkCell(sheet,`A${rowIndex}`,`#${auditExportSheetRef(group.sheetName)}!A1`,`Open ${group.label}`);
    sheetStyleCell(sheet,`B${rowIndex}`,band?AUDIT_EXPORT_STYLES.numberBand:AUDIT_EXPORT_STYLES.number);
    sheetStyleCell(sheet,`C${rowIndex}`,band?AUDIT_EXPORT_STYLES.numberBand:AUDIT_EXPORT_STYLES.number);
    const actioned=auditExportActionedCell(group);
    if(band)actioned.s=AUDIT_EXPORT_STYLES.numberBand;
    sheetSetCell(sheet,`D${rowIndex}`,actioned);
    sheetSetCell(sheet,`E${rowIndex}`,auditExportPercentCell(`B${rowIndex}`,`D${rowIndex}`,band?AUDIT_EXPORT_STYLES.percentBand:AUDIT_EXPORT_STYLES.percent));
    sheetSetCell(sheet,`F${rowIndex}`,auditExportBarCell(`E${rowIndex}`,band?AUDIT_EXPORT_STYLES.barBand:AUDIT_EXPORT_STYLES.bar));
  });
  sheetFreezeRows(sheet,4);
  return sheet;
}

function auditExportMilestoneSheet(group,preticked){
  const aoa=[[AUDIT_EXPORT_BACK_LINK,AUDIT_EXPORT_ACTIONED_NOTE],[...AUDIT_EXPORT_MILESTONE_HEADERS]];
  const meta=[];
  for(const line of group.lines){
    const row=line.row,indent='  '.repeat(Math.min(line.level,24));
    const equipment=[indent+clean(row.equipmentId),clean(row.equipmentDescription),clean(row.closestParent),clean(row.dependencies),clean(row.discipline),clean(row.upn),clean(row.systemName),clean(row.building),clean(row.itemMaster)];
    const findings=line.findings.length?line.findings:[null];
    /* Ticked from the start only when the equipment has findings and every one
       of them is in a pre-ticked group. */
    const startTicked=line.findings.length>0&&preticked&&line.findings.every(finding=>preticked.has(finding));
    findings.forEach((finding,offset)=>{
      aoa.push([offset===0?(startTicked?AUDIT_EXPORT_TICK:AUDIT_EXPORT_UNTICKED):'',line.level,...equipment,
        finding?AUDIT_EXPORT_SEVERITY_LABELS[finding.severity]||finding.severity.toUpperCase():'',
        finding?clean(finding.rule.title):'',finding?clean(finding.why):'',finding?clean(finding.recommendation):'','','']);
      meta.push({level:line.level,repeat:offset>0,severity:finding?finding.severity:'',flag:finding?AUDIT_EXPORT_FIELD_COLUMNS[clean(finding.field)]||'':''});
    });
  }
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=AUDIT_EXPORT_MILESTONE_WIDTHS.map(width=>({wch:width}));
  sheetStyleCell(sheet,'A1',AUDIT_EXPORT_STYLES.back);
  sheetLinkCell(sheet,'A1',`#${auditExportSheetRef(AUDIT_EXPORT_INDEX_SHEET)}!A1`,'Back to the milestone index');
  sheetStyleCell(sheet,'B1',AUDIT_EXPORT_STYLES.note);
  auditExportHeaderRow(sheet,2,AUDIT_EXPORT_MILESTONE_HEADERS,AUDIT_EXPORT_MILESTONE_ALIGN);
  meta.forEach((entry,offset)=>{
    const rowIndex=offset+3,nest=AUDIT_EXPORT_NEST_STYLES[entry.level%AUDIT_EXPORT_NEST_STYLES.length];
    const body=entry.repeat?AUDIT_EXPORT_STYLES.repeat:AUDIT_EXPORT_STYLES.text;
    sheetStyleCell(sheet,`A${rowIndex}`,entry.repeat?AUDIT_EXPORT_STYLES.actionedRepeat:AUDIT_EXPORT_STYLES.actioned);
    sheetStyleCell(sheet,`B${rowIndex}`,nest.level);
    sheetStyleCell(sheet,`C${rowIndex}`,entry.repeat?nest.repeat:nest.own);
    for(const column of AUDIT_EXPORT_ROW_COLUMNS)sheetStyleCell(sheet,`${column}${rowIndex}`,body);
    /* The cell the finding is about turns light red -- on the finding's own line,
       even when that is the equipment's ID cell. */
    if(entry.flag)sheetStyleCell(sheet,`${entry.flag}${rowIndex}`,AUDIT_EXPORT_STYLES.flag);
    if(entry.severity)sheetStyleCell(sheet,`L${rowIndex}`,AUDIT_EXPORT_SEVERITY_STYLES[entry.severity]||AUDIT_EXPORT_STYLES.text);
    const detail=entry.severity?AUDIT_EXPORT_STYLES.wrap:AUDIT_EXPORT_STYLES.wrapMuted;
    sheetStyleCell(sheet,`M${rowIndex}`,entry.severity?AUDIT_EXPORT_STYLES.text:AUDIT_EXPORT_STYLES.muted);
    sheetStyleCell(sheet,`N${rowIndex}`,detail);
    sheetStyleCell(sheet,`O${rowIndex}`,detail);
    sheetStyleCell(sheet,`P${rowIndex}`,AUDIT_EXPORT_STYLES.entry);
    sheetStyleCell(sheet,`Q${rowIndex}`,AUDIT_EXPORT_STYLES.entry);
  });
  const lastRow=Math.max(3,meta.length+2),lastColumn=auditColumnName(AUDIT_EXPORT_MILESTONE_HEADERS.length-1);
  sheetFreezeRows(sheet,2);
  sheetAutoFilter(sheet,`A2:${lastColumn}${Math.max(2,meta.length+2)}`);
  /* Tick-box dropdown on every Actioned cell; a ticked line turns green. */
  sheetXmlExtras(sheet,{dataValidations:[auditExportTickValidation(`A3:A${lastRow}`)],conditionalFormatting:[auditExportDoneFormat(`A3:${lastColumn}${lastRow}`,3)]});
  return sheet;
}

function auditExportFindingsSheet(result,groups){
  const milestoneByRow=new Map(),descriptionByRow=new Map();
  for(const group of groups)for(const line of group.lines){
    const key=auditExportRowKey(line.row);
    milestoneByRow.set(key,group.label);descriptionByRow.set(key,clean(line.row.equipmentDescription));
  }
  const aoa=[[AUDIT_EXPORT_BACK_LINK],[...AUDIT_EXPORT_FINDING_HEADERS]];
  const findings=result&&result.findings||[];
  for(const finding of findings){
    const key=auditExportFindingKey(finding);
    aoa.push([AUDIT_EXPORT_SEVERITY_LABELS[finding.severity]||finding.severity.toUpperCase(),milestoneByRow.get(key)||AUDIT_EXPORT_NO_MILESTONE,clean(finding.equipmentId),descriptionByRow.get(key)||'',
      clean(finding.rule.title),clean(finding.why),clean(finding.recommendation),clean(finding.field),printable(finding.actual),printable(finding.expected),clean(finding.sheet),finding.row||0]);
  }
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=AUDIT_EXPORT_FINDING_WIDTHS.map(width=>({wch:width}));
  auditExportBackLink(sheet);
  auditExportHeaderRow(sheet,2,AUDIT_EXPORT_FINDING_HEADERS,['center','left','left','left','left','left','left','left','left','left','left','right']);
  findings.forEach((finding,offset)=>{
    const rowIndex=offset+3;
    sheetStyleCell(sheet,`A${rowIndex}`,AUDIT_EXPORT_SEVERITY_STYLES[finding.severity]||AUDIT_EXPORT_STYLES.text);
    for(const column of ['B','C','D','E','H','I','J','K'])sheetStyleCell(sheet,`${column}${rowIndex}`,AUDIT_EXPORT_STYLES.text);
    sheetStyleCell(sheet,`F${rowIndex}`,AUDIT_EXPORT_STYLES.wrap);
    sheetStyleCell(sheet,`G${rowIndex}`,AUDIT_EXPORT_STYLES.wrap);
    sheetStyleCell(sheet,`L${rowIndex}`,AUDIT_EXPORT_STYLES.number);
  });
  sheetFreezeRows(sheet,2);
  sheetAutoFilter(sheet,`A2:${auditColumnName(AUDIT_EXPORT_FINDING_HEADERS.length-1)}${Math.max(2,findings.length+2)}`);
  return sheet;
}

function auditExportRulesSheet(result){
  const counts=new Map(),catalog=new Map();
  for(const rule of Object.values(SSM_AUDIT_RULES))if(rule.enabled)catalog.set(rule.id,rule);
  for(const finding of result&&result.findings||[]){
    counts.set(finding.rule.id,(counts.get(finding.rule.id)||0)+1);
    if(!catalog.has(finding.rule.id))catalog.set(finding.rule.id,finding.rule);
  }
  const entries=[...catalog.values()].sort((left,right)=>(counts.get(right.id)||0)-(counts.get(left.id)||0)||natCmp(left.title,right.title));
  const aoa=[[AUDIT_EXPORT_BACK_LINK],[...AUDIT_EXPORT_RULE_HEADERS]];
  for(const rule of entries)aoa.push([clean(rule.title),clean(rule.statement),EXPORT_SOURCE_LABELS[rule.source]||rule.source,EXPORT_CONFIDENCE_LABELS[rule.confidence]||rule.confidence,counts.get(rule.id)||0]);
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=AUDIT_EXPORT_RULE_WIDTHS.map(width=>({wch:width}));
  auditExportBackLink(sheet);
  auditExportHeaderRow(sheet,2,AUDIT_EXPORT_RULE_HEADERS,['left','left','left','left','right']);
  entries.forEach((rule,offset)=>{
    const rowIndex=offset+3;
    sheetStyleCell(sheet,`A${rowIndex}`,AUDIT_EXPORT_STYLES.text);
    sheetStyleCell(sheet,`B${rowIndex}`,AUDIT_EXPORT_STYLES.wrap);
    sheetStyleCell(sheet,`C${rowIndex}`,AUDIT_EXPORT_STYLES.text);
    sheetStyleCell(sheet,`D${rowIndex}`,AUDIT_EXPORT_STYLES.text);
    sheetStyleCell(sheet,`E${rowIndex}`,AUDIT_EXPORT_STYLES.number);
  });
  sheetFreezeRows(sheet,2);
  sheetAutoFilter(sheet,`A2:${auditColumnName(AUDIT_EXPORT_RULE_HEADERS.length-1)}${Math.max(2,entries.length+2)}`);
  return sheet;
}

export function buildAuditWorkbook(sourceResult,sessionName,options={}){
  const {result,preticked}=auditExportApplyPlan(sourceResult,options.plan);
  const workbook=XLSX.utils.book_new(),groups=auditExportGroups(result);
  const used=new Set([AUDIT_EXPORT_DASHBOARD_SHEET,AUDIT_EXPORT_INDEX_SHEET,AUDIT_EXPORT_FINDINGS_SHEET,AUDIT_EXPORT_RULES_SHEET,AUDIT_EXPORT_CALC_SHEET].map(name=>name.toLowerCase()));
  for(const group of groups)group.sheetName=auditExportSheetName(group.label,used);
  const generated=options.generatedAt instanceof Date?options.generatedAt:new Date();
  workbook.Dxfs=[...AUDIT_EXPORT_DXFS];
  const disciplines=auditExportDisciplines(result);
  addSheet(workbook,auditExportDashboardSheet(result,groups,disciplines,sessionName,generated),AUDIT_EXPORT_DASHBOARD_SHEET);
  addSheet(workbook,auditExportIndexSheet(groups),AUDIT_EXPORT_INDEX_SHEET);
  for(const group of groups)addSheet(workbook,auditExportMilestoneSheet(group,preticked),group.sheetName);
  addSheet(workbook,auditExportFindingsSheet(result,groups),AUDIT_EXPORT_FINDINGS_SHEET);
  addSheet(workbook,auditExportRulesSheet(result),AUDIT_EXPORT_RULES_SHEET);
  addSheet(workbook,auditExportCalcSheet(groups,disciplines),AUDIT_EXPORT_CALC_SHEET);
  workbook.Workbook=workbook.Workbook||{};workbook.Workbook.Sheets=workbook.SheetNames.map(name=>({name,Hidden:name===AUDIT_EXPORT_CALC_SHEET?1:0}));
  return workbook;
}

export async function exportSsmAuditXlsx(plan){
  const result=S.session&&S.session.result;if(!result){toast('Run an SSM Audit first');return;}
  const base=clean(S.session.name).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  /* A milestone workbook carries every equipment row and every finding. It is
     built then zipped with a real deflate, which takes a few seconds on a large
     registry, so the progress overlay stays up for the whole job. */
  try{
    await runWithProgress('Building the Excel report',S.session.name,async(checkpoint,report)=>{
      report(.1);await checkpoint();
      const workbook=buildAuditWorkbook(result,S.session.name,{plan});
      report(.45);await checkpoint();
      const blob=await workbookBlobCompact(workbook);
      report(1);downloadBlob(`${base}-Audit.xlsx`,blob);
    });
    toast('SSM Audit report exported');
  }catch(error){console.error('SSM Audit export failed',error);toast('The report could not be built');}
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

  const pairRows=[['UPN','Status','Match Basis','Target Equipment ID','Target Description','Target Parent ID','Target Parent Type','Target Header','Target Dependencies','Completed Project Equipment ID','Completed Project Description','Completed Project Parent ID','Completed Project Parent Type','Completed Project Header','Completed Project Dependencies','Placement Mismatch','Parent Cycle','Differences']];
  for(const system of result.systems)for(const pair of system.pairs){const target=pair.target,reference=pair.reference;pairRows.push([system.upn,pair.status,pair.matchReason,target&&target.tag||'',target&&target.role||'',target&&target.parentId||'',target&&target.parentRole||'',target&&target.headerName||'',target&&target.dependencyRoles.join('; ')||'',reference&&reference.tag||'',reference&&reference.role||'',reference&&reference.parentId||'',reference&&reference.parentRole||'',reference&&reference.headerName||'',reference&&reference.dependencyRoles.join('; ')||'',pair.placementMismatch?'Yes':'No',pair.cycle?'Yes':'No',pair.differences.join('; ')]);}
  const pairSheet=XLSX.utils.aoa_to_sheet(pairRows);pairSheet['!cols']=[{wch:12},{wch:20},{wch:38},{wch:30},{wch:34},{wch:30},{wch:34},{wch:30},{wch:40},{wch:36},{wch:34},{wch:36},{wch:34},{wch:30},{wch:40},{wch:20},{wch:14},{wch:54}];styleHeaderRow(pairSheet);pairSheet['!autofilter']={ref:`A1:R${Math.max(1,pairRows.length)}`};addSheet(workbook,pairSheet,'Equipment Mapping');

  const base=clean(comparison.targetName).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  downloadBlob(`${base}-Project-Comparison.xlsx`,workbookBlob(workbook));toast('Project comparison exported');
}
