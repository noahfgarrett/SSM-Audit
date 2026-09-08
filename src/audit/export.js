import { clean, natCmp } from '../core/text.js'
import { zipDeflateAvailable, zipEntries } from '../core/zip.js'
import { downloadBlob, sheetAutoFilter, sheetCellStyle, sheetFormulaCell, sheetFreezeRows, sheetLinkCell, sheetSetCell, sheetStyleCell, sheetXmlExtras, styleHeaderRow, workbookBlob, workbookBlobCompact } from '../core/download.js'
import { EXTO_REV21_COLUMNS, extoRev21Norm } from '../exto/rev21-contract.js'
import { S } from '../state.js'
import { runWithProgress, toast } from '../ui/feedback.js'
import { SSM_AUDIT_RULES } from './engine.js'
import { auditColumnName, auditNormId } from './model.js'

function addSheet(workbook,sheet,name){XLSX.utils.book_append_sheet(workbook,sheet,name);}
function printable(value){return typeof value==='string'?value:JSON.stringify(value);}
const EXPORT_SOURCE_LABELS={registry:'Registry Integrity',sop:'SSM SOP',logic:'Commissioning Logic',reference:'Selected references'},EXPORT_CONFIDENCE_LABELS={required:'Required',strong:'Strong pattern','description-rated':'Description based'};

/* ---- SSM Audit workbook ----
   One tab per L2 milestone so a commissioning engineer can work a milestone end
   to end: the full equipment tree of that milestone, one line per finding, and an
   `Actioned` column with a tick-box dropdown (☐ / ☑). Index and Dashboard read
   those ticks back with live COUNTIF formulas, so progress updates itself. */
const AUDIT_EXPORT_DASHBOARD_SHEET='Dashboard',AUDIT_EXPORT_INDEX_SHEET='Index',AUDIT_EXPORT_FINDINGS_SHEET='All Findings',AUDIT_EXPORT_RULES_SHEET='Rules',AUDIT_EXPORT_CALC_SHEET='Calc';
const AUDIT_EXPORT_NO_MILESTONE='No milestone',AUDIT_EXPORT_BACK_LINK='← Index';
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
/* Finding-tab columns, computed per layout: the level layout adds an
   L2 Milestone column after Item Master. The last column, Done, is a hidden
   helper that carries each equipment block's tick down to its finding lines so
   the whole block can turn green (the tick lives in one merged cell, and a
   conditional format on another row cannot read a merged neighbour). */
export function auditExportTabColumns(layout){
  const headers=['Actioned','Nest','Equipment ID','Description','Closest Parent','Dependencies','Discipline','UPN','System Name','Building','Item Master'];
  const widths=[10,6,36,38,28,34,22,8,30,10,26],align=['center','center','left','left','left','left','left','right','left','left','left'];
  if(layout==='level'){headers.push('L2 Milestone');widths.push(30);align.push('left');}
  headers.push('Severity','Finding','Why','What to do','Actioned By','Note','Done');
  widths.push(12,32,54,54,16,30,6);align.push('center','left','left','left','left','left','center');
  const index=Object.fromEntries(headers.map((header,at)=>[header,at]));
  const letterOf=header=>auditColumnName(index[header]);
  return {headers,widths,align,index,letterOf,equipmentEnd:index['Item Master']+(layout==='level'?1:0),doneIndex:index['Done'],lastVisibleIndex:index['Note']};
}
/* Which column a finding's `field` points at, so that cell can be shaded. */
const AUDIT_EXPORT_FIELD_HEADERS=Object.freeze({'Equipment ID':'Equipment ID','Closest Parent':'Closest Parent','Dependencies':'Dependencies','Dependency Project':'Dependencies','Discipline':'Discipline','UPN':'UPN','System Name':'System Name','Building':'Building','Item Master Unique Identifier':'Item Master'});
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
  overallBar:sheetCellStyle({fill:AUDIT_EXPORT_PALETTE.barTrack,align:'left',vertical:'center',numFmt:';;;'}),
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
  bar:sheetCellStyle({align:'left',numFmt:';;;'}),
  barBand:sheetCellStyle({fill:AUDIT_EXPORT_PALETTE.band,align:'left',numFmt:';;;'}),
  barWide:sheetCellStyle({align:'left',vertical:'center',numFmt:';;;'}),
  barWideBand:sheetCellStyle({fill:AUDIT_EXPORT_PALETTE.band,align:'left',vertical:'center',numFmt:';;;'}),
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

function auditExportEmptyBar(){return 0;}
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

/* One group per finding level: only flagged equipment, ordered by milestone
   then tag, each entry carrying just that level's findings. */
export function auditExportLevelGroups(result){
  const rows=result&&result.rows||[],findings=result&&result.findings||[],levelFor=auditExportNestLevels(rows);
  const rowByKey=new Map();for(const row of rows)rowByKey.set(auditExportRowKey(row),row);
  const groups=[];
  for(const severity of ['blocker','error','warning','info']){
    const severityFindings=findings.filter(finding=>finding.severity===severity);if(!severityFindings.length)continue;
    const byEquipment=new Map();
    for(const finding of severityFindings){const key=auditExportFindingKey(finding),entry=byEquipment.get(key)||{row:rowByKey.get(key)||null,findings:[]};entry.findings.push(finding);byEquipment.set(key,entry);}
    const lines=[...byEquipment.values()].filter(entry=>entry.row).map(entry=>({row:entry.row,level:levelFor(entry.row),findings:entry.findings}))
      .sort((left,right)=>natCmp(clean(left.row.milestone),clean(right.row.milestone))||natCmp(clean(left.row.equipmentId),clean(right.row.equipmentId)));
    groups.push({label:AUDIT_EXPORT_SEVERITY_LABELS[severity],severity,sheetName:'',rows:lines.map(line=>line.row),lines,equipmentCount:lines.length,findingCount:severityFindings.length});
  }
  return groups;
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
/* Progress bars are native Excel data bars on a percent cell. The cell shows the
   BAR ONLY (showValue off, and a blank number format for apps that ignore it);
   the % column next to it carries the number. */
function auditExportBarCell(percentCell,style){return sheetFormulaCell(percentCell,0,style);}
function auditExportDataBar(range,priority){return `<conditionalFormatting sqref="${range}"><cfRule type="dataBar" priority="${priority}"><dataBar minLength="0" maxLength="100" showValue="0"><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF${AUDIT_EXPORT_PALETTE.accent}"/></dataBar></cfRule></conditionalFormatting>`;}
function auditExportTickValidation(range){return `<dataValidation type="list" allowBlank="1" showDropDown="0" showErrorMessage="1" errorTitle="Actioned" error="Pick ${AUDIT_EXPORT_TICK} or ${AUDIT_EXPORT_UNTICKED} from the list." sqref="${range}"><formula1>"${AUDIT_EXPORT_TICK},${AUDIT_EXPORT_UNTICKED}"</formula1></dataValidation>`;}
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
export function auditExportApplyPlan(result,plan,actionedIds){
  const findings=[],preticked=new Set(),actioned=new Set(actionedIds||[]);
  for(const finding of result&&result.findings||[]){
    const mode=auditExportPlanMode(plan,finding);
    if(mode==='skip')continue;
    findings.push(finding);if(mode==='pretick'||actioned.has(finding.id))preticked.add(finding);
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
function auditExportDashboardSheet(result,groups,disciplines,sessionName,generated,groupLabel){
  const summary=result&&result.summary||{},severity=summary.severity||{};
  const columns=['Discipline','Equipment','Findings','Actioned','%','Progress'];
  const aoa=[
    [`SSM Audit — ${clean(sessionName)||'Registry'}`,'','','','',''],
    [`Generated ${auditExportDate(generated)}`,'','','','',''],
    [`Standard: ${clean(result&&result.standard)}`,'','','','',''],
    ['','','','','',''],
    [`OVERALL PROGRESS — equipment actioned across every ${groupLabel==='Finding level'?'finding level':'milestone'}`,'','','','',''],
    [auditExportEmptyBar(),'','','','',0],
    ['','','','','',''],
    ['Rows audited','Findings','Invalid','Rule broken','Check this','Notes'],
    [summary.rows||0,summary.findings||0,severity.blocker||0,severity.error||0,severity.warning||0,severity.info||0],
    ['','','','','',''],
    columns,
  ];
  const disciplineFirst=aoa.length+1;
  for(const discipline of disciplines)aoa.push([discipline.label,discipline.equipmentCount,discipline.findingCount,0,0,auditExportEmptyBar()]);
  if(!disciplines.length)aoa.push(['No equipment rows in this registry','','','','','']);
  const disciplineLast=disciplineFirst+Math.max(0,disciplines.length-1);
  aoa.push(['','','','','','']);
  aoa.push([groupLabel,'Equipment','Findings','Actioned','%','Progress']);
  const milestoneHeader=aoa.length,milestoneFirst=aoa.length+1;
  for(const group of groups)aoa.push([group.label,group.equipmentCount,group.findingCount,0,0,auditExportEmptyBar()]);
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
  if(overallPercent)sheetSetCell(sheet,'A6',auditExportBarCell('F6',AUDIT_EXPORT_STYLES.overallBar));else sheetStyleCell(sheet,'A6',AUDIT_EXPORT_STYLES.overallBar);
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
    sheetSetCell(sheet,`F${rowIndex}`,auditExportBarCell(`E${rowIndex}`,band?AUDIT_EXPORT_STYLES.barWideBand:AUDIT_EXPORT_STYLES.barWide));
  };
  auditExportHeaderRow(sheet,disciplineFirst-1,columns,['left','right','right','right','right','left']);
  disciplines.forEach((discipline,offset)=>{
    const rowIndex=disciplineFirst+offset,band=offset%2===1;
    progressRow(rowIndex,band,auditExportDisciplineActionedCell(offset,groups.length,band?AUDIT_EXPORT_STYLES.numberMidBand:AUDIT_EXPORT_STYLES.numberMid),null);
  });
  auditExportHeaderRow(sheet,milestoneHeader,[groupLabel,'Equipment','Findings','Actioned','%','Progress'],['left','right','right','right','right','left']);
  groups.forEach((group,offset)=>{
    const rowIndex=milestoneFirst+offset,band=offset%2===1;
    progressRow(rowIndex,band,auditExportActionedCell(group,band?AUDIT_EXPORT_STYLES.numberMidBand:AUDIT_EXPORT_STYLES.numberMid),{target:`#${auditExportSheetRef(group.sheetName)}!A1`,tooltip:`Open ${group.label}`});
  });
  const bars=[auditExportDataBar('A6:A6',1)];
  if(disciplines.length)bars.push(auditExportDataBar(`F${disciplineFirst}:F${disciplineLast}`,2));
  if(groups.length)bars.push(auditExportDataBar(`F${milestoneFirst}:F${milestoneLast}`,3));
  sheetXmlExtras(sheet,{conditionalFormatting:bars});
  return sheet;
}

function auditExportIndexSheet(groups,groupLabel){
  const firstRow=5;
  const aoa=[
    ['SSM Audit — Index','','','','',''],
    [AUDIT_EXPORT_DASHBOARD_SHEET,AUDIT_EXPORT_FINDINGS_SHEET,AUDIT_EXPORT_RULES_SHEET,'','',''],
    ['','','','','',''],
    [groupLabel,'Equipment','Findings','Actioned','%','Progress'],
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
  auditExportHeaderRow(sheet,4,[groupLabel,'Equipment','Findings','Actioned','%','Progress'],['left','right','right','right','right','left']);
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
  if(groups.length)sheetXmlExtras(sheet,{conditionalFormatting:[auditExportDataBar(`F${firstRow}:F${firstRow+groups.length-1}`,1)]});
  sheetFreezeRows(sheet,4);
  return sheet;
}

function auditExportFindingTab(group,preticked,layout){
  const columns=auditExportTabColumns(layout),blank=()=>new Array(columns.headers.length).fill('');
  const aoa=[[AUDIT_EXPORT_BACK_LINK,AUDIT_EXPORT_ACTIONED_NOTE],[...columns.headers]];
  const meta=[],blocks=[];
  for(const line of group.lines){
    const row=line.row,indent='  '.repeat(Math.min(line.level,24));
    const findings=line.findings.length?line.findings:[null];
    /* Ticked from the start only when the equipment has findings and every one
       of them is in a pre-ticked group. */
    const startTicked=line.findings.length>0&&preticked&&line.findings.every(finding=>preticked.has(finding));
    const first=aoa.length+1;
    findings.forEach((finding,offset)=>{
      const cells=blank();
      if(offset===0){
        cells[0]=startTicked?AUDIT_EXPORT_TICK:AUDIT_EXPORT_UNTICKED;cells[1]=line.level;
        cells[2]=indent+clean(row.equipmentId);cells[3]=clean(row.equipmentDescription);cells[4]=clean(row.closestParent);cells[5]=clean(row.dependencies);
        cells[6]=clean(row.discipline);cells[7]=clean(row.upn);cells[8]=clean(row.systemName);cells[9]=clean(row.building);cells[10]=clean(row.itemMaster);
        if(layout==='level')cells[columns.index['L2 Milestone']]=clean(row.milestone)||AUDIT_EXPORT_NO_MILESTONE;
      }
      if(finding){
        cells[columns.index['Severity']]=AUDIT_EXPORT_SEVERITY_LABELS[finding.severity]||finding.severity.toUpperCase();
        cells[columns.index['Finding']]=clean(finding.rule.title);cells[columns.index['Why']]=clean(finding.why);cells[columns.index['What to do']]=clean(finding.recommendation);
      }
      aoa.push(cells);
      meta.push({level:line.level,anchor:offset===0,severity:finding?finding.severity:'',flagHeader:finding?AUDIT_EXPORT_FIELD_HEADERS[clean(finding.field)]||'':'',blockFirst:first});
    });
    blocks.push({first,count:findings.length});
  }
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=columns.widths.map((width,at)=>at===columns.doneIndex?{wch:width,hidden:true}:{wch:width});
  sheetStyleCell(sheet,'A1',AUDIT_EXPORT_STYLES.back);
  sheetLinkCell(sheet,'A1',`#${auditExportSheetRef(AUDIT_EXPORT_INDEX_SHEET)}!A1`,'Back to the index');
  sheetStyleCell(sheet,'B1',AUDIT_EXPORT_STYLES.note);
  auditExportHeaderRow(sheet,2,columns.headers,columns.align);
  /* An equipment with several findings is ONE merged block: the equipment cells
     span its finding lines and one Actioned box covers them all. */
  sheet['!merges']=sheet['!merges']||[];
  for(const block of blocks){
    if(block.count<2)continue;
    for(let columnIndex=0;columnIndex<=columns.equipmentEnd;columnIndex++)
      sheet['!merges'].push({s:{r:block.first-1,c:columnIndex},e:{r:block.first-2+block.count,c:columnIndex}});
  }
  const severityLetter=columns.letterOf('Severity'),findingLetter=columns.letterOf('Finding'),whyLetter=columns.letterOf('Why'),toDoLetter=columns.letterOf('What to do'),byLetter=columns.letterOf('Actioned By'),noteLetter=columns.letterOf('Note'),doneLetter=auditColumnName(columns.doneIndex);
  meta.forEach((entry,offset)=>{
    const rowIndex=offset+3,nest=AUDIT_EXPORT_NEST_STYLES[entry.level%AUDIT_EXPORT_NEST_STYLES.length];
    if(entry.anchor){
      sheetStyleCell(sheet,`A${rowIndex}`,AUDIT_EXPORT_STYLES.actioned);
      sheetStyleCell(sheet,`B${rowIndex}`,nest.level);
      sheetStyleCell(sheet,`C${rowIndex}`,nest.own);
      for(let columnIndex=3;columnIndex<=columns.equipmentEnd;columnIndex++)sheetStyleCell(sheet,`${auditColumnName(columnIndex)}${rowIndex}`,AUDIT_EXPORT_STYLES.text);
    }
    /* The cell the finding is about turns light red. The equipment cells are one
       merged block, so the shade lands on the block's (anchor) cell. */
    if(entry.flagHeader)sheetStyleCell(sheet,`${columns.letterOf(entry.flagHeader)}${entry.blockFirst}`,AUDIT_EXPORT_STYLES.flag);
    if(entry.severity)sheetStyleCell(sheet,`${severityLetter}${rowIndex}`,AUDIT_EXPORT_SEVERITY_STYLES[entry.severity]||AUDIT_EXPORT_STYLES.text);
    const detail=entry.severity?AUDIT_EXPORT_STYLES.wrap:AUDIT_EXPORT_STYLES.wrapMuted;
    sheetStyleCell(sheet,`${findingLetter}${rowIndex}`,entry.severity?AUDIT_EXPORT_STYLES.text:AUDIT_EXPORT_STYLES.muted);
    sheetStyleCell(sheet,`${whyLetter}${rowIndex}`,detail);
    sheetStyleCell(sheet,`${toDoLetter}${rowIndex}`,detail);
    sheetStyleCell(sheet,`${byLetter}${rowIndex}`,AUDIT_EXPORT_STYLES.entry);
    sheetStyleCell(sheet,`${noteLetter}${rowIndex}`,AUDIT_EXPORT_STYLES.entry);
    /* Done carries the block's tick down its lines for the green format. */
    const formula=rowIndex===3?`IF(A3<>"",A3,"")`:`IF(A${rowIndex}<>"",A${rowIndex},${doneLetter}${rowIndex-1})`;
    sheetSetCell(sheet,`${doneLetter}${rowIndex}`,sheetFormulaCell(formula,entry.anchor?AUDIT_EXPORT_UNTICKED:'',null));
  });
  const lastRow=Math.max(3,meta.length+2),lastVisible=auditColumnName(columns.lastVisibleIndex);
  sheetFreezeRows(sheet,2);
  sheetAutoFilter(sheet,`A2:${lastVisible}${Math.max(2,meta.length+2)}`);
  /* Tick-box dropdown on every Actioned cell; the block turns green when its
     Done helper column carries the tick. */
  sheetXmlExtras(sheet,{dataValidations:[auditExportTickValidation(`A3:A${lastRow}`)],
    conditionalFormatting:[`<conditionalFormatting sqref="A3:${lastVisible}${lastRow}"><cfRule type="expression" dxfId="0" priority="1"><formula>$${doneLetter}3="${AUDIT_EXPORT_TICK}"</formula></cfRule></conditionalFormatting>`]});
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

function auditExportRulesSheet(result,disabledRules){
  const counts=new Map(),catalog=new Map(),off=new Set(disabledRules||[]);
  for(const rule of Object.values(SSM_AUDIT_RULES))if(rule.enabled&&!off.has(rule.id))catalog.set(rule.id,rule);
  for(const finding of result&&result.findings||[]){
    if(off.has(finding.rule.id))continue;
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
  const {result,preticked}=auditExportApplyPlan(sourceResult,options.plan,options.actionedIds);
  const layout=options.layout==='level'?'level':'milestone',groupLabel=layout==='level'?'Finding level':'Milestone';
  const workbook=XLSX.utils.book_new(),groups=layout==='level'?auditExportLevelGroups(result):auditExportGroups(result);
  const used=new Set([AUDIT_EXPORT_DASHBOARD_SHEET,AUDIT_EXPORT_INDEX_SHEET,AUDIT_EXPORT_FINDINGS_SHEET,AUDIT_EXPORT_RULES_SHEET,AUDIT_EXPORT_CALC_SHEET].map(name=>name.toLowerCase()));
  for(const group of groups)group.sheetName=auditExportSheetName(group.label,used);
  const generated=options.generatedAt instanceof Date?options.generatedAt:new Date();
  workbook.Dxfs=[...AUDIT_EXPORT_DXFS];
  const disciplines=auditExportDisciplines(result);
  addSheet(workbook,auditExportDashboardSheet(result,groups,disciplines,sessionName,generated,groupLabel),AUDIT_EXPORT_DASHBOARD_SHEET);
  addSheet(workbook,auditExportIndexSheet(groups,groupLabel),AUDIT_EXPORT_INDEX_SHEET);
  for(const group of groups)addSheet(workbook,auditExportFindingTab(group,preticked,layout),group.sheetName);
  addSheet(workbook,auditExportFindingsSheet(result,groups),AUDIT_EXPORT_FINDINGS_SHEET);
  addSheet(workbook,auditExportRulesSheet(result,options.disabledRules),AUDIT_EXPORT_RULES_SHEET);
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
      report(.04,'Collecting rows and findings');await checkpoint();
      const workbook=buildAuditWorkbook(result,S.session.name,{plan,layout:plan&&plan.layout,disabledRules:S.rules.disabled,actionedIds:[...(S.session.actioned||[])]});
      /* The next stretch is SheetJS serializing the workbook XML in one
         synchronous run -- the label paints first so the loader is honest
         about the wait on a large registry. */
      report(.3,'Writing the workbook');await checkpoint();
      const blob=await workbookBlobCompact(workbook,{onProgress:async(fraction,name)=>{report(.55+fraction*.42,`Compressing ${name.replace(/^xl\//,'').replace(/\.xml$/,'')}`);await checkpoint();}});
      report(1,'Report ready');downloadBlob(`${base}-Audit.xlsx`,blob);
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

/* ---- Updated Registry Export ----
   Preflight against the immutable import, never the corrected draft. Package
   output retains untouched part payloads, including forms and relationships.
   Target worksheet XML is DOM-serialized; ZIP bytes/metadata and XML spelling
   are not preserved exactly. Formula inputs may require Excel recalculation. */
export class AuditCorrectionExportError extends Error{
  constructor(code,message,index){super(index==null?message:`Correction ${index+1}: ${message}`);this.name='AuditCorrectionExportError';this.code=code;this.changeIndex=index;}
}
function auditCorrectionFail(code,message,index){throw new AuditCorrectionExportError(code,message,index);}
function auditCorrectionSourceKey(source){return JSON.stringify([source&&source.sheet,source&&source.row]);}
function auditCorrectionRows(snapshot){return snapshot&&Array.isArray(snapshot.snapshots)?snapshot.snapshots.flatMap(auditCorrectionRows):snapshot&&snapshot.rows||[];}
/* Import records trimmed display values, not raw storage values. Format a copy
   without cached text so validation neither mutates cells nor trusts stale w. */
function auditCorrectionCellValue(cell){return cell?clean(XLSX.utils.format_cell({...cell,w:undefined})):'';}
function auditCorrectionColumn(change,index){
  const aliases={'L2 Milestone':'milestone','L1 Milestone Parent':'milestoneParent'};
  const label=change.header||change.field;
  const column=EXTO_REV21_COLUMNS.find(column=>change.prop?column.field===change.prop:column.field===label||extoRev21Norm(column.header)===extoRev21Norm(label)||column.field===aliases[label]);
  if(!column)auditCorrectionFail('FIELD','The field is not an exported registry column.',index);
  if(change.header&&extoRev21Norm(change.header)!==extoRev21Norm(column.header))auditCorrectionFail('FIELD','The header and property identify different columns.',index);
  return column;
}
function auditCorrectionValue(value,index){
  if(value===null)return '';
  if(typeof value==='boolean'||typeof value==='number'&&Number.isFinite(value))return value;
  if(typeof value!=='string'||value.length>32767||/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(value))auditCorrectionFail('VALUE','The replacement is not a supported Excel cell value.',index);
  return value;
}
/* A tag match is not evidence of a mirror. Only the import's occurrence-paired
   cross-sheet sources may expand a canonical correction, and each must still
   match its complete original registry row before any cell is written. */
function auditCorrectionMirrorRows(workbook,canonical,bySource,mirrorOwners,index){
  const sources=canonical._sources;
  if(!Array.isArray(sources)||!sources.length)auditCorrectionFail('SOURCE','The imported mirror locations are invalid.',index);
  const seenSheets=new Set(),rows=[],canonicalKey=auditCorrectionSourceKey(canonical._source);
  for(const source of sources){
    const key=auditCorrectionSourceKey(source),matches=bySource.get(key);
    if(!source||!source.sheet||!Number.isInteger(source.row)||source.row<1||source.row>1048576||!matches||matches.length!==1)auditCorrectionFail('SOURCE','A mirrored row has no unique original physical location.',index);
    if(seenSheets.has(source.sheet)||mirrorOwners.get(key)?.size!==1)auditCorrectionFail('AMBIGUOUS','Mirrored copies must belong to one occurrence on distinct worksheets.',index);
    seenSheets.add(source.sheet);
    const row=matches[0],columns=row._source.columns,declared=source.columns;
    if(!columns||!declared||Object.keys(columns).length!==Object.keys(declared).length||Object.entries(columns).some(([prop,column])=>declared[prop]!==column))auditCorrectionFail('COLUMN','A mirrored row column mapping differs from the original import.',index);
    if(EXTO_REV21_COLUMNS.some(column=>clean(row[column.field])!==clean(canonical[column.field])))auditCorrectionFail('BASELINE','A mirrored row is not identical to its canonical baseline row.',index);
    const sheet=workbook.Sheets[source.sheet];if(!sheet)auditCorrectionFail('SHEET','An original mirrored worksheet is missing.',index);
    for(const field of EXTO_REV21_COLUMNS){
      const column=columns[field.field];if(column==null)continue;
      if(!Number.isInteger(column)||column<0||column>16383)auditCorrectionFail('COLUMN','A mirrored row has an invalid original column mapping.',index);
      const cell=sheet[XLSX.utils.encode_cell({r:source.row-1,c:column})];
      if(auditCorrectionCellValue(cell)!==clean(row[field.field]))auditCorrectionFail('CONFLICT','An original mirrored row has changed; no corrections were written.',index);
    }
    rows.push(row);
  }
  if(!rows.some(row=>auditCorrectionSourceKey(row._source)===canonicalKey))auditCorrectionFail('SOURCE','The canonical row is missing from its imported mirror group.',index);
  return rows;
}
function auditCorrectionPreflight(workbook,snapshot,changes){
  if(!snapshot||!Array.isArray(snapshot.rows)||!workbook||!workbook.Sheets||!Array.isArray(changes))auditCorrectionFail('BASELINE','The original baseline and a correction list are required.');
  const bySource=new Map(),byTag=new Map();
  for(const row of auditCorrectionRows(snapshot)){
    const sourceKey=auditCorrectionSourceKey(row._source),tag=auditNormId(row.equipmentId);
    const sources=bySource.get(sourceKey)||[];sources.push(row);bySource.set(sourceKey,sources);
    if(tag){const rows=byTag.get(tag)||[];rows.push(row);byTag.set(tag,rows);}
  }
  const canonicalBySource=new Map(),mirrorOwners=new Map();
  for(const row of snapshot.rows){
    const key=auditCorrectionSourceKey(row._source),matches=canonicalBySource.get(key)||[];matches.push(row);canonicalBySource.set(key,matches);
    for(const source of Array.isArray(row._sources)?row._sources:[]){const sourceKey=auditCorrectionSourceKey(source),owners=mirrorOwners.get(sourceKey)||new Set();owners.add(row);mirrorOwners.set(sourceKey,owners);}
  }
  const targets=new Map(),plan=[];
  const pending=changes.map((change,index)=>({change,index,mirrored:false})),verifiedMirrors=new Map();
  for(const {index,change,mirrored} of pending){
    if(!change||typeof change!=='object')auditCorrectionFail('CHANGE','The correction is invalid.',index);
    const explicit=Object.hasOwn(change,'source'),candidates=explicit?bySource.get(auditCorrectionSourceKey(change.source)):byTag.get(auditNormId(change.tag));
    if(!candidates||!candidates.length)auditCorrectionFail('SOURCE','No original row matches this correction.',index);
    if(candidates.length!==1)auditCorrectionFail('AMBIGUOUS','Multiple original rows match; an exact sheet and row are required.',index);
    const row=candidates[0],source=row._source,column=auditCorrectionColumn(change,index),prop=column.field;
    if(!source||typeof source.sheet!=='string'||!Number.isInteger(source.row)||source.row<1||source.row>1048576)auditCorrectionFail('SOURCE','The original row has no valid physical location.',index);
    if(change.tag!=null&&auditNormId(change.tag)!==auditNormId(row.equipmentId))auditCorrectionFail('SOURCE','The tag does not match the original physical row.',index);
    const c=source.columns&&source.columns[prop],tagColumn=source.columns&&source.columns.equipmentId;
    if(!Number.isInteger(c)||c<0||c>16383||!Number.isInteger(tagColumn)||tagColumn<0||tagColumn>16383)auditCorrectionFail('COLUMN','The original column mapping is missing or invalid.',index);
    if(explicit){
      const columns=change.source&&change.source.columns;
      if(!columns||columns[prop]!==c||Object.entries(columns).some(([field,at])=>source.columns[field]!==at))auditCorrectionFail('COLUMN','The correction column mapping differs from the original import.',index);
      if(!Object.hasOwn(change,'before'))auditCorrectionFail('BEFORE','The original value is required for a source-targeted correction.',index);
    }
    const before=Object.hasOwn(change,'before')?change.before:row[prop];
    if(before===undefined||clean(before)!==clean(row[prop]))auditCorrectionFail('BASELINE','The original value does not match the immutable baseline.',index);
    if(!mirrored){
      const canonicalKey=auditCorrectionSourceKey(source),canonicalRows=canonicalBySource.get(canonicalKey)||[];
      if(canonicalRows.length>1)auditCorrectionFail('AMBIGUOUS','Multiple canonical rows share the correction location.',index);
      const canonical=canonicalRows[0];
      if(canonical&&Object.hasOwn(canonical,'_sources')){
        let copies=verifiedMirrors.get(canonicalKey);
        if(!copies){copies=auditCorrectionMirrorRows(workbook,canonical,bySource,mirrorOwners,index);verifiedMirrors.set(canonicalKey,copies);}
        for(const copy of copies)if(auditCorrectionSourceKey(copy._source)!==canonicalKey)pending.push({index,mirrored:true,change:{...change,before,source:copy._source}});
      }
    }
    const sheet=workbook.Sheets[source.sheet];
    if(!sheet)auditCorrectionFail('SHEET','The original worksheet is missing.',index);
    const address=XLSX.utils.encode_cell({r:source.row-1,c}),tagAddress=XLSX.utils.encode_cell({r:source.row-1,c:tagColumn}),cell=sheet[address];
    if(auditCorrectionCellValue(sheet[tagAddress])!==clean(row.equipmentId))auditCorrectionFail('CONFLICT','The original equipment cell has changed.',index);
    if(cell&&(cell.f!=null||cell.F!=null||cell.t==='e'))auditCorrectionFail('FORMULA','Formula, array, and error cells cannot be replaced by a metadata correction.',index);
    if(auditCorrectionCellValue(cell)!==clean(before))auditCorrectionFail('CONFLICT','The original cell value has changed; no corrections were written.',index);
    if((sheet['!merges']||[]).some(range=>source.row-1>=range.s.r&&source.row-1<=range.e.r&&c>=range.s.c&&c<=range.e.c))auditCorrectionFail('MERGED','Merged cells cannot be corrected safely.',index);
    const value=auditCorrectionValue(change.value,index),key=JSON.stringify([source.sheet,address]),previous=targets.get(key);
    if(previous&&!Object.is(previous.value,value))auditCorrectionFail('CONFLICT','Two corrections request different values for the same cell.',index);
    const target={sheet,sheetName:source.sheet,address,row:source.row,column:c,value,index};
    if(!previous){targets.set(key,target);plan.push(target);}
  }
  for(const target of plan){
    for(const key of [target.address,'!ref']){
      const descriptor=Object.getOwnPropertyDescriptor(target.sheet,key);
      if(descriptor?descriptor.writable!==true:!Object.isExtensible(target.sheet))auditCorrectionFail('WORKBOOK','A target worksheet is read-only; no corrections were written.',target.index);
    }
    const ref=target.sheet['!ref'];
    if(ref&&!/^[A-Z]+[1-9][0-9]*(?::[A-Z]+[1-9][0-9]*)?$/.test(ref))auditCorrectionFail('WORKBOOK','A target worksheet range is invalid.',target.index);
  }
  return plan;
}
/* Read-only preview for an original sparse SheetJS workbook read with
   cellStyles:true to retain imported number formats. Counts include
   proven mirror cells and deduplicate identical writes. This checks cells, not
   package eligibility: callers must separately require an original safe XLSX. */
export function validateAuditCorrections(workbook,baselineSnapshot,changes){
  const plan=auditCorrectionPreflight(workbook,baselineSnapshot,changes);
  return {correctionCount:changes.length,cellCount:plan.length,sheetCount:new Set(plan.map(target=>target.sheetName)).size};
}
/* Legacy callers may omit source/before only when the baseline tag is unique.
   The return count is logical corrections, including any proven mirror writes.
   All validation runs before the first write. Styles and annotations survive. */
export function applyChangesToWorkbook(workbook,baselineSnapshot,changes){
  const plan=auditCorrectionPreflight(workbook,baselineSnapshot,changes);
  for(const target of plan){
    const cell={...target.sheet[target.address],t:typeof target.value==='number'?'n':typeof target.value==='boolean'?'b':'s',v:target.value};
    for(const key of ['w','f','F','D','h','r'])delete cell[key];
    sheetSetCell(target.sheet,target.address,cell);
  }
  return changes.length;
}

const AUDIT_PACKAGE_NS='http://schemas.openxmlformats.org/package/2006/relationships';
const AUDIT_SHEET_NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
function auditPackageChildren(node,name,namespace=node.namespaceURI){return Array.from(node.childNodes).filter(child=>child.nodeType===1&&child.localName===name&&child.namespaceURI===namespace);}
function auditPackageXml(bytes){
  const xml=new TextDecoder('utf-8',{fatal:true}).decode(bytes),document=new DOMParser().parseFromString(xml,'application/xml');
  if(!document.documentElement||document.doctype||document.getElementsByTagNameNS('*','parsererror').length)auditCorrectionFail('XML','The original workbook contains unsupported or invalid XML.');
  return document;
}
function auditPackagePath(base,target){
  if(!target||target.includes('\\'))auditCorrectionFail('PACKAGE','A workbook relationship has an invalid target.');
  const url=new URL(target,`https://xlsx.invalid/${base}`);
  if(url.origin!=='https://xlsx.invalid'||url.search||url.hash)auditCorrectionFail('PACKAGE','External workbook relationships cannot be edited.');
  return decodeURIComponent(url.pathname.slice(1));
}
function auditPackageRelationships(document){
  if(document.documentElement.localName!=='Relationships'||document.documentElement.namespaceURI!==AUDIT_PACKAGE_NS)auditCorrectionFail('PACKAGE','Workbook relationships are invalid.');
  const map=new Map();
  for(const relation of auditPackageChildren(document.documentElement,'Relationship')){
    const id=relation.getAttribute('Id');if(!id||map.has(id))auditCorrectionFail('PACKAGE','Workbook relationships are ambiguous.');map.set(id,relation);
  }
  return map;
}
function auditPackageElement(parent,name){return parent.ownerDocument.createElementNS(parent.namespaceURI,parent.prefix?`${parent.prefix}:${name}`:name);}
/* Only XLSX packages are supported: never silently downgrade a legacy, macro,
   signed, or unparseable workbook to a values-only copy. No source bytes mutate. */
export async function buildUpdatedRegistryBytes(sourceBytes,baselineSnapshot,changes,options={}){
  if(typeof DOMParser!=='function'||typeof XMLSerializer!=='function')auditCorrectionFail('XML','This browser does not provide the XML tools needed for safe export.');
  const bytes=sourceBytes instanceof ArrayBuffer?new Uint8Array(sourceBytes.slice(0)):ArrayBuffer.isView(sourceBytes)?new Uint8Array(sourceBytes.buffer,sourceBytes.byteOffset,sourceBytes.byteLength).slice():null;
  if(!bytes||bytes[0]!==0x50||bytes[1]!==0x4b)auditCorrectionFail('PACKAGE','Safe updated-registry export requires an original XLSX workbook.');
  const container=XLSX.CFB.read(bytes,{type:'array'}),entries=[],parts=new Map();
  container.FullPaths.forEach((path,index)=>{
    const name=path.slice(container.FullPaths[0].length),file=container.FileIndex[index];
    if(!name||name.endsWith('/')||name.startsWith('\u0001')||!file||file.type!==2)return;
    if(parts.has(name))auditCorrectionFail('PACKAGE','The original package contains duplicate parts.');
    const entry={name,data:new Uint8Array(file.content)};entries.push(entry);parts.set(name,entry);
  });
  if(entries.some(entry=>entry.name.toLowerCase().startsWith('_xmlsignatures/')))auditCorrectionFail('PACKAGE','Digitally signed workbooks cannot be corrected without invalidating their signatures.');
  const xmlAt=path=>{const entry=parts.get(path);if(!entry)auditCorrectionFail('PACKAGE','An original workbook part is missing.');return auditPackageXml(entry.data);};
  const roots=[...auditPackageRelationships(xmlAt('_rels/.rels')).values()].filter(relation=>relation.getAttribute('Type').endsWith('/officeDocument'));
  if(roots.length!==1||roots[0].getAttribute('TargetMode')==='External')auditCorrectionFail('PACKAGE','The original workbook relationship is missing or ambiguous.');
  const workbookPath=auditPackagePath('',roots[0].getAttribute('Target'));
  const types=xmlAt('[Content_Types].xml'),type=Array.from(types.getElementsByTagNameNS('*','Override')).filter(node=>node.getAttribute('PartName')===`/${workbookPath}`);
  if(type.length!==1||type[0].getAttribute('ContentType')!=='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')auditCorrectionFail('PACKAGE','Safe updated-registry export supports XLSX workbooks only.');
  const workbookXml=xmlAt(workbookPath),root=workbookXml.documentElement;
  if(root.localName!=='workbook'||root.namespaceURI!==AUDIT_SHEET_NS)auditCorrectionFail('PACKAGE','This workbook XML format is not supported for safe export.');
  const slash=workbookPath.lastIndexOf('/'),relationships=auditPackageRelationships(xmlAt(`${workbookPath.slice(0,slash+1)}_rels/${workbookPath.slice(slash+1)}.rels`));
  const sheetPaths=new Map(),sheetNames=new Set(),worksheetParts=new Set();
  for(const node of workbookXml.getElementsByTagNameNS(AUDIT_SHEET_NS,'sheet')){
    const name=node.getAttribute('name'),relation=relationships.get(node.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));
    if(sheetNames.has(name))auditCorrectionFail('PACKAGE','Worksheet names are ambiguous.');sheetNames.add(name);
    if(relation&&relation.getAttribute('TargetMode')!=='External'&&relation.getAttribute('Type').endsWith('/worksheet')){
      const path=auditPackagePath(workbookPath,relation.getAttribute('Target'));
      if(worksheetParts.has(path))auditCorrectionFail('PACKAGE','Multiple worksheets point to the same physical package part.');worksheetParts.add(path);sheetPaths.set(name,path);
    }
  }
  const workbook=XLSX.read(bytes,{type:'array',cellStyles:true}),plan=auditCorrectionPreflight(workbook,baselineSnapshot,changes),documents=new Map(),rowIndexes=new Map(),patches=[];
  for(const target of plan){
    const path=sheetPaths.get(target.sheetName);if(!path)auditCorrectionFail('PACKAGE','A targeted worksheet relationship is missing.',target.index);
    let rowsByNumber=rowIndexes.get(path);
    if(!rowsByNumber){
      const document=xmlAt(path),sheet=document.documentElement,data=auditPackageChildren(sheet,'sheetData');
      if(sheet.localName!=='worksheet'||sheet.namespaceURI!==AUDIT_SHEET_NS||data.length!==1)auditCorrectionFail('XML','A targeted worksheet has invalid cell data.',target.index);
      documents.set(path,document);rowsByNumber=new Map();rowIndexes.set(path,rowsByNumber);
      for(const row of auditPackageChildren(data[0],'row')){const key=row.getAttribute('r'),matches=rowsByNumber.get(key)||[];matches.push(row);rowsByNumber.set(key,matches);}
    }
    const rows=rowsByNumber.get(String(target.row));
    if(!rows||rows.length!==1)auditCorrectionFail('XML','A targeted physical row is missing or ambiguous.',target.index);
    const cells=auditPackageChildren(rows[0],'c'),matches=cells.filter(cell=>cell.getAttribute('r')===target.address);
    if(matches.length>1)auditCorrectionFail('XML','A targeted physical cell is ambiguous.',target.index);
    const cell=matches[0];if(cell&&auditPackageChildren(cell,'f').length)auditCorrectionFail('FORMULA','A formula cell cannot be replaced by a metadata correction.',target.index);
    patches.push({target,row:rows[0],cell,cells});
  }
  /* All physical cells and conflicts have now been checked. Only temporary DOMs
     are edited; no partially corrected workbook can escape on failure. */
  for(const {target,row,cell:existing,cells} of patches.sort((left,right)=>left.target.column-right.target.column)){
    const cell=existing||auditPackageElement(row,'c');
    if(!existing){cell.setAttribute('r',target.address);row.insertBefore(cell,cells.find(candidate=>XLSX.utils.decode_cell(candidate.getAttribute('r')).c>target.column)||null);}
    for(const name of ['v','is'])for(const child of auditPackageChildren(cell,name))cell.removeChild(child);
    const isString=typeof target.value==='string';cell.setAttribute('t',isString?'inlineStr':typeof target.value==='boolean'?'b':'n');
    const value=auditPackageElement(cell,isString?'is':'v');
    if(isString){const text=auditPackageElement(cell,'t');text.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');text.textContent=target.value.replace(/_x[0-9a-f]{4}_/gi,match=>`_x005F_${match.slice(1)}`).replace(/\r/g,'_x000D_');value.appendChild(text);}
    else value.textContent=typeof target.value==='boolean'?(target.value?'1':'0'):String(target.value);
    cell.insertBefore(value,cell.firstChild);
  }
  const encoder=new TextEncoder();for(const [path,document] of documents)parts.get(path).data=encoder.encode(new XMLSerializer().serializeToString(document));
  if(zipDeflateAvailable())return zipEntries(entries,options.onProgress);
  for(const entry of entries){const file=XLSX.CFB.find(container,`/${entry.name}`);file.content=entry.data;file.size=entry.data.length;}
  return new Uint8Array(XLSX.CFB.write(container,{fileType:'zip',type:'array',compression:true}));
}
export async function exportUpdatedRegistryXlsx(){
  const session=S.session,bytes=session&&session.sourceBytes,baseline=session&&session.baselineSnapshot;
  const changes=(session&&session.changes||[]).map(change=>({...change,...(change.source?{source:{...change.source,columns:{...change.source.columns}}}:{})}));
  if(!bytes){toast('The original workbook is not in memory — load the registry again first');return false;}
  if(!baseline){toast('The original audit baseline is missing; load the registry again first');return false;}
  if(!changes.length){toast('Nothing staged — action findings with a fix first');return false;}
  const base=clean(session.name).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  try{
    await runWithProgress('Building the updated registry','Checking original cells and staged corrections',async(checkpoint,report)=>{
      report(.05,'Checking every correction against the original workbook');await checkpoint();
      const updated=await buildUpdatedRegistryBytes(bytes,baseline,changes,{onProgress:async fraction=>{report(.4+fraction*.58,'Packaging the corrected copy');await checkpoint();}});
      report(1,'Corrected copy ready');downloadBlob(`${base}-Updated-Registry.xlsx`,new Blob([updated],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    });
    toast(`Updated registry exported: ${changes.length.toLocaleString()} corrections written`);
    return true;
  }catch(error){toast(error instanceof AuditCorrectionExportError?error.message:'The updated registry could not be built; no corrected copy was exported');return false;}
}

/* Current corrections and the review journal are separate: a reviewed finding
   is not necessarily corrected, and a staged change is not a source-file edit. */
export function buildAuditCorrectionsWorkbook(changes=[],reviewHistory=[],options={}){
  const workbook=XLSX.utils.book_new(),findings=new Map((options.baselineResult&&options.baselineResult.findings||[]).map(finding=>[finding.id,finding]));
  const sourceRows=new Map((options.baselineResult&&options.baselineResult.rows||[]).map(row=>[auditCorrectionSourceKey(row._source),row]));
  const dispositionLabels={'corrected-draft':'Corrected in draft',reviewed:'Reviewed',exception:'Exception'};
  const changeKey=change=>JSON.stringify([auditCorrectionSourceKey(change.source),change.prop||change.header||change.field,change.before,change.value]);
  const reviewByFinding=new Map(),reviewByChange=new Map();
  for(const review of reviewHistory){for(const id of review.findingIds||[])reviewByFinding.set(id,review);for(const change of review.changes||[])reviewByChange.set(changeKey(change),review);}
  const columns=['Equipment ID','Field','Before','After','Reason','Sheet','Row','Column','Cell','Status','Owner','Recorded At','Rule','Finding ID'];
  const correctionRow=(change,review,status)=>{
    const finding=findings.get(change.findingId),source=change.source||{},column=source.columns&&source.columns[change.prop||EXTO_REV21_COLUMNS.find(column=>extoRev21Norm(column.header)===extoRev21Norm(change.header||change.field))?.field];
    const letter=Number.isInteger(column)&&column>=0&&column<16384?auditColumnName(column):'';
    const value=value=>value==null?'':typeof value==='object'?JSON.stringify(value):value;
    return [clean(change.tag),clean(change.header||change.field||change.prop),value(change.before),value(change.value),clean(review&&review.reason||change.reason||finding&&finding.why),source.sheet||'',source.row||'',letter,letter&&source.row?`${letter}${source.row}`:'',status,clean(review&&review.owner),review&&review.at!=null?String(review.at):'',clean(change.ruleId||finding&&finding.rule&&finding.rule.id),clean(change.findingId)];
  };
  const log=[columns,...changes.map(change=>correctionRow(change,reviewByChange.get(changeKey(change))||reviewByFinding.get(change.findingId),'Staged in draft'))];
  const actions=[['Review ID','Recorded At','Owner','Disposition','Reason','Finding ID','Equipment ID','Field','Before','After','Sheet','Row','Column','Cell','Rule']];
  for(const review of reviewHistory){
    const disposition=dispositionLabels[review.disposition]||clean(review.disposition),ids=review.findingIds||[];
    const reviewedChanges=[...(review.changes||[])],represented=new Set(reviewedChanges.map(change=>change.findingId));
    for(const id of ids)if(!represented.has(id)){
      const finding=findings.get(id),row=sourceRows.get(auditCorrectionSourceKey(finding));
      reviewedChanges.push({findingId:id,tag:finding?.equipmentId,field:finding?.field,before:finding?.actual,source:row?row._source:{sheet:finding?.sheet,row:finding?.row},ruleId:finding?.rule?.id});
    }
    if(!reviewedChanges.length)reviewedChanges.push(null);
    for(const change of reviewedChanges){
      const cells=change?correctionRow(change,review,disposition):null;
      actions.push([clean(review.id),review.at==null?'':String(review.at),clean(review.owner),disposition,clean(review.reason),cells?cells[13]:'',cells?cells[0]:'',cells?cells[1]:'',cells?cells[2]:'',cells?cells[3]:'',cells?cells[5]:'',cells?cells[6]:'',cells?cells[7]:'',cells?cells[8]:'',cells?cells[12]:'']);
    }
  }
  const decorate=(rows,name,widths)=>{
    const sheet=XLSX.utils.aoa_to_sheet(rows);sheet['!cols']=widths.map(wch=>({wch}));styleHeaderRow(sheet);sheetFreezeRows(sheet,1);sheetAutoFilter(sheet,`A1:${auditColumnName(rows[0].length-1)}${rows.length}`);
    for(let r=1;r<rows.length;r++)for(let c=0;c<rows[0].length;c++)sheetStyleCell(sheet,XLSX.utils.encode_cell({r,c}),AUDIT_EXPORT_STYLES.wrap);
    addSheet(workbook,sheet,name);
  };
  decorate(log,'Correction Log',[30,28,42,42,58,26,8,9,12,22,22,26,26,30]);
  decorate(actions,'Actions',[28,26,22,24,58,38,30,28,42,42,26,8,9,12,26]);
  return workbook;
}
export async function exportAuditCorrectionsXlsx(){
  const session=S.session;if(!session||!session.baselineResult){toast('Run an SSM Audit first');return false;}
  const base=clean(session.name).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  try{
    const workbook=buildAuditCorrectionsWorkbook(session.changes||[],session.reviewHistory||[],{baselineResult:session.baselineResult});
    downloadBlob(`${base}-Corrections.xlsx`,await workbookBlobCompact(workbook));toast('Correction log and review actions exported');return true;
  }catch(_){toast('The correction log could not be built');return false;}
}

/* ---- Tracker Export ----
   A single shareable tab: one big bar at the top driven by the manual
   "Milestone actioned?" ticks, one row per L2 milestone (progress prefilled
   from in-app actioning), and one column per discipline showing that
   discipline's actioned share within the milestone. */
export function buildAuditTrackerWorkbook(currentResult,sessionName,options={}){
  const result=options.baselineResult||currentResult,actioned=new Set([...(options.actionedIds||[]),...(options.draftResolvedIds||[])]);
  const disabled=new Set(options.disabledRules||[]),excluded=new Set(options.excludedIds||[]),completed=new Set([...(options.completedEquipmentIds||[])].map(auditNormId));
  const inScope=finding=>!disabled.has(finding.rule?.id)&&!completed.has(auditNormId(finding.equipmentId))&&(!excluded.has(finding.id)||actioned.has(finding.id));
  const bySource=new Map(),byId=new Map();
  for(const row of [...(currentResult&&currentResult.rows||[]),...(result&&result.rows||[])]){
    if(row._source)bySource.set(auditCorrectionSourceKey(row._source),row);
  }
  for(const row of bySource.values()){const key=auditNormId(row.equipmentId),rows=byId.get(key)||[];rows.push(row);byId.set(key,rows);}
  const findingKey=finding=>JSON.stringify([finding.sheet||'',finding.row||0,finding.rule?.id||finding.ruleId||'',finding.row?'':auditNormId(finding.equipmentId)]);
  const baselineFindings=result&&result.findings||[],findings=baselineFindings.filter(inScope),baselineKeys=new Set(baselineFindings.map(findingKey));
  const baselineCounts=new Map();for(const finding of baselineFindings){const key=findingKey(finding);baselineCounts.set(key,(baselineCounts.get(key)||0)+1);}
  const doneKeys=new Set((currentResult&&currentResult.findings||[]).filter(finding=>actioned.has(finding.id)).map(findingKey));
  for(const finding of currentResult&&currentResult.findings||[])if(inScope(finding)&&!baselineKeys.has(findingKey(finding))){findings.push(finding);baselineKeys.add(findingKey(finding));}
  const groups=new Map(),discSet=new Set();
  for(const finding of findings){
    const candidates=byId.get(auditNormId(finding.equipmentId))||[],row=bySource.get(auditCorrectionSourceKey(finding))||(candidates.length===1?candidates[0]:null);
    const milestone=clean(row&&row.milestone)||'No L2 milestone';
    const discipline=clean(row&&row.discipline)||'No discipline';
    discSet.add(discipline);
    const group=groups.get(milestone)||{total:0,done:0,disc:new Map()};
    const key=findingKey(finding),done=actioned.has(finding.id)||baselineCounts.get(key)===1&&doneKeys.has(key);group.total++;if(done)group.done++;
    const cell=group.disc.get(discipline)||{total:0,done:0};cell.total++;if(done)cell.done++;group.disc.set(discipline,cell);
    groups.set(milestone,group);
  }
  const disciplines=[...discSet].sort(natCmp),milestones=[...groups.keys()].sort(natCmp);
  const headerSheetRow=8,firstRow=9,lastRow=8+milestones.length;
  const fixedHeaders=['L2 Milestone','Findings','Actioned in app','%','Progress','Milestone actioned?'];
  const aoa=[
    ['SSM Audit — Tracker','','','','',''],
    [`${clean(sessionName)} — generated ${(options.generatedAt||new Date()).toLocaleDateString()}`,'','','','',''],
    [`Action findings in the SSM Audit app, then tick "Milestone actioned?" (${AUDIT_EXPORT_TICK}) here when a whole milestone is closed out. The big bar follows the ticks.`,'','','','',''],
    ['','','','','',''],
    ['MILESTONES ACTIONED','','','','',''],
    [auditExportEmptyBar(),'','','','',''],
    ['','','','','',''],
    [...fixedHeaders,...disciplines],
  ];
  for(const milestone of milestones){
    const group=groups.get(milestone);
    const discCells=disciplines.map(discipline=>{const cell=group.disc.get(discipline);return cell?cell.done/cell.total:'';});
    aoa.push([milestone,group.total,group.done,0,auditExportEmptyBar(),group.total&&group.done===group.total?AUDIT_EXPORT_TICK:AUDIT_EXPORT_UNTICKED,...discCells]);
  }
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=[{wch:42},{wch:10},{wch:13},{wch:8},{wch:18},{wch:17},...disciplines.map(()=>({wch:14}))];
  sheet['!rows']=[{hpt:28},{hpt:16},{hpt:16},{hpt:8},{hpt:14},{hpt:34},{hpt:8},{hpt:30}];
  sheetStyleCell(sheet,'A1',AUDIT_EXPORT_STYLES.title);
  sheetStyleCell(sheet,'A2',AUDIT_EXPORT_STYLES.subtitle);
  sheetStyleCell(sheet,'A3',AUDIT_EXPORT_STYLES.note);
  sheetStyleCell(sheet,'A5',AUDIT_EXPORT_STYLES.overallLabel);
  const tickRange=milestones.length?`F${firstRow}:F${lastRow}`:'';
  const overallFormula=milestones.length?`COUNTIF(${tickRange},"${AUDIT_EXPORT_TICK}")/${milestones.length}`:'';
  if(overallFormula)sheetSetCell(sheet,'E6',sheetFormulaCell(overallFormula,0,AUDIT_EXPORT_STYLES.overallPercent));else sheetStyleCell(sheet,'E6',AUDIT_EXPORT_STYLES.overallPercent);
  if(overallFormula)sheetSetCell(sheet,'A6',auditExportBarCell('E6',AUDIT_EXPORT_STYLES.overallBar));else sheetStyleCell(sheet,'A6',AUDIT_EXPORT_STYLES.overallBar);
  for(const column of ['B','C','D'])sheetStyleCell(sheet,`${column}6`,AUDIT_EXPORT_STYLES.overallBar);
  auditExportHeaderRow(sheet,headerSheetRow,[...fixedHeaders,...disciplines],['left','right','right','right','left','center',...disciplines.map(()=>'right')]);
  milestones.forEach((milestone,offset)=>{
    const rowIndex=firstRow+offset,band=offset%2===1;
    sheet['!rows'][rowIndex-1]={hpt:22};
    sheetStyleCell(sheet,`A${rowIndex}`,band?AUDIT_EXPORT_STYLES.labelBand:AUDIT_EXPORT_STYLES.label);
    sheetStyleCell(sheet,`B${rowIndex}`,band?AUDIT_EXPORT_STYLES.numberMidBand:AUDIT_EXPORT_STYLES.numberMid);
    sheetStyleCell(sheet,`C${rowIndex}`,band?AUDIT_EXPORT_STYLES.numberMidBand:AUDIT_EXPORT_STYLES.numberMid);
    sheetSetCell(sheet,`D${rowIndex}`,auditExportPercentCell(`B${rowIndex}`,`C${rowIndex}`,band?AUDIT_EXPORT_STYLES.percentBand:AUDIT_EXPORT_STYLES.percent));
    sheetSetCell(sheet,`E${rowIndex}`,auditExportBarCell(`D${rowIndex}`,band?AUDIT_EXPORT_STYLES.barBand:AUDIT_EXPORT_STYLES.bar));
    sheetStyleCell(sheet,`F${rowIndex}`,AUDIT_EXPORT_STYLES.actioned);
    disciplines.forEach((discipline,at)=>{
      const address=`${auditColumnName(6+at)}${rowIndex}`;
      sheetStyleCell(sheet,address,band?AUDIT_EXPORT_STYLES.percentBand:AUDIT_EXPORT_STYLES.percent);
    });
  });
  const extras={conditionalFormatting:[auditExportDataBar('A6:A6',1)]};
  if(milestones.length){
    extras.conditionalFormatting.push(auditExportDataBar(`E${firstRow}:E${lastRow}`,2));
    extras.dataValidations=[auditExportTickValidation(tickRange)];
  }
  sheetXmlExtras(sheet,extras);
  sheetFreezeRows(sheet,headerSheetRow);
  const workbook=XLSX.utils.book_new();
  addSheet(workbook,sheet,'Tracker');
  return workbook;
}
export async function exportTrackerXlsx(){
  const session=S.session,result=session&&(session.rawResult||session.result);if(!result){toast('Run an SSM Audit first');return;}
  const workbook=buildAuditTrackerWorkbook(result,session.name,{baselineResult:session.baselineResult,actionedIds:session.actioned,draftResolvedIds:session.draftResolved,disabledRules:S.rules.disabled,excludedIds:session.excluded,completedEquipmentIds:session.status&&session.status.completed});
  const base=clean(session.name).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  try{
    const blob=await workbookBlobCompact(workbook,{});
    downloadBlob(`${base}-Tracker.xlsx`,blob);
    toast('Tracker exported');
  }catch(_){toast('The tracker could not be built');}
}
