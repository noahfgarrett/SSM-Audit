import { clean, natCmp } from '../core/text.js'
import { downloadBlob, sheetAutoFilter, sheetCellStyle, sheetFormulaCell, sheetFreezeRows, sheetLinkCell, sheetSetCell, sheetStyleCell, styleHeaderRow, workbookBlob } from '../core/download.js'
import { S } from '../state.js'
import { toast } from '../ui/feedback.js'
import { SSM_AUDIT_RULES } from './engine.js'
import { auditColumnName, auditNormId } from './model.js'

function addSheet(workbook,sheet,name){XLSX.utils.book_append_sheet(workbook,sheet,name);}
function printable(value){return typeof value==='string'?value:JSON.stringify(value);}
const EXPORT_SOURCE_LABELS={registry:'Registry Integrity',sop:'SSM SOP',logic:'Commissioning Logic'},EXPORT_CONFIDENCE_LABELS={required:'Required',strong:'Strong pattern','description-rated':'Description based'};

/* ---- SSM Audit workbook ----
   One tab per L2 milestone so a commissioning engineer can work a milestone end
   to end: the full equipment tree of that milestone, one line per finding, and a
   plain `Actioned` column they type Y into. Index and Dashboard read those Y
   marks back with live COUNTIF formulas, so progress updates itself. */
const AUDIT_EXPORT_DASHBOARD_SHEET='Dashboard',AUDIT_EXPORT_INDEX_SHEET='Index',AUDIT_EXPORT_FINDINGS_SHEET='All Findings',AUDIT_EXPORT_RULES_SHEET='Rules';
const AUDIT_EXPORT_NO_MILESTONE='No milestone',AUDIT_EXPORT_BAR_SEGMENTS=10,AUDIT_EXPORT_BACK_LINK='← Index';
const AUDIT_EXPORT_ACTIONED_NOTE='Type Y in the Actioned column on an equipment’s first line when it is closed out. Index and Dashboard progress update from those marks.';
const AUDIT_EXPORT_PALETTE=Object.freeze({ink:'173F5F',accent:'F26722',headerText:'FFFFFF',body:'21323F',muted:'8A96A3',repeat:'AAB4BE',band:'F4F7FA',line:'D9E1E9',link:'1B5FAA'});
const AUDIT_EXPORT_NEST_FILLS=Object.freeze(['FFFFFF','EDF3F8','DCE8F1','CBDCEA','B9D0E3','A8C5DC','97BAD5']);
const AUDIT_EXPORT_SEVERITY_LABELS=Object.freeze({blocker:"WON'T UPLOAD",error:'BREAKS A RULE',warning:'CHECK THIS',info:'NOTE'});
const AUDIT_EXPORT_SEVERITY_COLORS=Object.freeze({blocker:{fill:'8C1D18',color:'FFFFFF'},error:{fill:'D9531E',color:'FFFFFF'},warning:{fill:'F2B441',color:'40320A'},info:{fill:'6E8598',color:'FFFFFF'}});
const AUDIT_EXPORT_MILESTONE_HEADERS=Object.freeze(['Actioned','Nest','Equipment ID','Description','Closest Parent','Discipline','UPN','System Name','Severity','Finding','Why','What to do','Actioned By','Note']);
const AUDIT_EXPORT_MILESTONE_WIDTHS=Object.freeze([10,6,36,40,28,24,8,30,12,32,54,54,16,30]);
const AUDIT_EXPORT_FINDING_HEADERS=Object.freeze(['Severity','Milestone','Equipment ID','Description','Rule','Why','What to do','Field','Found','Expected','Sheet','Row']);
const AUDIT_EXPORT_FINDING_WIDTHS=Object.freeze([12,34,30,40,34,54,54,26,44,44,22,8]);
const AUDIT_EXPORT_RULE_HEADERS=Object.freeze(['Rule','What must be true','Source','Confidence','Findings count']);
const AUDIT_EXPORT_RULE_WIDTHS=Object.freeze([36,84,24,20,16]);

const AUDIT_EXPORT_STYLES=Object.freeze({
  title:sheetCellStyle({bold:true,size:18,color:AUDIT_EXPORT_PALETTE.ink}),
  subtitle:sheetCellStyle({size:11,color:AUDIT_EXPORT_PALETTE.muted}),
  note:sheetCellStyle({italic:true,size:10,color:AUDIT_EXPORT_PALETTE.muted}),
  back:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.link,underline:true}),
  kpiLabel:sheetCellStyle({bold:true,size:9,color:AUDIT_EXPORT_PALETTE.muted,align:'center',bottom:AUDIT_EXPORT_PALETTE.accent,bottomWeight:'medium'}),
  kpiValue:sheetCellStyle({bold:true,size:16,color:AUDIT_EXPORT_PALETTE.ink,align:'center'}),
  kpiPercent:sheetCellStyle({bold:true,size:16,color:AUDIT_EXPORT_PALETTE.accent,align:'center',numFmt:'0%'}),
  header:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.headerText,fill:AUDIT_EXPORT_PALETTE.ink,align:'left',vertical:'center',wrap:true}),
  headerRight:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.headerText,fill:AUDIT_EXPORT_PALETTE.ink,align:'right',vertical:'center',wrap:true}),
  headerCenter:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.headerText,fill:AUDIT_EXPORT_PALETTE.ink,align:'center',vertical:'center',wrap:true}),
  text:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,vertical:'top'}),
  wrap:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,vertical:'top',wrap:true}),
  wrapMuted:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.muted,vertical:'top',wrap:true}),
  muted:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.muted,vertical:'top'}),
  repeat:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.repeat,vertical:'top'}),
  linkText:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.link,underline:true,vertical:'top'}),
  linkBand:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.link,underline:true,fill:AUDIT_EXPORT_PALETTE.band,vertical:'top'}),
  number:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,align:'right'}),
  numberBand:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.body,fill:AUDIT_EXPORT_PALETTE.band,align:'right'}),
  percent:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.ink,align:'right',numFmt:'0%'}),
  percentBand:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.ink,fill:AUDIT_EXPORT_PALETTE.band,align:'right',numFmt:'0%'}),
  bar:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.accent,align:'left'}),
  barBand:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.accent,fill:AUDIT_EXPORT_PALETTE.band,align:'left'}),
  nest:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.muted,align:'center'}),
  nestRepeat:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.repeat,align:'center'}),
  actioned:sheetCellStyle({align:'center',bold:true,color:AUDIT_EXPORT_PALETTE.ink,border:AUDIT_EXPORT_PALETTE.line}),
  actionedRepeat:sheetCellStyle({align:'center',fill:AUDIT_EXPORT_PALETTE.band,border:AUDIT_EXPORT_PALETTE.line}),
  entry:sheetCellStyle({border:AUDIT_EXPORT_PALETTE.line,vertical:'top'}),
});
const AUDIT_EXPORT_NEST_STYLES=Object.freeze(AUDIT_EXPORT_NEST_FILLS.map(fill=>Object.freeze({
  own:sheetCellStyle({bold:true,color:AUDIT_EXPORT_PALETTE.body,fill,vertical:'top'}),
  repeat:sheetCellStyle({color:AUDIT_EXPORT_PALETTE.repeat,fill,vertical:'top'}),
})));
const AUDIT_EXPORT_SEVERITY_STYLES=Object.freeze(Object.fromEntries(Object.entries(AUDIT_EXPORT_SEVERITY_COLORS).map(([severity,colors])=>[severity,sheetCellStyle({bold:true,size:10,color:colors.color,fill:colors.fill,align:'center',vertical:'center'})])));

function auditExportEmptyBar(){return '░'.repeat(AUDIT_EXPORT_BAR_SEGMENTS);}
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
function auditExportActionedCell(group){return sheetFormulaCell(`COUNTIF(${auditExportSheetRef(group.sheetName)}!A:A,"Y")`,0,AUDIT_EXPORT_STYLES.number);}
function auditExportPercentCell(equipmentCell,actionedCell,style){return sheetFormulaCell(`IF(${equipmentCell}=0,0,${actionedCell}/${equipmentCell})`,0,style);}
function auditExportBarCell(percentCell,style){
  const filled=`MIN(${AUDIT_EXPORT_BAR_SEGMENTS},ROUND(${percentCell}*${AUDIT_EXPORT_BAR_SEGMENTS},0))`;
  return sheetFormulaCell(`REPT("█",${filled})&REPT("░",${AUDIT_EXPORT_BAR_SEGMENTS}-${filled})`,auditExportEmptyBar(),style);
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

function auditExportDashboardSheet(result,groups,sessionName,generated){
  const summary=result&&result.summary||{},severity=summary.severity||{},firstRow=9,lastRow=firstRow+Math.max(0,groups.length-1);
  const aoa=[
    [`SSM Audit — ${clean(sessionName)||'Registry'}`,'','','','','',''],
    [`Generated ${auditExportDate(generated)}`,'','','','','',''],
    [`Standard: ${clean(result&&result.standard)}`,'','','','','',''],
    ['','','','','','',''],
    ['Rows audited','Findings',"Won't upload",'Breaks a rule','Check this','Notes','Actioned'],
    [summary.rows||0,summary.findings||0,severity.blocker||0,severity.error||0,severity.warning||0,severity.info||0,0],
    ['','','','','','',''],
    ['Milestone','Equipment','Findings','Actioned','Progress','%'],
  ];
  for(const group of groups)aoa.push([group.label,group.equipmentCount,group.findingCount,0,auditExportEmptyBar(),0]);
  if(!groups.length)aoa.push(['No equipment rows in this registry','','','','','']);
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=[{wch:44},{wch:12},{wch:11},{wch:11},{wch:16},{wch:9},{wch:12}];
  sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:6}},{s:{r:1,c:0},e:{r:1,c:6}},{s:{r:2,c:0},e:{r:2,c:6}}];
  sheet['!rows']=[{hpt:30},{hpt:18},{hpt:18},{hpt:10},{hpt:18},{hpt:28}];
  sheetStyleCell(sheet,'A1',AUDIT_EXPORT_STYLES.title);
  sheetStyleCell(sheet,'A2',AUDIT_EXPORT_STYLES.subtitle);
  sheetStyleCell(sheet,'A3',AUDIT_EXPORT_STYLES.subtitle);
  for(let column=0;column<7;column++){
    sheetStyleCell(sheet,`${auditColumnName(column)}5`,AUDIT_EXPORT_STYLES.kpiLabel);
    sheetStyleCell(sheet,`${auditColumnName(column)}6`,column===6?AUDIT_EXPORT_STYLES.kpiPercent:AUDIT_EXPORT_STYLES.kpiValue);
  }
  if(groups.length)sheetSetCell(sheet,'G6',sheetFormulaCell(`IF(SUM(B${firstRow}:B${lastRow})=0,0,SUM(D${firstRow}:D${lastRow})/SUM(B${firstRow}:B${lastRow}))`,0,AUDIT_EXPORT_STYLES.kpiPercent));
  auditExportHeaderRow(sheet,8,['Milestone','Equipment','Findings','Actioned','Progress','%'],['left','right','right','right','left','right']);
  groups.forEach((group,offset)=>{
    const rowIndex=firstRow+offset,band=offset%2===1;
    sheetStyleCell(sheet,`A${rowIndex}`,band?AUDIT_EXPORT_STYLES.linkBand:AUDIT_EXPORT_STYLES.linkText);
    sheetLinkCell(sheet,`A${rowIndex}`,`#${auditExportSheetRef(group.sheetName)}!A1`,`Open ${group.label}`);
    sheetStyleCell(sheet,`B${rowIndex}`,band?AUDIT_EXPORT_STYLES.numberBand:AUDIT_EXPORT_STYLES.number);
    sheetStyleCell(sheet,`C${rowIndex}`,band?AUDIT_EXPORT_STYLES.numberBand:AUDIT_EXPORT_STYLES.number);
    const actioned=auditExportActionedCell(group);
    if(band){actioned.s=AUDIT_EXPORT_STYLES.numberBand;}
    sheetSetCell(sheet,`D${rowIndex}`,actioned);
    sheetSetCell(sheet,`E${rowIndex}`,auditExportBarCell(`F${rowIndex}`,band?AUDIT_EXPORT_STYLES.barBand:AUDIT_EXPORT_STYLES.bar));
    sheetSetCell(sheet,`F${rowIndex}`,auditExportPercentCell(`B${rowIndex}`,`D${rowIndex}`,band?AUDIT_EXPORT_STYLES.percentBand:AUDIT_EXPORT_STYLES.percent));
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

function auditExportMilestoneSheet(group){
  const aoa=[[AUDIT_EXPORT_BACK_LINK,AUDIT_EXPORT_ACTIONED_NOTE],[...AUDIT_EXPORT_MILESTONE_HEADERS]];
  const meta=[];
  for(const line of group.lines){
    const row=line.row,indent='  '.repeat(Math.min(line.level,24));
    const equipment=[indent+clean(row.equipmentId),clean(row.equipmentDescription),clean(row.closestParent),clean(row.discipline),clean(row.upn),clean(row.systemName)];
    const findings=line.findings.length?line.findings:[null];
    findings.forEach((finding,offset)=>{
      aoa.push(['',line.level,...equipment,
        finding?AUDIT_EXPORT_SEVERITY_LABELS[finding.severity]||finding.severity.toUpperCase():'',
        finding?clean(finding.rule.title):'',finding?clean(finding.why):'',finding?clean(finding.recommendation):'','','']);
      meta.push({level:line.level,repeat:offset>0,severity:finding?finding.severity:''});
    });
  }
  const sheet=XLSX.utils.aoa_to_sheet(aoa);
  sheet['!cols']=AUDIT_EXPORT_MILESTONE_WIDTHS.map(width=>({wch:width}));
  sheetStyleCell(sheet,'A1',AUDIT_EXPORT_STYLES.back);
  sheetLinkCell(sheet,'A1',`#${auditExportSheetRef(AUDIT_EXPORT_INDEX_SHEET)}!A1`,'Back to the milestone index');
  sheetStyleCell(sheet,'B1',AUDIT_EXPORT_STYLES.note);
  auditExportHeaderRow(sheet,2,AUDIT_EXPORT_MILESTONE_HEADERS,['center','center','left','left','left','left','right','left','center','left','left','left','left','left']);
  meta.forEach((entry,offset)=>{
    const rowIndex=offset+3,nest=AUDIT_EXPORT_NEST_STYLES[Math.min(entry.level,AUDIT_EXPORT_NEST_STYLES.length-1)];
    const body=entry.repeat?AUDIT_EXPORT_STYLES.repeat:AUDIT_EXPORT_STYLES.text;
    sheetStyleCell(sheet,`A${rowIndex}`,entry.repeat?AUDIT_EXPORT_STYLES.actionedRepeat:AUDIT_EXPORT_STYLES.actioned);
    sheetStyleCell(sheet,`B${rowIndex}`,entry.repeat?AUDIT_EXPORT_STYLES.nestRepeat:AUDIT_EXPORT_STYLES.nest);
    sheetStyleCell(sheet,`C${rowIndex}`,entry.repeat?nest.repeat:nest.own);
    for(const column of ['D','E','F','G','H'])sheetStyleCell(sheet,`${column}${rowIndex}`,body);
    if(entry.severity)sheetStyleCell(sheet,`I${rowIndex}`,AUDIT_EXPORT_SEVERITY_STYLES[entry.severity]||AUDIT_EXPORT_STYLES.text);
    const detail=entry.severity?AUDIT_EXPORT_STYLES.wrap:AUDIT_EXPORT_STYLES.wrapMuted;
    sheetStyleCell(sheet,`J${rowIndex}`,entry.severity?AUDIT_EXPORT_STYLES.text:AUDIT_EXPORT_STYLES.muted);
    sheetStyleCell(sheet,`K${rowIndex}`,detail);
    sheetStyleCell(sheet,`L${rowIndex}`,detail);
    sheetStyleCell(sheet,`M${rowIndex}`,AUDIT_EXPORT_STYLES.entry);
    sheetStyleCell(sheet,`N${rowIndex}`,AUDIT_EXPORT_STYLES.entry);
  });
  sheetFreezeRows(sheet,2);
  sheetAutoFilter(sheet,`A2:${auditColumnName(AUDIT_EXPORT_MILESTONE_HEADERS.length-1)}${Math.max(2,meta.length+2)}`);
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

export function buildAuditWorkbook(result,sessionName,options={}){
  const workbook=XLSX.utils.book_new(),groups=auditExportGroups(result);
  const used=new Set([AUDIT_EXPORT_DASHBOARD_SHEET,AUDIT_EXPORT_INDEX_SHEET,AUDIT_EXPORT_FINDINGS_SHEET,AUDIT_EXPORT_RULES_SHEET].map(name=>name.toLowerCase()));
  for(const group of groups)group.sheetName=auditExportSheetName(group.label,used);
  const generated=options.generatedAt instanceof Date?options.generatedAt:new Date();
  addSheet(workbook,auditExportDashboardSheet(result,groups,sessionName,generated),AUDIT_EXPORT_DASHBOARD_SHEET);
  addSheet(workbook,auditExportIndexSheet(groups),AUDIT_EXPORT_INDEX_SHEET);
  for(const group of groups)addSheet(workbook,auditExportMilestoneSheet(group),group.sheetName);
  addSheet(workbook,auditExportFindingsSheet(result,groups),AUDIT_EXPORT_FINDINGS_SHEET);
  addSheet(workbook,auditExportRulesSheet(result),AUDIT_EXPORT_RULES_SHEET);
  return workbook;
}

export function exportSsmAuditXlsx(){
  const result=S.session&&S.session.result;if(!result){toast('Run an SSM Audit first');return;}
  const workbook=buildAuditWorkbook(result,S.session.name);
  const base=clean(S.session.name).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  /* A milestone workbook carries every equipment row and every finding, so it is
     written compressed — uncompressed it runs several times larger. */
  downloadBlob(`${base}-Audit.xlsx`,workbookBlob(workbook,{compression:true}));toast('SSM Audit report exported');
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
