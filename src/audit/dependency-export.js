import { clean } from '../core/text.js'
import { sheetAutoFilter, sheetCellStyle, sheetFreezeRows, sheetStyleCell, styleHeaderRow } from '../core/download.js'
import { AUDIT_DEPENDENCY_SOURCE_KINDS } from './expected-dependencies.js'

const DEPENDENCY_EXPORT_MISSING_FILL='FBE3E1',DEPENDENCY_EXPORT_MATCH_FILL='E2F0D9';
const DEPENDENCY_EXPORT_STATUS_LABELS=Object.freeze({match:'Match',missing:'Missing'});
const DEPENDENCY_EXPORT_MAX_LOCATIONS=5,DEPENDENCY_EXPORT_MAX_AFFECTED=200;

function dependencyExportLocation(source){return `${source.label} ${source.sheet}, row ${source.row}`;}
/* Lists stay short enough for one Excel cell (32,767 characters); the
   remainder is counted, never dropped silently. */
function dependencyExportList(items,limit,total=items.length){
  const shown=items.slice(0,limit);
  return shown.join('; ')+(total>shown.length?`; and ${(total-shown.length).toLocaleString()} more`:'');
}
function dependencyExportLocations(expected){
  const unique=[...new Set(expected.flatMap(entry=>entry.sources).map(dependencyExportLocation))];
  const total=Math.max(unique.length,expected.reduce((sum,entry)=>sum+(entry.sourceCount||entry.sources.length),0));
  return dependencyExportList(unique,DEPENDENCY_EXPORT_MAX_LOCATIONS,total);
}
function dependencyExportSheet(rows,widths){
  const sheet=XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols']=widths.map(wch=>({wch}));
  styleHeaderRow(sheet,0);sheetFreezeRows(sheet,1);sheetAutoFilter(sheet,sheet['!ref']);
  return sheet;
}
function dependencyExportSummary(check,registryName){
  const counts=check.counts,byDiscipline=new Map();
  for(const entry of check.rows){
    const name=clean(entry.discipline)||'(blank)';
    const tally=byDiscipline.get(name)||{compared:0,match:0,missing:0};
    tally.compared++;tally[entry.status]++;byDiscipline.set(name,tally);
  }
  const rows=[
    ['Dependency check','Value'],
    ['Registry',clean(registryName)],
    ['Documents',check.kinds.map(kind=>AUDIT_DEPENDENCY_SOURCE_KINDS[kind]).join(', ')||'None selected'],
    ['Compared',counts.compared],['Match',counts.match],['Missing',counts.missing],['Not in registry',counts.notInRegistry],
    [],
    ['Discipline','Compared','Match','Missing'],
    ...[...byDiscipline].sort(([left],[right])=>left.localeCompare(right)).map(([name,tally])=>[name,tally.compared,tally.match,tally.missing]),
  ];
  const sheet=XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols']=[{wch:28},{wch:46},{wch:12},{wch:12}];
  styleHeaderRow(sheet,0);styleHeaderRow(sheet,8);
  return sheet;
}
function dependencyExportTable(check){
  const header=['Equipment ID','Status','Expected dependencies','Registry dependencies','Missing','Note','Source','Why','Discipline','Building','UPN','System Name'];
  const rows=[header,...check.rows.map(entry=>[
    entry.equipmentId,DEPENDENCY_EXPORT_STATUS_LABELS[entry.status],
    entry.expected.map(expected=>expected.tag).join('; '),entry.registryDependencies.join('; '),entry.missing.join('; '),
    entry.notes.join('. '),dependencyExportLocations(entry.expected),
    entry.expected.map(expected=>`${expected.tag}: ${expected.reason}`).join('. '),
    entry.discipline,entry.building,entry.upn,entry.systemName,
  ])];
  const sheet=dependencyExportSheet(rows,[30,10,34,34,30,44,40,60,24,18,8,34]);
  const missing=sheetCellStyle({fill:DEPENDENCY_EXPORT_MISSING_FILL,vertical:'top',wrap:true}),match=sheetCellStyle({fill:DEPENDENCY_EXPORT_MATCH_FILL,vertical:'top'});
  check.rows.forEach((entry,index)=>{
    const row=index+2;
    sheetStyleCell(sheet,`B${row}`,entry.status==='missing'?missing:match);
    if(entry.status==='missing')sheetStyleCell(sheet,`E${row}`,missing);
  });
  return sheet;
}
function dependencyExportNotInRegistry(check){
  const header=['Source tag','Document','Sheet','Row','Also found at','Registry rows affected'];
  const rows=[header,...check.notInRegistry.map(entry=>{
    const [first,...rest]=entry.sources;
    const documents=(entry.documents||[...new Set(entry.sources.map(source=>source.label))]).join('; ');
    const others=rest.map(source=>`${source.sheet}, row ${source.row}`);
    return [entry.tag,documents,first?.sheet||'',first?.row||'',
      dependencyExportList(others,DEPENDENCY_EXPORT_MAX_LOCATIONS,Math.max(others.length,(entry.sourceCount||entry.sources.length)-1)),
      dependencyExportList(entry.affects,DEPENDENCY_EXPORT_MAX_AFFECTED)];
  })];
  return dependencyExportSheet(rows,[30,24,24,8,40,50]);
}

export function buildDependencyCheckWorkbook(check,registryName){
  const workbook=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook,dependencyExportSummary(check,registryName),'Summary');
  XLSX.utils.book_append_sheet(workbook,dependencyExportTable(check),'Dependencies');
  XLSX.utils.book_append_sheet(workbook,dependencyExportNotInRegistry(check),'Not in registry');
  return workbook;
}
