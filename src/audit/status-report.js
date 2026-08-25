/* ---- Equipment Status Report tab ----
   Some exports carry a second tab recording where each equipment stands in
   commissioning. Equipment whose OA/BT step is Completed is finished on site,
   so its findings are dropped from every metric -- the registry structure is
   still audited in full (a completed parent still anchors its children).
   The registry itself lives on the 'Full Export' tab, which the normal
   sheet scan picks up by its headers; this file only reads the status tab. */

import { auditNormId } from './model.js'
import { sheetAoaAsync } from '../io/workbook.js'

const AUDIT_STATUS_NAME_HEADER='EQUIPMENT NAME';
const AUDIT_STATUS_COLUMNS=['RR OA/BT','DIST OA/BT','EQ OA/BT','SYS OA/BT'];

function auditStatusNorm(value){return String(value==null?'':value).replace(/\s+/g,' ').trim().toUpperCase();}

export function auditStatusSheetName(names){
  return (names||[]).find(name=>/equipment\s*status\s*report/i.test(String(name)))||'';
}

/* aoa -> {completed:Set<normalized tag>, equipment:[{name,steps}], completedRows,
   totalRows} or null when the sheet does not carry the expected headers.
   Completed means at least one OA/BT column says Completed and none says Not
   Started -- a row that says both is treated as not started, per the SOP
   owner's instruction. `steps` records which OA/BT columns said Completed. */
export function auditStatusCompleted(aoa){
  if(!Array.isArray(aoa))return null;
  let headerRow=-1,nameCol=-1,statusCols=[];
  for(let r=0;r<Math.min(aoa.length,25);r++){
    const cells=aoa[r]||[];const cols=[];let name=-1;
    for(let c=0;c<cells.length;c++){
      const value=auditStatusNorm(cells[c]);
      if(value===AUDIT_STATUS_NAME_HEADER)name=c;
      if(AUDIT_STATUS_COLUMNS.includes(value))cols.push({col:c,label:value});
    }
    if(name>=0&&cols.length){headerRow=r;nameCol=name;statusCols=cols;break;}
  }
  if(headerRow<0)return null;
  const completed=new Set(),byName=new Map();let totalRows=0;
  for(let r=headerRow+1;r<aoa.length;r++){
    const cells=aoa[r]||[];
    const name=String(cells[nameCol]==null?'':cells[nameCol]).trim();
    if(!name)continue;totalRows++;
    const values=statusCols.map(entry=>auditStatusNorm(cells[entry.col]));
    if(values.some(value=>value==='COMPLETED')&&!values.some(value=>value==='NOT STARTED')){
      const key=auditNormId(name);completed.add(key);
      const steps=statusCols.filter((entry,at)=>values[at]==='COMPLETED').map(entry=>entry.label);
      const existing=byName.get(key);
      if(existing)for(const step of steps){if(!existing.steps.includes(step))existing.steps.push(step);}
      else byName.set(key,{name,steps});
    }
  }
  return {completed,equipment:[...byName.values()],completedRows:completed.size,totalRows};
}

export async function auditStatusFromWorkbook(workbook,checkpoint){
  const name=auditStatusSheetName(workbook&&workbook.SheetNames);
  if(!name)return null;
  const parsed=await sheetAoaAsync(workbook.Sheets[name],async()=>{if(checkpoint)await checkpoint();});
  return auditStatusCompleted(parsed.aoa);
}
