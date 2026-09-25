/* ---- Equipment Status Report tab ----
   The user supplies a completed-only report. Every listed equipment identity
   is finished on site, regardless of step names or cell statuses,
   so its findings are dropped from every metric -- the registry structure is
   still audited in full (a completed parent still anchors its children).
   The registry itself lives on the 'Full Export' tab, which the normal
   sheet scan picks up by its headers; this file only reads the status tab. */

import { auditNormId } from './model.js'
import { sheetAoaAsync } from '../io/workbook.js'

const AUDIT_STATUS_ID_HEADERS=['EQUIPMENT ID','EQUIPMENT TAG','EQUIPMENT TAG/ID','EQUIPMENT TAG / ID','EQUIPMENT NAME'];

function auditStatusNorm(value){return String(value==null?'':value).replace(/\s+/g,' ').trim().toUpperCase();}

export function auditStatusSheetName(names){
  return (names||[]).find(name=>/equipment\s*status\s*report/i.test(String(name)))||'';
}

/* aoa -> {completed:Set<normalized tag>, equipment:[{name}], completedRows,
   totalRows} or null when the sheet does not carry the expected headers.
   Prefer an explicit ID/tag column over a descriptive equipment name. */
export function auditStatusCompleted(aoa){
  if(!Array.isArray(aoa))return null;
  let headerRow=-1,nameCol=-1;
  for(let r=0;r<Math.min(aoa.length,25);r++){
    const cells=(aoa[r]||[]).map(auditStatusNorm);
    for(const header of AUDIT_STATUS_ID_HEADERS){nameCol=cells.indexOf(header);if(nameCol>=0)break;}
    if(nameCol>=0){headerRow=r;break;}
  }
  if(headerRow<0)return null;
  const completed=new Set(),byName=new Map();let totalRows=0;
  for(let r=headerRow+1;r<aoa.length;r++){
    const cells=aoa[r]||[];
    const name=String(cells[nameCol]==null?'':cells[nameCol]).trim();
    const key=auditNormId(name);if(!key)continue;totalRows++;
    completed.add(key);
    if(!byName.has(key))byName.set(key,{name});
  }
  return {completed,equipment:[...byName.values()],completedRows:completed.size,totalRows};
}

export async function auditStatusFromWorkbook(workbook,checkpoint){
  const name=auditStatusSheetName(workbook&&workbook.SheetNames);
  if(!name)return null;
  const parsed=await sheetAoaAsync(workbook.Sheets[name],async()=>{if(checkpoint)await checkpoint();});
  return auditStatusCompleted(parsed.aoa);
}
