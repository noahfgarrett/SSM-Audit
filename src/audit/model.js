import { clean } from '../core/text.js'
import { sheetAoaAsync } from '../io/workbook.js'
import { EXTO_REV21_COLUMNS, extoRev21Norm } from '../exto/rev21-contract.js'

export const SSM_AUDIT_SCHEMA_VERSION=1;
export function auditNormId(value){return clean(value).replace(/\s+/g,' ').toUpperCase();}
export function auditSplitReferences(value){return clean(value).split(/\s*;\s*/).map(clean).filter(Boolean);}
export function auditColumnName(index){return XLSX.utils.encode_col(index);}

export function detectAuditRegistryHeader(aoa,maxRows=200){
  const wanted=new Map(EXTO_REV21_COLUMNS.map(column=>[extoRev21Norm(column.header),column]));
  let best=null;
  for(let row=0;row<Math.min(maxRows,aoa.length);row++){
    const fields={},headers=aoa[row]||[];
    for(let column=0;column<headers.length;column++){
      const contract=wanted.get(extoRev21Norm(headers[column]));
      if(contract)fields[contract.field]=column;
    }
    const score=Object.keys(fields).length;
    if(fields.equipmentId!=null&&fields.closestParent!=null&&fields.upn!=null&&(!best||score>best.score))best={headerRow:row,fields,score};
  }
  return best;
}

export function auditSnapshotFromAoa(aoa,source={}){
  const detected=detectAuditRegistryHeader(aoa);if(!detected)return null;
  const rowNums=Array.isArray(source.rowNums)?source.rowNums:null;
  const physicalRow=index=>Number.isInteger(rowNums&&rowNums[index])?rowNums[index]+1:index+1;
  const rows=[];
  for(let rowIndex=detected.headerRow+1;rowIndex<aoa.length;rowIndex++){
    const sourceRow=aoa[rowIndex]||[],record={};let populated=false;
    for(const column of EXTO_REV21_COLUMNS){
      const index=detected.fields[column.field],value=index==null?'':clean(sourceRow[index]);
      record[column.field]=value;if(value)populated=true;
    }
    if(!populated)continue;
    record._source=Object.freeze({file:clean(source.file),sheet:clean(source.sheet),row:physicalRow(rowIndex),columns:Object.freeze({...detected.fields})});
    rows.push(Object.freeze(record));
  }
  const missingHeaders=EXTO_REV21_COLUMNS.filter(column=>detected.fields[column.field]==null).map(column=>column.header);
  const headerRow=physicalRow(detected.headerRow);
  const missingHeaderEntries=missingHeaders.map(header=>Object.freeze({header,sheet:clean(source.sheet),headerRow}));
  return Object.freeze({schemaVersion:SSM_AUDIT_SCHEMA_VERSION,source:Object.freeze({file:clean(source.file),sheet:clean(source.sheet)}),
    headerRow,fields:Object.freeze({...detected.fields}),missingHeaders:Object.freeze(missingHeaders),missingHeaderEntries:Object.freeze(missingHeaderEntries),rows:Object.freeze(rows)});
}

export async function auditSnapshotFromWorkbook(workbook,fileName,checkpoint,report){
  const snapshots=[],names=workbook&&workbook.SheetNames||[];
  for(let index=0;index<names.length;index++){
    const name=names[index],parsed=await sheetAoaAsync(workbook.Sheets[name],async(done,total)=>{
      if(report)report((index+done/Math.max(1,total))/Math.max(1,names.length),`Scanning ${name}`);
      if(checkpoint)await checkpoint();
    });
    const snapshot=auditSnapshotFromAoa(parsed.aoa,{file:fileName,sheet:name,rowNums:parsed.rowNums});if(snapshot)snapshots.push(snapshot);
    if(report)report((index+1)/Math.max(1,names.length),`Scanned ${index+1} of ${names.length} tabs`);
    if(checkpoint)await checkpoint();
  }
  if(!snapshots.length)throw new Error('No Cx Registry tab with Equipment ID, Closest Parent, UPN, and Discipline headers was found');
  const rows=snapshots.flatMap(snapshot=>snapshot.rows),missingHeaders=[...new Set(snapshots.flatMap(snapshot=>snapshot.missingHeaders))];
  const missingHeaderEntries=snapshots.flatMap(snapshot=>snapshot.missingHeaderEntries);
  return Object.freeze({schemaVersion:SSM_AUDIT_SCHEMA_VERSION,source:Object.freeze({file:clean(fileName),sheets:Object.freeze(snapshots.map(snapshot=>snapshot.source.sheet))}),
    snapshots:Object.freeze(snapshots),rows:Object.freeze(rows),missingHeaders:Object.freeze(missingHeaders),missingHeaderEntries:Object.freeze(missingHeaderEntries)});
}

export function auditFingerprint(value){
  const text=String(value),bytes=new TextEncoder().encode(text);let hash=0x811c9dc5;
  for(const byte of bytes){hash^=byte;hash=Math.imul(hash,0x01000193);}
  return (hash>>>0).toString(16).padStart(8,'0');
}
