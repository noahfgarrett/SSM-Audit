import { auditSnapshotFromWorkbook, auditNormId } from '../audit/model.js'
import { auditStatusFromWorkbook } from '../audit/status-report.js'
import { runSsmAudit } from '../audit/engine.js'
import { auditPrepareInWorker } from '../audit/review.js'
import { auditReferenceSheets, auditReadReferenceWorkbook } from '../audit/references.js'
import { buildAuditUpdateRowsBytes, buildAuditActionsWorkbook } from '../audit/export.js'
import { workbookBytesCompact } from '../core/download.js'

const auditReviewWorkerCache={};

// Keep the workbook and temporary sheet arrays off the UI thread. Only the
// compact audit model and the untouched original bytes return to the app.
self.onmessage=async({data})=>{
  const report=(fraction,label)=>self.postMessage({type:'progress',fraction,label});
  try{
    if(data.kind==='actions-export'){
      if(data.baseline){auditReviewWorkerCache.baseline=data.baseline;auditReviewWorkerCache.file=data.file;auditReviewWorkerCache.workbook=null;}
      const {result,sessionName,options}=data.actionsWorkbook;
      report(.05,'Grouping action findings');
      const workbook=buildAuditActionsWorkbook(result,sessionName,{...options,onProgress:fraction=>report(.1+fraction*.55,'Building rule tabs')});
      report(.65,'Packaging Actions workbook');
      const bytes=new Uint8Array(await workbookBytesCompact(workbook,{onProgress:fraction=>report(.65+fraction*.34,'Packaging Actions workbook')}));
      report(1,'Actions workbook ready');self.postMessage({type:'result',prepared:{bytes}},[bytes.buffer]);return;
    }
    if(data.kind==='export'){
      if(data.baseline){auditReviewWorkerCache.baseline=data.baseline;auditReviewWorkerCache.file=data.file;auditReviewWorkerCache.workbook=null;}
      if(!auditReviewWorkerCache.file)throw new Error('Open the original registry before exporting corrections.');
      report(.02,'Reading original workbook');
      const source=new Uint8Array(await auditReviewWorkerCache.file.arrayBuffer());
      const bytes=await buildAuditUpdateRowsBytes(source,auditReviewWorkerCache.baseline,data.changes,{uploadTemplate:true,sourceWorkbook:auditReviewWorkerCache.workbook,completedEquipmentIds:data.completedEquipmentIds,onStage:report,onProgress:fraction=>report(.65+fraction*.34,'Packaging corrected copy')});
      report(1,'Corrected copy ready');self.postMessage({type:'result',prepared:{bytes}},[bytes.buffer]);return;
    }
    if(data.kind==='review'){
      const prepared=await auditPrepareInWorker(auditReviewWorkerCache,data,report);
      self.postMessage({type:'result',prepared});return;
    }
    report(.02,'Reading workbook');
    const bytes=new Uint8Array(await data.file.arrayBuffer());
    report(.06,'Opening workbook');
    const workbook=XLSX.read(bytes,{type:'array',dense:true});
    if(data.referenceKind){
      report(.2,'Finding reference sheets');
      const sheets=auditReferenceSheets(workbook,data.referenceKind);
      if(!sheets.length)throw new Error('No usable reference sheet was found.');
      const references=sheets.map((name,index)=>{
        report(.25+index/sheets.length*.65,'Reading reference entries');
        return auditReadReferenceWorkbook(workbook,data.referenceKind,name);
      });
      report(1,'Reference ready');
      self.postMessage({type:'result',references});return;
    }
    report(.2,'Reading registry tabs');
    const snapshot=await auditSnapshotFromWorkbook(workbook,data.fileName,null,(fraction,label)=>report(.2+fraction*.45,label));
    report(.68,'Checking completed equipment');
    const status=data.audit?await auditStatusFromWorkbook(workbook):null;
    if(status){
      let matched=0;
      for(const row of snapshot.rows)if(status.completed.has(auditNormId(row.equipmentId)))matched++;
      status.matched=matched;
    }
    report(.78,data.audit?'Running audit checks':'Preparing reference');
    const rawResult=data.audit?runSsmAudit(snapshot):null;
    report(.94,'Preparing results');
    self.postMessage({type:'result',snapshot,status,rawResult,bytes:data.audit?bytes:null},data.audit?[bytes.buffer]:[]);
  }catch(error){self.postMessage({type:'error',message:error?.message||'Could not read this registry'});}
};
