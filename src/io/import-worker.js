import { auditSnapshotFromWorkbook, auditNormId } from '../audit/model.js'
import { auditStatusFromWorkbook } from '../audit/status-report.js'
import { runSsmAudit } from '../audit/engine.js'
import { auditPrepareInWorker } from '../audit/review.js'

const auditReviewWorkerCache={};

// Keep the workbook and temporary sheet arrays off the UI thread. Only the
// compact audit model and the untouched original bytes return to the app.
self.onmessage=async({data})=>{
  const report=(fraction,label)=>self.postMessage({type:'progress',fraction,label});
  try{
    if(data.kind==='review'){
      const prepared=await auditPrepareInWorker(auditReviewWorkerCache,data,report);
      self.postMessage({type:'result',prepared});return;
    }
    report(.02,'Reading workbook');
    const bytes=new Uint8Array(await data.file.arrayBuffer());
    report(.06,'Opening workbook');
    const workbook=XLSX.read(bytes,{type:'array',dense:true});
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
