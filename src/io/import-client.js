export function importAuditWorkbook(file,{audit=true,referenceKind,report=()=>{}}={}){
  return new Promise((resolve,reject)=>{
    let worker,url,settled=false;
    const finish=(error,result)=>{
      if(settled)return;settled=true;
      worker?.terminate();if(url)URL.revokeObjectURL(url);
      if(error)reject(error);else resolve(result);
    };
    try{
      if(typeof Worker==='undefined')throw new Error('Background workbook processing is unavailable. Open this HTML in a current browser with local workers enabled.');
      const vendor=document.getElementById('sheetjs-runtime')?.textContent;
      if(!vendor)throw new Error('The workbook reader is missing. Download a fresh copy of the app.');
      url=URL.createObjectURL(new Blob([vendor,'\n',AUDIT_IMPORT_WORKER_SOURCE],{type:'text/javascript'}));
      worker=new Worker(url);
      worker.onmessage=({data})=>{
        if(settled)return;
        if(data.type==='progress')report(data.fraction,data.label);
        else if(data.type==='result')finish(null,data);
        else if(data.type==='error')finish(new Error(data.message));
      };
      worker.onerror=event=>{event.preventDefault?.();finish(new Error('Background workbook processing failed. Try reopening the app in a current browser.'));};
      worker.onmessageerror=()=>finish(new Error('Could not receive the workbook results. Try reopening the app.'));
      worker.postMessage({file,fileName:file.name,audit,...(referenceKind?{referenceKind}:{})});
    }catch(error){finish(error);}
  });
}

export function prepareAuditReview(session,changes,migration,migrationChanged,report=()=>{},options={}){
  return new Promise((resolve,reject)=>{
    let connection=session.reviewWorker,initial=false;
    try{
      if(!connection){
        if(typeof Worker==='undefined')throw new Error('Background review is unavailable. Open this HTML in a current browser with local workers enabled.');
        const vendor=document.getElementById('sheetjs-runtime')?.textContent;
        if(!vendor)throw new Error('The workbook reader is missing. Download a fresh copy of the app.');
        const url=URL.createObjectURL(new Blob([vendor,'\n',AUDIT_IMPORT_WORKER_SOURCE],{type:'text/javascript'}));
        let worker;try{worker=new Worker(url);}catch(error){URL.revokeObjectURL(url);throw error;}
        connection={worker,url,pending:null};session.reviewWorker=connection;initial=true;
        session.disposeReviewWorker=()=>{worker.terminate();URL.revokeObjectURL(url);const pending=connection.pending;connection.pending=null;session.reviewWorker=null;session.disposeReviewWorker=null;pending?.reject(new Error('The registry changed. No review was applied.'));};
        worker.onmessage=({data})=>{
          const pending=connection.pending;if(!pending)return;
          if(data.type==='progress'){pending.report(data.fraction,data.label);return;}
          connection.pending=null;
          if(data.type==='result')pending.resolve(data.prepared);else pending.reject(new Error(data.message||'The changes could not be checked.'));
        };
        worker.onerror=event=>{event.preventDefault?.();session.disposeReviewWorker?.();};
        worker.onmessageerror=()=>session.disposeReviewWorker?.();
      }
      if(connection.pending)throw new Error('Wait for the current review to finish.');
      connection.pending={resolve,reject,report};
      connection.worker.postMessage({kind:options.actionsWorkbook?'actions-export':options.export?'export':'review',...(options.actionsWorkbook?{actionsWorkbook:options.actionsWorkbook}:{}),changes,previousChanges:session.changes||[],references:options.references||session.references||{},referencesChanged:options.references!==undefined,completedEquipmentIds:options.export?[...(session.status?.completed||[])]:undefined,migration,migrationChanged,...(initial?{baseline:session.baselineSnapshot,file:new Blob([session.sourceBytes||new Uint8Array()])}:{})});
    }catch(error){if(connection?.pending?.reject===reject){connection.pending=null;session.disposeReviewWorker?.();}reject(error);}
  });
}
