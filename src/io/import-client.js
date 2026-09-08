export function importAuditWorkbook(file,{audit=true,report=()=>{}}={}){
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
      worker.postMessage({file,fileName:file.name,audit});
    }catch(error){finish(error);}
  });
}
