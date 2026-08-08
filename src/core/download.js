export function downloadBlob(filename,blob){
  const url=URL.createObjectURL(blob),anchor=document.createElement('a');
  anchor.href=url;anchor.download=filename;anchor.rel='noopener';
  document.body.appendChild(anchor);anchor.click();anchor.remove();
  setTimeout(()=>URL.revokeObjectURL(url),30000);
}

export function workbookBlob(workbook){
  return new Blob([XLSX.write(workbook,{bookType:'xlsx',type:'array',cellStyles:true})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

export function styleHeaderRow(sheet,rowIndex=0){
  if(!sheet||!sheet['!ref'])return;
  const range=XLSX.utils.decode_range(sheet['!ref']);
  for(let column=range.s.c;column<=range.e.c;column++){
    const address=XLSX.utils.encode_cell({r:rowIndex,c:column}),cell=sheet[address];if(!cell)continue;
    cell.s={font:{bold:true,color:{rgb:'FFFFFF'}},fill:{patternType:'solid',fgColor:{rgb:'173F5F'}},alignment:{vertical:'center'}};
  }
  sheet['!rows']=sheet['!rows']||[];sheet['!rows'][rowIndex]={hpt:24};
}
