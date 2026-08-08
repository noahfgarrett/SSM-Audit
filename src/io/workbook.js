export function readArrayBuffer(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=()=>resolve(reader.result);
    reader.onerror=()=>reject(reader.error||new Error('Could not read the workbook'));
    reader.readAsArrayBuffer(file);
  });
}

const DENSE_ROW_CHUNK=2000;
function cellValue(cell){
  if(!cell||(cell.v==null&&cell.f==null&&cell.w==null))return '';
  return cell.w!=null?cell.w:XLSX.utils.format_cell(cell);
}
export async function sheetAoaAsync(sheet,onChunk){
  if(!sheet||!sheet['!ref'])return {aoa:[],rowNums:[]};
  if(Array.isArray(sheet)){
    const aoa=[],rowNums=[];
    for(let start=0;start<sheet.length;start+=DENSE_ROW_CHUNK){
      const end=Math.min(start+DENSE_ROW_CHUNK,sheet.length);
      for(let rowIndex=start;rowIndex<end;rowIndex++){
        const cells=sheet[rowIndex];if(!cells)continue;
        const row=cells.map(cellValue);
        if(row.some(value=>value!=='')){aoa.push(row);rowNums.push(rowIndex);}
      }
      if(onChunk&&end<sheet.length)await onChunk(end,sheet.length);
    }
    return {aoa,rowNums};
  }
  const range=XLSX.utils.decode_range(sheet['!ref']),aoa=[],rowNums=[];
  for(let start=range.s.r;start<=range.e.r;start+=DENSE_ROW_CHUNK){
    const end=Math.min(start+DENSE_ROW_CHUNK-1,range.e.r);
    for(let rowIndex=start;rowIndex<=end;rowIndex++){
      const row=[];let populated=false;
      for(let column=range.s.c;column<=range.e.c;column++){
        const value=cellValue(sheet[XLSX.utils.encode_cell({r:rowIndex,c:column})]);
        row.push(value);if(value!=='')populated=true;
      }
      if(populated){aoa.push(row);rowNums.push(rowIndex);}
    }
    if(onChunk&&end<range.e.r)await onChunk(end-range.s.r+1,range.e.r-range.s.r+1);
  }
  return {aoa,rowNums};
}
