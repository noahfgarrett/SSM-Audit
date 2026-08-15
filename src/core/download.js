export function downloadBlob(filename,blob){
  const url=URL.createObjectURL(blob),anchor=document.createElement('a');
  anchor.href=url;anchor.download=filename;anchor.rel='noopener';
  document.body.appendChild(anchor);anchor.click();anchor.remove();
  setTimeout(()=>URL.revokeObjectURL(url),30000);
}

export function workbookBlob(workbook,options={}){
  return new Blob([XLSX.write(workbook,Object.assign({bookType:'xlsx',type:'array',cellStyles:true},options))],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

/* ---- worksheet decoration helpers ----
   Style objects follow the SheetJS per-cell `s` shape (font / fill / alignment /
   border / numFmt). Number formats are mirrored onto `z` as well, because that is
   the format channel every SheetJS build writes. */
export function sheetCellStyle(options={}){
  const style={},font={},alignment={};
  if(options.bold)font.bold=true;
  if(options.italic)font.italic=true;
  if(options.underline)font.underline=true;
  if(options.size)font.sz=options.size;
  if(options.color)font.color={rgb:options.color};
  if(Object.keys(font).length)style.font=font;
  if(options.fill)style.fill={patternType:'solid',fgColor:{rgb:options.fill},bgColor:{rgb:options.fill}};
  if(options.align)alignment.horizontal=options.align;
  if(options.vertical)alignment.vertical=options.vertical;
  if(options.wrap)alignment.wrapText=true;
  if(Object.keys(alignment).length)style.alignment=alignment;
  if(options.border){const side={style:'thin',color:{rgb:options.border}};style.border={top:side,bottom:side,left:side,right:side};}
  if(options.bottom){const bottom={style:options.bottomWeight||'thin',color:{rgb:options.bottom}};style.border=Object.assign({},style.border,{bottom});}
  if(options.numFmt)style.numFmt=options.numFmt;
  return Object.freeze(style);
}

export function sheetGrowRef(sheet,address){
  const cell=XLSX.utils.decode_cell(address);
  if(!sheet['!ref']){sheet['!ref']=XLSX.utils.encode_range({s:cell,e:cell});return;}
  const range=XLSX.utils.decode_range(sheet['!ref']);
  range.s.r=Math.min(range.s.r,cell.r);range.s.c=Math.min(range.s.c,cell.c);
  range.e.r=Math.max(range.e.r,cell.r);range.e.c=Math.max(range.e.c,cell.c);
  sheet['!ref']=XLSX.utils.encode_range(range);
}

export function sheetSetCell(sheet,address,cell){sheet[address]=cell;sheetGrowRef(sheet,address);return cell;}

export function sheetStyleCell(sheet,address,style){
  if(!style)return null;
  const cell=sheet[address]||sheetSetCell(sheet,address,{t:'s',v:''});
  cell.s=style;if(style.numFmt)cell.z=style.numFmt;
  return cell;
}

export function sheetLinkCell(sheet,address,target,tooltip){
  const cell=sheet[address]||sheetSetCell(sheet,address,{t:'s',v:''});
  cell.l={Target:target,Tooltip:tooltip||target};
  return cell;
}

/* A formula cell always carries a cached value so the workbook reads correctly
   before Excel recalculates. */
export function sheetFormulaCell(formula,cachedValue,style){
  const cell=typeof cachedValue==='number'?{t:'n',v:cachedValue,f:formula}:{t:'s',v:String(cachedValue==null?'':cachedValue),f:formula};
  if(style){cell.s=style;if(style.numFmt)cell.z=style.numFmt;}
  return cell;
}

export function sheetFreezeRows(sheet,rowCount){
  if(!sheet||!rowCount)return;
  const topLeft=`A${rowCount+1}`;
  sheet['!freeze']=topLeft;
  sheet['!views']=[{state:'frozen',ySplit:rowCount,topLeftCell:topLeft,activePane:'bottomLeft'}];
}

export function sheetAutoFilter(sheet,ref){if(sheet&&ref)sheet['!autofilter']={ref};}

export function styleHeaderRow(sheet,rowIndex=0){
  if(!sheet||!sheet['!ref'])return;
  const range=XLSX.utils.decode_range(sheet['!ref']);
  for(let column=range.s.c;column<=range.e.c;column++){
    const address=XLSX.utils.encode_cell({r:rowIndex,c:column}),cell=sheet[address];if(!cell)continue;
    cell.s={font:{bold:true,color:{rgb:'FFFFFF'}},fill:{patternType:'solid',fgColor:{rgb:'173F5F'}},alignment:{vertical:'center'}};
  }
  sheet['!rows']=sheet['!rows']||[];sheet['!rows'][rowIndex]={hpt:24};
}
