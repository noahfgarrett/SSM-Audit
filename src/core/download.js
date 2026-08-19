export function downloadBlob(filename,blob){
  const url=URL.createObjectURL(blob),anchor=document.createElement('a');
  anchor.href=url;anchor.download=filename;anchor.rel='noopener';
  document.body.appendChild(anchor);anchor.click();anchor.remove();
  setTimeout(()=>URL.revokeObjectURL(url),30000);
}

const XLSX_MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export function workbookBytes(workbook,options={}){
  const bytes=XLSX.write(workbook,Object.assign({bookType:'xlsx',type:'array',cellStyles:true},options));
  return workbookApplyXmlExtras(workbook,bytes,options);
}
export function workbookBlob(workbook,options={}){return new Blob([workbookBytes(workbook,options)],{type:XLSX_MIME});}

/* ---- worksheet XML extras ----
   The vendored SheetJS build writes cell styles but not data validation or
   conditional formatting. Both are plain worksheet XML, so a sheet may carry them
   as `!xmlExtras` ({dataValidations:[...], conditionalFormatting:[...]} -- raw XML
   fragments) and they are spliced into the written file afterwards. The Open XML
   schema fixes the element order: both fragments belong after <sheetData> and
   <mergeCells>, and BEFORE <hyperlinks>, <pageMargins>, <ignoredErrors>... so
   the splice point is the earliest of those later elements. Differential styles
   referenced by conditional formats (dxfId) go into styles.xml the same way. */
const XML_EXTRA_KEY='!xmlExtras';
const XML_SHEET_TAIL=['<hyperlinks','<printOptions','<pageMargins','<pageSetup','<headerFooter','<rowBreaks','<colBreaks','<customProperties','<cellWatches','<ignoredErrors','<smartTags','<drawing','<legacyDrawing','<picture','<oleObjects','<controls','<webPublishItems','<tableParts','<extLst','</worksheet>'];
export function sheetXmlExtras(sheet,extras){
  const current=sheet[XML_EXTRA_KEY]||{dataValidations:[],conditionalFormatting:[]};
  sheet[XML_EXTRA_KEY]={dataValidations:[...current.dataValidations,...(extras.dataValidations||[])],conditionalFormatting:[...current.conditionalFormatting,...(extras.conditionalFormatting||[])]};
}
export function spliceSheetXml(xml,fragment){
  if(!fragment)return xml;
  let at=-1;
  for(const marker of XML_SHEET_TAIL){const index=xml.indexOf(marker);if(index!==-1&&(at===-1||index<at))at=index;}
  return at===-1?xml+fragment:xml.slice(0,at)+fragment+xml.slice(at);
}
function sheetExtrasXml(extras){
  const formats=(extras.conditionalFormatting||[]).join('');
  const validations=extras.dataValidations||[];
  return formats+(validations.length?`<dataValidations count="${validations.length}">${validations.join('')}</dataValidations>`:'');
}
export function workbookApplyXmlExtras(workbook,bytes,options={}){
  const names=workbook&&workbook.SheetNames||[];
  const touched=names.filter(name=>{const extras=workbook.Sheets[name]&&workbook.Sheets[name][XML_EXTRA_KEY];return extras&&(extras.dataValidations.length||extras.conditionalFormatting.length);});
  const dxfs=Array.isArray(workbook.Dxfs)?workbook.Dxfs:[];
  if(!touched.length&&!dxfs.length)return bytes;
  const container=XLSX.CFB.read(new Uint8Array(bytes),{type:'array'}),decoder=new TextDecoder(),encoder=new TextEncoder();
  const replace=(path,transform)=>{const entry=XLSX.CFB.find(container,path);if(!entry)return;const next=transform(decoder.decode(entry.content));entry.content=encoder.encode(next);entry.size=entry.content.length;};
  names.forEach((name,index)=>{
    const extras=workbook.Sheets[name]&&workbook.Sheets[name][XML_EXTRA_KEY];if(!extras)return;
    const fragment=sheetExtrasXml(extras);if(!fragment)return;
    replace(`/xl/worksheets/sheet${index+1}.xml`,xml=>spliceSheetXml(xml,fragment));
  });
  if(dxfs.length)replace('/xl/styles.xml',xml=>xml.replace(/<dxfs count="0"\/>|<dxfs\/>/,`<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`));
  return XLSX.CFB.write(container,{fileType:'zip',type:'array',compression:options.compression!==false});
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
