import { zipDeflateAvailable, zipEntries } from './zip.js'

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
function workbookExtrasPending(workbook){
  const names=workbook&&workbook.SheetNames||[];
  const touched=names.some(name=>{const extras=workbook.Sheets[name]&&workbook.Sheets[name][XML_EXTRA_KEY];return extras&&(extras.dataValidations.length||extras.conditionalFormatting.length);});
  return touched||(Array.isArray(workbook.Dxfs)&&workbook.Dxfs.length>0);
}
/* Opens the written xlsx as a zip container and splices the extras in place. */
function workbookContainerWithExtras(workbook,bytes){
  const names=workbook&&workbook.SheetNames||[],dxfs=Array.isArray(workbook.Dxfs)?workbook.Dxfs:[];
  const container=XLSX.CFB.read(new Uint8Array(bytes),{type:'array'}),decoder=new TextDecoder(),encoder=new TextEncoder();
  const replace=(path,transform)=>{const entry=XLSX.CFB.find(container,path);if(!entry)return;const next=transform(decoder.decode(entry.content));entry.content=encoder.encode(next);entry.size=entry.content.length;};
  names.forEach((name,index)=>{
    const extras=workbook.Sheets[name]&&workbook.Sheets[name][XML_EXTRA_KEY];if(!extras)return;
    const fragment=sheetExtrasXml(extras);if(!fragment)return;
    replace(`/xl/worksheets/sheet${index+1}.xml`,xml=>spliceSheetXml(xml,fragment));
  });
  if(dxfs.length)replace('/xl/styles.xml',xml=>xml.replace(/<dxfs count="0"\/>|<dxfs\/>/,`<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`));
  return container;
}
export function workbookApplyXmlExtras(workbook,bytes,options={}){
  if(!workbookExtrasPending(workbook))return bytes;
  return XLSX.CFB.write(workbookContainerWithExtras(workbook,bytes),{fileType:'zip',type:'array',compression:options.compression!==false});
}
/* The zip entries of a written workbook (extras applied), [Content_Types].xml
   first as Excel expects. */
function workbookZipEntries(workbook,bytes){
  const container=workbookContainerWithExtras(workbook,bytes),entries=[];
  container.FullPaths.forEach((fullPath,index)=>{
    const name=fullPath.replace(/^Root Entry\/?/,'');const file=container.FileIndex[index];
    if(!name||name.endsWith('/')||name.startsWith('\u0001')||!file||!file.content)return;
    entries.push({name,data:file.content instanceof Uint8Array?file.content:new Uint8Array(file.content)});
  });
  const rank=name=>name==='[Content_Types].xml'?0:name==='_rels/.rels'?1:2;
  return entries.sort((left,right)=>rank(left.name)-rank(right.name));
}
/* Compact output: SheetJS writes the XML, a real deflate (CompressionStream)
   zips it. SheetJS's bundled compressor leaves large workbooks ~3x bigger than
   zlib does. Falls back to SheetJS's own bytes where CompressionStream is
   missing. Shared strings are on: the same finding text repeats on thousands
   of lines and is stored once. */
export async function workbookBytesCompact(workbook,options={}){
  const writeOptions=Object.assign({bookSST:true},options);
  if(!zipDeflateAvailable())return workbookBytes(workbook,Object.assign({compression:true},writeOptions));
  const bytes=XLSX.write(workbook,Object.assign({bookType:'xlsx',type:'array',cellStyles:true,compression:false},writeOptions));
  return zipEntries(workbookZipEntries(workbook,bytes));
}
export async function workbookBlobCompact(workbook,options={}){return new Blob([await workbookBytesCompact(workbook,options)],{type:XLSX_MIME});}

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
